import streamlit as st
import pyodbc
import pandas as pd
from datetime import datetime
import json
from dotenv import load_dotenv
import os
from sqlalchemy import create_engine, text


# Azure SQL Database connection
def init_db():
    try:
        server = st.secrets["db_server"]
        database = st.secrets["db_name"]
        username = st.secrets["db_user"]
        password = st.secrets["db_password"]
        connection_url = f"mssql+pymssql://{username}:{password}@{server}:1433/{database}"
        engine = create_engine(connection_url)
        conn = engine.connect()
        return conn
    except Exception as e:
        st.error(f"DB Connection Failed: {e}")
        raise


# Initialize database connection
conn = init_db()

# Streamlit App
st.title("Project Delivery Dashboard")

# Sidebar for navigation
st.sidebar.header("Navigation")
option = st.sidebar.selectbox("Choose an option", ["Add Project", "Submit Weekly Update"])

# Add Project
if option == "Add Project":
    st.header("Add New Project")
    with st.form("project_form"):
        project_name = st.text_input("Project Name")
        client_business_unit = st.text_input("Client / Business Unit")
        project_manager = st.text_input("Project Manager / Delivery Lead")
        start_date = st.date_input("Start Date")
        end_date = st.date_input("End Date")
        current_phase = st.selectbox("Current Phase", ["Build", "Test", "UAT", "Go-Live"])
        submit_project = st.form_submit_button("Submit Project")

        if submit_project:
            try:
                insert_stmt = text("""
                    INSERT INTO Projects (project_name, client_business_unit, project_manager, start_date, end_date, current_phase)
                    OUTPUT INSERTED.project_id
                    VALUES (:name, :client, :manager, :start, :end, :phase)
                """)
                result = conn.execute(insert_stmt, {
                    'name': project_name,
                    'client': client_business_unit,
                    'manager': project_manager,
                    'start': str(start_date),
                    'end': str(end_date),
                    'phase': current_phase
                })
                conn.commit()
                project_id = result.fetchone()[0]
                st.success(f"Project added successfully with project_id: {project_id}")
            except Exception as e:
                st.error(f"Failed to add project: {e}")
                raise

# Submit Weekly Update
elif option == "Submit Weekly Update":
    st.header("Weekly Project Update")
    projects = conn.execute(text("SELECT project_id, project_name FROM Projects")).fetchall()
    project_dict = {row.project_name: row.project_id for row in projects}

    if not project_dict:
        st.error("No projects found in the database. Please add a project first using the 'Add Project' section.")
    else:
        with st.form("update_form"):
            project_name = st.selectbox("Select Project", list(project_dict.keys()))
            week_ending_date = st.date_input("Week Ending Date")
            accomplishments = st.text_area("This Week’s Accomplishments (2-3 bullets)")
            decisions_needed = st.text_area("Key Decisions Needed / Escalations (1-2 bullets)")
            milestones = st.text_input("Key Milestones")
            status_indicator = st.selectbox("Status Indicator", ["On Track", "Delayed"])

            st.subheader("RAG Status")
            rag_areas = ["Scope", "Timeline", "Cost", "Quality", "Resources"]
            rag_data = {}
            for area in rag_areas:
                status = st.selectbox(f"Status for {area}", ["Green", "Amber", "Red"], key=f"rag_{area}")
                comment = st.text_input(f"Comment for {area}", key=f"comment_{area}")
                rag_data[area] = {"status": status, "comment": comment}

            st.subheader("Risks & Issues")
            risks = st.text_area("Top 2 Risks (Description, Owner, Mitigation)")
            issues = st.text_area("Top 2 Issues (Description, Owner, ETA)")

            st.subheader("Action Items / Dependencies")
            action_items = st.text_area("Pending Actions")
            client_inputs = st.checkbox("Client Inputs / Approvals Required")

            submit_update = st.form_submit_button("Submit Update")
            if submit_update:
                try:
                    project_id = project_dict[project_name]
                    insert_update = text("""
                        INSERT INTO Weekly_Updates (project_id, week_ending_date, accomplishments, decisions_needed, milestones, status_indicator)
                        OUTPUT INSERTED.update_id
                        VALUES (:pid, :week, :acc, :dec, :mile, :status)
                    """)
                    result = conn.execute(insert_update, {
                        'pid': project_id,
                        'week': str(week_ending_date),
                        'acc': accomplishments,
                        'dec': decisions_needed,
                        'mile': milestones,
                        'status': status_indicator
                    })
                    update_id = result.fetchone()[0]

                    for area, data in rag_data.items():
                        conn.execute(text("INSERT INTO RAG_Status (update_id, area, status, comment) VALUES (:uid, :area, :status, :comment)"), {
                            'uid': update_id,
                            'area': area,
                            'status': data['status'],
                            'comment': data['comment']
                        })

                    for risk in risks.split("\n"):
                        if risk.strip():
                            conn.execute(text("INSERT INTO Risks_Issues (update_id, type, description, owner, mitigation_eta) VALUES (:uid, 'Risk', :desc, 'TBD', 'TBD')"), {
                                'uid': update_id,
                                'desc': risk
                            })

                    for issue in issues.split("\n"):
                        if issue.strip():
                            conn.execute(text("INSERT INTO Risks_Issues (update_id, type, description, owner, mitigation_eta) VALUES (:uid, 'Issue', :desc, 'TBD', 'TBD')"), {
                                'uid': update_id,
                                'desc': issue
                            })

                    for action in action_items.split("\n"):
                        if action.strip():
                            conn.execute(text("INSERT INTO Action_Items (update_id, description, status, client_input_required) VALUES (:uid, :desc, 'Pending', :client)"), {
                                'uid': update_id,
                                'desc': action,
                                'client': 1 if client_inputs else 0
                            })

                    conn.commit()
                    st.success("Weekly update submitted successfully!")
                except Exception as e:
                    st.error(f"Failed to submit weekly update: {e}")
                    raise

# View Reports
# View Reports
elif option == "View Reports":
    import pdfkit
    import base64
    import tempfile

    st.header("Project Report Generator")
    with st.form("report_form"):
        report_type = st.selectbox("Report Type", ["Weekly Summary", "Project History"])
        week_ending_date = st.date_input("Select Week Ending Date")
        c = conn.cursor()
        c.execute("SELECT project_id, project_name FROM Projects")
        projects = c.fetchall()
        project_dict = {name: id for id, name in projects}
        project_name = st.selectbox("Select Project (Optional)", ["All"] + list(project_dict.keys()))
        generate_report = st.form_submit_button("Generate Report")

    if generate_report:
        c = conn.cursor()
        query = """
            SELECT p.project_name, p.client_business_unit, p.project_manager, p.start_date, p.end_date, p.current_phase,
                   w.accomplishments, w.decisions_needed, w.milestones, w.status_indicator,
                   r.area, r.status, r.comment,
                   ri.type, ri.description, ri.owner, ri.mitigation_eta,
                   a.description AS action_description, a.status AS action_status, a.client_input_required
            FROM Weekly_Updates w
            JOIN Projects p ON w.project_id = p.project_id
            LEFT JOIN RAG_Status r ON w.update_id = r.update_id
            LEFT JOIN Risks_Issues ri ON w.update_id = ri.update_id
            LEFT JOIN Action_Items a ON w.update_id = a.update_id
            WHERE CAST(w.week_ending_date AS DATE) = ?
        """
        params = [str(week_ending_date)]
        if project_name != "All":
            query += " AND p.project_name = ?"
            params.append(project_name)

        try:
            c.execute(query, params)
            data = c.fetchall()

            if data:
                project_data = {}
                for row in data:
                    project_key = row[0]  # project_name
                    if project_key not in project_data:
                        project_data[project_key] = {
                            'client_business_unit': row[1],
                            'project_manager': row[2],
                            'start_date': row[3],
                            'end_date': row[4],
                            'current_phase': row[5],
                            'accomplishments': row[6],
                            'decisions_needed': row[7],
                            'milestones': row[8],
                            'status_indicator': row[9],
                            'rag_status': [],
                            'risks_issues': [],
                            'action_items': []
                        }
                    if row[10]:
                        project_data[project_key]['rag_status'].append({
                            'area': row[10], 'status': row[11], 'comment': row[12]
                        })
                    if row[13]:
                        project_data[project_key]['risks_issues'].append({
                            'type': row[13], 'description': row[14], 'owner': row[15], 'mitigation_eta': row[16]
                        })
                    if row[17]:
                        project_data[project_key]['action_items'].append({
                            'description': row[17], 'status': row[18], 'client_input_required': row[19]
                        })

                # Generate HTML for PDF
                # html = "<h2 style='text-align:center;'>Weekly Report</h2>"
                # html += f"<p style='text-align:center;'>Week Ending: {week_ending_date.strftime('%Y-%m-%d')}</p><hr>"
                # for pname, details in project_data.items():
                #     html += f"""
                #     <h3>{pname}</h3>
                #     <p><strong>Client/BU:</strong> {details['client_business_unit']}<br>
                #     <strong>Project Manager:</strong> {details['project_manager']}<br>
                #     <strong>Duration:</strong> {details['start_date']} to {details['end_date']}<br>
                #     <strong>Phase:</strong> {details['current_phase']}<br>
                #     <strong>Status:</strong> {details['status_indicator']}</p>
                    
                #     <h4>Accomplishments</h4>
                #     <ul>{"".join([f"<li>{line}</li>" for line in details['accomplishments'].splitlines() if line])}</ul>
                    
                #     <h4>Decisions Needed</h4>
                #     <ul>{"".join([f"<li>{line}</li>" for line in details['decisions_needed'].splitlines() if line])}</ul>

                #     <h4>Milestones</h4>
                #     <p>{details['milestones'] or "None"}</p>

                #     <h4>RAG Status</h4>
                #     <ul>{"".join([f"<li><b>{r['area']}</b>: {r['status']} - {r['comment'] or 'No comment'}</li>" for r in details['rag_status']]) or "<li>No RAG status available</li>"}</ul>

                #     <h4>Risks & Issues</h4>
                #     <ul>{"".join([f"<li><b>{ri['type']}</b>: {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})</li>" for ri in details['risks_issues']]) or "<li>No risks or issues</li>"}</ul>

                #     <h4>Action Items</h4>
                #     <ul>{"".join([f"<li>{a['description']} - {a['status']} (Client Input: {'Yes' if a['client_input_required'] else 'No'})</li>" for a in details['action_items']]) or "<li>No action items</li>"}</ul>
                #     <hr>
                #     """

                html = "<h2 style='text-align:center;'>Weekly Report</h2>"
                html += f"<p style='text-align:center;'>Week Ending: {week_ending_date.strftime('%Y-%m-%d')}</p><hr>"

                for idx, (pname, details) in enumerate(project_data.items()):
                    html += f"""
                    <div style="{'page-break-before: always;' if idx != 0 else ''} font-family: Arial, sans-serif; padding: 20px;">
                        <h2 style="color:#003366;">{pname}</h2>
                        <p><strong>Client/BU:</strong> {details['client_business_unit']}<br>
                        <strong>Project Manager:</strong> {details['project_manager']}<br>
                        <strong>Duration:</strong> {details['start_date']} to {details['end_date']}<br>
                        <strong>Phase:</strong> {details['current_phase']}<br>
                        <strong>Status:</strong> <span style="color:{'green' if details['status_indicator']=='On Track' else 'red'}">{details['status_indicator']}</span></p>

                        <h4 style="color:#004080;">Accomplishments</h4>
                        <ul>{"".join([f"<li>{line}</li>" for line in details['accomplishments'].splitlines() if line])}</ul>

                        <h4 style="color:#004080;">Decisions Needed</h4>
                        <ul>{"".join([f"<li>{line}</li>" for line in details['decisions_needed'].splitlines() if line])}</ul>

                        <h4 style="color:#004080;">Milestones</h4>
                        <p>{details['milestones'] or "None"}</p>

                        <h4 style="color:#004080;">RAG Status</h4>
                        <ul>{"".join([f"<li><b>{r['area']}</b>: {r['status']} - {r['comment'] or 'No comment'}</li>" for r in details['rag_status']]) or "<li>No RAG status available</li>"}</ul>

                        <h4 style="color:#004080;">Risks & Issues</h4>
                        <ul>{"".join([f"<li><b>{ri['type']}</b>: {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})</li>" for ri in details['risks_issues']]) or "<li>No risks or issues</li>"}</ul>

                        <h4 style="color:#004080;">Action Items</h4>
                        <ul>{"".join([f"<li>{a['description']} - {a['status']} (Client Input: {'Yes' if a['client_input_required'] else 'No'})</li>" for a in details['action_items']]) or "<li>No action items</li>"}</ul>
                    </div>
                    """



                # 🖥️ Display Preview (Streamlit-native)
                st.markdown("## 📝 Report Preview")
                for pname, details in project_data.items():
                    with st.container():
                        st.markdown(f"### {pname}")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown(f"**Client/BU**: {details['client_business_unit']}")
                            st.markdown(f"**Project Manager**: {details['project_manager']}")
                            st.markdown(f"**Phase**: {details['current_phase']}")
                        with col2:
                            st.markdown(f"**Start Date**: {details['start_date']}")
                            st.markdown(f"**End Date**: {details['end_date']}")
                            st.markdown(f"**Status**: {details['status_indicator']}")
                        st.subheader("Accomplishments")
                        for line in details['accomplishments'].splitlines():
                            st.markdown(f"- {line}")
                        st.subheader("Decisions Needed")
                        for line in details['decisions_needed'].splitlines():
                            st.markdown(f"- {line}")
                        st.subheader("Milestones")
                        st.markdown(details['milestones'] or "- None")
                        st.subheader("RAG Status")
                        if details['rag_status']:
                            for r in details['rag_status']:
                                st.markdown(f"- **{r['area']}**: {r['status']} - {r['comment'] or 'No comment'}")
                        else:
                            st.markdown("- No RAG status available")
                        st.subheader("Risks & Issues")
                        if details['risks_issues']:
                            for ri in details['risks_issues']:
                                st.markdown(f"- **{ri['type']}**: {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})")
                        else:
                            st.markdown("- No risks or issues")
                        st.subheader("Action Items")
                        if details['action_items']:
                            for a in details['action_items']:
                                client_input = "Yes" if a['client_input_required'] else "No"
                                st.markdown(f"- {a['description']} - {a['status']} (Client Input: {client_input})")
                        else:
                            st.markdown("- No action items")
                        st.markdown("---")

                # 📄 Generate PDF & Download Button
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    pdfkit.from_string(html, tmpfile.name)
                    with open(tmpfile.name, "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode("utf-8")
                        href = f'<a href="data:application/pdf;base64,{base64_pdf}" download="Weekly_Report_{week_ending_date.strftime("%Y%m%d")}.pdf">📄 Download PDF Report</a>'
                        st.markdown(href, unsafe_allow_html=True)

            else:
                st.warning("No data found for the selected week/project.")
        except Exception as e:
            st.error(f"Error generating report: {str(e)}")


            
            # Chart for RAG Status Distribution
            c.execute(
                "SELECT area, status, COUNT(*) FROM RAG_Status r JOIN Weekly_Updates w ON r.update_id = w.update_id WHERE CAST(w.week_ending_date AS DATE) = ? GROUP BY area, status",
                (str(week_ending_date),)
            )
            rag_counts = c.fetchall()
            
            if rag_counts:
                areas = list(set([x[0] for x in rag_counts]))
                green_counts = [0] * len(areas)
                amber_counts = [0] * len(areas)
                red_counts = [0] * len(areas)
                
                for area, status, count in rag_counts:
                    idx = areas.index(area)
                    if status == "Green":
                        green_counts[idx] = count
                    elif status == "Amber":
                        amber_counts[idx] = count
                    elif status == "Red":
                        red_counts[idx] = count
                
                st.markdown("## RAG Status Distribution")
                st.write("""
                ```chartjs
                {
                    "type": "bar",
                    "data": {
                        "labels": """ + json.dumps(areas) + """,
                        "datasets": [
                            {
                                "label": "Green",
                                "data": """ + json.dumps(green_counts) + """,
                                "backgroundColor": "#00FF00",
                                "borderColor": "#00CC00",
                                "borderWidth": 1
                            },
                            {
                                "label": "Amber",
                                "data": """ + json.dumps(amber_counts) + """,
                                "backgroundColor": "#FFA500",
                                "borderColor": "#CC8400",
                                "borderWidth": 1
                            },
                            {
                                "label": "Red",
                                "data": """ + json.dumps(red_counts) + """,
                                "backgroundColor": "#FF0000",
                                "borderColor": "#CC0000",
                                "borderWidth": 1
                            }
                        ]
                    },
                    "options": {
                        "plugins": {
                            "title": {
                                "display": true,
                                "text": "RAG Status Distribution by Area",
                                "font": {
                                    "size": 18
                                }
                            }
                        },
                        "scales": {
                            "y": {
                                "beginAtZero": true,
                                "title": {
                                    "display": true,
                                    "text": "Count"
                                }
                            },
                            "x": {
                                "title": {
                                    "display": true,
                                    "text": "Area"
                                }
                            }
                        }
                    }
                }
                ```
                """)
            else:
                st.warning(f"No RAG status data available for {week_ending_date.strftime('%Y-%m-%d')}.")

# Close database connection
conn.close()