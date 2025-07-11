import streamlit as st
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import os
from sqlalchemy import create_engine, text
from xhtml2pdf import pisa
import io
import html

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

def convert_html_to_pdf(html_content):
    try:
        result = io.BytesIO()
        pdf = pisa.pisaDocument(io.StringIO(html_content), dest=result)
        if not pdf.err:
            return result.getvalue()
        st.error(f"xhtml2pdf error: {pdf.err}")
        return None
    except Exception as e:
        st.error(f"xhtml2pdf failed: {str(e)}")
        return None

# Initialize database connection
conn = init_db()

# Streamlit App
st.title("Project Delivery Dashboard")

# Sidebar for navigation
st.sidebar.header("Navigation")
option = st.sidebar.selectbox("Choose an option", ["Add Project", "Submit Weekly Update", "View Reports"])

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

                    unique_risks = set(risk.strip() for risk in risks.split("\n") if risk.strip())
                    unique_issues = set(issue.strip() for issue in issues.split("\n") if issue.strip())
                    existing_risks_issues = set(row[0] for row in conn.execute(text("SELECT description FROM Risks_Issues WHERE update_id = :uid"), {'uid': update_id}).fetchall())
                    for risk in unique_risks:
                        if risk not in existing_risks_issues:
                            conn.execute(text("INSERT INTO Risks_Issues (update_id, type, description, owner, mitigation_eta) VALUES (:uid, 'Risk', :desc, 'TBD', 'TBD')"), {
                                'uid': update_id,
                                'desc': risk
                            })
                    for issue in unique_issues:
                        if issue not in existing_risks_issues:
                            conn.execute(text("INSERT INTO Risks_Issues (update_id, type, description, owner, mitigation_eta) VALUES (:uid, 'Issue', :desc, 'TBD', 'TBD')"), {
                                'uid': update_id,
                                'desc': issue
                            })

                    unique_actions = set(action.strip() for action in action_items.split("\n") if action.strip())
                    existing_actions = set(row[0] for row in conn.execute(text("SELECT description FROM Action_Items WHERE update_id = :uid"), {'uid': update_id}).fetchall())
                    for action in unique_actions:
                        if action not in existing_actions:
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
elif option == "View Reports":
    st.header("Project Report Generator")

    with st.form("report_form"):
        report_type = st.selectbox("Report Type", ["Weekly Summary", "Project History"])
        week_ending_date = st.date_input("Select Week Ending Date")
        projects = conn.execute(text("SELECT project_id, project_name FROM Projects")).fetchall()
        project_dict = {row.project_name: row.project_id for row in projects}
        project_name = st.selectbox("Select Project (Optional)", ["All"] + list(project_dict.keys()))
        generate_report = st.form_submit_button("Generate Report")

    if generate_report:
        base_query = """
            SELECT DISTINCT p.project_name, p.client_business_unit, p.project_manager, p.start_date, p.end_date, p.current_phase,
                   w.accomplishments, w.decisions_needed, w.milestones, w.status_indicator,
                   r.area, r.status, r.comment,
                   ri.type, ri.description, ri.owner, ri.mitigation_eta,
                   a.description AS action_description, a.status AS action_status, a.client_input_required
            FROM Weekly_Updates w
            JOIN Projects p ON w.project_id = p.project_id
            LEFT JOIN RAG_Status r ON w.update_id = r.update_id
            LEFT JOIN (SELECT DISTINCT update_id, type, description, owner, mitigation_eta FROM Risks_Issues) ri ON w.update_id = ri.update_id
            LEFT JOIN (SELECT DISTINCT update_id, description, status, client_input_required FROM Action_Items) a ON w.update_id = a.update_id
            WHERE CAST(w.week_ending_date AS DATE) = :week
        """
        params = {'week': str(week_ending_date)}
        if project_name != "All":
            base_query += " AND p.project_name = :project_name"
            params['project_name'] = project_name

        try:
            result = conn.execute(text(base_query), params)
            data = result.fetchall()

            if data:
                project_data = {}
                for row in data:
                    project_key = html.escape(row[0])  # Sanitize project name
                    if project_key not in project_data:
                        project_data[project_key] = {
                            'client_business_unit': html.escape(row[1] or ''),
                            'project_manager': html.escape(row[2] or ''),
                            'start_date': row[3],
                            'end_date': row[4],
                            'current_phase': html.escape(row[5] or ''),
                            'accomplishments': html.escape(row[6] or ''),
                            'decisions_needed': html.escape(row[7] or ''),
                            'milestones': html.escape(row[8] or ''),
                            'status_indicator': html.escape(row[9] or ''),
                            'rag_status': [],
                            'risks_issues': [],
                            'action_items': []
                        }
                    if row[10]:
                        project_data[project_key]['rag_status'].append({
                            'area': html.escape(row[10]), 
                            'status': html.escape(row[11]), 
                            'comment': html.escape(row[12] or '')
                        })
                    if row[13]:
                        project_data[project_key]['risks_issues'].append({
                            'type': html.escape(row[13]), 
                            'description': html.escape(row[14]), 
                            'owner': html.escape(row[15]), 
                            'mitigation_eta': html.escape(row[16])
                        })
                    if row[17]:
                        project_data[project_key]['action_items'].append({
                            'description': html.escape(row[17]), 
                            'status': html.escape(row[18]), 
                            'client_input_required': row[19]
                        })

                # Generate HTML for PDF
                html_content = """
                <html>
                <head>
                    <style>
                        body { font-family: Times New Roman, Times, serif; font-size: 12pt; margin: 0.5in; line-height: 1.2; }
                        h1 { font-size: 18pt; text-align: center; color: #003366; margin: 10px 0; }
                        h2 { font-size: 14pt; color: #003366; margin: 8px 0; }
                        h3 { font-size: 12pt; color: #004080; margin: 6px 0; }
                        p, ul { margin: 4px 0; }
                        ul { padding-left: 20px; }
                        table { width: 100%; border-collapse: collapse; margin: 8px 0; }
                        th, td { padding: 4px; text-align: left; font-size: 12pt; }
                        .status-green { color: green; }
                        .status-red { color: red; }
                        .rag-green { color: green; }
                        .rag-orange { color: orange; }
                        .rag-red { color: red; }
                        hr { margin: 8px 0; border: 0.5px solid #ccc; }
                        .project-container { margin-bottom: 16px; }
                    </style>
                </head>
                <body>
                    <h1>Weekly Report for {}</h1>
                    <hr>
                """.format(week_ending_date.strftime('%Y-%m-%d'))

                for idx, (pname, project) in enumerate(project_data.items()):
                    rag_status = "".join(
                        f"<li><strong>{rag['area']}</strong>: <span class='rag-{rag['status'].lower()}'>{rag['status']}</span> - {rag['comment']}</li>"
                        for rag in project['rag_status']
                    ) or "<li>No RAG status available</li>"
                    risks_issues = "".join(
                        f"<li><strong>{ri['type']}</strong>: {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})</li>"
                        for ri in project['risks_issues']
                    ) or "<li>No risks or issues</li>"
                    action_items = "".join(
                        f"<li>{action['description']} - {action['status']} (Client Input: {'Yes' if action['client_input_required'] else 'No'})</li>"
                        for action in project['action_items']
                    ) or "<li>No action items</li>"

                    page_break = "page-break-before: always;" if idx != 0 else ""
                    html_content += f"""
                    <div class="project-container" style="{page_break}">
                        <h2>{pname}</h2>
                        <table>
                            <tr><th>Client/BU</th><td>{project['client_business_unit']}</td></tr>
                            <tr><th>Project Manager</th><td>{project['project_manager']}</td></tr>
                            <tr><th>Duration</th><td>{project['start_date']} to {project['end_date']}</td></tr>
                            <tr><th>Phase</th><td>{project['current_phase']}</td></tr>
                            <tr><th>Status</th><td><span class="status-{'green' if project['status_indicator'] == 'On Track' else 'red'}">{project['status_indicator']}</span></td></tr>
                        </table>
                        <h3>Accomplishments</h3>
                        <ul>{''.join(f'<li>{line}</li>' for line in project['accomplishments'].splitlines() if line.strip()) or '<li>None</li>'}</ul>
                        <h3>Decisions Needed</h3>
                        <ul>{''.join(f'<li>{line}</li>' for line in project['decisions_needed'].splitlines() if line.strip()) or '<li>None</li>'}</ul>
                        <h3>Milestones</h3>
                        <ul><li>{project['milestones'] or 'None'}</li></ul>
                        <h3>RAG Status</h3>
                        <ul>{rag_status}</ul>
                        <h3>Risks & Issues</h3>
                        <ul>{risks_issues}</ul>
                        <h3>Action Items</h3>
                        <ul>{action_items}</ul>
                    </div>
                    <hr>
                    """

                html_content += """
                </body>
                </html>
                """

                # Debug: Save HTML to file for inspection
                with open("debug_report.html", "w", encoding="utf-8") as f:
                    f.write(html_content)
                st.info("HTML content saved to debug_report.html for inspection.")

                # Generate LaTeX for fallback
                latex_content = r"""
                \documentclass[a4paper,11pt]{article}
                \usepackage[utf8]{inputenc}
                \usepackage{geometry}
                \geometry{margin=0.5in}
                \usepackage{xcolor}
                \usepackage{enumitem}
                \usepackage{booktabs}
                \usepackage{parskip}
                \usepackage{titlesec}
                \titleformat{\section}{\large\bfseries}{\thesection}{1em}{}
                \titleformat{\subsection}{\normalsize\bfseries}{\thesubsection}{1em}{}
                \usepackage[T1]{fontenc}
                \usepackage{times}
                \begin{document}
                \begin{center}
                    \textbf{\Large Weekly Report for """ + week_ending_date.strftime('%Y-%m-%d') + r"""}
                \end{center}
                \vspace{0.3cm}
                \hrule
                \vspace{0.3cm}
                """
                for project_name, project in project_data.items():
                    latex_content += r"""
                    \section*{""" + project_name.replace('&', r'\&') + r"""}
                    \begin{tabular}{p{0.4\textwidth} p{0.55\textwidth}}
                        \textbf{Client/BU} & """ + project['client_business_unit'].replace('&', r'\&') + r""" \\
                        \textbf{Project Manager} & """ + project['project_manager'].replace('&', r'\&') + r""" \\
                        \textbf{Dates} & """ + str(project['start_date']) + " -- " + str(project['end_date']) + r""" \\
                        \textbf{Current Phase} & """ + project['current_phase'].replace('&', r'\&') + r""" \\
                        \textbf{Status} & \textcolor{""" + ("green" if project['status_indicator'] == "On Track" else "red") + r"""}{""" + project['status_indicator'].replace('&', r'\&') + r"""} \\
                    \end{tabular}
                    \subsection*{Accomplishments}
                    \begin{itemize}[leftmargin=*]
                    """
                    if project['accomplishments']:
                        for item in project['accomplishments'].split('\n'):
                            if item.strip():
                                latex_content += r"\item " + item.strip().replace('&', r'\&') + "\n"
                    else:
                        latex_content += r"\item None" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \subsection*{Decisions Needed}
                    \begin{itemize}[leftmargin=*]
                    """
                    if project['decisions_needed']:
                        for item in project['decisions_needed'].split('\n'):
                            if item.strip():
                                latex_content += r"\item " + item.strip().replace('&', r'\&') + "\n"
                    else:
                        latex_content += r"\item None" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \subsection*{Milestones}
                    \begin{itemize}[leftmargin=*]
                    """
                    if project['milestones']:
                        latex_content += r"\item " + project['milestones'].replace('&', r'\&') + "\n"
                    else:
                        latex_content += r"\item None" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \subsection*{RAG Status}
                    \begin{itemize}[leftmargin=*]
                    """
                    for rag in project['rag_status']:
                        color = {"Green": "green", "Amber": "orange", "Red": "red"}.get(rag['status'], "black")
                        latex_content += r"\item \textbf{" + rag['area'].replace('&', r'\&') + r"}: \textcolor{" + color + r"}{" + rag['status'].replace('&', r'\&') + r"} -- " + (rag['comment'].replace('&', r'\&') if rag['comment'] else "No comment") + "\n"
                    if not project['rag_status']:
                        latex_content += r"\item No RAG status available" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \subsection*{Risks \& Issues}
                    \begin{itemize}[leftmargin=*]
                    """
                    for ri in project['risks_issues']:
                        latex_content += r"\item \textbf{" + ri['type'].replace('&', r'\&') + r"}: " + ri['description'].replace('&', r'\&') + r" (Owner: " + ri['owner'].replace('&', r'\&') + r", ETA: " + ri['mitigation_eta'].replace('&', r'\&') + r")" + "\n"
                    if not project['risks_issues']:
                        latex_content += r"\item No risks or issues" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \subsection*{Action Items}
                    \begin{itemize}[leftmargin=*]
                    """
                    for action in project['action_items']:
                        client_input = "Yes" if action['client_input_required'] else "No"
                        latex_content += r"\item " + action['description'].replace('&', r'\&') + r" -- " + action['status'].replace('&', r'\&') + r" (Client Input: " + client_input + r")" + "\n"
                    if not project['action_items']:
                        latex_content += r"\item No action items" + "\n"
                    latex_content += r"""
                    \end{itemize}
                    \vspace{0.3cm}
                    \hrule
                    \vspace{0.3cm}
                    """
                latex_content += r"""
                \end{document}
                """

                # Display Preview (Streamlit-native)
                st.markdown(f"<h2 style='text-align: center; color: #003366;'>Weekly Report for {week_ending_date.strftime('%Y-%m-%d')}</h2>", unsafe_allow_html=True)
                st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)
                for project_name, project in project_data.items():
                    with st.container():
                        st.markdown(f"<h3 style='color: #003366;'>{project_name}</h3>", unsafe_allow_html=True)
                        col1, col2 = st.columns([3, 2])
                        with col1:
                            st.markdown(f"<strong>Client/BU:</strong> {project['client_business_unit']}", unsafe_allow_html=True)
                            st.markdown(f"<strong>Project Manager:</strong> {project['project_manager']}", unsafe_allow_html=True)
                            st.markdown(f"<strong>Phase:</strong> {project['current_phase']}", unsafe_allow_html=True)
                        with col2:
                            st.markdown(f"<strong>Dates:</strong> {project['start_date']} - {project['end_date']}", unsafe_allow_html=True)
                            status_color = "green" if project['status_indicator'] == "On Track" else "red"
                            st.markdown(f"<strong>Status:</strong> <span style='color:{status_color};'>{project['status_indicator']}</span>", unsafe_allow_html=True)
                        
                        st.markdown("<h4 style='color: #004080; margin: 6px 0;'>Accomplishments</h4>", unsafe_allow_html=True)
                        accomplishments = project['accomplishments'].splitlines() if project['accomplishments'] else []
                        if accomplishments:
                            for line in accomplishments:
                                if line.strip():
                                    st.markdown(f"- {line.strip()}", unsafe_allow_html=True)
                        else:
                            st.markdown("- None", unsafe_allow_html=True)
                        
                        st.markdown("<h4 style='color: #004080; margin: 6px 0;'>Decisions Needed</h4>", unsafe_allow_html=True)
                        decisions = project['decisions_needed'].splitlines() if project['decisions_needed'] else []
                        if decisions:
                            for line in decisions:
                                if line.strip():
                                    st.markdown(f"- {line.strip()}", unsafe_allow_html=True)
                        else:
                            st.markdown("- None", unsafe_allow_html=True)
                        
                        st.markdown("<h4 style='color: #004080; margin: 6px 0;'>Milestones</h4>", unsafe_allow_html=True)
                        st.markdown(f"- {project['milestones']}" if project['milestones'] else "- None", unsafe_allow_html=True)
                        
                        with st.expander("RAG Status"):
                            if project['rag_status']:
                                for rag in project['rag_status']:
                                    color = {"Green": "green", "Amber": "orange", "Red": "red"}.get(rag['status'], "black")
                                    st.markdown(f"- <strong>{rag['area']}:</strong> <span style='color:{color};'>{rag['status']}</span> - {rag['comment'] or 'No comment'}", unsafe_allow_html=True)
                            else:
                                st.markdown("- No RAG status available")
                        
                        with st.expander("Risks & Issues"):
                            if project['risks_issues']:
                                for ri in project['risks_issues']:
                                    st.markdown(f"- <strong>{ri['type']}:</strong> {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})", unsafe_allow_html=True)
                            else:
                                st.markdown("- No risks or issues")
                        
                        with st.expander("Action Items"):
                            if project['action_items']:
                                for action in project['action_items']:
                                    client_input = "Yes" if action['client_input_required'] else "No"
                                    st.markdown(f"- {action['description']} - {action['status']} (Client Input: {client_input})", unsafe_allow_html=True)
                            else:
                                st.markdown("- No action items")
                        
                        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

                # Generate PDF & Download Button
                try:
                    pdf = convert_html_to_pdf(html_content)
                    if pdf:
                        st.download_button(
                            label="Download PDF Report",
                            data=pdf,
                            file_name=f"Weekly_Report_{week_ending_date.strftime('%Y%m%d')}.pdf",
                            mime="application/pdf"
                        )
                    else:
                        st.error("Failed to generate PDF. Check debug_report.html for issues.")
                        latex_file = io.StringIO(latex_content)
                        st.download_button(
                            label="Download Report as LaTeX (Compile to PDF)",
                            data=latex_file.getvalue().encode('utf-8'),
                            file_name=f"Weekly_Report_{week_ending_date.strftime('%Y%m%d')}.tex",
                            mime="application/x-tex"
                        )
                        st.info("Download the LaTeX file and compile it to PDF using a LaTeX editor like Overleaf (https://www.overleaf.com) or a local tool with the command: `latexmk -pdf report.tex`.")
                except Exception as e:
                    st.error(f"PDF generation failed: {e}")
                    latex_file = io.StringIO(latex_content)
                    st.download_button(
                        label="Download Report as LaTeX (Compile to PDF)",
                        data=latex_file.getvalue().encode('utf-8'),
                        file_name=f"Weekly_Report_{week_ending_date.strftime('%Y%m%d')}.tex",
                        mime="application/x-tex"
                    )
                    st.info("Download the LaTeX file and compile it to PDF using a LaTeX editor like Overleaf (https://www.overleaf.com) or a local tool with the command: `latexmk -pdf report.tex`.")

            else:
                st.warning(f"No data found for {week_ending_date.strftime('%Y-%m-%d')}.")
                dates = conn.execute(text("SELECT DISTINCT week_ending_date FROM Weekly_Updates ORDER BY week_ending_date")).fetchall()
                if dates:
                    st.info("Available week ending dates in database:")
                    for date in dates:
                        st.markdown(f"- {date[0]}")
                else:
                    st.info("No week ending dates found in Weekly_Updates table. Please submit a weekly update first.")
        except Exception as e:
            st.error(f"Error generating report: {str(e)}")

# Close database connection
conn.close()