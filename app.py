import streamlit as st
import pandas as pd
from datetime import datetime
import json
from dotenv import load_dotenv
import os
from sqlalchemy import create_engine, text
from xhtml2pdf import pisa
import io
import bcrypt

# Custom CSS for UI enhancements
st.markdown("""
    <style>
        /* General styling */
        body {
            font-family: 'Roboto', sans-serif;
        }
        .main {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
        }
        .stButton>button {
            background-color: #0058a3;
            color: white;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        .stButton>button:hover {
            background-color: #003f7d;
        }
        .stTextInput>div>input, .stDateInput>div>input, .stSelectbox>div>select {
            border: 1px solid #ced4da;
            border-radius: 5px;
            padding: 8px;
        }
        .stForm {
            border: 1px solid #e0e0e0;
            padding: 20px;
            border-radius: 10px;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .sidebar .sidebar-content {
            background-color: #e9ecef;
            padding: 15px;
            border-radius: 10px;
        }
        .sidebar .stButton>button {
            background-color: #dc3545;
            width: 100%;
            margin-top: 10px;
        }
        .sidebar .stButton>button:hover {
            background-color: #c82333;
        }
        .stSuccess {
            background-color: #d4edda;
            color: #155724;
            border-radius: 5px;
            padding: 10px;
        }
        .stError {
            background-color: #f8d7da;
            color: #721c24;
            border-radius: 5px;
            padding: 10px;
        }
        .stWarning {
            background-color: #fff3cd;
            color: #856404;
            border-radius: 5px;
            padding: 10px;
        }
        .report-card {
            border: 1px solid #e0e0e0;
            padding: 15px;
            border-radius: 10px;
            background-color: white;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .status-ontrack {
            color: #28a745;
            font-weight: bold;
        }
        .status-delayed {
            color: #dc3545;
            font-weight: bold;
        }
        .input-icon {
            display: flex;
            align-items: center;
        }
        .input-icon span {
            margin-right: 10px;
            font-size: 1.2em;
        }
    </style>
""", unsafe_allow_html=True)

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
        # Create Users table if it doesn't exist
        conn.execute(text("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Users')
            CREATE TABLE Users (
                user_id INT IDENTITY(1,1) PRIMARY KEY,
                username NVARCHAR(50) UNIQUE NOT NULL,
                password_hash NVARCHAR(255) NOT NULL
            )
        """))
        conn.commit()
        return conn, engine
    except Exception as e:
        st.error(f"DB Connection Failed: {e}")
        raise

def convert_html_to_pdf(html_content):
    result = io.BytesIO()
    pdf = pisa.pisaDocument(io.StringIO(html_content), dest=result)
    if not pdf.err:
        return result.getvalue()
    return None

# User authentication functions
def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def verify_password(password, password_hash):
    return bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8'))

def check_credentials(username, password, conn):
    try:
        result = conn.execute(
            text("SELECT password_hash FROM Users WHERE username = :username"),
            {'username': username}
        ).fetchone()
        if result and verify_password(password, result[0]):
            return True
        return False
    except Exception as e:
        st.error(f"Authentication error: {e}")
        return False

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.username = None

# Initialize database connection
conn, engine = init_db()

# Login page
if not st.session_state.authenticated:
    st.title("🔐 Login to Project Delivery Dashboard")
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("login_form"):
                st.markdown('<div class="input-icon"><span>👤</span><input placeholder="Username"></div>', unsafe_allow_html=True)
                username = st.text_input("", key="username_input")
                st.markdown('<div class="input-icon"><span>🔒</span><input placeholder="Password" type="password"></div>', unsafe_allow_html=True)
                password = st.text_input("", type="password", key="password_input")
                login_button = st.form_submit_button("Login")
                
                if login_button:
                    if not username or not password:
                        st.error("Please enter both username and password")
                    elif check_credentials(username, password, conn):
                        st.session_state.authenticated = True
                        st.session_state.username = username
                        st.rerun()
                    else:
                        st.error("Invalid username or password")
            
            with st.expander("Admin: Add New User"):
                with st.form("add_user_form"):
                    st.markdown('<div class="input-icon"><span>👤</span><input placeholder="New Username"></div>', unsafe_allow_html=True)
                    new_username = st.text_input("", key="new_username_input")
                    st.markdown('<div class="input-icon"><span>🔒</span><input placeholder="New Password" type="password"></div>', unsafe_allow_html=True)
                    new_password = st.text_input("", type="password", key="new_password_input")
                    add_user_button = st.form_submit_button("Add User")
                    
                    if add_user_button:
                        if not new_username or not new_password:
                            st.error("Please enter both username and password")
                        else:
                            try:
                                password_hash = hash_password(new_password)
                                conn.execute(
                                    text("INSERT INTO Users (username, password_hash) VALUES (:username, :password_hash)"),
                                    {'username': new_username, 'password_hash': password_hash}
                                )
                                conn.commit()
                                st.success(f"User {new_username} added successfully")
                            except Exception as e:
                                st.error(f"Failed to add user: {e}")
    
    conn.close()
    engine.dispose()
else:
    # Original dashboard code with UI enhancements
    with st.container():
        st.title("📊 Project Delivery Dashboard")
        st.markdown(f"Welcome, **{st.session_state.username}**", unsafe_allow_html=True)
    
    # Sidebar for navigation
    st.sidebar.image("https://via.placeholder.com/150x50?text=Logo", use_column_width=True)  # Replace with your logo URL
    st.sidebar.header("Navigation")
    option = st.sidebar.selectbox("Choose an option", ["Add Project", "Submit Weekly Update", "View Reports"], format_func=lambda x: f"📋 {x}")
    
    if st.sidebar.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.session_state.username = None
        st.rerun()

    # Add Project
    if option == "Add Project":
        with st.container():
            st.header("➕ Add New Project")
            with st.form("project_form"):
                project_name = st.text_input("Project Name", placeholder="Enter project name")
                client_business_unit = st.text_input("Client / Business Unit", placeholder="Enter client or business unit")
                project_manager = st.text_input("Project Manager / Delivery Lead", placeholder="Enter project manager name")
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input("Start Date")
                with col2:
                    end_date = st.date_input("End Date")
                current_phase = st.selectbox("Current Phase", ["Build", "Test", "UAT", "Go-Live"])
                submit_project = st.form_submit_button("Submit Project")

                if submit_project:
                    if not all([project_name, client_business_unit, project_manager, start_date, end_date]):
                        st.error("All fields are required!")
                    elif start_date > end_date:
                        st.error("Start Date cannot be later than End Date!")
                    else:
                        try:
                            conn.execute(text("SELECT 1"))
                            insert_stmt = text("""
                                INSERT INTO Projects (project_name, client_business_unit, project_manager, start_date, end_date, current_phase)
                                OUTPUT INSERTED.project_id
                                VALUES (:name, :client, :manager, :start, :end, :phase)
                            """)
                            result = conn.execute(insert_stmt, {
                                'name': project_name.strip(),
                                'client': client_business_unit.strip(),
                                'manager': project_manager.strip(),
                                'start': str(start_date),
                                'end': str(end_date),
                                'phase': current_phase
                            })
                            row = result.fetchone()
                            if row is None:
                                st.error("Failed to insert project: No project_id returned. Check database constraints or schema.")
                            else:
                                project_id = row[0]
                                conn.commit()
                                st.success(f"Project added successfully with project_id: {project_id}")
                        except Exception as e:
                            st.error(f"Failed to add project: {e}")
                            st.write("Debug Info: Check if 'Projects' table exists and has an auto-incrementing 'project_id' column.")
                            raise

    # Submit Weekly Update
    elif option == "Submit Weekly Update":
        with st.container():
            st.header("📅 Weekly Project Update")
            projects = conn.execute(text("SELECT project_id, project_name FROM Projects")).fetchall()
            project_dict = {row.project_name: row.project_id for row in projects}

            if not project_dict:
                st.error("No projects found in the database. Please add a project first using the 'Add Project' section.")
            else:
                with st.form("update_form"):
                    project_name = st.selectbox("Select Project", list(project_dict.keys()))
                    week_ending_date = st.date_input("Week Ending Date")
                    accomplishments = st.text_area("This Week’s Accomplishments (2-3 bullets)", placeholder="Enter accomplishments in bullet points")
                    decisions_needed = st.text_area("Key Decisions Needed / Escalations (1-2 bullets)", placeholder="Enter decisions needed")
                    milestones = st.text_input("Key Milestones", placeholder="Enter milestones")
                    status_indicator = st.selectbox("Status Indicator", ["On Track", "Delayed"])

                    st.subheader("RAG Status")
                    rag_areas = ["Scope", "Timeline", "Cost", "Quality", "Resources"]
                    rag_data = {}
                    for area in rag_areas:
                        col1, col2 = st.columns([1, 2])
                        with col1:
                            status = st.selectbox(f"Status for {area}", ["Green", "Amber", "Red"], key=f"rag_{area}")
                        with col2:
                            comment = st.text_input(f"Comment for {area}", key=f"comment_{area}", placeholder=f"Comment on {area} status")
                        rag_data[area] = {"status": status, "comment": comment}

                    st.subheader("Risks & Issues")
                    risks = st.text_area("Top 2 Risks (Description, Owner, Mitigation)", placeholder="Enter risks in bullet points")
                    issues = st.text_area("Top 2 Issues (Description, Owner, ETA)", placeholder="Enter issues in bullet points")

                    st.subheader("Action Items / Dependencies")
                    action_items = st.text_area("Pending Actions", placeholder="Enter action items in bullet points")
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
    elif option == "View Reports":
        with st.container():
            st.header("📄 Project Report Generator")
            with st.form("report_form"):
                report_type = st.selectbox("Report Type", ["Weekly Summary", "Project History"])
                week_ending_date = st.date_input("Select Week Ending Date")
                projects = conn.execute(text("SELECT project_id, project_name FROM Projects")).fetchall()
                project_dict = {row.project_name: row.project_id for row in projects}
                project_name = st.selectbox("Select Project (Optional)", ["All"] + list(project_dict.keys()))
                col1, col2 = st.columns(2)
                with col1:
                    preview_report = st.form_submit_button("Preview Report")
                with col2:
                    download_report = st.form_submit_button("Download PDF Report")

            if preview_report or download_report:
                base_query = """
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
                            pname = row[0]
                            if pname not in project_data:
                                project_data[pname] = {
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
                                    'risks_issues': set(),
                                    'action_items': set()
                                }
                            if row[10]:
                                project_data[pname]['rag_status'].append({
                                    'area': row[10], 'status': row[11], 'comment': row[12]
                                })
                            if row[13] and row[14]:
                                project_data[pname]['risks_issues'].add((
                                    row[13], row[14], row[15], row[16]
                                ))
                            if row[17]:
                                project_data[pname]['action_items'].add((
                                    row[17], row[18], row[19]
                                ))

                        # Convert sets back to lists for rendering
                        for pname, details in project_data.items():
                            details['risks_issues'] = [
                                {'type': ri[0], 'description': ri[1], 'owner': ri[2], 'mitigation_eta': ri[3]}
                                for ri in details['risks_issues']
                            ]
                            details['action_items'] = [
                                {'description': ai[0], 'status': ai[1], 'client_input_required': ai[2]}
                                for ai in details['action_items']
                            ]

                        # Generate HTML for PDF
                        html = f"""
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body {{ font-family: Arial, sans-serif; margin: 20px; line-height: 1.5; }}
                                h2 {{ color: #003366; margin-bottom: 10px; }}
                                h4 {{ color: #004080; margin-top: 15px; margin-bottom: 8px; }}
                                p {{ margin: 5px 0; }}
                                ul {{ margin: 5px 0; padding-left: 25px; }}
                                li {{ margin-bottom: 5px; }}
                                .project-container {{ margin-bottom: 30px; page-break-inside: avoid; }}
                                .header {{ text-align: center; margin-bottom: 20px; }}
                                .status-ontrack {{ color: green; }}
                                .status-delayed {{ color: red; }}
                            </style>
                        </head>
                        <body>
                            <div class="header">
                                <h2>Weekly Report</h2>
                                <p>Week Ending: {week_ending_date.strftime('%Y-%m-%d')}</p>
                            </div>
                        """
                        for idx, (pname, details) in enumerate(project_data.items()):
                            html += f"""
                            <div class="project-container" style="{'page-break-before: always;' if idx > 0 else ''}">
                                <h2>{pname}</h2>
                                <p><strong>Client/BU:</strong> {details['client_business_unit']}</p>
                                <p><strong>Project Manager:</strong> {details['project_manager']}</p>
                                <p><strong>Duration:</strong> {details['start_date']} to {details['end_date']}</p>
                                <p><strong>Phase:</strong> {details['current_phase']} | 
                                   <strong>Status:</strong> <span class="status-{'ontrack' if details['status_indicator']=='On Track' else 'delayed'}">{details['status_indicator']}</span></p>

                                <h4>Accomplishments</h4>
                                <ul>
                                    {"".join([f"<li>{line.strip()}</li>" for line in details['accomplishments'].splitlines() if line.strip()]) or "<li>No accomplishments reported</li>"}
                                </ul>

                                <h4>Decisions Needed</h4>
                                <ul>
                                    {"".join([f"<li>{line.strip()}</li>" for line in details['decisions_needed'].splitlines() if line.strip()]) or "<li>No decisions needed</li>"}
                                </ul>

                                <h4>Milestones</h4>
                                <p>{details['milestones'] or "None"}</p>

                                <h4>RAG Status</h4>
                                <ul>
                                    {"".join([f"<li><strong>{r['area']}</strong>: {r['status']} - {r['comment'] or 'No comment'}</li>" for r in details['rag_status']]) or "<li>No RAG status available</li>"}
                                </ul>

                                <h4>Risks & Issues</h4>
                                <ul>
                                    {"".join([f"<li><strong>{ri['type']}</strong>: {ri['description']} (Owner: {ri['owner']}, ETA: {ri['mitigation_eta']})</li>" for ri in details['risks_issues']]) or "<li>No risks or issues</li>"}
                                </ul>

                                <h4>Action Items</h4>
                                <ul>
                                    {"".join([f"<li>{a['description']} - {a['status']} (Client Input: {'Yes' if a['client_input_required'] else 'No'})</li>" for a in details['action_items']]) or "<li>No action items</li>"}
                                </ul>
                            </div>
                            """
                        html += """
                        </body>
                        </html>
                        """

                        # Generate PDF
                        pdf_data = convert_html_to_pdf(html)

                        if pdf_data:
                            if preview_report:
                                # Display preview
                                st.markdown("## 📝 Report Preview")
                                for pname, details in project_data.items():
                                    with st.container():
                                        st.markdown(f'<div class="report-card"><h3>{pname}</h3>', unsafe_allow_html=True)
                                        col1, col2 = st.columns(2)
                                        with col1:
                                            st.markdown(f"**Client/BU**: {details['client_business_unit']}")
                                            st.markdown(f"**Project Manager**: {details['project_manager']}")
                                            st.markdown(f"**Phase**: {details['current_phase']}")
                                        with col2:
                                            st.markdown(f"**Start Date**: {details['start_date']}")
                                            st.markdown(f"**End Date**: {details['end_date']}")
                                            st.markdown(f"**Status**: <span class='status-{'ontrack' if details['status_indicator']=='On Track' else 'delayed'}'>{details['status_indicator']}</span>", unsafe_allow_html=True)
                                        st.subheader("Accomplishments")
                                        for line in [l.strip() for l in details['accomplishments'].splitlines() if l.strip()]:
                                            st.markdown(f"- {line}")
                                        st.subheader("Decisions Needed")
                                        for line in [l.strip() for l in details['decisions_needed'].splitlines() if l.strip()]:
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
                                        st.markdown("</div>", unsafe_allow_html=True)

                            # Always provide download option
                            st.download_button(
                                label="📄 Download PDF Report",
                                data=pdf_data,
                                file_name=f"Weekly_Report_{week_ending_date.strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                key="auto_download" if download_report else "preview_download",
                                on_click=lambda: None
                            )
                        else:
                            st.error("Failed to generate PDF.")
                    else:
                        st.warning("No data found for the selected week/project.")
                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")

    # Close database connection
    conn.close()
    engine.dispose()