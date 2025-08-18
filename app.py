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

# Custom CSS for enhanced UI
st.markdown("""
    <style>
        /* General styling */
        body {
            font-family: 'Open Sans', sans-serif;
            background: linear-gradient(135deg, #e0eafc, #cfdef3);
            margin: 0;
            padding: 0;
        }
        .stApp {
            background: rgba(255, 255, 255, 0.9);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
            backdrop-filter: blur(5px);
        }
        .stHeader {
            color: #1e3a8a;
            font-size: 2em;
            font-weight: 700;
            text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
        }
        .stSubheader {
            color: #2b6cb0;
            font-size: 1.5em;
            font-weight: 600;
        }

        /* Button styling with animation */
        .stButton>button {
            background: linear-gradient(45deg, #1e40af, #3b82f6);
            color: white;
            border: none;
            border-radius: 10px;
            padding: 12px 25px;
            font-weight: 600;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        .stButton>button:hover {
            background: linear-gradient(45deg, #1e3a8a, #60a5fa);
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(30, 64, 175, 0.4);
        }
        .stButton>button:active {
            transform: translateY(0);
        }
        .stButton>button::after {
            content: '';
            position: absolute;
            width: 0;
            height: 0;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 50%;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            transition: width 0.6s ease, height 0.6s ease;
        }
        .stButton>button:hover::after {
            width: 200px;
            height: 200px;
        }
        .stButton>button:focus {
            outline: none;
            box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.5);
        }

        /* Input and form styling */
        .stTextInput>div>input, .stDateInput>div>input, .stSelectbox>div>select, .stTextArea>div>textarea {
            border: 2px solid #93c5fd;
            border-radius: 8px;
            padding: 10px;
            font-size: 1.1em;
            background: #f1f5f9;
            transition: border-color 0.3s, box-shadow 0.3s;
        }
        .stTextInput>div>input:focus, .stDateInput>div>input:focus, .stSelectbox>div>select:focus, .stTextArea>div>textarea:focus {
            border-color: #3b82f6;
            box-shadow: 0 0 8px rgba(59, 130, 246, 0.4);
        }
        .stForm {
            padding: 20px;
            border: 2px solid #e0e7ff;
            border-radius: 12px;
            background: #ffffff;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }

        /* Sidebar styling */
        .sidebar .sidebar-content {
            background: linear-gradient(135deg, #e0e7ff, #f1f5f9);
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }
        .sidebar .stHeader {
            color: #1e3a8a;
            font-size: 1.6em;
            font-weight: 700;
        }
        .sidebar .stButton>button {
            background: linear-gradient(45deg, #ef4444, #f87171);
            color: white;
            border-radius: 10px;
            padding: 10px;
            margin-top: 15px;
            transition: all 0.3s ease;
        }
        .sidebar .stButton>button:hover {
            background: linear-gradient(45deg, #dc2626, #f43f5e);
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(239, 68, 68, 0.4);
        }

        /* Feedback messages with icons */
        .stSuccess {
            background-color: #d4edda;
            color: #155724;
            border-radius: 8px;
            padding: 12px;
            border-left: 5px solid #28a745;
            display: flex;
            align-items: center;
        }
        .stSuccess::before {
            content: '✅';
            margin-right: 10px;
        }
        .stError {
            background-color: #f8d7da;
            color: #721c24;
            border-radius: 8px;
            padding: 12px;
            border-left: 5px solid #dc3545;
            display: flex;
            align-items: center;
        }
        .stError::before {
            content: '❌';
            margin-right: 10px;
        }
        .stWarning {
            background-color: #fff3cd;
            color: #856404;
            border-radius: 8px;
            padding: 12px;
            border-left: 5px solid #ffc107;
            display: flex;
            align-items: center;
        }
        .stWarning::before {
            content: '⚠️';
            margin-right: 10px;
        }

        /* Expander styling */
        .stExpander {
            border: 2px solid #e0e7ff;
            border-radius: 12px;
            margin-bottom: 15px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }
        .stExpander > div > div {
            padding: 15px;
        }

        /* Column spacing */
        .stColumns > div {
            padding: 8px;
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
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'qeUsers')
            CREATE TABLE qeUsers (
                user_id INT IDENTITY(1,1) PRIMARY KEY,
                username NVARCHAR(50) UNIQUE NOT NULL,
                password_hash NVARCHAR(255) NOT NULL
            )
        """))
        # Create Projects table if it doesn't exist
        conn.execute(text("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'qeProjects')
            CREATE TABLE qeProjects (
                project_id INT IDENTITY(1,1) PRIMARY KEY,
                project_name NVARCHAR(255) NOT NULL,
                client NVARCHAR(255),
                project_spoc NVARCHAR(255),
                technology_used NVARCHAR(255),
                artifacts_link NVARCHAR(MAX)
            )
        """))
        # Create Weekly_Updates table if it doesn't exist
        conn.execute(text("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'qeWeekly_Updates')
            CREATE TABLE qeWeekly_Updates (
                update_id INT IDENTITY(1,1) PRIMARY KEY,
                project_id INT NOT NULL,
                week_ending_date DATE,
                qe_overall_status NVARCHAR(10),
                qe_progress_percentage INT,
                current_week_progress_entry NVARCHAR(MAX),
                next_release_date DATE,
                qe_team_size INT,
                qe_current_week_task NVARCHAR(MAX),
                qe_automation_tools_used NVARCHAR(MAX),
                tc_created INT,
                tc_executed INT,
                tc_passed_first_round INT,
                effort_tc_execution FLOAT,
                tc_automated INT,
                effort_tc_automation FLOAT,
                defects_raised_internal INT,
                sit_defects INT,
                uat_defects INT,
                reopened_defects INT,
                FOREIGN KEY (project_id) REFERENCES qeProjects(project_id)
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
            text("SELECT password_hash FROM qeUsers WHERE username = :username"),
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
    st.title("Login to QE Weekly Status Dashboard")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
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
    
    # Admin option to add new users (for initial setup or admin access)
    with st.expander("Admin: Add New User"):
        with st.form("add_user_form"):
            new_username = st.text_input("New Username")
            new_password = st.text_input("New Password", type="password")
            auth_code = st.text_input("Authentication Code", type="password")
            add_user_button = st.form_submit_button("Add User")
            
            if add_user_button:
                if not new_username or not new_password or not auth_code:
                    st.error("Please enter all fields including the Authentication Code")
                elif auth_code != "SECURE123":
                    st.error("Invalid Authentication Code")
                else:
                    try:
                        password_hash = hash_password(new_password)
                        conn.execute(
                            text("INSERT INTO qeUsers (username, password_hash) VALUES (:username, :password_hash)"),
                            {'username': new_username, 'password_hash': password_hash}
                        )
                        conn.commit()
                        st.success(f"User {new_username} added successfully")
                    except Exception as e:
                        st.error(f"Failed to add user: {e}")
    
    conn.close()
    engine.dispose()
else:
    # Dashboard code
    st.title("QE Weekly Status Dashboard")
    st.sidebar.header(f"Welcome, {st.session_state.username}")
    
    # Logout button
    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.username = None
        st.rerun()

    # Sidebar for navigation
    st.sidebar.header("Navigation")
    option = st.sidebar.selectbox("Choose an option", ["Add Project", "Submit Weekly Update", "View Reports"])

    # Add Project
    if option == "Add Project":
        st.header("Add New Project")
        with st.form("project_form"):
            project_name = st.text_input("Project Name")
            client = st.text_input("Client")
            project_spoc = st.text_input("Project SPOC")
            technology_used = st.text_input("Technology Used")
            artifacts_link = st.text_input("Project Artifacts Link")
            submit_project = st.form_submit_button("Submit Project")

            if submit_project:
                if not all([project_name, client, project_spoc, technology_used, artifacts_link]):
                    st.error("All fields are required!")
                else:
                    try:
                        conn.execute(text("SELECT 1"))
                        insert_stmt = text("""
                            INSERT INTO qeProjects (project_name, client, project_spoc, technology_used, artifacts_link)
                            OUTPUT INSERTED.project_id
                            VALUES (:name, :client, :spoc, :tech, :link)
                        """)
                        result = conn.execute(insert_stmt, {
                            'name': project_name.strip(),
                            'client': client.strip(),
                            'spoc': project_spoc.strip(),
                            'tech': technology_used.strip(),
                            'link': artifacts_link.strip()
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
        st.header("Weekly QE Update")
        projects = conn.execute(text("SELECT project_id, project_name FROM qeProjects")).fetchall()
        project_dict = {row.project_name: row.project_id for row in projects}

        if not project_dict:
            st.error("No projects found in the database. Please add a project first using the 'Add Project' section.")
        else:
            with st.form("update_form"):
                project_name = st.selectbox("Select Project", list(project_dict.keys()))
                week_ending_date = st.date_input("Week Ending Date")

                st.subheader("QE Status & Progress")
                qe_overall_status = st.selectbox("QE Overall Status", ["GREEN", "AMBER", "RED"])
                qe_progress_percentage = st.number_input("QE Progress Percentage", min_value=0, max_value=100, step=1)
                current_week_progress_entry = st.text_area("Current Week Entry on Overall Progress")
                next_release_date = st.date_input("Next Release Date")

                st.subheader("QE Team & Resources")
                qe_team_size = st.number_input("QE Team Size", min_value=0, step=1)
                qe_current_week_task = st.text_area("QE Current Week Task")
                qe_automation_tools_used = st.text_area("QE Automation Tools Used")

                st.subheader("Test Case Metrics")
                tc_created = st.number_input("#TC Created", min_value=0, step=1)
                tc_executed = st.number_input("#TC Executed", min_value=0, step=1)
                tc_passed_first_round = st.number_input("#TC Passed in First Round of Validation", min_value=0, step=1)
                effort_tc_execution = st.number_input("Effort Spent on TC Execution (hours)", min_value=0.0, format="%.2f")
                tc_automated = st.number_input("#TC Automated", min_value=0, step=1)
                effort_tc_automation = st.number_input("Efforts Spent on TC Automation (hours)", min_value=0.0, format="%.2f")

                st.subheader("Defects & Quality Metrics")
                defects_raised_internal = st.number_input("Defects Raised (Internal)", min_value=0, step=1)
                sit_defects = st.number_input("#SIT Defects", min_value=0, step=1)
                uat_defects = st.number_input("#UAT Defects", min_value=0, step=1)
                reopened_defects = st.number_input("#Reopened Defects", min_value=0, step=1)

                submit_update = st.form_submit_button("Submit Update")
                if submit_update:
                    try:
                        project_id = project_dict[project_name]
                        insert_update = text("""
                            INSERT INTO qeWeekly_Updates (
                                project_id, week_ending_date, qe_overall_status, qe_progress_percentage,
                                current_week_progress_entry, next_release_date, qe_team_size,
                                qe_current_week_task, qe_automation_tools_used, tc_created,
                                tc_executed, tc_passed_first_round, effort_tc_execution,
                                tc_automated, effort_tc_automation, defects_raised_internal,
                                sit_defects, uat_defects, reopened_defects
                            )
                            OUTPUT INSERTED.update_id
                            VALUES (
                                :pid, :week, :status, :progress, :entry, :release, :size,
                                :task, :tools, :tc_created, :tc_executed, :tc_passed,
                                :effort_exec, :tc_auto, :effort_auto, :def_internal,
                                :sit, :uat, :reopened
                            )
                        """)
                        result = conn.execute(insert_update, {
                            'pid': project_id,
                            'week': str(week_ending_date),
                            'status': qe_overall_status,
                            'progress': qe_progress_percentage,
                            'entry': current_week_progress_entry,
                            'release': str(next_release_date) if next_release_date else None,
                            'size': qe_team_size,
                            'task': qe_current_week_task,
                            'tools': qe_automation_tools_used,
                            'tc_created': tc_created,
                            'tc_executed': tc_executed,
                            'tc_passed': tc_passed_first_round,
                            'effort_exec': effort_tc_execution,
                            'tc_auto': tc_automated,
                            'effort_auto': effort_tc_automation,
                            'def_internal': defects_raised_internal,
                            'sit': sit_defects,
                            'uat': uat_defects,
                            'reopened': reopened_defects
                        })
                        update_id = result.fetchone()[0]
                        conn.commit()
                        st.success("Weekly update submitted successfully!")
                    except Exception as e:
                        st.error(f"Failed to submit weekly update: {e}")
                        raise

    # View Reports
    elif option == "View Reports":
        st.header("QE Report Generator")

        with st.form("report_form"):
            week_ending_date = st.date_input("Select Week Ending Date")
            projects = conn.execute(text("SELECT project_id, project_name FROM qeProjects")).fetchall()
            project_dict = {row.project_name: row.project_id for row in projects}
            project_name = st.selectbox("Select Project (Optional)", ["All"] + list(project_dict.keys()))
            col1, col2 = st.columns(2)
            with col1:
                preview_report = st.form_submit_button("Preview Report")
            with col2:
                download_report = st.form_submit_button("Download PDF Report")

        if preview_report or download_report:
            base_query = """
                        SELECT p.project_name, p.client, p.project_spoc, p.technology_used, p.artifacts_link,
                               w.qe_overall_status, w.qe_progress_percentage, w.current_week_progress_entry, w.next_release_date,
                               w.qe_team_size, w.qe_current_week_task, w.qe_automation_tools_used,
                               w.tc_created, w.tc_executed, w.tc_passed_first_round, w.effort_tc_execution,
                               w.tc_automated, w.effort_tc_automation,
                               w.defects_raised_internal, w.sit_defects, w.uat_defects, w.reopened_defects
                        FROM Weekly_Updates w
                        JOIN qeProjects p ON w.project_id = p.project_id
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
                                'client': row[1],
                                'project_spoc': row[2],
                                'technology_used': row[3],
                                'artifacts_link': row[4],
                                'qe_overall_status': row[5],
                                'qe_progress_percentage': row[6],
                                'current_week_progress_entry': row[7],
                                'next_release_date': row[8],
                                'qe_team_size': row[9],
                                'qe_current_week_task': row[10],
                                'qe_automation_tools_used': row[11],
                                'tc_created': row[12],
                                'tc_executed': row[13],
                                'tc_passed_first_round': row[14],
                                'effort_tc_execution': row[15],
                                'tc_automated': row[16],
                                'effort_tc_automation': row[17],
                                'defects_raised_internal': row[18],
                                'sit_defects': row[19],
                                'uat_defects': row[20],
                                'reopened_defects': row[21]
                            }

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
                            .status-green {{ color: green; }}
                            .status-amber {{ color: orange; }}
                            .status-red {{ color: red; }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <h2>Weekly QE Status Report</h2>
                            <p>Week Ending: {week_ending_date.strftime('%Y-%m-%d')}</p>
                        </div>
                    """
                    for idx, (pname, details) in enumerate(project_data.items()):
                        status_class = details['qe_overall_status'].lower()
                        html += f"""
                        <div class="project-container" style="{'page-break-before: always;' if idx > 0 else ''}">
                            <h2>{pname}</h2>
                            <p><strong>Client:</strong> {details['client']}</p>
                            <p><strong>Project SPOC:</strong> {details['project_spoc']}</p>
                            <p><strong>Technology Used:</strong> {details['technology_used']}</p>
                            <p><strong>Artifacts Link:</strong> <a href="{details['artifacts_link']}">{details['artifacts_link']}</a></p>

                            <h4>QE Status & Progress</h4>
                            <p><strong>Overall Status:</strong> <span class="status-{status_class}">{details['qe_overall_status']}</span></p>
                            <p><strong>Progress Percentage:</strong> {details['qe_progress_percentage']}%</p>
                            <p><strong>Next Release Date:</strong> {details['next_release_date'] or 'N/A'}</p>
                            <h5>Current Week Progress Entry</h5>
                            <ul>
                                {"".join([f"<li>{line.strip()}</li>" for line in (details['current_week_progress_entry'] or '').splitlines() if line.strip()]) or "<li>No entry</li>"}
                            </ul>

                            <h4>QE Team & Resources</h4>
                            <p><strong>Team Size:</strong> {details['qe_team_size']}</p>
                            <h5>Current Week Task</h5>
                            <ul>
                                {"".join([f"<li>{line.strip()}</li>" for line in (details['qe_current_week_task'] or '').splitlines() if line.strip()]) or "<li>No tasks</li>"}
                            </ul>
                            <h5>Automation Tools Used</h5>
                            <ul>
                                {"".join([f"<li>{line.strip()}</li>" for line in (details['qe_automation_tools_used'] or '').splitlines() if line.strip()]) or "<li>None</li>"}
                            </ul>

                            <h4>Test Case Metrics</h4>
                            <p><strong>#TC Created:</strong> {details['tc_created']}</p>
                            <p><strong>#TC Executed:</strong> {details['tc_executed']}</p>
                            <p><strong>#TC Passed in First Round:</strong> {details['tc_passed_first_round']}</p>
                            <p><strong>Effort on TC Execution:</strong> {details['effort_tc_execution']} hours</p>
                            <p><strong>#TC Automated:</strong> {details['tc_automated']}</p>
                            <p><strong>Effort on TC Automation:</strong> {details['effort_tc_automation']} hours</p>

                            <h4>Defects & Quality Metrics</h4>
                            <p><strong>Defects Raised (Internal):</strong> {details['defects_raised_internal']}</p>
                            <p><strong>#SIT Defects:</strong> {details['sit_defects']}</p>
                            <p><strong>#UAT Defects:</strong> {details['uat_defects']}</p>
                            <p><strong>#Reopened Defects:</strong> {details['reopened_defects']}</p>
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
                                    st.markdown(f"### {pname}")
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        st.markdown(f"**Client**: {details['client']}")
                                        st.markdown(f"**Project SPOC**: {details['project_spoc']}")
                                        st.markdown(f"**Technology Used**: {details['technology_used']}")
                                    with col2:
                                        st.markdown(f"**Artifacts Link**: [{details['artifacts_link']}]({details['artifacts_link']})")
                                        st.markdown(f"**Overall Status**: {details['qe_overall_status']}")
                                        st.markdown(f"**Progress Percentage**: {details['qe_progress_percentage']}%")
                                    st.markdown(f"**Next Release Date**: {details['next_release_date'] or 'N/A'}")
                                    
                                    st.subheader("Current Week Progress Entry")
                                    for line in [l.strip() for l in (details['current_week_progress_entry'] or '').splitlines() if l.strip()]:
                                        st.markdown(f"- {line}")
                                    if not details['current_week_progress_entry']:
                                        st.markdown("- No entry")
                                    
                                    st.subheader("QE Team & Resources")
                                    st.markdown(f"**Team Size**: {details['qe_team_size']}")
                                    st.markdown("**Current Week Task**")
                                    for line in [l.strip() for l in (details['qe_current_week_task'] or '').splitlines() if l.strip()]:
                                        st.markdown(f"- {line}")
                                    if not details['qe_current_week_task']:
                                        st.markdown("- No tasks")
                                    st.markdown("**Automation Tools Used**")
                                    for line in [l.strip() for l in (details['qe_automation_tools_used'] or '').splitlines() if l.strip()]:
                                        st.markdown(f"- {line}")
                                    if not details['qe_automation_tools_used']:
                                        st.markdown("- None")
                                    
                                    st.subheader("Test Case Metrics")
                                    st.markdown(f"- **#TC Created**: {details['tc_created']}")
                                    st.markdown(f"- **#TC Executed**: {details['tc_executed']}")
                                    st.markdown(f"- **#TC Passed in First Round**: {details['tc_passed_first_round']}")
                                    st.markdown(f"- **Effort on TC Execution**: {details['effort_tc_execution']} hours")
                                    st.markdown(f"- **#TC Automated**: {details['tc_automated']}")
                                    st.markdown(f"- **Effort on TC Automation**: {details['effort_tc_automation']} hours")
                                    
                                    st.subheader("Defects & Quality Metrics")
                                    st.markdown(f"- **Defects Raised (Internal)**: {details['defects_raised_internal']}")
                                    st.markdown(f"- **#SIT Defects**: {details['sit_defects']}")
                                    st.markdown(f"- **#UAT Defects**: {details['uat_defects']}")
                                    st.markdown(f"- **#Reopened Defects**: {details['reopened_defects']}")
                                    st.markdown("---")

                        # Always provide download option
                        st.download_button(
                            label="📄 Download PDF Report",
                            data=pdf_data,
                            file_name=f"Weekly_QE_Report_{week_ending_date.strftime('%Y%m%d')}.pdf",
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