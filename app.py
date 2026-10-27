import hashlib
import secrets
import string
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st  # type: ignore[reportMissingImports]


# =========================================================
# CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Company Task Manager",
    page_icon="✅",
    layout="wide",
)

FILE_PATH = Path(__file__).resolve().parent / "data" / "TaskManager.xlsx"

TASK_STATUSES = ["To Do", "In Progress", "Review", "Completed"]
TASK_PRIORITIES = ["Low", "Medium", "High", "Urgent"]

TASK_COLUMNS = [
    "Task ID",
    "Task",
    "Description",
    "Project",
    "Assigned To",
    "Priority",
    "Status",
    "Due Date",
    "Created By",
    "Created At",
    "Updated At",
    "Status Changed At",
    "Completion Date",
    "Recurrence",
    "Recurrence Day",
    "Recurrence Start Date",
    "Recurrence End Date",
    "Recurring Task ID",
    "Occurrence Date",
]

USER_COLUMNS = [
    "User ID",
    "Name",
    "Email",
    "Department",
    "Role",
    "Reporting Manager",
    "Manager ID",
    "Active",
    "Password Hash",
    "First Login",
]


# =========================================================
# COMMON HELPERS
# =========================================================

def clean_text(value):
    """Convert Excel values to clean strings."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_column(df, column):
    """Ensure a column exists and contains clean strings."""
    if column not in df.columns:
        df[column] = ""
    df[column] = df[column].apply(clean_text)
    return df


def ensure_file_exists():
    if not FILE_PATH.exists():
        st.error(
            f"Excel file not found: {FILE_PATH}. "
            "Please check the data folder and workbook."
        )
        st.stop()


def read_sheet(sheet_name):
    ensure_file_exists()

    try:
        return pd.read_excel(FILE_PATH, sheet_name=sheet_name)
    except Exception as exc:
        st.error(f"Unable to read the '{sheet_name}' sheet: {exc}")
        st.stop()


def save_sheet(df, sheet_name):
    """Save one worksheet without changing the other worksheets."""
    ensure_file_exists()

    try:
        with pd.ExcelWriter(
            FILE_PATH,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        if sheet_name == "Tasks":
            load_tasks.clear()

    except PermissionError:
        st.error(
            "Excel workbook is currently locked. "
            "Close TaskManager.xlsx in Excel and try again."
        )
        st.stop()

    except Exception as exc:
        st.error(f"Unable to save the '{sheet_name}' sheet: {exc}")
        st.stop()


def load_users():
    users = read_sheet("Users")

    for column in USER_COLUMNS:
        if column not in users.columns:
            users[column] = ""

    for column in USER_COLUMNS:
        users[column] = (
            users[column].astype(object).fillna("").astype(str).str.strip()
        )

    return users


def save_users(users):
    save_sheet(users, "Users")


@st.cache_data
def load_tasks():
    df = pd.read_excel(FILE_PATH, sheet_name="Tasks")

    for column in TASK_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    for date_column in [
        "Due Date",
        "Recurrence Start Date",
        "Recurrence End Date",
        "Occurrence Date",
    ]:
        if date_column in df.columns:
            df[date_column] = pd.to_datetime(
                df[date_column], errors="coerce"
            ).dt.date

    return df


def save_tasks(tasks):
    save_sheet(tasks, "Tasks")


def normalize_assigned_ids(tasks, users):
    """
    Convert Assigned To values to User IDs.
    Supports both current User IDs and older employee names.
    """
    tasks = tasks.copy()

    if "Assigned To" not in tasks.columns:
        tasks["Assigned To"] = ""

    users = users.copy()

    users["User ID"] = users["User ID"].apply(clean_text)
    users["Name"] = users["Name"].apply(clean_text)

    valid_ids = set(users["User ID"])
    name_to_id = dict(zip(users["Name"], users["User ID"]))

    assigned_values = (
        tasks["Assigned To"].fillna("").astype(str).str.strip()
    )

    tasks["__Assigned ID"] = assigned_values.where(
        assigned_values.isin(valid_ids),
        assigned_values.map(name_to_id),
    ).fillna("")

    return tasks


def next_task_id(tasks):
    """Generate the next unused TASK-### identifier."""
    existing_ids = (
        tasks["Task ID"].fillna("").astype(str).str.strip()
        if "Task ID" in tasks.columns
        else pd.Series(dtype=str)
    )

    used_numbers = []

    for task_id in existing_ids:
        if task_id.upper().startswith("TASK-"):
            suffix = task_id[5:]
            if suffix.isdigit():
                used_numbers.append(int(suffix))

    next_number = max(used_numbers, default=0) + 1

    while f"TASK-{next_number:03d}" in set(existing_ids):
        next_number += 1

    return f"TASK-{next_number:03d}"


def next_recurring_id(tasks):
    """Generate the next recurring-task group identifier."""
    existing_ids = (
        tasks["Recurring Task ID"].fillna("").astype(str).str.strip()
        if "Recurring Task ID" in tasks.columns
        else pd.Series(dtype=str)
    )

    used_numbers = []
    for recurring_id in existing_ids:
        if recurring_id.upper().startswith("REC-"):
            suffix = recurring_id[4:]
            if suffix.isdigit():
                used_numbers.append(int(suffix))

    next_number = max(used_numbers, default=0) + 1
    while f"REC-{next_number:03d}" in set(existing_ids):
        next_number += 1

    return f"REC-{next_number:03d}"


def load_status_history():
    """Load the status audit sheet, creating it when needed."""
    ensure_file_exists()

    try:
        history = pd.read_excel(FILE_PATH, sheet_name="Status History")
    except ValueError:
        history = pd.DataFrame()
    except Exception as exc:
        st.error(f"Unable to read the 'Status History' sheet: {exc}")
        st.stop()

    history_columns = [
        "History ID",
        "Task ID",
        "Old Status",
        "New Status",
        "Changed By",
        "Changed At",
    ]

    for column in history_columns:
        if column not in history.columns:
            history[column] = ""

    return history[history_columns].copy()


def save_status_history(history):
    save_sheet(history, "Status History")


def next_history_id(history):
    existing_ids = (
        history["History ID"].fillna("").astype(str).str.strip()
        if "History ID" in history.columns
        else pd.Series(dtype=str)
    )

    used_numbers = []
    for history_id in existing_ids:
        if history_id.upper().startswith("HIST-"):
            suffix = history_id[5:]
            if suffix.isdigit():
                used_numbers.append(int(suffix))

    return f"HIST-{max(used_numbers, default=0) + 1:04d}"


def update_task_status(task_id, new_status, changed_by_id):
    """Update status, timestamps and write an audit-history record."""
    tasks = load_tasks().copy()

        # Fix pandas dtype issue for timestamp columns
    for col in ["Updated At", "Status Changed At", "Completion Date"]:
        if col not in tasks.columns:
            tasks[col] = ""
        tasks[col] = tasks[col].astype(object)

    history = load_status_history()

    
    tasks["Task ID"] = tasks["Task ID"].apply(clean_text)
    task_id = clean_text(task_id)
    new_status = clean_text(new_status)

    matching = tasks["Task ID"] == task_id

    if not matching.any():
        return False, f"Task ID {task_id} was not found."

    row_index = tasks.index[matching][0]
    old_status = clean_text(tasks.loc[row_index, "Status"])

    if old_status == new_status:
        return False, "No status change was made."

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tasks.loc[row_index, "Status"] = new_status
    tasks.loc[row_index, "Updated At"] = now
    tasks.loc[row_index, "Status Changed At"] = now

    # Keep the first completion timestamp. If an old completed task has no
    # Completion Date, use the current change time when it is marked completed.
    if new_status.casefold() == "completed":
        existing_completion = clean_text(tasks.loc[row_index, "Completion Date"])
        if not existing_completion:
            tasks.loc[row_index, "Completion Date"] = now

    save_tasks(tasks)

    history = load_status_history()
    history = pd.concat(
        [
            history,
            pd.DataFrame(
                [{
                    "History ID": next_history_id(history),
                    "Task ID": task_id,
                    "Old Status": old_status,
                    "New Status": new_status,
                    "Changed By": clean_text(changed_by_id),
                    "Changed At": now,
                }]
            ),
        ],
        ignore_index=True,
    )
    save_status_history(history)

    return True, "Status updated successfully."


def update_recurring_day(recurring_id, new_day):
    """Update the recurrence day for a recurring task series."""
    tasks = load_tasks().copy()
    recurring_id = clean_text(recurring_id)
    new_day = clean_text(new_day)
    if not recurring_id or not new_day:
        return False, "Please select a recurring task and recurrence day."
    mask = tasks["Recurring Task ID"].astype(str).str.strip() == recurring_id
    if not mask.any():
        return False, "Recurring task series not found."
    tasks.loc[mask, "Recurrence Day"] = new_day
    save_tasks(tasks)
    return True, f"Recurrence day updated to {new_day}."


def role_visible_tasks(tasks, users, user_id, user_name, role):
    """Enforce role-based task visibility, including full manager hierarchy."""

    tasks_with_ids = normalize_assigned_ids(tasks, users).copy()

    current_id = clean_text(user_id)
    current_name = clean_text(user_name).casefold()
    role_lower = clean_text(role).casefold()

    tasks_with_ids["__Assigned ID"] = (
        tasks_with_ids["__Assigned ID"].apply(clean_text)
    )

    if role_lower == "admin":
        return tasks_with_ids.drop(columns=["__Assigned ID"])

    users = users.copy()

    for column in ["User ID", "Name", "Role", "Manager ID",
                   "Reporting Manager"]:
        if column not in users.columns:
            users[column] = ""

        users[column] = users[column].apply(clean_text)

    # Employee: own tasks only
    if "manager" not in role_lower:
        return tasks_with_ids[
            tasks_with_ids["__Assigned ID"] == current_id
        ].drop(columns=["__Assigned ID"])

    # Build reporting hierarchy
    users["Role"] = users["Role"].str.casefold()

    name_to_ids = {}
    for _, row in users.iterrows():
        name = row["Name"].casefold()
        if name:
            name_to_ids.setdefault(name, []).append(row["User ID"])

    children = {}

    for _, row in users.iterrows():
        employee_id = row["User ID"]

        if not employee_id:
            continue

        manager_id = row["Manager ID"]

        if manager_id and manager_id != employee_id:
            children.setdefault(manager_id, set()).add(employee_id)

        reporting_name = row["Reporting Manager"].casefold()

        if reporting_name:
            for matching_manager_id in name_to_ids.get(reporting_name, []):
                if matching_manager_id and matching_manager_id != employee_id:
                    children.setdefault(
                        matching_manager_id, set()
                    ).add(employee_id)

    # Include direct and indirect reports
    allowed_ids = {current_id}
    queue = [current_id]

    while queue:
        manager_id = queue.pop()

        for employee_id in children.get(manager_id, set()):
            if employee_id not in allowed_ids:
                allowed_ids.add(employee_id)
                queue.append(employee_id)

    # Do not include Admin accounts in a Manager's team
    admin_ids = set(
        users.loc[users["Role"] == "admin", "User ID"]
    )
    allowed_ids -= admin_ids
    allowed_ids.add(current_id)

    return tasks_with_ids[
        tasks_with_ids["__Assigned ID"].isin(allowed_ids)
    ].drop(columns=["__Assigned ID"])



def make_employee_labels(users):
    """Create employee dropdown labels, including IDs for duplicate names."""
    users = users.copy()
    users["User ID"] = users["User ID"].apply(clean_text)
    users["Name"] = users["Name"].apply(clean_text)

    users = users[
        (users["User ID"] != "") & (users["Name"] != "")
    ].copy()

    name_counts = users["Name"].value_counts()
    labels = {}

    for _, employee in users.iterrows():
        employee_id = employee["User ID"]
        employee_name = employee["Name"]

        labels[employee_id] = (
            f"{employee_name} ({employee_id})"
            if name_counts[employee_name] > 1
            else employee_name
        )

    return users, labels


# =========================================================
# PASSWORD HELPERS
# =========================================================

def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password, password_hash):
    return hash_password(password) == clean_text(password_hash)


def generate_temporary_password(length=10):
    characters = string.ascii_letters + string.digits + "!@#$%"
    return "".join(secrets.choice(characters) for _ in range(length))


def authenticate_user(user_id, password):
    users = load_users()
    user_id = clean_text(user_id)

    matching = users[users["User ID"] == user_id]

    if matching.empty:
        return None

    user = matching.iloc[0]

    if clean_text(user["Active"]).lower() != "yes":
        return None

    password_hash = clean_text(user["Password Hash"])

    if not password_hash or password_hash.lower() == "nan":
        return None

    if verify_password(password, password_hash):
        return user

    return None


# =========================================================
# INITIAL ADMIN PASSWORD SETUP
# =========================================================

def setup_admin_password():
    if not FILE_PATH.exists():
        st.title("🔐 Initial Admin Setup")
        st.info(
            "The task workbook does not exist yet. Create the first Admin "
            "account to initialize it."
        )

        with st.form("initial_admin_setup"):
            admin_id = st.text_input("Admin User ID *")
            admin_name = st.text_input("Admin Name *")
            admin_email = st.text_input("Email")
            password = st.text_input("Create Admin Password *", type="password")
            confirm = st.text_input("Confirm Admin Password *", type="password")
            submitted = st.form_submit_button(
                "Create Workbook and Admin",
                width="stretch",
            )

        if submitted:
            admin_id = admin_id.strip()
            admin_name = admin_name.strip()

            if not admin_id:
                st.error("Please enter an Admin User ID.")
            elif not admin_name:
                st.error("Please enter the Admin name.")
            elif len(password) < 8:
                st.error("Password must contain at least 8 characters.")
            elif password != confirm:
                st.error("Passwords do not match.")
            else:
                if FILE_PATH.exists():
                    st.warning(
                        "A workbook was created while setup was open. "
                        "Refresh the app to continue."
                    )
                    st.rerun()

                admin = {
                    "User ID": admin_id,
                    "Name": admin_name,
                    "Email": admin_email.strip(),
                    "Department": "",
                    "Role": "Admin",
                    "Reporting Manager": "",
                    "Manager ID": "",
                    "Active": "Yes",
                    "Password Hash": hash_password(password),
                    "First Login": "No",
                }

                try:
                    FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
                    with pd.ExcelWriter(FILE_PATH, engine="openpyxl") as writer:
                        pd.DataFrame([admin], columns=USER_COLUMNS).to_excel(
                            writer, sheet_name="Users", index=False
                        )
                        pd.DataFrame(columns=TASK_COLUMNS).to_excel(
                            writer, sheet_name="Tasks", index=False
                        )
                        pd.DataFrame(columns=["Project"]).to_excel(
                            writer, sheet_name="Projects", index=False
                        )
                        pd.DataFrame(
                            columns=[
                                "History ID",
                                "Task ID",
                                "Old Status",
                                "New Status",
                                "Changed By",
                                "Changed At",
                            ]
                        ).to_excel(
                            writer, sheet_name="Status History", index=False
                        )
                except Exception as exc:
                    st.error(f"Unable to create the task workbook: {exc}")
                    st.stop()

                st.success("Workbook and Admin account created successfully.")
                st.rerun()

        st.stop()

    users = load_users()

    admin_users = users[
        (users["Role"].str.lower() == "admin")
        & (users["Active"].str.lower() == "yes")
    ]

    if admin_users.empty:
        return

    admin_index = admin_users.index[0]
    existing_hash = clean_text(users.loc[admin_index, "Password Hash"])

    if existing_hash and existing_hash.lower() != "nan":
        return

    st.title("🔐 Initial Admin Setup")
    st.info("Create a password for the first active Admin account.")

    st.write(f"**Admin:** {users.loc[admin_index, 'Name']}")
    st.write(f"**User ID:** {users.loc[admin_index, 'User ID']}")

    with st.form("admin_password_setup"):
        password = st.text_input("Create Admin Password", type="password")
        confirm = st.text_input("Confirm Password", type="password")

        submitted = st.form_submit_button(
            "🔐 Set Admin Password",
            width="stretch",
        )

    if submitted:
        if len(password) < 8:
            st.error("Password must contain at least 8 characters.")
        elif password != confirm:
            st.error("Passwords do not match.")
        else:
            users.loc[admin_index, "Password Hash"] = hash_password(password)
            users.loc[admin_index, "First Login"] = "No"
            save_users(users)
            st.success("Admin password created successfully!")
            st.rerun()

    st.stop()


# =========================================================
# SESSION STATE & LOGIN
# =========================================================

setup_admin_password()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "current_user" not in st.session_state:
    st.session_state.current_user = None

if "must_change_password" not in st.session_state:
    st.session_state.must_change_password = False

if "password_change_success" not in st.session_state:
    st.session_state.password_change_success = False

if "task_status_updated" not in st.session_state:
    st.session_state.task_status_updated = False


# ---------------- LOGIN PAGE ----------------

if not st.session_state.logged_in:

    if st.session_state.pop("password_change_success", False):
        st.success(
            "Password changed successfully! "
            "Please log in with your new password."
        )

    st.title("🔐 Company Task Manager")
    st.subheader("Login")

    with st.form("login_form"):
        login_user_id = st.text_input("User ID", placeholder="Example: EMP001")
        login_password = st.text_input("Password", type="password")

        submitted = st.form_submit_button(
            "🔑 Login",
            width="stretch",
        )

    if submitted:
        user = authenticate_user(login_user_id, login_password)

        if user is None:
            st.error("Invalid User ID or Password.")
        else:
            st.session_state.logged_in = True
            st.session_state.current_user = user
            st.session_state.must_change_password = (
                clean_text(user.get("First Login", "No")).lower() == "yes"
            )
            st.rerun()

    st.stop()


# ---------------- MANDATORY FIRST LOGIN PASSWORD CHANGE ----------------

if st.session_state.must_change_password:
    st.title("🔐 Change Your Temporary Password")

    st.warning(
        "You must create a new password before accessing the application."
    )

    with st.form("first_login_password_change"):
        new_password = st.text_input("New Password", type="password")
        confirm_password = st.text_input("Confirm New Password", type="password")

        submitted = st.form_submit_button(
            "🔒 Set New Password",
            width="stretch",
        )

    if submitted:
        if len(new_password) < 8:
            st.error("Password must contain at least 8 characters.")
        elif new_password != confirm_password:
            st.error("Passwords do not match.")
        else:
            current_user_id = clean_text(
                st.session_state.current_user["User ID"]
            )

            users = load_users()
            matching = users["User ID"] == current_user_id

            if not matching.any():
                st.error("Account not found. Please contact the administrator.")
                st.stop()

            user_index = users.index[matching][0]
            users.loc[user_index, "Password Hash"] = hash_password(new_password)
            users.loc[user_index, "First Login"] = "No"

            save_users(users)

            # Force a fresh login after changing the password.
            st.session_state.logged_in = False
            st.session_state.current_user = None
            st.session_state.must_change_password = False
            st.session_state.password_change_success = True
            st.rerun()

    st.stop()


# =========================================================
# CURRENT USER & SIDEBAR NAVIGATION
# =========================================================

current_user = st.session_state.current_user
current_user = st.session_state.current_user

current_user_id = clean_text(current_user.get("User ID", ""))
current_user_name = clean_text(current_user.get("Name", ""))
current_user_role = clean_text(current_user.get("Role", ""))

if current_user_role == "Admin":
    menu_options = [
        "Dashboard",
        "Create Task",
        "My Tasks",
        "Employees",
        "Projects",
        "Recurring Tasks",
        "Reports",
        "Change Password",
        "Password Reset",
    ]
elif current_user_role.strip().casefold() == "manager":
    menu_options = [
        "Dashboard",
        "Create Task",
        "My Tasks",
        "Reports",
        "Change Password",
    ]

else:
    menu_options = [
        "Dashboard",
        "My Tasks",
        "Reports",
    ]


st.sidebar.title("📋 Task Manager")
st.sidebar.write(f"👤 **{current_user_name}**")
st.sidebar.write(f"🔐 Role: **{current_user_role}**")
st.sidebar.divider()

page = st.sidebar.radio("Navigation", menu_options)

st.sidebar.divider()

if st.sidebar.button("🚪 Logout", width="stretch"):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.must_change_password = False
    st.rerun()


# =========================================================
# DASHBOARD
# =========================================================

if page == "Dashboard":
    st.title("✅ Company Task Manager")
    st.write(f"Today: **{date.today().strftime('%d %B %Y')}**")

    tasks = load_tasks()
    users = load_users()

    dashboard_tasks = role_visible_tasks(
        tasks,
        users,
        current_user_id,
        current_user_name,
        current_user_role,
    )

    today = date.today()

    dashboard_tasks["Due Date"] = pd.to_datetime(
        dashboard_tasks["Due Date"], errors="coerce"
    ).dt.date

    dashboard_tasks["Status"] = (
        dashboard_tasks["Status"].fillna("").astype(str).str.strip()
    )

    active_tasks = dashboard_tasks[
        dashboard_tasks["Status"].str.lower() != "completed"
    ]

    today_tasks = int((active_tasks["Due Date"] == today).sum())
    upcoming_tasks = int((active_tasks["Due Date"] > today).sum())
    overdue_tasks = int((active_tasks["Due Date"] < today).sum())

    completed_tasks = int(
        (dashboard_tasks["Status"].str.lower() == "completed").sum()
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📅 Due Today", today_tasks)
    col2.metric("🔜 Upcoming", upcoming_tasks)
    col3.metric("🔴 Overdue", overdue_tasks)
    col4.metric("✅ Completed", completed_tasks)

    st.divider()
    st.subheader("📋 Today's Tasks")

    today_df = active_tasks[active_tasks["Due Date"] == today]

    if today_df.empty:
        st.success("🎉 No pending tasks for today!")
    else:
        st.dataframe(today_df, width="stretch", hide_index=True)

    st.divider()
    st.subheader("🔴 Overdue Tasks")

    overdue_df = active_tasks[active_tasks["Due Date"] < today].sort_values(
        "Due Date"
    )

    if overdue_df.empty:
        st.success("🎉 No overdue tasks!")
    else:
        st.dataframe(overdue_df, width="stretch", hide_index=True)

    st.divider()
    st.subheader("🔜 Upcoming Tasks")

    upcoming_df = active_tasks[active_tasks["Due Date"] > today].sort_values(
        "Due Date"
    )

    if upcoming_df.empty:
        st.info("No upcoming tasks.")
    else:
        st.dataframe(upcoming_df, width="stretch", hide_index=True)




# =========================================================
# CREATE TASK
# =========================================================

elif page == "Create Task":

    st.title("➕ Create New Task")
    st.write("Create a task for yourself or another employee.")

    users = load_users()

    # Clean user data
    active_users = users[
        users["Active"].fillna("").astype(str).str.strip().str.lower() == "yes"
    ].copy()

    active_users["User ID"] = active_users["User ID"].apply(clean_text)
    active_users["Name"] = active_users["Name"].fillna("").astype(str).str.strip()
    active_users["Role"] = active_users["Role"].fillna("").astype(str).str.strip()
    active_users["Manager ID"] = active_users["Manager ID"].apply(clean_text)
    active_users["Reporting Manager"] = (
        active_users["Reporting Manager"].fillna("").astype(str).str.strip()
    )

    current_id = clean_text(current_user_id)
    current_name = current_user_name.strip().casefold()
    current_role = current_user_role.strip().casefold()

    # Manager: self + own eligible team members
    if current_role == "manager":

        is_self = active_users["User ID"] == current_id

        is_team_member = (
            (active_users["Manager ID"] == current_id)
            | (
                active_users["Reporting Manager"].str.casefold()
                == current_name
            )
        )

        role_lower = active_users["Role"].str.casefold()

        is_eligible_team = (
            is_team_member
            & (role_lower != "admin")
            & (~role_lower.str.contains("manager", na=False))
        )

        active_users = active_users[
            is_self | is_eligible_team
        ].copy()

    # Employee cannot create tasks
    elif current_role == "employee":
        st.warning("You do not have permission to create tasks.")
        st.stop()

    # Create dropdown labels
    active_users, employee_labels = make_employee_labels(active_users)
    employee_ids = active_users["User ID"].tolist()

    if not employee_ids:
        st.warning(
            "No active users found. Please check the Active column "
            "in the Users sheet."
        )

    else:
        with st.form("create_task_form"):

            task_name = st.text_input(
                "Task Name *",
                placeholder="Example: Prepare Amazon shipment report",
            )

            description = st.text_area(
                "Description",
                placeholder="Enter task details...",
            )

            col1, col2 = st.columns(2)

            with col1:
                project = st.text_input(
                    "Project",
                    placeholder="Example: Amazon Operations",
                )

                assigned_to = st.selectbox(
                    "Assigned To *",
                    employee_ids,
                    format_func=lambda emp_id: employee_labels[emp_id],
                )

                priority = st.selectbox("Priority", TASK_PRIORITIES)

            with col2:
                status = st.selectbox("Status", TASK_STATUSES)

                due_date = st.date_input(
                    "Due Date *",
                    value=date.today(),
                )

                st.text_input(
                    "Created By",
                    value=current_user_name,
                    disabled=True,
                )

            st.divider()
            st.subheader("🔁 Recurring Task")

            recurrence = st.selectbox(
                "Recurrence",
                ["None", "Weekly"],
                help="Weekly recurrence will create future occurrences automatically.",
            )

            recurrence_enabled = recurrence == "Weekly"

            recurrence_day_options = [
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
                "Sunday",
            ]

            default_day = recurrence_day_options[due_date.weekday()]

            recurrence_col1, recurrence_col2 = st.columns(2)

            with recurrence_col1:
                recurrence_day = st.selectbox(
                    "Recurrence Day",
                    recurrence_day_options,
                    index=recurrence_day_options.index(default_day),
                    disabled=not recurrence_enabled,
                    help="The first Due Date must fall on this weekday.",
                )

            with recurrence_col2:
                recurrence_start_date = st.date_input(
                    "Recurrence Start Date",
                    value=due_date,
                    disabled=not recurrence_enabled,
                )

                recurrence_end_date = st.date_input(
                    "Recurrence End Date",
                    value=due_date,
                    disabled=not recurrence_enabled,
                )

            submitted = st.form_submit_button(
                "💾 Save Task",
                width="stretch",
            )

        if submitted:
            if not task_name.strip():
                st.error("Please enter a Task Name.")
            elif recurrence_enabled and due_date.weekday() != recurrence_day_options.index(recurrence_day):
                st.error(
                    f"For a weekly recurring task, the Due Date must be a {recurrence_day}."
                )
            elif recurrence_enabled and recurrence_start_date > recurrence_end_date:
                st.error("Recurrence End Date cannot be earlier than Recurrence Start Date.")
            elif recurrence_enabled and not (recurrence_start_date <= due_date <= recurrence_end_date):
                st.error(
                    "The first Due Date must fall between the Recurrence Start Date and Recurrence End Date."
                )
            else:
                tasks = load_tasks().copy()
                task_id = next_task_id(tasks)
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                initial_completion_date = (
                    now if status.casefold() == "completed" else ""
                )

                recurring_id = next_recurring_id(tasks) if recurrence_enabled else ""
                occurrence_date = due_date if recurrence_enabled else ""

                new_task = {
                    "Task ID": task_id,
                    "Task": task_name.strip(),
                    "Description": description.strip(),
                    "Project": project.strip(),
                    "Assigned To": assigned_to,
                    "Priority": priority,
                    "Status": status,
                    "Due Date": due_date,
                    "Created By": current_user_id,
                    "Created At": now,
                    "Updated At": now,
                    "Status Changed At": "",
                    "Completion Date": initial_completion_date,
                    "Recurrence": recurrence,
                    "Recurrence Day": recurrence_day if recurrence_enabled else "",
                    "Recurrence Start Date": recurrence_start_date if recurrence_enabled else "",
                    "Recurrence End Date": recurrence_end_date if recurrence_enabled else "",
                    "Recurring Task ID": recurring_id,
                    "Occurrence Date": occurrence_date,
                }

                tasks = pd.concat(
                    [tasks, pd.DataFrame([new_task])],
                    ignore_index=True,
                )

                save_tasks(tasks)
                st.success(f"✅ {task_id} created successfully!")

# =========================================================
# MY TASKS
# =========================================================

elif page == "My Tasks":
    st.title("📋 My Tasks")

    tasks = load_tasks().copy()
    users = load_users()

    if "Assigned To" not in tasks.columns:
        st.error("Assigned To column is missing from the Tasks sheet.")
        st.stop()

    if "Task ID" not in tasks.columns:
        st.error("Task ID column is missing from the Tasks sheet.")
        st.stop()

    all_visible_tasks = role_visible_tasks(
        tasks,
        users,
        current_user_id,
        current_user_name,
        current_user_role,
    )

    if all_visible_tasks.empty:
        st.info("📭 No tasks found for your account.")
    else:
        filtered_tasks = all_visible_tasks.copy()

        filtered_tasks["Due Date"] = pd.to_datetime(
            filtered_tasks["Due Date"], errors="coerce"
        )

        for date_column in ["Status Changed At", "Completion Date"]:
            if date_column not in filtered_tasks.columns:
                filtered_tasks[date_column] = ""

        filtered_tasks["Status"] = (
            filtered_tasks["Status"].fillna("").astype(str).str.strip()
        )
        filtered_tasks["Priority"] = (
            filtered_tasks["Priority"].fillna("").astype(str).str.strip()
        )

        today = pd.Timestamp.today().normalize()

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            date_filter = st.selectbox(
                "📅 Date Filter",
                [
                    "All",
                    "Today",
                    "Tomorrow",
                    "This Week",
                    "This Month",
                    "Exact Date",
                    "Overdue",
                ],
            )

        with col2:
            available_statuses = sorted(
                filtered_tasks["Status"]
                .replace("", pd.NA)
                .dropna()
                .unique()
                .tolist()
            )
            status_filter = st.selectbox(
                "🔄 Status",
                ["All"] + available_statuses,
            )

        with col3:
            available_priorities = sorted(
                filtered_tasks["Priority"]
                .replace("", pd.NA)
                .dropna()
                .unique()
                .tolist()
            )
            priority_filter = st.selectbox(
                "⚡ Priority",
                ["All"] + available_priorities,
            )

        with col4:
            exact_date = st.date_input(
                "📆 Exact Date",
                value=date.today(),
                disabled=date_filter != "Exact Date",
            )

        # Date filtering
        if date_filter == "Today":
            filtered_tasks = filtered_tasks[
                filtered_tasks["Due Date"] == today
            ]
        elif date_filter == "Tomorrow":
            filtered_tasks = filtered_tasks[
                filtered_tasks["Due Date"] == today + pd.Timedelta(days=1)
            ]
        elif date_filter == "This Week":
            week_start = today - pd.Timedelta(days=today.weekday())
            week_end = week_start + pd.Timedelta(days=7)
            filtered_tasks = filtered_tasks[
                (filtered_tasks["Due Date"] >= week_start)
                & (filtered_tasks["Due Date"] < week_end)
            ]
        elif date_filter == "This Month":
            month_start = today.replace(day=1)
            next_month = month_start + pd.offsets.MonthBegin(1)
            filtered_tasks = filtered_tasks[
                (filtered_tasks["Due Date"] >= month_start)
                & (filtered_tasks["Due Date"] < next_month)
            ]
        elif date_filter == "Exact Date":
            selected_date = pd.Timestamp(exact_date)
            filtered_tasks = filtered_tasks[
                filtered_tasks["Due Date"] == selected_date
            ]
        elif date_filter == "Overdue":
            filtered_tasks = filtered_tasks[
                (filtered_tasks["Due Date"] < today)
                & (filtered_tasks["Status"].str.casefold() != "completed")
            ]

        if status_filter != "All":
            filtered_tasks = filtered_tasks[
                filtered_tasks["Status"] == status_filter
            ]

        if priority_filter != "All":
            filtered_tasks = filtered_tasks[
                filtered_tasks["Priority"] == priority_filter
            ]

        st.write(f"### Tasks: {len(filtered_tasks)}")

        if filtered_tasks.empty:
            st.info("No tasks match the selected filters.")
        else:
            # One table only: status is edited directly inside this table.
            editable_tasks = filtered_tasks.copy()

            # Display employee names in the UI while keeping User IDs in the
            # Tasks sheet/backend for reliable task ownership and updates.
            users_for_display = users.copy()
            users_for_display["User ID"] = users_for_display["User ID"].apply(clean_text)
            users_for_display["Name"] = users_for_display["Name"].apply(clean_text)

            id_to_name = dict(
                zip(
                    users_for_display["User ID"],
                    users_for_display["Name"],
                )
            )

            for employee_column in ["Assigned To", "Created By"]:
                if employee_column in editable_tasks.columns:
                    editable_tasks[employee_column] = (
                        editable_tasks[employee_column]
                        .apply(clean_text)
                        .map(lambda value: id_to_name.get(value, value))
                    )

            editable_tasks["Due Date"] = pd.to_datetime(
                editable_tasks["Due Date"], errors="coerce"
            ).dt.date

            table_columns = [
                col for col in [
                    "Task ID",
                    "Task",
                    "Description",
                    "Project",
                    "Assigned To",
                    "Priority",
                    "Status",
                    "Due Date",
                    "Created By",
                    "Created At",
                    "Updated At",
                    "Status Changed At",
                    "Completion Date",
                ]
                if col in editable_tasks.columns
            ]

            editable_tasks = editable_tasks[table_columns].copy()
            editable_tasks = editable_tasks.drop_duplicates("Task ID")

            edited_tasks = st.data_editor(
                editable_tasks,
                width="stretch",
                hide_index=True,
                disabled=[
                    col for col in table_columns
                    if col not in {"Status"}
                ],
                column_config={
                    "Status": st.column_config.SelectboxColumn(
                        "Status",
                        options=TASK_STATUSES,
                        required=True,
                    ),
                    "Due Date": st.column_config.DateColumn(
                        "Due Date",
                        format="DD/MM/YYYY",
                    ),
                },
                key="my_tasks_editor",
            )

            original_statuses = (
                editable_tasks.set_index("Task ID")["Status"]
                .fillna("")
                .astype(str)
                .str.strip()
                .to_dict()
            )

            edited_statuses = (
                edited_tasks.set_index("Task ID")["Status"]
                .fillna("")
                .astype(str)
                .str.strip()
                .to_dict()
            )

            changes = [
                (task_id, new_status)
                for task_id, new_status in edited_statuses.items()
                if task_id in original_statuses
                and new_status != original_statuses[task_id]
            ]

            if changes:
                updated_count = 0
                errors = []

                for task_id, new_status in changes:
                    saved, message = update_task_status(
                        task_id,
                        new_status,
                        current_user_id,
                    )
                    if saved:
                        updated_count += 1
                    else:
                        errors.append(f"{task_id}: {message}")

                if errors:
                    for error in errors:
                        st.error(error)

                if updated_count:
                    st.success(
                        f"✅ {updated_count} task status change(s) saved successfully."
                    )
                    st.rerun()

            st.caption(
                "Status changes are saved automatically. Each change is recorded "
                "in the Status History sheet with the user and exact timestamp."
            )


# =========================================================
# EMPLOYEES
# =========================================================

elif page == "Employees":
    st.title("👥 Employees")
    
    # Only Admin can manage employees
    if current_user_role.strip().lower() != "admin":
        st.error("Access denied. Only Admin can add or manage employees.")
        st.stop()

    st.write("Manage employees, roles and reporting structure.")

    users = load_users()

    # ---------------- UPDATE MANAGER IDs ----------------

    st.subheader("🔗 Manager Mapping")

    if st.button("🔄 Update Manager IDs"):
        manager_rows = users[
            users["Role"].str.lower().apply(
                lambda value: value == "admin" or "manager" in value
            )
        ].copy()

        manager_names = manager_rows["Name"].str.casefold()
        duplicate_names = manager_names[
            manager_names.ne("")
        ].duplicated(keep=False)

        if duplicate_names.any():
            st.error(
                "Duplicate manager names found. "
                "Please make manager names unique before mapping."
            )
        else:
            manager_map = dict(
                zip(manager_rows["Name"].str.casefold(), manager_rows["User ID"])
            )

            updated_users = users.copy()
            unresolved = []

            for idx, employee in updated_users.iterrows():
                role_value = clean_text(employee.get("Role", "")).lower()
                employee_id = clean_text(employee.get("User ID", ""))

                if role_value == "admin" or "manager" in role_value:
                    updated_users.loc[idx, "Manager ID"] = employee_id
                    continue

                reporting_name = clean_text(
                    employee.get("Reporting Manager", "")
                )

                if not reporting_name or reporting_name.casefold() in {
                    "nan", "none"
                }:
                    updated_users.loc[idx, "Manager ID"] = ""
                elif reporting_name.casefold() in manager_map:
                    updated_users.loc[idx, "Manager ID"] = (
                        manager_map[reporting_name.casefold()]
                    )
                else:
                    unresolved.append(
                        f"{employee.get('Name', '')}: {reporting_name}"
                    )

            if unresolved:
                st.error(
                    "Some Reporting Manager names could not be matched. "
                    "No changes were saved."
                )
                st.write(unresolved)
            else:
                save_users(updated_users)
                users = updated_users
                st.success("Manager IDs updated successfully!")

    st.divider()
    st.subheader("➕ Add Employee")

    active_names = users[
        users["Active"].str.lower() == "yes"
    ]["Name"].dropna().astype(str).tolist()

    with st.form("add_employee_form"):
        col1, col2 = st.columns(2)

        with col1:
            employee_id = st.text_input(
                "Employee ID *",
                placeholder="Example: EMP001",
            )
            employee_name = st.text_input(
                "Employee Name *",
                placeholder="Example: Gaurav Modi",
            )
            email = st.text_input("Email", placeholder="employee@company.com")
            department = st.text_input(
                "Department",
                placeholder="Example: Operations",
            )

        with col2:
            employee_role = st.selectbox(
                "Role *",
                ["Admin", "Manager", "Employee"],
            )
            reporting_manager = st.selectbox(
                "Reporting Manager",
                ["None"] + sorted(set(active_names)),
            )
            active = st.selectbox("Status", ["Yes", "No"])

        submitted = st.form_submit_button(
            "💾 Add Employee",
            width="stretch",
        )

    if submitted:
        employee_id = employee_id.strip()
        employee_name = employee_name.strip()

        if not employee_id:
            st.error("Please enter Employee ID.")
        elif not employee_name:
            st.error("Please enter Employee Name.")
        elif employee_id in users["User ID"].tolist():
            st.error("This Employee ID already exists.")
        else:
            temporary_password = generate_temporary_password()

            new_employee = {
                "User ID": employee_id,
                "Name": employee_name,
                "Email": email.strip(),
                "Department": department.strip(),
                "Role": employee_role,
                "Reporting Manager": (
                    "" if reporting_manager == "None" else reporting_manager
                ),
                "Manager ID": "",
                "Active": active,
                "Password Hash": hash_password(temporary_password),
                "First Login": "Yes",
            }

            users = pd.concat(
                [users, pd.DataFrame([new_employee])],
                ignore_index=True,
            )

            save_users(users)

            st.success(f"✅ {employee_name} added successfully!")
            st.write("**Temporary password:**")
            st.code(temporary_password)
            st.warning(
                "Share this password securely. The employee must change "
                "it after their first login."
            )

    st.divider()
    st.subheader("👥 Employee List")

    employee_list = users.copy()

    if current_user_role == "Manager":
        employee_list = employee_list[
            employee_list["Reporting Manager"].str.casefold()
            == current_user_name.casefold()
        ]

    employee_columns = [
        column for column in [
            "User ID",
            "Name",
            "Email",
            "Department",
            "Role",
            "Reporting Manager",
            "Manager ID",
            "Active",
        ]
        if column in employee_list.columns
    ]

    if employee_list.empty:
        st.info("No employees have been added yet.")
    else:
        st.dataframe(
            employee_list[employee_columns],
            width="stretch",
            hide_index=True,
        )


# =========================================================
# PROJECTS
# =========================================================

elif page == "Projects":
    st.title("📁 Projects")

    projects = read_sheet("Projects")

    st.dataframe(
        projects,
        width="stretch",
        hide_index=True,
    )


# =========================================================
# REPORTS - ROLE-BASED HRMS ANALYTICS
# =========================================================

elif page == "Recurring Tasks":
    st.title("🔁 Recurring Tasks")
    st.caption("Only Admin can edit the recurrence day after a recurring task has been saved.")

    tasks = load_tasks().copy()
    recurring_tasks = tasks[
        tasks["Recurrence"].astype(str).str.strip().str.casefold() == "weekly"
    ].copy() if "Recurrence" in tasks.columns else pd.DataFrame()

    if recurring_tasks.empty:
        st.info("No weekly recurring tasks found.")
    else:
        recurring_ids = [
            rid for rid in recurring_tasks["Recurring Task ID"].dropna().astype(str).str.strip().unique().tolist()
            if rid
        ]
        if not recurring_ids:
            st.info("No recurring task series found.")
        else:
            def recurring_label(rid):
                row = recurring_tasks[recurring_tasks["Recurring Task ID"].astype(str).str.strip() == rid].iloc[0]
                return f"{rid} — {clean_text(row.get('Task', ''))} — {clean_text(row.get('Assigned To', ''))}"

            selected_recurring_id = st.selectbox("Recurring Task", recurring_ids, format_func=recurring_label)
            selected_rows = recurring_tasks[recurring_tasks["Recurring Task ID"].astype(str).str.strip() == selected_recurring_id]
            current_day = clean_text(selected_rows.iloc[0].get("Recurrence Day", "Monday")) or "Monday"
            day_options = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            if current_day not in day_options:
                current_day = "Monday"

            new_recurrence_day = st.selectbox(
                "Recurrence Day", day_options, index=day_options.index(current_day),
                help="This controls future occurrences for this recurring task series.",
            )

            if st.button("💾 Save Recurrence Day", type="primary"):
                ok, message = update_recurring_day(selected_recurring_id, new_recurrence_day)
                if ok:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)

            st.divider()
            display_recurring = recurring_tasks.copy()
            users_for_display = load_users()
            id_to_name = dict(zip(users_for_display["User ID"].astype(str).str.strip(), users_for_display["Name"].astype(str).str.strip()))
            display_recurring["Assigned To"] = display_recurring["Assigned To"].map(lambda x: id_to_name.get(str(x).strip(), str(x).strip()))
            st.dataframe(
                display_recurring[["Recurring Task ID", "Task", "Assigned To", "Recurrence Day", "Recurrence Start Date", "Recurrence End Date"]]
                .drop_duplicates(subset=["Recurring Task ID"]),
                width="stretch", hide_index=True,
            )


elif page == "Reports":

    st.title("📊 Reports Dashboard")
    st.caption("Employee productivity, task completion and delay analytics")

    tasks = load_tasks().copy()
    users = load_users().copy()

    if tasks.empty:
        st.info("No task data available.")
        st.stop()

    # -----------------------------------------------------
    # CLEAN USER DATA
    # -----------------------------------------------------

    users["User ID"] = users["User ID"].apply(clean_text)
    users["Name"] = users["Name"].fillna("").astype(str).str.strip()
    users["Role"] = users["Role"].fillna("").astype(str).str.strip()

    if "Manager ID" in users.columns:
        users["Manager ID"] = users["Manager ID"].apply(clean_text)
    else:
        users["Manager ID"] = ""

    if "Reporting Manager" in users.columns:
        users["Reporting Manager"] = (
            users["Reporting Manager"].fillna("").astype(str).str.strip()
        )
    else:
        users["Reporting Manager"] = ""

    current_id = clean_text(current_user_id)
    current_name = current_user_name.strip()
    role = current_user_role.strip().casefold()

    # -----------------------------------------------------
    # BUILD REPORTING HIERARCHY
    # Manager can see direct and indirect reports
    # -----------------------------------------------------

    id_to_name = dict(zip(users["User ID"], users["Name"]))

    name_to_ids = {}
    for _, user_row in users.iterrows():
        user_name = user_row["Name"].casefold()
        if user_name:
            name_to_ids.setdefault(user_name, []).append(
                user_row["User ID"]
            )

    # Build each manager's direct-report list
    children = {}

    for _, user_row in users.iterrows():

        employee_id = user_row["User ID"]

        manager_id = user_row["Manager ID"]
        if manager_id and manager_id != employee_id:
            children.setdefault(manager_id, set()).add(employee_id)

        # Name-based fallback for older user records
        reporting_name = user_row["Reporting Manager"].casefold()

        if reporting_name:
            matching_manager_ids = name_to_ids.get(reporting_name, [])

            for manager_id in matching_manager_ids:
                if manager_id and manager_id != employee_id:
                    children.setdefault(manager_id, set()).add(
                        employee_id
                    )

    # Find all direct and indirect reports
    visible_ids = {current_id}
    pending_ids = [current_id]

    while pending_ids:

        manager_id = pending_ids.pop()

        for employee_id in children.get(manager_id, set()):

            if employee_id not in visible_ids:
                visible_ids.add(employee_id)
                pending_ids.append(employee_id)

    # -----------------------------------------------------
    # ROLE-BASED ACCESS
    # -----------------------------------------------------

    if role == "admin":

        permitted_users = users.copy()

    elif role == "manager":

        permitted_users = users[
            users["User ID"].isin(visible_ids)
        ].copy()

    else:

        permitted_users = users[
            users["User ID"] == current_id
        ].copy()

    if permitted_users.empty:
        st.warning("No employee records found for your account.")
        st.stop()

    # -----------------------------------------------------
    # NORMALIZE TASK ASSIGNMENT IDs
    # -----------------------------------------------------

    report_tasks = normalize_assigned_ids(tasks, users).copy()

    report_tasks["__Assigned ID"] = (
        report_tasks["__Assigned ID"].apply(clean_text)
    )

    # Restrict tasks to permitted employees
    permitted_ids = set(permitted_users["User ID"])

    report_tasks = report_tasks[
        report_tasks["__Assigned ID"].isin(permitted_ids)
    ].copy()

    if report_tasks.empty:
        st.info("No reportable tasks found for your access level.")
        st.stop()

    # -----------------------------------------------------
    # PREPARE TASK FIELDS
    # -----------------------------------------------------

    for col in ["Task", "Status", "Priority"]:
        if col not in report_tasks.columns:
            report_tasks[col] = ""

    report_tasks["Status"] = (
        report_tasks["Status"].fillna("").astype(str).str.strip()
    )

    report_tasks["Priority"] = (
        report_tasks["Priority"].fillna("Unknown").astype(str).str.strip()
    )

    report_tasks["Employee"] = (
        report_tasks["__Assigned ID"].map(id_to_name)
        .fillna(report_tasks["__Assigned ID"])
    )

    # Parse dates safely
    for date_col in ["Created At", "Due Date", "Updated At", "Status Changed At", "Completion Date"]:
        if date_col not in report_tasks.columns:
            report_tasks[date_col] = pd.NaT

        report_tasks[date_col] = pd.to_datetime(
            report_tasks[date_col],
            errors="coerce"
        )

    # Convert timezone-aware timestamps to local-naive timestamps
    for date_col in ["Created At", "Due Date", "Updated At", "Status Changed At", "Completion Date"]:
        try:
            if report_tasks[date_col].dt.tz is not None:
                report_tasks[date_col] = (
                    report_tasks[date_col].dt.tz_localize(None)
                )
        except (AttributeError, TypeError):
            pass

    today = pd.Timestamp(date.today())

    # -----------------------------------------------------
    # REPORT FILTERS
    # -----------------------------------------------------

    st.subheader("🔎 Report Filters")

    col1, col2 = st.columns(2)

    with col1:
        period = st.selectbox(
            "Reporting Period",
            [
                "Daily",
                "Weekly",
                "Monthly",
                "Last 3 Months",
                "All Time",
            ],
            index=1,
        )

    permitted_users["Employee Label"] = (
        permitted_users["Name"]
        + " ("
        + permitted_users["User ID"]
        + ")"
    )

    employee_options = ["All Employees"] + (
        permitted_users["Employee Label"].tolist()
    )

    with col2:
        selected_employee = st.selectbox(
            "Employee Name",
            employee_options,
        )

    # Apply date range using task creation date
    if period == "Daily":
        start_date = today
        end_date = today + pd.Timedelta(days=1)

    elif period == "Weekly":
        start_date = today - pd.Timedelta(days=today.weekday())
        end_date = start_date + pd.Timedelta(days=7)

    elif period == "Monthly":
        start_date = today.replace(day=1)
        end_date = start_date + pd.DateOffset(months=1)

    elif period == "Last 3 Months":
        start_date = today - pd.DateOffset(months=3)
        end_date = today + pd.Timedelta(days=1)

    else:
        start_date = None
        end_date = None

    if start_date is not None:
        report_tasks = report_tasks[
            (report_tasks["Created At"] >= start_date)
            & (report_tasks["Created At"] < end_date)
        ].copy()

    # Filter selected employee
    if selected_employee != "All Employees":

        selected_row = permitted_users[
            permitted_users["Employee Label"] == selected_employee
        ]

        if not selected_row.empty:
            selected_id = selected_row.iloc[0]["User ID"]

            report_tasks = report_tasks[
                report_tasks["__Assigned ID"] == selected_id
            ].copy()

    st.caption(
        f"Period: {period} | Employee: {selected_employee}"
    )

    if report_tasks.empty:
        st.info("No tasks found for the selected filters.")
        st.stop()

    # -----------------------------------------------------
    # TASK STATUS CLASSIFICATION
    # -----------------------------------------------------

    status_lower = report_tasks["Status"].str.casefold()

    completed_mask = status_lower.isin(
        ["completed", "complete", "done", "closed"]
    )

    # Every task not marked completed is treated as open
    pending_mask = ~completed_mask

    completed_tasks = report_tasks[completed_mask].copy()
    pending_tasks = report_tasks[pending_mask].copy()

    total_tasks = len(report_tasks)
    completed_count = len(completed_tasks)
    pending_count = len(pending_tasks)

    # Delayed completed tasks: use the recorded completion date.
    # Fall back to Updated At for older tasks created before Completion Date
    # was introduced.
    completed_tasks["__Completion Check"] = completed_tasks["Completion Date"]
    completed_tasks["__Completion Check"] = completed_tasks["__Completion Check"].fillna(
        completed_tasks["Updated At"]
    )

    delayed_completed = completed_tasks[
        completed_tasks["Due Date"].notna()
        & completed_tasks["__Completion Check"].notna()
        & (
            completed_tasks["__Completion Check"].dt.normalize()
            > completed_tasks["Due Date"].dt.normalize()
        )
    ]

    delayed_count = len(delayed_completed)

    completion_pct = (
        completed_count / total_tasks * 100
        if total_tasks else None
    )

    pending_pct = (
        pending_count / total_tasks * 100
        if total_tasks else None
    )

    delay_pct = (
        delayed_count / completed_count * 100
        if completed_count else None
    )

    def show_pct(value):
        return f"{value:.1f}%" if value is not None else "N/A"

    # -----------------------------------------------------
    # KPI CARDS
    # -----------------------------------------------------

    st.subheader("📌 Key Performance Indicators")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Total Tasks", total_tasks)

    c2.metric(
        "Completion %",
        show_pct(completion_pct),
        f"{completed_count} completed",
    )

    c3.metric(
        "Pending %",
        show_pct(pending_pct),
        f"{pending_count} open",
    )

    c4.metric(
        "Delay %",
        show_pct(delay_pct),
        f"{delayed_count} completed late",
    )

    st.caption(
        "Completion % = completed ÷ total. "
        "Pending % = open ÷ total. "
        "Delay % = completed after due date ÷ completed tasks. "
        "Tasks without a verifiable completion date are not counted as delayed."
    )

    # -----------------------------------------------------
    # STATUS AND PRIORITY CHARTS
    # -----------------------------------------------------

    st.divider()
    st.subheader("📊 Task Analysis")

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Tasks by Status**")
        status_counts = report_tasks["Status"].replace(
            "", "Unknown"
        ).value_counts()
        st.bar_chart(status_counts)

    with col2:
        st.write("**Tasks by Priority**")
        priority_counts = report_tasks["Priority"].replace(
            "", "Unknown"
        ).value_counts()
        st.bar_chart(priority_counts)

    # -----------------------------------------------------
    # EMPLOYEE-WISE COMPARISON
    # -----------------------------------------------------

    st.subheader("👥 Employee-wise Performance")

    employee_summary = report_tasks.groupby("Employee").agg(
        Total_Tasks=("Task", "size"),
        Completed_Tasks=(
            "Status",
            lambda s: s.str.casefold().isin(
                ["completed", "complete", "done", "closed"]
            ).sum(),
        ),
    )

    employee_summary["Completion %"] = (
        employee_summary["Completed_Tasks"]
        / employee_summary["Total_Tasks"] * 100
    ).round(1)

    st.bar_chart(
        employee_summary[
            ["Total_Tasks", "Completed_Tasks"]
        ]
    )

    st.dataframe(
        employee_summary.reset_index(),
        use_container_width=True,
        hide_index=True,
    )

    # -----------------------------------------------------
    # TASK CREATION TREND
    # -----------------------------------------------------

    st.subheader("📈 Task Creation Trend")

    dated_tasks = report_tasks.dropna(
        subset=["Created At"]
    ).copy()

    if not dated_tasks.empty:

        dated_tasks["Creation Date"] = (
            dated_tasks["Created At"].dt.normalize()
        )

        daily_trend = (
            dated_tasks.groupby("Creation Date")
            .size()
            .rename("Tasks Created")
        )

        if period == "Daily":
            st.bar_chart(daily_trend)

        elif period == "Weekly":
            weekly_trend = (
                dated_tasks.set_index("Created At")
                .resample("W-MON")
                .size()
                .rename("Tasks Created")
            )
            st.line_chart(weekly_trend)

        elif period == "Monthly":
            monthly_trend = (
                dated_tasks.set_index("Created At")
                .resample("MS")
                .size()
                .rename("Tasks Created")
            )
            st.line_chart(monthly_trend)

        elif period == "Last 3 Months":
            monthly_trend = (
                dated_tasks.set_index("Created At")
                .resample("MS")
                .size()
                .rename("Tasks Created")
            )
            st.line_chart(monthly_trend)

        else:
            st.line_chart(daily_trend)

    else:
        st.info("No valid task creation dates available for the trend chart.")

    # -----------------------------------------------------
    # TASK DETAIL TABLE
    # -----------------------------------------------------

    st.subheader("📋 Task Details")

    display_columns = [
        col for col in [
            "Task ID",
            "Task",
            "Employee",
            "Project",
            "Status",
            "Priority",
            "Due Date",
            "Created At",
            "Completion Date",
            "Status Changed At",
        ]
        if col in report_tasks.columns
    ]

    st.dataframe(
        report_tasks[display_columns],
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# CHANGE PASSWORD
# =========================================================

elif page == "Change Password":
    st.title("🔐 Change Password")
    st.write("Update your account password.")

    with st.form("change_password_form"):
        current_password = st.text_input("Current Password", type="password")
        new_password = st.text_input("New Password", type="password")
        confirm_password = st.text_input(
            "Confirm New Password",
            type="password",
        )

        submitted = st.form_submit_button("Update Password")

    if submitted:
        if not current_password or not new_password or not confirm_password:
            st.error("Please fill in all fields.")
        elif len(new_password) < 8:
            st.error("Your new password must contain at least 8 characters.")
        elif new_password != confirm_password:
            st.error("The new passwords do not match.")
        elif current_password == new_password:
            st.error(
                "Your new password must be different from your current password."
            )
        else:
            users = load_users()
            matching = users["User ID"] == current_user_id

            if not matching.any():
                st.error("User account not found. Contact your administrator.")
            else:
                user_index = users.index[matching][0]
                stored_hash = clean_text(
                    users.loc[user_index, "Password Hash"]
                )

                if not verify_password(current_password, stored_hash):
                    st.error("Your current password is incorrect.")
                else:
                    users.loc[user_index, "Password Hash"] = hash_password(
                        new_password
                    )
                    save_users(users)

                    # Log the user out so the new credentials are used next time.
                    st.session_state.logged_in = False
                    st.session_state.current_user = None
                    st.session_state.must_change_password = False
                    st.session_state.password_change_success = True
                    st.rerun()


# =========================================================
# PASSWORD RESET
# =========================================================

elif page == "Password Reset":
    st.title("🔑 Password Reset")

    if current_user_role.strip().casefold() != "admin":
        st.error("Access denied. Only Admin can reset passwords.")
        st.stop()

    st.write("Generate a temporary password for an employee.")

    users = load_users()

    eligible_users = users[
        (users["User ID"] != current_user_id)
        & (users["Role"].str.lower() != "admin")
    ].copy()

    employee_options = {}

    for _, user in eligible_users.iterrows():
        user_id = clean_text(user["User ID"])
        user_name = clean_text(user["Name"])

        if user_id:
            employee_options[f"{user_name} ({user_id})"] = user_id

    if not employee_options:
        st.warning("No eligible employees found.")
    else:
        selected_employee = st.selectbox(
            "Select Employee",
            list(employee_options.keys()),
        )

        if st.button("Generate Temporary Password"):
            selected_user_id = employee_options[selected_employee]

            matching = users["User ID"] == selected_user_id

            if not matching.any():
                st.error("Employee not found.")
            else:
                user_index = users.index[matching][0]
                temporary_password = secrets.token_urlsafe(9)

                users.loc[user_index, "Password Hash"] = hash_password(
                    temporary_password
                )
                users.loc[user_index, "First Login"] = "Yes"

                save_users(users)

                st.success("Temporary password generated.")
                st.code(temporary_password)
                st.warning(
                    "Share this password securely. The employee must "
                    "change it after logging in."
                )