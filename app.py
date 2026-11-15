
import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from pathlib import Path
from datetime import datetime

# ==========================================
# APPLICATION CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="Saraswati Puja Food Manager",
    page_icon="🍱",
    layout="wide"
)

DB_PATH = Path(__file__).parent / "food_collection.db"

# ==========================================
# DATABASE CONNECTION
# ==========================================

def get_connection():
    return sqlite3.connect(DB_PATH)

def initialize_database():

    with get_connection() as conn:

        conn.execute("""
            CREATE TABLE IF NOT EXISTS students (

                student_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                class_name TEXT DEFAULT '',
                section TEXT DEFAULT '',
                food_status TEXT DEFAULT 'Not collected',
                updated_at TEXT

            )
        """)

        conn.commit()

initialize_database()

# ==========================================
# STATUS NORMALIZATION
# ==========================================

def normalize_status(value):

    value = str(value).strip().lower()

    if value in ["collected", "yes", "1", "true"]:
        return "Collected"

    return "Not collected"

# ==========================================
# EXCEL IMPORT FUNCTION
# ==========================================

def import_excel(uploaded_file):

    df = pd.read_excel(uploaded_file, dtype=str)

    df.columns = [
        str(col).strip().lower().replace(" ", "_")
        for col in df.columns
    ]

    if "student_id" not in df.columns or "name" not in df.columns:
        raise ValueError(
            "Excel must contain Student_ID and Name columns."
        )

    for column in ["class", "section", "food_status"]:
        if column not in df.columns:
            df[column] = ""

    df = df.fillna("")

    df["student_id"] = df["student_id"].astype(str).str.strip()
    df["name"] = df["name"].astype(str).str.strip()

    # Remove records with missing mandatory fields
    df = df[
        (df["student_id"] != "") &
        (df["name"] != "")
    ]

    # Prevent duplicate student IDs in the uploaded file
    df = df.drop_duplicates(
        subset=["student_id"],
        keep="last"
    )

    records = []

    for _, row in df.iterrows():

        records.append((
            row["student_id"],
            row["name"],
            row["class"],
            row["section"],
            normalize_status(row["food_status"]),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

    with get_connection() as conn:

        conn.executemany("""
            INSERT INTO students
            (
                student_id,
                name,
                class_name,
                section,
                food_status,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?)

            ON CONFLICT(student_id)
            DO UPDATE SET

                name = excluded.name,
                class_name = excluded.class_name,
                section = excluded.section

        """, records)

        conn.commit()

    return len(records)

# ==========================================
# FETCH DASHBOARD STATISTICS
# ==========================================

def get_statistics():

    with get_connection() as conn:

        total = conn.execute("""
            SELECT COUNT(*) FROM students
        """).fetchone()[0]

        collected = conn.execute("""
            SELECT COUNT(*) FROM students
            WHERE food_status = 'Collected'
        """).fetchone()[0]

        pending = total - collected

    return total, collected, pending

# ==========================================
# STUDENT SEARCH
# ==========================================

def search_students(keyword, status, limit, offset):

    query = """
        SELECT *
        FROM students
        WHERE (
            student_id LIKE ?
            OR name LIKE ?
            OR class_name LIKE ?
            OR section LIKE ?
        )
    """

    search_term = f"%{keyword}%"

    params = [
        search_term,
        search_term,
        search_term,
        search_term
    ]

    if status != "All":

        query += " AND food_status = ?"
        params.append(status)

    query += """
        ORDER BY name COLLATE NOCASE
        LIMIT ? OFFSET ?
    """

    params.extend([limit, offset])

    with get_connection() as conn:

        return pd.read_sql_query(
            query,
            conn,
            params=params
        )

def get_matching_count(keyword, status):

    search_term = f"%{keyword}%"

    query = """
        SELECT COUNT(*)
        FROM students
        WHERE (
            student_id LIKE ?
            OR name LIKE ?
            OR class_name LIKE ?
            OR section LIKE ?
        )
    """

    params = [
        search_term,
        search_term,
        search_term,
        search_term
    ]

    if status != "All":
        query += " AND food_status = ?"
        params.append(status)

    with get_connection() as conn:

        return conn.execute(
            query,
            params
        ).fetchone()[0]

# ==========================================
# UPDATE FOOD COLLECTION STATUS
# ==========================================

def update_status(student_id, current_status):

    new_status = (
        "Not collected"
        if current_status == "Collected"
        else "Collected"
    )

    with get_connection() as conn:

        conn.execute("""
            UPDATE students
            SET food_status = ?,
                updated_at = ?
            WHERE student_id = ?
        """, (
            new_status,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            student_id
        ))

        conn.commit()

# ==========================================
# EXPORT DATABASE TO EXCEL
# ==========================================

def export_excel():

    with get_connection() as conn:

        df = pd.read_sql_query("""
            SELECT
                student_id AS Student_ID,
                name AS Name,
                class_name AS Class,
                section AS Section,
                food_status AS Food_Status,
                updated_at AS Updated_At
            FROM students
            ORDER BY name COLLATE NOCASE
        """, conn)

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Food Collection"
        )

    return output.getvalue()

# ==========================================
# APPLICATION HEADER
# ==========================================

st.title("🍱 Saraswati Puja")
st.subheader("Student Food Collection Management System")

st.caption(
    "Manage student food distribution records efficiently."
)

st.divider()

# ==========================================
# SIDEBAR - EXCEL UPLOAD
# ==========================================

st.sidebar.title("📂 Data Management")

uploaded_file = st.sidebar.file_uploader(
    "Upload Student Excel File",
    type=["xlsx"]
)

if uploaded_file is not None:

    if st.sidebar.button(
        "Import Student Records",
        use_container_width=True
    ):

        try:

            count = import_excel(uploaded_file)

            st.sidebar.success(
                f"{count} records imported successfully!"
            )

            st.rerun()

        except Exception as error:

            st.sidebar.error(str(error))

st.sidebar.divider()

st.sidebar.info(
    "Existing students retain their saved collection status "
    "when the same Student IDs are imported again."
)

# ==========================================
# DASHBOARD
# ==========================================

total, collected, pending = get_statistics()

col1, col2, col3 = st.columns(3)

col1.metric(
    "Total Students",
    total
)

col2.metric(
    "Food Collected",
    collected
)

col3.metric(
    "Pending Collection",
    pending
)

if total > 0:

    progress = collected / total

    st.progress(
        progress,
        text=f"{progress * 100:.1f}% Food Distribution Completed"
    )

st.divider()

# ==========================================
# SEARCH AND FILTER
# ==========================================

st.subheader("🔍 Find Student")

col1, col2 = st.columns([3, 1])

with col1:

    keyword = st.text_input(
        "Search by Student ID, Name, Class or Section",
        placeholder="Enter student name or ID..."
    )

with col2:

    status_filter = st.selectbox(
        "Filter by Status",
        ["All", "Collected", "Not collected"]
    )

# ==========================================
# PAGINATION
# ==========================================

PAGE_SIZE = 50

matching_count = get_matching_count(
    keyword,
    status_filter
)

total_pages = max(
    1,
    (matching_count + PAGE_SIZE - 1) // PAGE_SIZE
)

page = st.number_input(
    "Page",
    min_value=1,
    max_value=total_pages,
    value=1,
    step=1
)

offset = (page - 1) * PAGE_SIZE

students = search_students(
    keyword,
    status_filter,
    PAGE_SIZE,
    offset
)

# ==========================================
# STUDENT RECORD DISPLAY
# ==========================================

st.subheader("📋 Student Records")

st.caption(
    f"Showing {len(students)} of {matching_count} matching records"
)

if students.empty:

    st.warning(
        "No student records found. Upload an Excel file first."
    )

else:

    for _, student in students.iterrows():

        student_id = student["student_id"]
        name = student["name"]
        class_name = student["class_name"]
        section = student["section"]
        current_status = student["food_status"]

        with st.container(border=True):

            col1, col2, col3 = st.columns(
                [3, 2, 2]
            )

            with col1:

                st.markdown(f"**{name}**")

                st.caption(
                    f"ID: {student_id}"
                )

            with col2:

                st.write(
                    f"Class: {class_name} - {section}"
                )

                if current_status == "Collected":

                    st.success("✓ Collected")

                else:

                    st.warning("⏳ Not collected")

            with col3:

                button_label = (
                    "↩ Mark Pending"
                    if current_status == "Collected"
                    else "✓ Mark Collected"
                )

                if st.button(
                    button_label,
                    key=f"update_{student_id}",
                    use_container_width=True
                ):

                    update_status(
                        student_id,
                        current_status
                    )

                    st.rerun()

# ==========================================
# EXCEL DOWNLOAD
# ==========================================

st.divider()

st.subheader("📥 Download Updated Records")

st.download_button(
    label="Download Updated Excel File",
    data=export_excel(),
    file_name="Saraswati_Puja_Food_Collection.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    disabled=(total == 0)
)

st.caption(
    "All collection status changes are saved automatically "
    "in the local SQLite database."
)