import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="QA Assignment", layout="wide")
st.title("📊 QA Assignment — Stage 4")

# --- File upload ---
uploaded_file = st.file_uploader("Upload QA Excel file", type=["xlsx"])
if uploaded_file is None:
    st.info("Please upload an Excel file to proceed.")
    st.stop()

# --- Load workbook ---
wb = load_workbook(uploaded_file)
qa_ws = wb["QA"]
mp_ws = wb["MP"]

# --- Streamlit inputs ---
backlog_mode = st.radio("Are you expecting to be in backlog today?", ["No", "Yes"]) == "Yes"

absent_input = st.text_input("Type names of absent members (comma-separated)").strip()
absent_list = [name.strip().title() for name in absent_input.split(",") if name.strip()]
if absent_list:
    st.warning(f"🟡 Absent today: {', '.join(absent_list)}")
else:
    st.success("✅ Everyone is present.")

custom_input = st.text_input("Specify custom product limits (Name:Limit, comma-separated)").strip()
custom_limits = {}
if custom_input:
    for entry in custom_input.split(","):
        if ":" in entry:
            name, limit = entry.split(":")
            name = name.strip().title()
            try:
                limit = int(limit.strip())
                custom_limits[name] = limit
            except ValueError:
                st.warning(f"⚠️ Invalid limit for {name}, ignoring.")

# --- Read preferences from MP sheet ---
preferences = {}
for row in mp_ws.iter_rows(min_row=2, values_only=True):
    name, divs = row[0], row[1]
    if name and divs:
        div_list = [d.strip().title() for d in divs.split(",") if d.strip()]
        preferences[name] = div_list

# --- Filter absentees ---
active_preferences = {name: divs for name, divs in preferences.items() if name not in absent_list}
team_members = list(active_preferences.keys())
num_members = len(team_members)
if num_members == 0:
    st.error("❌ No active team members available for assignment!")
    st.stop()

# --- Build QA data ---
qa_rows = []
brand_rows = defaultdict(list)
priority_rows = []
priority_override_rows = []
normal_rows = []

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    assigned_to = row[0]
    division = str(row[17]).strip() if row[17] else ""
    m_value = row[12]
    brand = row[14]
    workflow = str(row[8]).strip() if row[8] else ""
    col_ag = row[32]  # AG
    col_ah = row[33]  # AH
    col_aq = row[42]  # AQ (for backlog)

    # Priority override
    if isinstance(col_ag, (int, float)) or isinstance(col_ah, (int, float)):
        priority_override_rows.append((i, division, brand, workflow, col_aq))

    if m_value is not None and str(m_value).strip() != "":
        qa_rows.append((i, assigned_to, division, brand, workflow, col_aq))
        brand_rows[brand].append((i, division, workflow, col_aq))
        if workflow == "Prioritise in Workflow":
            priority_rows.append((i, division, brand, workflow, col_aq))
        else:
            normal_rows.append((i, division, brand, workflow, col_aq))

# --- Apply backlog sorting ---
if backlog_mode:
    st.info("🕐 Backlog mode ON — sorting all rows by earliest AQ date.")
    def sort_key(x):
        date_val = x[-1]
        return date_val if isinstance(date_val, datetime) else datetime.max
    qa_rows.sort(key=sort_key)
    priority_override_rows.sort(key=sort_key)
    priority_rows.sort(key=sort_key)
    normal_rows.sort(key=sort_key)
    for brand in brand_rows:
        brand_rows[brand].sort(key=sort_key)
else:
    st.success("🚀 Backlog mode OFF — assigning in normal order.")

# --- Assignment trackers ---
assignments = {name: [] for name in team_members}
counts = {name: 0 for name in team_members}
DEFAULT_TARGET = 100

def member_limit(member):
    return custom_limits.get(member, DEFAULT_TARGET)

def assign_rows(rows):
    for r, div, brand, workflow, *_ in rows:
        eligible = [m for m in team_members if counts[m] < member_limit(m)]
        if not eligible:
            qa_ws[f"A{r}"].value = "Backlog"
            continue
        chosen = min(eligible, key=lambda x: counts[x])
        qa_ws[f"A{r}"].value = chosen
        assignments[chosen].append(r)
        counts[chosen] += 1

def assign_brand_block(member, rows):
    remaining_capacity = member_limit(member) - counts[member]
    if remaining_capacity <= 0:
        return 0
    for r, div, workflow, *_ in rows[:remaining_capacity]:
        qa_ws[f"A{r}"].value = member
        assignments[member].append(r)
        counts[member] += 1
    return len(rows[:remaining_capacity])

# --- Stage 4 assignment logic ---
if priority_override_rows:
    st.info(f"🚨 Found {len(priority_override_rows)} AG/AH priority rows. Assigning first...")
    assign_rows(priority_override_rows)
else:
    st.success("✅ No AG/AH priority rows found.")

# Preferred divisions
for member in team_members:
    prefs = active_preferences[member]
    for pref_div in prefs:
        for brand, rows in brand_rows.items():
            unassigned = [r for r in rows if qa_ws[f"A{r[0]}"].value in [None, ""] and r[1] == pref_div]
            if unassigned:
                assign_brand_block(member, unassigned)

# Remaining brands
for brand, rows in brand_rows.items():
    unassigned = [r for r in rows if qa_ws[f"A{r[0]}"].value in [None, ""]]
    if not unassigned:
        continue
    eligible = [m for m in team_members if counts[m] < member_limit(m)]
    if eligible:
        chosen = min(eligible, key=lambda x: counts[x])
        assign_brand_block(chosen, unassigned)
    else:
        for r, div, workflow, *_ in unassigned:
            qa_ws[f"A{r}"].value = "Backlog"

# Convert formulas to values
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            cell.value = cell.value

# --- Prepare file for download ---
output_buffer = BytesIO()
wb.save(output_buffer)
output_buffer.seek(0)

st.download_button(
    label="📥 Download Assigned QA Excel",
    data=output_buffer,
    file_name=f"QA_Assigned_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Summary ---
st.subheader("📊 Summary of assignments:")
for name, rows in assignments.items():
    st.write(f"- {name}: {len(rows)} products (Limit: {member_limit(name)})")

backlog_count = sum(1 for r in qa_rows if qa_ws[f"A{r[0]}"].value == "Backlog")
st.write(f"- Backlog: {backlog_count} products")
