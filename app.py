import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict

st.set_page_config(page_title="QA Assignment Stage 4", layout="wide")
st.title("QA Assignment Script — Stage 4 📝")

# --- Upload file ---
file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if file is None:
    st.warning("Please upload an Excel file to proceed.")
    st.stop()

# --- Backlog mode ---
backlog_mode = st.radio("Are you expecting to be in backlog today?", ("No", "Yes")) == "Yes"

# --- Absentees ---
absent_input = st.text_input("Enter names of absentees (comma-separated), leave blank if none:")
absent_list = []
if absent_input:
    absent_list = [name.strip().title() for name in absent_input.split(",") if name.strip()]
    st.info(f"🟡 Absent today: {', '.join(absent_list)}")
else:
    st.success("✅ Everyone is present.")

# --- Custom limits ---
custom_input = st.text_input(
    "Any member with a specific product limit? (format: Name:Limit, comma-separated)"
)
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

# --- Load workbook ---
wb = load_workbook(file)
qa_ws = wb["QA"]
mp_ws = wb["MP"]

# --- Read MP preferences ---
preferences = {}
for row in mp_ws.iter_rows(min_row=2, values_only=True):
    name, divs = row[0], row[1]
    if name and divs:
        div_list = [d.strip().title() for d in divs.split(",") if d.strip()]
        preferences[name] = div_list

# --- Filter out absentees ---
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
priority_override_rows = []  # AG/AH numeric priority
normal_rows = []

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    assigned_to = row[0]
    division = str(row[17]).strip() if row[17] else ""
    m_value = row[12]
    brand = row[14]
    workflow = str(row[8]).strip() if row[8] else ""
    col_ag = row[32]
    col_ah = row[33]
    col_aq = row[42]

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
    st.success("🚀 Backlog mode OFF — normal assignment order.")

# --- Assignment trackers ---
assignments = {name: [] for name in team_members}
counts = {name: 0 for name in team_members}

# --- Calculate fair target ---
total_products = len([r for r in qa_ws.iter_rows(min_row=2, values_only=True) if r[12] is not None and str(r[12]).strip() != ""])
num_active = len(team_members)

if not custom_limits:
    even_split = total_products // num_active
    remainder = total_products % num_active
    custom_limits = {}
    for i, member in enumerate(team_members):
        limit = even_split + (1 if i < remainder else 0)
        custom_limits[member] = limit

DEFAULT_TARGET = max(custom_limits.values())

st.subheader("📏 Calculated per-member targets:")
for name, limit in custom_limits.items():
    st.write(f"  - {name}: {limit} products")

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

# --- Assign priority override rows ---
if priority_override_rows:
    st.subheader(f"🚨 Assigning {len(priority_override_rows)} AG/AH priority rows...")
    assign_rows(priority_override_rows)

# --- Assign preferred divisions ---
for member in team_members:
    prefs = active_preferences[member]
    for pref_div in prefs:
        for brand, rows in brand_rows.items():
            unassigned = [r for r in rows if qa_ws[f"A{r[0]}"].value in [None, ""] and r[1] == pref_div]
            if unassigned:
                assign_brand_block(member, unassigned)

# --- Assign remaining brands ---
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

# --- Convert formulas to values ---
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            cell.value = cell.value

# --- Save file ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")
st.download_button(
    "📥 Download Assigned File",
    data=open(output_path, "rb").read(),
    file_name=output_path,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("📊 Assignment Summary:")
for name, rows in assignments.items():
    limit = member_limit(name)
    st.write(f"  - {name}: {len(rows)} products (Limit: {limit})")
backlog_count = sum(1 for r in qa_rows if qa_ws[f"A{r[0]}"].value == "Backlog")
st.write(f"  - Backlog: {backlog_count} products")
