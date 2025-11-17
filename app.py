import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="QA Assignment Tool", layout="wide")

st.title("📊 QA Assignment Tool")

# --- Upload Excel file ---
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file is None:
    st.warning("Please upload an Excel file to proceed.")
    st.stop()

# --- Ask for backlog mode ---
backlog_mode = st.checkbox("Backlog mode?", value=False, help="Sort by earliest AQ date if enabled.")

# --- Input active members ---
st.subheader("Active Members Today")
st.markdown("Enter members in the format `Name` or `Name:Limit`, separated by commas.")
working_input = st.text_input("Who is working today?", placeholder="Barry, Annie, Tony:50, Tomas")

if not working_input:
    st.warning("Please enter at least one team member.")
    st.stop()

# --- Process member input ---
active_members = []
custom_limits = {}
for part in working_input.split(","):
    part = part.strip()
    if not part:
        continue
    if ":" in part:
        name, limit = part.split(":")
        name = name.strip().title()
        try:
            limit = int(limit.strip())
            active_members.append(name)
            custom_limits[name] = limit
        except ValueError:
            st.warning(f"Invalid limit for {name}, ignoring. Using default split.")
            active_members.append(name.strip().title())
    else:
        active_members.append(part.strip().title())

if not active_members:
    st.error("❌ No team members entered!")
    st.stop()

st.success(f"Active members: {', '.join(active_members)}")
if custom_limits:
    for name, limit in custom_limits.items():
        st.info(f"Custom limit: {name} = {limit} products")

# --- Load workbook ---
wb = load_workbook(uploaded_file)
qa_ws = wb["QA"]
assignments_ws = wb["Assignments"]

# --- Read brand-to-member assignments ---
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand, member = row[0], row[1]
    if brand and member:
        brand_to_member[brand.strip().title()] = member.strip().title()

# --- Build QA data grouped by brand ---
brand_blocks = defaultdict(list)
qa_rows = []
for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    m_value = row[12]
    brand = row[14]
    workflow = str(row[8]).strip() if row[8] else ""
    col_aq = row[42]
    if m_value is not None and str(m_value).strip() != "":
        qa_rows.append((i, brand, workflow, col_aq))
        if brand:
            brand_blocks[brand.strip().title()].append(i)

# --- Apply backlog sorting if needed ---
if backlog_mode:
    st.info("🕐 Backlog mode ON — sorting all rows by earliest AQ date.")
    def sort_key(row_idx):
        date_val = qa_ws[f"AQ{row_idx}"].value
        return date_val if isinstance(date_val, datetime) else datetime.max
    for brand in brand_blocks:
        brand_blocks[brand].sort(key=sort_key)
else:
    st.info("🚀 Backlog mode OFF — assigning in normal order.")

# --- Calculate per-member limits ---
total_products = len(qa_rows)
total_custom = sum(custom_limits.values())
num_remaining_members = len(active_members) - len(custom_limits)
remaining_products = max(0, total_products - total_custom)
default_limit = remaining_products // num_remaining_members if num_remaining_members > 0 else 0
remainder = remaining_products % num_remaining_members if num_remaining_members > 0 else 0

member_limits = {}
idx = 0
for member in active_members:
    if member in custom_limits:
        member_limits[member] = custom_limits[member]
    else:
        member_limits[member] = default_limit + (1 if idx < remainder else 0)
        idx += 1

assignments = {member: [] for member in active_members}
counts = {member: 0 for member in active_members}

st.subheader("Per-Member Targets")
for member, limit in member_limits.items():
    st.write(f"- {member}: {limit} products")

# --- Step 1: Pre-assign brands ---
for brand, rows in brand_blocks.items():
    if brand in brand_to_member:
        member = brand_to_member[brand]
        if member not in active_members:
            eligible = [m for m in active_members if counts[m] < member_limits[m]]
            if not eligible:
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                continue
            member = max(eligible, key=lambda x: member_limits[x] - counts[x])
        for r in rows:
            qa_ws[f"A{r}"].value = member
            assignments[member].append(r)
            counts[member] += 1

# --- Step 2: Assign remaining ---
for brand, rows in brand_blocks.items():
    unassigned = [r for r in rows if qa_ws[f"A{r}"].value in [None, ""]]
    if not unassigned:
        continue
    block_size = len(unassigned)
    eligible = [m for m in active_members if counts[m] + block_size <= member_limits[m]]
    if eligible:
        member = min(eligible, key=lambda x: counts[x])
        for r in unassigned:
            qa_ws[f"A{r}"].value = member
            assignments[member].append(r)
            counts[member] += 1
    else:
        for r in unassigned:
            eligible_split = [m for m in active_members if counts[m] < member_limits[m]]
            if not eligible_split:
                qa_ws[f"A{r}"].value = "Backlog"
            else:
                member = min(eligible_split, key=lambda x: counts[x])
                qa_ws[f"A{r}"].value = member
                assignments[member].append(r)
                counts[member] += 1

# --- Convert formulas to values ---
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            cell.value = cell.value

# --- Save to BytesIO for download ---
output = BytesIO()
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
wb.save(output)
output.seek(0)

st.success("✅ Assignment complete!")

st.subheader("Assignment Summary")
for member, rows in assignments.items():
    limit = member_limits[member]
    st.write(f"- {member}: {len(rows)} products (Target: {limit})")
backlog_count = sum(1 for r in qa_rows if qa_ws[f"A{r[0]}"].value == "Backlog")
st.write(f"- Backlog: {backlog_count} products")

st.download_button(
    label="📥 Download Assigned Excel",
    data=output,
    file_name=f"QA_Assigned_{timestamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
