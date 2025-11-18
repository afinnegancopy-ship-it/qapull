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

# --- Backlog mode checkbox ---
backlog_mode = st.checkbox("Backlog mode?", value=False, help="Sort by earliest AQ date if enabled.")

# --- Input active members ---
st.subheader("Active Members Today")
st.markdown("Enter members in the format `Name` or `Name:Limit`, separated by commas.")
working_input = st.text_input("Who is working today?", placeholder="Barry, Annie, Tony:50, Tomas")

if not working_input:
    st.warning("Please enter at least one team member.")
    st.stop()

# --- Process members ---
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
            custom_limits[name] = limit
            active_members.append(name)
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
    col_aq = row[42]
    if m_value is not None and str(m_value).strip() != "":
        qa_rows.append((i, brand, col_aq))
        if brand:
            brand_blocks[brand.strip().title()].append(i)

# --- Backlog sorting ---
if backlog_mode:
    st.info("🕐 Backlog mode ON — sorting all rows by earliest AQ date.")
    def sort_key(row_idx):
        date_val = qa_ws[f"AQ{row_idx}"].value
        return date_val if isinstance(date_val, datetime) else datetime.max
    for brand in brand_blocks:
        brand_blocks[brand].sort(key=sort_key)
else:
    st.info("🚀 Backlog mode OFF — assigning in normal order.")

# --- Compute global even target (Option B priority) ---
# First: enforce no one can exceed 100
MAX_PER_MEMBER = 100
num_members = len(active_members)
max_possible = num_members * MAX_PER_MEMBER

if len(qa_rows) > max_possible:
    st.error("❌ Not enough capacity. Total products exceed 100 per member limit.")
    st.stop()

# Even split respecting MAX 100 & custom limits
member_limits = {}
remaining_products = len(qa_rows)

# Apply custom limits first
total_custom = sum(custom_limits.values())
if total_custom > remaining_products:
    total_custom = remaining_products

# Assign custom limits
total_custom_assigned = 0
for m in active_members:
    if m in custom_limits:
        member_limits[m] = min(custom_limits[m], MAX_PER_MEMBER)
        total_custom_assigned += member_limits[m]
    else:
        member_limits[m] = 0

remaining_products -= total_custom_assigned

# Distribute remaining evenly to non-custom
non_custom = [m for m in active_members if m not in custom_limits]
if non_custom:
    per = min(MAX_PER_MEMBER, remaining_products // len(non_custom))
    rem = remaining_products % len(non_custom)
    for idx, m in enumerate(non_custom):
        member_limits[m] = min(MAX_PER_MEMBER, per + (1 if idx < rem else 0))

# --- Assignment structures ---
assignments = {m: [] for m in active_members}
counts = {m: 0 for m in active_members}

st.subheader("Per-Member Targets (Even Split, Max 100)")
for member, limit in member_limits.items():
    st.write(f"- {member}: {limit} products")

# --- Step 1: Brand Cohesion (Option B: keep together only if it doesn't break even split) ---
for brand, rows in brand_blocks.items():
    block_size = len(rows)

    # Try preferred member if exists
    preferred = brand_to_member.get(brand, None)
    candidate_members = active_members.copy()

    if preferred in candidate_members:
        candidate_members = [preferred] + [m for m in candidate_members if m != preferred]

    assigned = False
    for m in candidate_members:
        if counts[m] + block_size <= member_limits[m]:
            for r in rows:
                qa_ws[f"A{r}"].value = m
                assignments[m].append(r)
            counts[m] += block_size
            assigned = True
            break

    # If brand block doesn't fit anywhere, assign row-by-row later
    if not assigned:
        for r in rows:
            qa_ws[f"A{r}"].value = None

# --- Step 2: Row-by-row fill for unassigned rows ---
for brand, rows in brand_blocks.items():
    for r in rows:
        if qa_ws[f"A{r}"].value:
            continue
        eligible = [m for m in active_members if counts[m] < member_limits[m]]
        if not eligible:
            qa_ws[f"A{r}"].value = "Backlog"
        else:
            m = min(eligible, key=lambda x: counts[x])
            qa_ws[f"A{r}"].value = m
            assignments[m].append(r)
            counts[m] += 1

# --- Convert formulas to values ---
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            cell.value = cell.value

# --- Save to BytesIO ---
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
