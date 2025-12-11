import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict, deque
import math

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="QA Assignment Tool", layout="wide")
st.title("QA Assignment Tool 📊")
st.write("Assigns products to QA team. (Rewritten with split-if-unbalanced logic — interleaved split)")

# ---------------------------
# Helpers
# ---------------------------

def title_or_none(val):
    return val.strip().title() if isinstance(val, str) and val.strip() else None


def remaining_capacity(member, limits, counts):
    return max(0, limits.get(member, 0) - counts.get(member, 0))


def compute_loads_after_assignment(counts, member_to_add, add):
    # returns a sorted list of loads descending
    new_counts = counts.copy()
    new_counts[member_to_add] = new_counts.get(member_to_add, 0) + add
    loads = sorted(new_counts.values(), reverse=True)
    return loads


def top_and_second(loads):
    if not loads:
        return 0, 0
    if len(loads) == 1:
        return loads[0], 0
    return loads[0], loads[1]


# ---------------------------
# File upload
# ---------------------------
uploaded_file = st.file_uploader("Upload QA Template", type=["xlsx"]) 
if not uploaded_file:
    st.info("Please upload an Excel (.xlsx) file containing 'QA' and 'Assignments' sheets.")
    st.stop()

# Save to temp file
temp_file_path = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
with open(temp_file_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

# Load workbook
wb = load_workbook(temp_file_path)
if "QA" not in wb.sheetnames or "Assignments" not in wb.sheetnames:
    st.error("Excel file must contain 'QA' and 'Assignments' sheets.")
    st.stop()

qa_ws = wb["QA"]
assignments_ws = wb["Assignments"]

# Options
backlog_mode = st.checkbox("Backlog mode (sort by earliest AQ date)", value=False)
st.write("Enter active members today (e.g: Ross:100, Phoebe:80, Monica)")
working_input = st.text_input("Active members")

if not working_input:
    st.error("Please enter at least one active member.")
    st.stop()

# Parse active members and custom limits
active_members = []
member_limits = {}
for part in working_input.split(","):
    part = part.strip()
    if not part:
        continue
    if ":" in part:
        name, limit = part.split(":", 1)
        name = name.strip().title()
        try:
            lim = int(limit.strip())
        except Exception:
            lim = 100
        active_members.append(name)
        member_limits[name] = lim
    else:
        active_members.append(part.strip().title())

if not active_members:
    st.error("No active members parsed from input!")
    st.stop()

# default limits for members not specified
for m in active_members:
    if m not in member_limits:
        member_limits[m] = 100

# Read brand->member preassignments from Assignments sheet
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand, member = row[0], row[1]
    if brand and member:
        brand_to_member[title_or_none(brand)] = title_or_none(member)

# Build brand blocks (preserving file order)
brand_blocks = defaultdict(list)
row_brand_order = []
qa_rows = []

# Based on your original script's indexing: row[12] -> column M-value, row[14] -> brand, row[8] -> workflow, row[42] -> AQ
for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    m_value = row[12]  # M column condition used before to include row
    brand = row[14]
    col_aq = row[42]

    if m_value is not None and str(m_value).strip() != "":
        qa_rows.append((i, brand, col_aq))
        if brand:
            btitle = title_or_none(brand)
            if btitle not in brand_blocks:
                row_brand_order.append(btitle)
            brand_blocks[btitle].append(i)

# Backlog mode: sort each brand block by AQ date (earliest first)
if backlog_mode:
    def row_aq_date(row_idx):
        try:
            val = qa_ws[f"AQ{row_idx}"].value
            return val if isinstance(val, datetime) else datetime.max
        except Exception:
            return datetime.max
    for b in brand_blocks:
        brand_blocks[b].sort(key=row_aq_date)

# Prepare blocks preserving first seen order
blocks = []
for b in row_brand_order:
    rows = brand_blocks[b]
    blocks.append((b, rows.copy()))

# Sort blocks largest -> smallest for smart reservation
blocks.sort(key=lambda x: len(x[1]), reverse=True)

# State
counts = {m: 0 for m in active_members}  # current assigned counts
assignments = {m: [] for m in active_members}
assigned_blocks = []  # records of assigned block fragments: {'brand', 'rows', 'member'}
backlog_rows = []

# Utility functions for assignment

def assign_rows_to_member(member, rows):
    for r in rows:
        qa_ws[f"A{r}"].value = member
    assignments[member].extend(rows)
    counts[member] += len(rows)
    assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': member})


def assign_fragment_to_member(member, rows):
    # This function records fragments individually (used by splits)
    for r in rows:
        qa_ws[f"A{r}"].value = member
    assignments[member].extend(rows)
    counts[member] += len(rows)
    assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': member})

# The imbalance threshold is 30% (user choice). We compare top vs second-top after assignment.
IMBALANCE_RATIO = 1.30

# We'll iterate through blocks and try to assign them using the new rule
blocks_queue = deque(blocks)
iteration = 0
max_iterations = 20000

while blocks_queue and iteration < max_iterations:
    iteration += 1
    current_brand, rows = blocks_queue.popleft()
    block_size = len(rows)

    # Skip empty brand
    if block_size == 0:
        continue

    # If preassigned to a specific member and that member is active
    pre_member = brand_to_member.get(current_brand)
    if pre_member and pre_member in active_members:
        cap = remaining_capacity(pre_member, member_limits, counts)
        if cap >= block_size:
            # assign whole
            current_brand = current_brand  # keep for assign_rows_to_member
            assign_rows_to_member(pre_member, rows)
            continue
        else:
            # If preassigned member can't take whole block, we'll try to give as much as possible then process remainder
            take = min(cap, block_size)
            if take > 0:
                assign_rows_to_member(pre_member, rows[:take])
            remaining = rows[take:]
            if remaining:
                blocks_queue.appendleft((current_brand, remaining))
            continue

    # Choose best single candidate (one who would accept the entire block and has max remaining capacity)
    candidates_can_take = [m for m in active_members if remaining_capacity(m, member_limits, counts) >= block_size]
    if candidates_can_take:
        # choose the one with the most remaining capacity, tie-break on lowest current count
        candidates_can_take.sort(key=lambda m: (-remaining_capacity(m, member_limits, counts), counts[m]))
        best_candidate = candidates_can_take[0]

        # Simulate assigning whole block to best_candidate and compute top and second loads
        loads_after = compute_loads_after_assignment(counts, best_candidate, block_size)
        top, second = top_and_second(loads_after)

        # If assignment would make top > 1.30 * second -> we will split
        if second == 0:
            would_imbalance = (top > 0 and second == 0)
        else:
            would_imbalance = (top > IMBALANCE_RATIO * second)

        if not would_imbalance:
            # Assign whole block to the best candidate normally
            current_brand = current_brand
            assign_rows_to_member(best_candidate, rows)
            continue
        # else we fall through to splitting logic

    # If no single candidate can take block, or splitting is required by imbalance, we will attempt to split

    # Decide members eligible for split (members with remaining capacity)
    eligible_members = [m for m in active_members if remaining_capacity(m, member_limits, counts) > 0]
    if not eligible_members:
        # No capacity anywhere -> backlog the whole block
        for r in rows:
            qa_ws[f"A{r}"].value = "Backlog"
            backlog_rows.append(r)
        continue

    # Compute remaining capacities
    rem_caps = {m: remaining_capacity(m, member_limits, counts) for m in eligible_members}
    total_rem_cap = sum(rem_caps.values())
    if total_rem_cap == 0:
        for r in rows:
            qa_ws[f"A{r}"].value = "Backlog"
            backlog_rows.append(r)
        continue

    # Determine tentative split counts proportional to remaining capacity
    # Use floor of proportional allocation, then distribute leftover one by one to highest remaining capacity
    tentative = {}
    for m in eligible_members:
        proportion = rem_caps[m] / total_rem_cap
        tentative[m] = math.floor(proportion * block_size)

    assigned_sum = sum(tentative.values())
    remaining_to_assign = block_size - assigned_sum

    # Distribute leftovers to members sorted by remaining capacity (desc)
    members_by_cap = sorted(eligible_members, key=lambda m: -rem_caps[m])
    idx = 0
    while remaining_to_assign > 0 and members_by_cap:
        m = members_by_cap[idx % len(members_by_cap)]
        # only give if we haven't exceeded that member's remaining cap
        if tentative[m] < rem_caps[m]:
            tentative[m] += 1
            remaining_to_assign -= 1
        idx += 1
        # break guard
        if idx > block_size * 5:
            break

    # Final guard: ensure no one gets more than their remaining capacity
    for m in tentative:
        if tentative[m] > rem_caps[m]:
            tentative[m] = rem_caps[m]

    # If after all allocation we still haven't assigned all rows (rare), put leftovers to backlog
    total_assigned = sum(tentative.values())
    leftover_after_caps = block_size - total_assigned

    # Interleaved assignment: cycle through members and give one row at a time until their quota is met
    quotas = {m: tentative[m] for m in eligible_members}
    assigned_order = []  # list of (member, row)
    row_iter = iter(rows)

    # Build a queue of members who still need rows
    from collections import deque as _deque
    member_queue = _deque([m for m in eligible_members if quotas[m] > 0])

    # Interleaved distribute
    while member_queue and rows:
        m = member_queue.popleft()
        if quotas[m] <= 0:
            continue
        # assign one row
        r = rows.pop(0)
        assign_fragment_to_member(m, [r])
        quotas[m] -= 1
        if quotas[m] > 0:
            member_queue.append(m)

    # Any leftover rows that couldn't be assigned (due to capacity changes) -> backlog
    if rows:
        for r in rows:
            qa_ws[f"A{r}"].value = "Backlog"
            backlog_rows.append(r)

# After loop

# Final pass: convert formulas to values (attempt to keep values as-is)
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            # Evaluate formula if possible, otherwise keep current value
            try:
                cell.value = cell.value
            except Exception:
                pass

# Save output
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")
st.write(f"📄 Saved as: {output_path}")

# Summary
st.write("📊 Summary of assignments:")
for member in active_members:
    limit = member_limits[member]
    st.write(f"- {member}: {len(assignments.get(member, []))} products (Target: {limit})")
st.write(f"- Backlog (explicitly set): {len(backlog_rows)} products")

# Download button
with open(output_path, "rb") as f:
    st.download_button(
        label="📥 Download Assigned Excel",
        data=f,
        file_name=output_path,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
