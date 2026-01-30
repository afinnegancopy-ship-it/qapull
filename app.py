import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="QA Assignment Tool", layout="wide")
st.title("QA Assignment Tool 📊")
st.write("Even split with brand integrity. Priorities: 1) Perfect balance, 2) Pre-assignments, 3) Keep brands together")

# ---------------------------
# Helpers
# ---------------------------

def get_header_map(worksheet):
    """Build a dictionary mapping header names to column indices (1-based)."""
    header_map = {}
    for col_idx, cell in enumerate(worksheet[1], start=1):
        if cell.value:
            header_map[cell.value.strip()] = col_idx
            header_map[cell.value.strip().lower()] = col_idx
    return header_map


def get_col_index(header_map, *possible_names):
    """Get column index from header map, trying multiple possible header names."""
    for name in possible_names:
        if name in header_map:
            return header_map[name]
        if name.lower() in header_map:
            return header_map[name.lower()]
    raise KeyError(f"Could not find column with any of these headers: {possible_names}")


def title_or_none(val):
    return val.strip().title() if isinstance(val, str) and val.strip() else None


def calculate_exact_targets(active_members, total_products, member_limits):
    """
    Calculate EXACT targets for perfect distribution, respecting member limits.
    Redistributes overflow from limited members to unlimited ones.
    """
    remaining = total_products
    targets = {m: 0 for m in active_members}
    locked = set()
    
    for _ in range(len(active_members) + 1):  # Max iterations = number of members
        unlocked = [m for m in active_members if m not in locked]
        if not unlocked:
            break
        
        per_person, remainder = divmod(remaining, len(unlocked))
        changed = False
        
        for i, m in enumerate(unlocked):
            fair_share = per_person + (1 if i < remainder else 0)
            limit = member_limits.get(m, 999)
            
            if limit < fair_share:
                targets[m] = limit
                locked.add(m)
                changed = True
            else:
                targets[m] = fair_share
        
        if not changed:
            break
        remaining = total_products - sum(targets[m] for m in locked)
    
    return targets


# ---------------------------
# File upload
# ---------------------------
uploaded_file = st.file_uploader("Upload QA Template", type=["xlsx"])
if not uploaded_file:
    st.info("Please upload an Excel (.xlsx) file containing 'QA' and 'Assignments' sheets.")
    st.stop()

temp_file_path = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
with open(temp_file_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

wb = load_workbook(temp_file_path)
if "QA" not in wb.sheetnames or "Assignments" not in wb.sheetnames:
    st.error("Excel file must contain 'QA' and 'Assignments' sheets.")
    st.stop()

qa_ws = wb["QA"]
assignments_ws = wb["Assignments"]

# ---------------------------
# Build header maps
# ---------------------------
qa_headers = get_header_map(qa_ws)
assignments_headers = get_header_map(assignments_ws)

try:
    COL_ASSIGNED = get_col_index(qa_headers, "Assigned", "assigned", "ASSIGNED")
    COL_PIM_PARENT_ID = get_col_index(qa_headers, "Pim Parent ID", "pim parent id", "PIM Parent ID")
    COL_BRAND = get_col_index(qa_headers, "Brand", "brand", "BRAND")
    COL_BT_IMAGE_DATE = get_col_index(qa_headers, "Bt Image Date", "bt image date", "BT Image Date",
                                       "Enrichment QA Date", "enrichment qa date")
except KeyError as e:
    st.error(f"Missing required column in QA sheet: {e}")
    st.stop()

try:
    COL_ASSIGN_BRAND = get_col_index(assignments_headers, "BRAND", "Brand", "brand")
    COL_ASSIGN_QAER = get_col_index(assignments_headers, "Qaer", "qaer", "QAER", "QA", "Member", "member")
except KeyError as e:
    st.error(f"Missing required column in Assignments sheet: {e}")
    st.stop()

with st.expander("📋 Detected Column Mappings"):
    st.write("**QA Sheet:**")
    st.write(f"- Assigned: Column {get_column_letter(COL_ASSIGNED)}")
    st.write(f"- Pim Parent ID: Column {get_column_letter(COL_PIM_PARENT_ID)}")
    st.write(f"- Brand: Column {get_column_letter(COL_BRAND)}")
    st.write(f"- BT Image Date: Column {get_column_letter(COL_BT_IMAGE_DATE)}")
    st.write("**Assignments Sheet:**")
    st.write(f"- Brand: Column {get_column_letter(COL_ASSIGN_BRAND)}")
    st.write(f"- Qaer: Column {get_column_letter(COL_ASSIGN_QAER)}")

# ---------------------------
# Options
# ---------------------------
st.subheader("⚙️ Configuration")

backlog_mode = st.checkbox("Backlog mode (sort by earliest BT Image date)", value=False)

st.write("Enter active members today (e.g: Ross:100, Phoebe:80, Monica)")
working_input = st.text_input("Active members")

if not working_input:
    st.error("Please enter at least one active member.")
    st.stop()

# Parse active members
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
        except:
            lim = 999
        active_members.append(name)
        member_limits[name] = lim
    else:
        active_members.append(part.strip().title())

for m in active_members:
    if m not in member_limits:
        member_limits[m] = 999

# Preassignments from Assignments sheet
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand = row[COL_ASSIGN_BRAND - 1] if len(row) >= COL_ASSIGN_BRAND else None
    member = row[COL_ASSIGN_QAER - 1] if len(row) >= COL_ASSIGN_QAER else None
    if brand and member:
        brand_to_member[title_or_none(brand)] = title_or_none(member)

if brand_to_member:
    with st.expander(f"📌 Pre-assigned Brands ({len(brand_to_member)})"):
        for brand, member in sorted(brand_to_member.items()):
            status = "✅" if member in active_members else "⚠️ (not active today)"
            st.write(f"- {brand} → {member} {status}")

# ---------------------------
# Build brand blocks (read all data in one pass)
# ---------------------------
brand_blocks = defaultdict(list)
row_brand_order = []
row_dates = {}  # Only populate if backlog_mode

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    pim_parent_id = row[COL_PIM_PARENT_ID - 1] if len(row) >= COL_PIM_PARENT_ID else None
    brand = row[COL_BRAND - 1] if len(row) >= COL_BRAND else None

    if pim_parent_id is not None and str(pim_parent_id).strip():
        btitle = title_or_none(brand) if brand else "No Brand"
        if btitle not in brand_blocks:
            row_brand_order.append(btitle)
        brand_blocks[btitle].append(i)
        
        if backlog_mode:
            date_val = row[COL_BT_IMAGE_DATE - 1] if len(row) >= COL_BT_IMAGE_DATE else None
            row_dates[i] = date_val if isinstance(date_val, datetime) else datetime.max

if backlog_mode:
    for b in brand_blocks:
        brand_blocks[b].sort(key=lambda r: row_dates.get(r, datetime.max))

# Build blocks list
blocks = []
for b in row_brand_order:
    pre_member = brand_to_member.get(b)
    is_preassigned = pre_member is not None and pre_member in active_members
    blocks.append({
        'brand': b,
        'rows': brand_blocks[b],
        'size': len(brand_blocks[b]),
        'preassigned_to': pre_member if is_preassigned else None
    })

blocks.sort(key=lambda x: (0 if x['preassigned_to'] else 1, x['size']))

# Calculate targets
total_products = sum(b['size'] for b in blocks)
targets = calculate_exact_targets(active_members, total_products, member_limits)

total_capacity = sum(targets.values())
if total_capacity < total_products:
    shortfall = total_products - total_capacity
    st.warning(f"⚠️ Total capacity ({total_capacity}) is less than products ({total_products}). {shortfall} will go to backlog.")

st.write("📊 **Exact Targets for Perfect Split:**")
target_cols = st.columns(len(active_members))
for i, m in enumerate(active_members):
    with target_cols[i]:
        limit_info = f" (limit: {member_limits[m]})" if member_limits[m] < 999 else ""
        st.metric(m, f"{targets[m]} products", delta=limit_info if limit_info else None)

st.divider()

# ---------------------------
# FAST ASSIGNMENT ALGORITHM
# ---------------------------
# Instead of writing to Excel cell-by-cell, we build a dict of row -> assignee
# then write everything in one batch at the end

row_assignments = {}  # row_number -> member_name
counts = {m: 0 for m in active_members}
member_rows = {m: [] for m in active_members}  # For final balancing

for block in blocks:
    brand = block['brand']
    rows = list(block['rows'])  # Copy
    preassigned_to = block['preassigned_to']
    
    if not rows:
        continue
    
    # Pre-assigned member gets first dibs
    if preassigned_to:
        room = targets[preassigned_to] - counts[preassigned_to]
        if room > 0:
            take = min(room, len(rows))
            for r in rows[:take]:
                row_assignments[r] = preassigned_to
                member_rows[preassigned_to].append(r)
            counts[preassigned_to] += take
            rows = rows[take:]
    
    # Distribute remaining
    while rows:
        # Find members with room, sorted by most room first
        with_room = [(m, targets[m] - counts[m]) for m in active_members if counts[m] < targets[m]]
        if not with_room:
            # Backlog
            for r in rows:
                row_assignments[r] = "Backlog"
            break
        
        with_room.sort(key=lambda x: -x[1])
        
        # Try to keep brand together
        brand_size = len(rows)
        assigned = False
        
        for m, room in with_room:
            if room >= brand_size:
                # This member can take the whole remaining brand
                for r in rows:
                    row_assignments[r] = m
                    member_rows[m].append(r)
                counts[m] += brand_size
                rows = []
                assigned = True
                break
        
        if not assigned:
            # Must split - distribute to members with room
            for m, room in with_room:
                if not rows:
                    break
                take = min(room, len(rows))
                for r in rows[:take]:
                    row_assignments[r] = m
                    member_rows[m].append(r)
                counts[m] += take
                rows = rows[take:]

# ---------------------------
# FAST FINAL BALANCE
# ---------------------------
# Calculate all moves needed upfront, then apply

over = {m: counts[m] - targets[m] for m in active_members if counts[m] > targets[m]}
under = {m: targets[m] - counts[m] for m in active_members if counts[m] < targets[m]}

final_adjustments = 0
while over and under:
    # Get most over and most under
    from_member = max(over, key=over.get)
    to_member = max(under, key=under.get)
    
    if not member_rows[from_member]:
        break
    
    # Move one product
    row_to_move = member_rows[from_member].pop()
    member_rows[to_member].append(row_to_move)
    row_assignments[row_to_move] = to_member
    
    counts[from_member] -= 1
    counts[to_member] += 1
    final_adjustments += 1
    
    # Update over/under
    if counts[from_member] <= targets[from_member]:
        del over[from_member]
    else:
        over[from_member] = counts[from_member] - targets[from_member]
    
    if counts[to_member] >= targets[to_member]:
        del under[to_member]
    else:
        under[to_member] = targets[to_member] - counts[to_member]

# ---------------------------
# BATCH WRITE TO EXCEL (single pass)
# ---------------------------
assigned_col_letter = get_column_letter(COL_ASSIGNED)
for row_num, assignee in row_assignments.items():
    qa_ws[f"{assigned_col_letter}{row_num}"].value = assignee

# ---------------------------
# Results Summary
# ---------------------------
st.subheader("📈 Final Distribution")

result_cols = st.columns(len(active_members))
for i, m in enumerate(active_members):
    with result_cols[i]:
        diff = counts[m] - targets[m]
        if diff == 0:
            st.success(f"**{m}**: {counts[m]} ✓")
        elif diff > 0:
            st.warning(f"**{m}**: {counts[m]} (+{diff})")
        else:
            st.error(f"**{m}**: {counts[m]} ({diff})")

backlog_count = sum(1 for v in row_assignments.values() if v == "Backlog")
if backlog_count:
    st.warning(f"📦 **Backlog**: {backlog_count} products")

if final_adjustments > 0:
    st.info(f"🔄 Made {final_adjustments} final adjustments for perfect balance")

# ---------------------------
# Save (skip formula conversion - only touch Assigned column)
# ---------------------------
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")

with open(output_path, "rb") as f:
    st.download_button(
        label="📥 Download Assigned Excel",
        data=f,
        file_name=output_path,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
