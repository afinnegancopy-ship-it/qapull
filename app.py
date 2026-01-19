import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict, deque
import math

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="QA Assignment Tool", layout="wide")
st.title("QA Assignment Tool 📊")
st.write("Assigns products to QA team with configurable distribution modes.")

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


def remaining_capacity(member, limits, counts):
    return max(0, limits.get(member, 0) - counts.get(member, 0))


def compute_loads_after_assignment(counts, member_to_add, add):
    new_counts = counts.copy()
    new_counts[member_to_add] = new_counts.get(member_to_add, 0) + add
    return sorted(new_counts.values(), reverse=True)


def top_and_second(loads):
    if not loads:
        return 0, 0
    if len(loads) == 1:
        return loads[0], 0
    return loads[0], loads[1]


def calculate_targets(active_members, member_limits, total_products):
    """Calculate ideal target for each member based on their limits and total products."""
    # First, calculate proportional targets based on limits
    total_capacity = sum(member_limits[m] for m in active_members)
    
    if total_capacity == 0:
        return {m: 0 for m in active_members}
    
    # Proportional targets, but capped at each member's limit
    targets = {}
    for m in active_members:
        proportion = member_limits[m] / total_capacity
        ideal = round(total_products * proportion)
        targets[m] = min(ideal, member_limits[m])
    
    # Adjust to ensure we hit total exactly
    assigned = sum(targets.values())
    diff = total_products - assigned
    
    # Distribute remainder to members with most remaining capacity
    members_by_slack = sorted(active_members, 
                               key=lambda m: member_limits[m] - targets[m], 
                               reverse=True)
    i = 0
    while diff > 0 and i < len(members_by_slack) * 2:
        m = members_by_slack[i % len(members_by_slack)]
        if targets[m] < member_limits[m]:
            targets[m] += 1
            diff -= 1
        i += 1
    
    return targets


def distance_from_target(member, counts, targets):
    """How far is this member from their target? Positive = room to add."""
    return targets.get(member, 0) - counts.get(member, 0)


def rebalance_assignments(qa_ws, assignments, counts, targets, active_members, assigned_col_letter, max_iterations=1000):
    """Post-process to move products from over-target members to under-target members."""
    iterations = 0
    moves_made = 0
    
    while iterations < max_iterations:
        iterations += 1
        
        # Find most over-target and most under-target members
        over_target = [(m, counts[m] - targets[m]) for m in active_members if counts[m] > targets[m]]
        under_target = [(m, targets[m] - counts[m]) for m in active_members if counts[m] < targets[m]]
        
        if not over_target or not under_target:
            break
        
        # Sort to get most imbalanced
        over_target.sort(key=lambda x: -x[1])
        under_target.sort(key=lambda x: -x[1])
        
        from_member = over_target[0][0]
        to_member = under_target[0][0]
        
        # Check if rebalancing would improve things
        current_diff = counts[from_member] - counts[to_member]
        if current_diff <= 1:
            break
        
        # Move one product
        if assignments[from_member]:
            row_to_move = assignments[from_member].pop()
            assignments[to_member].append(row_to_move)
            counts[from_member] -= 1
            counts[to_member] += 1
            qa_ws[f"{assigned_col_letter}{row_to_move}"].value = to_member
            moves_made += 1
        else:
            break
    
    return moves_made


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

# Display detected columns
with st.expander("📋 Detected Column Mappings"):
    st.write("**QA Sheet:**")
    st.write(f"- Assigned: Column {get_column_letter(COL_ASSIGNED)} (index {COL_ASSIGNED})")
    st.write(f"- Pim Parent ID: Column {get_column_letter(COL_PIM_PARENT_ID)} (index {COL_PIM_PARENT_ID})")
    st.write(f"- Brand: Column {get_column_letter(COL_BRAND)} (index {COL_BRAND})")
    st.write(f"- BT Image Date: Column {get_column_letter(COL_BT_IMAGE_DATE)} (index {COL_BT_IMAGE_DATE})")
    st.write("**Assignments Sheet:**")
    st.write(f"- Brand: Column {get_column_letter(COL_ASSIGN_BRAND)} (index {COL_ASSIGN_BRAND})")
    st.write(f"- Qaer: Column {get_column_letter(COL_ASSIGN_QAER)} (index {COL_ASSIGN_QAER})")

# ---------------------------
# Options
# ---------------------------
st.subheader("⚙️ Configuration")

col1, col2 = st.columns(2)

with col1:
    distribution_mode = st.radio(
        "Distribution Mode",
        options=["Balanced (Brand Priority)", "Balanced (Even Priority)", "Strict Even Split"],
        index=1,
        help="""
        • **Brand Priority**: Keeps brands together when possible, may have some imbalance
        • **Even Priority**: Aims for even distribution with post-rebalancing, still tries to group brands
        • **Strict Even Split**: Round-robin distribution, ignores brand grouping entirely
        """
    )

with col2:
    backlog_mode = st.checkbox("Backlog mode (sort by earliest BT Image date)", value=False)

# Advanced settings
with st.expander("🔧 Advanced Settings"):
    adv_col1, adv_col2 = st.columns(2)
    with adv_col1:
        IMBALANCE_RATIO = st.slider("Imbalance Ratio (Brand Priority mode)", 1.05, 2.0, 1.30, 0.05,
                                     help="Lower = more aggressive splitting for balance")
        SPLIT_SIZE_THRESHOLD = st.slider("Min Brand Size to Split", 5, 100, 50, 5,
                                          help="Brands smaller than this won't be split in Brand Priority mode")
    with adv_col2:
        TARGET_TOLERANCE = st.slider("Target Tolerance % (Even Priority)", 0.0, 0.3, 0.1, 0.05,
                                      help="How much over-target before forcing a split")
        REBALANCE_ENABLED = st.checkbox("Enable Post-Rebalancing", value=True,
                                         help="Move products after initial assignment to even out distribution")

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
            lim = 100
        active_members.append(name)
        member_limits[name] = lim
    else:
        active_members.append(part.strip().title())

for m in active_members:
    if m not in member_limits:
        member_limits[m] = 100

# Preassignments
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand = row[COL_ASSIGN_BRAND - 1] if len(row) >= COL_ASSIGN_BRAND else None
    member = row[COL_ASSIGN_QAER - 1] if len(row) >= COL_ASSIGN_QAER else None
    if brand and member:
        brand_to_member[title_or_none(brand)] = title_or_none(member)

# Build brand blocks
brand_blocks = defaultdict(list)
row_brand_order = []
qa_rows = []

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    pim_parent_id = row[COL_PIM_PARENT_ID - 1] if len(row) >= COL_PIM_PARENT_ID else None
    brand = row[COL_BRAND - 1] if len(row) >= COL_BRAND else None
    bt_image_date = row[COL_BT_IMAGE_DATE - 1] if len(row) >= COL_BT_IMAGE_DATE else None

    if pim_parent_id is not None and str(pim_parent_id).strip():
        qa_rows.append((i, brand, bt_image_date))
        if brand:
            btitle = title_or_none(brand)
            if btitle not in brand_blocks:
                row_brand_order.append(btitle)
            brand_blocks[btitle].append(i)

if backlog_mode:
    date_col_letter = get_column_letter(COL_BT_IMAGE_DATE)

    def row_aq_date(row_idx):
        try:
            val = qa_ws[f"{date_col_letter}{row_idx}"].value
        except:
            val = None
        return val if isinstance(val, datetime) else datetime.max

    for b in brand_blocks:
        brand_blocks[b].sort(key=row_aq_date)

blocks = []
for b in row_brand_order:
    blocks.append((b, brand_blocks[b].copy()))
blocks.sort(key=lambda x: len(x[1]), reverse=True)

# Calculate totals and targets
total_products = sum(len(rows) for _, rows in blocks)
targets = calculate_targets(active_members, member_limits, total_products)

# Display targets
st.write("📊 **Calculated Targets:**")
target_cols = st.columns(len(active_members))
for i, m in enumerate(active_members):
    with target_cols[i]:
        st.metric(m, f"Target: {targets[m]}", f"Max: {member_limits[m]}")

st.divider()

# Initialize
counts = {m: 0 for m in active_members}
assignments = {m: [] for m in active_members}
assigned_blocks = []
backlog_rows = []
assigned_col_letter = get_column_letter(COL_ASSIGNED)

# ---------------------------
# STRICT EVEN SPLIT MODE
# ---------------------------
if distribution_mode == "Strict Even Split":
    from itertools import cycle
    
    # Flatten all products with their brand info
    all_rows = []
    for brand, rows in blocks:
        for row in rows:
            all_rows.append((row, brand))
    
    # Create member cycle
    member_cycle = cycle(active_members)
    current_member_idx = 0
    
    for row, brand in all_rows:
        # Find next member with capacity
        attempts = 0
        while attempts < len(active_members):
            member = next(member_cycle)
            if remaining_capacity(member, member_limits, counts) > 0:
                qa_ws[f"{assigned_col_letter}{row}"].value = member
                assignments[member].append(row)
                counts[member] += 1
                break
            attempts += 1
        else:
            # All members at capacity
            qa_ws[f"{assigned_col_letter}{row}"].value = "Backlog"
            backlog_rows.append(row)

# ---------------------------
# BALANCED (BRAND PRIORITY) MODE - Original logic
# ---------------------------
elif distribution_mode == "Balanced (Brand Priority)":
    blocks_queue = deque(blocks)
    iteration = 0
    max_iterations = 20000

    while blocks_queue and iteration < max_iterations:
        iteration += 1
        current_brand, rows = blocks_queue.popleft()
        block_size = len(rows)
        if block_size == 0:
            continue

        # preassignment
        pre_member = brand_to_member.get(current_brand)
        if pre_member and pre_member in active_members:
            cap = remaining_capacity(pre_member, member_limits, counts)
            if cap >= block_size:
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
                assignments[pre_member].extend(rows)
                counts[pre_member] += len(rows)
                assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': pre_member})
                continue
            else:
                take = min(cap, block_size)
                if take > 0:
                    for r in rows[:take]:
                        qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
                    assignments[pre_member].extend(rows[:take])
                    counts[pre_member] += take
                    assigned_blocks.append({'brand': current_brand, 'rows': rows[:take].copy(), 'member': pre_member})
                remaining = rows[take:]
                if remaining:
                    blocks_queue.appendleft((current_brand, remaining))
                continue

        if block_size < SPLIT_SIZE_THRESHOLD:
            candidates_can_take = [m for m in active_members if remaining_capacity(m, member_limits, counts) >= block_size]
            if candidates_can_take:
                candidates_can_take.sort(key=lambda m: (-remaining_capacity(m, member_limits, counts), counts[m]))
                best = candidates_can_take[0]
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = best
                assignments[best].extend(rows)
                counts[best] += block_size
                assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': best})
            else:
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                    backlog_rows.append(r)
            continue

        candidates_can_take = [m for m in active_members if remaining_capacity(m, member_limits, counts) >= block_size]
        best_candidate = candidates_can_take[0] if candidates_can_take else None
        would_imbalance = False
        if best_candidate:
            loads_after = compute_loads_after_assignment(counts, best_candidate, block_size)
            top, second = top_and_second(loads_after)
            if second == 0:
                would_imbalance = (top > 0 and second == 0)
            else:
                would_imbalance = (top > IMBALANCE_RATIO * second)

        if not best_candidate or would_imbalance:
            eligible = [m for m in active_members if remaining_capacity(m, member_limits, counts) > 0]
            rem_caps = {m: remaining_capacity(m, member_limits, counts) for m in eligible}
            total = sum(rem_caps.values())
            if total == 0:
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                    backlog_rows.append(r)
                continue

            tentative = {m: math.floor(rem_caps[m] / total * block_size) for m in eligible}
            assigned_sum = sum(tentative.values())
            remaining_to_assign = block_size - assigned_sum
            members_by_cap = sorted(eligible, key=lambda m: -rem_caps[m])
            idx = 0
            while remaining_to_assign > 0 and members_by_cap:
                m = members_by_cap[idx % len(members_by_cap)]
                if tentative[m] < rem_caps[m]:
                    tentative[m] += 1
                    remaining_to_assign -= 1
                idx += 1
                if idx > block_size * 5:
                    break

            quotas = {m: tentative[m] for m in eligible}
            member_queue = deque([m for m in eligible if quotas[m] > 0])
            while member_queue and rows:
                m = member_queue.popleft()
                if quotas[m] <= 0:
                    continue
                r = rows.pop(0)
                qa_ws[f"{assigned_col_letter}{r}"].value = m
                assignments[m].append(r)
                counts[m] += 1
                quotas[m] -= 1
                if quotas[m] > 0:
                    member_queue.append(m)
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                backlog_rows.append(r)
        else:
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = best_candidate
            assignments[best_candidate].extend(rows)
            counts[best_candidate] += block_size
            assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': best_candidate})

# ---------------------------
# BALANCED (EVEN PRIORITY) MODE - New target-aware logic
# ---------------------------
elif distribution_mode == "Balanced (Even Priority)":
    blocks_queue = deque(blocks)
    iteration = 0
    max_iterations = 20000

    while blocks_queue and iteration < max_iterations:
        iteration += 1
        current_brand, rows = blocks_queue.popleft()
        block_size = len(rows)
        if block_size == 0:
            continue

        # Handle preassignments first
        pre_member = brand_to_member.get(current_brand)
        if pre_member and pre_member in active_members:
            cap = remaining_capacity(pre_member, member_limits, counts)
            if cap >= block_size:
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
                assignments[pre_member].extend(rows)
                counts[pre_member] += len(rows)
                assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': pre_member})
                continue
            else:
                take = min(cap, block_size)
                if take > 0:
                    for r in rows[:take]:
                        qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
                    assignments[pre_member].extend(rows[:take])
                    counts[pre_member] += take
                    assigned_blocks.append({'brand': current_brand, 'rows': rows[:take].copy(), 'member': pre_member})
                remaining = rows[take:]
                if remaining:
                    blocks_queue.appendleft((current_brand, remaining))
                continue

        # Find eligible members (have capacity AND below target tolerance)
        eligible = [m for m in active_members if remaining_capacity(m, member_limits, counts) > 0]
        
        if not eligible:
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                backlog_rows.append(r)
            continue

        # Sort by distance from target (most room first)
        eligible.sort(key=lambda m: -distance_from_target(m, counts, targets))
        
        best_member = eligible[0]
        room_to_target = distance_from_target(best_member, counts, targets)
        cap = remaining_capacity(best_member, member_limits, counts)
        
        # Decision: Can we assign the whole brand without going too far over target?
        tolerance_threshold = targets[best_member] * (1 + TARGET_TOLERANCE)
        projected_count = counts[best_member] + block_size
        
        if projected_count <= tolerance_threshold and cap >= block_size:
            # Assign whole brand to best member
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = best_member
            assignments[best_member].extend(rows)
            counts[best_member] += block_size
            assigned_blocks.append({'brand': current_brand, 'rows': rows.copy(), 'member': best_member})
        else:
            # Split proportionally based on distance from target
            distances = {m: max(0, distance_from_target(m, counts, targets)) for m in eligible}
            total_distance = sum(distances.values())
            
            if total_distance == 0:
                # Everyone at or over target - split by remaining capacity
                rem_caps = {m: remaining_capacity(m, member_limits, counts) for m in eligible}
                total_cap = sum(rem_caps.values())
                if total_cap == 0:
                    for r in rows:
                        qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                        backlog_rows.append(r)
                    continue
                distances = rem_caps
                total_distance = total_cap
            
            # Calculate proportional shares
            shares = {}
            for m in eligible:
                proportion = distances[m] / total_distance if total_distance > 0 else 0
                shares[m] = min(
                    math.floor(block_size * proportion),
                    remaining_capacity(m, member_limits, counts)
                )
            
            # Distribute remainder
            assigned_so_far = sum(shares.values())
            remainder = block_size - assigned_so_far
            
            # Give remainder to members furthest from target
            members_by_distance = sorted(eligible, key=lambda m: -distances[m])
            idx = 0
            while remainder > 0 and idx < len(members_by_distance) * 3:
                m = members_by_distance[idx % len(members_by_distance)]
                if shares[m] < remaining_capacity(m, member_limits, counts):
                    shares[m] += 1
                    remainder -= 1
                idx += 1
            
            # Interleaved assignment
            member_queue = deque([m for m in eligible if shares[m] > 0])
            rows_copy = rows.copy()
            
            while member_queue and rows_copy:
                m = member_queue.popleft()
                if shares[m] <= 0:
                    continue
                r = rows_copy.pop(0)
                qa_ws[f"{assigned_col_letter}{r}"].value = m
                assignments[m].append(r)
                counts[m] += 1
                shares[m] -= 1
                if shares[m] > 0:
                    member_queue.append(m)
            
            # Any remaining go to backlog
            for r in rows_copy:
                qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                backlog_rows.append(r)

    # Post-rebalancing
    if REBALANCE_ENABLED:
        moves = rebalance_assignments(qa_ws, assignments, counts, targets, active_members, assigned_col_letter)
        if moves > 0:
            st.info(f"🔄 Post-rebalancing moved {moves} products to even out distribution.")

# ---------------------------
# Results
# ---------------------------

# Convert formulas to values
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            try:
                cell.value = cell.value
            except:
                pass

# Save output
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")
st.write(f"📄 Saved as: {output_path}")

# Summary with variance analysis
st.subheader("📊 Assignment Summary")

summary_data = []
for member in active_members:
    count = len(assignments.get(member, []))
    target = targets[member]
    limit = member_limits[member]
    variance = count - target
    variance_pct = (variance / target * 100) if target > 0 else 0
    summary_data.append({
        "Member": member,
        "Assigned": count,
        "Target": target,
        "Limit": limit,
        "Variance": variance,
        "Variance %": f"{variance_pct:+.1f}%"
    })

import pandas as pd
summary_df = pd.DataFrame(summary_data)
st.dataframe(summary_df, use_container_width=True, hide_index=True)

# Distribution metrics
assigned_counts = [len(assignments.get(m, [])) for m in active_members]
if assigned_counts:
    max_count = max(assigned_counts)
    min_count = min(assigned_counts)
    spread = max_count - min_count
    avg_count = sum(assigned_counts) / len(assigned_counts)
    
    metric_cols = st.columns(4)
    with metric_cols[0]:
        st.metric("Max Assigned", max_count)
    with metric_cols[1]:
        st.metric("Min Assigned", min_count)
    with metric_cols[2]:
        st.metric("Spread (Max-Min)", spread)
    with metric_cols[3]:
        st.metric("Backlog", len(backlog_rows))

# Visual distribution
st.subheader("📈 Distribution Chart")
try:
    import plotly.express as px
    
    chart_df = pd.DataFrame({
        "Member": active_members,
        "Assigned": [len(assignments.get(m, [])) for m in active_members],
        "Target": [targets[m] for m in active_members]
    })
    
    fig = px.bar(chart_df, x="Member", y=["Assigned", "Target"], 
                 barmode="group", 
                 title="Assigned vs Target",
                 color_discrete_map={"Assigned": "#4CAF50", "Target": "#2196F3"})
    st.plotly_chart(fig, use_container_width=True)
except ImportError:
    st.write("Install plotly for distribution charts: `pip install plotly`")

# Download
with open(output_path, "rb") as f:
    st.download_button(
        label="📥 Download Assigned Excel",
        data=f,
        file_name=output_path,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
