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
st.write("Assigns products to QA team with smart distribution.")

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


def distance_from_target(member, counts, targets):
    """How far is this member from their target? Positive = room to add."""
    return targets.get(member, 0) - counts.get(member, 0)


def calculate_targets(active_members, member_limits, total_products):
    """Calculate ideal target for each member based on their limits and total products."""
    total_capacity = sum(member_limits[m] for m in active_members)
    
    if total_capacity == 0:
        return {m: 0 for m in active_members}
    
    targets = {}
    for m in active_members:
        proportion = member_limits[m] / total_capacity
        ideal = round(total_products * proportion)
        targets[m] = min(ideal, member_limits[m])
    
    # Adjust to hit total exactly
    assigned = sum(targets.values())
    diff = total_products - assigned
    
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
    
    while diff < 0 and i < len(members_by_slack) * 2:
        m = members_by_slack[i % len(members_by_slack)]
        if targets[m] > 0:
            targets[m] -= 1
            diff += 1
        i += 1
    
    return targets


def get_best_member_for_brand(brand_size, active_members, counts, targets, member_limits, max_overshoot_pct):
    """
    Find the best member to assign a whole brand to.
    Prioritises members furthest from target who can take the brand without excessive overshoot.
    """
    candidates = []
    
    for m in active_members:
        cap = remaining_capacity(m, member_limits, counts)
        if cap <= 0:
            continue
            
        dist = distance_from_target(m, counts, targets)
        projected = counts[m] + brand_size
        target = targets[m]
        
        # Calculate overshoot
        if target > 0:
            overshoot_pct = (projected - target) / target if projected > target else 0
        else:
            overshoot_pct = float('inf') if projected > 0 else 0
        
        can_fit = cap >= brand_size
        within_tolerance = overshoot_pct <= max_overshoot_pct
        
        candidates.append({
            'member': m,
            'distance': dist,
            'cap': cap,
            'can_fit': can_fit,
            'within_tolerance': within_tolerance,
            'overshoot_pct': overshoot_pct
        })
    
    if not candidates:
        return None, False
    
    # Priority 1: Can fit whole brand within tolerance
    fitting = [c for c in candidates if c['can_fit'] and c['within_tolerance']]
    if fitting:
        # Among those, pick the one furthest from target
        fitting.sort(key=lambda c: -c['distance'])
        return fitting[0]['member'], True
    
    # Priority 2: Can fit whole brand (even if over tolerance)
    fitting = [c for c in candidates if c['can_fit']]
    if fitting:
        # Pick the one with least overshoot
        fitting.sort(key=lambda c: c['overshoot_pct'])
        return fitting[0]['member'], True
    
    # Priority 3: Nobody can fit whole brand - will need to split
    # Return the one with most capacity
    candidates.sort(key=lambda c: -c['cap'])
    return candidates[0]['member'], False


def smart_rebalance_by_brands(qa_ws, assignments, counts, targets, brand_assignments, 
                               active_members, assigned_col_letter, max_iterations=500):
    """
    Rebalance by moving whole small brands from over-target to under-target members.
    This maintains brand integrity during rebalancing.
    """
    moves_made = 0
    iterations = 0
    
    # Build brand -> rows mapping from current assignments
    member_brands = defaultdict(list)  # member -> list of (brand, [rows])
    for member in active_members:
        brands_for_member = defaultdict(list)
        for row in assignments[member]:
            brand = brand_assignments.get(row, "Unknown")
            brands_for_member[brand].append(row)
        for brand, rows in brands_for_member.items():
            member_brands[member].append((brand, rows))
    
    while iterations < max_iterations:
        iterations += 1
        
        # Find most over and under target
        over = [(m, counts[m] - targets[m]) for m in active_members if counts[m] > targets[m]]
        under = [(m, targets[m] - counts[m]) for m in active_members if counts[m] < targets[m]]
        
        if not over or not under:
            break
        
        over.sort(key=lambda x: -x[1])
        under.sort(key=lambda x: -x[1])
        
        from_member = over[0][0]
        to_member = under[0][0]
        
        # Current spread
        current_spread = counts[from_member] - counts[to_member]
        if current_spread <= 1:
            break
        
        # Find a small brand to move that would improve balance
        best_brand_to_move = None
        best_improvement = 0
        
        for brand, rows in member_brands[from_member]:
            brand_size = len(rows)
            
            # Would moving this brand improve the spread?
            new_from_count = counts[from_member] - brand_size
            new_to_count = counts[to_member] + brand_size
            
            # Check capacity
            if new_to_count > member_limits[to_member]:
                continue
            
            new_spread = abs(new_from_count - new_to_count)
            improvement = current_spread - new_spread
            
            # Only move if it improves things and doesn't flip the imbalance too much
            if improvement > 0 and new_to_count <= new_from_count + 2:
                if improvement > best_improvement:
                    best_improvement = improvement
                    best_brand_to_move = (brand, rows)
        
        if best_brand_to_move:
            brand, rows = best_brand_to_move
            
            # Move the brand
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = to_member
                assignments[from_member].remove(r)
                assignments[to_member].append(r)
            
            counts[from_member] -= len(rows)
            counts[to_member] += len(rows)
            
            # Update brand tracking
            member_brands[from_member].remove(best_brand_to_move)
            member_brands[to_member].append(best_brand_to_move)
            
            moves_made += 1
        else:
            # No whole brand can be moved to improve balance
            break
    
    return moves_made


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

with st.expander("🔧 Advanced Settings"):
    col1, col2 = st.columns(2)
    with col1:
        MAX_OVERSHOOT_PCT = st.slider(
            "Max Overshoot % (before splitting)", 
            0.0, 0.5, 0.15, 0.05,
            help="How much over target (%) before forcing a brand split. Lower = more even, but more splits."
        )
        FORCE_SPLIT_THRESHOLD = st.slider(
            "Force Split Spread Threshold",
            1, 20, 5, 1,
            help="If spread exceeds this, force split remaining brands regardless of size."
        )
    with col2:
        REBALANCE_BRANDS = st.checkbox(
            "Rebalance by moving whole brands", 
            value=True,
            help="After assignment, try to move small brands to even out distribution."
        )
        SHOW_BRAND_DETAILS = st.checkbox("Show brand assignment details", value=False)

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

# Preassignments from Assignments sheet
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand = row[COL_ASSIGN_BRAND - 1] if len(row) >= COL_ASSIGN_BRAND else None
    member = row[COL_ASSIGN_QAER - 1] if len(row) >= COL_ASSIGN_QAER else None
    if brand and member:
        brand_to_member[title_or_none(brand)] = title_or_none(member)

# Show pre-assignments
if brand_to_member:
    with st.expander(f"📌 Pre-assigned Brands ({len(brand_to_member)})"):
        for brand, member in sorted(brand_to_member.items()):
            status = "✅" if member in active_members else "⚠️ (not active today)"
            st.write(f"- {brand} → {member} {status}")

# Build brand blocks
brand_blocks = defaultdict(list)
row_brand_order = []
row_to_brand = {}  # Track which brand each row belongs to

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    pim_parent_id = row[COL_PIM_PARENT_ID - 1] if len(row) >= COL_PIM_PARENT_ID else None
    brand = row[COL_BRAND - 1] if len(row) >= COL_BRAND else None

    if pim_parent_id is not None and str(pim_parent_id).strip():
        btitle = title_or_none(brand) if brand else "No Brand"
        row_to_brand[i] = btitle
        if btitle not in brand_blocks:
            row_brand_order.append(btitle)
        brand_blocks[btitle].append(i)

if backlog_mode:
    date_col_letter = get_column_letter(COL_BT_IMAGE_DATE)

    def row_date(row_idx):
        try:
            val = qa_ws[f"{date_col_letter}{row_idx}"].value
        except:
            val = None
        return val if isinstance(val, datetime) else datetime.max

    for b in brand_blocks:
        brand_blocks[b].sort(key=row_date)

# Build blocks list
blocks = []
for b in row_brand_order:
    blocks.append((b, brand_blocks[b].copy()))

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

# ---------------------------
# SMART EVEN SPLIT ALGORITHM
# ---------------------------
# Strategy:
# 1. Process pre-assigned brands first (to their assigned member)
# 2. Sort remaining brands by size (largest first - harder to place)
# 3. Assign whole brands to member furthest from target
# 4. Only split if necessary to maintain balance
# 5. Rebalance by moving whole small brands

counts = {m: 0 for m in active_members}
assignments = {m: [] for m in active_members}
brand_assignments_log = []  # Track for summary
backlog_rows = []
assigned_col_letter = get_column_letter(COL_ASSIGNED)

# Separate pre-assigned and unassigned brands
preassigned_blocks = []
unassigned_blocks = []

for brand, rows in blocks:
    pre_member = brand_to_member.get(brand)
    if pre_member and pre_member in active_members:
        preassigned_blocks.append((brand, rows, pre_member))
    else:
        unassigned_blocks.append((brand, rows))

# Sort unassigned by size (largest first)
unassigned_blocks.sort(key=lambda x: -len(x[1]))

# ---------------------------
# PHASE 1: Process pre-assigned brands
# ---------------------------
for brand, rows, pre_member in preassigned_blocks:
    block_size = len(rows)
    cap = remaining_capacity(pre_member, member_limits, counts)
    
    if cap >= block_size:
        # Assign all to pre-assigned member
        for r in rows:
            qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
        assignments[pre_member].extend(rows)
        counts[pre_member] += block_size
        brand_assignments_log.append({
            'brand': brand, 
            'size': block_size, 
            'member': pre_member, 
            'split': False,
            'preassigned': True
        })
    else:
        # Partial assignment to pre-assigned member
        if cap > 0:
            for r in rows[:cap]:
                qa_ws[f"{assigned_col_letter}{r}"].value = pre_member
            assignments[pre_member].extend(rows[:cap])
            counts[pre_member] += cap
            brand_assignments_log.append({
                'brand': brand, 
                'size': cap, 
                'member': pre_member, 
                'split': True,
                'preassigned': True
            })
        
        # Remaining goes to others or backlog
        remaining_rows = rows[cap:]
        # Add to unassigned for processing
        if remaining_rows:
            unassigned_blocks.append((brand + " (overflow)", remaining_rows))

# Re-sort unassigned after adding overflow
unassigned_blocks.sort(key=lambda x: -len(x[1]))

# ---------------------------
# PHASE 2: Process unassigned brands (largest first)
# ---------------------------
for brand, rows in unassigned_blocks:
    block_size = len(rows)
    
    if block_size == 0:
        continue
    
    # Calculate current spread
    current_counts = list(counts.values())
    current_spread = max(current_counts) - min(current_counts) if current_counts else 0
    
    # Adjust overshoot tolerance based on current spread
    # If spread is already high, be more aggressive about splitting
    dynamic_overshoot = MAX_OVERSHOOT_PCT
    if current_spread > FORCE_SPLIT_THRESHOLD:
        dynamic_overshoot = 0.05  # Very strict - force more splits
    
    # Find best member for this brand
    best_member, can_fit_whole = get_best_member_for_brand(
        block_size, active_members, counts, targets, member_limits, dynamic_overshoot
    )
    
    if best_member is None:
        # No capacity anywhere
        for r in rows:
            qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
            backlog_rows.append(r)
        continue
    
    if can_fit_whole:
        # Assign whole brand to best member
        for r in rows:
            qa_ws[f"{assigned_col_letter}{r}"].value = best_member
        assignments[best_member].extend(rows)
        counts[best_member] += block_size
        brand_assignments_log.append({
            'brand': brand, 
            'size': block_size, 
            'member': best_member, 
            'split': False,
            'preassigned': False
        })
    else:
        # Need to split - distribute proportionally by distance from target
        eligible = [m for m in active_members if remaining_capacity(m, member_limits, counts) > 0]
        
        if not eligible:
            for r in rows:
                qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                backlog_rows.append(r)
            continue
        
        # Calculate shares based on distance from target
        distances = {m: max(0, distance_from_target(m, counts, targets)) for m in eligible}
        total_distance = sum(distances.values())
        
        if total_distance == 0:
            # Everyone at or over target - split by remaining capacity
            caps = {m: remaining_capacity(m, member_limits, counts) for m in eligible}
            total_cap = sum(caps.values())
            if total_cap == 0:
                for r in rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
                    backlog_rows.append(r)
                continue
            distances = caps
            total_distance = total_cap
        
        # Calculate proportional shares
        shares = {}
        for m in eligible:
            proportion = distances[m] / total_distance
            shares[m] = min(
                math.floor(block_size * proportion),
                remaining_capacity(m, member_limits, counts)
            )
        
        # Distribute remainder
        assigned_so_far = sum(shares.values())
        remainder = block_size - assigned_so_far
        members_by_distance = sorted(eligible, key=lambda m: -distances[m])
        
        idx = 0
        while remainder > 0 and idx < len(members_by_distance) * 3:
            m = members_by_distance[idx % len(members_by_distance)]
            if shares[m] < remaining_capacity(m, member_limits, counts):
                shares[m] += 1
                remainder -= 1
            idx += 1
        
        # Assign rows
        rows_copy = rows.copy()
        for m in eligible:
            if shares[m] > 0:
                member_rows = rows_copy[:shares[m]]
                rows_copy = rows_copy[shares[m]:]
                
                for r in member_rows:
                    qa_ws[f"{assigned_col_letter}{r}"].value = m
                assignments[m].extend(member_rows)
                counts[m] += len(member_rows)
                
                brand_assignments_log.append({
                    'brand': brand, 
                    'size': len(member_rows), 
                    'member': m, 
                    'split': True,
                    'preassigned': False
                })
        
        # Any remaining to backlog
        for r in rows_copy:
            qa_ws[f"{assigned_col_letter}{r}"].value = "Backlog"
            backlog_rows.append(r)

# ---------------------------
# PHASE 3: Rebalance by moving whole brands
# ---------------------------
rebalance_moves = 0
if REBALANCE_BRANDS:
    rebalance_moves = smart_rebalance_by_brands(
        qa_ws, assignments, counts, targets, row_to_brand,
        active_members, assigned_col_letter
    )
    if rebalance_moves > 0:
        st.info(f"🔄 Rebalancing moved {rebalance_moves} whole brand(s) to improve distribution.")

# ---------------------------
# Save and display results
# ---------------------------

# Convert formulas to values
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            try:
                cell.value = cell.value
            except:
                pass

timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")

# Summary
st.subheader("📊 Assignment Summary")

import pandas as pd

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

summary_df = pd.DataFrame(summary_data)
st.dataframe(summary_df, use_container_width=True, hide_index=True)

# Metrics
assigned_counts = [len(assignments.get(m, [])) for m in active_members]
if assigned_counts:
    max_count = max(assigned_counts)
    min_count = min(assigned_counts)
    spread = max_count - min_count
    
    metric_cols = st.columns(4)
    with metric_cols[0]:
        st.metric("Max Assigned", max_count)
    with metric_cols[1]:
        st.metric("Min Assigned", min_count)
    with metric_cols[2]:
        st.metric("Spread", spread, delta=None, delta_color="inverse")
    with metric_cols[3]:
        st.metric("Backlog", len(backlog_rows))

# Brand integrity stats
brands_kept_whole = sum(1 for b in brand_assignments_log if not b['split'])
brands_split = sum(1 for b in brand_assignments_log if b['split'])
preassigned_count = sum(1 for b in brand_assignments_log if b['preassigned'])

st.write("**Brand Integrity:**")
integrity_cols = st.columns(3)
with integrity_cols[0]:
    st.metric("Brands Kept Whole", brands_kept_whole)
with integrity_cols[1]:
    st.metric("Brands Split", brands_split)
with integrity_cols[2]:
    st.metric("Pre-assigned Brands", preassigned_count)

# Chart
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
    st.write("Install plotly for charts: `pip install plotly`")

# Brand details
if SHOW_BRAND_DETAILS:
    st.subheader("📋 Brand Assignment Details")
    brand_df = pd.DataFrame(brand_assignments_log)
    if not brand_df.empty:
        brand_df = brand_df.rename(columns={
            'brand': 'Brand',
            'size': 'Products',
            'member': 'Assigned To',
            'split': 'Was Split',
            'preassigned': 'Pre-assigned'
        })
        st.dataframe(brand_df, use_container_width=True, hide_index=True)

# Download
with open(output_path, "rb") as f:
    st.download_button(
        label="📥 Download Assigned Excel",
        data=f,
        file_name=output_path,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
