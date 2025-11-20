import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict, deque
import os

# ---------------------------
# Streamlit UI
# ---------------------------
st.title("QA Assignment Tool 📊 (Smart reservation + A1 rebalancing)")
st.write("Assigns products to QA team while respecting brand blocks. No splitting unless block > individual limit.")

# ---------------------------
# File upload
# ---------------------------
uploaded_file = st.file_uploader("Upload QA Template", type=["xlsx"])
if uploaded_file is not None:
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

    # Backlog mode (sort by AQ date)
    backlog_mode = st.checkbox("Backlog mode (sort by earliest AQ date)", value=False)

    # Active members input
    st.write("Enter active members today (e.g: Ross:100, Phoebe:80, Monica)")
    working_input = st.text_input("Active members")

    if working_input:
        # Parse active members and custom limits
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
                    st.warning(f"Invalid limit for {name}, using default 100")
                    active_members.append(name)
            else:
                active_members.append(part.strip().title())

        if not active_members:
            st.error("No active members entered!")
            st.stop()

        st.write(f"✅ Active members: {', '.join(active_members)}")
        if custom_limits:
            st.write("Custom limits:")
            for name, limit in custom_limits.items():
                st.write(f"- {name}: {limit}")

        # Read brand->member preassignments
        brand_to_member = {}
        for row in assignments_ws.iter_rows(min_row=2, values_only=True):
            brand, member = row[0], row[1]
            if brand and member:
                brand_to_member[brand.strip().title()] = member.strip().title()

        # Build brand blocks (preserving file order)
        brand_blocks = defaultdict(list)
        row_brand_order = []  # list of (brand_title) in first-seen order
        qa_rows = []

        # Assuming header is row 1 and data starts at row 2
        for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
            # Keep the same column references you used originally:
            m_value = row[12]   # original check
            brand = row[14]
            workflow = str(row[8]).strip() if row[8] else ""
            col_aq = row[42]

            if m_value is not None and str(m_value).strip() != "":
                qa_rows.append((i, brand, workflow, col_aq))
                if brand:
                    btitle = brand.strip().title()
                    if btitle not in brand_blocks:
                        row_brand_order.append(btitle)
                    brand_blocks[btitle].append(i)

        # Optionally sort brand blocks by AQ date inside each brand if backlog_mode
        if backlog_mode:
            st.info("🕐 Backlog mode ON — sorting all rows by earliest AQ date within brands.")

            def row_aq_date(row_idx):
                # fallback to max date if missing
                try:
                    val = qa_ws[f"AQ{row_idx}"].value
                    return val if isinstance(val, datetime) else datetime.max
                except Exception:
                    return datetime.max

            for b in brand_blocks:
                brand_blocks[b].sort(key=row_aq_date)

        # Create a list of blocks as tuples: (brand, [rows...])
        # We'll process in order of descending block size (largest first) for Option C
        blocks = []
        for b in row_brand_order:
            rows = brand_blocks[b]
            blocks.append((b, rows.copy()))

        # Sort blocks largest -> smallest (smart reservation)
        blocks.sort(key=lambda x: len(x[1]), reverse=True)

        # Set up member limits and tracking
        member_limits = {member: custom_limits.get(member, 100) for member in active_members}
        counts = {member: 0 for member in active_members}  # assigned counts
        assignments = {member: [] for member in active_members}  # row indices per member

        st.write("📏 Calculated per-member targets:")
        for m, lim in member_limits.items():
            st.write(f"- {m}: {lim} products")

        # Helper structures:
        # assigned_blocks: list of dicts {'brand':str, 'rows':[rows], 'member':str}
        # This records blocks that have been assigned so we can undo them if needed.
        assigned_blocks = []

        # Helper functions
        def remaining_capacity(member):
            return member_limits[member] - counts[member]

        def choose_member_for_block(block_size):
            """
            Choose a member who currently has remaining_capacity >= block_size.
            Tie-breakers: prefer member with largest remaining_capacity, then lowest counts.
            Returns member name or None.
            """
            candidates = [m for m in active_members if remaining_capacity(m) >= block_size]
            if not candidates:
                return None
            # sort by (remaining_capacity desc, counts asc)
            candidates.sort(key=lambda x: (-remaining_capacity(x), counts[x]))
            return candidates[0]

        def assign_block_to_member(brand, rows, member):
            """Assign the whole block (rows) to member and update tracking."""
            for r in rows:
                qa_ws[f"A{r}"].value = member
            assignments[member].extend(rows)
            counts[member] += len(rows)
            assigned_blocks.append({'brand': brand, 'rows': rows.copy(), 'member': member})
            st.write(f"Assigned brand '{brand}' ({len(rows)} rows) -> {member}")

        def unassign_block_record(block_record):
            """Unassign a previously assigned block (reset A cell) and update tracking."""
            brand = block_record['brand']
            rows = block_record['rows']
            member = block_record['member']
            for r in rows:
                qa_ws[f"A{r}"].value = None
            # remove rows from assignments[member]; counts decrement
            for r in rows:
                if r in assignments[member]:
                    assignments[member].remove(r)
            counts[member] -= len(rows)
            st.write(f"Unassigned (undo) brand '{brand}' ({len(rows)} rows) from {member}")

        # STEP 0: Apply preassigned brands from assignments sheet, but using block logic (no splitting unless overflow)
        # We'll mark those as assigned first (they are "earlier" for the undo policy).
        st.write("🔁 Applying preassigned brands (from Assignments sheet) where possible...")
        for brand, member in brand_to_member.items():
            if brand not in brand_blocks:
                continue
            if member not in active_members:
                # skip if designated member not working
                continue

            rows = brand_blocks[brand].copy()
            if not rows:
                continue
            block_size = len(rows)

            # If block fits in the remaining capacity of designated member -> assign
            cap = remaining_capacity(member)
            if block_size <= cap:
                assign_block_to_member(brand, rows, member)
                # Also remove this block from 'blocks' list so it's not processed again
                blocks = [b for b in blocks if b[0] != brand]
            else:
                # If block > member's limit, we must allow overflow behavior
                member_total_limit = member_limits[member]
                if block_size > member_total_limit:
                    # We will assign up to remaining capacity now and keep the rest as a smaller block to handle later
                    take = cap
                    if take > 0:
                        to_assign = rows[:take]
                        assign_block_to_member(brand, to_assign, member)
                    # update the brand block in blocks to the remaining rows
                    remaining = rows[take:]
                    # replace the brand's block in 'blocks' with the leftover
                    new_blocks = []
                    for bname, brows in blocks:
                        if bname == brand:
                            if remaining:
                                new_blocks.append((bname, remaining))
                        else:
                            new_blocks.append((bname, brows))
                    blocks = new_blocks
                else:
                    # block fits in member's full limit, but not in remaining capacity.
                    # For preassign step we will not force undoing; leave block for main assignment pass.
                    continue

        # Recompute blocks sorted largest->smallest after possible modifications
        blocks.sort(key=lambda x: len(x[1]), reverse=True)

        # Main assignment loop:
        st.write("🔧 Starting main assignment pass (largest -> smallest) with intelligent rebalancing...")
        # Use deque for efficient popping from left
        blocks_queue = deque(blocks)

        # This list will collect blocks that could not be assigned even after rebalancing and will be backlogged
        backlog_rows = []

        iteration = 0
        max_iterations = 20000  # safety to avoid infinite loops

        while blocks_queue and iteration < max_iterations:
            iteration += 1
            brand, rows = blocks_queue.popleft()
            block_size = len(rows)

            # First, try to find a member with remaining capacity >= block_size
            member = choose_member_for_block(block_size)

            if member:
                # fits entirely -> assign
                assign_block_to_member(brand, rows, member)
                continue

            # If no member currently has enough remaining capacity,
            # check if block size > any single member limit (overflow case)
            max_individual_limit = max(member_limits.values())
            if block_size > max_individual_limit:
                # overflow: we must allocate chunks up to remaining capacity across members (allowed only because block > limit)
                remaining_rows = rows.copy()
                assigned_any = False
                # assign to members in order of largest remaining capacity first
                members_by_capacity = sorted(active_members, key=lambda m: -remaining_capacity(m))
                for m in members_by_capacity:
                    cap = remaining_capacity(m)
                    if cap <= 0:
                        continue
                    take = min(len(remaining_rows), cap)
                    to_assign = remaining_rows[:take]
                    for r in to_assign:
                        qa_ws[f"A{r}"].value = m
                    assignments[m].extend(to_assign)
                    counts[m] += take
                    # record this partial assignment as an assigned block for undo purposes (brand remains same, rows is that chunk)
                    assigned_blocks.append({'brand': brand, 'rows': to_assign.copy(), 'member': m})
                    st.write(f"(Overflow) Assigned {len(to_assign)} of brand '{brand}' -> {m}")
                    remaining_rows = remaining_rows[take:]
                    assigned_any = True
                    if not remaining_rows:
                        break
                if remaining_rows:
                    # Could not assign all rows (no capacity left anywhere)
                    # Put remaining into backlog
                    for r in remaining_rows:
                        qa_ws[f"A{r}"].value = "Backlog"
                        backlog_rows.append(r)
                    st.write(f"Brand '{brand}' overflow remainder moved to Backlog ({len(remaining_rows)} rows).")
                continue

            # Otherwise (block_size <= max_individual_limit) and no member has remaining_capacity now:
            # We must unassign earlier assigned blocks (smallest first) until some member can take this block whole.
            st.write(f"Need to free capacity to place brand '{brand}' ({block_size} rows). Attempting undo of smallest earlier blocks...")

            # Prepare list of earlier assigned blocks (assigned_blocks) that we can undo
            # We should not undo any blocks that are preassigned and mapped to a designated member? 
            # (We allowed preassign to assign earlier only when fit; those are valid assigned_blocks too.)
            if not assigned_blocks:
                # Nothing to undo; no one has capacity -> move to backlog
                st.write(f"No earlier blocks to undo and no capacity — moving brand '{brand}' to Backlog.")
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)
                continue

            # Sort earlier assigned blocks by ascending size (A1 policy)
            # But only consider blocks that are currently assigned (i.e., still in assigned_blocks)
            assigned_sorted = sorted(enumerate(assigned_blocks), key=lambda x: len(x[1]['rows']))  # (index, record)
            freed_records = []  # will hold the blocks we unassign
            freed_total_by_member = {}  # track freed capacity by member

            placed = False
            # We'll iterate, unassigning the smallest blocks, and after each unassignment check if any member now can hold the block whole
            for idx, rec in assigned_sorted:
                # Unassign this record
                # NOTE: can't modify assigned_blocks while iterating assigned_sorted; we will collect indices to remove later
                rec_member = rec['member']
                rec_rows = rec['rows']
                # perform unassignment
                unassign_block_record(rec)
                freed_records.append(rec)
                # mark for removal by setting to None in assigned_blocks
                assigned_blocks[idx] = None

                # After unassignment, check if any member can now fit block_size
                if choose_member_for_block(block_size):
                    placed = True
                    break
            # Clean up assigned_blocks list by removing None entries
            assigned_blocks = [r for r in assigned_blocks if r is not None]

            if not placed:
                # Even after undoing everything, block still cannot fit (all members full or block too big)
                # If block_size > max_individual_limit we should have hit overflow earlier, so here it's that nobody has capacity
                # Move the block to backlog
                st.write(f"Even after undoing earlier blocks, brand '{brand}' cannot be placed. Moving to backlog.")
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)
                # The freed_records should be re-queued for assignment
                for rec in freed_records:
                    # push these freed blocks back to the queue to be assigned later
                    blocks_queue.append((rec['brand'], rec['rows']))
                continue
            else:
                # We've freed enough capacity to place the current block.
                chosen_member = choose_member_for_block(block_size)
                if chosen_member is None:
                    # defensive check — should not happen
                    st.write("Unexpected: no member found after freeing. Moving to backlog.")
                    for r in rows:
                        qa_ws[f"A{r}"].value = "Backlog"
                        backlog_rows.append(r)
                    # requeue freed_records
                    for rec in freed_records:
                        blocks_queue.append((rec['brand'], rec['rows']))
                    continue

                # Assign current block to chosen_member
                assign_block_to_member(brand, rows, chosen_member)

                # Requeue freed records for reassignment (we push them to left to assign them sooner)
                # But to minimize disturbance, requeue them sorted smallest->largest (they were freed in that order)
                for rec in freed_records:
                    # rec may have been partially previously handled; simply append to left so they get assigned soon
                    blocks_queue.appendleft((rec['brand'], rec['rows']))

                # continue main loop

        # After main loop, any blocks remaining in queue were not handled due to iterations cap; move to backlog as safe fallback
        if blocks_queue:
            st.write("⚠ Reached iteration cap or leftover blocks — moving remaining blocks to backlog.")
            while blocks_queue:
                brand, rows = blocks_queue.popleft()
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)

        # Final pass: ensure no formulas left (convert formulas to values if data_type == 'f')
        for row in qa_ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    cell.value = cell.value

        # Save output
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = f"QA_Assignment_{timestamp}.xlsx"
        wb.save(output_path)

        st.success("✅ Assignment complete!")
        st.write(f"📄 Saved as: {output_path}")

        # Summary
        st.write("📊 Summary of assignments:")
        for member, rows in assignments.items():
            limit = member_limits[member]
            st.write(f"- {member}: {len(rows)} products (Target: {limit})")
        st.write(f"- Backlog (explicitly set): {len(backlog_rows)} products")

        # Provide download
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Download Assigned Excel",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
