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
                    active_members.append(name)
            else:
                active_members.append(part.strip().title())

        if not active_members:
            st.error("No active members entered!")
            st.stop()

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

        for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
            m_value = row[12]
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

        if backlog_mode:
            def row_aq_date(row_idx):
                try:
                    val = qa_ws[f"AQ{row_idx}"].value
                    return val if isinstance(val, datetime) else datetime.max
                except Exception:
                    return datetime.max
            for b in brand_blocks:
                brand_blocks[b].sort(key=row_aq_date)

        blocks = []
        for b in row_brand_order:
            rows = brand_blocks[b]
            blocks.append((b, rows.copy()))

        # Sort blocks largest -> smallest for smart reservation
        blocks.sort(key=lambda x: len(x[1]), reverse=True)

        member_limits = {member: custom_limits.get(member, 100) for member in active_members}
        counts = {member: 0 for member in active_members}
        assignments = {member: [] for member in active_members}

        assigned_blocks = []

        def remaining_capacity(member):
            return member_limits[member] - counts[member]

        def choose_member_for_block(block_size):
            candidates = [m for m in active_members if remaining_capacity(m) >= block_size]
            if not candidates:
                return None
            candidates.sort(key=lambda x: (-remaining_capacity(x), counts[x]))
            return candidates[0]

        def assign_block_to_member(brand, rows, member):
            for r in rows:
                qa_ws[f"A{r}"].value = member
            assignments[member].extend(rows)
            counts[member] += len(rows)
            assigned_blocks.append({'brand': brand, 'rows': rows.copy(), 'member': member})

        def unassign_block_record(block_record):
            rows = block_record['rows']
            member = block_record['member']
            for r in rows:
                qa_ws[f"A{r}"].value = None
            for r in rows:
                if r in assignments[member]:
                    assignments[member].remove(r)
            counts[member] -= len(rows)

        # STEP 0: Apply preassigned brands
        for brand, member in brand_to_member.items():
            if brand not in brand_blocks:
                continue
            if member not in active_members:
                continue
            rows = brand_blocks[brand].copy()
            if not rows:
                continue
            block_size = len(rows)
            cap = remaining_capacity(member)
            if block_size <= cap:
                assign_block_to_member(brand, rows, member)
                blocks = [b for b in blocks if b[0] != brand]
            else:
                member_total_limit = member_limits[member]
                if block_size > member_total_limit:
                    take = cap
                    if take > 0:
                        assign_block_to_member(brand, rows[:take], member)
                    remaining = rows[take:]
                    new_blocks = []
                    for bname, brows in blocks:
                        if bname == brand:
                            if remaining:
                                new_blocks.append((bname, remaining))
                        else:
                            new_blocks.append((bname, brows))
                    blocks = new_blocks
                else:
                    continue

        blocks.sort(key=lambda x: len(x[1]), reverse=True)
        blocks_queue = deque(blocks)
        backlog_rows = []

        iteration = 0
        max_iterations = 20000

        while blocks_queue and iteration < max_iterations:
            iteration += 1
            brand, rows = blocks_queue.popleft()
            block_size = len(rows)
            member = choose_member_for_block(block_size)

            if member:
                assign_block_to_member(brand, rows, member)
                continue

            max_individual_limit = max(member_limits.values())
            if block_size > max_individual_limit:
                remaining_rows = rows.copy()
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
                    assigned_blocks.append({'brand': brand, 'rows': to_assign.copy(), 'member': m})
                    remaining_rows = remaining_rows[take:]
                    if not remaining_rows:
                        break
                if remaining_rows:
                    for r in remaining_rows:
                        qa_ws[f"A{r}"].value = "Backlog"
                        backlog_rows.append(r)
                continue

            if not assigned_blocks:
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)
                continue

            assigned_sorted = sorted(enumerate(assigned_blocks), key=lambda x: len(x[1]['rows']))
            freed_records = []
            placed = False

            for idx, rec in assigned_sorted:
                unassign_block_record(rec)
                freed_records.append(rec)
                assigned_blocks[idx] = None
                if choose_member_for_block(block_size):
                    placed = True
                    break

            assigned_blocks = [r for r in assigned_blocks if r is not None]

            if not placed:
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)
                for rec in freed_records:
                    blocks_queue.append((rec['brand'], rec['rows']))
                continue
            else:
                chosen_member = choose_member_for_block(block_size)
                if chosen_member is None:
                    for r in rows:
                        qa_ws[f"A{r}"].value = "Backlog"
                        backlog_rows.append(r)
                    for rec in freed_records:
                        blocks_queue.append((rec['brand'], rec['rows']))
                    continue

                assign_block_to_member(brand, rows, chosen_member)

                for rec in freed_records:
                    blocks_queue.appendleft((rec['brand'], rec['rows']))

        if blocks_queue:
            while blocks_queue:
                brand, rows = blocks_queue.popleft()
                for r in rows:
                    qa_ws[f"A{r}"].value = "Backlog"
                    backlog_rows.append(r)

        # Final pass: convert formulas to values
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

        # Download button
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Download Assigned Excel",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
