import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict
import os

# --- Streamlit UI ---
st.title("QA Assignment Tool 📊")
st.write("Auto Assigns Products to QA Team")

# --- File upload ---
uploaded_file = st.file_uploader("Upload QA Template", type=["xlsx"])
if uploaded_file is not None:
    # Save to a temp file
    temp_file_path = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    with open(temp_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # --- Load workbook ---
    wb = load_workbook(temp_file_path)
    if "QA" not in wb.sheetnames or "Assignments" not in wb.sheetnames:
        st.error("Excel file must contain 'QA' and 'Assignments' sheets.")
        st.stop()
        
    qa_ws = wb["QA"]
    assignments_ws = wb["Assignments"]

    # --- Backlog mode ---
    backlog_mode = st.checkbox("Backlog mode (sort by earliest AQ date)", value=False)

    # --- Active members input ---
    st.write("Enter active members today (e.g: Ross, Phoebe, Monica: 20) Press Enter once added")
    working_input = st.text_input("Active members")
    
    if working_input:
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

        # --- Set member limits ---
        member_limits = {member: custom_limits.get(member, 100) for member in active_members}
        assignments = {member: [] for member in active_members}
        counts = {member: 0 for member in active_members}

        st.write("📏 Calculated per-member targets:")
        for member, limit in member_limits.items():
            st.write(f"- {member}: {limit} products")

        # --- Step 1: Pre-assign brands from Assignments sheet ---
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
                    if counts[member] < member_limits[member]:
                        qa_ws[f"A{r}"].value = member
                        assignments[member].append(r)
                        counts[member] += 1
                    else:
                        qa_ws[f"A{r}"].value = "Backlog"

        # --- Step 2: Assign remaining brand blocks ---
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
                remaining_rows = list(unassigned)
                eligible_members = sorted(active_members, key=lambda x: counts[x])
                for member in eligible_members[:2]:
                    capacity = member_limits[member] - counts[member]
                    if capacity <= 0:
                        continue
                    to_assign = remaining_rows[:capacity]
                    for r in to_assign:
                        qa_ws[f"A{r}"].value = member
                        assignments[member].append(r)
                        counts[member] += 1
                    remaining_rows = remaining_rows[capacity:]
                    if not remaining_rows:
                        break
                for r in remaining_rows:
                    qa_ws[f"A{r}"].value = "Backlog"

        # --- Convert formulas to values ---
        for row in qa_ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    cell.value = cell.value

        # --- Save timestamped file ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = f"QA_Assignment_{timestamp}.xlsx"
        wb.save(output_path)

        st.success("✅ Assignment complete!")
        st.write(f"📄 Saved as: {output_path}")

        # --- Summary ---
        st.write("📊 Summary of assignments:")
        for member, rows in assignments.items():
            limit = member_limits[member]
            st.write(f"- {member}: {len(rows)} products (Target: {limit})")
        backlog_count = sum(1 for r in qa_rows if qa_ws[f"A{r[0]}"].value == "Backlog")
        st.write(f"- Backlog: {backlog_count} products")

        # --- Download link ---
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Download Assigned Excel",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



