import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict

st.title("Product Assignment Tool")

# --- File upload ---
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file is not None:

    # --- Load workbook ---
    wb = load_workbook(uploaded_file)
    
    if "QA" not in wb.sheetnames or "Assignment" not in wb.sheetnames:
        st.error("❌ The workbook must have 'QA' and 'Assignment' sheets.")
    else:
        qa_ws = wb["QA"]
        assign_ws = wb["Assignment"]

        # --- Build assignment mapping (Brand → Team Member) ---
        brand_to_member = {}
        for row in assign_ws.iter_rows(min_row=2, values_only=True):
            brand, member = row[0], row[1]
            if brand and member:
                brand_to_member[brand.strip().title()] = member.strip().title()

        # --- Determine active team members ---
        team_members = list(set(brand_to_member.values()))
        num_members = len(team_members)
        if num_members == 0:
            st.error("❌ No active team members found in Assignment tab!")
        else:

            # --- Count products per brand ---
            brand_rows = defaultdict(list)
            for row in qa_ws.iter_rows(min_row=2):
                brand = row[14].value  # Column O
                if brand:
                    brand_rows[str(brand).strip().title()].append(row)

            # --- Custom product limits ---
            st.subheader("Custom Product Limits (Optional)")
            st.write("Format: Name:Limit, separated by commas. Example: Alice:10,Bob:15")
            custom_input = st.text_input("Enter custom limits:")

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

            # --- Calculate even split for those without custom limits ---
            total_products = sum(len(rows) for rows in brand_rows.values())
            assigned_counts = {member: 0 for member in team_members}

            even_split = total_products // num_members
            remainder = total_products % num_members
            for i, member in enumerate(team_members):
                if member not in custom_limits:
                    custom_limits[member] = even_split + (1 if i < remainder else 0)

            # --- Helper: remaining capacity ---
            def remaining_capacity(member):
                return custom_limits.get(member, 0) - assigned_counts.get(member, 0)

            # --- Assign brands ---
            backlog_count = 0
            for brand, rows in brand_rows.items():
                member = brand_to_member.get(brand)
                if not member or remaining_capacity(member) <= 0:
                    for r in rows:
                        r[0].value = "Backlog"
                    backlog_count += len(rows)
                    continue

                if len(rows) <= remaining_capacity(member):
                    for r in rows:
                        r[0].value = member
                    assigned_counts[member] += len(rows)
                else:
                    for r in rows[:remaining_capacity(member)]:
                        r[0].value = member
                    for r in rows[remaining_capacity(member):]:
                        r[0].value = "Backlog"
                        backlog_count += 1
                    assigned_counts[member] += remaining_capacity(member)

            # --- Save timestamped file ---
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            output_path = f"Assigned_{timestamp}.xlsx"
            wb.save(output_path)

            # --- Show summary ---
            st.success("✅ Assignment complete!")
            st.write(f"📄 Saved as: `{output_path}`")

            st.subheader("📊 Summary of assignments")
            for member, count in assigned_counts.items():
                limit = custom_limits.get(member)
                st.write(f"- {member}: {count} products (Limit: {limit})")
            st.write(f"- Backlog: {backlog_count} products")

            # --- Download button ---
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Assigned File",
                    data=f,
                    file_name=output_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
