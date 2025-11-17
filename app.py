import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="QA Assignment", layout="wide")
st.title("📊 QA Assignment")

# --- Upload workbook ---
uploaded_file = st.file_uploader("Upload QA Excel file", type=["xlsx"])
if uploaded_file is None:
    st.info("Please upload an Excel file to proceed.")
    st.stop()

# --- Load sheets directly with pandas ---
qa_df = pd.read_excel(uploaded_file, sheet_name="QA")
mp_df = pd.read_excel(uploaded_file, sheet_name="MP")
assignments_df = pd.read_excel(uploaded_file, sheet_name="Assignments")

# --- Streamlit inputs ---
backlog_mode = st.radio("Are you expecting to be in backlog today?", ["No", "Yes"]) == "Yes"

absent_input = st.text_input("Type names of absent members (comma-separated)").strip()
absent_list = [name.strip().title() for name in absent_input.split(",") if name.strip()]
if absent_list:
    st.warning(f"🟡 Absent today: {', '.join(absent_list)}")
else:
    st.success("✅ Everyone is present.")

custom_input = st.text_input("Specify custom product limits (Name:Limit, comma-separated)").strip()
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

# --- Preferences ---
preferences = {}
for _, row in mp_df.iterrows():
    name, divs = row[0], row[1]
    if name and divs:
        div_list = [d.strip().title() for d in str(divs).split(",") if d.strip()]
        preferences[name] = div_list

# --- Filter absentees ---
active_preferences = {name: divs for name, divs in preferences.items() if name not in absent_list}
team_members = list(active_preferences.keys())
num_members = len(team_members)
if num_members == 0:
    st.error("❌ No active team members available for assignment!")
    st.stop()

# --- Prepare QA data ---
qa_df['Assigned'] = None
qa_df['PriorityOverride'] = qa_df.iloc[:, 32].combine_first(qa_df.iloc[:, 33])
qa_df['AQDate'] = qa_df.iloc[:, 42]
qa_df['Division'] = qa_df.iloc[:, 17].astype(str).str.strip().str.title()
qa_df['Brand'] = qa_df.iloc[:, 14].astype(str).str.strip().str.title()
qa_df['Workflow'] = qa_df.iloc[:, 8].astype(str).str.strip()

if backlog_mode:
    qa_df.sort_values('AQDate', inplace=True)

DEFAULT_TARGET = 100
counts = {m: 0 for m in team_members}

def member_limit(member):
    return custom_limits.get(member, DEFAULT_TARGET)

# --- Filter only rows with Column M populated ---
product_df = qa_df[qa_df.iloc[:, 12].notna()].copy()
total_products = len(product_df)

# --- Compute perfectly fair targets per member ---
base_target = total_products // num_members
remainder = total_products % num_members

# Initialize ideal targets
ideal_targets = {m: min(member_limit(m), base_target) for m in team_members}

# Distribute remainder fairly
sorted_members = sorted(team_members)
for i in range(remainder):
    member = sorted_members[i % len(team_members)]
    ideal_targets[member] = min(member_limit(member), ideal_targets[member] + 1)

def eligible_members():
    return [m for m in team_members if counts[m] < ideal_targets[m]]

# --- Read Assignments tab and map Brand -> Member ---
brand_to_member = {}
for _, row in assignments_df.iterrows():
    brand = str(row[0]).strip().title()
    member = str(row[1]).strip().title()
    if brand and member:
        brand_to_member[brand] = member

# --- Step 1: Assign based on Assignments tab ---
for brand, assigned_member in brand_to_member.items():
    if assigned_member not in team_members:
        # Assigned member is absent or invalid
        continue
    
    brand_mask = product_df['Brand'] == brand
    brand_rows = product_df[brand_mask & product_df['Assigned'].isna()].index.tolist()
    brand_size = len(brand_rows)
    
    remaining_capacity = ideal_targets[assigned_member] - counts[assigned_member]
    
    if remaining_capacity >= brand_size:
        product_df.loc[brand_rows, 'Assigned'] = assigned_member
        counts[assigned_member] += brand_size
    else:
        # Assign as much as possible to the assigned_member
        to_assign = brand_rows[:remaining_capacity]
        product_df.loc[to_assign, 'Assigned'] = assigned_member
        counts[assigned_member] += len(to_assign)
        
        # Remaining products assigned to other eligible members
        remaining_rows = brand_rows[remaining_capacity:]
        for idx in remaining_rows:
            eligible = eligible_members()
            if not eligible:
                product_df.at[idx, 'Assigned'] = "Backlog"
                continue
            chosen = min(eligible, key=lambda m: counts[m])
            product_df.at[idx, 'Assigned'] = chosen
            counts[chosen] += 1

# --- Step 2: Assign remaining products using brand blocks + preferences ---
brands_remaining = product_df.loc[product_df['Assigned'].isna(), 'Brand'].unique()
for brand in brands_remaining:
    brand_mask = product_df['Brand'] == brand
    brand_rows = product_df[brand_mask & product_df['Assigned'].isna()].index.tolist()
    brand_size = len(brand_rows)
    
    eligible = eligible_members()
    if not eligible:
        product_df.loc[brand_rows, 'Assigned'] = "Backlog"
        continue
    
    # Prefer members who like any division in this brand
    brand_divisions = product_df.loc[brand_rows, 'Division'].unique()
    pref_eligible = [m for m in eligible if any(div in active_preferences.get(m, []) for div in brand_divisions)]
    
    if pref_eligible:
        chosen = min(pref_eligible, key=lambda m: counts[m])
    else:
        chosen = min(eligible, key=lambda m: counts[m])
    
    remaining_capacity = ideal_targets[chosen] - counts[chosen]
    if remaining_capacity >= brand_size:
        product_df.loc[brand_rows, 'Assigned'] = chosen
        counts[chosen] += brand_size
    else:
        to_assign = brand_rows[:remaining_capacity]
        product_df.loc[to_assign, 'Assigned'] = chosen
        counts[chosen] += len(to_assign)
        
        remaining_rows = brand_rows[remaining_capacity:]
        for idx in remaining_rows:
            eligible = eligible_members()
            if not eligible:
                product_df.at[idx, 'Assigned'] = "Backlog"
                continue
            chosen = min(eligible, key=lambda m: counts[m])
            product_df.at[idx, 'Assigned'] = chosen
            counts[chosen] += 1

# --- Step 3: Assign remaining priority override products ---
priority_mask = product_df['Assigned'].isna() & product_df['PriorityOverride'].notna()
for idx in product_df[priority_mask].index:
    eligible = eligible_members()
    if eligible:
        chosen = min(eligible, key=lambda m: counts[m])
        product_df.at[idx, 'Assigned'] = chosen
        counts[chosen] += 1
    else:
        product_df.at[idx, 'Assigned'] = "Backlog"

# --- Step 4: Any remaining unassigned products ---
remaining_mask = product_df['Assigned'].isna()
for idx in product_df[remaining_mask].index:
    eligible = eligible_members()
    if eligible:
        chosen = min(eligible, key=lambda m: counts[m])
        product_df.at[idx, 'Assigned'] = chosen
        counts[chosen] += 1
    else:
        product_df.at[idx, 'Assigned'] = "Backlog"

# --- Write assignments back ---
qa_df['Assigned'] = qa_df.index.map(lambda idx: product_df.at[idx, 'Assigned'] if idx in product_df.index else "")

# --- Download ---
output_buffer = BytesIO()
with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
    qa_df.to_excel(writer, sheet_name="QA", index=False)
    mp_df.to_excel(writer, sheet_name="MP", index=False)
    assignments_df.to_excel(writer, sheet_name="Assignments", index=False)
output_buffer.seek(0)

st.download_button(
    label="📥 Download Assigned QA Excel",
    data=output_buffer,
    file_name=f"QA_Assigned_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Dashboard & Summary ---
st.subheader("📊 Summary of Assignments")
summary_data = []
for member in team_members:
    assigned_count = sum(product_df['Assigned'] == member)
    limit = member_limit(member)
    summary_data.append({
        "Team Member": member,
        "Assigned": assigned_count,
        "Limit": limit,
        "Remaining Capacity": limit - assigned_count
    })

summary_df = pd.DataFrame(summary_data)
st.dataframe(summary_df)

backlog_count = (product_df['Assigned'] == "Backlog").sum()
st.write(f"**Backlog:** {backlog_count} products")

st.bar_chart(summary_df.set_index("Team Member")["Assigned"])
