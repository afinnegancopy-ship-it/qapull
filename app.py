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

# Filter absentees
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

# --- 1️⃣ Assign by preferences ---
for member in team_members:
    prefs = active_preferences[member]
    for div in prefs:
        mask = product_df['Assigned'].isna() & (product_df['Division'] == div)
        rows_to_assign = product_df[mask].index
        remaining_capacity = member_limit(member) - counts[member]
        assign_count = min(len(rows_to_assign), remaining_capacity)
        if assign_count > 0:
            product_df.loc[rows_to_assign[:assign_count], 'Assigned'] = member
            counts[member] += assign_count

# --- 2️⃣ Assign priority override products ---
priority_mask = product_df['Assigned'].isna() & product_df['PriorityOverride'].notna()
for idx in product_df[priority_mask].index:
    eligible = [m for m in team_members if counts[m] < member_limit(m)]
    if eligible:
        chosen = min(eligible, key=lambda x: counts[x])
        product_df.at[idx, 'Assigned'] = chosen
        counts[chosen] += 1
    else:
        product_df.at[idx, 'Assigned'] = "Backlog"

# --- 3️⃣ Assign remaining products by brand globally ---
remaining_mask = product_df['Assigned'].isna()
brands = product_df.loc[remaining_mask, 'Brand'].unique()
for brand in brands:
    brand_mask = remaining_mask & (product_df['Brand'] == brand)
    rows_idx = product_df[brand_mask].index.tolist()
    while rows_idx:
        eligible = [m for m in team_members if counts[m] < member_limit(m)]
        if not eligible:
            product_df.loc[rows_idx, 'Assigned'] = "Backlog"
            break
        chosen = max(eligible, key=lambda m: member_limit(m) - counts[m])
        remaining_capacity = member_limit(chosen) - counts[chosen]
        assign_count = min(len(rows_idx), remaining_capacity)
        product_df.loc[rows_idx[:assign_count], 'Assigned'] = chosen
        counts[chosen] += assign_count
        rows_idx = rows_idx[assign_count:]

# --- Write assignments back ---
qa_df['Assigned'] = qa_df.index.map(lambda idx: product_df.at[idx, 'Assigned'] if idx in product_df.index else "")

# --- Download ---
output_buffer = BytesIO()
with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
    qa_df.to_excel(writer, sheet_name="QA", index=False)
    mp_df.to_excel(writer, sheet_name="MP", index=False)
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
