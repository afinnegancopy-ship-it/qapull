import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="QA Assignment", layout="wide")
st.title("📊 QA Assignment")

# --- File upload ---
uploaded_file = st.file_uploader("Upload QA Excel file", type=["xlsx"])
if uploaded_file is None:
    st.info("Please upload an Excel file to proceed.")
    st.stop()

# --- Load sheets with pandas ---
qa_df = pd.read_excel(uploaded_file, sheet_name="QA", engine='openpyxl')
mp_df = pd.read_excel(uploaded_file, sheet_name="MP", engine='openpyxl')

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

# --- Read preferences from MP sheet ---
preferences = {}
for _, row in mp_df.iterrows():
    name = row[0]
    divs = row[1]
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
qa_df['Assigned'] = None  # Header changed to Assigned
qa_df['PriorityOverride'] = qa_df.iloc[:, 32].combine_first(qa_df.iloc[:, 33])  # AG/AH columns
qa_df['AQDate'] = qa_df.iloc[:, 42]  # AQ column
qa_df['Division'] = qa_df.iloc[:, 17].astype(str).str.strip().str.title()
qa_df['Brand'] = qa_df.iloc[:, 14].astype(str).str.strip().str.title()
qa_df['Workflow'] = qa_df.iloc[:, 8].astype(str).str.strip()

# --- Backlog sorting ---
if backlog_mode:
    st.info("🕐 Backlog mode ON — sorting all rows by earliest AQ date.")
    qa_df.sort_values('AQDate', inplace=True)
else:
    st.success("🚀 Backlog mode OFF — assigning in normal order.")

# --- Assignment helpers ---
DEFAULT_TARGET = 100
counts = {member: 0 for member in team_members}
assignments = {member: [] for member in team_members}

def member_limit(member):
    return custom_limits.get(member, DEFAULT_TARGET)

def assign_block(member, rows_idx):
    """Assign a block of rows to a member respecting remaining capacity."""
    remaining_capacity = member_limit(member) - counts[member]
    assign_count = min(len(rows_idx), remaining_capacity)
    if assign_count <= 0:
        return
    qa_df.loc[rows_idx[:assign_count], 'Assigned'] = member
    assignments[member].extend(rows_idx[:assign_count])
    counts[member] += assign_count

# --- Step 1: Priority override rows ---
priority_override_rows = qa_df[qa_df['PriorityOverride'].notna()]
for idx, row in priority_override_rows.iterrows():
    eligible = [m for m in team_members if counts[m] < member_limit(m)]
    if eligible:
        chosen = min(eligible, key=lambda x: counts[x])
        qa_df.at[idx, 'Assigned'] = chosen
        assignments[chosen].append(idx)
        counts[chosen] += 1
    else:
        qa_df.at[idx, 'Assigned'] = "Backlog"

# --- Step 2: Preferred divisions with brand & workflow priority ---
for member in team_members:
    prefs = active_preferences[member]
    for pref_div in prefs:
        unassigned_df = qa_df[(qa_df['Assigned'].isna()) & (qa_df['Division'] == pref_div)]
        for brand, brand_group in unassigned_df.groupby('Brand'):
            # Workflow prioritization
            priority_rows = brand_group[brand_group['Workflow'] == "Prioritise in Workflow"].index.tolist()
            normal_rows = brand_group[brand_group['Workflow'] != "Prioritise in Workflow"].index.tolist()
            assign_block(member, priority_rows)
            assign_block(member, normal_rows)

# --- Step 3: Remaining unassigned rows ---
unassigned_idx = qa_df[qa_df['Assigned'].isna()].index
for idx in unassigned_idx:
    eligible = [m for m in team_members if counts[m] < member_limit(m)]
    if eligible:
        chosen = min(eligible, key=lambda x: counts[x])
        qa_df.at[idx, 'Assigned'] = chosen
        assignments[chosen].append(idx)
        counts[chosen] += 1
    else:
        qa_df.at[idx, 'Assigned'] = "Backlog"

# --- Prepare file for download ---
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

# --- Summary ---
st.subheader("📊 Summary of assignments:")
for name, rows in assignments.items():
    st.write(f"- {name}: {len(rows)} products (Limit: {member_limit(name)})")

backlog_count = (qa_df['Assigned'] == "Backlog").sum()
st.write(f"- Backlog: {backlog_count} products")

