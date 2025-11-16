import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="QA Assignment - Stage 4", layout="wide")
st.title("📊 QA Assignment — Stage 4 (Optimized & Safe Columns)")

# --- File upload ---
uploaded_file = st.file_uploader("Upload QA Excel file", type=["xlsx"])
if uploaded_file is None:
    st.info("Please upload an Excel file to proceed.")
    st.stop()

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

# --- Load Excel into DataFrames ---
qa_df = pd.read_excel(uploaded_file, sheet_name="QA")
mp_df = pd.read_excel(uploaded_file, sheet_name="MP")

# --- Normalize QA column names ---
qa_df.columns = [str(c).strip().upper() for c in qa_df.columns]
st.write("📝 QA Sheet Columns:", qa_df.columns.tolist())

# --- Check for essential columns ---
required_cols = ['DIVISION','BRAND','WORKFLOW','AG','AH','AQ']
for col in required_cols:
    if col not in qa_df.columns:
        st.error(f"❌ Column '{col}' not found in QA sheet. Please check your Excel file.")
        st.stop()

# --- Build preferences ---
preferences = {}
for _, row in mp_df.iterrows():
    name, divs = row[0], row[1]
    if pd.notna(name) and pd.notna(divs):
        div_list = [d.strip().title() for d in str(divs).split(",") if d.strip()]
        preferences[name] = div_list

# --- Filter absentees ---
active_preferences = {name: divs for name, divs in preferences.items() if name not in absent_list}
team_members = list(active_preferences.keys())
num_members = len(team_members)
if num_members == 0:
    st.error("❌ No active team members available for assignment!")
    st.stop()

# --- Initialize assignment tracking ---
qa_df['AssignedTo'] = None
counts = {name: 0 for name in team_members}
DEFAULT_TARGET = 100
def member_limit(member):
    return custom_limits.get(member, DEFAULT_TARGET)

# --- Helper functions ---
def eligible_members():
    return [m for m in team_members if counts[m] < member_limit(m)]

def assign_row(row_index, member):
    qa_df.at[row_index, 'AssignedTo'] = member
    counts[member] += 1

# --- Stage 4: Priority override rows (AG/AH numeric) ---
priority_mask = qa_df['AG'].apply(lambda x: pd.notna(x) and isinstance(x, (int, float))) | \
                qa_df['AH'].apply(lambda x: pd.notna(x) and isinstance(x, (int, float)))
priority_rows = qa_df[priority_mask].index.tolist()

# --- Stage 4: Backlog sorting ---
qa_df['AQ'] = pd.to_datetime(qa_df['AQ'], errors='coerce')
if backlog_mode:
    st.info("🕐 Backlog mode ON — sorting by AQ date.")
    qa_df.sort_values(by='AQ', inplace=True)
    priority_rows = sorted(priority_rows, key=lambda i: qa_df.at[i, 'AQ'] if pd.notna(qa_df.at[i, 'AQ']) else pd.Timestamp.max)
else:
    st.success("🚀 Backlog mode OFF — normal order.")

# --- Assign priority override rows first ---
for idx in priority_rows:
    eligible = eligible_members()
    if not eligible:
        qa_df.at[idx, 'AssignedTo'] = 'Backlog'
    else:
        chosen = min(eligible, key=lambda x: counts[x])
        assign_row(idx, chosen)

# --- Assign preferred divisions ---
for member, prefs in active_preferences.items():
    for div in prefs:
        mask = qa_df['AssignedTo'].isna() & (qa_df['DIVISION'] == div)
        unassigned_idx = qa_df[mask].index.tolist()
        for idx in unassigned_idx:
            if counts[member] >= member_limit(member):
                break
            assign_row(idx, member)

# --- Assign remaining by brand (brand blocks together) ---
for brand, group in qa_df[qa_df['AssignedTo'].isna()].groupby('BRAND'):
    unassigned_idx = group.index.tolist()
    while unassigned_idx:
        eligible = eligible_members()
        if not eligible:
            for idx in unassigned_idx:
                qa_df.at[idx, 'AssignedTo'] = 'Backlog'
            break
        chosen = min(eligible, key=lambda x: counts[x])
        capacity = member_limit(chosen) - counts[chosen]
        for idx in unassigned_idx[:capacity]:
            assign_row(idx, chosen)
        unassigned_idx = unassigned_idx[capacity:]

# --- Assignment preview ---
st.subheader("👀 Preview of Assignments")
st.dataframe(qa_df[['DIVISION','BRAND','WORKFLOW','AssignedTo']].head(50))
st.info("Scroll horizontally and vertically to preview more rows.")

# --- Prepare Excel file for download ---
output_buffer = BytesIO()
with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
    qa_df.to_excel(writer, sheet_name='QA', index=False)
    mp_df.to_excel(writer, sheet_name='MP', index=False)
output_buffer.seek(0)

st.download_button(
    label="📥 Download Assigned QA Excel",
    data=output_buffer,
    file_name=f"QA_Assigned_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Summary ---
st.subheader("📊 Summary of assignments:")
for member in team_members:
    st.write(f"- {member}: {counts[member]} products (Limit: {member_limit(member)})")
backlog_count = (qa_df['AssignedTo'] == 'Backlog').sum()
st.write(f"- Backlog: {backlog_count} products")
