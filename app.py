import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

@@ -13,7 +12,7 @@
st.info("Please upload an Excel file to proceed.")
st.stop()

# --- Load sheets with pandas for processing ---
# --- Load sheets directly with pandas ---
qa_df = pd.read_excel(uploaded_file, sheet_name="QA")
mp_df = pd.read_excel(uploaded_file, sheet_name="MP")

@@ -117,27 +116,14 @@ def member_limit(member):
counts[chosen] += assign_count
rows_idx = rows_idx[assign_count:]

# --- Safe write to existing Excel (preserves formulas and other sheets) ---
uploaded_file.seek(0)
wb = load_workbook(uploaded_file)
qa_ws = wb['QA']

# Ensure 'Assigned' column exists
header = [cell.value for cell in qa_ws[1]]
if 'Assigned' not in header:
    qa_ws.insert_cols(1)
    qa_ws.cell(row=1, column=1, value='Assigned')
    assigned_col_idx = 1
else:
    assigned_col_idx = header.index('Assigned') + 1

# Write assignments safely
for i, idx in enumerate(product_df.index, start=2):
    qa_ws.cell(row=i, column=assigned_col_idx, value=product_df.at[idx, 'Assigned'])
# --- Write assignments back ---
qa_df['Assigned'] = qa_df.index.map(lambda idx: product_df.at[idx, 'Assigned'] if idx in product_df.index else "")

# Save to buffer
# --- Download ---
output_buffer = BytesIO()
wb.save(output_buffer)
with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
    qa_df.to_excel(writer, sheet_name="QA", index=False)
    mp_df.to_excel(writer, sheet_name="MP", index=False)
output_buffer.seek(0)

st.download_button(
