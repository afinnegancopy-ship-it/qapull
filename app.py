import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict, deque
import math

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="QA Assignment Tool", layout="wide")
st.title("QA Assignment Tool 📊")
st.write("Assigns products to QA team. (Updated: split large brands only if >30% imbalance)")

# ---------------------------
# Helpers
# ---------------------------

def title_or_none(val):
    return val.strip().title() if isinstance(val, str) and val.strip() else None


def remaining_capacity(member, limits, counts):
    return max(0, limits.get(member, 0) - counts.get(member, 0))


def compute_loads_after_assignment(counts, member_to_add, add):
    new_counts = counts.copy()
    new_counts[member_to_add] = new_counts.get(member_to_add, 0) + add
    return sorted(new_counts.values(), reverse=True)


def top_and_second(loads):
    if not loads: return 0,0
    if len(loads)==1: return loads[0],0
    return loads[0], loads[1]

# ---------------------------
# File upload
# ---------------------------
uploaded_file = st.file_uploader("Upload QA Template", type=["xlsx"]) 
if not uploaded_file:
    st.info("Please upload an Excel (.xlsx) file containing 'QA' and 'Assignments' sheets.")
    st.stop()

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

# Options
backlog_mode = st.checkbox("Backlog mode (sort by earliest AQ date)", value=False)
st.write("Enter active members today (e.g: Ross:100, Phoebe:80, Monica)")
working_input = st.text_input("Active members")

if not working_input:
    st.error("Please enter at least one active member.")
    st.stop()

# Parse active members and custom limits
active_members = []
member_limits = {}
for part in working_input.split(","):
    part = part.strip()
    if not part: continue
    if ":" in part:
        name, limit = part.split(":",1)
        name = name.strip().title()
        try: lim = int(limit.strip())
        except: lim=100
        active_members.append(name)
        member_limits[name] = lim
    else:
        active_members.append(part.strip().title())

for m in active_members:
    if m not in member_limits: member_limits[m]=100

# Preassignments
brand_to_member = {}
for row in assignments_ws.iter_rows(min_row=2, values_only=True):
    brand, member = row[0], row[1]
    if brand and member:
        brand_to_member[title_or_none(brand)] = title_or_none(member)

# Build brand blocks
brand_blocks = defaultdict(list)
row_brand_order=[]
qa_rows=[]
for i,row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True),start=2):
    m_value=row[12]; brand=row[14]; col_aq=row[42]
    if m_value is not None and str(m_value).strip():
        qa_rows.append((i,brand,col_aq))
        if brand:
            btitle=title_or_none(brand)
            if btitle not in brand_blocks: row_brand_order.append(btitle)
            brand_blocks[btitle].append(i)

if backlog_mode:
    def row_aq_date(row_idx):
        try: val=qa_ws[f"AQ{row_idx}"].value
        except: val=None
        return val if isinstance(val,datetime) else datetime.max
    for b in brand_blocks: brand_blocks[b].sort(key=row_aq_date)

blocks=[]
for b in row_brand_order: blocks.append((b,brand_blocks[b].copy()))
blocks.sort(key=lambda x:len(x[1]),reverse=True)

counts={m:0 for m in active_members}
assignments={m:[] for m in active_members}
assigned_blocks=[]
backlog_rows=[]
IMBALANCE_RATIO=1.30
SPLIT_SIZE_THRESHOLD=50  # only split brands >=50

blocks_queue=deque(blocks)
iteration=0; max_iterations=20000

while blocks_queue and iteration<max_iterations:
    iteration+=1
    current_brand,rows=blocks_queue.popleft()
    block_size=len(rows)
    if block_size==0: continue

    # preassignment
    pre_member=brand_to_member.get(current_brand)
    if pre_member and pre_member in active_members:
        cap=remaining_capacity(pre_member, member_limits, counts)
        if cap>=block_size:
            for r in rows: qa_ws[f"A{r}"].value=pre_member
            assignments[pre_member].extend(rows)
            counts[pre_member]+=len(rows)
            assigned_blocks.append({'brand':current_brand,'rows':rows.copy(),'member':pre_member})
            continue
        else:
            take=min(cap,block_size)
            if take>0:
                for r in rows[:take]: qa_ws[f"A{r}"].value=pre_member
                assignments[pre_member].extend(rows[:take])
                counts[pre_member]+=take
                assigned_blocks.append({'brand':current_brand,'rows':rows[:take].copy(),'member':pre_member})
            remaining=rows[take:]
            if remaining: blocks_queue.appendleft((current_brand,remaining))
            continue

    # Decide splitting
    if block_size<SPLIT_SIZE_THRESHOLD:
        # small brand -> assign whole to one candidate
        candidates_can_take=[m for m in active_members if remaining_capacity(m,member_limits,counts)>=block_size]
        if candidates_can_take:
            candidates_can_take.sort(key=lambda m:(-remaining_capacity(m,member_limits,counts),counts[m]))
            best=candidates_can_take[0]
            for r in rows: qa_ws[f"A{r}"].value=best
            assignments[best].extend(rows)
            counts[best]+=block_size
            assigned_blocks.append({'brand':current_brand,'rows':rows.copy(),'member':best})
        else:
            for r in rows: qa_ws[f"A{r}"].value="Backlog"; backlog_rows.append(r)
        continue

    # large brand -> check imbalance
    candidates_can_take=[m for m in active_members if remaining_capacity(m,member_limits,counts)>=block_size]
    best_candidate=candidates_can_take[0] if candidates_can_take else None
    would_imbalance=False
    if best_candidate:
        loads_after=compute_loads_after_assignment(counts,best_candidate,block_size)
        top,second=top_and_second(loads_after)
        if second==0: would_imbalance=(top>0 and second==0)
        else: would_imbalance=(top>IMBALANCE_RATIO*second)

    if not best_candidate or would_imbalance:
        # split interleaved proportional
        eligible=[m for m in active_members if remaining_capacity(m,member_limits,counts)>0]
        rem_caps={m:remaining_capacity(m,member_limits,counts) for m in eligible}
        total=sum(rem_caps.values())
        if total==0: 
            for r in rows: qa_ws[f"A{r}"].value="Backlog"; backlog_rows.append(r)
            continue

        tentative={m:math.floor(rem_caps[m]/total*block_size) for m in eligible}
        assigned_sum=sum(tentative.values())
        remaining_to_assign=block_size-assigned_sum
        members_by_cap=sorted(eligible,key=lambda m:-rem_caps[m])
        idx=0
        while remaining_to_assign>0 and members_by_cap:
            m=members_by_cap[idx%len(members_by_cap)]
            if tentative[m]<rem_caps[m]: tentative[m]+=1; remaining_to_assign-=1
            idx+=1
            if idx>block_size*5: break

        quotas={m:tentative[m] for m in eligible}
        row_iter=iter(rows)
        from collections import deque as _deque
        member_queue=_deque([m for m in eligible if quotas[m]>0])
        while member_queue and rows:
            m=member_queue.popleft()
            if quotas[m]<=0: continue
            r=rows.pop(0)
            qa_ws[f"A{r}"].value=m
            assignments[m].append(r)
            counts[m]+=1
            quotas[m]-=1
            if quotas[m]>0: member_queue.append(m)
        for r in rows: qa_ws[f"A{r}"].value="Backlog"; backlog_rows.append(r)
    else:
        # assign whole to best_candidate
        for r in rows: qa_ws[f"A{r}"].value=best_candidate
        assignments[best_candidate].extend(rows)
        counts[best_candidate]+=block_size
        assigned_blocks.append({'brand':current_brand,'rows':rows.copy(),'member':best_candidate})

# Convert formulas to values
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type=="f":
            try: cell.value=cell.value
            except: pass

# Save output
timestamp=datetime.now().strftime("%Y%m%d_%H%M")
output_path=f"QA_Assignment_{timestamp}.xlsx"
wb.save(output_path)

st.success("✅ Assignment complete!")
st.write(f"📄 Saved as: {output_path}")

# Summary
st.write("📊 Summary of assignments:")
for member in active_members:
    limit=member_limits[member]
    st.write(f"- {member}: {len(assignments.get(member,[]))} products (Target: {limit})")
st.write(f"- Backlog (explicitly set): {len(backlog_rows)} products")

# Download
with open(output_path,"rb") as f:
    st.download_button(label="📥 Download Assigned Excel", data=f, file_name=output_path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
