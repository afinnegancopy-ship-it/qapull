"""
===========================
QA Assignment Script — Stage 4 (Backlog Mode + Priority Override + Custom Limits + Preference Filter + Brand Blocks)
===========================

Stage 4 Features:no
- NEW: "Backlog Mode" — if enabled, sorts all products by AQ date (earliest first)
- Prioritizes rows where AG or AH contain numbers
- Keeps all Stage 3 logic:
  ✅ Absent members handled
  ✅ Division preferences respected
  ✅ Brand blocks stay together
  ✅ Per-member custom limits
  ✅ Balanced assignment
  ✅ Backlog fallback
"""

import openpyxl
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict

# --- File path ---
file_path = r"\\BTBNAS\DigitalAssets\Photostudio\0 - Enrichment\QA\2025\PIM QA 2025\06.11.2025.xlsx"

# --- Ask for backlog mode ---
backlog_mode_input = input("Are you expecting to be in backlog today? (yes/no): ").strip().lower()
backlog_mode = backlog_mode_input in ["yes", "y"]

# --- Ask for absentees ---
absent_input = input("Is anyone absent today? (Type names separated by commas, or press Enter if no one is absent): ").strip().lower()
absent_list = []
if absent_input and absent_input not in ["no", "none", "n"]:
    absent_list = [name.strip().title() for name in absent_input.split(",") if name.strip()]
    print(f"🟡 Absent today: {', '.join(absent_list)}")
else:
    print("✅ Everyone is present.")

# --- Ask for custom product limits ---
custom_input = input("Does any member have a specific product count limit? (format: Name:Limit, separated by commas, or press Enter for none): ").strip()
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
                print(f"⚠️ Invalid limit for {name}, ignoring.")

# --- Load workbook safely ---
wb = load_workbook(file_path)
qa_ws = wb["QA"]
mp_ws = wb["MP"]

# --- Read mapping (MP sheet) with multiple preferences ---
preferences = {}
for row in mp_ws.iter_rows(min_row=2, values_only=True):
    name, divs = row[0], row[1]
    if name and divs:
        div_list = [d.strip().title() for d in divs.split(",") if d.strip()]
        preferences[name] = div_list

# --- Filter out absentees ---
active_preferences = {name: divs for name, divs in preferences.items() if name not in absent_list}
team_members = list(active_preferences.keys())
num_members = len(team_members)
if num_members == 0:
    raise ValueError("❌ No active team members available for assignment!")

# --- Build QA data ---
qa_rows = []
brand_rows = defaultdict(list)
priority_rows = []
priority_override_rows = []  # AG/AH numeric priority
normal_rows = []

for i, row in enumerate(qa_ws.iter_rows(min_row=2, values_only=True), start=2):
    assigned_to = row[0]
    division = str(row[17]).strip() if row[17] else ""
    m_value = row[12]
    brand = row[14]
    workflow = str(row[8]).strip() if row[8] else ""
    col_ag = row[32]  # Column AG
    col_ah = row[33]  # Column AH
    col_aq = row[42]  # Column AQ (for backlog mode)

    # Detect Stage 3 priority (AG or AH numeric)
    if isinstance(col_ag, (int, float)) or isinstance(col_ah, (int, float)):
        priority_override_rows.append((i, division, brand, workflow, col_aq))

    if m_value is not None and str(m_value).strip() != "":
        qa_rows.append((i, assigned_to, division, brand, workflow, col_aq))
        brand_rows[brand].append((i, division, workflow, col_aq))
        if workflow == "Prioritise in Workflow":
            priority_rows.append((i, division, brand, workflow, col_aq))
        else:
            normal_rows.append((i, division, brand, workflow, col_aq))

# --- Apply backlog sorting if needed ---
if backlog_mode:
    print("🕐 Backlog mode ON — sorting all rows by earliest AQ date.")
    def sort_key(x):
        date_val = x[-1]
        return date_val if isinstance(date_val, datetime) else datetime.max
    qa_rows.sort(key=sort_key)
    priority_override_rows.sort(key=sort_key)
    priority_rows.sort(key=sort_key)
    normal_rows.sort(key=sort_key)
    for brand in brand_rows:
        brand_rows[brand].sort(key=sort_key)
else:
    print("🚀 Backlog mode OFF — assigning in normal order.")

# --- Initialize assignment trackers ---
assignments = {name: [] for name in team_members}
counts = {name: 0 for name in team_members}
DEFAULT_TARGET = 100

# --- Helper: get member's current limit ---
def member_limit(member):
    return custom_limits.get(member, DEFAULT_TARGET)

# --- Helper: assign individual rows (used for priority overrides) ---
def assign_rows(rows):
    for r, div, brand, workflow, *_ in rows:
        eligible = [m for m in team_members if counts[m] < member_limit(m)]
        if not eligible:
            qa_ws[f"A{r}"].value = "Backlog"
            continue
        chosen = min(eligible, key=lambda x: counts[x])
        qa_ws[f"A{r}"].value = chosen
        assignments[chosen].append(r)
        counts[chosen] += 1

# --- Helper: assign brand blocks ---
def assign_brand_block(member, rows):
    remaining_capacity = member_limit(member) - counts[member]
    if remaining_capacity <= 0:
        return 0
    for r, div, workflow, *_ in rows[:remaining_capacity]:
        qa_ws[f"A{r}"].value = member
        assignments[member].append(r)
        counts[member] += 1
    return len(rows[:remaining_capacity])

# --- STAGE 4 STEP 1️⃣: Assign AG/AH priority rows first ---
if priority_override_rows:
    print(f"\n🚨 Found {len(priority_override_rows)} AG/AH priority rows. Assigning first...")
    assign_rows(priority_override_rows)
else:
    print("\n✅ No AG/AH priority rows found.")

# --- STAGE 4 STEP 2️⃣: Assign preferred divisions next ---
for member in team_members:
    prefs = active_preferences[member]
    for pref_div in prefs:
        for brand, rows in brand_rows.items():
            unassigned = [r for r in rows if qa_ws[f"A{r[0]}"].value in [None, ""] and r[1] == pref_div]
            if unassigned:
                assign_brand_block(member, unassigned)

# --- STAGE 4 STEP 3️⃣: Assign remaining brands normally ---
for brand, rows in brand_rows.items():
    unassigned = [r for r in rows if qa_ws[f"A{r[0]}"].value in [None, ""]]
    if not unassigned:
        continue
    eligible = [m for m in team_members if counts[m] < member_limit(m)]
    if eligible:
        chosen = min(eligible, key=lambda x: counts[x])
        assign_brand_block(chosen, unassigned)
    else:
        for r, div, workflow, *_ in unassigned:
            qa_ws[f"A{r}"].value = "Backlog"

# --- Convert formulas to values ---
for row in qa_ws.iter_rows():
    for cell in row:
        if cell.data_type == "f":
            cell.value = cell.value

# --- Save timestamped file ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
file_base, file_ext = file_path.rsplit(".", 1)
output_path = f"{file_base}_{timestamp}.xlsx"
wb.save(output_path)

# --- Summary ---
print("\n✅ Assignment complete!")
print(f"📄 Saved as: {output_path}")
print("\n📊 Summary of assignments:")
for name, rows in assignments.items():
    limit = member_limit(name)
    print(f"  - {name}: {len(rows)} products (Limit: {limit})")
backlog_count = sum(1 for r in qa_rows if qa_ws[f"A{r[0]}"].value == "Backlog")
print(f"  - Backlog: {backlog_count} products")
