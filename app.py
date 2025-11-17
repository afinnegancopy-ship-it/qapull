# --- Read Assignments tab ---
assignments_df = pd.read_excel(uploaded_file, sheet_name="Assignments")
# Make mapping: Brand -> Member
brand_to_member = {}
for _, row in assignments_df.iterrows():
    brand = str(row[0]).strip().title()
    member = str(row[1]).strip().title()
    if brand and member:
        brand_to_member[brand] = member

# --- Assign products based on Assignments tab ---
for brand, assigned_member in brand_to_member.items():
    if assigned_member not in team_members:
        # Assigned member is absent or not in active team, skip
        continue
    
    brand_mask = product_df['Brand'] == brand
    brand_rows = product_df[brand_mask & product_df['Assigned'].isna()].index.tolist()
    
    remaining_capacity = ideal_targets[assigned_member] - counts[assigned_member]
    brand_size = len(brand_rows)
    
    if remaining_capacity >= brand_size:
        # Assign entire brand block
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
