import csv
import openpyxl

# Step 1: Load JP sellers from P0 CSV
jp_lookup = {}
with open('P0_20260404.csv', 'r', encoding='latin1') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row.get('embr_region') == 'JP':
            mid = row.get('merchant_customer_id', '').strip()
            if mid and mid not in jp_lookup:
                jp_lookup[mid] = {
                    'launch_channel': row.get('launch_channel', ''),
                    'esm_program': row.get('esm_program', '')
                }

print(f"JP sellers in P0: {len(jp_lookup)}")

# Step 2: Load Excel (read-only for speed)
wb = openpyxl.load_workbook('02_Validation Result_ooc.xlsx', read_only=True)
ws = wb.active

# Read all rows
all_rows = list(ws.rows)
header_row = all_rows[0]
header = [cell.value for cell in header_row]
print(f"Excel columns: {len(header)}, data rows: {len(all_rows)-1}")

# Find matching rows
matched = []
for row in all_rows[1:]:
    mid_val = row[3].value
    if mid_val is None:
        continue
    mid_str = str(int(mid_val)) if isinstance(mid_val, (int, float)) else str(mid_val).strip()
    if mid_str in jp_lookup:
        vals = [cell.value for cell in row]
        info = jp_lookup[mid_str]
        vals.append(info['launch_channel'])
        vals.append(info['esm_program'])
        matched.append(vals)

wb.close()
print(f"Matched rows: {len(matched)}")

# Step 3: Write output
wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "Matched JP Sellers"

# Write header
out_header = header + ['launch_channel', 'esm_program']
ws_out.append(out_header)

# Write data
for vals in matched:
    ws_out.append(vals)

wb_out.save('Matched_JP_Sellers.xlsx')
print("Output saved to: Matched_JP_Sellers.xlsx")
