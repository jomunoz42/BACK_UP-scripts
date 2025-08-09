from openpyxl import load_workbook
import re
from collections import defaultdict

wb = load_workbook('buffer.xlsx')
ws = wb.active

# Step 1 - Move column M to A and keep only first word
for row in range(1, ws.max_row + 1):
    value = ws[f"M{row}"].value
    if value is not None:
        # Ensure it's a string before splitting
        first_word = str(value).split()[0] if str(value).strip() != "" else ""
        ws[f"A{row}"].value = first_word
    else:
        ws[f"A{row}"].value = None
    ws[f"M{row}"].value = None  # Clear original col M

# Step 2 - Move column R to B
for row in range(1, ws.max_row + 1):
    value = ws[f"R{row}"].value
    ws[f"B{row}"].value = value
    ws[f"R{row}"].value = None  # Clear original col R

# Step 1 — Keep only columns A and B, delete all others
max_col = ws.max_column
for col in range(max_col, 2, -1):  # start from the end and go down to C
    ws.delete_cols(col)

# Step 2 — Delete rows in which column B contains the given keywords
keywords = ["Villa", "Galé", "Arroios III", "Antonio Pedro 2"]

# Loop from bottom to top so deleting rows doesn't mess up indexes
for row in range(ws.max_row, 0, -1):
    cell_value = str(ws[f"B{row}"].value or "")
    if any(keyword in cell_value for keyword in keywords):
        ws.delete_rows(row)

# Step 1 — Read values from col B
values = [str(ws[f'B{row}'].value or "") for row in range(1, ws.max_row + 1)]

# Helper: extract data structures
def parse_row(text):
    text = text.strip()
    if text == "Porteira":
        return ("porteira", None)

    if "|" not in text:
        # Case: "Double open shower (604)" -> only 604
        m = re.search(r"\((\d+)\)", text)
        if m:
            return ("single_number", [m.group(1)])
        return ("unknown", [])

    description, after_pipe = map(str.strip, text.split("|", 1))

    # Remove trailing/leading commas and spaces
    after_pipe = after_pipe.strip().strip(",")

    # Handle comma separated multi‑pipes: "208, Quadruplos ensuite | 409"
    # We will process them later in grouping
    return (description, after_pipe)

# Step 2 — Group lines by description for multi‑row merging logic
groups = defaultdict(list) # description -> [after_pipe values]
for v in values:
    typ, data = parse_row(v)
    if typ == "porteira":
        groups[v].append(None)  # keep as is
    elif typ == "single_number":
        groups[v].append(data[0])
    elif typ != "unknown":
        groups[typ].append(data)

# Step 3 — Function to process each group according to your rules
def process_group(description, after_pipe_list):
    if description == "Porteira":
        return ["Porteira"]

    results = []

    # Flatten and clean multiple "208, 409" type strings
    numbers = []
    for x in after_pipe_list:
        parts = re.split(r",\s*", str(x))
        for p in parts:
            p = p.strip()
            if not p:
                continue
            # Handle 503-7 type -> "503 (7)"
            dash_match = re.match(r"(\d+)-(\d+)", p)
            paren_match = re.match(r"(\d+)\s*\(([\d, ]+)\)", p)
            if dash_match:
                numbers.append((dash_match.group(1), dash_match.group(2)))
            elif paren_match:
                num = paren_match.group(1)
                for bed in paren_match.group(2).split(","):
                    numbers.append((num, bed.strip()))
            else:
                # Just plain number
                numbers.append((p, None))

    # Special cases
    if description.startswith("506 Dorm 12 Female"):
        # Separate into A and B groups
        groupA = []
        groupB = []
        allbeds = defaultdict(list)  # For each base number -> beds
        for num, bed in numbers:
            if bed is None:
                results.append(num)
            else:
                bed_num = int(bed)
                if 1 <= bed_num <= 6:
                    groupA.append(bed_num)
                elif 7 <= bed_num <= 12:
                    groupB.append(bed_num)
        groupA.sort()
        groupB.sort()
        if groupA:
            results.append(f"506-A ({','.join(map(str, groupA))})")
        if groupB:
            results.append(f"506-B ({','.join(map(str, groupB))})")

    elif description.startswith("503 + 505 Dorm 2x8 females"):
        merged = defaultdict(list)
        for num, bed in numbers:
            if bed is None:
                results.append(num)
            else:
                merged[num].append(int(bed))
        for num in merged:
            merged[num] = sorted(set(merged[num]))
            results.append(f"{num} ({','.join(map(str, merged[num]))})")

    elif description.startswith("501 Dorm 8 male"):
        merged = defaultdict(list)
        for num, bed in numbers:
            if bed is None:
                results.append(num)
            else:
                merged[num].append(int(bed))
        for num in merged:
            merged[num] = sorted(set(merged[num]))
            results.append(f"{num} ({','.join(map(str, merged[num]))})")

    elif "Quadruplos ensuite" in description:
        only_nums = []
        for num, bed in numbers:
            if num:
                only_nums.append(num)
        results.append(", ".join(only_nums))

    else:
        # Generic: keep only number after pipe
        for num, bed in numbers:
            if num:
                if bed:
                    results.append(f"{num} ({bed})")
                else:
                    results.append(num)

    return results

# Step 4 — Apply processing
new_values = []
for v in values:
    typ, data = parse_row(v)
    if typ == "porteira":
        new_values.append("Porteira")
    elif typ == "single_number":
        new_values.append(data[0])
    else:
        # Process group for this description
        desc = typ
        outputs = process_group(typ, groups[typ])
        # We just take the merged final result for this whole group
        if outputs:
            merged_final = ", ".join(outputs)
            new_values.append(merged_final)
        else:
            new_values.append("")

# Step 5 — Write back to column B
for row, val in enumerate(new_values, start=1):
    ws[f"B{row}"].value = val


wb.save("buffer.xlsx")

print("Done.")