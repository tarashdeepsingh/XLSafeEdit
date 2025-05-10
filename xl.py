import zipfile
import shutil
import os
from lxml import etree

def update_excel_from_json(xlsx_path, output_path, data_json, target_sheet_name, has_header=True):
    temp_dir = "temp_excel"

    # Step 1: Unzip Excel (keep it in memory)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    # Step 2: Locate correct sheet by name
    workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
    workbook_tree = etree.parse(workbook_path)
    sheets = workbook_tree.findall(".//main:sheets/main:sheet", namespaces=ns)

    sheet_filename = None

    for i, sheet in enumerate(sheets, start=1):
        if sheet.attrib.get("name") == target_sheet_name:
            sheet_filename = f"sheet{i}.xml"
            break

    if not sheet_filename:
        raise ValueError(f"Sheet named '{target_sheet_name}' not found!")

    sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', sheet_filename)

    # Step 3: Load sharedStrings.xml (optional)
    shared_strings_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
    shared_strings = []
    if os.path.exists(shared_strings_path):
        sst = etree.parse(shared_strings_path)
        for si in sst.findall(".//main:si", namespaces=ns):
            t = si.find(".//main:t", namespaces=ns)
            shared_strings.append(t.text if t is not None else "")

    # Step 4: Parse the worksheet
    tree = etree.parse(sheet_path)
    root = tree.getroot()
    rows = root.findall(".//main:sheetData/main:row", namespaces=ns)

    def set_cell_value(row, col_letter, value):
        row_num = row.attrib["r"]
        cell_ref = f"{col_letter}{row_num}"
        cell = None

        # Find the cell in the row
        for c in row.findall("main:c", namespaces=ns):
            if c.attrib.get("r") == cell_ref:
                cell = c
                break

        if cell is None:
            cell = etree.Element(f"{{{ns['main']}}}c", r=cell_ref)
            row.append(cell)

        # Set the value type (string or numeric)
        if isinstance(value, (int, float)):
            cell.attrib["t"] = "n"
        else:
            cell.attrib["t"] = "str"

        # Set the actual value
        v = cell.find("main:v", namespaces=ns)
        if v is None:
            v = etree.SubElement(cell, f"{{{ns['main']}}}v")
        v.text = str(value)

    # Step 5: Update matching rows with data_json
    updates = 0
    data_iter = iter(data_json)

    for row in rows:
        if has_header and int(row.attrib["r"]) == 1:
            continue  # Skip header row if specified

        try:
            entry = next(data_iter)
        except StopIteration:
            break  # No more data to write

        for col_letter, value in entry.items():
            set_cell_value(row, col_letter, value)

        print(f"✅ Updated Row {row.attrib['r']}")
        updates += 1

    print(f"\n✅ Total rows updated: {updates}")

    # Step 6: Save updated sheet
    tree.write(sheet_path, xml_declaration=True, encoding='UTF-8')

    # Step 7: Repack Excel (only if modifications are done)
    shutil.make_archive("final_excel", 'zip', temp_dir)
    shutil.move("final_excel.zip", output_path)
    shutil.rmtree(temp_dir)

    print(f"\n✅ Excel updated successfully at: {output_path}")


# Example usage
data_json = [
    {"K": "Mumbai", "L": "Maharashtra", "M": "India", "N": 34.5, "O": 76.5, "P": 88.3},
    {"K": "Mumbai", "L": "Maharashtra", "M": "India", "N": 36.8, "O": 69.2, "P": 91.0},
    {"K": "Mumbai", "L": "Maharashtra", "M": "India", "N": 384.47, "O": 97.2, "P": 23.3}
]

update_excel_from_json(
    r"C:\Users\Tarashdeep Singh\Desktop\test2.xlsx",
    r"C:\Users\Tarashdeep Singh\Desktop\updated_test.xlsx",
    data_json,
    target_sheet_name="Data Sheet",
    has_header=True
)
