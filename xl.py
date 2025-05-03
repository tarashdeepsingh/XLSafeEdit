import zipfile
import shutil
import os
from lxml import etree

def update_excel_from_json(xlsx_path, output_path, data_json):
    temp_dir = "temp_excel"

    # Step 1: Unzip Excel
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    # Step 2: Locate correct sheet by name
    workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
    workbook_tree = etree.parse(workbook_path)
    sheets = workbook_tree.findall(".//main:sheets/main:sheet", namespaces=ns)

    sheet_filename = None
    target_sheet_name = "Data Sheet"

    for i, sheet in enumerate(sheets, start=1):
        if sheet.attrib.get("name") == target_sheet_name:
            sheet_filename = f"sheet{i}.xml"
            break

    if not sheet_filename:
        raise ValueError(f"Sheet named '{target_sheet_name}' not found!")

    sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', sheet_filename)

    # Step 3: Load sharedStrings.xml
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

        for c in row.findall("main:c", namespaces=ns):
            if c.attrib.get("r") == cell_ref:
                cell = c
                break

        if cell is None:
            cell = etree.Element(f"{{{ns['main']}}}c", r=cell_ref)
            row.append(cell)

        cell.attrib["t"] = "n" if isinstance(value, (int, float)) else "str"
        v = cell.find("main:v", namespaces=ns)
        if v is None:
            v = etree.SubElement(cell, f"{{{ns['main']}}}v")
        v.text = str(value)

    # Step 5: Update matching rows with data_json
    updates = 0
    data_iter = iter(data_json)

    for row in rows:
        if int(row.attrib["r"]) == 1:
            continue  # Skip header

        try:
            entry = next(data_iter)
        except StopIteration:
            break  # No more data to write

        set_cell_value(row, "J", entry["area_name"])
        set_cell_value(row, "K", entry["parent_area"])
        set_cell_value(row, "L", entry["state_value"])
        set_cell_value(row, "M", entry["national_value"])
        print(f"✅ Updated Row {row.attrib['r']}")
        updates += 1

    print(f"\n✅ Total rows updated: {updates}")

    # Step 6: Save updated sheet
    tree.write(sheet_path, xml_declaration=True, encoding='UTF-8')

    # Step 7: Repack Excel
    shutil.make_archive("final_excel", 'zip', temp_dir)
    shutil.move("final_excel.zip", output_path)
    shutil.rmtree(temp_dir)

    print(f"\n✅ Excel updated successfully at: {output_path}")


# Example usage
data_json = [
    {
        "area_name": "Mumbai",
        "parent_area": "India",
        "state_value": 76.5,
        "national_value": 88.3
    },
    {
        "area_name": "Mumbai",
        "parent_area": "India",
        "state_value": 69.2,
        "national_value": 91.0
    },
    {
        "area_name": "Mumbai",
        "parent_area": "India",
        "state_value": 97.2,
        "national_value": 23.3
    }
]

update_excel_from_json(
    r"C:\Users\Tarashdeep Singh\Desktop\test.xlsx",
    r"C:\Users\Tarashdeep Singh\Desktop\updated_test.xlsx",
    data_json
)
