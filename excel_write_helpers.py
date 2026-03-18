import io
import pathlib
import shutil
import zipfile
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter


def _update_checkboxes(out_path: str, sheet_position_updates: dict[str, dict[tuple, bool]]) -> None:
    """
    Modify <x:Checked> in VML XML for checkboxes identified by (row, col).

    sheet_position_updates: {sheet_name: {(row, col): True/False}}
    """
    from lxml import etree
    from excel_read_helpers import _find_sheet_path, _find_vml_files_for_sheet

    ns_v = "urn:schemas-microsoft-com:vml"
    ns_x = "urn:schemas-microsoft-com:office:excel"

    # Read all zip contents up front
    with zipfile.ZipFile(out_path, "r") as z:
        all_names = z.namelist()
        file_contents = {name: z.read(name) for name in all_names}

    modified: dict[str, bytes] = {}

    for sheet_name, position_updates in sheet_position_updates.items():
        if not position_updates:
            continue

        with zipfile.ZipFile(out_path, "r") as z:
            sheet_path = _find_sheet_path(z, sheet_name)
            if not sheet_path:
                continue
            vml_files = _find_vml_files_for_sheet(z, sheet_path)

        for vml_file in vml_files:
            if vml_file not in file_contents:
                continue

            root = etree.fromstring(file_contents[vml_file], etree.XMLParser(recover=True))
            changed = False

            for shape in root.findall(f".//{{{ns_v}}}shape"):
                cd = shape.find(f".//{{{ns_x}}}ClientData")
                if cd is None or cd.get("ObjectType") != "Checkbox":
                    continue

                anchor_el = cd.find(f"{{{ns_x}}}Anchor")
                if anchor_el is None or not anchor_el.text:
                    continue

                parts = [p.strip() for p in anchor_el.text.split(",")]
                row, col = 0, 0
                if len(parts) >= 8:
                    col = int(parts[0]) + 1
                    row = ((int(parts[2]) + int(parts[6])) // 2) + 1
                elif len(parts) >= 4:
                    col = int(parts[0]) + 1
                    row = int(parts[2]) + 1

                if (row, col) not in position_updates:
                    continue

                should_check = position_updates[(row, col)]
                checked_el = cd.find(f"{{{ns_x}}}Checked")

                if should_check:
                    if checked_el is None:
                        checked_el = etree.SubElement(cd, f"{{{ns_x}}}Checked")
                    checked_el.text = "1"
                else:
                    if checked_el is not None:
                        cd.remove(checked_el)

                changed = True

            if changed:
                modified[vml_file] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

    if not modified:
        return

    # Rewrite the zip with the modified VML files
    buf = io.BytesIO()
    with zipfile.ZipFile(out_path, "r") as zin:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                zout.writestr(name, modified.get(name, zin.read(name)))

    with open(out_path, "wb") as f:
        f.write(buf.getvalue())


def write_cells(
    file_path: str,
    cells: dict[str, list[dict]],
    value: str | None = None,
    out_path: str | None = None,
) -> str:
    """
    Write values into cells, handling both regular input cells and VML checkboxes.

    Parameters
    ----------
    file_path : str
        Source workbook path.
    cells : dict[sheet_name -> list[cell_dict]]
        Same structure returned by extract_cells — each dict must have
        "sheet", "cell", and "type" keys. When *value* is None, each dict
        must also have a "value" key; cells whose "value" is None are skipped.
    value : str | None
        Fixed value to write into every cell. If None, uses item["value"]
        from each cell dict instead.
    out_path : str | None
        Destination path. If None, overwrites file_path in-place.

    Returns
    -------
    str
        Path of the written file.
    """
    out_path = out_path or file_path

    if out_path != file_path:
        shutil.copy2(file_path, out_path)

    # Separate input cells from checkbox cells
    input_writes: dict[str, list[tuple[str, str]]] = {}   # sheet -> [(coord, val)]
    checkbox_writes: dict[str, list[tuple[str, str]]] = {}  # sheet -> [(coord, val)]

    for sheet_name, cell_list in cells.items():
        for item in cell_list:
            cell_value = value if value is not None else item.get("value")
            if cell_value is None:
                continue
            coord = item["cell"]
            if item.get("type") == "checkbox":
                checkbox_writes.setdefault(sheet_name, []).append((coord, cell_value))
            else:
                input_writes.setdefault(sheet_name, []).append((coord, cell_value))

    # --- Write regular input cells via openpyxl ---
    wb = openpyxl.load_workbook(out_path)
    written = 0
    for sheet_name, writes in input_writes.items():
        if sheet_name not in wb.sheetnames:
            print(f"  [write]  WARNING: sheet '{sheet_name}' not found — skipped")
            continue
        ws = wb[sheet_name]
        for coord, val in writes:
            ws[coord] = val
            written += 1
    wb.save(out_path)
    wb.close()

    # --- Write checkboxes ---
    if checkbox_writes:
        from excel_read_helpers import extract_checkboxes

        # Build (row, col) -> should_check map for VML update
        # Also collect linked cells to update via openpyxl
        sheet_position_updates: dict[str, dict[tuple, bool]] = {}
        linked_cell_writes: dict[str, list[tuple[str, bool]]] = {}  # sheet -> [(coord, bool)]

        for sheet_name, writes in checkbox_writes.items():
            # Re-extract checkboxes to get linked cells and verify positions
            cbs = extract_checkboxes(file_path, sheet_name)
            linked_map: dict[str, str] = {}  # coord -> linked_cell
            for cb in cbs:
                if cb["linked_cell"]:
                    coord = f"{get_column_letter(cb['col'])}{cb['row']}"
                    linked_map[coord] = cb["linked_cell"]

            for coord, val in writes:
                should_check = val == "1"
                col_str, row = coordinate_from_string(coord)
                col = column_index_from_string(col_str)
                sheet_position_updates.setdefault(sheet_name, {})[(row, col)] = should_check

                linked = linked_map.get(coord)
                if linked:
                    if "!" in linked:
                        linked = linked.split("!")[-1]
                    linked_cell_writes.setdefault(sheet_name, []).append((linked, should_check))

            written += len(writes)

        # Update VML XML
        _update_checkboxes(out_path, sheet_position_updates)

        # Update linked cells via openpyxl
        if linked_cell_writes:
            wb = openpyxl.load_workbook(out_path)
            for sheet_name, lc_writes in linked_cell_writes.items():
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                for coord, checked in lc_writes:
                    ws[coord] = checked
            wb.save(out_path)
            wb.close()

    print(f"  [write]  wrote {written} cells → {pathlib.Path(out_path).name}")
    return out_path
