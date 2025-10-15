from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path
import csv
import re


def column_index(cell_ref: str) -> int:
    letters = re.sub(r"\d", "", cell_ref or "")
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch.upper()) - 64)
    return index


def read_sheet_rows(xlsx_path: Path, sheet_name: str = "sheet1") -> list[list[str]]:
    with ZipFile(xlsx_path) as zf:
        shared_strings = []
        shared_strings_path = "xl/sharedStrings.xml"
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

        if shared_strings_path in zf.namelist():
            root = ET.fromstring(zf.read(shared_strings_path))
            for si in root.findall(f"{ns}si"):
                text_parts = []
                for t in si.iter(f"{ns}t"):
                    text_parts.append(t.text or "")
                shared_strings.append("".join(text_parts))

        sheet_path = f"xl/worksheets/{sheet_name}.xml"
        root = ET.fromstring(zf.read(sheet_path))
        sheet_data = root.find(f"{ns}sheetData")

        rows: list[list[str]] = []
        for row in sheet_data.findall(f"{ns}row"):
            cells = {}
            for cell in row.findall(f"{ns}c"):
                ref = cell.attrib.get("r", "")
                idx = column_index(ref)
                value = ""
                cell_type = cell.attrib.get("t")
                v = cell.find(f"{ns}v")
                if cell_type == "s" and v is not None:
                    value = shared_strings[int(v.text or "0")]
                elif v is not None:
                    value = v.text or ""
                else:
                    inline = cell.find(f"{ns}is")
                    if inline is not None:
                        t_elem = inline.find(f"{ns}t")
                        if t_elem is not None:
                            value = t_elem.text or ""
                cells[idx] = value
            if cells:
                max_idx = max(cells.keys())
                row_values = [""] * max_idx
                for idx, value in cells.items():
                    row_values[idx - 1] = value
                rows.append(row_values)
        return rows


def main():
    root = Path(__file__).resolve().parents[1]
    xlsx_path = root / "data" / "School_numbers" / "School numbers dataset.xlsx"
    rows = read_sheet_rows(xlsx_path)
    out_path = xlsx_path.parent / "school_numbers_sheet1.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
