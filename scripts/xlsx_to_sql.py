"""Convert one XLSX or a directory of split XLSX files into auditable SQLite SQL.

This importer reads cached cell values. It does not evaluate formulas; formula evaluation
belongs to Milestone 3. Each workbook sheet becomes a source table with Excel columns
A, B, ..., while row-one text is kept as a field label in metadata.
"""
from __future__ import annotations
import json, pathlib, re, sys, zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"x": MAIN}

def q(value: str) -> str: return "'" + value.replace("'", "''") + "'"
def qi(value: str) -> str: return '"' + value.replace('"', '""') + '"'
def col_index(ref: str) -> int:
    label = re.match(r"[A-Z]+", ref).group(0)
    result = 0
    for char in label: result = result * 26 + ord(char) - 64
    return result - 1
def col_label(index: int) -> str:
    result = ""; index += 1
    while index: index, rem = divmod(index - 1, 26); result = chr(65 + rem) + result
    return result
def scalar(text: str | None, kind: str | None, shared: list[str], inline: str | None):
    if kind == "inlineStr": return inline or ""
    if text is None: return None
    if kind == "s": return shared[int(text)]
    if kind == "b": return 1 if text == "1" else 0
    if kind in ("str", "e"): return text
    try:
        value = float(text)
        return int(value) if value.is_integer() else value
    except ValueError: return text

def workbook(path: pathlib.Path):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if "xl/sharedStrings.xml" in z.namelist():
            root=ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared=["".join(t.text or "" for t in item.findall(".//x:t",NS)) for item in root]
        wb=ET.fromstring(z.read("xl/workbook.xml")); rel=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets={item.attrib["Id"]:item.attrib["Target"] for item in rel}
        for item in wb.find("x:sheets",NS):
            target=targets[item.attrib[f"{{{REL}}}id"]]; target=target.lstrip("/") if target.startswith("/") else "xl/"+target
            xml=ET.fromstring(z.read(target)); rows=[]; max_col=-1
            for row in xml.findall(".//x:sheetData/x:row",NS):
                values={}
                for cell in row.findall("x:c",NS):
                    c=col_index(cell.attrib["r"]); max_col=max(max_col,c); node=cell.find("x:v",NS); inline=cell.find("x:is",NS)
                    values[c]=scalar(node.text if node is not None else None,cell.attrib.get("t"),shared,"".join(t.text or "" for t in inline.findall(".//x:t",NS)) if inline is not None else None)
                rows.append((int(row.attrib["r"]),values))
            yield item.attrib["name"],rows,max_col

def main(source: pathlib.Path, output: pathlib.Path):
    files=[source] if source.is_file() else sorted(source.glob("*.xlsx"))
    if not files: raise SystemExit("未找到 .xlsx 文件")
    statements=["PRAGMA foreign_keys=ON;","CREATE TABLE IF NOT EXISTS _source_column_labels(table_name TEXT NOT NULL,column_name TEXT NOT NULL,label TEXT NOT NULL,PRIMARY KEY(table_name,column_name));"]
    seen=set()
    for file in files:
        for name,rows,max_col in workbook(file):
            table=name
            if table in seen: table=f"{file.stem}_{name}"
            seen.add(table); columns=[col_label(i) for i in range(max_col+1)]
            first=next((values for number,values in rows if number==1),{})
            data_rows=[values for number,values in rows if number>1 and any(value not in (None,'') for value in values.values())]
            affinities=[]
            for index in range(len(columns)):
                values=[row.get(index) for row in data_rows if row.get(index) not in (None,'')]
                affinities.append("REAL" if values and all(isinstance(value,(int,float)) for value in values) else "TEXT")
            statements += [f"DROP TABLE IF EXISTS {qi(table)};",f"DELETE FROM _source_column_labels WHERE table_name={q(table)};",f"CREATE TABLE {qi(table)} ({','.join(qi(c)+' '+affinities[i] for i,c in enumerate(columns))});"]
            statements += [f"INSERT INTO _source_column_labels VALUES({q(table)},{q(c)},{q(str(first.get(i) if first.get(i) not in (None,'') else c))});" for i,c in enumerate(columns)]
            for values in data_rows:
                ordered=[values.get(i) for i in range(len(columns))]
                sql_values=','.join('NULL' if v is None else str(v) if isinstance(v,(int,float)) else q(v) for v in ordered)
                statements.append(f"INSERT INTO {qi(table)} VALUES({sql_values});")
    output.write_text("\n".join(statements)+"\n",encoding="utf-8")
    print(json.dumps({"files":len(files),"tables":len(seen),"output":str(output)},ensure_ascii=False))

if __name__ == "__main__":
    if len(sys.argv)!=3: raise SystemExit("usage: xlsx_to_sql.py SOURCE OUTPUT")
    main(pathlib.Path(sys.argv[1]),pathlib.Path(sys.argv[2]))
