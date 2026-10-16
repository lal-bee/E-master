"""完整预检查后，修改原始 OOXML 包中的目标单元格。"""
from __future__ import annotations

import copy
import io
import posixpath
from dataclasses import dataclass, field
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.utils.cell import column_index_from_string, range_boundaries
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH

from .rules import (CODE_FIELD, DATE_FIELDS, DEFAULT_MAPPING, DEVICE_FIELD, NAME_FIELD,
                    TARGET_FIELDS, YearOnlyError, load_mapping, normalize_date, normalize_device_code)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorkbookError(ValueError):
    """文件无法读取或不符合本工具的输入要求。"""


@dataclass(frozen=True)
class Issue:
    kind: str
    sheet: str
    row: int
    field: str
    cell: str
    value: str
    message: str


@dataclass
class ConversionResult:
    issues: list[Issue] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    skipped_sheets: list[str] = field(default_factory=list)
    rows: int = 0
    subcategory_filled: int = 0
    dates_converted: int = 0
    dates_cleared: int = 0
    codes_formatted: int = 0
    output: bytes | None = field(default=None, repr=False)

    @property
    def success(self):
        return self.output is not None and not self.issues

    def save(self, path: Path, *, source: Path | None = None):
        """独占创建新文件；绝不覆盖原件或已有结果。"""
        path = Path(path)
        if source is not None and path.resolve() == Path(source).resolve():
            raise WorkbookError("结果文件不能覆盖原始 Excel")
        if not self.success:
            raise WorkbookError("预检查未通过，不能生成结果文件")
        with path.open("xb") as stream:
            stream.write(self.output)
        return path


def _text(element):
    # 只读实际文本，排除拼音注音节点。
    return "".join(t.text or "" for t in element.findall(f"{{{NS}}}t") +
                   element.findall(f"{{{NS}}}r/{{{NS}}}t"))


def _value(cell, strings):
    if cell is None:
        return None
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        inline = cell.find(f"{{{NS}}}is")
        return _text(inline) if inline is not None else None
    value = cell.find(f"{{{NS}}}v")
    if value is None or value.text is None:
        return None
    if kind == "s":
        return strings[int(value.text)]
    return value.text


def _parts(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = {item.get("Id"): item for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
    properties = workbook.find(f"{{{NS}}}workbookPr")
    epoch = MAC_EPOCH if properties is not None and properties.get("date1904") in {"1", "true"} else WINDOWS_EPOCH
    for sheet in workbook.findall(f"{{{NS}}}sheets/{{{NS}}}sheet"):
        if sheet.get("state", "visible") != "visible" or sheet.get("name", "").lower().startswith("hidden@"):
            yield sheet, None, epoch
            continue
        relationship = rels[sheet.get(f"{{{REL}}}id")]
        if relationship.get("TargetMode") == "External":
            raise WorkbookError("不支持外部工作表")
        target = relationship.get("Target", "")
        part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
        if not part.startswith("xl/") or part not in archive.namelist():
            raise WorkbookError("工作表部件路径无效")
        yield sheet, part, epoch


def _formats(archive):
    if "xl/styles.xml" not in archive.namelist():
        return []
    styles = ET.fromstring(archive.read("xl/styles.xml"))
    formats = dict(BUILTIN_FORMATS)
    formats.update({int(n.get("numFmtId")): n.get("formatCode", "")
                    for n in styles.findall(f"{{{NS}}}numFmts/{{{NS}}}numFmt")})
    return [formats.get(int(n.get("numFmtId", "0")), "")
            for n in styles.findall(f"{{{NS}}}cellXfs/{{{NS}}}xf")]


def _write_cells(xml: bytes, changes: dict[str, str | None]) -> bytes:
    """只替换值节点，保留原单元格样式、其他节点和命名空间。"""
    document = minidom.parseString(xml)
    root = document.documentElement
    prefix = root.prefix + ":" if root.prefix else ""
    def element(name):
        return document.createElementNS(NS, prefix + name)
    data = root.getElementsByTagNameNS(NS, "sheetData")[0]
    rows = {int(n.getAttribute("r")): n for n in data.childNodes
            if n.nodeType == n.ELEMENT_NODE and n.localName == "row"}
    for coordinate, value in changes.items():
        letters = coordinate.rstrip("0123456789")
        row = rows[int(coordinate[len(letters):])]
        cells = [c for c in row.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName == "c"]
        cell = next((c for c in cells if c.getAttribute("r") == coordinate), None)
        if cell is None:
            if value is None:
                continue
            cell = element("c")
            cell.setAttribute("r", coordinate)
            column = column_index_from_string(letters)
            following = next((c for c in cells if column_index_from_string(c.getAttribute("r").rstrip("0123456789")) > column), None)
            row.insertBefore(cell, following)
        for child in list(cell.childNodes):
            if child.nodeType == child.ELEMENT_NODE and child.localName in {"v", "is", "f"}:
                cell.removeChild(child)
        if cell.hasAttribute("t"):
            cell.removeAttribute("t")
        if value is not None:
            cell.setAttribute("t", "inlineStr")
            inline, text = element("is"), element("t")
            text.setAttribute("xml:space", "preserve")
            text.appendChild(document.createTextNode(value))
            inline.appendChild(text)
            extension = next((c for c in cell.childNodes if c.nodeType == c.ELEMENT_NODE
                              and c.namespaceURI == NS and c.localName == "extLst"), None)
            cell.insertBefore(inline, extension)
    return document.toxml(encoding="utf-8")


def convert_workbook(source: Path | bytes, mapping_path: Path = DEFAULT_MAPPING) -> ConversionResult:
    """扫描整本、收集全部阻断问题；成功后才写入副本并产生输出字节。"""
    if isinstance(source, bytes):
        content = source
    else:
        source = Path(source)
        if source.suffix.lower() != ".xlsx":
            raise WorkbookError("仅支持 .xlsx 文件")
        content = source.read_bytes()
    result, pending = ConversionResult(), {}
    try:
        with ZipFile(io.BytesIO(content)) as archive:
            if sum(info.file_size for info in archive.infolist()) > 256 * 1024 * 1024:
                raise WorkbookError("Excel 解压后过大，请缩小文件后重试")
            strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                strings = [_text(n) for n in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
            formats = _formats(archive)
            mapping = None
            for metadata, part, epoch in _parts(archive):
                name = metadata.get("name", "")
                if metadata.get("state", "visible") != "visible" or name.lower().startswith("hidden@"):
                    result.skipped_sheets.append(name)
                    continue
                xml = archive.read(part)
                root = ET.fromstring(xml)
                data = root.find(f"{{{NS}}}sheetData")
                if data is None:
                    result.skipped_sheets.append(name)
                    continue
                header = next((r for r in data if r.get("r") == "1"), None)
                fields = {}
                if header is not None:
                    for cell in header:
                        value = _value(cell, strings)
                        if value in TARGET_FIELDS:
                            if value in fields:
                                result.issues.append(Issue("STRUCTURE", name, 1, value, cell.get("r"), value, "目标表头重复"))
                            fields[value] = cell.get("r").rstrip("0123456789")
                if not fields:
                    result.skipped_sheets.append(name)
                    continue
                result.sheets.append(name)
                if NAME_FIELD in fields and CODE_FIELD not in fields:
                    result.issues.append(Issue("STRUCTURE", name, 1, NAME_FIELD, fields[NAME_FIELD] + "1", NAME_FIELD,
                                               "缺少设备小类编码列，请补充该表头"))
                # 合并单元格和跨列公式不能安全地按行改写。
                regions = [(n.get("ref"), "目标字段含合并单元格，请先取消合并", "")
                           for n in root.findall(f"{{{NS}}}mergeCells/{{{NS}}}mergeCell")]
                regions += [(n.get("ref"), "目标字段位于公式范围，请先转换为明确值", "=" + (n.text or ""))
                            for n in data.findall(f"{{{NS}}}row/{{{NS}}}c/{{{NS}}}f") if n.get("ref")]
                for reference, message, original in regions:
                    first_col, first_row, last_col, last_row = range_boundaries(reference)
                    if last_row < 2:
                        continue
                    for field_name, letters in fields.items():
                        if first_col <= column_index_from_string(letters) <= last_col:
                            number = max(2, first_row)
                            coordinate = letters + str(number)
                            cell = data.find(f"{{{NS}}}row[@r='{number}']/{{{NS}}}c[@r='{coordinate}']")
                            if original and cell is not None and cell.find(f"{{{NS}}}f") is not None:
                                continue  # 公式本身由下面逐行检查定位。
                            result.issues.append(Issue("STRUCTURE", name, number, field_name, coordinate,
                                                       original or _value(cell, strings) or "", message + f"（{reference}）"))
                if NAME_FIELD in fields or CODE_FIELD in fields:
                    if mapping is None:
                        mapping = load_mapping(mapping_path)
                changes = {}
                for row in data:
                    number = int(row.get("r", "0"))
                    if number < 2:
                        continue
                    cells = {c.get("r"): c for c in row if c.tag == f"{{{NS}}}c"}
                    if not any(c.find(f"{{{NS}}}f") is not None or _value(c, strings) not in {None, ""} for c in cells.values()):
                        continue
                    result.rows += 1
                    def get(field):
                        coordinate = fields[field] + str(number)
                        cell = cells.get(coordinate)
                        return coordinate, cell, _value(cell, strings)
                    formula_fields = set()
                    for field_name in fields:
                        coordinate, cell, value = get(field_name)
                        if cell is not None and cell.find(f"{{{NS}}}f") is not None:
                            formula_fields.add(field_name)
                            formula = cell.find(f"{{{NS}}}f")
                            result.issues.append(Issue("STRUCTURE", name, number, field_name, coordinate,
                                                       "=" + (formula.text or ""), "目标字段含公式，请先转换为明确值"))
                    if NAME_FIELD in fields or CODE_FIELD in fields:
                        coordinate, _, value = get(NAME_FIELD) if NAME_FIELD in fields else ("", None, None)
                        code_value = get(CODE_FIELD)[2] if CODE_FIELD in fields else None
                        if not ({NAME_FIELD, CODE_FIELD} & formula_fields):
                            if value not in {None, ""} or code_value not in {None, ""}:
                                if value not in mapping:
                                    if NAME_FIELD not in fields:
                                        result.issues.append(Issue("UNMATCHED", name, number, CODE_FIELD,
                                                                  get(CODE_FIELD)[0], code_value or "", "有编码但缺少设备小类名称，无法匹配"))
                                    else:
                                        result.issues.append(Issue("UNMATCHED", name, number, NAME_FIELD, coordinate, value or "", "设备小类名称找不到对应编码"))
                                elif CODE_FIELD in fields:
                                    changes[fields[CODE_FIELD] + str(number)] = mapping[value]
                                    result.subcategory_filled += 1
                    for field_name in DATE_FIELDS:
                        if field_name not in fields or field_name in formula_fields:
                            continue
                        coordinate, cell, value = get(field_name)
                        numeric = cell is not None and cell.get("t", "n") == "n"
                        # 0.00 明确表示小数年月时恢复 Excel 丢掉的末尾零（2026.10）。
                        style = int(cell.get("s", "0")) if cell is not None else 0
                        fmt = formats[style] if style < len(formats) else ""
                        if numeric and value and fmt == "0.00" and "." in value and len(value.split(".")[0]) == 4:
                            value = value.ljust(len(value.split(".")[0]) + 3, "0")
                        try:
                            valid_type = cell is None or cell.get("t", "n") in {"n", "s", "inlineStr", "d"}
                            converted = normalize_date(value, numeric=numeric, epoch=epoch) if valid_type else None
                        except YearOnlyError:
                            result.issues.append(Issue("YEAR_ONLY", name, number, field_name, coordinate, value, "日期只有年份，请补充月份或完整日期"))
                            continue
                        changes[coordinate] = converted
                        if value is not None and str(value).strip():
                            if converted is None:
                                result.dates_cleared += 1
                            else:
                                result.dates_converted += 1
                    if DEVICE_FIELD in fields and DEVICE_FIELD not in formula_fields:
                        coordinate, cell, value = get(DEVICE_FIELD)
                        style = int(cell.get("s", "0")) if cell is not None else 0
                        fmt = formats[style] if style < len(formats) else ""
                        numeric = cell is not None and cell.get("t", "n") == "n"
                        converted = normalize_device_code(value, numeric=numeric, number_format=fmt)
                        changes[coordinate] = converted
                        if converted is not None:
                            result.codes_formatted += 1
                pending[part] = (xml, changes)
            if not result.sheets:
                raise WorkbookError("没有找到包含目标字段的业务工作表；请确认第 1 行为表头")
            if result.issues:
                # 预检查阶段只计算候选值，不修改任何原始 XML、不生成输出。
                result.subcategory_filled = result.dates_converted = result.dates_cleared = result.codes_formatted = 0
                return result
            rewritten = {part: _write_cells(xml, changes) if changes else xml
                         for part, (xml, changes) in pending.items()}
            output = io.BytesIO()
            with ZipFile(output, "w") as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(copy.copy(info), rewritten.get(info.filename, archive.read(info.filename)))
            result.output = output.getvalue()
    except (BadZipFile, ET.ParseError, KeyError, IndexError, TypeError) as exc:
        raise WorkbookError("无法读取 Excel，请确认文件为有效且未加密的 .xlsx") from exc
    return result
