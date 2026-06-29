#!/usr/bin/env python3
"""
把 full_scan_once.py 生成的 CSV 转成一个可打开的 .xlsx 工作簿。

用途：
    python3 build_excel_report.py outputs/20260606_190619 outputs/上海社会学相关实习筛选_20260606.xlsx

为什么不用 openpyxl / pandas：
    这台机器当时没有安装 openpyxl、xlsxwriter、pandas。
    为了让脚本在“只有 Python 标准库”的环境中也能运行，这里直接按 Office Open XML
    规范打包一个基础 xlsx 文件。

生成的工作簿包含：
    1. 摘要
    2. 明确匹配
    3. 近似匹配

注意：
    这是一个轻量 xlsx 生成器，只实现本项目需要的功能：
    sheet、表头样式、自动筛选、冻结首行、列宽和换行。
"""
import csv
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape


FIELD_LABELS = [
    ("match_level", "匹配等级"),
    ("title", "岗位名称"),
    ("company", "公司"),
    ("industry", "行业"),
    ("address", "工作地点"),
    ("salary", "薪资"),
    ("degree", "学历要求"),
    ("refresh_time", "刷新时间"),
    ("url", "详情链接"),
    ("matched_keywords", "命中关键词"),
    ("fit_reason", "推荐理由"),
    ("evidence", "命中证据"),
    ("description", "职位描述"),
    ("parse_status", "正文解析"),
    ("parse_warning", "解析警告"),
    ("uuid", "岗位ID"),
]


def read_csv(path):
    """读取 full_scan_once.py 输出的 UTF-8 BOM CSV。"""
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def cell_ref(row, col):
    """把行列数字转换成 Excel 坐标，例如 (1, 1) -> A1。"""
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def clean_sheet_name(name):
    """Excel sheet 名不能包含某些字符，且最长 31 个字符。"""
    return re.sub(r"[][*/\\\\?:]", "", name)[:31]


def xml_text(value):
    """清理并转义 XML 文本。"""
    value = "" if value is None else str(value)
    value = value.replace("\x00", "")
    return escape(value)


def write_cell(value, row, col, style=None):
    """生成一个单元格 XML。

    简化处理：
    - 少数字段按数字写入，便于 Excel 排序。
    - 其他字段都用 inlineStr，避免额外维护 sharedStrings.xml。
    """
    ref = cell_ref(row, col)
    style_attr = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, (int, float)) or (isinstance(value, str) and value.isdigit() and col in (2, 9, 10)):
        return f'<c r="{ref}"{style_attr}><v>{xml_text(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{xml_text(value)}</t></is></c>'


def sheet_xml(rows, widths, freeze=True):
    """生成单个 worksheet XML。"""
    max_cols = max((len(row) for row in rows), default=1)
    max_rows = max(len(rows), 1)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        f'<dimension ref="A1:{cell_ref(max_rows, max_cols)}"/>',
        '<sheetViews><sheetView workbookViewId="0">',
    ]
    if freeze:
        parts.append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
    parts.append('</sheetView></sheetViews>')
    parts.append('<sheetFormatPr defaultRowHeight="16"/>')
    parts.append("<cols>")
    for index in range(1, max_cols + 1):
        width = widths.get(index, 16)
        parts.append(f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>')
    parts.append("</cols>")
    parts.append("<sheetData>")
    for row_index, row in enumerate(rows, start=1):
        height = 34 if row_index == 1 else 54
        parts.append(f'<row r="{row_index}" ht="{height}" customHeight="1">')
        for col_index, value in enumerate(row, start=1):
            # 表头使用蓝底白字；推荐理由、证据、描述、警告列开启换行并顶部对齐。
            style = 1 if row_index == 1 else 2 if col_index in (16, 17, 18, 20) else None
            parts.append(write_cell(value, row_index, col_index, style=style))
        parts.append("</row>")
    parts.append("</sheetData>")
    if max_rows > 1 and max_cols > 1:
        parts.append(f'<autoFilter ref="A1:{cell_ref(max_rows, max_cols)}"/>')
    parts.append("</worksheet>")
    return "".join(parts)


def workbook_xml(sheet_names):
    """生成工作簿主体 XML，登记每个 sheet。"""
    sheets = []
    for index, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{xml_text(name)}" sheetId="{index}" r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        + "".join(sheets)
        + "</sheets></workbook>"
    )


def workbook_rels(sheet_count):
    """生成 workbook.xml.rels，连接工作簿、各 sheet 和 styles.xml。"""
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )


def root_rels():
    """生成根关系文件 _rels/.rels。"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def content_types(sheet_count):
    """生成 [Content_Types].xml，声明 xlsx 包内各文件类型。"""
    overrides = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(overrides)
        + "</Types>"
    )


def styles_xml():
    """生成样式表。

    当前只定义三种样式：
    0. 默认样式
    1. 表头：蓝底白字、居中、自动换行
    2. 长文本：顶部对齐、自动换行
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Arial"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def core_xml():
    """生成文档元数据。"""
    now = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>上海社会学相关实习筛选</dc:title>'
        '<dc:creator>Codex</dc:creator>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_xml(sheet_names):
    """生成 Office 扩展属性，主要登记 sheet 名。"""
    names = "".join(f"<vt:lpstr>{xml_text(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Codex</Application>"
        f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{names}</vt:vector></TitlesOfParts>'
        "</Properties>"
    )


def min_salary(salary_str):
    """从薪资字符串中提取最低日薪，无法解析时返回 None。"""
    m = re.search(r"(\d+)", salary_str or "")
    return int(m.group(1)) if m else None


def rows_for_matches(rows):
    """把 CSV 字典列表转换为工作表二维数组，过滤日薪低于 100 的岗位。"""
    header = [label for _, label in FIELD_LABELS]
    output = [header]
    for row in rows:
        sal = min_salary(row.get("salary", ""))
        if sal is not None and sal < 100:
            continue
        output.append([row.get(field, "") for field, _ in FIELD_LABELS])
    return output


def read_metadata(source):
    """读取扫描脚本写出的元数据；旧输出目录没有该文件时返回空 dict。"""
    path = source / "matches.json"
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def summary_rows(explicit, approximate, metadata=None):
    """生成摘要 sheet 的数据。"""
    metadata = metadata or {}
    levels = {"高": 0, "中": 0, "低": 0}
    for row in approximate:
        levels[row.get("match_level", "")] = levels.get(row.get("match_level", ""), 0) + 1
    scanned_jobs = metadata.get("scanned_jobs", "")
    source_desc = "来自上海、1天内发布、实习岗位列表"
    if metadata.get("list_url_page_1"):
        source_desc = metadata["list_url_page_1"]
    return [
        ["项目", "数值", "说明"],
        ["扫描岗位总数", scanned_jobs, source_desc],
        ["直接匹配", len(explicit), "岗位描述直接出现社会学、社会科学、人文社科、社科、用户研究、用户洞察或用研"],
        ["近似匹配", len(approximate), "未直接出现上述词，但包含研究、调研、行业分析、社会议题等能力信号"],
        ["近似匹配-高", levels.get("高", 0), "建议优先查看"],
        ["近似匹配-中", levels.get("中", 0), "建议按兴趣和公司筛选"],
        ["近似匹配-低", levels.get("低", 0), "保留为宽松候选，避免漏掉可迁移机会"],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""],
    ]


def build_xlsx(explicit, approximate, output, metadata=None):
    """把三个 sheet 写入 xlsx zip 包。"""
    sheets = [
        ("摘要", summary_rows(explicit, approximate, metadata), {1: 18, 2: 16, 3: 90}),
        ("直接匹配", rows_for_matches(explicit), default_widths()),
        ("近似匹配", rows_for_matches(approximate), default_widths()),
    ]
    sheet_names = [clean_sheet_name(name) for name, _, _ in sheets]
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types(len(sheets)))
        archive.writestr("_rels/.rels", root_rels())
        archive.writestr("docProps/core.xml", core_xml())
        archive.writestr("docProps/app.xml", app_xml(sheet_names))
        archive.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(sheets)))
        archive.writestr("xl/styles.xml", styles_xml())
        for index, (_, rows, widths) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(rows, widths, freeze=True))


def default_widths():
    """匹配结果 sheet 的列宽。数字是 Excel 的字符宽度单位。"""
    return {
        1: 12,
        2: 8,
        3: 26,
        4: 20,
        5: 18,
        6: 10,
        7: 32,
        8: 14,
        9: 10,
        10: 10,
        11: 12,
        12: 20,
        13: 14,
        14: 52,
        15: 28,
        16: 60,
        17: 70,
        18: 80,
        19: 14,
        20: 42,
        21: 18,
    }


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: build_excel_report.py <scan_output_dir> <output.xlsx>")
    source = Path(sys.argv[1])
    output = Path(sys.argv[2])
    explicit = read_csv(source / "explicit_matches.csv")
    approximate = read_csv(source / "approximate_matches.csv")
    metadata = read_metadata(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    build_xlsx(explicit, approximate, output, metadata)
    print(output)


if __name__ == "__main__":
    main()
