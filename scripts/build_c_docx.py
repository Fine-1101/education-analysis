"""把 C 模块的正文草稿、图表汇编成可交付给组长合并的 Word 章节。

先运行 c_questions.py 生成 outputs/c_questions/，再运行本脚本输出 docs/C_题目与作答表现.docx。
正文文字全部来自 正文草稿.md，本脚本只负责排版与插图插表，不重复维护内容。
"""

from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

PROJECT_DIR = Path(__file__).resolve().parent.parent
OUT = PROJECT_DIR / "outputs" / "c_questions"
DOCX_PATH = PROJECT_DIR / "docs" / "C_题目与作答表现.docx"

BODY_FONT = "宋体"
HEAD_FONT = "微软雅黑"

# 正文草稿的二级标题 -> 该节末尾要插入的图表（kind, 文件名, 题注, 宽度cm）
# 题注里的 图 C-n / 表 C-n 编号被 正文草稿.md 的交叉引用直接使用，调整顺序需同步改 c_questions.py
FIGURES = {
    "三": [
        ("table", "过滤规则与处理数量", "表 C-1 记录过滤规则与前后数量", None),
        ("table", "状态取值与得分分布", "表 C-2 作答状态取值与对应得分", None),
    ],
    "四": [
        ("figure", "首次正确率排序.png", "图 C-1 各题目首次正确率排序", 13.0),
        ("figure", "首次与最终正确率对比.png", "图 C-2 各题目首次与最终正确率对比", 11.5),
        ("figure", "重试提升幅度散点.png", "图 C-3 重试提升幅度与首次正确率的关系", 14.0),
        ("table", "候选异常题清单", "表 C-3 候选需复核题目清单（编号省略 Question_ 前缀）", None),
        ("table", "题目指标表", "表 C-4 38 道题目的作答表现指标（按首次正确率升序，编号省略 Question_ 前缀）", None),
    ],
    "五": [
        ("figure", "参与人数与得分率散点.png", "图 C-4 题目参与人数与平均得分率", 14.0),
    ],
}

# 题目指标表列很多，Word 里只放正文用到的 8 列，其余列（重试转对占比、个例最高尝试等）留在 CSV
METRIC_COLUMNS = {
    "title_ID": "题目编号",
    "full_score": "满分",
    "participating_students": "参与学生数",
    "submission_count": "提交量",
    "mean_attempts": "平均尝试",
    "first_accuracy": "首次正确率",
    "final_accuracy": "最终正确率",
    "mean_final_score_rate": "平均得分率",
    "max_attempts": "最高尝试",
    "retry_solved_share": "重试转对占比",
    "命中判据数": "命中判据",
}

# 每张表在 Word 里展示的原始列；未列出的表按 CSV 全部列展示
TABLE_COLUMNS = {
    "题目指标表": [
        "title_ID", "full_score", "participating_students", "submission_count",
        "mean_attempts", "first_accuracy", "final_accuracy", "mean_final_score_rate",
    ],
    "候选异常题清单": [
        "title_ID", "full_score", "participating_students", "submission_count",
        "mean_attempts", "max_attempts", "first_accuracy", "final_accuracy", "判据",
    ],
}

# 各表的列宽（cm），第一列为题目编号时留宽
COLUMN_WIDTHS = {
    "过滤规则与处理数量": [4.6, 2.4, 2.4, 1.8, 4.4],
    "状态取值与得分分布": [3.6, 2.2, 2.2, 2.2, 2.6, 2.8],
    "候选异常题清单": [2.7, 0.9, 1.3, 1.3, 1.2, 1.2, 1.5, 1.5, 3.1],
    "题目指标表": [3.4, 1.2, 1.9, 1.7, 1.7, 2.0, 2.0, 2.0],
}

# 百分比列按百分数显示，其余保留 1–2 位小数
PERCENT_COLUMNS = {
    "first_accuracy",
    "final_accuracy",
    "mean_first_score_rate",
    "mean_final_score_rate",
    "retry_solved_share",
    "first_accuracy_by_submission",
    "首次与最终差值",
    "平均得分率",
    "首次正确率",
    "最终正确率",
    "重试转对占比",
    "得分为零占比",
}


def set_run(run, size, *, bold=False, font=BODY_FONT, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    if color is not None:
        run.font.color.rgb = color


def add_paragraph(doc, text, *, size=12, bold=False, font=BODY_FONT, align=None,
                  space_after=6, first_line_indent=None, color=None):
    paragraph = doc.add_paragraph()
    if align is not None:
        paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing = 1.4
    if first_line_indent is not None:
        paragraph.paragraph_format.first_line_indent = Cm(first_line_indent)
    set_run(paragraph.add_run(text), size, bold=bold, font=font, color=color)
    return paragraph


def add_caption(doc, text):
    add_paragraph(
        doc,
        text,
        size=9,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=10,
        color=RGBColor(0x55, 0x55, 0x55),
    )


def add_figure(doc, path, caption, width_cm):
    if not path.exists():
        add_paragraph(doc, f"[缺少图片：{path.name}]", 10)
        return
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.keep_with_next = True  # 图与题注不跨页分离
    paragraph.add_run().add_picture(str(path), width=Cm(width_cm))
    add_caption(doc, caption)


def format_cell(value, column):
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)):
        if column in PERCENT_COLUMNS:
            return f"{value * 100:.1f}%"
        if float(value).is_integer():
            return str(int(value))
        return f"{value:.2f}"
    return str(value)


def mark_header_row(row):
    """跨页表格在每页重复表头，便于组长合并后阅读。"""
    element = row._element.get_or_add_trPr()
    element.append(element.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"}))


def prevent_row_split(row):
    """禁止单行跨页拆开，否则判据这类长文本列会在翻页处断成两截。"""
    element = row._element.get_or_add_trPr()
    element.append(element.makeelement(qn("w:cantSplit"), {}))


def add_table(doc, rows, columns, caption, widths_cm):
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, name in enumerate(columns):
        cell = table.rows[0].cells[index]
        cell.width = Cm(widths_cm[index])
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_after = Pt(1)
        set_run(paragraph.add_run(name), 8.5, bold=True, font=HEAD_FONT)
    mark_header_row(table.rows[0])
    prevent_row_split(table.rows[0])
    for row in rows:
        new_row = table.add_row()
        prevent_row_split(new_row)
        cells = new_row.cells
        for index, name in enumerate(columns):
            cell = cells[index]
            cell.width = Cm(widths_cm[index])
            paragraph = cell.paragraphs[0]
            paragraph.alignment = (
                WD_ALIGN_PARAGRAPH.LEFT if index == 0 else WD_ALIGN_PARAGRAPH.CENTER
            )
            paragraph.paragraph_format.space_after = Pt(1)
            set_run(paragraph.add_run(format_cell(row.get(name), name)), 8.5)
    add_caption(doc, caption)


def load_table(name):
    frame = pd.read_csv(OUT / "tables" / f"{name}.csv")
    columns = TABLE_COLUMNS.get(name, list(frame.columns))
    if name == "题目指标表":
        frame = frame.sort_values("first_accuracy")
    frame = frame[columns].rename(columns=METRIC_COLUMNS)
    # 编号前缀固定为 Question_，表内省略以腾出列宽，完整编号见同名 CSV
    if "题目编号" in frame.columns:
        frame["题目编号"] = frame["题目编号"].str.removeprefix("Question_")
    return frame.to_dict("records"), list(frame.columns)


CJK = lambda char: "\u4e00" <= char <= "\u9fff" or char in "，。；：、（）／％"


def join_lines(lines):
    """中文行之间不补空格，涉及英文或数字时才补，避免正文出现奇怪的空隙。"""
    text = lines[0]
    for nxt in lines[1:]:
        if text and nxt and not (CJK(text[-1]) and CJK(nxt[0])):
            text += " "
        text += nxt
    return text


def load_draft_sections():
    """正文草稿按二级标题切块，块内空行分段，段内硬换行合并。"""
    sections = []
    heading, lines, paragraphs = None, [], []

    def close_paragraph():
        if lines:
            paragraphs.append(join_lines(lines))
            lines.clear()

    def close_section():
        close_paragraph()
        if heading is not None:
            sections.append((heading, list(paragraphs)))
        paragraphs.clear()

    for raw in (OUT / "正文草稿.md").read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if line.startswith("## "):
            close_section()
            heading = line[3:].strip()
        elif not line:
            close_paragraph()
        elif not line.startswith("# "):
            lines.append(line)
    close_section()
    return sections


def main():
    doc = Document()
    doc.styles["Normal"].font.name = BODY_FONT
    doc.styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    for section in doc.sections:
        # python-docx 模板默认 Letter，合并进团队 Word 前统一成 A4
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = section.bottom_margin = Cm(2.5)
        section.left_margin = section.right_margin = Cm(2.5)

    add_paragraph(doc, "C 题目与作答表现", size=18, bold=True, font=HEAD_FONT,
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_paragraph(
        doc,
        "主要完成人：C ｜ 数据来源：ChinaVis 2024 赛道1 ｜ 代码：scripts/c_questions.py ｜ "
        "口径依据：outputs/c_questions/audit/口径与指标公式.md",
        size=9,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=14,
        color=RGBColor(0x77, 0x77, 0x77),
    )

    for heading, paragraphs in load_draft_sections():
        add_paragraph(
            doc, heading, size=14, bold=True, font=HEAD_FONT, space_after=8
        ).paragraph_format.keep_with_next = True
        for body in paragraphs:
            add_paragraph(doc, body, first_line_indent=0.74, space_after=8)
        for kind, name, caption, width in FIGURES.get(heading[:1], []):
            if kind == "figure":
                add_figure(doc, OUT / "figures" / name, caption, width)
            else:
                rows, columns = load_table(name)
                add_table(doc, rows, columns, caption, COLUMN_WIDTHS[name])

    add_paragraph(
        doc, "附：本模块交付文件", size=14, bold=True, font=HEAD_FONT, space_after=8
    ).paragraph_format.keep_with_next = True
    for line in [
        "代码：scripts/c_questions.py（顺序执行，输出全部落在 outputs/c_questions/）",
        "表格：tables/题目指标表.csv、tables/候选异常题清单.csv、tables/质量检查表.csv、"
        "tables/字段说明.csv、tables/过滤规则与处理数量.csv",
        "明细：intermediate/学生题目作答汇总.csv（50482 个学生×题目组合）",
        "审计：audit/ 下 11 个留档文件与口径与指标公式.md、官方数据说明摘录.md",
        "公共职责：全组数据质量问题清单.md（13 项，含移交 A、B、D 的待办）",
    ]:
        add_paragraph(doc, f"· {line}", size=10.5, space_after=2)

    DOCX_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX_PATH)
    print("已保存：", DOCX_PATH)


if __name__ == "__main__":
    main()
