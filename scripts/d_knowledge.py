from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch

# 根据脚本位置定位项目目录，避免依赖个人电脑的绝对路径
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw" / "chinavis2024_education"
OUT_DIR = ROOT / "outputs" / "d_knowledge"

FIGURE_DPI = 300

STUDENT_FILE = RAW_DIR / "Data_StudentInfo.csv"
TITLE_FILE = RAW_DIR / "Data_TitleInfo.csv"
RECORD_DIR = RAW_DIR / "Data_SubmitRecord"

# A 模块已确定的日志去重口径：九个业务字段完全相同即视为重复记录
BUSINESS_COLUMNS = [
    "class",
    "student_ID",
    "title_ID",
    "time",
    "state",
    "score",
    "method",
    "memory",
    "timeconsume",
]

# 正常作答状态；源数据中存在 6 条 U+FFFD 乱码值，单独审计
STANDARD_STATES = [
    "Absolutely_Correct",
    "Partially_Correct",
    "Absolutely_Error",
    "Error1",
    "Error2",
    "Error3",
    "Error4",
    "Error5",
    "Error6",
    "Error7",
    "Error8",
    "Error9",
]

# 字段含义，供字段说明表使用（只写数据能够支持的解释，不臆造业务含义）
FIELD_MEANINGS = {
    ("Data_StudentInfo", "index"): "行号（各表独立）",
    ("Data_StudentInfo", "student_ID"): "学生匿名编号",
    ("Data_StudentInfo", "sex"): "性别",
    ("Data_StudentInfo", "age"): "年龄（岁）",
    ("Data_StudentInfo", "major"): "专业编码",
    ("Data_TitleInfo", "index"): "行号（各表独立）",
    ("Data_TitleInfo", "title_ID"): "题目编号",
    ("Data_TitleInfo", "score"): "题目满分",
    ("Data_TitleInfo", "knowledge"): "一级知识点编码",
    ("Data_TitleInfo", "sub_knowledge"): "二级知识点编码",
    ("Data_SubmitRecord", "index"): "行号（文件内 0 起）",
    ("Data_SubmitRecord", "class"): "班级",
    ("Data_SubmitRecord", "time"): "提交时间（Unix 秒）",
    ("Data_SubmitRecord", "state"): "作答状态",
    ("Data_SubmitRecord", "score"): "本题得分",
    ("Data_SubmitRecord", "title_ID"): "题目编号",
    ("Data_SubmitRecord", "method"): "解题方法编码",
    ("Data_SubmitRecord", "memory"): "内存占用（单位未说明）",
    ("Data_SubmitRecord", "timeconsume"): "耗时（单位未说明，含 '--'）",
    ("Data_SubmitRecord", "student_ID"): "学生编号",
    ("Data_SubmitRecord", "source_file"): "来源文件名（程序附加）",
}

FIELD_NOTES = {
    ("Data_StudentInfo", "index"): "不作为关联键",
    ("Data_StudentInfo", "student_ID"): "学生表内唯一，20 位小写十六进制",
    ("Data_TitleInfo", "title_ID"): "一题可对应多个知识点，故不唯一",
    ("Data_TitleInfo", "score"): "经比对，等于该题观察到的最大得分，按满分处理",
    ("Data_TitleInfo", "knowledge"): "8 个取值，均为 5 位随机码，无名称映射",
    ("Data_TitleInfo", "sub_knowledge"): "15 个取值，均以 <knowledge>_ 开头",
    ("Data_SubmitRecord", "index"): "每个文件从 0 开始，跨文件重复，不能作唯一提交 ID",
    ("Data_SubmitRecord", "class"): "正常为 Class1-Class15，另有 6 条字面值 class",
    ("Data_SubmitRecord", "time"): "整数秒级时间戳；数据未附说明，单位按数值范围推断",
    ("Data_SubmitRecord", "state"): "12 个正常取值 + 6 条 U+FFFD 乱码（源文件损坏）",
    ("Data_SubmitRecord", "score"): "0 至该题满分；与 state 逻辑一致",
    ("Data_SubmitRecord", "memory"): "0-65536，知识点与掌握情况模块不使用",
    ("Data_SubmitRecord", "timeconsume"): "含 2612 条 '--'；知识点与掌握情况模块不使用",
}


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def ensure_output_dirs() -> dict:
    """创建输出目录，允许重复运行；模块说明与流程图放在 outputs/d_knowledge/report/。"""
    dirs = {}
    for name in ("figures", "tables", "audit", "intermediate", "report"):
        path = OUT_DIR / name
        path.mkdir(parents=True, exist_ok=True)
        dirs[name] = path
    return dirs


def write_csv(frame: pd.DataFrame, path: Path) -> None:
    """统一以 utf-8-sig 保存 CSV，便于 Excel 直接打开中文。"""
    frame.to_csv(path, index=False, encoding="utf-8-sig")


def make_check(
    check_id: str,
    target: str,
    item: str,
    raw_count: int,
    abnormal_count: int,
    handled_count: int,
    rule: str,
    after_count: int,
    note: str = "",
) -> dict:
    """构造一行质量检查记录。"""
    return {
        "check_id": check_id,
        "target": target,
        "item": item,
        "raw_count": raw_count,
        "abnormal_count": abnormal_count,
        "handled_count": handled_count,
        "rule": rule,
        "after_count": after_count,
        "note": note,
    }


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------
def read_students() -> pd.DataFrame:
    """读取学生信息表，所有字段按文本读取，避免编号被推断为数值。"""
    students = pd.read_csv(
        STUDENT_FILE,
        dtype="string",
    )
    print("学生信息表：", STUDENT_FILE)
    print("行数、列数：", students.shape)
    return students


def read_titles() -> pd.DataFrame:
    """读取题目-知识点表，所有编码按文本读取。"""
    titles = pd.read_csv(TITLE_FILE, dtype="string")
    titles["score_num"] = pd.to_numeric(titles["score"], errors="coerce")
    print("\n题目-知识点表：", TITLE_FILE)
    print("行数、列数：", titles.shape)
    return titles


def read_submission_files() -> tuple[pd.DataFrame, pd.DataFrame]:
    """读取全部班级提交文件，返回合并记录与逐文件概况。"""
    files = sorted(
        RECORD_DIR.glob("SubmitRecord-Class*.csv"),
        key=lambda path: int(path.stem.split("Class")[1]),
    )
    if not files:
        raise FileNotFoundError(f"未找到提交记录文件，请检查目录：{RECORD_DIR}")

    parts = []
    profiles = []
    for file in files:
        part = pd.read_csv(file, dtype="string")
        part["source_file"] = file.name
        parts.append(part)
        profiles.append(
            {
                "table": "Data_SubmitRecord",
                "file": file.name,
                "rows": len(part),
                "columns": part.shape[1] - 1,  # 不计程序附加的 source_file
                "size_bytes": file.stat().st_size,
            }
        )

    records = pd.concat(parts, ignore_index=True)
    profile = pd.DataFrame(profiles)
    print("\n提交记录文件数：", len(files))
    print("提交记录总行数：", len(records))
    return records, profile


# ---------------------------------------------------------------------------
# 字段与表级检查
# ---------------------------------------------------------------------------
def build_field_dictionary(
    students: pd.DataFrame,
    titles: pd.DataFrame,
    records: pd.DataFrame,
) -> pd.DataFrame:
    """生成字段说明表：原始类型、缺失数、唯一值数、示例。"""
    rows = []
    for table, frame in (
        ("Data_StudentInfo", students),
        ("Data_TitleInfo", titles.drop(columns=["score_num"])),
        ("Data_SubmitRecord", records),
    ):
        for column in frame.columns:
            series = frame[column]
            example = ""
            if series.notna().any():
                example = str(series.dropna().iloc[0])
            rows.append(
                {
                    "table": table,
                    "field": column,
                    "meaning": FIELD_MEANINGS.get((table, column), ""),
                    "read_dtype": str(series.dtype),
                    "missing": int(series.isna().sum()),
                    "unique": int(series.nunique(dropna=True)),
                    "example": example,
                    "note": FIELD_NOTES.get((table, column), ""),
                }
            )
    return pd.DataFrame(rows)


def check_student_table(students: pd.DataFrame) -> list:
    """学生表字段与取值检查。"""
    checks = []
    duplicate_ids = int(students["student_ID"].duplicated().sum())
    checks.append(
        make_check(
            "S01", "学生表", "student_ID 唯一性", len(students), duplicate_ids, 0,
            "不处理（先检查）", len(students) - duplicate_ids,
            f"唯一学生数：{students['student_ID'].nunique()}",
        )
    )
    missing = int(students.isna().sum().sum())
    checks.append(
        make_check(
            "S02", "学生表", "字段缺失值", len(students) * students.shape[1], missing, 0,
            "不处理（先检查）", len(students) * students.shape[1],
            "各字段均无缺失",
        )
    )
    sex_values = set(students["sex"].dropna().unique())
    checks.append(
        make_check(
            "S03", "学生表", "sex 取值合法性", len(students),
            0 if sex_values <= {"male", "female"} else len(students), 0,
            "不处理（先检查）", len(students),
            f"取值：{sorted(sex_values)}",
        )
    )
    ages = pd.to_numeric(students["age"], errors="coerce")
    checks.append(
        make_check(
            "S04", "学生表", "age 数值范围", len(students),
            int(ages.isna().sum() + ages.lt(14).sum() + ages.gt(60).sum()), 0,
            "不处理（先检查）", len(students),
            f"范围 {int(ages.min())}-{int(ages.max())} 岁",
        )
    )
    return checks


def check_title_table(titles: pd.DataFrame) -> tuple[list, dict]:
    """题目-知识点表的字段、满分、层级与关联检查。"""
    checks = []
    unique_titles = titles["title_ID"].nunique()
    multi_rows = len(titles) - unique_titles

    checks.append(
        make_check(
            "T01", "题目表", "title_ID 唯一性", len(titles), multi_rows, 0,
            "不删除：重复题号为合法的一题多知识点展开", unique_titles,
            f"唯一题目 {unique_titles} 道，{multi_rows} 行为多关联展开",
        )
    )

    missing = int(titles.isna().sum().sum())
    title_field_count = titles.shape[1] - 1  # 不计程序附加的 score_num
    checks.append(
        make_check(
            "T02", "题目表", "字段缺失值", len(titles) * title_field_count, missing, 0,
            "不处理（先检查）", len(titles) * title_field_count,
            "title_ID/knowledge/sub_knowledge/score 均无缺失",
        )
    )

    triple_dup = int(
        titles.duplicated(subset=["title_ID", "knowledge", "sub_knowledge"]).sum()
    )
    checks.append(
        make_check(
            "T03", "题目表", "题目-主知识点-子知识点完全重复", len(titles), triple_dup,
            triple_dup, "重复三元组去重保留一条", len(titles) - triple_dup,
            "实测无完全重复关联",
        )
    )

    sub_pair_dup = int(
        titles.duplicated(subset=["title_ID", "sub_knowledge"]).sum()
    )
    checks.append(
        make_check(
            "T04", "题目表", "题目-子知识点重复关联", len(titles), sub_pair_dup,
            sub_pair_dup, "题目-子知识点映射去重", len(titles) - sub_pair_dup,
            "实测无重复",
        )
    )

    # 同一主知识点下挂两个子知识点时，两行会共享 (title_ID, knowledge)。
    # 这是合法关联，合并为一条主知识点映射，不计为重复异常。
    knowledge_pair_rows = int(
        len(titles[["title_ID", "knowledge"]].drop_duplicates())
    )
    merged_rows = len(titles) - knowledge_pair_rows
    checks.append(
        make_check(
            "T05", "题目表", "题目-主知识点映射合并", len(titles), 0, merged_rows,
            "同一主知识点下的多个子知识点合并为一条主知识点关联（合法关联，非重复）",
            knowledge_pair_rows,
            "合并的 {} 行来自同一主知识点挂多个子知识点的题目".format(merged_rows),
        )
    )

    score_conflict = int(titles.groupby("title_ID")["score_num"].nunique().gt(1).sum())
    checks.append(
        make_check(
            "T06", "题目表", "同一题目满分冲突", unique_titles, score_conflict,
            score_conflict, "满分冲突题目剔除并审计", unique_titles - score_conflict,
            "实测无冲突",
        )
    )

    blank_knowledge = int(
        titles["knowledge"].isna().sum()
        + titles["knowledge"].astype("string").str.strip().eq("").fillna(False).sum()
    )
    checks.append(
        make_check(
            "T07", "题目表", "knowledge 缺失或空白", len(titles), blank_knowledge,
            0, "不处理（先检查）", len(titles),
            f"知识点 {titles['knowledge'].nunique()} 个",
        )
    )

    prefix = titles["sub_knowledge"].str.split("_").str[0]
    prefix_mismatch = int((prefix != titles["knowledge"]).sum())
    checks.append(
        make_check(
            "T08", "题目表", "sub_knowledge 前缀与 knowledge 不符", len(titles),
            prefix_mismatch, 0, "不处理（先检查）", len(titles),
            "全部子知识点均以父知识点编码加下划线开头",
        )
    )

    parent_counts = titles.groupby("sub_knowledge")["knowledge"].nunique()
    multi_parent = int(parent_counts.gt(1).sum())
    checks.append(
        make_check(
            "T09", "题目表", "一个子知识点归属多个主知识点",
            int(titles["sub_knowledge"].nunique()), multi_parent, 0,
            "不处理（先检查）", int(titles["sub_knowledge"].nunique()),
            "层级关系唯一，无跨父节点",
        )
    )

    knowledge_per_title = titles.groupby("title_ID")["knowledge"].nunique()
    multi_knowledge_titles = int(knowledge_per_title.gt(1).sum())
    checks.append(
        make_check(
            "T10", "题目表", "一题对应多个主知识点", unique_titles,
            multi_knowledge_titles, 0, "保留真实多关联，不删除题号", unique_titles,
            "分布：1 个知识点 {} 道，2 个知识点 {} 道".format(
                int(knowledge_per_title.eq(1).sum()), multi_knowledge_titles
            ),
        )
    )

    sub_per_title = titles.groupby("title_ID")["sub_knowledge"].nunique()
    checks.append(
        make_check(
            "T11", "题目表", "一题对应多个子知识点", unique_titles,
            int(sub_per_title.gt(1).sum()), 0,
            "保留真实多关联，不删除题号", unique_titles,
            "分布：1 个子知识点 {} 道，2 个子知识点 {} 道".format(
                int(sub_per_title.eq(1).sum()), int(sub_per_title.gt(1).sum())
            ),
        )
    )

    non_numeric_score = int(titles["score_num"].isna().sum())
    checks.append(
        make_check(
            "T12", "题目表", "score 非数字", len(titles), non_numeric_score, 0,
            "不处理（先检查）", len(titles),
            "满分均可解析为数值，取值 {}".format(
                sorted(titles["score_num"].dropna().unique().tolist())
            ),
        )
    )

    association = {
        "unique_titles": unique_titles,
        "knowledge_count": int(titles["knowledge"].nunique()),
        "sub_knowledge_count": int(titles["sub_knowledge"].nunique()),
        "title_knowledge_rows": int(
            len(titles[["title_ID", "knowledge"]].drop_duplicates())
        ),
        "title_sub_rows": int(
            len(titles[["title_ID", "sub_knowledge"]].drop_duplicates())
        ),
        "multi_knowledge_titles": multi_knowledge_titles,
        "multi_sub_titles": int(sub_per_title.gt(1).sum()),
    }
    return checks, association


# ---------------------------------------------------------------------------
# 知识点映射与关联结构
# ---------------------------------------------------------------------------
def build_knowledge_maps(titles: pd.DataFrame) -> dict:
    """分别构建题目→主知识点、题目→子知识点、题目→满分三套映射。

    注意：两套映射必须分开使用。若把子知识点展开结果直接汇总到主知识点，
    同一主知识点下挂两个子知识点的题目会被重复计入。
    """
    question_info = (
        titles.groupby("title_ID", as_index=False)
        .agg(full_score=("score_num", "first"))
    )
    title_knowledge = titles[["title_ID", "knowledge"]].drop_duplicates()
    title_sub = titles[["title_ID", "sub_knowledge"]].drop_duplicates()
    knowledge_sub = (
        titles[["knowledge", "sub_knowledge"]]
        .drop_duplicates()
        .sort_values(["knowledge", "sub_knowledge"])
    )
    return {
        "question_info": question_info,
        "title_knowledge": title_knowledge,
        "title_sub": title_sub,
        "knowledge_sub": knowledge_sub,
    }


def build_association_tables(titles: pd.DataFrame, maps: dict) -> dict:
    """生成知识点关联汇总、共享题目、多知识点题目等审计表。"""
    knowledge_summary = (
        maps["title_knowledge"]
        .groupby("knowledge", as_index=False)
        .agg(question_count=("title_ID", "nunique"))
        .sort_values("question_count", ascending=False)
    )
    sub_count = (
        maps["knowledge_sub"]
        .groupby("knowledge", as_index=False)
        .agg(sub_knowledge_count=("sub_knowledge", "nunique"))
    )
    knowledge_summary = knowledge_summary.merge(
        sub_count, on="knowledge", how="left"
    )
    knowledge_summary["association_rows"] = knowledge_summary["question_count"]
    knowledge_summary["share_of_questions"] = (
        knowledge_summary["question_count"] / titles["title_ID"].nunique()
    ).round(4)
    knowledge_summary["note"] = (
        "占比分母为唯一题目数 38；一题多知识点时各知识点占比合计大于 1"
    )
    knowledge_summary = knowledge_summary[
        ["knowledge", "sub_knowledge_count", "question_count",
         "association_rows", "share_of_questions", "note"]
    ]

    sub_summary = (
        titles[["knowledge", "sub_knowledge", "title_ID"]]
        .drop_duplicates()
        .groupby(["knowledge", "sub_knowledge"], as_index=False)
        .agg(question_count=("title_ID", "nunique"))
        .sort_values(["knowledge", "question_count"], ascending=[True, False])
    )

    # 知识点之间的共享题目
    overlap_rows = []
    knowledge_values = sorted(maps["title_knowledge"]["knowledge"].unique())
    for i in range(len(knowledge_values)):
        for j in range(i + 1, len(knowledge_values)):
            left = knowledge_values[i]
            right = knowledge_values[j]
            shared = set(
                maps["title_knowledge"].loc[
                    maps["title_knowledge"]["knowledge"].eq(left), "title_ID"
                ]
            ) & set(
                maps["title_knowledge"].loc[
                    maps["title_knowledge"]["knowledge"].eq(right), "title_ID"
                ]
            )
            if shared:
                overlap_rows.append(
                    {
                        "knowledge_a": left,
                        "knowledge_b": right,
                        "shared_question_count": len(shared),
                        "shared_titles": "|".join(sorted(shared)),
                    }
                )
    overlap = pd.DataFrame(
        overlap_rows,
        columns=["knowledge_a", "knowledge_b", "shared_question_count", "shared_titles"],
    )

    # 一题多知识点明细
    per_title = titles.groupby("title_ID", as_index=False).agg(
        full_score=("score_num", "first"),
        knowledge_count=("knowledge", "nunique"),
        knowledge_list=("knowledge", lambda s: "|".join(sorted(set(s)))),
        sub_knowledge_count=("sub_knowledge", "nunique"),
        sub_knowledge_list=("sub_knowledge", lambda s: "|".join(sorted(set(s)))),
    )
    multi = per_title[per_title["knowledge_count"].gt(1)].copy()

    # 同一主知识点下挂多个子知识点的题目
    same_knowledge_multi_sub = (
        titles.groupby(["title_ID", "knowledge"], as_index=False)
        .agg(sub_knowledge_count=("sub_knowledge", "nunique"))
    )
    same_knowledge_multi_sub = same_knowledge_multi_sub[
        same_knowledge_multi_sub["sub_knowledge_count"].gt(1)
    ]

    reconciliation = pd.DataFrame(
        [
            {"item": "唯一题目数（nunique title_ID）",
             "value": int(titles["title_ID"].nunique()),
             "note": "全库题目总量以此为准"},
            {"item": "主知识点数", "value": int(titles["knowledge"].nunique()),
             "note": ""},
            {"item": "子知识点数", "value": int(titles["sub_knowledge"].nunique()),
             "note": ""},
            {"item": "题目-主知识点关联行数", "value": len(maps["title_knowledge"]),
             "note": "一题多知识点会多计"},
            {"item": "题目-子知识点关联行数", "value": len(maps["title_sub"]),
             "note": "一题多子知识点会多计"},
            {"item": "主知识点题数合计（按知识点分别统计后再相加）",
             "value": int(knowledge_summary["question_count"].sum()),
             "note": "= 38 + 5 道多归属题目"},
            {"item": "子知识点题数合计（按子知识点分别统计后再相加）",
             "value": int(sub_summary["question_count"].sum()),
             "note": "= 38 + 6 道多归属题目"},
            {"item": "一题多主知识点题目数", "value": int(len(multi)),
             "note": ""},
            {"item": "一题多子知识点题目数",
             "value": int(per_title["sub_knowledge_count"].gt(1).sum()),
             "note": ""},
            {"item": "同一主知识点下多子知识点题目数",
             "value": int(len(same_knowledge_multi_sub)),
             "note": "汇总到主知识点时必须先去重"},
            {"item": "存在共享题目的知识点对", "value": int(len(overlap)),
             "note": ""},
        ]
    )

    return {
        "knowledge_summary": knowledge_summary,
        "sub_summary": sub_summary,
        "overlap": overlap,
        "multi_knowledge": multi,
        "same_knowledge_multi_sub": same_knowledge_multi_sub,
        "reconciliation": reconciliation,
    }


def write_filtering_notes(
    path: Path,
    counts: dict,
    duplicate_groups: int,
    cross_file_groups: int,
) -> None:
    """把过滤规则与数量写成可读说明，风格与 A 模块的 deduplication_notes 一致。"""
    text = (
        "知识点与掌握情况模块数据过滤与数量说明（阶段2）\n"
        "范围：data/raw/chinavis2024_education 下的 15 个班级提交文件。\n"
        "原始数据只读；本阶段只做检查与计数，清洗后的全量记录在阶段3输出。\n"
        "\n"
        "一、数量核对\n"
        f"15 个班级文件原始行数合计：{counts['raw_total']}\n"
        f"合并后原始记录数：{counts['raw_total']}\n"
        f"班级字段字面值为 class 的记录数：{counts['literal_class']}\n"
        f"班级字段修复后记录数：{counts['after_class_repair']}（只改字段值，行数不变）\n"
        f"九字段完全重复被排除记录数：{counts['duplicate_excluded']}"
        f"（重复组 {duplicate_groups} 组，跨文件重复组 {cross_file_groups} 组）\n"
        f"去重后记录数：{counts['after_dedup']}\n"
        f"无法匹配学生被排除记录数：{counts['unmatched_students']}\n"
        f"用于后续分析的记录数：{counts['analysis_records']}\n"
        f"核对：{counts['raw_total']} - {counts['duplicate_excluded']}"
        f" - {counts['unmatched_students']} = {counts['analysis_records']}\n"
        "\n"
        "二、过滤规则\n"
        "1. 班级字段：值等于字面值 class 时，按来源文件名修复为 ClassN。\n"
        "2. 重复记录：class、student_ID、title_ID、time、state、score、method、memory、"
        "timeconsume 九个业务字段完全相同，保留读取顺序第一条。\n"
        "   数据未提供唯一提交 ID，该口径是分析假设，不能证明被排除记录一定是系统重复；\n"
        "   被排除记录保存在 duplicate_records_excluded.csv。\n"
        "3. 学生匹配：提交记录中的 student_ID 无法在学生表找到时不猜测、不纠正，\n"
        "   记录保存在 unmatched_student_records.csv 并排除出分析集。\n"
        "4. 有效性：score 或 time 非数字、score 小于 0、score 大于该题满分的记录判为无效。\n"
        "   实测无效记录 0 条。\n"
        "5. 题目匹配：提交记录中的 title_ID 必须在题目表中存在，实测未知题号 0 条。\n"
        "6. 知识点：不删除任何题号；一题多知识点是合法关联，按两套独立映射分别统计。\n"
        "\n"
        "三、与 A 模块口径的关系\n"
        "上述规则与 A 模块已确定的公共口径一致；本模块独立读取原始数据复现，\n"
        "不读取 outputs/a_students 下的任何中间结果。\n"
    )
    path.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------------
# 提交记录质量检查与过滤数量
# ---------------------------------------------------------------------------
def repair_class_field(records: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """按来源文件名修复字面值为 class 的班级字段。"""
    expected_class = (
        records["source_file"]
        .str.replace("SubmitRecord-", "", regex=False)
        .str.replace(".csv", "", regex=False)
    )
    bad_class = records["class"].str.strip().eq("class").fillna(False)

    repairs = records.loc[bad_class].copy()
    repairs["original_class"] = repairs["class"]
    repairs["corrected_class"] = expected_class[bad_class]

    repaired = records.copy()
    repaired.loc[bad_class, "class"] = expected_class[bad_class]
    return repaired, repairs


def check_submission_table(
    records: pd.DataFrame,
    students: pd.DataFrame,
    titles: pd.DataFrame,
    question_info: pd.DataFrame,
) -> dict:
    """提交记录的字段、去重、关联、取值检查，并返回各阶段数量。"""
    checks = []
    raw_total = len(records)

    missing = int(records.isna().sum().sum())
    # 原始提交表 10 个字段（含 index），不含程序附加的 source_file
    record_field_count = records.shape[1] - 1
    checks.append(
        make_check(
            "R01", "提交记录", "字段缺失值",
            raw_total * record_field_count, missing, 0,
            "不处理（先检查）", raw_total * record_field_count,
            "10 个原始字段均无缺失（含 index）",
        )
    )

    checks.append(
        make_check(
            "R02", "提交记录", "来源文件数与总行数", len(records), 0, 0,
            "逐文件读取后合并", len(records),
            "15 个班级文件逐文件行数见 table_profile.csv",
        )
    )

    # 班级字段修复
    repaired, repairs = repair_class_field(records)
    checks.append(
        make_check(
            "R03", "提交记录", "班级字段字面值 class", raw_total, len(repairs),
            len(repairs), "按来源文件名修复为 ClassN（行数不变）", len(repaired),
            "修复明细见 class_field_repairs.csv",
        )
    )
    class_values = sorted(repaired["class"].dropna().unique())
    checks.append(
        make_check(
            "R04", "提交记录", "修复后班级取值", len(repaired),
            0 if len(class_values) == 15 else len(repaired), 0,
            "不处理（先检查）", len(repaired),
            f"修复后班级取值 {len(class_values)} 类，无字面值 class",
        )
    )

    # 九字段去重
    duplicate_mask = repaired.duplicated(subset=BUSINESS_COLUMNS, keep="first")
    excluded = repaired.loc[duplicate_mask].copy()
    deduplicated = repaired.loc[~duplicate_mask].copy()
    group_mask = repaired.duplicated(subset=BUSINESS_COLUMNS, keep=False)
    group_count = int(
        repaired.loc[group_mask, BUSINESS_COLUMNS].drop_duplicates().shape[0]
    )
    checks.append(
        make_check(
            "R05", "提交记录", "九字段完全重复记录", raw_total, len(excluded),
            len(excluded), "保留读取顺序第一条，被排除记录单独存档", len(deduplicated),
            f"重复组 {group_count} 组；明细见 duplicate_records_excluded.csv",
        )
    )

    # 重复组是否跨文件
    group_table = (
        repaired.loc[group_mask]
        .groupby(BUSINESS_COLUMNS, dropna=False)
        .agg(occurrences=("source_file", "size"),
             source_file_count=("source_file", "nunique"))
        .reset_index()
    )
    cross_file_groups = int(group_table["source_file_count"].gt(1).sum())
    checks.append(
        make_check(
            "R06", "提交记录", "重复组跨文件情况", group_count, cross_file_groups, 0,
            "不处理（用于判断重复性质）", group_count,
            "实测跨文件重复组 {} 组；为 0 说明重复均发生在同一文件内部".format(
                cross_file_groups
            ),
        )
    )

    # 无法匹配学生
    unknown_student_mask = ~deduplicated["student_ID"].isin(students["student_ID"])
    unmatched_students = deduplicated.loc[unknown_student_mask].copy()
    analysis_records = deduplicated.loc[~unknown_student_mask].copy()
    checks.append(
        make_check(
            "R07", "提交记录", "无法匹配学生表", len(deduplicated),
            len(unmatched_students), len(unmatched_students),
            "不猜测、不纠正，单独存档并排除", len(analysis_records),
            "明细见 unmatched_student_records.csv",
        )
    )

    # 无法匹配题目
    known_titles = set(titles["title_ID"])
    unknown_title_mask = ~deduplicated["title_ID"].isin(known_titles)
    unknown_titles = deduplicated.loc[unknown_title_mask]
    checks.append(
        make_check(
            "R08", "提交记录", "无法匹配题目表", len(deduplicated),
            len(unknown_titles), 0, "不处理（先检查）", len(deduplicated),
            f"提交记录中的不同题号 {deduplicated['title_ID'].nunique()} 个，均在题目表中",
        )
    )

    # 分数有效性
    full_lookup = question_info.set_index("title_ID")["full_score"]
    earned = pd.to_numeric(analysis_records["score"], errors="coerce")
    full_score = analysis_records["title_ID"].map(full_lookup)
    invalid_score = (
        full_score.isna()
        | full_score.le(0)
        | earned.isna()
        | earned.lt(0)
        | earned.gt(full_score)
    )
    checks.append(
        make_check(
            "R09", "提交记录", "得分越界或非数字", len(analysis_records),
            int(invalid_score.sum()), int(invalid_score.sum()),
            "标记为无效作答，不参与指标计算", int((~invalid_score).sum()),
            "实测无越界得分",
        )
    )

    # 时间有效性
    time_numeric = pd.to_numeric(analysis_records["time"], errors="coerce")
    invalid_time = int(time_numeric.isna().sum())
    checks.append(
        make_check(
            "R10", "提交记录", "时间非数字", len(analysis_records), invalid_time,
            invalid_time, "标记为无效作答", len(analysis_records) - invalid_time,
            "时间范围 {}~{}（Unix 秒）".format(
                int(time_numeric.min()), int(time_numeric.max())
            ),
        )
    )

    # 状态取值
    abnormal_state = ~repaired["state"].isin(STANDARD_STATES)
    state_anomalies = repaired.loc[abnormal_state].copy()
    checks.append(
        make_check(
            "R11", "提交记录", "state 非标准取值", len(repaired),
            int(abnormal_state.sum()), 0,
            "保留并审计；知识点与掌握情况模块按 score 评分，不使用 state", len(repaired),
            "源文件内即为 U+FFFD 替换字符，属数据损坏，见 state_anomaly_records.csv",
        )
    )

    # 状态与得分一致性（仅对标准状态）
    standard = repaired.loc[~abnormal_state].copy()
    standard["earned"] = pd.to_numeric(standard["score"], errors="coerce")
    standard["full"] = standard["title_ID"].map(full_lookup)
    correct_violation = int(
        (standard["state"].eq("Absolutely_Correct")
         & standard["earned"].ne(standard["full"])).sum()
    )
    partial_violation = int(
        (standard["state"].eq("Partially_Correct")
         & ~(standard["earned"].gt(0) & standard["earned"].lt(standard["full"]))).sum()
    )
    error_violation = int(
        (standard["state"].isin(
            ["Absolutely_Error"] + [f"Error{i}" for i in range(1, 10)]
        ) & standard["earned"].ne(0)).sum()
    )
    checks.append(
        make_check(
            "R12", "提交记录", "state 与 score 逻辑一致性", len(standard),
            correct_violation + partial_violation + error_violation,
            correct_violation + partial_violation + error_violation,
            "不一致记录标记并审计（实测为 0）", len(standard),
            "Absolutely_Correct=满分，Partially_Correct=部分分，Error*=0 分",
        )
    )

    # timeconsume 非数字
    consume_numeric = pd.to_numeric(repaired["timeconsume"], errors="coerce")
    checks.append(
        make_check(
            "R13", "提交记录", "timeconsume 非数字", len(repaired),
            int(consume_numeric.isna().sum()), 0,
            "知识点与掌握情况模块不使用该字段，仅记录", len(repaired),
            "取值形如 '--'，会影响时间类分析",
        )
    )

    # 跨班学生：修复前字面值 class 会让同一学生看似跨班，修复后应消失
    pre_multi_class = int(records.groupby("student_ID")["class"].nunique().gt(1).sum())
    post_multi_class = int(repaired.groupby("student_ID")["class"].nunique().gt(1).sum())
    checks.append(
        make_check(
            "R14", "提交记录", "同一学生出现在多个班级（修复后）", len(repaired),
            post_multi_class, 0,
            "保留，班级仅作参考维度", len(repaired),
            "修复前 {} 名，均由字面值 class 造成；修复后 {} 名".format(
                pre_multi_class, post_multi_class
            ),
        )
    )

    # 学生/题目覆盖
    students_without_records = set(students["student_ID"]) - set(repaired["student_ID"])
    checks.append(
        make_check(
            "R15", "提交记录", "学生表中无任何提交的学生",
            students["student_ID"].nunique(), len(students_without_records), 0,
            "不处理（先检查）", students["student_ID"].nunique(),
            "所有学生均有提交记录",
        )
    )
    unanswered = set(titles["title_ID"]) - set(repaired["title_ID"])
    checks.append(
        make_check(
            "R16", "提交记录", "题目表中从未被作答的题目",
            titles["title_ID"].nunique(), len(unanswered), 0,
            "不处理（先检查）", titles["title_ID"].nunique(),
            "38 道题均有作答记录",
        )
    )

    # 满分与观察到的最大得分
    observed_max = (
        pd.to_numeric(repaired["score"], errors="coerce")
        .groupby(repaired["title_ID"])
        .max()
        .reindex(question_info["title_ID"])
    )
    max_mismatch = int(observed_max.ne(question_info.set_index("title_ID")["full_score"]).sum())
    checks.append(
        make_check(
            "R17", "题目表", "题目满分与观察到的最大得分", len(question_info),
            max_mismatch, 0, "不处理（用于确认 score 即满分）", len(question_info),
            "38 道题全部一致，可确认题目表的 score 为满分",
        )
    )

    return {
        "checks": checks,
        "repairs": repairs,
        "excluded": excluded,
        "duplicate_groups": group_table,
        "unmatched_students": unmatched_students,
        "state_anomalies": state_anomalies,
        "deduplicated": deduplicated,
        "analysis_records": analysis_records,
        "counts": {
            "raw_total": raw_total,
            "literal_class": int(len(repairs)),
            "after_class_repair": len(repaired),
            "duplicate_excluded": int(len(excluded)),
            "after_dedup": len(deduplicated),
            "unmatched_students": int(len(unmatched_students)),
            "analysis_records": len(analysis_records),
            "duplicate_groups": group_count,
            "cross_file_groups": cross_file_groups,
        },
    }


# ---------------------------------------------------------------------------
# 数量对账
# ---------------------------------------------------------------------------
def build_count_reconciliation(counts: dict, profile: pd.DataFrame) -> pd.DataFrame:
    """按处理顺序记录每一步的数量，并给出算术核对结论。"""
    record_rows = int(
        profile.loc[profile["table"].eq("Data_SubmitRecord"), "rows"].sum()
    )
    rows = [
        {"step": 1, "item": "15 个班级文件原始行数合计",
         "count": record_rows,
         "note": "与逐文件行数之和一致"},
        {"step": 2, "item": "合并后原始记录数", "count": counts["raw_total"],
         "note": "应等于第 1 步"},
        {"step": 3, "item": "班级字段字面值 class 记录数",
         "count": counts["literal_class"], "note": "只改字段值，行数不变"},
        {"step": 4, "item": "班级字段修复后记录数",
         "count": counts["after_class_repair"], "note": "应等于第 2 步"},
        {"step": 5, "item": "九字段重复被排除记录数",
         "count": counts["duplicate_excluded"], "note": "保留每组第一条"},
        {"step": 6, "item": "去重后记录数", "count": counts["after_dedup"],
         "note": "= 第 4 步 − 第 5 步"},
        {"step": 7, "item": "无法匹配学生被排除记录数",
         "count": counts["unmatched_students"], "note": "单独存档，不自动纠正"},
        {"step": 8, "item": "用于后续分析的记录数",
         "count": counts["analysis_records"],
         "note": "= 第 6 步 − 第 7 步"},
    ]
    reconciliation = pd.DataFrame(rows)

    equation_ok = (
        counts["raw_total"] - counts["duplicate_excluded"] - counts["unmatched_students"]
        == counts["analysis_records"]
    )
    sum_ok = record_rows == counts["raw_total"]
    reconciliation = pd.concat(
        [
            reconciliation,
            pd.DataFrame(
                [
                    {"step": 9, "item": "算术核对：原始 − 重复 − 未匹配 = 分析记录",
                     "count": counts["analysis_records"],
                     "note": "通过" if equation_ok else "不通过，需核查"},
                    {"step": 10, "item": "算术核对：逐文件行数之和 = 合并后原始记录数",
                     "count": counts["raw_total"],
                     "note": "通过" if sum_ok else "不通过，需核查"},
                ]
            ),
        ],
        ignore_index=True,
    )
    if not (equation_ok and sum_ok):
        raise ValueError("提交记录数量对账不通过，请检查读取与过滤步骤")
    return reconciliation


# ---------------------------------------------------------------------------
# 阶段2：数据质量检查
# ---------------------------------------------------------------------------
def run_quality_checks(
    students: pd.DataFrame,
    titles: pd.DataFrame,
    records: pd.DataFrame,
    profile: pd.DataFrame,
    dirs: dict,
) -> dict:
    """执行阶段2的字段、关联、质量检查与数量对账，并写出审计文件。"""
    audit_dir = dirs["audit"]

    # 表级概况：学生表与题目表各一行，提交记录逐文件一行
    profile = pd.concat(
        [
            pd.DataFrame(
                [
                    {"table": "Data_StudentInfo", "file": STUDENT_FILE.name,
                     "rows": len(students), "columns": students.shape[1],
                     "size_bytes": STUDENT_FILE.stat().st_size},
                    {"table": "Data_TitleInfo", "file": TITLE_FILE.name,
                     "rows": len(titles), "columns": titles.shape[1] - 1,
                     "size_bytes": TITLE_FILE.stat().st_size},
                ]
            ),
            profile,
        ],
        ignore_index=True,
    )
    write_csv(profile, audit_dir / "table_profile.csv")

    # 字段说明
    field_dictionary = build_field_dictionary(students, titles, records)
    write_csv(field_dictionary, audit_dir / "field_dictionary.csv")

    # 学生表检查
    checks = [
        make_check(
            "D01", "全部数据表", "字段读取类型", 20, 0, 0,
            "统一以 dtype=string 读取", 20,
            "避免 student_ID、title_ID、知识点编码被推断为数值而丢失前导零或大小写",
        )
    ]
    checks.extend(check_student_table(students))

    # 题目表检查与映射
    title_checks, association = check_title_table(titles)
    checks.extend(title_checks)
    maps = build_knowledge_maps(titles)
    write_csv(maps["question_info"], audit_dir / "question_info.csv")
    write_csv(maps["title_knowledge"], audit_dir / "title_knowledge_map.csv")
    write_csv(maps["title_sub"], audit_dir / "title_sub_knowledge_map.csv")
    write_csv(maps["knowledge_sub"], audit_dir / "knowledge_sub_knowledge_map.csv")

    # 关联结构
    association_tables = build_association_tables(titles, maps)
    write_csv(
        association_tables["knowledge_summary"],
        audit_dir / "knowledge_association_summary.csv",
    )
    write_csv(
        association_tables["sub_summary"],
        audit_dir / "sub_knowledge_summary.csv",
    )
    write_csv(association_tables["overlap"], audit_dir / "knowledge_overlap.csv")
    write_csv(
        association_tables["multi_knowledge"],
        audit_dir / "multi_knowledge_questions.csv",
    )
    write_csv(
        association_tables["same_knowledge_multi_sub"],
        audit_dir / "same_knowledge_multi_sub_questions.csv",
    )
    write_csv(
        association_tables["reconciliation"],
        audit_dir / "association_reconciliation.csv",
    )

    # 提交记录检查
    submission = check_submission_table(
        records, students, titles, maps["question_info"]
    )
    checks.extend(submission["checks"])

    write_csv(submission["repairs"], audit_dir / "class_field_repairs.csv")
    write_csv(submission["excluded"], audit_dir / "duplicate_records_excluded.csv")
    write_csv(
        submission["duplicate_groups"],
        audit_dir / "duplicate_groups.csv",
    )
    write_csv(
        submission["unmatched_students"],
        audit_dir / "unmatched_student_records.csv",
    )
    write_csv(
        submission["state_anomalies"],
        audit_dir / "state_anomaly_records.csv",
    )

    # 汇总检查表
    checks_frame = pd.DataFrame(checks)
    write_csv(checks_frame, audit_dir / "data_quality_checks.csv")

    # 数量对账
    counts = submission["counts"]
    reconciliation = build_count_reconciliation(counts, profile)
    write_csv(reconciliation, audit_dir / "count_reconciliation.csv")
    write_filtering_notes(
        audit_dir / "filtering_notes.txt",
        counts,
        counts["duplicate_groups"],
        counts["cross_file_groups"],
    )

    # 控制台摘要
    print("\n—— 表级概况 ——")
    print(profile.to_string(index=False))
    print("\n—— 知识点关联结构 ——")
    print(association_tables["reconciliation"].to_string(index=False))
    print("\n—— 过滤前后数量 ——")
    print(reconciliation.to_string(index=False))
    print("\n—— 数据质量检查（异常数不为 0 的项）——")
    print(
        checks_frame.loc[checks_frame["abnormal_count"].gt(0)].to_string(index=False)
    )
    print("\n审计文件目录：", audit_dir)
    print(
        "逐文件行数合计：",
        int(profile.loc[profile["table"].eq("Data_SubmitRecord"), "rows"].sum()),
    )
    print("合并后原始记录：", counts["raw_total"])
    print("九字段去重排除：", counts["duplicate_excluded"])
    print("去重后记录：", counts["after_dedup"])
    print("未匹配学生排除：", counts["unmatched_students"])
    print("分析记录：", counts["analysis_records"])
    print(
        "核对：{} − {} − {} = {}".format(
            counts["raw_total"],
            counts["duplicate_excluded"],
            counts["unmatched_students"],
            counts["analysis_records"],
        )
    )

    return {
        "students": students,
        "titles": titles,
        "records": records,
        "profile": profile,
        "maps": maps,
        "submission": submission,
        "association": association_tables,
        "counts": counts,
        "checks": checks_frame,
        "dirs": dirs,
    }


# ---------------------------------------------------------------------------
# 阶段3：有效作答、最终作答与知识点指标
# ---------------------------------------------------------------------------
ATTEMPT_KEYS = ["student_ID", "title_ID"]
KNOWLEDGE = "knowledge"
SUB_KNOWLEDGE = "sub_knowledge"


def prepare_valid_records(records: pd.DataFrame, question_info: pd.DataFrame) -> dict:
    """补充满分与得分率，并划分有效、无效记录。"""
    full_lookup = question_info.set_index("title_ID")["full_score"]
    data = records.copy()
    data["full_score"] = pd.to_numeric(
        data["title_ID"].map(full_lookup), errors="coerce"
    )
    data["earned"] = pd.to_numeric(data["score"], errors="coerce")
    data["time_numeric"] = pd.to_numeric(data["time"], errors="coerce")
    invalid_mask = (
        data["full_score"].isna()
        | data["full_score"].le(0)
        | data["earned"].isna()
        | data["earned"].lt(0)
        | data["earned"].gt(data["full_score"])
        | data["time_numeric"].isna()
    )
    valid = data.loc[~invalid_mask].copy()
    invalid = data.loc[invalid_mask].copy()
    # 单题得分率 = 得分 / 该题满分
    valid["score_rate"] = valid["earned"] / valid["full_score"]
    return {"valid": valid, "invalid": invalid}


def select_extreme_attempts(valid: pd.DataFrame, method: str) -> dict:
    """按学生-题目组合取时间最大（或最小）的有效作答。

    同一组合在最晚（最早）时刻可能有多条记录：得分率相同则保留一条；
    得分率不同则无法判断先后，整个组合排除并单独审计。
    """
    extreme_time = valid.groupby(ATTEMPT_KEYS)["time_numeric"].transform(method)
    initial = valid.loc[valid["time_numeric"].eq(extreme_time)].copy()
    rows_at_extreme = len(initial)
    multi_row_pairs = int(initial.groupby(ATTEMPT_KEYS).size().gt(1).sum())
    multi_rows = initial.loc[
        initial.groupby(ATTEMPT_KEYS)["time_numeric"].transform("size").gt(1)
    ].sort_values(ATTEMPT_KEYS)

    tie = initial.groupby(ATTEMPT_KEYS)["score_rate"].nunique()
    ambiguous_keys = tie[tie.gt(1)].reset_index()[ATTEMPT_KEYS]
    ambiguous = initial.merge(ambiguous_keys, on=ATTEMPT_KEYS, how="inner")

    rows = initial.merge(
        ambiguous_keys.assign(_ambiguous=True), on=ATTEMPT_KEYS, how="left"
    )
    rows = rows.loc[rows["_ambiguous"].isna()].drop(columns="_ambiguous")
    attempts = (
        rows.sort_values(ATTEMPT_KEYS)
        .drop_duplicates(ATTEMPT_KEYS)
        .reset_index(drop=True)
    )
    return {
        "attempts": attempts,
        "ambiguous": ambiguous,
        "multi_rows": multi_rows,
        "rows_at_extreme": rows_at_extreme,
        "multi_row_pairs": multi_row_pairs,
        "extra_rows": rows_at_extreme - len(attempts),
    }


def student_level_rate(
    attempts: pd.DataFrame, mapping: pd.DataFrame, level: str
) -> pd.DataFrame:
    """学生-知识点得分率：该生在该知识点下各题得分率的等权平均。"""
    expanded = attempts.merge(mapping, on="title_ID", how="inner")
    return expanded.groupby(["student_ID", level], as_index=False).agg(
        attempted_questions=("title_ID", "nunique"),
        score_rate=("score_rate", "mean"),
    )


def build_class_lookup(records: pd.DataFrame) -> pd.DataFrame:
    """取每个学生的班级，仅作参考维度。"""
    return records.groupby("student_ID", as_index=False).agg(
        class_name=("class", "first")
    )


def build_coverage(
    records: pd.DataFrame,
    valid: pd.DataFrame,
    attempts: pd.DataFrame,
    mapping: pd.DataFrame,
    level: str,
    n_students: int,
    sub_counts: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """知识点覆盖：题量、参与学生、提交记录、有效作答、最终作答。"""
    question = mapping.groupby(level, as_index=False).agg(
        question_count=("title_ID", "nunique")
    )
    expanded_attempts = attempts.merge(mapping, on="title_ID", how="inner")
    attempt_stats = expanded_attempts.groupby(level, as_index=False).agg(
        student_count=("student_ID", "nunique"),
        final_attempt_count=("title_ID", "size"),
    )
    expanded_records = records.merge(mapping, on="title_ID", how="inner")
    record_stats = expanded_records.groupby(level, as_index=False).agg(
        submission_count=("title_ID", "size")
    )
    expanded_valid = valid.merge(mapping, on="title_ID", how="inner")
    valid_stats = expanded_valid.groupby(level, as_index=False).agg(
        valid_submission_count=("title_ID", "size")
    )

    coverage = (
        question.merge(attempt_stats, on=level, how="left")
        .merge(record_stats, on=level, how="left")
        .merge(valid_stats, on=level, how="left")
    )
    for column in (
        "student_count",
        "final_attempt_count",
        "submission_count",
        "valid_submission_count",
    ):
        coverage[column] = coverage[column].fillna(0).astype(int)
    coverage["coverage_rate"] = (coverage["student_count"] / n_students).round(6)
    coverage["no_data_students"] = n_students - coverage["student_count"]

    columns = [level]
    if sub_counts is not None:
        coverage = coverage.merge(sub_counts, on=level, how="left")
        columns.append("sub_knowledge_count")
    columns += [
        "question_count",
        "student_count",
        "coverage_rate",
        "no_data_students",
        "submission_count",
        "valid_submission_count",
        "final_attempt_count",
    ]
    coverage = coverage[columns].sort_values(
        ["question_count", level], ascending=[False, True]
    )
    return coverage.reset_index(drop=True)


def build_performance(
    long: pd.DataFrame, coverage: pd.DataFrame, level: str
) -> pd.DataFrame:
    """知识点整体表现：对有数据学生的得分率做等权平均。"""
    grouped = long.groupby(level)
    performance = pd.DataFrame(
        {
            "n_students": grouped["student_ID"].nunique(),
            "mean_score_rate": grouped["score_rate"].mean(),
            "median_score_rate": grouped["score_rate"].median(),
            "p25_score_rate": grouped["score_rate"].quantile(0.25),
            "p75_score_rate": grouped["score_rate"].quantile(0.75),
            "std_score_rate": grouped["score_rate"].std(),
            "min_score_rate": grouped["score_rate"].min(),
            "max_score_rate": grouped["score_rate"].max(),
            "mean_attempted_questions": grouped["attempted_questions"].mean(),
            "full_mark_students": grouped["score_rate"].apply(
                lambda values: int((values >= 1 - 1e-12).sum())
            ),
            "below_60_students": grouped["score_rate"].apply(
                lambda values: int((values < 0.6).sum())
            ),
        }
    ).reset_index()

    performance = coverage[[level]].merge(performance, on=level, how="left")
    # 得分率类指标保留完整精度，避免舍入影响对账与复算
    performance["mean_attempted_questions"] = performance[
        "mean_attempted_questions"
    ].round(2)
    performance = performance.merge(
        coverage[[level, "question_count", "no_data_students"]],
        on=level,
        how="left",
    )
    return performance[
        [
            level,
            "question_count",
            "n_students",
            "no_data_students",
            "mean_score_rate",
            "median_score_rate",
            "p25_score_rate",
            "p75_score_rate",
            "std_score_rate",
            "min_score_rate",
            "max_score_rate",
            "full_mark_students",
            "below_60_students",
            "mean_attempted_questions",
        ]
    ]


def build_matrix(
    long: pd.DataFrame,
    level: str,
    student_order: list,
    level_order: list,
) -> pd.DataFrame:
    """学生×知识点宽表；无数据组合保留为空值，不填 0。"""
    matrix = long.pivot(index="student_ID", columns=level, values="score_rate")
    matrix = matrix.reindex(index=student_order, columns=level_order)
    matrix.index.name = "student_ID"
    return matrix.reset_index()


def build_sensitivity(
    valid: pd.DataFrame,
    first_attempts: pd.DataFrame,
    latest_attempts: pd.DataFrame,
    maps: dict,
) -> dict:
    """三种口径的敏感性对照：首次、最终、提交记录加权。"""
    first_student = first_attempts.groupby("student_ID")["score_rate"].mean()
    latest_student = latest_attempts.groupby("student_ID")["score_rate"].mean()

    overall = pd.DataFrame(
        [
            {
                "scope": "overall",
                "first_attempt_student_mean": round(float(first_student.mean()), 6),
                "latest_attempt_student_mean": round(float(latest_student.mean()), 6),
                "submission_weighted_record_mean": round(
                    float(valid["score_rate"].mean()), 6
                ),
                "pair_weighted_latest_mean": round(
                    float(latest_attempts["score_rate"].mean()), 6
                ),
                "n_students_first": int(first_student.size),
                "n_students_latest": int(latest_student.size),
                "n_records": int(len(valid)),
                "n_pairs": int(len(latest_attempts)),
                "latest_minus_first": round(
                    float(latest_student.mean() - first_student.mean()), 6
                ),
            }
        ]
    )

    frames = {}
    for level, mapping in (
        (KNOWLEDGE, maps["title_knowledge"]),
        (SUB_KNOWLEDGE, maps["title_sub"]),
    ):
        first_long = student_level_rate(first_attempts, mapping, level)
        latest_long = student_level_rate(latest_attempts, mapping, level)
        first_mean = first_long.groupby(level)["score_rate"].mean()
        latest_mean = latest_long.groupby(level)["score_rate"].mean()
        expanded_valid = valid.merge(mapping, on="title_ID", how="inner")
        submission_mean = expanded_valid.groupby(level)["score_rate"].mean()
        n_first = first_long.groupby(level)["student_ID"].nunique()
        n_latest = latest_long.groupby(level)["student_ID"].nunique()
        n_records = expanded_valid.groupby(level).size()

        frame = pd.DataFrame(
            {
                level: first_mean.index,
                "n_students_first": n_first.reindex(first_mean.index)
                .astype(int)
                .values,
                "n_students_latest": n_latest.reindex(first_mean.index)
                .astype(int)
                .values,
                "n_records": n_records.reindex(first_mean.index).astype(int).values,
                "first_attempt": first_mean.round(6).values,
                "latest_attempt": latest_mean.reindex(first_mean.index)
                .round(6)
                .values,
                "submission_weighted": submission_mean.reindex(first_mean.index)
                .round(6)
                .values,
            }
        )
        frame["latest_minus_first"] = (
            frame["latest_attempt"] - frame["first_attempt"]
        ).round(6)
        frame["submission_minus_latest"] = (
            frame["submission_weighted"] - frame["latest_attempt"]
        ).round(6)
        frames[level] = frame

    return {"overall": overall, **frames}


def traceability_diff(
    attempts: pd.DataFrame,
    mapping: pd.DataFrame,
    level: str,
    long: pd.DataFrame,
) -> tuple:
    """从最终作答重新计算学生-知识点得分率，与已保存的长表比对。"""
    expanded = attempts.merge(mapping, on="title_ID", how="inner")
    recomputed = expanded.groupby(["student_ID", level])["score_rate"].mean()
    saved = long.set_index(["student_ID", level])["score_rate"]
    joined = pd.concat(
        [recomputed.rename("recomputed"), saved.rename("saved")],
        axis=1,
        join="outer",
    )
    difference = (joined["recomputed"] - joined["saved"]).abs().max()
    missing = int(joined.isna().any(axis=1).sum())
    return float(difference), missing


def build_phase3_reconciliation(
    latest_info: dict,
    long_k: pd.DataFrame,
    long_s: pd.DataFrame,
    maps: dict,
    coverage_k: pd.DataFrame,
    coverage_s: pd.DataFrame,
    performance_k: pd.DataFrame,
    performance_s: pd.DataFrame,
    matrix_k: pd.DataFrame,
    matrix_s: pd.DataFrame,
    n_students: int,
) -> pd.DataFrame:
    """阶段3数据对账：数量、唯一性、范围、无数据与可追溯性。"""
    attempts = latest_info["attempts"]
    rows = []

    def add(check_id, item, expected, actual, passed, note=""):
        rows.append(
            {
                "check_id": check_id,
                "item": item,
                "expected": str(expected),
                "actual": str(actual),
                "result": "通过" if passed else "不通过",
                "note": note,
            }
        )

    def add_info(check_id, item, expected, actual, note=""):
        rows.append(
            {
                "check_id": check_id,
                "item": item,
                "expected": str(expected),
                "actual": str(actual),
                "result": "信息",
                "note": note,
            }
        )

    pair_rows = len(attempts)
    unique_pairs = attempts[ATTEMPT_KEYS].drop_duplicates().shape[0]
    add(
        "P01",
        "最终作答行数 = 学生-题目唯一组合数",
        pair_rows,
        unique_pairs,
        pair_rows == unique_pairs,
        "每个组合只保留一条最终有效作答",
    )

    duplicate_pairs = int(attempts.duplicated(ATTEMPT_KEYS).sum())
    add("P02", "最终作答在学生-题目组合上重复", 0, duplicate_pairs, duplicate_pairs == 0)

    ambiguous_count = len(latest_info["ambiguous"])
    add(
        "P03",
        "最晚时刻得分率冲突而被排除的组合数",
        0,
        ambiguous_count,
        ambiguous_count == 0,
        "若不为 0，见 ambiguous_latest_pairs.csv",
    )

    add_info(
        "P04",
        "最晚时刻同组合多条记录（得分率相同，保留一条）",
        "—",
        "{} 个组合，多出 {} 行".format(
            latest_info["multi_row_pairs"], latest_info["extra_rows"]
        ),
        "多行明细见 latest_multiple_same_time.csv",
    )

    duplicate_long_k = int(long_k.duplicated(["student_ID", KNOWLEDGE]).sum())
    add("P05", "学生-主知识点组合重复", 0, duplicate_long_k, duplicate_long_k == 0)
    duplicate_long_s = int(long_s.duplicated(["student_ID", SUB_KNOWLEDGE]).sum())
    add("P06", "学生-子知识点组合重复", 0, duplicate_long_s, duplicate_long_s == 0)

    knowledge_rows = len(maps["title_knowledge"])
    add("P07", "主知识点关联行数与阶段2一致", 43, knowledge_rows, knowledge_rows == 43)
    sub_rows = len(maps["title_sub"])
    add("P08", "子知识点关联行数与阶段2一致", 44, sub_rows, sub_rows == 44)

    unique_titles = maps["question_info"]["title_ID"].nunique()
    add("P09", "全库唯一题目数", 38, unique_titles, unique_titles == 38)

    rate_min = float(
        min(
            attempts["score_rate"].min(),
            long_k["score_rate"].min(),
            long_s["score_rate"].min(),
        )
    )
    rate_max = float(
        max(
            attempts["score_rate"].max(),
            long_k["score_rate"].max(),
            long_s["score_rate"].max(),
        )
    )
    nan_in_long = int(
        long_k["score_rate"].isna().sum() + long_s["score_rate"].isna().sum()
    )
    add(
        "P10",
        "得分率范围与空值",
        "[0, 1] 且长表无空值",
        "min={:.6f}, max={:.6f}, 长表空值={}".format(rate_min, rate_max, nan_in_long),
        rate_min >= 0 and rate_max <= 1 and nan_in_long == 0,
    )

    missing_k = n_students * len(coverage_k) - len(long_k)
    missing_s = n_students * len(coverage_s) - len(long_s)
    nan_k = int(matrix_k.drop(columns="student_ID").isna().sum().sum())
    nan_s = int(matrix_s.drop(columns="student_ID").isna().sum().sum())
    zero_k = int((long_k["score_rate"] == 0).sum())
    zero_s = int((long_s["score_rate"] == 0).sum())
    add(
        "P11",
        "无数据与 0 分区分（无数据保留空值，不填 0）",
        "宽表空值数 = 无数据组合数",
        "主：无数据 {} / 空值 {} / 0 分 {}；子：无数据 {} / 空值 {} / 0 分 {}".format(
            missing_k, nan_k, zero_k, missing_s, nan_s, zero_s
        ),
        missing_k == nan_k and missing_s == nan_s,
        "0 分是真实作答结果，空值表示无有效作答",
    )

    diff_k, mismatch_k = traceability_diff(
        attempts, maps["title_knowledge"], KNOWLEDGE, long_k
    )
    diff_s, mismatch_s = traceability_diff(
        attempts, maps["title_sub"], SUB_KNOWLEDGE, long_s
    )
    add(
        "P12",
        "学生-知识点得分率可由最终作答完整复算",
        "最大绝对差 < 1e-9，且无缺失匹配",
        "主：最大差 {:.2e}，未匹配 {}；子：最大差 {:.2e}，未匹配 {}".format(
            diff_k, mismatch_k, diff_s, mismatch_s
        ),
        diff_k < 1e-9 and diff_s < 1e-9 and mismatch_k == 0 and mismatch_s == 0,
    )

    sum_k = int(coverage_k["question_count"].sum())
    sum_s = int(coverage_s["question_count"].sum())
    add(
        "P13",
        "按知识点统计的题量合计",
        "主 43 / 子 44",
        "主 {} / 子 {}".format(sum_k, sum_s),
        sum_k == 43 and sum_s == 44,
        "一题多知识点会重复计入各知识点",
    )

    add_info(
        "P14",
        "最终作答组合总数",
        "50,482（阶段1探查值，供对照）",
        pair_rows,
        "以实际运行结果为准",
    )

    recomputed_k = long_k.groupby(KNOWLEDGE)["score_rate"].mean()
    recomputed_s = long_s.groupby(SUB_KNOWLEDGE)["score_rate"].mean()
    joined_k = (
        performance_k.set_index(KNOWLEDGE)["mean_score_rate"]
        .to_frame("perf")
        .join(recomputed_k)
    )
    joined_s = (
        performance_s.set_index(SUB_KNOWLEDGE)["mean_score_rate"]
        .to_frame("perf")
        .join(recomputed_s)
    )
    perf_diff = float(
        max(
            (joined_k["perf"] - joined_k["score_rate"]).abs().max(),
            (joined_s["perf"] - joined_s["score_rate"]).abs().max(),
        )
    )
    add(
        "P15",
        "知识点整体表现与长表学生等权均值一致",
        "< 1e-9",
        "{:.2e}".format(perf_diff),
        perf_diff < 1e-9,
    )

    add_info(
        "P16",
        "知识点展开后的合计与全库总数（展开会重复计入一题多知识点）",
        "全库：提交记录 231,811，学生-题目组合 50,482",
        "主：提交记录 {}，最终作答 {}；子：提交记录 {}，最终作答 {}".format(
            int(coverage_k["submission_count"].sum()),
            int(coverage_k["final_attempt_count"].sum()),
            int(coverage_s["submission_count"].sum()),
            int(coverage_s["final_attempt_count"].sum()),
        ),
        "全库总量始终使用去重记录数与唯一组合数，不使用展开合计",
    )

    return pd.DataFrame(rows)


def write_metric_definitions(path: Path, stats: dict) -> None:
    """写出指标定义、公式、分子分母与统计范围。"""
    sensitivity = stats["sensitivity"]
    text = f"""# 知识点与掌握情况模块指标定义与计算口径

本文件由 `scripts/d_knowledge.py` 在阶段3自动生成，文中数值来自本次运行。
所有指标均由 `data/raw/chinavis2024_education/` 的原始数据重新计算得到。

## 1. 数据范围

- 原始提交记录：{stats['raw_records']:,} 条
- 九字段去重后：{stats['dedup_records']:,} 条
- 剔除无法匹配学生记录后：{stats['analysis_records']:,} 条
- 有效记录：{stats['valid_records']:,} 条（无效 {stats['invalid_records']} 条）
- 学生：{stats['students']:,} 名
- 唯一题目：{stats['questions']} 道
- 主知识点：{stats['knowledge']} 个；子知识点：{stats['sub_knowledge']} 个
- 最终有效作答（学生-题目组合）：{stats['final_pairs']:,} 个

## 2. 记录级处理

### 2.1 去重

九个业务字段（class、student_ID、title_ID、time、state、score、method、memory、timeconsume）
完全相同即视为重复，保留读取顺序第一条。该口径与 A 模块一致。
数据没有唯一提交 ID，因此该口径是分析假设。

### 2.2 有效记录

同时满足以下条件才算有效：

- score、time 可转为数值；
- 0 ≤ score ≤ 该题满分；
- 满分取自 Data_TitleInfo.score（已核实等于该题观察到的最大得分）。

### 2.3 最终有效作答（主口径）

按 (student_ID, title_ID) 分组，取 time 最大的有效记录。
若最晚时刻有多条记录且得分率不同，无法判断先后，整个组合排除并审计。
同一时刻多条但得分率相同的，保留一条。

### 2.4 首次有效作答（仅用于敏感性分析）

按 (student_ID, title_ID) 分组，取 time 最小的有效记录，冲突处理规则同上。
首次作答不替代主口径。

## 3. 单题得分率

r(s, q) = score(s, q) / full_score(q)

- 分子：学生 s 在题目 q 上最终有效作答的得分。
- 分母：题目 q 的满分。
- 统计范围：每个学生-题目组合只取一条最终有效作答。
- 取值范围：0 到 1。

## 4. 学生-知识点得分率

R(s, k) = (1 / n(s, k)) × Σ r(s, q)，q ∈ Q(s, k)

- 分子：该生在该知识点下有有效最终作答的各题得分率之和。
- 分母 n(s, k)：该生在该知识点下**有有效最终作答的题目数**，不是该知识点的总题数。
- 统计范围：只对至少有一道有效作答的学生-知识点组合计算。
- 无作答的组合不产生记录，在宽表中保留为空值，**不记为 0**。
- 一题多知识点：该题得分率同时计入其关联的每个知识点。由于分子、分母按同一比例拆分，
  得分率数值不变（(score/2)/(full/2) 与 score/full 相等），受影响的只是覆盖题量和作答次数的计数。

## 5. 知识点整体表现

Mean(k) = (1 / |S(k)|) × Σ R(s, k)，s ∈ S(k)

- 分子：该知识点下所有有数据学生的 R(s, k) 之和。
- 分母 |S(k)|：该知识点下**有有效得分率的学生数**。
- 权重：学生等权。提交次数多或作答题目多的学生不会获得更大权重。
- 无数据学生不参与平均。

## 6. 覆盖指标

- 覆盖题量 question_count = nunique(title_ID)，按知识点分别统计；
  合计为 43（主）/ 44（子），大于唯一题目 38，因为一题多知识点。
- 参与学生数 student_count = 该知识点下有有效最终作答的 nunique(student_ID)。
- 覆盖率 coverage_rate = student_count / {stats['students']}。
- 关联提交记录数 submission_count：去重后提交记录中，题目属于该知识点的记录数；
  一题多知识点会重复计入各知识点，**不能跨知识点相加当作全库记录总数**。
- 有效提交记录数 valid_submission_count：同上，但只统计有效记录。
- 最终作答次数 final_attempt_count：该知识点下的学生-题目组合数。
- 全库唯一题目数始终为 nunique(title_ID) = 38，不能用关联行数代替。

## 7. 0 分与无数据

- 0 分：学生提交了该知识点下的题目，最终有效作答得分为 0，计入平均。
- 无数据：学生在该知识点下没有任何有效作答，保留为空值，不参与平均。
- 两者在统计表与宽表中严格区分。

## 8. 敏感性口径与单位

| 口径 | 单位 | 总体结果 |
|---|---|---:|
| 首次有效作答，学生等权 | 学生 | {sensitivity['first_attempt_student_mean']:.4f} |
| 最终有效作答，学生等权（主口径） | 学生 | {sensitivity['latest_attempt_student_mean']:.4f} |
| 全部有效提交记录加权 | 提交记录 | {sensitivity['submission_weighted_record_mean']:.4f} |
| 最终作答组合等权（参考） | 学生-题目组合 | {sensitivity['pair_weighted_latest_mean']:.4f} |

三种主要口径的统计单位不同：前两者先算学生值再平均，第三者按记录加权，
不能直接混用，也不能把提交次数当作学生数。

## 9. 解释边界

- 最终作答得分率反映的是"多次尝试后的收尾状态"，会受到重试行为影响；
  它不等于真实知识掌握程度，也不能区分"一次做对"和"反复试错后做对"。
- 首次作答与最终作答的差异可用于观察重试的影响，但首次作答同样不是掌握程度的直接测量。
- 各知识点题量差异较大（{stats['question_min']}–{stats['question_max']} 题），
  题量少的知识点其均值更容易受个别题目影响。
"""
    path.write_text(text, encoding="utf-8")


def run_metrics(ctx: dict) -> None:
    """阶段3：从清洗后的记录计算最终作答、覆盖、表现与敏感性。"""
    dirs = ctx["dirs"]
    tables_dir = dirs["tables"]
    audit_dir = dirs["audit"]
    intermediate_dir = dirs["intermediate"]

    students = ctx["students"]
    titles = ctx["titles"]
    records = ctx["submission"]["analysis_records"]
    maps = ctx["maps"]
    counts = ctx["counts"]
    n_students = int(students["student_ID"].nunique())

    # 1. 有效记录
    valid_data = prepare_valid_records(records, maps["question_info"])
    valid = valid_data["valid"]
    invalid = valid_data["invalid"]
    write_csv(invalid, audit_dir / "invalid_attempts.csv")

    # 2. 最终作答与首次作答
    latest_info = select_extreme_attempts(valid, "max")
    first_info = select_extreme_attempts(valid, "min")
    latest_attempts = latest_info["attempts"]
    first_attempts = first_info["attempts"]
    write_csv(latest_attempts, intermediate_dir / "attempts_latest.csv")
    write_csv(first_attempts, intermediate_dir / "attempts_first.csv")
    write_csv(latest_info["ambiguous"], audit_dir / "ambiguous_latest_pairs.csv")
    write_csv(
        latest_info["multi_rows"], audit_dir / "latest_multiple_same_time.csv"
    )

    # 3. 清洗后的记录（班级已修复、重复已排除、未匹配学生已剔除）
    write_csv(records, intermediate_dir / "records_clean.csv")

    # 4. 学生-知识点长表
    class_lookup = build_class_lookup(records)
    long_k = student_level_rate(latest_attempts, maps["title_knowledge"], KNOWLEDGE)
    long_k = long_k.merge(class_lookup, on="student_ID", how="left")[
        ["student_ID", "class_name", KNOWLEDGE, "attempted_questions", "score_rate"]
    ]
    long_s = student_level_rate(latest_attempts, maps["title_sub"], SUB_KNOWLEDGE)
    long_s = long_s.merge(class_lookup, on="student_ID", how="left")[
        ["student_ID", "class_name", SUB_KNOWLEDGE, "attempted_questions", "score_rate"]
    ]
    write_csv(long_k, tables_dir / "student_knowledge_score_rate.csv")
    write_csv(long_s, tables_dir / "student_sub_knowledge_score_rate.csv")

    # 5. 覆盖与表现
    sub_counts = maps["knowledge_sub"].groupby("knowledge", as_index=False).agg(
        sub_knowledge_count=("sub_knowledge", "nunique")
    )
    coverage_k = build_coverage(
        records,
        valid,
        latest_attempts,
        maps["title_knowledge"],
        KNOWLEDGE,
        n_students,
        sub_counts,
    )
    coverage_s = build_coverage(
        records,
        valid,
        latest_attempts,
        maps["title_sub"],
        SUB_KNOWLEDGE,
        n_students,
    )
    performance_k = build_performance(long_k, coverage_k, KNOWLEDGE)
    performance_s = build_performance(long_s, coverage_s, SUB_KNOWLEDGE)
    write_csv(coverage_k, tables_dir / "knowledge_coverage.csv")
    write_csv(performance_k, tables_dir / "knowledge_performance.csv")
    write_csv(coverage_s, tables_dir / "sub_knowledge_coverage.csv")
    write_csv(performance_s, tables_dir / "sub_knowledge_performance.csv")

    # 6. 题目-知识点映射
    question_map = (
        titles[["title_ID", "knowledge", "sub_knowledge"]]
        .merge(maps["question_info"], on="title_ID", how="left")
        .sort_values(["knowledge", "sub_knowledge", "title_ID"])
        .reset_index(drop=True)
    )
    write_csv(question_map, tables_dir / "knowledge_question_map.csv")

    # 7. 学生×知识点宽表
    student_order = sorted(students["student_ID"].tolist())
    matrix_k = build_matrix(
        long_k, KNOWLEDGE, student_order, coverage_k[KNOWLEDGE].tolist()
    )
    matrix_s = build_matrix(
        long_s, SUB_KNOWLEDGE, student_order, coverage_s[SUB_KNOWLEDGE].tolist()
    )
    write_csv(matrix_k, intermediate_dir / "student_knowledge_matrix.csv")
    write_csv(matrix_s, intermediate_dir / "student_sub_knowledge_matrix.csv")

    # 8. 敏感性
    sensitivity = build_sensitivity(valid, first_attempts, latest_attempts, maps)
    write_csv(sensitivity["overall"], tables_dir / "sensitivity_overall.csv")
    write_csv(sensitivity[KNOWLEDGE], tables_dir / "sensitivity_knowledge.csv")
    write_csv(
        sensitivity[SUB_KNOWLEDGE], tables_dir / "sensitivity_sub_knowledge.csv"
    )

    # 9. 数据对账
    reconciliation = build_phase3_reconciliation(
        latest_info,
        long_k,
        long_s,
        maps,
        coverage_k,
        coverage_s,
        performance_k,
        performance_s,
        matrix_k,
        matrix_s,
        n_students,
    )
    write_csv(reconciliation, audit_dir / "phase3_reconciliation.csv")
    if (reconciliation["result"] == "不通过").any():
        failed = reconciliation.loc[
            reconciliation["result"].eq("不通过"), "check_id"
        ].tolist()
        raise ValueError(f"阶段3数据对账不通过：{failed}")

    # 10. 指标定义
    stats = {
        "raw_records": counts["raw_total"],
        "dedup_records": counts["after_dedup"],
        "analysis_records": counts["analysis_records"],
        "valid_records": len(valid),
        "invalid_records": len(invalid),
        "students": n_students,
        "questions": int(maps["question_info"]["title_ID"].nunique()),
        "knowledge": int(maps["title_knowledge"]["knowledge"].nunique()),
        "sub_knowledge": int(maps["title_sub"]["sub_knowledge"].nunique()),
        "final_pairs": len(latest_attempts),
        "question_min": int(coverage_k["question_count"].min()),
        "question_max": int(coverage_k["question_count"].max()),
        "sensitivity": sensitivity["overall"].iloc[0].to_dict(),
    }
    write_metric_definitions(audit_dir / "metric_definitions.md", stats)

    # 11. 控制台摘要
    print("\n" + "=" * 70)
    print("阶段3：有效作答与知识点指标")
    print("=" * 70)
    print("有效记录：", len(valid), " 无效记录：", len(invalid))
    print(
        "最终作答行数：",
        len(latest_attempts),
        " 学生-题目唯一组合：",
        latest_attempts[ATTEMPT_KEYS].drop_duplicates().shape[0],
    )
    print(
        "最晚时刻多行组合：",
        latest_info["multi_row_pairs"],
        " 多出记录：",
        latest_info["extra_rows"],
        " 得分率冲突排除：",
        len(latest_info["ambiguous"]),
    )
    print("首次作答组合：", len(first_attempts))
    print("\n—— 主知识点覆盖 ——")
    print(coverage_k.to_string(index=False))
    print("\n—— 主知识点表现 ——")
    print(performance_k.to_string(index=False))
    print("\n—— 子知识点覆盖 ——")
    print(coverage_s.to_string(index=False))
    print("\n—— 子知识点表现 ——")
    print(performance_s.to_string(index=False))
    print("\n—— 敏感性（总体）——")
    print(sensitivity["overall"].to_string(index=False))
    print("\n—— 数据对账 ——")
    print(reconciliation.to_string(index=False))
    print("\n统计表目录：", tables_dir)


# ---------------------------------------------------------------------------
# 阶段4：可视化图表与流程图
# ---------------------------------------------------------------------------
def setup_plot_style() -> None:
    """统一中文字体与负号显示，沿用 A 模块设置。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
    plt.rcParams["axes.unicode_minus"] = False


def read_figure_table(tables_dir: Path, name: str, **kwargs) -> pd.DataFrame:
    """读取阶段3统计表，缺失时给出明确错误。"""
    path = tables_dir / name
    if not path.exists():
        raise FileNotFoundError(f"缺少阶段3统计表，无法绘图：{path}")
    return pd.read_csv(path, **kwargs)


def plot_knowledge_coverage(tables_dir: Path, figures_dir: Path) -> Path:
    """图1：主知识点覆盖题量横向柱状图（同时标注参与学生数）。"""
    coverage = read_figure_table(
        tables_dir, "knowledge_coverage.csv", dtype={"knowledge": "string"}
    )
    coverage = coverage.sort_values(
        ["question_count", "knowledge"], ascending=[False, True]
    ).reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(10, 6))
    single_question = coverage["question_count"].eq(1)
    colors = np.where(single_question, "#EE6666", "#5470C6")

    bars = ax.barh(
        coverage["knowledge"],
        coverage["question_count"],
        color=colors,
        height=0.62,
        zorder=3,
    )
    ax.invert_yaxis()
    labels = [
        f"{int(question)} 道｜{int(student):,} 人"
        for question, student in zip(
            coverage["question_count"], coverage["student_count"]
        )
    ]
    ax.bar_label(bars, labels=labels, padding=5, fontsize=9, color="#333333")

    max_question = int(coverage["question_count"].max())
    ax.set_xlim(0, max_question * 1.62)
    ax.set_xticks(np.arange(0, max_question + 2, 2))
    ax.set_xlabel("题目数量（道）", fontsize=11)
    ax.set_ylabel("主知识点编码", fontsize=11)
    ax.set_title("各主知识点覆盖题量", fontsize=15, pad=12)
    ax.tick_params(axis="both", labelsize=10)
    ax.grid(axis="x", linestyle="--", alpha=0.25, zorder=0)
    ax.set_axisbelow(True)
    ax.spines[["top", "right"]].set_visible(False)

    legend_handles = [
        Patch(facecolor="#5470C6", label="覆盖题量（道）"),
        Patch(facecolor="#EE6666", label="仅 1 道题的知识点"),
    ]
    ax.legend(
        handles=legend_handles,
        loc="lower right",
        frameon=False,
        fontsize=9,
    )
    fig.text(
        0.5,
        0.015,
        "注：题量按知识点分别统计，一题多知识点会重复计入；全库唯一题目共 38 道。"
        "柱末标注为「题量｜参与学生数」。",
        ha="center",
        fontsize=8.5,
        color="#555555",
    )
    fig.tight_layout(rect=(0, 0.055, 1, 1))

    path = figures_dir / "知识点覆盖题量.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def plot_knowledge_performance(tables_dir: Path, figures_dir: Path) -> Path:
    """图2：主知识点平均得分率柱状图（含样本量与参考线）。"""
    performance = read_figure_table(
        tables_dir, "knowledge_performance.csv", dtype={"knowledge": "string"}
    )
    performance = performance.sort_values(
        "mean_score_rate", ascending=False
    ).reset_index(drop=True)
    overall = read_figure_table(tables_dir, "sensitivity_overall.csv")
    overall_rate = float(overall.loc[0, "latest_attempt_student_mean"]) * 100

    rates = performance["mean_score_rate"].to_numpy(dtype=float) * 100
    single_question = performance["question_count"].eq(1)
    colors = np.where(single_question, "#EE6666", "#5470C6")
    labels = [
        f"{knowledge}*" if single else str(knowledge)
        for knowledge, single in zip(
            performance["knowledge"], single_question
        )
    ]

    fig, ax = plt.subplots(figsize=(11, 6))
    bars = ax.bar(labels, rates, color=colors, width=0.62, zorder=3)
    for bar, rate, students in zip(
        bars, rates, performance["n_students"]
    ):
        center = bar.get_x() + bar.get_width() / 2
        ax.text(
            center,
            rate + 1.2,
            f"{rate:.2f}%",
            ha="center",
            va="bottom",
            fontsize=9,
            color="#333333",
        )
        ax.text(
            center,
            2.0,
            f"n={int(students):,}",
            ha="center",
            va="bottom",
            fontsize=8,
            color="white",
        )

    ax.axhline(
        overall_rate,
        color="#B45309",
        linestyle="--",
        linewidth=1.3,
        zorder=4,
        label=f"全体学生等权均值 {overall_rate:.2f}%",
    )
    ax.set_ylim(0, 100)
    ax.set_xlabel("主知识点编码", fontsize=11)
    ax.set_ylabel("平均得分率（%）", fontsize=11)
    ax.set_title("各主知识点平均得分率（最终有效作答）", fontsize=15, pad=12)
    ax.tick_params(axis="both", labelsize=10)
    ax.grid(axis="y", linestyle="--", alpha=0.25, zorder=0)
    ax.set_axisbelow(True)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(
        handles=[
            Patch(facecolor="#5470C6", label="平均得分率"),
            Patch(facecolor="#EE6666", label="* 仅 1 道题（稳定性较弱）"),
            Line2D(
                [0],
                [0],
                color="#B45309",
                linestyle="--",
                linewidth=1.3,
                label=f"全体学生等权均值 {overall_rate:.2f}%",
            ),
        ],
        loc="upper left",
        frameon=False,
        fontsize=8.5,
    )
    fig.text(
        0.5,
        0.015,
        "注：* k4W1c、s8Y2f 各只有 1 道题，均值即该题的平均得分率，稳定性较弱。"
        "参考线为全体学生等权均值，不是各知识点均值的算术平均；"
        "得分率不等同于真实知识掌握程度。",
        ha="center",
        fontsize=8.5,
        color="#555555",
    )
    fig.tight_layout(rect=(0, 0.06, 1, 1))

    path = figures_dir / "知识点得分率.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def plot_student_knowledge_heatmap(
    tables_dir: Path, intermediate_dir: Path, figures_dir: Path
) -> Path:
    """图3：全部学生×主知识点得分率热力图；无数据灰色，0 分保留在色阶中。"""
    matrix = pd.read_csv(
        intermediate_dir / "student_knowledge_matrix.csv",
        dtype={"student_ID": "string"},
    )
    if "knowledge" in matrix.columns:
        knowledge_columns = [c for c in matrix.columns if c != "student_ID"]
    else:
        knowledge_columns = [c for c in matrix.columns if c != "student_ID"]

    values = matrix[knowledge_columns].to_numpy(dtype=float)
    row_mean = np.nanmean(values, axis=1)
    order = np.argsort(-row_mean, kind="stable")
    values = values[order]
    masked = np.ma.masked_invalid(values)

    cmap = plt.get_cmap("RdYlGn").copy()
    cmap.set_bad("#D9D9D9")

    fig, ax = plt.subplots(figsize=(11, 15))
    image = ax.imshow(
        masked,
        aspect="auto",
        cmap=cmap,
        vmin=0,
        vmax=1,
        interpolation="nearest",
    )
    ax.set_xticks(range(len(knowledge_columns)))
    ax.set_xticklabels(knowledge_columns, rotation=45, ha="right", fontsize=10)
    ax.set_yticks([])
    ax.set_xlabel("主知识点编码", fontsize=11)
    ax.set_ylabel(
        f"学生（按有数据知识点的等权平均降序，共 {len(matrix):,} 人）",
        fontsize=11,
    )
    ax.set_title(
        f"学生×主知识点得分率热力图（全量 {len(matrix):,} 名学生）",
        fontsize=16,
        pad=14,
    )
    ax.set_xticks(np.arange(-0.5, len(knowledge_columns), 1), minor=True)
    ax.grid(which="minor", axis="x", color="white", linewidth=0.8)
    ax.tick_params(which="minor", bottom=False)

    colorbar = fig.colorbar(image, ax=ax, fraction=0.035, pad=0.02)
    colorbar.set_label("学生-知识点得分率", fontsize=10)
    colorbar.set_ticks([0, 0.25, 0.5, 0.75, 1.0])
    colorbar.set_ticklabels(["0%", "25%", "50%", "75%", "100%"])
    colorbar.ax.tick_params(labelsize=9)

    missing_handle = Patch(
        facecolor="#D9D9D9", edgecolor="#999999", label="无有效作答（不计 0 分）"
    )
    fig.legend(
        handles=[missing_handle],
        loc="lower left",
        bbox_to_anchor=(0.10, 0.008),
        frameon=False,
        fontsize=9,
    )
    fig.text(
        0.34,
        0.014,
        "彩色格：得分率 0%（深红）→ 100%（深绿）；真实 0 分保留在色阶中。",
        ha="left",
        fontsize=9,
        color="#555555",
    )
    fig.tight_layout(rect=(0, 0.04, 1, 1))

    path = figures_dir / "学生知识点得分率热力图.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def plot_score_distribution(tables_dir: Path, figures_dir: Path) -> Path:
    """图4：学生-知识点得分率箱线图（不以提交记录为样本）。"""
    long = read_figure_table(
        tables_dir,
        "student_knowledge_score_rate.csv",
        dtype={"student_ID": "string"},
    )
    performance = read_figure_table(
        tables_dir, "knowledge_performance.csv", dtype={"knowledge": "string"}
    ).sort_values("mean_score_rate", ascending=False)
    order = performance["knowledge"].tolist()
    counts = dict(zip(performance["knowledge"], performance["n_students"]))

    groups = [
        long.loc[long["knowledge"].eq(k), "score_rate"].dropna().to_numpy()
        * 100
        for k in order
    ]
    means = [values.mean() for values in groups]
    labels = [
        f"{knowledge}\n(n={int(counts[knowledge]):,})" for knowledge in order
    ]

    fig, ax = plt.subplots(figsize=(12, 6.5))
    box = ax.boxplot(
        groups,
        tick_labels=labels,
        patch_artist=True,
        widths=0.55,
        showfliers=False,
        boxprops={"facecolor": "#CBD5F0", "edgecolor": "#4C5B8A"},
        medianprops={"color": "#B45309", "linewidth": 1.8},
        whiskerprops={"color": "#4C5B8A"},
        capprops={"color": "#4C5B8A"},
    )

    # 离群值用固定随机种子的横向抖动散点展示，避免同一取值的点完全重叠
    rng = np.random.default_rng(2024)
    for index, values in enumerate(groups):
        q1, q3 = np.percentile(values, [25, 75])
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        outliers = values[(values < lower) | (values > upper)]
        if outliers.size:
            jitter = rng.uniform(-0.10, 0.10, outliers.size)
            ax.scatter(
                np.full(outliers.size, index + 1) + jitter,
                outliers,
                s=7,
                facecolor="#888888",
                edgecolor="none",
                alpha=0.35,
                zorder=3,
            )
    ax.scatter(
        range(1, len(groups) + 1),
        means,
        marker="D",
        s=30,
        facecolor="white",
        edgecolor="#B45309",
        linewidth=1.0,
        zorder=4,
    )

    ax.set_ylim(-3, 104)
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    ax.set_xlabel("主知识点编码（按平均得分率降序）", fontsize=11)
    ax.set_ylabel("学生-知识点得分率（%）", fontsize=11)
    ax.set_title("各主知识点学生得分率分布", fontsize=15, pad=40)
    ax.tick_params(axis="both", labelsize=9.5)
    ax.grid(axis="y", linestyle="--", alpha=0.25, zorder=0)
    ax.set_axisbelow(True)
    ax.spines[["top", "right"]].set_visible(False)

    legend_handles = [
        Line2D([0], [0], color="#B45309", linewidth=2, label="中位数"),
        Line2D(
            [0],
            [0],
            marker="D",
            color="none",
            markerfacecolor="white",
            markeredgecolor="#B45309",
            label="均值",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            color="none",
            markerfacecolor="#999999",
            label="离群值",
        ),
    ]
    ax.legend(
        handles=legend_handles,
        loc="lower center",
        bbox_to_anchor=(0.5, 1.01),
        ncol=3,
        frameon=False,
        fontsize=9,
    )
    fig.text(
        0.5,
        0.015,
        "注：样本为 10,810 个学生-知识点得分率（每名学生每个知识点一个值），不是提交记录。"
        "箱体为 P25–P75，中线为中位数，菱形为均值，灰点为离群值；"
        "各知识点中位数均为 100%，差异主要体现在低分尾部。",
        ha="center",
        fontsize=8.5,
        color="#555555",
    )
    fig.tight_layout(rect=(0, 0.07, 1, 1))

    path = figures_dir / "知识点得分率分布.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def render_figures(dirs: dict) -> None:
    """生成四张正式图表。"""
    setup_plot_style()
    tables_dir = dirs["tables"]
    figures_dir = dirs["figures"]
    intermediate_dir = dirs["intermediate"]

    paths = [
        plot_knowledge_coverage(tables_dir, figures_dir),
        plot_knowledge_performance(tables_dir, figures_dir),
        plot_student_knowledge_heatmap(
            tables_dir, intermediate_dir, figures_dir
        ),
        plot_score_distribution(tables_dir, figures_dir),
    ]
    print("\n—— 图表生成 ——")
    for path in paths:
        print("已保存：", path)


def render_flowchart(report_dir: Path) -> Path:
    """生成与代码执行顺序一致的知识点与掌握情况模块数据处理流程图。"""
    setup_plot_style()
    steps = [
        ("1. 原始数据读取", "学生表、题目表、15 个班级提交文件"),
        ("2. 数据质量检查", "字段缺失、取值、满分冲突、题目覆盖"),
        ("3. class 字段修复", "6 条字面值 class 按来源文件修复"),
        ("4. 完全重复记录处理", "九字段完全相同的 1,006 条记录单独存档并排除"),
        ("5. 学生信息匹配", "排除 1 条无法匹配学生表的记录"),
        ("6. 有效作答筛选", "分数与时间可解析、且 0 ≤ 得分 ≤ 满分"),
        ("7. 最终作答提取", "每生每题取最晚有效记录（50,482 组）；另取首次作答"),
        ("8. 主/子知识点分别关联", "43 条 / 44 条，两套映射不混用"),
        ("9. 学生-知识点得分率", "学生在该知识点下各题得分率等权平均"),
        ("10. 知识点覆盖与表现统计", "题量、参与学生、提交记录、学生等权均值"),
        ("11. 敏感性分析", "首次 / 最终 / 全部提交加权三种口径对照"),
        ("12. 可视化图表生成", "覆盖题量、得分率、热力图、箱线图"),
    ]

    fig, ax = plt.subplots(figsize=(9, 13.8))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    top = 0.925
    gap = 0.0725
    height = 0.055
    width = 0.64
    center_x = 0.5

    for index, (title, detail) in enumerate(steps):
        y = top - index * gap
        box = FancyBboxPatch(
            (center_x - width / 2, y - height),
            width,
            height,
            boxstyle="round,pad=0.004,rounding_size=0.012",
            linewidth=1.2,
            edgecolor="#5470C6",
            facecolor="#EAF0FB",
            zorder=2,
        )
        ax.add_patch(box)
        ax.text(
            center_x,
            y - height / 2 + 0.010,
            title,
            ha="center",
            va="center",
            fontsize=11.2,
            color="#1F2D3D",
            fontweight="bold",
            zorder=3,
        )
        ax.text(
            center_x,
            y - height / 2 - 0.013,
            detail,
            ha="center",
            va="center",
            fontsize=8.6,
            color="#555555",
            zorder=3,
        )
        if index < len(steps) - 1:
            ax.annotate(
                "",
                xy=(center_x, y - gap + 0.002),
                xytext=(center_x, y - height - 0.002),
                arrowprops={
                    "arrowstyle": "-|>",
                    "color": "#5470C6",
                    "linewidth": 1.4,
                    "shrinkA": 0,
                    "shrinkB": 0,
                },
                zorder=1,
            )

    fig.suptitle("知识点与掌握情况模块数据处理流程图", fontsize=18, y=0.975)
    fig.text(
        0.5,
        0.016,
        "注：步骤与 scripts/d_knowledge.py 的实际执行顺序一致；全流程只读原始数据。\n"
        "统计表、审计记录与图表写入 outputs/d_knowledge/，模块说明与流程图写入 outputs/d_knowledge/report/。",
        ha="center",
        va="bottom",
        fontsize=8.5,
        linespacing=1.5,
        color="#555555",
    )
    fig.subplots_adjust(left=0.03, right=0.97, top=0.95, bottom=0.05)

    path = report_dir / "知识点与掌握情况模块数据处理流程图.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def render_overall_flowchart(report_dir: Path) -> Path:
    """绘制项目整体处理流程图，并标注各模块当前是否已接入。"""
    setup_plot_style()
    report_dir.mkdir(parents=True, exist_ok=True)

    module_states = [
        ("A", "A 学生与班级", "学生构成、班级参与、班级成绩",
         (ROOT / "outputs" / "a_students").exists()),
        ("B", "B 答题行为与时间", "时间分布、活跃天数、重试行为",
         (ROOT / "outputs" / "b_behaviors").exists()),
        ("C", "C 题目与作答表现", "首次/最终正确率、题目指标",
         (ROOT / "outputs" / "c_questions").exists()),
        ("D", "D 知识点与掌握情况", "覆盖、得分率、热力图",
         (ROOT / "outputs" / "d_knowledge").exists()),
    ]

    fig, ax = plt.subplots(figsize=(14, 9))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    def draw_box(center_x, center_y, width, height, title, detail, *,
                 face="#EAF0FB", edge="#5470C6", dashed=False,
                 title_color="#1F2D3D", detail_color="#555555",
                 title_size=12, detail_size=8.5):
        patch = FancyBboxPatch(
            (center_x - width / 2, center_y - height / 2),
            width,
            height,
            boxstyle="round,pad=0.004,rounding_size=0.012",
            linewidth=1.2,
            edgecolor=edge,
            facecolor=face,
            linestyle="--" if dashed else "-",
            zorder=2,
        )
        ax.add_patch(patch)
        ax.text(
            center_x,
            center_y + height * 0.18,
            title,
            ha="center",
            va="center",
            fontsize=title_size,
            fontweight="bold",
            color=title_color,
            zorder=3,
        )
        ax.text(
            center_x,
            center_y - height * 0.16,
            detail,
            ha="center",
            va="center",
            fontsize=detail_size,
            color=detail_color,
            linespacing=1.35,
            zorder=3,
        )

    def arrow(x1, y1, x2, y2, color="#5470C6", linestyle="-"):
        ax.annotate(
            "",
            xy=(x2, y2),
            xytext=(x1, y1),
            arrowprops={
                "arrowstyle": "-|>",
                "color": color,
                "linewidth": 1.3,
                "linestyle": linestyle,
                "shrinkA": 0,
                "shrinkB": 0,
            },
            zorder=1,
        )

    # 1. 原始数据
    draw_box(
        0.5, 0.895, 0.74, 0.075,
        "原始数据（data/raw/chinavis2024_education）",
        "学生表 1,364 行｜题目表 44 行｜15 个班级提交文件 232,818 行",
        title_size=12.5, detail_size=8.6,
    )
    # 2. 全组统一口径
    draw_box(
        0.5, 0.775, 0.84, 0.075,
        "全组统一口径",
        "编号文本｜时间单位与时区｜状态编码｜九字段去重｜有效作答｜得分率｜缺失值",
        title_size=12.5, detail_size=8.6,
    )
    arrow(0.5, 0.8575, 0.5, 0.8135)

    # 3. 四个模块并行
    module_centers = [0.14, 0.38, 0.62, 0.86]
    for center_x, (code, title, detail, active) in zip(module_centers, module_states):
        draw_box(
            center_x, 0.575, 0.19, 0.16,
            title, detail,
            face="#EAF0FB" if active else "#F7F7F7",
            edge="#5470C6" if active else "#999999",
            dashed=not active,
            title_color="#1F2D3D" if active else "#777777",
            detail_color="#555555" if active else "#888888",
            title_size=11, detail_size=8.2,
        )
        ax.text(
            center_x, 0.527,
            "状态：已接入" if active else "状态：待接入",
            ha="center", va="center",
            fontsize=8.2,
            color="#2E7D32" if active else "#999999",
            zorder=3,
        )

    # 口径到四模块的连接线
    ax.plot([0.5, 0.5], [0.7375, 0.70], color="#5470C6", linewidth=1.3, zorder=1)
    ax.plot([0.14, 0.86], [0.70, 0.70], color="#5470C6", linewidth=1.3, zorder=1)
    for center_x, (_, _, _, active) in zip(module_centers, module_states):
        arrow(
            center_x, 0.70, center_x, 0.6565,
            color="#5470C6" if active else "#999999",
            linestyle="-" if active else "--",
        )

    # 4. 模块输出
    draw_box(
        0.5, 0.37, 0.80, 0.085,
        "各模块独立输出",
        "outputs/a_students｜outputs/b_behaviors｜outputs/c_questions｜outputs/d_knowledge\n"
        "每模块含 figures / tables / audit / intermediate",
        title_size=12, detail_size=8.3,
    )
    # 模块到输出的连接线
    for center_x, (_, _, _, active) in zip(module_centers, module_states):
        ax.plot(
            [center_x, center_x], [0.495, 0.455],
            color="#5470C6" if active else "#999999",
            linewidth=1.2,
            linestyle="-" if active else "--",
            zorder=1,
        )
    ax.plot([0.14, 0.86], [0.455, 0.455], color="#5470C6", linewidth=1.3, zorder=1)
    arrow(0.5, 0.455, 0.5, 0.4135)

    # 5. D 汇总
    draw_box(
        0.5, 0.245, 0.70, 0.075,
        "D 汇总全组指标字典与整体处理流程",
        "outputs/d_knowledge/report/ 内：全组指标字典.md｜项目整体处理流程图.png",
        title_size=12, detail_size=8.3,
    )
    arrow(0.5, 0.3275, 0.5, 0.2835)

    # 6. A 合并 Word
    draw_box(
        0.5, 0.13, 0.58, 0.075,
        "A 合并 Word 实验报告",
        "各模块正文、图表与表格",
        title_size=12, detail_size=8.3,
    )
    arrow(0.5, 0.2075, 0.5, 0.1685)

    fig.suptitle("项目整体数据处理流程图", fontsize=18, y=0.965)
    fig.text(
        0.5,
        0.022,
        "注：A、D 已接入现有脚本与输出；B、C 按分工计划预留，待其脚本与输出接入后更新。\n"
        "全流程只读原始数据，各模块输出到独立目录，公共口径见 outputs/d_knowledge/report/全组指标字典.md。",
        ha="center",
        va="bottom",
        fontsize=8.5,
        linespacing=1.5,
        color="#555555",
    )
    fig.subplots_adjust(left=0.02, right=0.98, top=0.945, bottom=0.02)

    path = report_dir / "项目整体处理流程图.png"
    fig.savefig(path, dpi=FIGURE_DPI)
    plt.close(fig)
    return path


def main() -> None:
    dirs = ensure_output_dirs()
    students = read_students()
    titles = read_titles()
    records, profile = read_submission_files()
    ctx = run_quality_checks(students, titles, records, profile, dirs)
    run_metrics(ctx)
    render_figures(dirs)
    render_flowchart(dirs["report"])
    render_overall_flowchart(dirs["report"])


if __name__ == "__main__":
    main()
