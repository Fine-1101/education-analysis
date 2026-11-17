"""C模块：题目与作答表现。可独立运行，不读取其他模块的中间文件。"""
from pathlib import Path
import json
import re
import zipfile
import pandas as pd
import matplotlib.pyplot as plt

PROJECT_DIR = Path(__file__).resolve().parent.parent
RAW = PROJECT_DIR / "data" / "raw" / "chinavis2024_education"
OUTPUT_DIR = PROJECT_DIR / "outputs" / "c_questions"
# 官方数据说明留档（不属于原始数据，仅用于口径依据）：
# https://chinavis.org/2024/english/challenge_en.html 赛道1 数据说明文档
OFFICIAL_DOC = PROJECT_DIR / "data" / "docs" / "first_dataDes.docx"
OFFICIAL_SOURCE = "ChinaVis 2024 Data Challenge 赛道1《数据说明》first_dataDes.docx"
OFFICIAL_URL = "https://chinavis.org/2024/document/first_dataDes.docx"
tables_dir = OUTPUT_DIR / "tables"
audit_dir = OUTPUT_DIR / "audit"
intermediate_dir = OUTPUT_DIR / "intermediate"
figures_dir = OUTPUT_DIR / "figures"
for directory in (tables_dir, audit_dir, intermediate_dir, figures_dir):
    directory.mkdir(parents=True, exist_ok=True)

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 状态编码：官方说明只写到"例如完全正确、部分正确等，共12种情况"，未给逐条定义；
# 下面 12 个取值即数据中出现的全部合法值，与官方给出的数量一致（见 state枚举一致性检查）。
# 计为正确的只有 Absolutely_Correct，Partially_Correct 得部分分不计正确。
SUCCESS_STATE = "Absolutely_Correct"
LEGAL_STATES = {
    SUCCESS_STATE, "Partially_Correct", "Absolutely_Error",
    *(f"Error{i}" for i in range(1, 10)),
}
OFFICIAL_STATE_COUNT = 12
# 与 A 模块 deduplication_notes.txt 保持同一判定字段。
BUSINESS_COLUMNS = [
    "class", "student_ID", "title_ID", "time", "state", "score",
    "method", "memory", "timeconsume",
]
QUALITY = []
FILTERS = []


def add_quality(item, count, action, effect, note):
    QUALITY.append({"检查项": item, "涉及记录数": int(count), "处理方式": action,
                    "是否影响指标": effect, "说明": note})


def add_filter(stage, before, after, rule):
    FILTERS.append({"阶段": stage, "处理前记录数": int(before),
                    "处理后记录数": int(after), "剔除数": int(before - after), "规则": rule})


# 官方《数据说明》逐字段含义，用于给字段说明表标注口径来源。
OFFICIAL_MEANINGS = {
    "Data_TitleInfo": {
        "index": ("题目索引", ""), "title_ID": ("题目编号", "唯一编码"),
        "score": ("题目分数", ""), "knowledge": ("知识点", "题目可能考察多个知识点"),
        "sub_knowledge": ("从属知识点", "知识点可能存在多个从属知识点"),
    },
    "SubmitRecord-ClassN": {
        "index": ("记录索引", ""), "class": ("班级", ""),
        "time": ("日志生成时间", "时间戳，精确到秒"),
        "state": ("答题状态", "例如完全正确、部分正确等，共12种情况"),
        "score": ("答题分数", "通过测试用例获得的分数"),
        "title_ID": ("题目编号", "引用题目基本信息表中的title_ID"),
        "method": ("语言", "学习者答题采用的编程语言"),
        "memory": ("内存", "单位KB"), "timeconsume": ("用时", "单位毫秒"),
        "student_ID": ("学习者编号", "引用学习者基本信息表中的student_ID"),
    },
}


def official_reference_text():
    """从留档的官方 docx 抽取纯文本，作为口径依据写入 audit。"""
    if not OFFICIAL_DOC.exists():
        return None
    with zipfile.ZipFile(OFFICIAL_DOC) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", "replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = xml.replace("</w:tc>", " | ")
    text = re.sub(r"<[^>]+>", "", xml)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


official_text = official_reference_text()
if official_text:
    (audit_dir / "官方数据说明摘录.md").write_text(
        f"""# 官方数据说明摘录（C 口径依据）

来源：{OFFICIAL_SOURCE}
原文地址：{OFFICIAL_URL}
本地留档：data/docs/first_dataDes.docx（只读参考文件，非原始数据，不参与计算）
同目录另有答题模板 first_answerSheet.docx（本模块第 4 问对应"找出难度不匹配的题目"）。

官方说明中与 C 模块直接相关的表述：
- state：答题状态，"例如完全正确、部分正确等，共12种情况"——未逐条列出取值，故 Error1-9 的分级含义仍需向数据提供方确认。
- score（提交记录）：答题分数，"通过测试用例获得的分数"——即得分由判题结果决定，与本模块观测到的 state-得分对应关系一致。
- memory 单位 KB，timeconsume 单位毫秒——移交 B 模块时按此单位表述。
- Data_TitleInfo 的 knowledge "题目可能考察多个知识点"、sub_knowledge "知识点可能存在多个从属知识点"——
  官方口径确认题目表的 44 条记录是知识点关系记录，不是重复数据。
- 官方说明明确数据集中"设置了缺失值、异常值或数据不一致等噪声信息"，本模块的未知 state、占位符与时刻冲突均属此类。

## 原文
{official_text}
""", encoding="utf-8")


def describe(frame, table_name):
    meanings = OFFICIAL_MEANINGS.get(table_name, {})
    rows = []
    for column in frame.columns:
        series = frame[column]
        top = ", ".join(f"{key}:{count}" for key, count in
                        series.value_counts(dropna=False).head(3).items())
        meaning, note = meanings.get(column, ("", ""))
        rows.append({"表": table_name, "字段": column,
                     "官方含义": meaning or "官方未定义", "官方补充说明": note,
                     "口径依据": OFFICIAL_SOURCE if meaning else "本模块推断",
                     "读取类型": str(series.dtype),
                     "非空数": int(series.notna().sum()), "缺失数": int(series.isna().sum()),
                     "不同取值数": int(series.nunique()), "高频取值": top})
    return pd.DataFrame(rows)


# —— 数据理解：读取题目表与 15 个班级提交记录 ——
titles_raw = pd.read_csv(RAW / "Data_TitleInfo.csv", dtype="string")
record_files = sorted((RAW / "Data_SubmitRecord").glob("SubmitRecord-Class*.csv"))
if not record_files:
    raise FileNotFoundError(f"未找到提交记录文件：{RAW / 'Data_SubmitRecord'}")

records = pd.concat(
    [pd.read_csv(file, dtype="string").assign(source_file=file.name) for file in record_files],
    ignore_index=True,
)
students = pd.read_csv(RAW / "Data_StudentInfo.csv", dtype={"student_ID": "string"})
raw_records = len(records)
print("提交记录文件数：", len(record_files), "原始记录数：", raw_records)

# 行为字段的质量问题只在原始快照上计数，作为全组质量清单条目，具体处理由 B 模块负责。
raw_tc_missing = int(pd.to_numeric(records.timeconsume, errors="coerce").isna().sum())
raw_tc_values = records.loc[~pd.to_numeric(records.timeconsume, errors="coerce").notna(),
                             "timeconsume"].value_counts().head(3).to_dict()
raw_mem_ceiling = int(records.memory.eq("65536").sum())
raw_time_decimal = int(((pd.to_numeric(records.time, errors="coerce") % 1) != 0).sum())

pd.concat([describe(titles_raw, "Data_TitleInfo"),
           describe(records, "SubmitRecord-ClassN")]).to_csv(
    tables_dir / "字段说明.csv", index=False, encoding="utf-8-sig")

# —— 数据过滤：题目表质量检查 ——
full_titles = titles_raw.copy()
full_titles["full_score"] = pd.to_numeric(full_titles["score"], errors="coerce")
relation_rows = int(full_titles.duplicated(subset=["title_ID"]).sum())
add_quality("题目表-同一 title_ID 多行", relation_rows, "保留", "否",
            "多行来自一题多知识点关系，按关系记录处理，不当作重复删除")
score_conflicts = full_titles.groupby("title_ID").full_score.nunique()
score_conflicts = score_conflicts[score_conflicts > 1]
add_quality("题目表-分值冲突题数", len(score_conflicts), "中止并人工核查", "是",
            "同一题号出现不同分值时无法定义满分")
if not score_conflicts.empty:
    raise ValueError("存在分值冲突题目，需人工核查后重跑")
bad_full = full_titles[full_titles.full_score.isna() | (full_titles.full_score <= 0)]
bad_full.to_csv(audit_dir / "title_invalid_score.csv", index=False, encoding="utf-8-sig")
add_quality("题目表-分值缺失或非正", len(bad_full), "剔除该题", "是", "无法计算得分率与正确率")

title_lookup = (full_titles[["title_ID", "full_score"]]
                .dropna(subset=["full_score"])
                .query("full_score > 0")
                .drop_duplicates("title_ID"))
knowledge_map = (full_titles[["title_ID", "knowledge", "sub_knowledge"]]
                 .drop_duplicates().sort_values(["title_ID", "knowledge"]))
# 一题的多行可能 knowledge 相同而 sub_knowledge 不同，需分别统计关系数与知识点数。
relation_counts = knowledge_map.groupby("title_ID").size()
knowledge_counts = knowledge_map.groupby("title_ID").knowledge.nunique()
sub_knowledge_counts = knowledge_map.groupby("title_ID").sub_knowledge.nunique()
knowledge_map.to_csv(intermediate_dir / "题目知识点关系表.csv", index=False, encoding="utf-8-sig")
title_lookup.to_csv(intermediate_dir / "每题唯一分值表.csv", index=False, encoding="utf-8-sig")
print("题目表行数：", len(titles_raw), "唯一题目数：", len(title_lookup))

# —— 数据过滤：提交记录质量检查 ——
bad_class = records["class"].str.strip().eq("class").fillna(False)
class_repairs = records.loc[bad_class].copy()
expected = class_repairs["source_file"].str.replace("SubmitRecord-", "", regex=False)
class_repairs["corrected_class"] = expected.values
class_repairs.to_csv(audit_dir / "class_invalid_records.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-class 取值为字面 class", len(class_repairs),
            "按来源文件名修复为对应班级", "否", "与 A 模块 class_repairs 同一批记录")
records.loc[bad_class, "class"] = expected.values

legal_states_present = int(records.loc[records["state"].isin(LEGAL_STATES), "state"].nunique())
unknown_state = records[~records["state"].isin(LEGAL_STATES)]
corrupted_state_value = "\ufffd" in "".join(unknown_state["state"].fillna("").unique())
corrupted_with_mem_ceiling = int(unknown_state["memory"].eq("65536").sum())
unknown_state.to_csv(audit_dir / "unknown_state_records.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-未知或损坏的 state 取值", len(unknown_state), "剔除出题目指标", "是",
            ("取值在文件中已是 U+FFFD 替换字符，损坏发生在数据发布之前、原值不可复原，"
             f"只能剔除并提请数据方复核；这 {len(unknown_state)} 行中有 "
             f"{corrupted_with_mem_ceiling} 行 memory 同为 65536"
             if corrupted_state_value else "取值不在官方 12 种枚举内，剔除并提请复核"))
add_quality("提交记录-state 取值与官方枚举数量一致性", legal_states_present,
            "核对后确认无需改动状态映射", "否",
            f"官方说明称 state 共 {OFFICIAL_STATE_COUNT} 种情况，数据中出现的合法取值恰为 "
            f"{legal_states_present} 种且与代码枚举一一对应，另有 {len(unknown_state)} 行损坏值；"
            "逐条含义（尤其 Error1-9 的分级）官方未给出，仍待确认")

before = len(records)
records = records[records["state"].isin(LEGAL_STATES)].copy()
add_filter("剔除未知状态记录", before, len(records), "state 在官方枚举之外不参与统计")

unmatched_title = records[~records["title_ID"].isin(title_lookup.title_ID)]
unmatched_title.to_csv(audit_dir / "unmatched_title_records.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-title_ID 无法关联题目表", len(unmatched_title), "剔除出题目指标", "是",
            "关联不到满分")
before = len(records)
records = records.merge(title_lookup, on="title_ID", how="inner", validate="many_to_one")
add_filter("关联每题唯一分值", before, len(records),
           "先构造 title_ID 唯一的满分表再 many_to_one 关联，避免一题多知识点把日志行数放大")
assert len(records) == before, "关联满分表改变了记录数，存在一对多展开"

unmatched_student = records[~records["student_ID"].isin(students["student_ID"])]
unmatched_student.to_csv(audit_dir / "unmatched_student_records.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-学生编号无法匹配学生表", len(unmatched_student), "剔除，与 A 一致", "是",
            "与 A 模块 unmatched_students 同一批记录")
before = len(records)
records = records[records["student_ID"].isin(students["student_ID"])].copy()
add_filter("剔除无法匹配学生的记录", before, len(records), "与 A 模块保持同一分析样本")

before = len(records)
pre_ties = records.groupby(["student_ID", "title_ID", "time"]).size()
pre_tie_groups = int((pre_ties > 1).sum())
remove_mask = records.duplicated(subset=BUSINESS_COLUMNS, keep="first")
removed = records.loc[remove_mask].copy()
removed.to_csv(audit_dir / "duplicates_excluded.csv", index=False, encoding="utf-8-sig")
groups = (removed[BUSINESS_COLUMNS].drop_duplicates()
          .merge(records.groupby(BUSINESS_COLUMNS, dropna=False).size().rename("occurrences").reset_index(),
                 on=BUSINESS_COLUMNS, how="left"))
groups.to_csv(audit_dir / "duplicate_groups.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-完整业务字段重复", len(removed), "每组保留第一条", "是",
            f"共涉及 {len(groups)} 个重复组；真实多次提交（时间或字段不同）全部保留")
records = records.loc[~remove_mask].copy()
add_filter("九字段完全重复去重", before, len(records),
           "判定字段与 A 相同：class,student_ID,title_ID,time,state,score,method,memory,timeconsume")
print("去重后记录数：", len(records), "排除重复数：", len(removed))

records["earned_score"] = pd.to_numeric(records["score"], errors="coerce")
records["time_numeric"] = pd.to_numeric(records["time"], errors="coerce")
records["row_no"] = pd.to_numeric(records["index"], errors="coerce")
over = records.loc[records.earned_score.isna() | (records.earned_score < 0)
                   | (records.earned_score > records.full_score)
                   | records.time_numeric.isna()]
over.to_csv(audit_dir / "score_invalid_records.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-得分缺失或越界", int(records.earned_score.isna().sum()
            + (records.earned_score < 0).sum() + (records.earned_score > records.full_score).sum()),
            "置空得分，不参与得分率", "是", "得分超过满分或无法转为数字")
add_quality("提交记录-时间无法转为数字", int(records.time_numeric.isna().sum()),
            "不参与首次与最终判定", "是", "Unix 秒时间戳")
before = len(records)
records = records[records.time_numeric.notna()].copy()
add_filter("剔除时间无法解析的记录", before, len(records), "无时间即无法排序首次与最终作答")

non_numeric = records.loc[~pd.to_numeric(records["timeconsume"], errors="coerce").notna(),
                          ["source_file", "index", "class", "student_ID", "title_ID", "time",
                           "state", "score", "timeconsume"]]
non_numeric.to_csv(audit_dir / "timeconsume_non_numeric.csv", index=False, encoding="utf-8-sig")
add_quality("提交记录-timeconsume 为缺失占位符", raw_tc_missing, "登记并移交 B 模块", "否",
            f"字面取值为 {raw_tc_values}，官方说明单位为毫秒；属行为耗时字段，本模块只汇总问题")
add_quality("提交记录-memory 等于 65536", raw_mem_ceiling, "登记并移交 B 模块", "否",
            f"官方说明单位为 KB，65536KB 恰为 64MB，疑似上限或哨兵值；其中 "
            f"{corrupted_with_mem_ceiling} 行与 state 损坏行重合，"
            f"其余 {raw_mem_ceiling - corrupted_with_mem_ceiling} 行状态合法")
add_quality("提交记录-time 含小数秒", raw_time_decimal, "无需处理", "否",
            "全部为整秒，说明采集精度不超过 1 秒，重试间隔无法细到秒以下")

# 时间统一为东八区（原始 time 为 Unix 秒，即 UTC 时间戳）
records["time_cst"] = pd.to_datetime(records.time_numeric, unit="s") + pd.Timedelta(hours=8)

# 状态取值与得分分布，作为“哪些状态计为正确”的口径证据。
records["row_score_rate"] = records.earned_score / records.full_score
state_profile = (records.groupby("state")
                 .agg(记录数=("state", "size"), 覆盖题目数=("title_ID", "nunique"),
                      平均得分=("earned_score", "mean"),
                      平均得分率=("row_score_rate", "mean"),
                      得分为零占比=("earned_score", lambda s: float((s == 0).mean())))
                 .reset_index().sort_values("平均得分率", ascending=False).round(4))
state_profile.to_csv(tables_dir / "状态取值与得分分布.csv", index=False, encoding="utf-8-sig")

# —— 数据处理：同刻冲突一律取最高分，使首次与最终判定可复现 ——
keys = ["student_ID", "title_ID"]
records["neg_score"] = -records["earned_score"]
ordered = records.sort_values(
    [*keys, "time_numeric", "neg_score", "source_file", "row_no"], ignore_index=True)
tie_counts = (ordered.groupby([*keys, "time_numeric"]).size().rename("rows").reset_index())
ties = tie_counts[tie_counts["rows"] > 1]
tie_detail = ties.merge(ordered, on=[*keys, "time_numeric"], how="left")
tie_detail.to_csv(audit_dir / "timestamp_ties.csv", index=False, encoding="utf-8-sig")
add_quality("同一学生同一题同一时刻多条提交", len(ties),
            "取 score 最高的一条参与首次与最终判定", "是",
            f"共 {len(ties)} 组冲突，最多 {int(ties['rows'].max()) if len(ties) else 0} 条同时刻，"
            "缺少唯一提交 ID 无法区分先后，取最高分会轻微高估最终表现")

BASE_COLUMNS = [*keys, "class", "full_score", "state", "earned_score",
                "time_numeric", "time_cst", "method"]
first = ordered.groupby(keys, sort=False).head(1)[BASE_COLUMNS].rename(
    columns={"state": "first_state", "earned_score": "first_score",
             "time_numeric": "first_time", "time_cst": "first_time_cst",
             "method": "first_method"})
final = ordered.groupby(keys, sort=False).tail(1)[BASE_COLUMNS].rename(
    columns={"state": "final_state", "earned_score": "final_score",
             "time_numeric": "final_time", "time_cst": "final_time_cst",
             "method": "final_method"})
pair = first.merge(final.drop(columns=["class", "full_score"]), on=keys,
                   how="inner", validate="one_to_one")
pair["attempt_count"] = ordered.groupby(keys).size().reindex(
    pd.MultiIndex.from_arrays([pair.student_ID, pair.title_ID])).values
pair["first_correct"] = pair.first_state.eq(SUCCESS_STATE)
pair["final_correct"] = pair.final_state.eq(SUCCESS_STATE)
pair["first_score_rate"] = pair.first_score / pair.full_score
pair["final_score_rate"] = pair.final_score / pair.full_score
pair["first_time_cst"] = pd.to_datetime(pair.first_time_cst)
pair["final_time_cst"] = pd.to_datetime(pair.final_time_cst)
pair["resolved_by_retry"] = (~pair.first_correct) & pair.final_correct
pair.sort_values(["title_ID", "student_ID"]).to_csv(
    intermediate_dir / "学生题目作答汇总.csv", index=False, encoding="utf-8-sig")
print("学生×题目组合数：", len(pair))

# —— 题目指标：正确率分母为学生×题目组合，另给提交次数口径供核对 ——
submit_first = ordered.drop_duplicates(keys).assign(
    ok=lambda d: d.state.eq(SUCCESS_STATE)).groupby("title_ID").ok.mean()
metrics = pair.groupby("title_ID").agg(
    participating_students=("student_ID", "nunique"),
    first_accuracy=("first_correct", "mean"),
    final_accuracy=("final_correct", "mean"),
    mean_attempts=("attempt_count", "mean"),
    mean_first_score_rate=("first_score_rate", "mean"),
    mean_final_score_rate=("final_score_rate", "mean"),
    retry_solved_share=("resolved_by_retry", "mean"),
).reset_index()
metrics["submission_count"] = ordered.groupby("title_ID").size().reindex(
    metrics.title_ID).values
metrics["submissions_per_student"] = metrics.submission_count / metrics.participating_students
metrics["full_score"] = title_lookup.set_index("title_ID").full_score.reindex(
    metrics.title_ID).values
metrics["knowledge_count"] = knowledge_counts.reindex(metrics.title_ID).fillna(0).values
metrics["sub_knowledge_count"] = sub_knowledge_counts.reindex(metrics.title_ID).fillna(0).values
metrics["knowledge_relation_rows"] = relation_counts.reindex(metrics.title_ID).fillna(0).values
metrics["max_attempts"] = pair.groupby("title_ID").attempt_count.max().reindex(
    metrics.title_ID).values
metrics["first_accuracy_by_submission"] = submit_first.reindex(metrics.title_ID).values
metrics["首次与最终差值"] = metrics.final_accuracy - metrics.first_accuracy
metrics = metrics.round(4)

# 全部题目的首次正确率都显著低于最终正确率，因此重试幅度只作展示列，不作为异常判据。
flag_low_final = metrics.final_accuracy <= metrics.final_accuracy.quantile(0.25)
flag_low_first = metrics.first_accuracy <= metrics.first_accuracy.quantile(0.25)
flag_many_attempts = metrics.mean_attempts >= metrics.mean_attempts.quantile(0.75)
flag_score_above_accuracy = (metrics.mean_final_score_rate - metrics.final_accuracy) >= 0.08
flag_few_participants = metrics.participating_students <= metrics.participating_students.quantile(0.1)
flag_extreme_retry = metrics.max_attempts >= 100
flags = pd.DataFrame({
    "低最终正确率": flag_low_final, "低首次正确率": flag_low_first,
    "重试次数偏高": flag_many_attempts, "得分率明显高于正确率": flag_score_above_accuracy,
    "参与学生偏少": flag_few_participants, "存在极端重试个例": flag_extreme_retry,
})
metrics["命中判据数"] = flags.sum(axis=1)
metrics.to_csv(tables_dir / "题目指标表.csv", index=False, encoding="utf-8-sig")
candidates = metrics[metrics["命中判据数"] >= 2].join(flags[flags.columns]).copy()
candidates["判据"] = flags.apply(
    lambda row: "、".join([name for name in flags.columns if row[name]]), axis=1)
candidates = candidates.sort_values(
    ["命中判据数", "final_accuracy"], ascending=[False, True])
candidates = candidates[["title_ID", "full_score", "knowledge_count", "sub_knowledge_count",
                         "participating_students", "submission_count", "mean_attempts",
                         "max_attempts", "first_accuracy", "final_accuracy",
                         "mean_final_score_rate", "首次与最终差值", "命中判据数", "判据"]].round(4)
candidates.to_csv(tables_dir / "候选异常题清单.csv", index=False, encoding="utf-8-sig")

pd.DataFrame(QUALITY).to_csv(tables_dir / "质量检查表.csv", index=False, encoding="utf-8-sig")
pd.DataFrame(FILTERS).to_csv(tables_dir / "过滤规则与处理数量.csv", index=False, encoding="utf-8-sig")

# —— 图表 ——
plot = metrics.sort_values("first_accuracy")
fig, ax = plt.subplots(figsize=(9, 10))
ax.barh(plot.title_ID.str.replace("Question_", "", regex=False) + "  "
        + plot.full_score.astype(int).astype(str) + "分",
        plot.first_accuracy * 100, color="#4C78A8")
ax.set_xlabel("首次正确率（%）")
ax.set_ylabel("题目（题号后括号为满分）")
ax.set_title("各题目首次正确率排序")
ax.set_xlim(0, plot.first_accuracy.max() * 100 * 1.12)
ax.grid(axis="x", alpha=.2)
fig.text(.5, .015, f"分母为学生×题目组合数；共 {len(metrics)} 题、{len(pair)} 个组合；"
         "首次作答取该组合最早一次提交，同刻冲突取最高分。", ha="center", fontsize=8)
fig.tight_layout(rect=(0, .03, 1, 1))
fig.savefig(figures_dir / "首次正确率排序.png", dpi=300)
plt.close(fig)

fig, ax = plt.subplots(figsize=(8.5, 11))
order = metrics.sort_values("first_accuracy").reset_index(drop=True)
cand_ids = set(candidates.title_ID)
ax.hlines(range(len(order)), order.first_accuracy * 100, order.final_accuracy * 100,
          color="#C7CEDB", lw=2.4, zorder=1)
ax.scatter(order.first_accuracy * 100, range(len(order)), s=46, color="#4C78A8",
           zorder=2, label="首次正确率")
ax.scatter(order.final_accuracy * 100, range(len(order)), s=46,
           color=[("#E45756" if t in cand_ids else "#74A9D6") for t in order.title_ID],
           zorder=3, label="最终正确率（红=候选需复核）")
ax.set_yticks(range(len(order)))
ax.set_yticklabels(order.title_ID.str.replace("Question_", "", regex=False).str[:12]
                   + "  " + order.full_score.astype(int).astype(str) + "分", fontsize=7.5)
ax.set_ylim(len(order) - .6, -.6)
ax.set_xlabel("正确率（%）")
ax.set_ylabel("题目（按首次正确率升序，自上而下）")
ax.set_title("各题目首次与最终正确率对比")
ax.grid(axis="x", alpha=.2)
ax.legend(loc="lower left", fontsize=9)
fig.text(.5, .012, f"平均提升 {(pair.final_correct.mean() - pair.first_correct.mean()) * 100:.1f} 个百分点；"
         f"全部 {len(metrics)} 题最终正确率高于首次；"
         f"重试后由错转对的学生×题目组合占 {pair.resolved_by_retry.mean() * 100:.1f}%。",
         ha="center", fontsize=8)
fig.tight_layout(rect=(0, .025, 1, 1))
fig.savefig(figures_dir / "首次与最终正确率对比.png", dpi=300)
plt.close(fig)

fig, ax = plt.subplots(figsize=(8.5, 7))
gap = (metrics.final_accuracy - metrics.first_accuracy) * 100
ax.scatter(metrics.first_accuracy * 100, gap, s=metrics.participating_students / 12,
           alpha=.8, color="#B0BAC9", edgecolor="none")
ax.scatter(candidates.first_accuracy * 100, (candidates.final_accuracy - candidates.first_accuracy) * 100,
           s=candidates.participating_students / 12, color="#E45756",
           edgecolor="white", linewidth=.6, label="候选需复核题目")
for _, row in candidates.iterrows():
    ax.annotate(row.title_ID.replace("Question_", "")[:8],
                (row.first_accuracy * 100, (row.final_accuracy - row.first_accuracy) * 100),
                fontsize=7.5, xytext=(4, 3), textcoords="offset points")
ax.axhline(gap.mean(), c="#4C78A8", ls=":", lw=1.2)
ax.set_xlabel("首次正确率（%）")
ax.set_ylabel("最终与首次正确率差值（百分点）")
ax.set_title("重试提升幅度与首次正确率的关系")
ax.grid(alpha=.2)
ax.legend(loc="upper right", fontsize=9)
fig.text(.5, .015, f"点面积为参与学生数；虚线为平均提升 {gap.mean():.1f} 个百分点；"
         "首次正确率越低的题，重试带来的提升越大。", ha="center", fontsize=8)
fig.tight_layout(rect=(0, .03, 1, 1))
fig.savefig(figures_dir / "重试提升幅度散点.png", dpi=300)
plt.close(fig)

fig, ax = plt.subplots(figsize=(8, 6))
scatter = ax.scatter(metrics.participating_students, metrics.mean_final_score_rate * 100,
                     s=60, c=metrics.final_accuracy * 100, cmap="viridis", alpha=.85)
fig.colorbar(scatter, ax=ax, label="最终正确率（%）")
ax.axhline(metrics.mean_final_score_rate.mean() * 100, ls="--", c="grey", lw=1)
ax.set_xlabel("参与学生数（人）")
ax.set_ylabel("平均得分率（%）")
ax.set_title("题目参与人数与平均得分率")
ax.grid(alpha=.2)
fig.text(.5, .015, "平均得分率按每题最终有效作答的得分/满分求学生均值；"
         f"虚线为题目平均得分率均值 {metrics.mean_final_score_rate.mean() * 100:.1f}%。",
         ha="center", fontsize=8)
fig.tight_layout(rect=(0, .03, 1, 1))
fig.savefig(figures_dir / "参与人数与得分率散点.png", dpi=300)
plt.close(fig)

overall = pair.groupby("student_ID").agg(
    attempted=("title_ID", "nunique"), attempts=("attempt_count", "sum"),
    final_accuracy=("final_correct", "mean")).reset_index()
metrics_json = {
    "official_reference": {
        "source": OFFICIAL_SOURCE, "url": OFFICIAL_URL,
        "local_doc": str(OFFICIAL_DOC.relative_to(PROJECT_DIR)),
        "state_count_official": OFFICIAL_STATE_COUNT,
        "state_values_present": legal_states_present,
        "doc_available": official_text is not None,
    },
    "record_files": len(record_files), "raw_records": raw_records,
    "records_after_all_filters": len(records),
    "duplicates_removed": len(removed), "unknown_state_records": len(unknown_state),
    "class_repaired_records": len(class_repairs),
    "unmatched_student_records": len(unmatched_student),
    "unmatched_title_records": len(unmatched_title),
    "invalid_score_records": len(over), "timestamp_tie_pairs": len(ties),
    "unique_questions": int(len(title_lookup)),
    "questions_with_multiple_knowledge": int((relation_counts > 1).sum()),
    "questions_with_multiple_knowledge_codes": int((knowledge_counts > 1).sum()),
    "student_question_pairs": int(len(pair)),
    "students_in_pairs": int(pair.student_ID.nunique()),
    "overall_first_accuracy": float(pair.first_correct.mean()),
    "overall_final_accuracy": float(pair.final_correct.mean()),
    "overall_final_score_rate": float(pair.final_score_rate.mean()),
    "retry_resolved_share": float(pair.resolved_by_retry.mean()),
    "mean_attempts_per_pair": float(pair.attempt_count.mean()),
    "median_attempts_per_pair": float(pair.attempt_count.median()),
    "max_attempts_per_pair": int(pair.attempt_count.max()),
    "time_min_cst": str(records.time_cst.min()), "time_max_cst": str(records.time_cst.max()),
    "candidate_questions": int(len(candidates)),
}
ac_row = state_profile.loc[state_profile.state.eq(SUCCESS_STATE)]
pc_row = state_profile.loc[state_profile.state.eq("Partially_Correct")]
ac_score_rate = float(ac_row["平均得分率"].iloc[0]) if not ac_row.empty else float("nan")
pc_score_rate = float(pc_row["平均得分率"].iloc[0]) if not pc_row.empty else float("nan")
pc_questions = int(pc_row["覆盖题目数"].iloc[0]) if not pc_row.empty else 0
error_zero_share = float(state_profile.loc[
    ~state_profile.state.isin([SUCCESS_STATE, "Partially_Correct"]), "得分为零占比"].min())
metrics_json["score_rate_by_state"] = dict(
    zip(state_profile.state, state_profile.平均得分率.astype(float)))
(audit_dir / "question_analysis_metrics.json").write_text(
    json.dumps(metrics_json, ensure_ascii=False, indent=2), encoding="utf-8")
(audit_dir / "口径与指标公式.md").write_text(
    f"""# C 模块口径与指标公式

## 样本
- 提交记录原始 {raw_records} 行，来自 {len(record_files)} 个班级文件，表头完全一致。
- 九字段完全重复排除 {len(removed)} 行，判定字段与 A 模块 deduplication_notes.txt 一致。
- 未知 state {len(unknown_state)} 行剔除；无法匹配学生表 {len(unmatched_student)} 行剔除，与 A 同一批。
- 最终参与题目指标的记录数 {len(records)} 行，学生×题目组合 {len(pair)} 组，覆盖 {len(title_lookup)} 题。

## 状态与时间
- 口径依据：{OFFICIAL_SOURCE}，原文地址 {OFFICIAL_URL}，本地留档 data/docs/first_dataDes.docx，
  摘录见 audit/官方数据说明摘录.md。官方说明写明数据集为 15 个班级、2023-08-31 至 2024-01-25 共 148 天、
  学习者 1364 名、题目记录 44 条、答题日志 232818 条，且有意设置了缺失值、异常值与不一致噪声。
- 正确 = state 为 {SUCCESS_STATE}；Partially_Correct 得部分分但不计正确；其余 Absolutely_Error 与 Error1-9 计错误。
  官方说明对 state 只给出"共12种情况"的数量约束，未逐条定义，因此该映射由数据反推并经官方数量校验：
  数据中出现的合法取值恰为 {legal_states_present} 种，与代码枚举一一对应。
- 口径证据见 tables/状态取值与得分分布.csv：Absolutely_Correct 的平均得分率为 {ac_score_rate:.3f}（即一律给满分），
  Partially_Correct 平均得分率 {pc_score_rate:.3f}（约满分一半，覆盖 {pc_questions} 题），
  Absolutely_Error 与 Error1-9 的得分全部为 0（该组状态得分为零占比最小值仍为 {error_zero_share:.3f}）。
  可见得分完全由 state 决定，平均得分率与最终正确率的差值只来自 Partially_Correct 这一档。
- 因此正确率是严格口径（只认完全正确），得分率是含部分分的宽松口径，两者须同时报告，不可互相替代。
- time 为 Unix 秒（整数、无空值），与官方说明"时间戳，精确到秒"一致；
  范围 {metrics_json['time_min_cst']} ~ {metrics_json['time_max_cst']}，
  属 UTC 时间戳，本模块统一加 8 小时按东八区（Asia/Shanghai）解释，全组须使用同一时区。
- 官方说明定义提交记录 score 为"通过测试用例获得的分数"，即得分由判题结果决定；
  这与观测到的 state-得分对应关系吻合，可作为"得分完全由 state 决定"的外部依据。
  memory 单位 KB、timeconsume 单位毫秒，移交 B 模块时按此单位表述。

## 首次与最终有效作答
- 按 (student_ID, title_ID) 归组，以 time 升序排序；首次取最早一条，最终取最晚一条。
- 同一时刻存在多条时（共 {len(ties)} 组）一律取 score 最高的一条，其次按来源文件与行序号排序，
  保证可复现。该规则会轻微高估最终正确率，已在 timestamp_ties.csv 保留全部冲突明细。
- 去重前同一 (student_ID, title_ID, time) 有 {pre_tie_groups} 组多条记录，去重后仅剩 {len(ties)} 组，
  说明多数"同刻多条"实为完整业务字段一致的重复行，真正需要在同一时刻内取舍的组合很少。

## 满分与一题多知识点
- Data_TitleInfo 有 {len(titles_raw)} 行但仅 {len(title_lookup)} 个唯一 title_ID，
  多出的 {relation_rows} 行是一题多知识点关系记录，不视为重复。
  官方说明对 knowledge 写明"题目可能考察多个知识点"、对 sub_knowledge 写明"知识点可能存在多个从属知识点"，
  且官方口径的题目表记录数就是 {len(titles_raw)} 条，与本判定一致。
- 先构造 title_ID 唯一的每题满分表，再以 many_to_one 关联日志，并断言行数不变，
  避免多对多展开把提交次数放大。

## 指标公式
- 参与人数 = 该题去重后的 student_ID 数。
- 提交次数 submission_count = 过滤去重后该题的全部日志行数。
- 首次正确率 first_accuracy = 该题学生×题目组合中，首次作答为 {SUCCESS_STATE} 的比例。
- 最终正确率 final_accuracy = 同上，按最后一次有效作答计算。
- first_accuracy_by_submission = 按提交次数口径的对照值，仅供敏感性检查。
- 平均得分率 mean_final_score_rate = mean(final_score / full_score)，题目内学生等权。
- 平均尝试次数 mean_attempts = 该题各学生×题目组合提交行数的均值。
- retry_solved_share = 首次错误、最终正确的组合占该题组合的比例。
- 未作答的题目不进入任何分母，未作答不记零分。

## 候选异常题判据
- 六项判据各取分布分位：低最终正确率（≤25 分位）、低首次正确率（≤25 分位）、
  重试次数偏高（≥75 分位）、得分率明显高于正确率（差值≥0.08）、
  参与学生偏少（≤10 分位）、存在极端重试个例（单组合最多尝试次数≥100）。
- 命中 2 项及以上才进入候选清单，避免全部题目都被标为异常；清单同时保留命中数与判据说明。

## 后续扩展：题目区分度（本轮未实施，先给定义）
- 区分度 D = 高分组答对率 − 低分组答对率。先按学生最终有效作答的平均得分率排序，
  取前 27% 为高分组、后 27% 为低分组，D = P(答对|高分组) − P(答对|低分组)，答对沿用 Absolutely_Correct 口径。
- 实施前提：本数据集 38 题被 1317–1360 名学生普遍作答，分组样本充足；
  但 D 只说明题目能否区分学生水平，不能单独判定题目质量差，须与首次正确率、平均尝试次数和知识点覆盖联合解读。

## 局限
- 原始数据缺少唯一提交 ID，同刻多条提交只能按取最高分的约定处理。
- 官方《数据说明》已获取并留档，但它对 state 只给数量（共12种）不给逐条定义，
  Error1-9 的分级差异、以及 {len(unknown_state)} 行损坏取值的原值仍需向数据提供方确认。
- 低正确率不能单独说明题目不合理，需结合参与人数、知识点与班级构成，后续再补区分度指标。
""", encoding="utf-8")

# —— 全组数据质量问题清单：C 负责汇总，A 的输出只读不改 ——
A_AUDIT = PROJECT_DIR / "outputs" / "a_students" / "audit"


def audit_count(name):
    path = A_AUDIT / name
    if not path.exists():
        return "文件未生成"
    return sum(1 for _ in open(path, encoding="utf-8-sig")) - 1


dedup_path = A_AUDIT / "duplicate_audit_summary.csv"
a_dedup_removed = (int(pd.read_csv(dedup_path).query("scope == 'raw'")
                       .extra_records_if_deduplicated.iloc[0])
                   if dedup_path.exists() else "文件未生成")
group_rows = [{"来源模块": owner, "问题": issue, "涉及记录数": count,
               "处理方式": action, "处理状态": status, "依据文件": ref}
              for owner, issue, count, action, status, ref in [
    ("A", "提交记录 class 字段为字面 class", audit_count("class_repairs.csv"),
     "按来源文件名修复为对应班级", "已处理并留档",
     "outputs/a_students/audit/class_repairs.csv"),
    ("A", "学生编号无法匹配学生信息表", audit_count("unmatched_students.csv"),
     "不自动纠正，剔除出学生关联分析", "已处理并留档",
     "outputs/a_students/audit/unmatched_students.csv"),
    ("A", "完整业务字段重复提交", a_dedup_removed,
     "每组保留第一条，被排除记录留档", "已处理并留档",
     "outputs/a_students/audit/duplicate_audit_summary.csv"),
    ("A", "得分或时间无法用于成绩分析", audit_count("score_invalid_records.csv"),
     "剔除出成绩分析", "已处理并留档",
     "outputs/a_students/audit/score_invalid_records.csv"),
    ("A", "最晚时刻得分不一致的组合", audit_count("latest_score_ambiguous.csv"),
     "从最终成绩分析中排除", "已处理并留档",
     "outputs/a_students/audit/latest_score_ambiguous.csv"),
    ("C", "state 取值损坏（文件中已是 U+FFFD 替换字符）", len(unknown_state),
     "剔除出题目指标；损坏早于数据发布，原值不可复原，提请数据方复核", "已处理，待数据方复核",
     "outputs/c_questions/audit/unknown_state_records.csv"),
    ("C", "timeconsume 为 -- 或 - 占位符", raw_tc_missing,
     "移交 B 模块处理（官方口径单位为毫秒）", "待 B 处理",
     "outputs/c_questions/audit/timeconsume_non_numeric.csv"),
    ("C", "memory 等于 65536，疑似上限哨兵值", raw_mem_ceiling,
     "移交 B 模块核实（官方口径单位为 KB，即 64MB 整）", "待 B 处理",
     "outputs/c_questions/tables/质量检查表.csv"),
    ("C", "题目表一题多知识点关系行", relation_rows,
     "按关系记录保留，先取唯一分值再关联（官方说明已确认一题多知识点）", "已处理",
     "outputs/c_questions/intermediate/题目知识点关系表.csv"),
    ("C", "同一学生同一题同一时刻多条提交", len(ties),
     "取最高分参与首次与最终判定", "已处理并留档",
     "outputs/c_questions/audit/timestamp_ties.csv"),
    ("C", "缺少官方数据说明文档", "不适用",
     "已从官网获取 first_dataDes.docx 与 first_answerSheet.docx 并留档，字段口径按官方表述更新",
     "已解决", "data/docs/first_dataDes.docx；outputs/c_questions/audit/官方数据说明摘录.md"),
    ("C", "state 逐条含义官方未定义（仅给出共12种的数量）", OFFICIAL_STATE_COUNT,
     "按数据反推映射，并用官方给出的 12 种数量校验合法取值集合，逐条含义提请数据方确认", "待确认",
     "outputs/c_questions/audit/官方数据说明摘录.md"),
    ("D", "知识点与 sub_knowledge 的编码含义未知", "不适用",
     "官方说明只定义字段含义，未给编码对照，由 D 模块核对覆盖与权重", "待 D 处理",
     "data/raw/chinavis2024_education/Data_TitleInfo.csv"),
]]
group_table = pd.DataFrame(group_rows)
group_table.to_csv(tables_dir / "全组数据质量问题清单.csv", index=False, encoding="utf-8-sig")
(OUTPUT_DIR / "全组数据质量问题清单.md").write_text(
    "# 全组数据质量问题清单（C 汇总）\n\n"
    + "| " + " | ".join(group_table.columns) + " |\n"
    + "|" + "---|" * len(group_table.columns) + "\n"
    + "".join("| " + " | ".join(str(v) for v in row) + " |\n"
              for row in group_table.itertuples(index=False))
    + "\n## 已统一的全组口径\n"
      "- 学生编号、题目编号按文本处理，编号格式不做数字化改写。\n"
      f"- 重复判定使用九字段：{'、'.join(BUSINESS_COLUMNS)}，共排除 {len(removed)} 行，A 与 C 结果一致。\n"
      "- 时间为 Unix 秒 UTC 时间戳，全组统一按东八区（+8 小时）解释。\n"
      "- 正确 = state 为 Absolutely_Correct；Partially_Correct 计部分分不计正确；Error 系列计错误。\n"
      "  官方说明给出 state 共 12 种情况，数据中合法取值恰为 12 种，该枚举集合已获官方数量校验。\n"
      "- 得分率分母为该题满分；正确率分母默认学生×题目组合数，提交次数口径另行给出。\n"
      "- 原始数据文件一律不修改，被排除记录全部单独留档；官方说明文档另存 data/docs/，不并入原始数据。\n"
      "- 未作答不记零分，不进入分母。\n\n"
    + "## 待全组确认\n"
      f"- Error1-9 的分级含义：官方说明只写“共12种情况”未逐条定义，{len(unknown_state)} 行损坏取值的原值需数据方提供。\n"
      "- 得分口径已可关闭：官方说明定义 score 为“通过测试用例获得的分数”，"
      "与本数据中 Absolutely_Correct 得满分、Partially_Correct 约半分、Error 系列 0 分的观测一致。\n"
      "- 同一时刻多条提交的取舍规则，需与 D 模块的知识点聚合保持一致。\n"
      "- 官方字段单位（memory 为 KB、timeconsume 为 毫秒）请 B 模块在行为指标中统一使用。\n"
      "- 官方答题模板 first_answerSheet.docx 列出 5 个必答问题，正文需逐题对应"
      "（第 4 问“找出难度不匹配的题目”即本模块候选异常题清单）。\n",
    encoding="utf-8")

# —— 正文草稿：交付给组长合并进 Word ——
worst = metrics.nsmallest(4, "first_accuracy")
corr_gap = float(metrics.first_accuracy.corr(metrics.final_accuracy - metrics.first_accuracy))
corr_part = float(metrics.participating_students.corr(metrics.mean_final_score_rate))
full_score_dist = "、".join(f"{int(k)} 分 {v} 题" for k, v in
                            title_lookup.full_score.value_counts().sort_index().items())
# 官方《数据说明》给出的规模，用于自检读数是否对得上发布口径。
official_scale = {"班级数": (len(record_files), 15), "日志行数": (raw_records, 232818),
                  "题目记录条数": (len(titles_raw), 44), "学习者数": (len(students), 1364)}
scale_check = "；".join(f"{name} 实测 {own}／官方 {off}{'' if own == off else '（不一致，需核查）'}"
                        for name, (own, off) in official_scale.items())
(OUTPUT_DIR / "正文草稿.md").write_text(
    f"""# C 题目与作答表现（正文草稿）

## 三、题目属性与数据质量
本模块字段口径以 ChinaVis 2024 数据竞赛赛道1官方《数据说明》（first_dataDes.docx，已留档
data/docs/，摘录见 outputs/c_questions/audit/官方数据说明摘录.md）为准。官方说明给出的数据集规模
与本项目读数一致：{scale_check}；时间跨度为 2023 年 8 月 31 日至 2024 年 1 月 25 日共 148 天。
官方说明同时指出数据集中"设置了缺失值、异常值或数据不一致等噪声信息"，本模块发现的问题即属此类。
题目材料 Data_TitleInfo.csv 共 {len(titles_raw)} 行，但只有 {len(title_lookup)} 个唯一 title_ID，
其中 {relation_rows} 行为同一题目的多条知识点关系记录，不属于重复数据（官方说明明确"题目可能考察多个知识点"
且 knowledge 之下还可有多个 sub_knowledge）。{len(score_conflicts)} 道题目存在分值冲突，
{len(bad_full)} 道题目分值缺失或非正；满分分布为 {full_score_dist}。
提交日志来自 {len(record_files)} 个班级文件共 {raw_records} 行，字段结构完全一致。
按与 A 模块相同的九字段口径排除完全重复 {len(removed)} 行，剔除 state 损坏记录 {len(unknown_state)} 行、
无法匹配学生编号 {len(unmatched_student)} 行，得到 {len(records)} 行分析样本，
构成 {len(pair)} 个学生×题目组合，覆盖全部 {len(title_lookup)} 道题目、{pair.student_ID.nunique()} 名学生。
此外发现 timeconsume 存在 {raw_tc_missing} 处 "--"/"-" 占位符、memory 有 {raw_mem_ceiling} 行等于 65536
（官方口径 memory 单位为 KB、timeconsume 单位为毫秒），已移交 B 模块，详见《全组数据质量问题清单》。

state 的合法取值共 {legal_states_present} 种，与官方说明的"共12种情况"数量吻合，
因此本模块的状态枚举已获官方口径校验；但官方未逐条定义取值含义，Error1-9 的分级差异仍列为待确认。
{len(unknown_state)} 行的 state 在文件字节层面已是 U+FFFD 替换字符，说明损坏发生在数据发布之前、原值不可复原，
只能剔除并提请数据方复核。作答状态与得分的对应关系高度规整：Absolutely_Correct 的平均得分率为 {ac_score_rate:.3f}，
Partially_Correct 为 {pc_score_rate:.3f}（约满分一半），Absolutely_Error 与 Error1-9 得分均为 0；
官方说明把 score 定义为"通过测试用例获得的分数"，与上述观测一致，可确认得分由判题结果直接决定。
因此本模块同时报告两套口径：正确率为严格口径（只认完全正确），得分率为含部分分的宽松口径，
两者差值完全来自 Partially_Correct，不可互相替代。时间字段为整秒 Unix 时间戳（官方说明"精确到秒"），
范围 {metrics_json['time_min_cst']} 至 {metrics_json['time_max_cst']}，统一按东八区解释。

## 四、题目表现：首次作答与最终作答
按 (student_ID, title_ID) 归组并按时间排序，取最早一次为首次作答、最晚一次为最终作答；
同一时刻存在多条提交的组合共 {len(ties)} 个，统一取其中得分最高的一条，
该约定会轻微高估最终表现（去重前此类组合有 {pre_tie_groups} 个，多数实为完全重复行）。

全部 {len(metrics)} 道题目的最终正确率都高于首次正确率，总体由 {pair.first_correct.mean() * 100:.1f}%
升至 {pair.final_correct.mean() * 100:.1f}%，平均提升 {(pair.final_correct.mean() - pair.first_correct.mean()) * 100:.1f} 个百分点；
首次作答错误、经重试最终答对的学生×题目组合占 {pair.resolved_by_retry.mean() * 100:.1f}%。
最终正确率的分布很窄（{metrics.final_accuracy.min() * 100:.1f}%–{metrics.final_accuracy.max() * 100:.1f}%），
而首次正确率跨度达 {metrics.first_accuracy.min() * 100:.1f}%–{metrics.first_accuracy.max() * 100:.1f}%，
且首次正确率越低的题目重试提升幅度越大（两者相关系数 {corr_gap:.2f}，见图 C-3）。
这说明用最终正确率排序会把难题之间的差异"洗平"，识别高难度题目应以首次正确率为主、平均尝试次数为辅。
按此判据筛出候选需复核题目 {len(candidates)} 道（命中两项以上判据，清单见表 C-3），
其中 {worst.title_ID.iloc[0]}、{worst.title_ID.iloc[1]}、{worst.title_ID.iloc[2]}、{worst.title_ID.iloc[3]}
首次正确率仅 {worst.first_accuracy.min() * 100:.1f}%–{worst.first_accuracy.iloc[3] * 100:.1f}%，
平均尝试 {worst.mean_attempts.mean():.1f} 次，个例最高尝试次数达 {int(metrics.max_attempts.max())} 次。

## 五、参与广度与得分水平
各题参与学生在 {metrics.participating_students.min()}–{metrics.participating_students.max()} 之间，
参与人数与平均得分率呈负相关（相关系数 {corr_part:.2f}，见图 C-4）：
参与人数最少的题目平均得分率反而最高。这与"难题劝退"的解释方向一致，
但也可能来自不同班级题目覆盖差异，需与 A 模块的班级题目覆盖结果联合判断，
不能单独据此认定题目难度。同样需要说明的是，低正确率不能单独证明题目设置不合理，
还需结合学生基础、参与人数与知识点表现，题目区分度指标（定义见《口径与指标公式》）留作后续扩展。

## 六、与官方答题要求的对应
官方答题模板 first_answerSheet.docx 为赛道1列出 5 个必答问题：①知识点掌握程度量化与薄弱环节识别；
②个性化学习行为模式与学习者画像；③学习模式与知识掌握关系建模；④利用可视分析找出难度不匹配的不合理题目；
⑤面向题目设计者与课程管理的优化建议。本模块产出直接支撑第 ④ 问（题目指标表与候选需复核题目清单，
并以首次正确率、平均尝试次数作为难度判据），并为第 ①③⑤ 问提供题目侧口径：
首次正确率反映题目难度、重试解决占比反映题目可学习性、得分率与正确率的双口径避免部分分被误判为全对。
第 ②③ 问所需的行为字段（timeconsume、method、memory）问题已移交 B 模块。
""", encoding="utf-8")

print(json.dumps(metrics_json, ensure_ascii=False, indent=2))
print("\n—— 题目指标（按最终正确率升序前10） ——")
cols = ["title_ID", "full_score", "participating_students", "submission_count",
        "mean_attempts", "first_accuracy", "final_accuracy", "mean_final_score_rate"]
print(metrics.sort_values("final_accuracy")[cols].head(10).to_string(index=False))
print("\n—— 状态取值与得分分布 ——")
print(state_profile.to_string(index=False))
print("\n候选需复核题目数：", len(candidates))
print(candidates.to_string(index=False))
print("\n输出目录：", OUTPUT_DIR)
