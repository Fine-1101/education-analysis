from pathlib import Path
import pandas as pd

# 根据脚本位置定位项目目录
PROJECT_DIR = Path(__file__).resolve().parent.parent

# 学生信息文件的位置
student_file = (
    PROJECT_DIR
    / "data"
    / "raw"
    / "chinavis2024_education"
    / "Data_StudentInfo.csv"
)

# 将学生编号按文本读取，避免格式被改变
students = pd.read_csv(
    student_file,
    dtype={"student_ID": "string"},
)

print("数据文件：", student_file)
print("\n前5行数据：")
print(students.head().to_string(index=False))

print("\n行数和列数：", students.shape)
print("\n字段名称：", students.columns.tolist())

print("\n各字段缺失数量：")
print(students.isna().sum())
print("\n不同学生编号数量：")
print(students["student_ID"].nunique())

print("\n完全重复的记录数：")
print(students.duplicated().sum())

print("\n重复学生编号涉及的记录：")
duplicate_ids = students[
    students.duplicated(subset=["student_ID"], keep=False)
]

if duplicate_ids.empty:
    print("没有重复学生编号")
else:
    print(duplicate_ids.to_string(index=False))

print("\n性别分布：")
print(students["sex"].value_counts(dropna=False))

print("\n年龄分布：")
print(students["age"].value_counts(dropna=False).sort_index())

print("\n专业编码分布：")
print(students["major"].value_counts(dropna=False))
import matplotlib.pyplot as plt

# 创建输出文件夹
output_dir = PROJECT_DIR / "outputs" / "a_students"
output_dir.mkdir(parents=True, exist_ok=True)
figures_dir = output_dir / "figures"
tables_dir = output_dir / "tables"
audit_dir = output_dir / "audit"
intermediate_dir = output_dir / "intermediate"
for directory in (figures_dir, tables_dir, audit_dir, intermediate_dir):
    directory.mkdir(parents=True, exist_ok=True)

# 设置中文字体和负号显示
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 按年龄统计人数
age_counts = students["age"].value_counts().sort_index()

# 保存统计表，便于核对图片数据
age_counts.rename_axis("age").reset_index(
    name="student_count"
).to_csv(
    tables_dir / "age_distribution.csv",
    index=False,
    encoding="utf-8-sig",
)

# 绘图
fig, ax = plt.subplots(figsize=(8, 5))

bars = ax.bar(
    age_counts.index.astype(str),
    age_counts.values,
    color="#5470C6",
)

ax.bar_label(bars, padding=3)
ax.set_title("学生年龄分布")
ax.set_xlabel("年龄（岁）")
ax.set_ylabel("学生人数（人）")
ax.set_ylim(0, age_counts.max() * 1.15)

fig.tight_layout()

figure_path = figures_dir / "学生年龄分布.png"
fig.savefig(figure_path, dpi=300)
plt.close(fig)

print("\n图片已保存：", figure_path)
# 找到所有班级的提交记录文件
record_dir = (
    PROJECT_DIR
    / "data"
    / "raw"
    / "chinavis2024_education"
    / "Data_SubmitRecord"
)

record_files = sorted(record_dir.glob("SubmitRecord-Class*.csv"))

print("\n找到的班级记录文件数：", len(record_files))

if not record_files:
    raise FileNotFoundError(f"未找到提交记录，请检查目录：{record_dir}")

# 本轮只读取需要的两列，减少内存占用
records = pd.concat(
    [
        pd.read_csv(
            file,
            usecols=["class", "student_ID"],
            dtype={"class": "string", "student_ID": "string"},
        )
        for file in record_files
    ],
    ignore_index=True,
)

print("\n提交记录总行数：", len(records))

print("\n班级和学生编号的缺失数量：")
print(records.isna().sum())

# 同一学生的多次提交只计为一个班级参与学生
class_counts = (
    records.groupby("class")["student_ID"]
    .nunique()
    .rename("有提交记录的学生数")
)

print("\n各班级有提交记录的学生数：")
print(class_counts)

# 检查一个学生是否在多个班级出现
student_class_counts = records.groupby("student_ID")["class"].nunique()

print("\n出现在多个班级的学生数：")
print((student_class_counts > 1).sum())

# 检查提交者是否都能在学生信息表中找到
record_student_ids = records["student_ID"].dropna().unique()
unknown_ids = set(record_student_ids) - set(students["student_ID"])

print("\n提交记录中的不同学生数：", len(record_student_ids))
print("无法匹配学生信息表的学生数：", len(unknown_ids))
# 1. 检查班级值为“class”的记录来自哪些文件
print("\n——班级字段异常记录——")

for file in record_files:
    part = pd.read_csv(file, dtype="string")

    suspicious = part[
        part["class"].str.strip().eq("class").fillna(False)
    ]

    if not suspicious.empty:
        print(f"\n文件：{file.name}")
        print(suspicious.to_string(index=False))

# 2. 查看无法匹配学生表的编号
print("\n——无法匹配的学生编号——")
print(sorted(unknown_ids))

# 3. 查看跨班级学生及其记录数量
multi_class_ids = student_class_counts[
    student_class_counts > 1
].index

print("\n——跨班级学生的记录数量——")

multi_class_details = (
    records[records["student_ID"].isin(multi_class_ids)]
    .groupby(["student_ID", "class"])
    .size()
    .reset_index(name="记录数量")
)

print(multi_class_details.to_string(index=False))
print("\n——修复班级字段——")

parts = []
repair_logs = []

for file in record_files:
    part = pd.read_csv(file, dtype="string")

    # 例如：SubmitRecord-Class1.csv → Class1
    expected_class = file.stem.replace("SubmitRecord-", "")

    # 只修复已经确认的字面值“class”
    bad_class = part["class"].str.strip().eq("class").fillna(False)

    if bad_class.any():
        log = part.loc[bad_class].copy()
        log["source_file"] = file.name
        log["original_class"] = log["class"]
        log["corrected_class"] = expected_class
        repair_logs.append(log)

        part.loc[bad_class, "class"] = expected_class

    part["source_file"] = file.name
    parts.append(part)

# 保存为新变量，保留此前的原始检查结果
clean_records = pd.concat(parts, ignore_index=True)

if repair_logs:
    repairs = pd.concat(repair_logs, ignore_index=True)
    repairs.to_csv(
        audit_dir / "class_repairs.csv",
        index=False,
        encoding="utf-8-sig",
    )
    print("已修复班级字段的记录数：", len(repairs))

print("修复后记录总数：", len(clean_records))

class_check = clean_records.groupby("student_ID")["class"].nunique()
print("修复后跨班级学生数：", (class_check > 1).sum())

# 单独导出无法匹配学生信息的记录
unmatched = clean_records[
    ~clean_records["student_ID"].isin(students["student_ID"])
]

unmatched.to_csv(
    audit_dir / "unmatched_students.csv",
    index=False,
    encoding="utf-8-sig",
)

print("\n——无法匹配学生信息的提交记录——")
print(unmatched.to_string(index=False))
from difflib import get_close_matches

print("\n——检查无法匹配的学生编号——")

known_ids = students["student_ID"].dropna().tolist()

for unknown_id in unmatched["student_ID"].dropna().unique():
    print("\n异常编号：", unknown_id)

    # 检查是否仅有前后空格差异
    space_matches = [
        sid for sid in known_ids
        if sid.strip() == unknown_id.strip()
    ]
    print("去除前后空格后匹配：", space_matches)

    # 检查是否仅有大小写差异
    case_matches = [
        sid for sid in known_ids
        if sid.lower() == unknown_id.lower()
    ]
    print("忽略大小写后匹配：", case_matches)

    # 查找相似编号，仅供人工核查，不自动替换
    candidates = get_close_matches(
        unknown_id,
        known_ids,
        n=3,
        cutoff=0.8,
    )
    print("相似编号候选：", candidates)

    for candidate in candidates:
        print("候选编号在提交记录中的班级分布：")
        print(
            clean_records.loc[
                clean_records["student_ID"].eq(candidate),
                "class",
            ].value_counts()
        )

# 不猜测修改学生编号；未匹配记录保存在 unmatched_students.csv 中。
matched_records = clean_records[
    clean_records["student_ID"].isin(students["student_ID"])
].copy()

analysis_records = matched_records.merge(
    students[["student_ID", "sex", "age", "major"]],
    on="student_ID",
    how="left",
    validate="many_to_one",
)

# 尚未完成完整日志去重，以下提交量均为初步记录数。
class_summary = (
    analysis_records.groupby("class")
    .agg(
        participating_students=("student_ID", "nunique"),
        record_count=("student_ID", "size"),
    )
    .reset_index()
)
class_summary["records_per_student"] = (
    class_summary["record_count"] / class_summary["participating_students"]
).round(2)
class_summary["class_number"] = (
    class_summary["class"].str.extract(r"(\d+)$")[0].astype(int)
)
class_summary = (
    class_summary.sort_values("class_number")
    .drop(columns="class_number")
)
summary_path = intermediate_dir / "class_summary_initial.csv"
class_summary.to_csv(summary_path, index=False, encoding="utf-8-sig")

print("\n——学生关联结果——")
print("关联前记录数：", len(clean_records))
print("未匹配记录数：", len(unmatched))
print("成功关联记录数：", len(analysis_records))
print("成功关联的不同学生数：", analysis_records["student_ID"].nunique())
print("\n——班级参与情况初步统计——")
print(class_summary.to_string(index=False))
print("\n班级汇总表已保存：", summary_path)

# 完整业务字段重复检查：只检查，不自动删除疑似重复提交。
business_columns = [
    "class", "student_ID", "title_ID", "time", "state", "score",
    "method", "memory", "timeconsume",
]
audit_rows = []
duplicate_groups = []
for scope, frame in [("raw", pd.concat([
    pd.read_csv(file, dtype="string").assign(source_file=file.name)
    for file in record_files
], ignore_index=True)), ("class_repaired", clean_records)]:
    mask = frame.duplicated(subset=business_columns, keep=False)
    candidates = frame.loc[mask].copy()
    extra_count = int(frame.duplicated(subset=business_columns, keep="first").sum())
    group_count = len(candidates[business_columns].drop_duplicates())
    candidates.to_csv(intermediate_dir / f"duplicate_records_{scope}.csv",
                      index=False, encoding="utf-8-sig")
    audit_rows.append({"scope": scope, "total_records": len(frame),
                       "duplicate_groups": group_count,
                       "involved_records": int(mask.sum()),
                       "extra_records_if_deduplicated": extra_count})
    if not candidates.empty:
        groups = (candidates.groupby(business_columns, dropna=False)
                  .size().reset_index(name="occurrences"))
        groups.insert(0, "scope", scope)
        duplicate_groups.append(groups)

audit = pd.DataFrame(audit_rows)
audit.to_csv(audit_dir / "duplicate_audit_summary.csv",
             index=False, encoding="utf-8-sig")
groups = (pd.concat(duplicate_groups, ignore_index=True) if duplicate_groups
          else pd.DataFrame(columns=["scope", *business_columns, "occurrences"]))
groups.to_csv(intermediate_dir / "duplicate_groups.csv", index=False, encoding="utf-8-sig")
print("\n——完整业务字段重复检查（未删除记录）——")
print(audit.to_string(index=False))
print("检查结果目录：", output_dir)

# 第一轮去重口径：完整业务字段相同则保留读取顺序中的第一条。
# 数据未提供可确认唯一提交的ID；该口径是分析假设，不代表确认其为错误记录。
numeric_times = pd.to_numeric(clean_records["time"], errors="coerce")
print("\n——时间与序号检查——")
print("无法转为数字的时间数：", numeric_times.isna().sum())
print("含小数秒的时间数：", ((numeric_times.dropna() % 1) != 0).sum())
candidate_rows = clean_records.loc[
    clean_records.duplicated(subset=business_columns, keep=False)
]
print("疑似重复组内的行序号不同的组数：",
      (candidate_rows.groupby(business_columns, dropna=False)["index"].nunique() > 1).sum())

remove_mask = clean_records.duplicated(subset=business_columns, keep="first")
removed_records = clean_records.loc[remove_mask].copy()
removed_records.to_csv(audit_dir / "duplicates_excluded.csv", index=False,
                       encoding="utf-8-sig")
deduplicated_records = clean_records.loc[~remove_mask].copy()
deduplicated_records.to_csv(intermediate_dir / "records_deduplicated.csv", index=False,
                            encoding="utf-8-sig")
final_records = deduplicated_records[
    deduplicated_records["student_ID"].isin(students["student_ID"])
].copy()
final_summary = final_records.groupby("class").agg(
    participating_students=("student_ID", "nunique"),
    submission_count=("student_ID", "size"),
).reset_index()
final_summary["submissions_per_student"] = (
    final_summary["submission_count"] / final_summary["participating_students"]
).round(2)
final_summary["class_number"] = final_summary["class"].str.extract(r"(\d+)$")[0].astype(int)
final_summary = final_summary.sort_values("class_number").drop(columns="class_number")
final_summary.to_csv(tables_dir / "class_summary_deduplicated.csv", index=False,
                     encoding="utf-8-sig")
comparison = class_summary.merge(final_summary, on="class", suffixes=("_before", "_after"))
comparison["excluded_duplicate_count"] = comparison["record_count"] - comparison["submission_count"]
comparison.to_csv(audit_dir / "class_dedup_comparison.csv", index=False,
                  encoding="utf-8-sig")
report = (
    "第一轮日志去重说明\n"
    "范围：修复班级字段后的全部提交记录。\n"
    "判定字段：" + ", ".join(business_columns) + "。\n"
    "index为行序号、source_file为来源，不作为唯一提交标识。\n"
    "未发现本地官方说明文档，未能核实时间单位；数字时间均为整数时也不能证明原始采集精度。\n"
    "由于缺少可靠的唯一提交ID，无法排除同一时间单位内真实重复提交。\n"
    "第一轮分析约定：完整业务字段一致时保留第一条，不同时间或其他字段不同的提交均保留。\n"
    "原始文件保持不变，被排除记录单独保存；未匹配学生编号不自动纠正。\n"
    f"原始记录数：{len(clean_records)}\n"
    f"按上述口径排除重复记录数：{len(removed_records)}\n"
    f"去重后记录数：{len(deduplicated_records)}\n"
    f"去重后未匹配学生记录数：{len(deduplicated_records) - len(final_records)}\n"
    f"用于学生关联分析的记录数：{len(final_records)}\n"
    "对比表同时保留去重前后的班级统计，供敏感性检查。\n"
)
(audit_dir / "deduplication_notes.txt").write_text(report, encoding="utf-8")
print("\n——第一轮去重结果——")
print(report)
print(final_summary.to_string(index=False))

# 使用去重且成功关联学生信息的数据绘制班级对比图。
chart_specs = [
    ("participating_students", "各班级参与学生人数", "参与学生数（人）", "班级参与人数.png", "%.0f"),
    ("submissions_per_student", "各班级参与学生人均提交量", "人均提交量（次/人）", "班级人均提交量.png", "%.2f"),
]
for column, title, ylabel, filename, label_format in chart_specs:
    fig, ax = plt.subplots(figsize=(12, 5))
    bars = ax.bar(final_summary["class"], final_summary[column],
                  color="#5470C6", width=0.7)
    ax.bar_label(bars, fmt=label_format, padding=3, fontsize=8)
    ax.set_title(title)
    ax.set_xlabel("班级")
    ax.set_ylabel(ylabel)
    ax.set_ylim(0, final_summary[column].max() * 1.18)
    ax.tick_params(axis="x", rotation=45)
    ax.set_axisbelow(True)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines[["top", "right"]].set_visible(False)
    fig.text(0.5, 0.015,
             "口径：业务字段去重后、可关联学生信息的提交；人数为有提交记录的学生，人均分母为参与学生数。",
             ha="center", fontsize=8, color="#555555")
    fig.tight_layout(rect=(0, 0.055, 1, 1))
    chart_path = figures_dir / filename
    fig.savefig(chart_path, dpi=300)
    plt.close(fig)
    print("班级图表已保存：", chart_path)
