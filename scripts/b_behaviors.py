from pathlib import Path
import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parent.parent
INPUT = PROJECT_DIR / "outputs/a_students/intermediate/records_deduplicated.csv"
OUT = PROJECT_DIR / "outputs/b_behaviors"
for sub in ("audit", "intermediate", "tables", "figures"):
    (OUT / sub).mkdir(parents=True, exist_ok=True)

records = pd.read_csv(INPUT, dtype="string")
print("行数、列数：", records.shape)
print("字段：", records.columns.tolist())
print(records.head().to_string(index=False))
print("\ntime 前 10 个值：")
print(records["time"].head(10).tolist())
print("\ntime 描述：")
print(pd.to_numeric(records["time"], errors="coerce").describe())

# ===== 第二段：时间质量检查 + 特征提取 =====
import json

# 1. 转北京时间
records["time_numeric"] = pd.to_numeric(records["time"], errors="coerce")
dt_utc = pd.to_datetime(records["time_numeric"], unit="s", utc=True, errors="coerce")
records["dt_bj"] = dt_utc.dt.tz_convert("Asia/Shanghai")
records["date"] = records["dt_bj"].dt.date
records["hour"] = records["dt_bj"].dt.hour
records["weekday"] = records["dt_bj"].dt.day_of_week  # 0=周一
records["weekday_name"] = records["dt_bj"].dt.day_name()

# 2. 时间质量检查
total = len(records)
na_time = int(records["time_numeric"].isna().sum())
non_int = int(((records["time_numeric"].dropna() % 1) != 0).sum())
le_zero = int((records["time_numeric"] <= 0).sum())
out_of_range = int((~records["dt_bj"].between(
    pd.Timestamp("2023-08-01", tz="Asia/Shanghai"),
    pd.Timestamp("2024-02-01", tz="Asia/Shanghai"))).sum())
checks = pd.DataFrame([
    {"检查项": "总记录数", "规则": "输入去重日志全部行", "命中数": total, "处理方式": "统计基数"},
    {"检查项": "time 无法转数字", "规则": "to_numeric 后为 NaN", "命中数": na_time, "处理方式": "标记，不参与时间分析"},
    {"检查项": "time 含小数秒", "规则": "value % 1 != 0", "命中数": non_int, "处理方式": "保留，按秒取整展示"},
    {"检查项": "time <= 0", "规则": "非正时间戳", "命中数": le_zero, "处理方式": "标记为异常"},
    {"检查项": "时间超出 2023-08~2024-02", "规则": "北京时间范围外", "命中数": out_of_range, "处理方式": "标记为异常，不参与时间分析"},
])
checks.to_csv(OUT / "audit/time_quality_check.csv", index=False, encoding="utf-8-sig")

# 3. 异常时间单独导出（不删除）
invalid = records[records["time_numeric"].isna() | (records["time_numeric"] <= 0)]
invalid.to_csv(OUT / "audit/time_invalid_records.csv", index=False, encoding="utf-8-sig")

# 4. 有效提交（时间分析用）
valid = records[records["time_numeric"].notna() & (records["time_numeric"] > 0)].copy()
valid.to_csv(OUT / "intermediate/records_with_time_features.csv", index=False, encoding="utf-8-sig")

# 5. 时间范围与单位说明
info = {
    "time_unit": "Unix 秒级时间戳 (UTC)",
    "display_timezone": "Asia/Shanghai (UTC+8)",
    "min_bj": str(valid["dt_bj"].min()),
    "max_bj": str(valid["dt_bj"].max()),
    "total_records": total,
    "valid_time_records": len(valid),
    "invalid_time_records": len(invalid),
}
(OUT / "audit/time_unit_conclusion.json").write_text(
    json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")

print("\n—— 时间质量检查表 ——")
print(checks.to_string(index=False))
print("\n—— 时间结论 ——")
print(json.dumps(info, ensure_ascii=False, indent=2))
print("\n北京时间样例：")
print(valid[["time", "dt_bj", "date", "hour", "weekday_name"]].head(10).to_string(index=False))

# ===== 第三段：学生行为汇总 + 单题尝试次数 =====
import matplotlib.pyplot as plt

# 修掉之前的弃用警告（上一行已写入 weekday，这里兼容即可）
# 学生行为汇总：一行一个学生
student_summary = valid.groupby("student_ID").agg(
    class_name=("class", "first"),
    active_days=("date", "nunique"),          # 活跃天数 = 有有效提交的不同日期数
    submission_count=("student_ID", "size"),  # 提交次数
    attempted_questions=("title_ID", "nunique"),  # 参与题量
).reset_index()
student_summary.to_csv(OUT / "tables/student_behavior_summary.csv",
                       index=False, encoding="utf-8-sig")

# 单题尝试次数：按“学生+题目”组合统计，真实重试不删除
attempt = valid.groupby(["student_ID", "title_ID"]).agg(
    attempt_count=("title_ID", "size"),
    first_time=("dt_bj", "min"),
    last_time=("dt_bj", "max"),
    final_state=("state", "last"),
).reset_index()
attempt["is_retry"] = attempt["attempt_count"] > 1
attempt.to_csv(OUT / "tables/question_attempt_summary.csv",
               index=False, encoding="utf-8-sig")

# 时间统计表
daily = valid.groupby("date").size().rename("submission_count").reset_index()
daily.to_csv(OUT / "tables/daily_submission.csv", index=False, encoding="utf-8-sig")

print("\n学生行为汇总（前 5 行）：")
print(student_summary.head().to_string(index=False))
print("\n活跃天数分布：")
print(student_summary["active_days"].describe().to_string())
print("\n单题尝试次数分布：")
print(attempt["attempt_count"].describe().to_string())
print("\n重试对数占比：", round(attempt["is_retry"].mean() * 100, 2), "%")

# ===== 第四段：三张图 =====
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 图1：每日提交量折线图
fig, ax = plt.subplots(figsize=(12, 5))
ax.plot(daily["date"], daily["submission_count"], color="#5470C6", linewidth=1.2)
peak = daily.loc[daily["submission_count"].idxmax()]
ax.scatter([peak["date"]], [peak["submission_count"]], color="#EE6666", zorder=5)
ax.annotate(f"峰值 {peak['submission_count']} 次\n{peak['date']}",
            xy=(peak["date"], peak["submission_count"]),
            xytext=(10, -30), textcoords="offset points", fontsize=8)
ax.set(title="每日提交量变化", xlabel="日期", ylabel="提交次数")
ax.grid(alpha=.2)
fig.text(.5, .015, "口径：A 模块去重后有效提交；时间按北京时间；日界按自然日。",
         ha="center", fontsize=8)
fig.tight_layout(rect=(0, .05, 1, 1))
fig.savefig(OUT / "figures/每日提交量折线图.png", dpi=300)
plt.close(fig)

# 图2：星期 × 小时热力图
pivot = valid.pivot_table(index="weekday", columns="hour",
                          values="student_ID", aggfunc="size", fill_value=0)
weekday_labels = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
fig, ax = plt.subplots(figsize=(14, 4.5))
im = ax.imshow(pivot.values, aspect="auto", cmap="YlOrRd")
ax.set_yticks(range(7)); ax.set_yticklabels(weekday_labels)
ax.set_xticks(range(len(pivot.columns))); ax.set_xticklabels(pivot.columns)
ax.set(title="星期 × 小时提交热力图", xlabel="小时（北京时间）", ylabel="星期")
fig.colorbar(im, ax=ax, label="提交次数")
fig.text(.5, .015, "口径：有效提交按北京时间归入星期与小时；颜色越深提交越多。",
         ha="center", fontsize=8)
fig.tight_layout(rect=(0, .05, 1, 1))
fig.savefig(OUT / "figures/星期小时热力图.png", dpi=300)
plt.close(fig)

# 图3：活跃天数分布
fig, ax = plt.subplots(figsize=(10, 5))
ax.hist(student_summary["active_days"], bins=30, color="#91CC75", edgecolor="white")
ax.set(title="学生活跃天数分布", xlabel="活跃天数（天）", ylabel="学生人数")
ax.grid(axis="y", alpha=.2)
fig.text(.5, .015, "口径：活跃天数 = 有有效提交的不同日期数；不代表每天都学习。",
         ha="center", fontsize=8)
fig.tight_layout(rect=(0, .05, 1, 1))
fig.savefig(OUT / "figures/活跃天数分布图.png", dpi=300)
plt.close(fig)

print("\n三张图已保存到：", OUT / "figures")