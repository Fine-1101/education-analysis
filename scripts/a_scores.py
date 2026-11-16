"""A模块成绩分析。先运行 a_students.py，再运行本脚本。"""
from pathlib import Path
import json
import pandas as pd
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'outputs/a_students'
RAW = ROOT / 'data/raw/chinavis2024_education'
students = pd.read_csv(RAW / 'Data_StudentInfo.csv', dtype={'student_ID': 'string'})
records = pd.read_csv(OUT / 'intermediate/records_deduplicated.csv', dtype='string')
records = records[records.student_ID.isin(students.student_ID)].copy()
titles = pd.read_csv(RAW / 'Data_TitleInfo.csv', dtype='string')
titles['full_score'] = pd.to_numeric(titles['score'], errors='coerce')
if titles.groupby('title_ID').full_score.nunique().gt(1).any():
    raise ValueError('同一题目分值冲突，需人工核查')
# 题目表的一题多行对应知识点关系，不可直接与日志多对多连接。
lookup = titles[['title_ID', 'full_score']].drop_duplicates()
data = records.merge(lookup, on='title_ID', how='left', validate='many_to_one')
assert len(data) == len(records)
data['earned_score'] = pd.to_numeric(data.score, errors='coerce')
data['time_numeric'] = pd.to_numeric(data.time, errors='coerce')
bad = (data.full_score.isna() | data.full_score.le(0)
       | data.earned_score.isna() | data.earned_score.lt(0)
       | data.earned_score.gt(data.full_score) | data.time_numeric.isna())
data.loc[bad].to_csv(OUT / 'audit/score_invalid_records.csv', index=False, encoding='utf-8-sig')
valid = data.loc[~bad].copy()
valid['score_rate'] = valid.earned_score / valid.full_score
keys = ['student_ID', 'title_ID']
# 最晚时刻若有不同结果，无法确定先后；仅从最终成绩分析中排除该组合。
latest = valid[valid.time_numeric.eq(valid.groupby(keys).time_numeric.transform('max'))].copy()
tie_variation = latest.groupby(keys).score_rate.nunique()
ambiguous_keys = tie_variation[tie_variation.gt(1)].reset_index()[keys]
ambiguous = latest.merge(ambiguous_keys, on=keys, how='inner')
ambiguous.to_csv(OUT / 'audit/latest_score_ambiguous.csv', index=False, encoding='utf-8-sig')
latest = latest.merge(ambiguous_keys.assign(ambiguous=True), on=keys, how='left')
latest = latest[latest.ambiguous.isna()].drop_duplicates(keys)
assert not latest.duplicated(keys).any()
latest.to_csv(OUT / 'intermediate/student_question_latest.csv', index=False, encoding='utf-8-sig')
base = records.groupby('student_ID').agg(
    class_name=('class','first'), submission_count=('student_ID','size'),
    attempted_questions=('title_ID','nunique')).reset_index()
rates = latest.groupby('student_ID').agg(
    scored_questions=('title_ID','nunique'), mean_score_rate=('score_rate','mean')).reset_index()
summary = students.merge(base, on='student_ID', how='left', validate='one_to_one').merge(rates, on='student_ID', how='left', validate='one_to_one')
summary.to_csv(OUT / 'tables/student_summary.csv', index=False, encoding='utf-8-sig')
cs = summary.groupby('class_name').agg(
    participants=('student_ID','size'), scored_students=('mean_score_rate','count'),
    mean_score_rate=('mean_score_rate','mean'), median_score_rate=('mean_score_rate','median'),
    mean_scored_questions=('scored_questions','mean')).reset_index()
cs['order'] = cs.class_name.str.extract(r'(\d+)$')[0].astype(int)
cs = cs.sort_values('order').drop(columns='order')
cs.to_csv(OUT / 'tables/class_score_summary.csv', index=False, encoding='utf-8-sig')
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
fig, ax = plt.subplots(figsize=(12,5))
groups = [summary.loc[summary.class_name.eq(c), 'mean_score_rate'].dropna().mul(100).tolist() for c in cs.class_name]
ax.boxplot(groups, tick_labels=cs.class_name.tolist(), patch_artist=True,
           boxprops={'facecolor':'#CBD5F0'}, medianprops={'color':'#B45309'})
ax.set(title='各班级学生最终有效作答平均得分率分布', xlabel='班级', ylabel='学生平均得分率（%）', ylim=(-3,103))
ax.tick_params(axis='x', rotation=45)
ax.grid(axis='y', alpha=.2)
fig.text(.5,.015,'每题取最终有效作答；学生内题目等权；未作答不计零；各班题目覆盖可能不同。',ha='center',fontsize=8)
fig.tight_layout(rect=(0,.05,1,1))
fig.savefig(OUT / 'figures/班级得分率分布.png', dpi=300)
plt.close(fig)
metrics = {'title_rows':len(titles),'unique_questions':len(lookup),
           'matched_records':len(records),'invalid_score_or_time_records':int(bad.sum()),
           'ambiguous_student_question_pairs':len(ambiguous_keys),
           'latest_pairs':len(latest),'students':len(summary),
           'students_with_scores':int(summary.mean_score_rate.notna().sum()),
           'overall_mean_student_score_rate':float(summary.mean_score_rate.mean()),
           'question_coverage_min':float(summary.scored_questions.min()),
           'question_coverage_max':float(summary.scored_questions.max())}
(OUT / 'audit/score_analysis_metrics.json').write_text(json.dumps(metrics,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(metrics, ensure_ascii=False, indent=2))
print(cs.to_string(index=False))
