# 1118/papers/v1 — 真实经营数据先导论文

## 文件
- 论文V1_卷烟多品规关联信息增益研究.docx：完整中文论文（含英文题名、结构化摘要、公式文字、三线表、双语图注）。
- 论文V1_完整初稿.md / 论文V1_图文阅读版.html：与Word对应，可在GitHub检索/阅读。
- code/01_run_real_experiments.py：从原始XLS和旧版raw_panel_v1.csv核对并计算7组先导实验；不使用未来信息构造图。
- code/02_build_paper.py：从真实结果JSON自动生成本文Word/Markdown/HTML。
- figures/Fig01—Fig07：真实计算的PNG和SVG（不含假想TFT/GAT准确率）。
- results：源Excel结构、数据审计、算法精度、月度测试汇总（不包含完整SKU预测明细）。

## 实证结论（真实运行）
- 产品：203个，月份：19，面板：3149条。
- 非缺失销量：2390条，缺失：759条，显式零值：460条。
- 原始销售XLS与面板成功匹配2390条，数值一致率100%。
- 一月提前期、5个月独立测试（478条）：Lag1 wMAPE 60.24%，HistGB 71.45%，GraphHistGB 71.43%。
- GraphHistGB相对于HistGB仅降低0.021个百分点，探索性块式Bootstrap区间为（-0.261，0.307）个百分点。
- 未训练TFT、GAT、MTGNN，数据是品规×月份，不是逐日交易流水。

## 复现步骤
从仓库根目录执行：
    pip install pandas numpy scikit-learn matplotlib xlrd python-docx
    python 1118/papers/v1/code/01_run_real_experiments.py
    python 1118/papers/v1/code/02_build_paper.py

## 稿件成熟度说明
V1是**有实际数据和先导实验的完整科研工作稿**，不是宣称图网络创新已成立的投稿定稿。
先补充更长、更高频的SKU销量和正式供给事件记录，再考虑训练TFT/MTGNN/GAT。
不得将本V1的小样本负面结果或原始相关图写成消费者替代/缩投因果作用。
