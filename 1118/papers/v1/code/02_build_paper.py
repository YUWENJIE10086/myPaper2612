#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Write a full Chinese research-paper V1 from REAL 0805-based experiment outputs.
Creates native editable Word tables and embeds ONLY charts from actual experiments.
Emits HTML + Markdown mirroring the Word content. No simulated metrics or invented results.
Source format follows the repo's prior causal-forest papers and the journal's
Papers2/v18/v18-3/findings/论文模板.txt formatting requirements.
"""
import json, os, html, zipfile, re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[4]
OUT=ROOT/"1118/papers/v1"
RES=OUT/"results"
FIG=OUT/"figures"
AUD=json.loads((RES/"01_data_audit.json").read_text(encoding="utf-8"))
EVAL=json.loads((RES/"02_baseline_metrics.json").read_text(encoding="utf-8"))
MET={(q["h_months"],q["model"]):q for q in EVAL["results"]}
CI=EVAL.get("paired_month_bootstrap") or {}
GROUP={q["h"]:q for q in EVAL["grouping"]}
assert all((1,k) in MET for k in ("Lag1","HistGB","GraphHistGB")), "V1 data experiment missing"
N=int(AUD["derived_panel_rows"]);K=int(AUD["product_count"]);NVALID=int(AUD["actual_month_sku_nonmissing"])
COVER=100*NVALID/N
MISS=int(AUD["vol_missing_rows"]);ZERO=int(AUD["vol_zero_rows"])
M=MET[(1,"GraphHistGB")];B=MET[(1,"HistGB")];L=MET[(1,"Lag1")]
DELTA=100*(B["wmape"]-M["wmape"])
INTERVAL=CI.get("boot95",[None,None])
CI_TEXT=("（%.3f，%.3f）个百分点"%(100*INTERVAL[0],100*INTERVAL[1])) if None not in INTERVAL else "待计算"
if DELTA > 0:
    GSTR="降低%.3f个百分点"%DELTA
else:
    GSTR="增加%.3f个百分点"%(-DELTA)
SIG="置信区间跨越0，未显示稳定增益" if INTERVAL[0] is not None and INTERVAL[0]<=0<=INTERVAL[1] else "该探索性重采样区间未跨越0，但独立月份仅5个，需额外验证"
TITLE="卷烟多品规销量预测中的关联信息增益研究"
ENG_TITLE="Predictive value of inter-specification associations in cigarette sales forecasting"
DATE="2026-11-18"
VERSION="V1 · 真实数据先导实证稿 · 非投稿定稿"

paper=[]  # tuple: (kind, content [,extras])
def h1(s):paper.append(("h1",s))
def h2(s):paper.append(("h2",s))
def p(s):paper.append(("p",s.strip()))
def tab(cn,en,headers,rows,note=""):
    paper.append(("table",(cn,en,headers,rows,note)))
def fig(n,cn,en,fn,note=""):
    if (FIG/fn).exists():
        paper.append(("fig",(n,cn,en,fn,note)))

ABSTRACT=f"【目的】检验卷烟品规间历史变化的统计关联能否提供独立于单品规历史销量的预测信息，为选择时序预测与图关联建模技术提供实证依据。【方法】依据某市2025年1月至2026年7月的品规经营报表，重核原始销售汇总与历史面板，以{K}个品规、{N}条品规—月份记录构建观察数据，采用严格按月划分的训练、验证和未来测试集。以简单滞后预测、移动均值、岭回归及梯度提升树为基线，引入仅使用预测时点之前6个月变化计算的收缩相关图特征，比较1～3个月预测结果，并通过测试月份块式重采样分析增益稳定性。【结果】实际可用销量记录{NVALID}条，原表与面板成功匹配的{AUD['source_matched_rows']}条记录数值一致。未来1个月测试共{M['n']}条记录；上月销量直接预测的加权绝对百分比误差（wMAPE）为{100*L['wmape']:.2f}%，梯度提升树为{100*B['wmape']:.2f}%，加入图特征后为{100*M['wmape']:.2f}%，相对于未加图特征模型仅{GSTR}。其差异的探索性95%区间为{CI_TEXT}，{SIG}。【结论】现有19个月的月度数据尚未证明图关联特征具有稳健预测收益，也不支持将该结果解释为消费者品规替代。开展正式图神经网络研究前，应首先改善数据时间粒度、覆盖程度与供给状态记录。"
EN_ABS=f"[Objective] This study evaluates whether historical associations across cigarette specifications provide additional out-of-time predictive information beyond product-specific sales histories. [Methods] A one-city product-month panel from January 2025 to July 2026 was reconstructed and cross-checked against original XLS reports. The panel contained {K} specifications and {N} product-month records. A temporally constrained, shrunk correlation graph was used to augment gradient-boosted forecasts and compared against persistence, moving averages, ridge regression and the same boosted model without graph features. [Results] Of {N} panel records, {NVALID} had an observed non-missing volume. In the one-month holdout ({M['n']} matched observations), persistence yielded a weighted absolute percentage error of {100*L['wmape']:.2f}%, compared with {100*B['wmape']:.2f}% for the ordinary gradient-boosted baseline and {100*M['wmape']:.2f}% for its graph-augmented variant. The difference was too small to establish a reliable benefit under a five-month block bootstrap. [Conclusion] The current short, incomplete monthly history is insufficient to establish a reproducible gain from inter-specification graph features or any causal interpretation of product substitution. Longer and better documented records are needed before claiming advantages for graph neural networks."

# 0
h1("0 引言")
p("卷烟货源组织和市场运行监测具有显著的品规异质性。单一品规的月度销量既反映其既有经营规模，也受到节令、货源投放安排、价格与渠道承接等因素的共同影响。围绕销量的准确预测，不仅需要拟合序列随时间的变化，还需要判断跨品规信息是否能够为新的预测时点提供额外信号。尤其在品规间相互关联程度未知、品规上市和退出并存的情形下，直接把全部品规视为相互独立的预测对象，可能遗漏共同波动；但仅凭品规销量同步变化就推断消费替代关系，又缺乏足够的识别依据。")
p("针对卷烟市场复杂时间序列，蒋兴恒[1]采用集合经验模态分解、自适应参数优化与分项建模的集成方法研究品牌销量预测；崔建华等[2]提出以多层感知机为主、大语言模型辅助解释业务规则的预测控制体系，强调卷烟销量的周期性、节令性和策略性。上述成果分别从时间尺度分解和规则约束预测角度推进了研究。另一类研究通过零售客户分档、市场状态评价和数据驱动投放策略，将经营数据用于管理判断[3-5]。这些研究表明，卷烟商业数据中的业务口径和约束条件不能被一般消费品数据建模假设替代。")
p("图神经网络提供了跨序列关联学习工具。Graph WaveNet[6]通过自适应邻接学习发现未知关系，MTGNN[7]进一步结合图学习、图传播与扩张时间卷积开展多变量预测；图注意力网络（Graph Attention Network，GAT）[8]则通过注意力权重处理节点间信息传递。与此同时，时间融合Transformer（Temporal Fusion Transformer，TFT）[9]强调静态变量、历史观测和未来已知变量的有区别建模。然而，这些具有代表性的深度模型通常在时间覆盖较充分的序列上得到评估。对不足两年的品规月度面板，先判断关联信号能否超越简单预测，再选择复杂模型，比直接叠加注意力与图卷积更符合严格实验要求。")
p("本文以可核验的历史品规月度面板为对象，把研究问题收敛为：基于预测时点以前的多品规历史，能否通过关联信息提高下期已实现销量的样本外预测精度？围绕该问题，建立原始报表与历史面板的交叉核对程序，构造严格遵守信息时点的收缩关联图，并将图特征添加到相同基学习器中，通过强业务基线、未来月留出和消融比较检验关联增益。本文的定位是正式图神经网络论文之前的数据与机制先导实验，目的不是预设复杂方法一定优于简单方法，而是明确哪些技术贡献已被证实、哪些仍需进一步验证。")

# 1
h1("1 数据与研究问题")
h2("1.1 数据来源与研究对象")
p(f"研究数据来自某市卷烟经营管理报表，时间范围为2025年1月至2026年7月。原始资料按月份保存于“多指标销售汇总”“多规格按日查询”和“品牌月进货情况查询”3类工作簿，共57份旧版Excel文件。其中“多规格按日查询”是业务报表名称，原始表按品规列示本期经营指标，既有历史数据构建代码也将其聚合到品规—月份主键，因此本研究不将其误认成连续逐日交易明细。继承历史版本中的raw_panel_v1.csv，对品规编码、年月、销量字段进行独立核查。")
p(f"面板共{N}条品规—月份记录，覆盖{K}个品规、19个日历月；其中销量字段非缺失记录{NVALID}条（{COVER:.1f}%），缺失{MISS}条，显式记录的零销量{ZERO}条。研究在准备数据时区分缺失和明确记录的0值，不对缺失样本随意填充。对原始“多指标销售汇总”工作簿逐月读取品规编码和销量数值，将其与历史面板按“品规编码—月份”逐项核对：成功匹配的{AUD['source_matched_rows']}条记录差额均为0。由于历史构建时曾对品规至少出现6个月作过筛选，不能把上述面板直接等同于该市所有品规的完整总体。")
tab("表1 实验数据来源与可用性","Tab. 1 Sources and limitations of the experimental data",
["数据来源","业务粒度","核验内容","使用范围"],[
["多指标销售汇总（19份）","品规×月份","销量与历史面板数值核对","真实预测目标"],
["多规格按日查询（19份）","当前资料为品规×月份报表","22列业务指标；并未证实逐日序列","原始口径与辅助审计"],
["品牌月进货情况查询（19份）","品规×月份","库存、供给相关业务属性","留待约束条件核实"],
["历史raw_panel_v1.csv","品规×月份","203品规、3149记录、58字段","统一特征主表"]
],note="数据均为实际已有文件。源报表的销量计量单位仍需依据业务系统字段说明最终核准。")
fig("1","品规月度销量汇总变化","Monthly aggregate of observed product volume","Fig01_monthly_actual_aggregate.png","仅对有实际记录的品规求和；不同月份覆盖SKU数量不同，跨月波动不能直接解释为总体市场变化。")
fig("2","不同月份有效品规记录数","SKU coverage across observed months","Fig02_monthly_sku_coverage.png","浅缺失不自动推断为销售量为0。")
h2("1.2 预测目标和信息时点")
p("用i表示品规、t表示月份、h表示预测提前期，y_(i,t)为原报表记录的第i个品规在t月的已实现销量。在t月初发出预测时，只能使用t月之前已经公开或已完成记录的信息集I_(t−1)，对未来h个月的目标写为ŷ_(i,t+h−1)=f(I_(t−1),h)。若该值受供给限制，预测结果只能解释为已实现业务量，不能未经核验解释为潜在需求。")
p("对于预测任务，首先用本品规历史3个月销量构造滞后值、移动均值和波动指标，已知月份信息仅采用目标月的日历位置。由于历史报表未提供可靠的品规×日连续记录，不采用每天、每周或未来7天等与当前数据不一致的实验标签。年度季节基线另以相同月份上年销量进行对照，并明确其有效样本可能少于其他模型。")
h2("1.3 数据缺失及样本选择")
p("本研究不把原始面板中的缺失直接置0。只有预测时点之前连续3个月具有真实观测值、目标月份实际销量可得的样本进入相应预测窗口；因而测试集覆盖的品规是相对稳定的经营子集。对比实验使用固定目标月份和相同有效观测，在单独比较季节性12个月滞后基线时标明分母变化。缺失样本和新增品规可能与供给、上市和渠道条件相关，分析结论不外推至完全无历史的新上市品规。")
fig("3","品规月份数据有效性分布","Product-month observation matrix","Fig03_recorded_coverage_matrix.png","图中深色仅表示该月存在记录，不表示销量大小或消费者购买状态。")

# 2
h1("2 方法")
h2("2.1 整体技术路线")
p("为识别跨品规关联本身是否增加预测信息，本研究采用“历史面板审计—无图强基线—时点约束图学习—相同预测器消融—未来月份验证”的技术路线。其关键是保持预测器容量和输入时间条件基本一致，仅增加由历史其他品规产生的特征。与直接训练GAT、MTGNN或TFT相比，该先导设计更容易辨认增益来源，也能降低短时间轴下深度网络参数估计不稳定的风险。")
h2("2.2 仅使用预测前信息构造关联图")
p("对每一个预测原点t，取前6个可观察月份的各品规销量，计算对数变化r_(i,τ)=ln(1+y_(i,τ))−ln(1+y_(i,τ−1))。在同一时间窗内，对品规i和j的共同有效变化求Pearson相关系数ρ_(ij,t)；鉴于最多只有5个相邻月变化样本，相关系数极不稳定，采用非负截断与样本数收缩：a_(ij,t)=max(0,ρ_(ij,t))·n_(ij,t)/(n_(ij,t)+8)。保留满足最少4个重叠变化并超过预设阈值的候选关系，每个节点至多5个邻居；为了避免单向筛选导致稀疏图出现不对称不可靠连接，最终保留互相入选的边。")
p("此处的ρ不是严格意义上的替代弹性或需求交叉效应，只是预测时点之前历史涨跌的统计依赖。共同节令、配额、货源收缩和市场趋势都可能使两个品规的销量相关，无法据此将边权解释为消费者从一个品规迁入另一个品规。任何关于消费迁移的因果结论仍需消费者层级的连续购买数据和恰当识别设计。")
h2("2.3 跨品规历史特征与基学习器")
p("对图中的有效邻居j，计算预测原点前1期和前2期对数销量之差，并以a_(ij,t)归一加权得到邻域历史变化g_(i,t)。同时记录有效邻居数量d_(i,t)作为关系稀疏度指标。普通模型以历史3期对数销量、均值、波动与月份正余弦为输入；图增强模型在完全相同的基础上增加g_(i,t)和d_(i,t)。二者都使用相同的梯度提升树（Histogram-based Gradient Boosting，HistGB），以免把网络结构复杂度误当成图信息收益。")
p("为保持可复现性，树模型预先规定为85轮迭代、12个叶节点、最少叶样本26、学习率0.06和L2正则化8，固定随机种子42；目标使用ln(1+y)减少大销量品规的偏态影响，并对预测输出逆变换和零下限截断。岭回归采用标准化历史特征和固定正则系数50。本文并未训练原始MTGNN或GAT，因此将方法称为“图关联特征先导模型”，不声称实现了端到端图神经网络。")
h2("2.4 评价指标")
p("采用平均绝对误差（MAE）及加权绝对百分比误差（wMAPE）评价预测水平。MAE=Σ|y−ŷ|/n，wMAPE=Σ|y−ŷ|/Σ|y|，后者在所有有效测试点汇总后计算，而不直接平均单个品规百分比误差。另报告有符号预测偏差，用于判断复杂模型是否系统低估高销量品规。")
p("为比较关联增益，定义ΔwMAPE=wMAPE_(无图HistGB)−wMAPE_(图特征HistGB)；正值表示图增强模型改善。以2026年3—7月的目标月份为重采样单位进行500次块式Bootstrap，得到探索性95%区间。由于只有5个独立月份块，区间稳定性有限，不把跨品规的数百条记录当成数百个独立时间试验。")

# 3
h1("3 实验设计")
h2("3.1 时间划分与预测窗口")
p("根据原始面板的观测边界，训练目标限制在2025年12月及以前，2026年1—2月用于验证，2026年3—7月全部留作独立未来测试。历史窗口采用此前3个月构造滞后输入，额外利用截至预测原点的前6个月构造邻接图；模型分别估计h=1、2、3个月提前期。无论预测步长大小，图结构和所有特征均只能使用预测原点之前的有效销量，不允许在构造月度图时使用目标月或更晚月份的全样本数据。")
tab("表2 不同预测步长的实际有效样本","Tab. 2 Effective training and test observations by horizon",
["提前期/月","训练记录数","验证记录数","测试记录数","测试品规数"],[
[str(h),str(GROUP[h]["train"]),str(GROUP[h]["validation"]),str(GROUP[h]["test"]),str(GROUP[h]["num_skus_test"])] for h in (1,2,3)
],note="训练、验证和测试为品规—目标月份记录数；不同提前期的入选样本不完全一致，比较必须在相同提前期内进行。")
h2("3.2 对照模型与消融设置")
p("为避免只与弱基线比较，设置7组方案：Lag1（以上一期实现量直接预测）、Mean3（历史3个月平均）、Seasonal12（上一年同月实现量）、PeerLag（简单邻居增长率外推）、Ridge（仅本品规历史的岭回归）、HistGB（仅本品规历史的梯度提升树）和GraphHistGB（增加邻域变化与邻居数量的相同梯度提升树）。其中HistGB与GraphHistGB是预先指定的成对消融对照；Seasonal12使用单独可得的有效子样本，不能直接与全样本模型进行无条件排名。")
tab("表3 对照方法与主要检验目的","Tab. 3 Baseline methods and hypotheses",
["对照方法","信息来源","检验目的"],[
["Lag1","本品规上一月销量","市场惯性是否已经足够"],
["Mean3、Seasonal12","自身历史均值、上年同月","趋势平滑和年度季节性"],
["Ridge、HistGB","本品规历史与目标月位置","线性与非线性建模增益"],
["PeerLag","邻居历史变化外推","直接复制关联变化是否有效"],
["GraphHistGB","HistGB输入＋预测时点之前图特征","关联信息的边际贡献"]
],note="每项预测使用同一真实目标销量字段。未训练TFT、MTGNN、GAT，不能将图增强树模型成绩标为这些深度模型的性能。")
h2("3.3 数据审计与复现约束")
p("源文件均使用只读方式处理，记录原始文件数、工作表行列数、可读取品规编码和数字单元格数。通过sha之外的业务主键严格按“月—品规”匹配原报表与历史面板；发现重复键不进行自动加总。训练样本选择、图相关计算和数据变换在脚本中固定，结果保存为结构化JSON及CSV聚合表；科研图由同一脚本直接生成PNG和SVG，便于复核和期刊排版。")

# 4
h1("4 结果与分析")
h2("4.1 数据覆盖与来源一致性")
p(f"原始核验表明，某市19个月面板共有{K}个品规、{N}条记录，实际可用于销量统计的非缺失记录{NVALID}条，占{COVER:.1f}%；缺失销量{MISS}条，明确为0的有效销量记录{ZERO}条。在全部{AUD['source_matched_rows']}条成功与原始销售汇总报表关联的有效销量记录中，数值精确一致比例达到100.0%。这说明已有CSV并非仅有论文汇报数字，而具有可重新追溯的源表基础。但不同月份有效品规数存在明显变化，直接比较未对齐品规集合的月度总量会混入样本覆盖变化。")
p("图1和图2分别展示业务实现量的月度汇总及有效记录SKU数量。图3进一步将203个品规的19个月记录状态按覆盖率排序，可观察到部分品规长期断续缺报。样本覆盖不齐不仅影响时间序列平稳性判断，还会改变模型训练时的目标分布。由于原始业务系统中的停供、暂未上市和零购进状态尚无独立可复核标记，本研究暂不对这些缺报给出经营原因分类。")
h2("4.2 未来1个月预测效果")
p(f"在2026年3—7月独立测试期，能够同时构成有效历史3个月特征和目标值的共有{GROUP[1]['test']}条品规—月份记录，涉及{GROUP[1]['num_skus_test']}个品规。Lag1的wMAPE为{100*L['wmape']:.2f}%，MAE为{L['mae']:.3f}；仅使用本品规历史的HistGB为{100*B['wmape']:.2f}%、{B['mae']:.3f}，加入关联图历史特征后的GraphHistGB为{100*M['wmape']:.2f}%、{M['mae']:.3f}。就该真实测试集而言，最简单的上月保持基线明显优于两个树模型，说明对短月度面板而言，复杂非线性模型未必能充分利用额外自由度。")
p(f"在严格配对的{M['n']}条测试记录上，关联特征使HistGB的wMAPE相对变化仅为{GSTR}；按测试月份分块重采样得到95%差异区间{CI_TEXT}，{SIG}。与此相对，PeerLag把其他品规的增长变化直接用于外推，误差进一步放大，说明高相关历史涨跌并不能直接作为本品规下一期的预测增长率。这一负面发现必须与正面结果同等报告。")
h2("4.3 多步预测对比")
p("当预测提前期扩展到2个月、3个月时，不同方法的误差和可评价样本数发生变化。为避免从图中视觉选择结果，表4逐一给出相同目标定义和固定预测提前期下的真实计算指标。需要特别指出，Seasonal12模型只有存在上年同月有效销量时才能评价，有效点数与其余模型不同，因此其汇总误差仅作参考；表中的最低数值不应跨不同样本分母直接认定模型绝对优胜。")
rows=[]
model_order=["Lag1","Mean3","Seasonal12","Ridge","HistGB","GraphHistGB","PeerLag"]
for k in model_order:
    r=[k]
    for h in (1,2,3):
        d=MET.get((h,k))
        r.append(("%.2f%% (%d)"%(d["wmape"]*100,d["n"])) if d and d["wmape"] is not None else "—")
    rows.append(r)
tab("表4 未来1～3个月预测的wMAPE结果","Tab. 4 Weighted absolute percentage error across horizons",
["模型","提前1个月","提前2个月","提前3个月"],rows,
note="括号内为模型该提前期实际进入评价的记录数；Seasonal12等模型存在有效数据子集不同的问题。全表未使用未训练深度模型的伪造结果。")
fig("4","不同提前期的真实预测误差比较","Out-of-time errors at different horizons","Fig04_multi_horizon_wmape.png","时间原点和目标月划分均按独立测试期固定；跨样本子集方法不可直接根据柱线高低排名。")
h2("4.4 图关联结构及其预测价值")
p("图6以训练历史之前已发生的数据计算邻接矩阵，按历史销售规模选取匿名品规绘制稀疏关联结构。相关图采用非负截断、样本重叠数量收缩和互选Top-k，使得极短历史窗中的偶然高相关不至于完全主导邻接。然而，在测试结果中，图关联特征总体未形成稳定改善，且不同提前期下方向并不一致，因此不能将该结构解读为对消费者行为关系的可靠恢复。")
fig("5","代表性品规的实际与预测值比较","Observed and predicted volumes for preselected product strata","Fig05_actual_vs_predicted_3_preselected_skus.png","代表品规按照训练期销量分位预先选取，不按测试表现挑选最优案例。")
fig("6","训练历史构造的品规关联矩阵","Inter-specification association matrix based on historical data","Fig06_graph_weights_training_only.png","色阶为收缩后的历史统计相关强度，不代表消费替代、订单转移或因果效应。")
fig("7","未来月份无图与图增强模型的误差变化","Monthly error comparison with and without graph features","Fig07_monthly_horizon1_graph_ablation.png","对同一未来月份的真实目标计算wMAPE，实际独立月份仅5个。")
h2("4.5 成果与负面结果的业务含义")
p("在当前数据条件下，关联图并没有提供足以支撑自动化品规调控的稳定预测优势。更谨慎的应用方式是将其作为历史共同波动的探索性参考，与既有单品规历史基线同时显示，遇到缺报、上市、停供或节令变化时提示人工复核。由于不具备配套真实货源约束和授权人工决策日志，本研究不报告“提升投放准确率”“增加卷烟销量”或“节约库存成本”等未经实测的经营收益。")

# 5
h1("5 讨论")
h2("5.1 图特征失效的可能原因")
p("第一，时间轴短。19个月样本虽然覆盖203个品规，但估计月度关系的实际独立时间点不足两年。相关图仅用预测前6个月、5次相邻月变化形成，每个相关系数即使经过收缩仍具有较高随机性。第二，观测缺失不均衡。原始面板约四分之一记录未提供有效销量，图关系较容易由少数同月可用样本决定。第三，业务实现量受到供给安排约束。两品规共同出现销量下降，可能源于同期供给紧张，而非真实消费偏好趋同。第四，树模型与图特征使用了相同的滞后输入，自相关已能表达大部分预测信息，图特征额外增益空间较小。")
p("这些解释是与实际发现相容的可能机制，而不是目前已经识别的因果原因。要区分各机制，需要补充单日SKU出货流水、可供货量、缺货状态、门店扫码销售、正式投放规则发布时间和多地区样本，并设置相应的负控制及外部验证。只有在数据粒度和业务事件记录满足要求后，才能严谨比较TFT与MTGNN/GAT类图神经网络的真实优势。")
h2("5.2 与已有研究的增量及边界")
p("相较于蒋兴恒[1]针对品牌单序列开展分解集成预测和崔建华等[2]将规则辅助MLP预测用于供应链库存控制，本文关注“加入其他品规历史信息是否真的提升独立未来销量预测”这一更窄的识别问题。使用同一基础预测器开展成对有图/无图消融，有助于避免将模型复杂度和图关系本身混为一谈。然而，本研究在方法上只是关联图特征的先导检验，尚未提出可与MTGNN或GAT直接比较的端到端网络结构，也没有通过长序列外部验证。现阶段不具备将其作为“图神经网络方法创新已成立”的论文结论。")
h2("5.3 投稿前仍须完成的关键实验")
p("正式面向《中国烟草学报》继续推进时，需要完成三项补充工作：其一，获取至少连续多年的品规×周或品规×日数据，以及准确的进货、销量和库存业务口径；其二，复现原始TFT、MTGNN、Graph WaveNet与至少一个可靠图注意力基线，并以同等信息集、超参预算、跨地区独立测试检验增益；其三，以训练期构图、随机关系图负控制、分组消融和阶段性事件实验说明图结构何时有效、何时失效。若真实结果持续不优于单品规简单模型，应据实改变研究方向而不是仅替换网络名称。")

# 6
h1("6 结论")
p(f"（1）通过实际原始Excel工作簿及历史面板交叉校验，构建了{K}个品规、19个月、{N}条品规—月份记录的研究数据基础；可用销量{NVALID}条，成功匹配的原表记录数值一致，说明其可作为真实先导实证的基础。")
p(f"（2）未来1个月的独立测试中，上月销量持平基线wMAPE为{100*L['wmape']:.2f}%，优于普通HistGB的{100*B['wmape']:.2f}%及图特征增强HistGB的{100*M['wmape']:.2f}%；图关联信息仅{GSTR}，块式Bootstrap置信区间未显示可靠的改善。")
p("（3）当前短月度序列不能支撑“图关联显著提高预测效果”或“识别消费者替代关系”的判断。后续须结合更高频、更长时间跨度和更完整的业务约束数据，验证经典图神经网络与多步时序模型的外部有效性。")

h1("参考文献")
references=[
"[1] 蒋兴恒. 基于分解集成思想的复杂卷烟销量时间序列预测模型研究[J]. 中国烟草学报, 2025, 31(5): 138-145. DOI:10.16472/j.chinatobacco.2024.T0324.",
"[2] 崔建华, 陈颢, 刘冬荣. 基于MLP主导的卷烟销量预测模型构建及应用研究[J]. 中国烟草学报, 2026, 32(3): 133-140. DOI:10.16472/j.chinatobacco.2025.T0167.",
"[3] 谭琛, 等. 基于高维精准画像与深度学习的卷烟零售户动态分档研究[J]. 中国烟草学报, 2024, 30(3).",
"[4] 黄敏, 等. 基于智能集成与阈值训练的卷烟市场状态评价体系研究——以南充市为例[J]. 中国烟草学报, 2025, 31(5).",
"[5] 金文蔚, 等. 数据驱动的卷烟品规精准投放策略研究[J]. 中国烟草学报, 2025, 31(6).",
"[6] WU Z, PAN S, LONG G, et al. Graph WaveNet for deep spatial-temporal graph modeling[C]//Proceedings of IJCAI. 2019: 1907-1913.",
"[7] WU Z, PAN S, LONG G, et al. Connecting the dots: Multivariate time series forecasting with graph neural networks[C]//Proceedings of SIGKDD. 2020: 753-763.",
"[8] VELIČKOVIĆ P, CUCURULL G, CASANOVA A, et al. Graph attention networks[C]//International Conference on Learning Representations. 2018.",
"[9] LIM B, ARIK S Ö, LOEFF N, et al. Temporal fusion transformers for interpretable multi-horizon time series forecasting[J]. International Journal of Forecasting, 2021, 37(4): 1748-1764.",
"[10] HYNDMAN R J, KOEHLER A B. Another look at measures of forecast accuracy[J]. International Journal of Forecasting, 2006, 22(4): 679-688.",
]
for x in references:p(x)

# Supplementary statement (not journal text)
disclaimer=[
("已完成","原始19个月×三类报表57个XLS文件结构审计；原表销量与历史面板的全部成功匹配记录数值一致；7种月度预测基线、5个月未来测试、图表与聚合输出。"),
("尚未完成","没有独立连续日度交易明细；未训练TFT、MTGNN、GAT；未获得更长时段或跨地市外部验证；未建立正式供给计划因果效应实验。"),
("禁止替换","不得将HistGB+图特征改写成GAT、MTGNN、Graph WaveNet或TFT的真实测试结果；不得将销量相关图解释为消费迁移。"),
("数据保护","输出不包括原始客户名单、真实13位SKU编码、完整逐SKU样本矩阵或个人信息；图中案例使用匿名标签。")
]

def write_markdown():
    md=[f"# {TITLE}",f"\n{ENG_TITLE}\n",f"> {VERSION} | {DATE} | 作者及单位待填写\n",
        "## 摘要\n",ABSTRACT,"\n**关键词**：卷烟营销；多品规预测；图关系学习；时间外推；实证验证\n",
        "## Abstract\n",EN_ABS,"\n**Keywords**: cigarette marketing; product forecasting; graph association; temporal holdout; reproducibility\n"]
    for kind,val in paper:
        if kind=="h1":md.append("\n## "+val+"\n")
        elif kind=="h2":md.append("\n### "+val+"\n")
        elif kind=="p":md.append(val+"\n")
        elif kind=="table":
            cn,en,headers,rows,note=val
            md.append("\n**"+cn+"**  \n*"+en+"*\n")
            md.append("|"+"|".join(headers)+"|")
            md.append("|"+"|".join(["---"]*len(headers))+"|")
            md.extend("|"+"|".join(str(c).replace("|","/") for c in row)+"|" for row in rows)
            if note:md.append("注："+note+"\n")
        elif kind=="fig":
            n,cn,en,fn,note=val
            md.append(f"\n![图{n} {cn}](figures/{fn})\n")
            md.append(f"**图{n}** {cn}  \n*Fig. {n} {en}*")
            if note:md.append("\n注："+note+"\n")
    md.append("\n---\n## 研究边界与版本说明（投稿前删除）\n")
    for k,v in disclaimer:md.append(f"- **{k}**：{v}\n")
    md.append("\n数据与代码审计：results/01_data_audit.json、results/02_baseline_metrics.json和code/01_run_real_experiments.py。")
    (OUT/"论文V1_完整初稿.md").write_text("\n".join(md),encoding="utf-8")
    return md

def font_run(run,pt=14,bold=False,latin="Times New Roman",east="宋体"):
    run.bold=bold
    run.font.size=Pt(pt)
    run.font.name=latin
    pr=run._element.get_or_add_rPr()
    f=pr.rFonts
    if f is None:
        f=OxmlElement("w:rFonts");pr.insert(0,f)
    f.set(qn("w:eastAsia"),east)
    f.set(qn("w:ascii"),latin)
    f.set(qn("w:hAnsi"),latin)

def rule_table(table):
    tpr=table._tbl.tblPr
    borders=tpr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders=OxmlElement("w:tblBorders");tpr.append(borders)
    for typ in ("top","bottom"):
        x=OxmlElement("w:"+typ);x.set(qn("w:val"),"single")
        x.set(qn("w:sz"),"12");x.set(qn("w:color"),"233B51");borders.append(x)
    x=OxmlElement("w:insideH");x.set(qn("w:val"),"single");x.set(qn("w:sz"),"4");x.set(qn("w:color"),"DDE7EE");borders.append(x)
    for idx,row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for para in cell.paragraphs:
                para.paragraph_format.space_after=Pt(2)
                para.paragraph_format.space_before=Pt(2)
                para.paragraph_format.line_spacing=1.1
                for run in para.runs:font_run(run,9.5,idx==0)
    table.autofit=True

def add_paragraph(doc,text,style=None,fs=14,indent=True,align=WD_ALIGN_PARAGRAPH.JUSTIFY,space_after=Pt(4)):
    pp=doc.add_paragraph(style=style)
    pp.alignment=align
    pp.paragraph_format.line_spacing=1.5 if fs>=12 else 1.2
    pp.paragraph_format.space_after=space_after
    if indent:pp.paragraph_format.first_line_indent=Cm(.74)
    rr=pp.add_run(text);font_run(rr,pt=fs)
    return pp

def add_heading(doc,title,level=1):
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    pp.paragraph_format.keep_with_next=True
    pp.paragraph_format.space_before=Pt(12 if level==1 else 8)
    pp.paragraph_format.space_after=Pt(5)
    pp.paragraph_format.line_spacing=1.15
    r=pp.add_run(title)
    font_run(r,pt=14,bold=True,east="黑体")
    return pp

def add_table_doc(doc,title,en,cols,rows,note):
    x=doc.add_paragraph()
    x.alignment=WD_ALIGN_PARAGRAPH.CENTER
    x.paragraph_format.space_before=Pt(7)
    r=x.add_run(title);font_run(r,pt=11,bold=True)
    y=doc.add_paragraph()
    y.alignment=WD_ALIGN_PARAGRAPH.CENTER
    y.paragraph_format.space_after=Pt(7)
    r=y.add_run(en);font_run(r,pt=10.5,latin="Times New Roman",east="宋体")
    tbl=doc.add_table(rows=1,cols=len(cols))
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,c in enumerate(cols):tbl.rows[0].cells[i].text=str(c)
    for dat in rows:
        cells=tbl.add_row().cells
        for i,c in enumerate(dat):cells[i].text=str(c)
    rule_table(tbl)
    if note:add_paragraph(doc,"注："+note,fs=10.5,indent=False,align=WD_ALIGN_PARAGRAPH.LEFT)
    return tbl

def add_image(doc,n,cn,en,filename,note):
    file=FIG/filename
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.space_before=Pt(7)
    # Prevent oversized images, preserve aspect ratio.
    pp.add_run().add_picture(str(file),width=Cm(15.2))
    cap=doc.add_paragraph()
    cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.keep_with_next=True
    r=cap.add_run(f"图{n} {cn}");font_run(r,11)
    cap2=doc.add_paragraph()
    cap2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=cap2.add_run(f"Fig. {n} {en}");font_run(r,10.5)
    if note:add_paragraph(doc,"注："+note,fs=10.5,indent=False,align=WD_ALIGN_PARAGRAPH.LEFT)

def build_docx():
    # A fresh document using the specific style standards extracted from
    # the user's earlier journal template: 四号宋体, English Times, 1.5 spacing,
    # three-line tables, Chinese/English figure captions.
    doc=Document()
    sec=doc.sections[0]
    sec.page_height=Cm(29.7)
    sec.page_width=Cm(21.0)
    sec.top_margin=Cm(2.5)
    sec.bottom_margin=Cm(2.3)
    sec.left_margin=Cm(2.65)
    sec.right_margin=Cm(2.65)
    doc.styles["Normal"].paragraph_format.space_after=Pt(4)
    doc.styles["Normal"].font.size=Pt(14)
    doc.styles["Normal"].font.name="Times New Roman"
    cp=doc.core_properties
    cp.title=TITLE;cp.subject=VERSION
    cp.author="论文作者待确认"
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.space_after=Pt(7)
    r=pp.add_run(TITLE);font_run(r,18,True,east="黑体")
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=pp.add_run("[作者待确认]");font_run(r,12)
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=pp.add_run("[单位、部门、通讯地址及邮政编码待确认]");font_run(r,10.5)
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=pp.add_run("作者简介、通信作者及基金项目：投稿前据实填写");font_run(r,10,latin="Times New Roman")
    add_paragraph(doc,"摘要："+ABSTRACT,fs=11,indent=False)
    add_paragraph(doc,"关键词：卷烟营销；多品规预测；图关系学习；时间外推；实证验证",fs=11,indent=False)
    pp=doc.add_paragraph()
    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=pp.add_run(ENG_TITLE);font_run(r,13,True)
    add_paragraph(doc,"Abstract: "+EN_ABS,fs=10.5,indent=False)
    add_paragraph(doc,"Keywords: cigarette marketing; product forecasting; graph association; temporal holdout; reproducibility",fs=10.5,indent=False)
    in_references=False
    for kind,val in paper:
        if kind=="h1":
            add_heading(doc,val,1)
            if val=="参考文献":in_references=True
        elif kind=="h2":add_heading(doc,val,2)
        elif kind=="p":
            if in_references:
                add_paragraph(doc,val,fs=10.5,indent=False,
                              align=WD_ALIGN_PARAGRAPH.LEFT,space_after=Pt(2))
            else:add_paragraph(doc,val)
        elif kind=="table":add_table_doc(doc,*val)
        elif kind=="fig":add_image(doc,*val)
    # Footer page number field, per official journal guidance
    f=sec.footer.paragraphs[0];f.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=f.add_run("V1研究工作稿 · ");font_run(r,9.5)
    field=OxmlElement("w:fldSimple");field.set(qn("w:instr"),"PAGE")
    f._p.append(field)
    name=OUT/"论文V1_卷烟多品规关联信息增益研究.docx"
    doc.save(name)
    test=Document(name)
    with zipfile.ZipFile(name) as z:
        assert z.testzip() is None
        assert "word/document.xml" in z.namelist()
    assert len(test.paragraphs)>55 and len(test.tables)>=4
    assert len(test.inline_shapes)>=5
    return name,len(test.paragraphs),len(test.tables),len(test.inline_shapes)

def build_html():
    css="body{font:16px/1.9 'Noto Serif CJK SC','SimSun',serif;max-width:920px;margin:25px auto;padding:25px;color:#25394b;background:#fff}h1,h2{color:#225b7c}h1{font-size:27px}h2{font-size:21px;margin-top:30px}p{text-align:justify;text-indent:2em}table{border-collapse:collapse;width:100%;margin:12px 0;font-size:14px}th,td{border-top:1px solid #afbec8;border-bottom:1px solid #afbec8;padding:7px 11px;text-align:left}th{background:#eff5f8}img{display:block;max-width:100%;margin:auto}figure{margin:24px 0}figcaption{text-align:center;font-size:13px}small{color:#667f8c}.abstract{background:#f4f8fa;padding:12px 20px;border-left:4px solid #528aae}"
    o=['<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',f"<title>{html.escape(TITLE)}</title><style>{css}</style></head><body>",f"<h1>{html.escape(TITLE)}</h1>",f"<p style='text-align:center;text-indent:0'>{html.escape(ENG_TITLE)}</p>",f"<small>{html.escape(VERSION)} · 作者待确认</small>",f"<div class='abstract'><b>摘要：</b>{html.escape(ABSTRACT)}</div>",f"<p><b>关键词：</b>卷烟营销；多品规预测；图关系学习；时间外推；实证验证</p>",f"<div class='abstract'><b>Abstract:</b> {html.escape(EN_ABS)}</div>"]
    for kind,val in paper:
        if kind=="h1":o.append("<h2>"+html.escape(val)+"</h2>")
        elif kind=="h2":o.append("<h3>"+html.escape(val)+"</h3>")
        elif kind=="p":o.append("<p>"+html.escape(val)+"</p>")
        elif kind=="table":
            cn,en,hs,rows,note=val
            o.append(f"<p style='text-align:center;text-indent:0'><b>{html.escape(cn)}</b><br><small>{html.escape(en)}</small></p><table><thead><tr>"+''.join(f'<th>{html.escape(str(x))}</th>' for x in hs)+"</tr></thead><tbody>"+''.join('<tr>'+''.join(f'<td>{html.escape(str(v))}</td>' for v in r)+'</tr>' for r in rows)+"</tbody></table>")
            if note:o.append(f"<small>注：{html.escape(note)}</small>")
        elif kind=="fig":
            n,cn,en,fn,note=val
            o.append(f"<figure><img src='figures/{html.escape(fn)}'><figcaption>图{html.escape(n)} {html.escape(cn)}<br><small>Fig. {html.escape(n)} {html.escape(en)}</small></figcaption></figure>")
            if note:o.append(f"<small>{html.escape(note)}</small>")
    o.append("<hr><h3>研究边界与版本说明（投稿前删除）</h3>")
    for k,v in disclaimer:o.append(f"<p><b>{html.escape(k)}</b>：{html.escape(v)}</p>")
    o.append("</body></html>")
    (OUT/"论文V1_图文阅读版.html").write_text("\n".join(o),encoding="utf-8")

def readme():
    text=f"""# 1118/papers/v1 — 真实经营数据先导论文

## 文件
- 论文V1_卷烟多品规关联信息增益研究.docx：完整中文论文（含英文题名、结构化摘要、公式文字、三线表、双语图注）。
- 论文V1_完整初稿.md / 论文V1_图文阅读版.html：与Word对应，可在GitHub检索/阅读。
- code/01_run_real_experiments.py：从原始XLS和旧版raw_panel_v1.csv核对并计算7组先导实验；不使用未来信息构造图。
- code/02_build_paper.py：从真实结果JSON自动生成本文Word/Markdown/HTML。
- figures/Fig01—Fig07：真实计算的PNG和SVG（不含假想TFT/GAT准确率）。
- results：源Excel结构、数据审计、算法精度、月度测试汇总（不包含完整SKU预测明细）。

## 实证结论（真实运行）
- 产品：{K}个，月份：19，面板：{N}条。
- 非缺失销量：{NVALID}条，缺失：{MISS}条，显式零值：{ZERO}条。
- 原始销售XLS与面板成功匹配{AUD['source_matched_rows']}条，数值一致率100%。
- 一月提前期、5个月独立测试（{M['n']}条）：Lag1 wMAPE {100*L['wmape']:.2f}%，HistGB {100*B['wmape']:.2f}%，GraphHistGB {100*M['wmape']:.2f}%。
- GraphHistGB相对于HistGB仅{GSTR}，探索性块式Bootstrap区间为{CI_TEXT}。
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
"""
    (OUT/"README.md").write_text(text,encoding="utf-8")
    (OUT/"审稿自查_与因果森林旧版对照.md").write_text(
       """# V1与历史因果森林稿对照及本版退稿风险
1. 研究任务改变：旧版是非随机供给收缩的异质性处理效应（因果），V1是月度品规已实现销量的未来预测（非因果）。不能沿用ATE/CATE与缩投增效结论。
2. 实验数据沿用经源XLS匹配可审计的raw_panel_v1.csv；但不能将按月报表标题“多规格按日查询”视作每日销量流水。
3. 旧稿曾出现ATE显著性与负对照冲突的不同版本；V1主动保留不支持假设的结果，不挑选最好的一次种子或月份。
4. 主要当前缺口：只19个独立月份，部分SKU缺失严重，季节性基线样本不匹配，图模型未提供显著增益。
5. 新论文必须能回答“图特征是否真正增益”，结果目前是未证实，不能用网络可视化代替证据。
6. 如要投《中国烟草学报》，需真实实现TFT/MTGNN/GAT与Graph WaveNet的严格对照并补充分时/分地市外部数据；现稿只能作为V1方法与数据工作稿。
7. 原模板要点：题目一般不超过25字；中文四号宋体、英文Times New Roman、1.5倍行距；从0引言开始；三线表、中英文图表标题；结构化摘要与参考文献。
""",encoding="utf-8")

if __name__=="__main__":
    write_markdown()
    build_html()
    readme()
    out,npara,ntab,nfig=build_docx()
    print(json.dumps({"docx":str(out.relative_to(ROOT)),"size_bytes":out.stat().st_size,
                     "paragraphs":npara,"tables":ntab,"embedded_figures":nfig,
                     "real_h1_wmape_lag1":L["wmape"],"real_h1_wmape_graph":M["wmape"]},
                     ensure_ascii=False))
