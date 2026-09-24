#!/usr/bin/env python3
"""Screen and publish the 2026 Information, Communication & Society review."""

from __future__ import annotations

import csv
import os
import re
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data/raw/Information, Communication & Society-全部论文-2026-09-20.xlsx"
CROPPED = ROOT / "data/processed/Information, Communication & Society-2026-01-01至2026-09-22-裁剪.xlsx"
CURATED = ROOT / "outputs/ICS-2026-计算传播学精选.xlsx"
CURATED_CSV = ROOT / "outputs/ICS-2026-计算传播学精选.csv"
SCREENING_CSV = ROOT / "screening/ICS-2026-screening-log.csv"
INDEX_MD = ROOT / "reference/INDEX.md"
PRESENTATION_MD = ROOT / "outputs/ICS-2026-计算传播学精选-汇报版.md"

START_DATE = date(2026, 1, 1)
END_DATE = date(2026, 9, 22)
FIXED_XLSX_TIME = datetime(2026, 9, 22, 0, 0, 0)


# The DOI is the stable key.  Reasons record the positive method evidence seen
# in the supplied abstract, not a claim that the full text has been read.
INCLUDE = {
    "10.1080/1369118x.2026.2733507": ("主题建模、文本社群检测与时间比较", "计算文本分析"),
    "10.1080/1369118x.2026.2715573": ("Meta 广告库全集、投放支出与受众定向数据", "平台数字痕迹"),
    "10.1080/1369118x.2026.2713664": ("Facebook 捐赠数字痕迹，并含预注册题序实验", "数字痕迹＋在线实验"),
    "10.1080/1369118x.2026.2716778": ("LLM 分类、动态提及网络与图机器学习", "机器学习＋网络分析"),
    "10.1080/1369118x.2026.2695177": ("基于代理的社会模拟，操控网络同质性与选择性暴露", "计算模拟"),
    "10.1080/1369118x.2026.2696929": ("1200 万条 Facebook 帖文、LLM 分类、时间序列与回归", "大规模平台数据＋LLM"),
    "10.1080/1369118x.2026.2699262": ("在线评论的词嵌入与语义距离测量", "计算文本分析"),
    "10.1080/1369118x.2026.2691775": ("6829 篇新闻的主题建模与情感分析", "计算文本分析"),
    "10.1080/1369118x.2026.2685121": ("对 ChatGPT-4o mini 选举回答进行系统审计", "生成式 AI 审计"),
    "10.1080/1369118x.2026.2686314": ("Facebook 品牌帖文构成的四个二部网络", "网络分析"),
    "10.1080/1369118x.2026.2686315": ("2.29 万帖文、115 万评论与广义线性混合模型", "大规模平台数据建模"),
    "10.1080/1369118x.2026.2652497": ("Meta 广告与定向数据集、4760 条广告的多模态编码", "平台数据＋多模态分析"),
    "10.1080/1369118x.2026.2667914": ("2×2×2 的社交媒体假新闻在线实验", "在线实验"),
    "10.1080/1369118x.2026.2665253": ("代表性样本的被动浏览记录与恶意域名数据匹配", "被动数字痕迹"),
    "10.1080/1369118x.2026.2665252": ("3.8 万条 Reddit 评论的结构主题模型", "计算文本分析"),
    "10.1080/1369118x.2026.2659287": ("400 万条 X 帖文、转发网络、语言对齐与回音室分数", "网络分析＋计算文本"),
    "10.1080/1369118x.2026.2659279": ("340 万条微博与双重差分设计", "平台数据＋因果推断"),
    "10.1080/1369118x.2026.2654668": ("多层 Telegram 数据、纵向内容、URL 与网络分析", "多层平台数据分析"),
    "10.1080/1369118x.2026.2659280": ("TikTok 视频与评论的计算多模态分析", "计算多模态分析"),
    "10.1080/1369118x.2026.2647352": ("X 帖文的社会网络分析，并与访谈互证", "网络分析＋混合方法"),
    "10.1080/1369118x.2026.2648695": ("2500 万场在线比赛的个体数据与结构引力模型", "大规模行为数据建模"),
    "10.1080/1369118x.2026.2636138": ("145 篇全文的 BERTopic 建模", "计算型综述"),
    "10.1080/1369118x.2026.2655885": ("3879 条 Reddit 评论的主题建模", "计算文本分析"),
    "10.1080/1369118x.2026.2645882": ("Reddit 帖文与评论的事件时间线和时间序列分析", "平台数据＋时间序列"),
    "10.1080/1369118x.2026.2642851": ("语言、音频与主题特征的多模态时间序列", "计算多模态分析"),
    "10.1080/1369118x.2026.2642840": ("4131 条推文的隐喻频率与原创性数据化测量", "计算文本测量"),
    "10.1080/1369118x.2026.2633222": ("五种语言约 1500 万条推文的计算与质性分析", "大规模多语种计算文本"),
    "10.1080/1369118x.2026.2633216": ("三平台趋势数据与跨平台注意力效应估计", "跨平台数据建模"),
    "10.1080/1369118x.2026.2631709": ("69,421 个专利家族的主题聚类", "大规模计算文本"),
    "10.1080/1369118x.2026.2624702": ("在中英文预训练语言模型上实施 FMAT 偏见审计", "语言模型审计"),
    "10.1080/1369118x.2026.2623523": ("16 年 Reddit 跨版块用户、链接、转帖与评论重叠分析", "纵向网络分析"),
    "10.1080/1369118x.2026.2703856": ("移动经验抽样与真实政治广告截图的数据捐赠", "数据捐赠＋数字痕迹"),
    "10.1080/1369118x.2026.2678364": ("Python 抓取 5789 条 Reddit 评论，使用 Top2Vec 与 LLM 辅助主题整理", "计算文本分析＋LLM"),
    "10.1080/1369118x.2026.2669800": ("Discord 导出 1.33 万条申请，使用 R、正则表达式与随机抽样后开展主题分析", "计算采集与预处理＋质性分析"),
    "10.1080/1369118x.2026.2636134": ("预注册 3×3 在线实验，操控 Instagram 仿真帖文的敌意类型与提示方式", "在线实验"),
}


FULLTEXT_VERIFIED = {
    "10.1080/1369118x.2026.2733507",
    "10.1080/1369118x.2026.2715573",
    "10.1080/1369118x.2026.2713664",
    "10.1080/1369118x.2026.2716778",
    "10.1080/1369118x.2026.2695177",
    "10.1080/1369118x.2026.2696929",
    "10.1080/1369118x.2026.2699262",
    "10.1080/1369118x.2026.2691775",
    "10.1080/1369118x.2026.2685121",
    "10.1080/1369118x.2026.2686314",
    "10.1080/1369118x.2026.2686315",
    "10.1080/1369118x.2026.2652497",
    "10.1080/1369118x.2026.2667914",
    "10.1080/1369118x.2026.2665253",
    "10.1080/1369118x.2026.2665252",
    "10.1080/1369118x.2026.2659287",
    "10.1080/1369118x.2026.2659279",
    "10.1080/1369118x.2026.2654668",
    "10.1080/1369118x.2026.2659280",
    "10.1080/1369118x.2026.2647352",
    "10.1080/1369118x.2026.2648695",
    "10.1080/1369118x.2026.2636138",
    "10.1080/1369118x.2026.2655885",
    "10.1080/1369118x.2026.2645882",
    "10.1080/1369118x.2026.2642851",
    "10.1080/1369118x.2026.2642840",
    "10.1080/1369118x.2026.2633222",
    "10.1080/1369118x.2026.2633216",
    "10.1080/1369118x.2026.2631709",
    "10.1080/1369118x.2026.2624702",
    "10.1080/1369118x.2026.2623523",
    "10.1080/1369118x.2026.2729550",
    "10.1080/1369118x.2026.2713666",
    "10.1080/1369118x.2026.2663190",
    "10.1080/1369118x.2026.2713667",
    "10.1080/1369118x.2026.2703856",
    "10.1080/1369118x.2026.2674878",
    "10.1080/1369118x.2026.2667921",
    "10.1080/1369118x.2026.2678366",
    "10.1080/1369118x.2026.2678364",
    "10.1080/1369118x.2026.2669800",
    "10.1080/1369118x.2026.2636134",
    "10.1080/1369118x.2026.2715572",
    "10.1080/1369118x.2026.2713669",
    "10.1080/1369118x.2026.2674925",
    "10.1080/1369118x.2026.2681885",
    "10.1080/1369118x.2026.2678363",
    "10.1080/1369118x.2026.2683531",
    "10.1080/1369118x.2026.2659281",
    "10.1080/1369118x.2026.2652514",
    "10.1080/1369118x.2026.2645883",
    "10.1080/1369118x.2026.2636130",
    "10.1080/1369118x.2026.2613437",
}


EXCLUDE_OVERRIDES = {
    "10.1080/1369118x.2026.2729550": "全文显示使用九国 Generations and Gender Survey II 与加权多项 Logit；属于传统调查数据分析。",
    "10.1080/1369118x.2026.2713666": "全文显示使用 Ofcom 多年调查与 MAIHDA；没有平台数字痕迹或计算内容分析。",
    "10.1080/1369118x.2026.2663190": "全文显示为四国配额调查、潜在剖面分析与结构模型；属于传统调查研究。",
    "10.1080/1369118x.2026.2713667": "全文摘要显示使用 PATH 纵向调查，研究自然广告暴露；不是在线实验或计算方法。",
    "10.1080/1369118x.2026.2674878": "全文显示以少量 TikTok 视频及评论开展话语分析；爬虫只用于取数，核心分析为质性解释。",
    "10.1080/1369118x.2026.2667921": "全文通过人工汇编和系统编码比较 72 种数据访问工具；研究对象与计算方法有关，但本身未使用计算分析。",
    "10.1080/1369118x.2026.2678366": "全文显示为 23 次半结构化访谈与人工主题分析。",
    "10.1080/1369118x.2026.2715572": "全文显示为荷兰代表性样本的预注册横截面在线问卷（n=2,196），使用 EFA、CFA、SEM 和调节分析；没有计算传播方法或实验操控。",
    "10.1080/1369118x.2026.2713669": "全文以 Kiwi Farms 的公开帖文、平台文件与新闻材料重建平台迁移过程，采用质性案例分析；没有计算分析。",
    "10.1080/1369118x.2026.2674925": "全文显示为 Prolific 横截面在线问卷（n=375）与 PLS-SEM；没有实验操控或计算传播方法。",
    "10.1080/1369118x.2026.2681885": "全文显示为瑞典四波面板问卷（n=3,530），使用 RI-CLPM 与 LGCM；属于传统纵向调查分析。",
    "10.1080/1369118x.2026.2678363": "全文显示为两次相互关联的在线调查；第二次核验任务由人工按量规评分，分析使用 t 检验、相关和 OLS，不是随机实验。",
    "10.1080/1369118x.2026.2683531": "出版方页面标明该文为书评，不是经验研究论文。",
    "10.1080/1369118x.2026.2659281": "全文虽查询 Stack Exchange Data Dumps，但只用查询取得讨论串，核心材料缩减为 22 个线程并作溯因式质性分析；没有计算分析。",
    "10.1080/1369118x.2026.2652514": "全文使用 Wakoopa 数字踪迹制作个体访谈提示材料，核心证据来自 73 次访谈及反思式主题分析；踪迹没有进入计算建模。",
    "10.1080/1369118x.2026.2645883": "全文先人工核验 407 个短视频，再选取 22 个视频进行人工多模态编码和质性解释；没有自动化或计算多模态分析。",
    "10.1080/1369118x.2026.2636130": "全文为全国在线问卷（n=2,373），向受访者展示带标签的既有图片并询问是否见过；只随机了题项顺序，没有随机处理条件。",
    "10.1080/1369118x.2026.2613437": "全文确认为 2×4 随机组间在线情境实验；但研究涉及 AI 生成性影像虐待及色情网站情境，按用户本轮主题选择不收录。",
}


PENDING: dict[str, str] = {}


# These papers remain in the collection but sit on one boundary of the
# user-authorized project definition.  The labels are project-specific screening
# judgments, not claims about a universal disciplinary taxonomy.
MEDIUM_RELEVANCE = {
    "10.1080/1369118x.2026.2715573": "论文的资料来自 Meta 广告库 API，但六类核心内容变量由两名编码员人工标注，后续主要比较卡方统计量和 Cramér’s V；计算方法没有进入核心内容测量，因此列为方法边界。",
    "10.1080/1369118x.2026.2685121": "研究对 ChatGPT-4o mini 的选举回答进行审计，议题高度相关；但核心判断由研究者逐项核查事实并人工编码，计算环节主要是三元词筛除无法核验的陈述，因此列为方法边界。",
    "10.1080/1369118x.2026.2652497": "研究结合人工多模态内容编码、Meta 定向数据和广告库比较定位与定制；核心内容测量明确为人工编码，未显示计算分析进入内容测量，因此列为方法边界。",
    "10.1080/1369118x.2026.2703856": "研究对象是在线政治广告，但核心设计是移动经验抽样、截图捐赠、人工内容编码和一般统计模型；没有计算内容分析或平台行为建模，因此列为方法边界。",
    "10.1080/1369118x.2026.2669800": "使用 R 和正则表达式处理大规模 Discord 申请，但核心解释来自随机抽样后的主题分析；计算环节主要集中在采集与预处理。",
    "10.1080/1369118x.2026.2642840": "全文显示隐喻及其原创性由两名编码员按 MIP(VU) 人工识别，程序只计算派生指标并拟合一般统计模型；没有自动化文本识别，因此列为方法边界。",
    "10.1080/1369118x.2026.2631709": "使用主题聚类分析大规模专利文本，但核心问题偏向云基础设施与汽车创新，与传播研究的联系相对间接。",
}


METHOD_OVERRIDES = {
    "10.1080/1369118x.2026.2733507": "研究从欧洲委员会 Press Corner 收集 2016 年 1 月 1 日至 2024 年 6 月 30 日期间含“disinformation”的 238 篇官方演讲，在 R 中组建语料库。计算环节先构建词共现网络，用 igraph 的 Leiden 算法识别词汇社群；再将文本小写化、去停用词和领域停用词、去标点数字并词干化，用 LDA 提取 11 个主题。随后在 MAXQDA 中进行演绎与归纳并用的质性内容分析，同一编码者隔 20 天重新编码全部语料，稳定性 κ=0.95。",
    "10.1080/1369118x.2026.2715573": "研究用 Meta 广告库 API 收集 2022 年瑞典议会选举前四周八个议会政党投放的 2,395 条 Facebook 和 Instagram 广告，每条广告为分析单位。两名编码员先试编码，再人工标注主题、政策表达、价值表达、负面性、批评对象和行动号召；60 条复编码样本的 Cohen’s κ 为 0.70–0.85。分析同时使用 Meta 提供的跨平台发布、花费上界以及性别×六个年龄组的 12 项投放占比，主要用卡方检验和 Cramér’s V 比较政党差异。",
    "10.1080/1369118x.2026.2713664": "研究于 2023 年 2–6 月在匈牙利开展数据捐赠，最终得到 755 份完整问卷，并与同一批参与者通过 Facebook Data Download Packages 捐赠的发帖、分享、评论和反应记录配对。数字踪迹主分析覆盖问卷前 182 天，并以 30、60 和 365 天窗口做稳健性检验。发帖、分享和评论的政治性由 LLM 文本分类器判定，评论和反应另用 Facebook 页面类别与研究者自建政治页面名单交叉校验。作者将数字行为按 Guess 等人的规则转为有序类别，用 Spearman 相关和绝对一致性 ICC(2,1) 比较问卷与数字踪迹。另有预注册题序实验，随机分配参与者先回答总体行为题或具体内容题。",
    "10.1080/1369118x.2026.2716778": "研究回溯收集 2021 年 1 月 20 日至 2024 年 4 月的 X 移民议题数据，按重大移民事件与数据分布分成三个时段。作者建立三个有向 @提及网络，分别含 53,807、56,319 和 56,727 个节点以及 64,606、69,145 和 73,873 条唯一边。帖文的不文明得分由基于 BERT 的 Detoxify 模型产生；移民立场经人工与自动程序结合识别，并汇总为用户级立场。网络影响用开源软件计算，时间与立场差异用线性混合效应模型估计，组内外连结用缩放 E-I 指数测量。",
    "10.1080/1369118x.2026.2695177": "论文采用基于代理模型（ABM），以有界置信意见动力学为基础，在 NetLogo 中设置 2,500 个代理人，并用加泰罗尼亚 2011 年调查数据初始化国家认同。模型采用线上与线下两层网络，交叉操控随机／同质网络和过滤气泡开／关四种条件；每种条件重复模拟 100 次，运行至 400 个时间步，结果在 RStudio 中分析。",
    "10.1080/1369118x.2026.2696929": "研究用 CrowdTangle 收集 53 个群组和 4 个页面在 2021 年 1 月至 2023 年 12 月发布的 12,156,409 条亲博索纳罗 Facebook 帖文。作者用 love/angry 反应构建情感极化指数，用 shares/comments 构建参与平衡指数，再以 30 天滚动标准差超过第 95 百分位识别 96 个波动日和 7 个连续失稳期。后续聚焦失稳期的 1,161,126 条帖文，其中 712,287 条含文本。研究者人工标注 2,245 条帖文的六类政治行动者，按 80/20 划分训练与验证集，对 GPT-4o-2024-08-06 微调 3 轮，整体 F1=0.85。最后用两组多项 Logit 模型估计行动者提及与情绪、参与模式的关系。",
    "10.1080/1369118x.2026.2699262": "研究收集知乎唐山打人、徐州锁链女性和西安地铁拖拽三起事件发生后一个月内的帖文、回答和评论，清理重复、无关、垃圾和信息量不足内容后留下 288,359 条评论。作者先用 HanLP 做分词、词性与命名实体识别，再用 Qwen-Max-Latest 零样本分类人物指称是否为个人/群体及性别化/中性，从中得到 11,006 个指称词，1% 随机样本人工复核一致率为 95%。道德词来自经极性标注的 CMFD2.0。词嵌入以清华知乎 Word2Vec 为初始模型，用本研究 20 万余条评论增量训练 8 轮，窗口为 5、最低词频为 5，最终含 33,640 个 300 维向量。研究通过正负道德词向量差构造道德轴，以余弦相似度和欧氏距离测量指称词与道德框架的关联。",
    "10.1080/1369118x.2026.2691775": "研究从 LexisNexis 检索 2010 年 6 月至 2024 年 12 月的英文 AI 与政治新闻，初始得到 7,291 篇，排除 462 篇重复文章后保留来自 12 家主流媒体的 6,829 篇。文本用 LexisNexisTools 导入 R，统一日期与媒体名称；用 tidytext 分词、去停用词并保留关键二元短语，形成 5,119,881 个 token、132,054 个唯一词的文档—词项矩阵。LDA 通过 Gibbs 抽样估计（α=0.1，δ=0.1，5,000 次迭代，burn-in=1,000，seed=42），以拟合度、语义可解释性和分析连贯性确定 15 个主题。情感分析以 6,829 条标题为单位，并用 Bing、sentimentr、NRC 和 AFINN 四种词典交叉比较，南北差异用 Mann–Whitney 检验。",
    "10.1080/1369118x.2026.2685121": "研究围绕候选人、Project 2025、选举信息、选举错误信息和选举威胁设计 107 个提示词，由两名研究助理在 2024 年 9 月 10 日至 11 月 5 日每周两次访问未登录、无痕模式下的 ChatGPT-4o mini。本文取 10 月 25 日的 107 份回答，以三元词规则过滤无法核验的语句，研究者再人工补充筛除，最终得到 592 条可核验陈述。每条陈述由两名编码者依政府、竞选网站和独立新闻等来源判定准确性、错误类型与模糊限定；完整回答另人工编码严重虚构与关键遗漏，分歧经讨论达成一致。",
    "10.1080/1369118x.2026.2686314": "研究于 2024 年 5 月通过 CrowdTangle 收集 1,674 条含“renewable energy”的品牌化 Facebook 帖文/合作广告，样本涉及 850 个发布页面和 801 个付费或合作伙伴。作者在 R igraph 中分别构建伙伴—页面、伙伴—页面类别、伙伴类别—页面和伙伴类别—页面类别四个二部网络，用度数衡量合作数量，用 Louvain 方法识别凝聚子网络，并在 Gephi 中可视化。",
    "10.1080/1369118x.2026.2686315": "研究用性暴力相关中文关键词收集 2020 年 1 月 1 日至 2024 年 12 月 31 日的微博原创帖，筛除无图片、少于 10 个词或少于 10 条评论的帖文后留下 22,925 条帖文和 1,151,782 条评论。文本框架先用 LDA 探索，再以 1,800 条人工标注帖文训练中文 RoBERTa 分类器（交叉验证 precision/recall=0.929）。视觉框架以 VGG16 提取特征、PCA 降维、k-means 聚类为 18 类，再用 ViT 区分聊天截图。性别对立也由以 1,800 条人工编码评论训练的 RoBERTa 识别（precision=0.929，recall=0.925）。分析用 R glmmTMB 估计帖文随机效应的广义线性混合模型，并检验多模态框架、发布者性别与评论者性别的三重交互。",
    "10.1080/1369118x.2026.2678364": "研究用 Python 抓取 r/loseit 截至 2024 年 2 月提及 Mounjaro、Ozempic、Wegovy 或 Zepbound 的全部帖文与评论，去重后保留 5,789 条评论。Top2Vec 先将语料聚成 572 个主题；GPT-4o 只负责按研究者给定标准建议删除无关主题，所有取舍均由作者复核。清理后的材料再由研究者人工归纳，从 76 个初始簇合并为 18 个簇和 4 个主题。",
    "10.1080/1369118x.2026.2652497": "研究连接 Meta Ad Targeting Dataset 与 Meta 广告库，取得 2021 年德国联邦大选前四周的 4,775 条定向记录和 4,760 条广告。四名编码员对完整广告做人工多模态内容分析，并把地域、性别、年龄和议题拆成 17 个定位／定制类别；59 条复编码广告的 Krippendorff’s α 为 0.66–0.90。作者据此构造定位量、定制量与匹配量，再用带稳健标准误的截断 Poisson 模型解释展示次数，并控制花费、预估受众、账户、发布时间和广告形式。",
    "10.1080/1369118x.2026.2667914": "研究通过 Qualtrics 招募 271 名美国参与者，采用 2（真假）×2（来源可信度）×2（社会认同线索高低）的被试内在线实验。每人随机查看 8 条仿真 Twitter 新闻帖；真假文章与媒体来源分别经过两个 MTurk 预试，点赞、转发和评论数用于操控社会认同线索。结果变量是逐帖真假判断是否正确，分析先做 t 检验，再用二元 Logit 混合效应模型估计三类线索及平台有用性感知的作用。",
    "10.1080/1369118x.2026.2665253": "研究使用 2022 年 6 月 1,134 名美国成年人的 RealityMine 被动浏览记录，覆盖约 630 万次访问和约 64,000 个域名。作者把每个访问域名送入 VirusTotal API，至少被两家安全供应商判定恶意时才计为恶意域名，再按个人汇总访问次数和不同恶意域名数。群体差异通过分位数回归估计，并用 LOWESS 和 bootstrap 检查浏览强度能否解释人口差异。",
    "10.1080/1369118x.2026.2665252": "研究用 PRAW 收集 r/SexWorkers 与 r/OnlyFans 自 2020 年 3 月以来的 2,392 个帖文及 72,664 条评论。删除低互动帖、少于 10 词的评论、重复项和低相关内容后，保留 38,461 条评论。结构主题模型先用 Lee–Mimno 方法与 searchK 在 10–30 个主题间比较 held-out likelihood、残差、语义连贯性和 lower bound，最终确定 17 个主题；两名编码者命名主题，第三人裁决，并结合主题相关网络和每个主题随机抽取的 15 个帖文做定性解释。",
    "10.1080/1369118x.2026.2659287": "研究收集新冠疫情、俄乌战争和加沙战争三个议题下 400 多万条德语 X 帖文，并按周构建 12 个转发网络。作者先用 Leiden 划分社群，以 PageRank 识别关键账户，再分别计算 EchoGAE 网络同质性、ECS 语言一致性和二者合成的 ECI 回音室指标；跨时段用户与社群重叠用于判断同一批极化群体能否在不同危机中持续或重新聚合。",
    "10.1080/1369118x.2026.2659279": "研究建立 3,851 家中国 A 股公司 2007—2023 年面板，并抓取 31 个省级环保部门官方微博的 340 多万条帖文，共形成 32,369 个公司—年份观察。企业绿色转型用完整年报中 113 个环境词的词典频次测量；省级环保微博首次持续发帖被视为分期处理。主分析采用公司、省份和年份固定效应的分阶段 DID，并以事件研究、500 次安慰剂、Callaway–Sant’Anna CSDID、地形起伏度工具变量和中介模型做识别与稳健性检验。",
    "10.1080/1369118x.2026.2654668": "研究分析 2020 年 4 月至 2022 年 4 月德国 Querdenken 运动的 Telegram 数据。内容部分从 200 个频道的 630,966 条帖文中抽取每周浏览量最高的 40 条，共人工编码 4,160 条，再用 Sieve bootstrap 检验线性、单调和非单调趋势；URL 部分从 785,887 个链接中清理出 327,420 个外链并按域名统计。网络部分由转发关系滚雪球得到 6,008 个频道和 36,137,880 条帖文，经 disparity backbone 筛到 5,535 个节点、79,404 条边，再用 Louvain、HITS 和图结构指标刻画社群。",
    "10.1080/1369118x.2026.2659280": "研究从 TikTok Research API 收集 2023 年美国 #science 视频，先得 37,133 条，以 BERT 筛出 21,617 条科学视频；主分析使用 4,311 个账号发布、封面可识别身份的 10,800 条视频，并配对 448,002 条评论。DeepFace 识别封面中的感知性别与种族，GPT-4o-mini 识别言语攻击性和学科，两个 RoBERTa 模型测量情感强度与正式度，Perspective API 计算评论毒性，各模型均抽样人工验证。参与度用负二项回归，评论毒性用以账号为随机效应的线性混合模型。",
    "10.1080/1369118x.2026.2647352": "研究用 Brandwatch 收集 2023 年 11—12 月含“Spotify”和“Wrapped”的 22,283 条葡萄牙语 X 消息，其中 12,745 次转发用于构建账号—转发网络。作者按入度选出被转发至少 50 次的 41 条消息，这些消息覆盖近一万次转发，再与 2024 年 2 月至 2025 年 4 月对 38 名巴西 Spotify 用户的半结构化访谈合并做主题分析。网络只负责识别传播结构和高影响消息，四个解释主题来自三位作者对帖文与访谈的反复人工阅读。",
    "10.1080/1369118x.2026.2648695": "研究整合 aoe2insights.com 与 aoe2.net 的《帝国时代 II：决定版》完整对局记录，排除多人局后保留 24,040,177 场两人比赛，其中 8,098,734 场为玩家自主匹配，涉及 189 国、1,290,919 名玩家。作者构造 189×189×7×24 的国家对—星期—小时矩阵，共 6,001,128 个观察，以 CEPII 距离、共同语言、宗教、接壤和殖民关系以及联合国移民存量、Facebook Social Connectedness Index 解释互动量；主模型为含高维固定效应、按国家对聚类标准误的 PPML 结构引力模型。",
    "10.1080/1369118x.2026.2636138": "研究按预注册 PRISMA-ScR 流程检索 Scopus、Web of Science、Academic Search Premier、Communication Source，并用 Python scholarly 组合 7,840 个查询补充 Google Scholar。2,213 条记录去重后人工筛查 1,766 条，最终纳入 145 篇全文。作者人工编码平台、语言、内容模态和可供性，同时把去除参考文献的全文输入 BERTopic，以 all-MiniLM-L6-v2 生成嵌入、UMAP 降维、HDBSCAN 聚类（最小簇为 3），再人工命名主题。",
    "10.1080/1369118x.2026.2655885": "研究抓取 r/LovecraftCountry 与 r/Watchmen 在 2019 年 12 月至 2022 年 12 月的 7,601 个讨论串。作者比较 25 个 BERTopic 模型，在 HDBSCAN 最小簇 15–255 间选择连贯度 0.601、最小簇 45 的模型，得到 81 个主题；只保留种族相关主题概率最高的讨论串，最终分析 1,764 个串和 3,879 条评论。随后由首位作者在 NVivo 逐行编码，第二位研究者复核每三条评论中的一条，并用批判性技术文化话语分析解释角色与话语权力。",
    "10.1080/1369118x.2026.2645882": "研究从 Academic Torrents 的 Pushshift 档案取得 r/UAP（2014 年 9 月起）和 r/UFO（2017 年 10 月起）至 2024 年 12 月的 1,205,147 行帖文与评论，并整理 2018—2024 年 50 余个政府披露事件。Gemma 9B 与 Gemini 2.0 Flash 对 11,818 个 r/UAP 讨论串和 51,546 个 r/UFO 讨论串做七维分类，研究者再把标签人工合并为覆盖 91% 语料的 50 类；抽样 50 个串的人机一致性为 Krippendorff’s α=.701、κ=.69。日级序列用 CCF 检验披露事件的滞后关系，另由主题—用户关联矩阵投影出跨版主题网络。",
    "10.1080/1369118x.2026.2642851": "研究配对 2016—2018 年 InfoWars 每日销售额、931 期《Alex Jones Show》和 49,417 篇网站文章。节目经 Google Speech-to-Text 转写，Parselmouth 与 Librosa 提取 RMS 音量和 mel 频谱音高；文本分别计算八类 LIWC 词项，并以 STM 为文章和节目各建 15 个主题，最终语料为 45,917 篇文章和 930 份转录。49 个日级变量进入向量自回归模型，再用 Granger 因果检验判断内容特征与销售额的双向时序关系。",
    "10.1080/1369118x.2026.2636134": "研究于 2024 年 8 月在奥地利实施预注册的 3（不文明／不宽容／文明）×3（警觉提示／警觉加行动提示／无提示）实验；清理后从 594 人中保留 450 名非 LGBTQIA+ 旁观者。每人按组别随机查看 3 条仿真 Instagram 新闻帖，敌意评论和两步式社区提示均由研究者直接操控。可接受度、负面情绪、干预选择和审核偏好在每条刺激后测量，分析采用参与者随机效应、稳健极大似然估计的线性混合模型。",
    "10.1080/1369118x.2026.2633222": "研究用 Twitter Academic API 收集 2019 年 7 月至 2021 年 6 月含“价值观”对应词的五语帖文，约含英语 900 万、日语 500 万、德语 45 万、意大利语 30 万和韩语 20 万条。每种语言先对高频词建立共现矩阵并在 Visone 聚类，再以 Biterm Topic Model 从完整短文本中提取 20 个主题；2—5 元短语用 PMI 排序。十名母语编码者随后分析每种语言前 150 个短语，并对 30 个关键短语各抽取 30 条帖文，共精读 900 条。",
    "10.1080/1369118x.2026.2633216": "研究逐分钟抓取 2023 年微博、抖音和今日头条热榜，分别得到 124,751、53,214 和 120,020 个话题。跨平台匹配先用多语种 sentence-transformer 生成 512 维向量并计算余弦相似度，经人工编码 5,092 对候选话题后把阈值定为 0.8346，最终得到 39,253 对匹配和 54,539 个独立话题；政治话题由 GPT-4o 零样本分类。第二项研究对 196 个持续两小时以上的巴以战争微博话题形成 3,437 个 20 分钟观察，采用带话题、时段和小时固定效应的动态负二项模型估计跨平台上榜的滞后作用。",
    "10.1080/1369118x.2026.2624702": "研究用 FMAT 审计 BERT、RoBERTa、ALBERT 及四种中文预训练掩码语言模型，测试性别、职业地位和年龄三类关联。中文多字词采用伪对数似然：以等长 [MASK] 序列替换目标词，累加各 token 的对数概率并按长度归一；每个模板使用最小差异的 A/B 两版并取平均。统计单位为单条查询，分别做单样本 t 检验、方向符号检验和 2,000 次 bootstrap 置信区间。",
    "10.1080/1369118x.2026.2623523": "研究覆盖 11 个主要政治类 Reddit 子版块的 16 年帖文与评论，并分别计算发帖用户、评论用户、共享域名、具体 URL、跨版链接和话语身份的两两重叠。作者以各版活跃人数和贡献率建立期望重叠，再用观察／期望之比构成标准化相似度 NSS；话语差异则以 r/politics 为参照语料，在 AntConc 4.3.1 中用 log-likelihood 做 keyness 分析，最后回到原帖做定性解释。",
    "10.1080/1369118x.2026.2703856": "研究在德国 2021 年大选和美国 2022 年中期选举前各开展两周移动经验抽样，让面板成员上传实际遇到的政治广告截图并立即评分。德国先收 137 人的 611 张截图，筛得 55 人的 174 张政治广告；美国先收 530 人的 1,816 张，筛得 119 人的 419 张。广告政党与议题由人工编码并和选前问卷匹配；广告评价经主轴因子分析形成三项指标，随后分别用含受访者固定效应及聚类标准误的 OLS、以受访者为随机效应的因果中介模型和 SEM 稳健性检验。",
    "10.1080/1369118x.2026.2669800": "研究从四个公开 Discord 社群导出 2021 年 6 月至 2023 年 9 月的 Axie Infinity 奖学金申请，用 R 去重，并用正则表达式识别条目、年龄和国家。删除少于 50 词的消息后保留 13,329 份申请，再按各频道原分布随机抽取 10%，即 1,333 条。核心分析是在 NVivo 中先归纳编码、再演绎复核的人工主题分析；Discord 导出、正则表达式与随机抽样只承担材料整理工作。",
    "10.1080/1369118x.2026.2642840": "研究通过 Apify 收集 2023 年 3 月 18 日至 5 月 28 日两名总统候选人和五个主要政党账号的 4,202 条 X 帖文，清理后保留 4,131 条。两名土耳其语编码员按 MIP(VU) 逐条人工识别隐喻并判断原创性，400 条复编码样本的 κ 与 Krippendorff’s α 均为 .81；程序只据人工标签计算隐喻频率和原创性。点赞、转发和回复用带政治人物随机效应的负二项 GLMM 分析，并控制文本长度、标签、视觉内容和离选举日的时间。",
    "10.1080/1369118x.2026.2631709": "研究用 Lens 检索联网与自动驾驶汽车专利，从 168,033 条结果经扩展专利家族去重得到 69,421 个专利家族。作者把专利摘要输入 BERTopic，以 all-MiniLM-L6-v2 生成语义嵌入并聚类；人工审查后选择车辆通信 375 件、机器视觉 470 件、网络架构 1,264 件和边缘计算 765 件四个簇。每簇再统计主要申请人，并对代表性专利的全文和图示做定性分析。",
}


METHOD_LABELS = {
    "10.1080/1369118x.2026.2733507": "Leiden 词共现社群检测、LDA 主题模型与质性内容分析",
    "10.1080/1369118x.2026.2715573": "人工定量内容分析、卡方检验与 Cramér’s V",
    "10.1080/1369118x.2026.2713664": "LLM 内容分类、数据捐赠效度校验、Spearman 相关与 ICC(2,1)",
    "10.1080/1369118x.2026.2716778": "Detoxify 分类、动态提及网络、线性混合效应模型与 E-I 指数",
    "10.1080/1369118x.2026.2695177": "NetLogo 基于代理模型与 2×2 条件模拟",
    "10.1080/1369118x.2026.2696929": "GPT-4o 微调多标签分类、滚动窗口波动检测与多项 Logit",
    "10.1080/1369118x.2026.2699262": "HanLP 实体抽取、Qwen 分类、Word2Vec 增量训练与道德轴投影",
    "10.1080/1369118x.2026.2691775": "Gibbs 抽样 LDA 主题模型、四词典情感分析与 Mann–Whitney 检验",
    "10.1080/1369118x.2026.2685121": "三元词筛选与双人事实核验/内容编码",
    "10.1080/1369118x.2026.2686314": "二部网络、度数分析与 Louvain 社群检测",
    "10.1080/1369118x.2026.2686315": "RoBERTa 文本分类、VGG16/PCA/k-means/ViT 视觉识别与 GLMM",
    "10.1080/1369118x.2026.2678364": "Top2Vec 主题建模、GPT-4o 辅助筛题与人工主题分析",
    "10.1080/1369118x.2026.2652497": "人工多模态内容分析、定位—定制匹配测量与截断 Poisson 回归",
    "10.1080/1369118x.2026.2667914": "2×2×2 被试内随机在线实验与二元 Logit 混合模型",
    "10.1080/1369118x.2026.2665253": "RealityMine 被动浏览测量、VirusTotal 多引擎匹配与分位数回归",
    "10.1080/1369118x.2026.2665252": "结构主题模型、searchK 模型选择与定性内容分析",
    "10.1080/1369118x.2026.2659287": "Leiden 转发社群、EchoGAE/ECS/ECI 回音室测量与 PageRank",
    "10.1080/1369118x.2026.2659279": "年报词典测量、分阶段 DID、CSDID 与工具变量检验",
    "10.1080/1369118x.2026.2654668": "Sieve-bootstrap 趋势检验、URL 域分析、disparity backbone、Louvain 与 HITS",
    "10.1080/1369118x.2026.2659280": "BERT 筛选、DeepFace 身份识别、GPT/RoBERTa/Perspective 内容测量与混合模型",
    "10.1080/1369118x.2026.2647352": "转发网络入度分析、半结构化访谈与人工主题分析",
    "10.1080/1369118x.2026.2648695": "PPML 结构引力模型与高维固定效应",
    "10.1080/1369118x.2026.2636138": "PRISMA-ScR、all-MiniLM-L6-v2/UMAP/HDBSCAN BERTopic 与人工编码",
    "10.1080/1369118x.2026.2655885": "BERTopic 模型比较与批判性技术文化话语分析",
    "10.1080/1369118x.2026.2645882": "Gemma/Gemini 主题分类、交叉相关时间序列与二部网络投影",
    "10.1080/1369118x.2026.2642851": "LIWC、Parselmouth/Librosa 音频测量、STM、VAR 与 Granger 检验",
    "10.1080/1369118x.2026.2636134": "预注册 3×3 混合设计在线实验与线性混合模型",
    "10.1080/1369118x.2026.2633222": "词共现网络、Biterm Topic Model、PMI 短语排序与多语种定性编码",
    "10.1080/1369118x.2026.2633216": "sentence-transformer 话题匹配、GPT-4o 分类与动态固定效应负二项模型",
    "10.1080/1369118x.2026.2624702": "FMAT、中文多 token 伪对数似然与 bootstrap 偏见检验",
    "10.1080/1369118x.2026.2623523": "标准化重叠相似度（NSS）、log-likelihood 关键词分析与定性回读",
    "10.1080/1369118x.2026.2703856": "移动经验抽样、广告截图数据捐赠、主轴因子分析与因果中介模型",
    "10.1080/1369118x.2026.2669800": "正则表达式辅助整理、分层随机抽样与人工主题分析",
    "10.1080/1369118x.2026.2642840": "MIP(VU) 人工隐喻编码与负二项 GLMM",
    "10.1080/1369118x.2026.2631709": "all-MiniLM-L6-v2/BERTopic 专利聚类与代表性专利精读",
}


RECOMMENDATION_OVERRIDES = {
    "10.1080/1369118x.2026.2733507": "值得关注的地方在于，论文没有把“虚假信息”当作现成标签，而是把政策机构如何定义问题本身变成研究对象。计算分析显示，欧盟委员会长期借助战争和传染病隐喻，将责任推向敌对国家、平台和易受骗的公民；到 2023—2024 年，AI 又被纳入放大风险的叙事。它适合用于讨论问题框架如何预先限定平台治理、内容清除和媒介素养等政策答案。",
    "10.1080/1369118x.2026.2715573": "这篇文章把政治传播研究中常被忽略的付费广告置于中心，并同时观察广告说了什么、花了多少钱、投给了谁。结果显示瑞典左右翼政党在负面竞选和人口定向上采取不同策略，也说明平台的付费功能并非中性的传播通道，而会进入政党的受众选择与资源配置。",
    "10.1080/1369118x.2026.2713664": "这是一篇很有用的方法校准研究。它直接拿问卷回答与同一人的 Facebook 行为记录比较，发现用户往往低估一般使用、又高估政治参与，而且题目从总体到具体的排列更可靠。对依赖自报社交媒体行为的研究来说，这些结果能具体说明测量误差会出现在哪里，也展示了数据捐赠怎样用于验证传统问卷。",
    "10.1080/1369118x.2026.2716778": "论文把立场识别、不文明语言测量和互动网络连在一起，研究的不只是“哪一方更不文明”，而是不文明表达怎样沿着提及关系扩散并跨越意识形态边界。它为冲突螺旋理论提供了可追踪的动态证据，也适合作为大型语言模型与网络分析组合使用的案例。",
    "10.1080/1369118x.2026.2695177": "这篇文章适合用来检验一个经常被过度简化的判断：过滤算法是否会自动制造并固化回音室。模型把网络同质性与选择性暴露分开操控，结果显示信息暴露确实会影响回音室的形成，但人为形成的回音室即使处在过滤算法下也未必长期维持。它的价值在于把“算法—网络结构—意见变化”拆成可重复检验的条件；作者还在 OSF 公开了模型、输入和输出数据。",
    "10.1080/1369118x.2026.2696929": "文章用三年纵向数据挑战了“高度党派化社群始终按固定方式运作”的想象。卢拉就职和 2023 年 1 月事件之后，这些亲博索纳罗社群的情绪强度下降、内部讨论增加，说明政治事件会重新组织看似封闭的回音室。两个行为指标也为长期比较情感反应与互动方式提供了可复用的测量思路。",
    "10.1080/1369118x.2026.2699262": "它把“如何称呼事件中的人”与“如何对他们作道德判断”放进同一计算框架。研究发现，性别化称谓更容易连接广泛而两极化的道德基础，性别中立称谓则更偏向程序和角色评价。这种设计能把网络性别争论中的道德边界从抽象概念转成可比较的语义关系。",
    "10.1080/1369118x.2026.2691775": "文章兼有长时段、跨地区与可比较的媒体样本，能看到 AI 报道从技术想象走向政治化的过程。剑桥分析事件和 ChatGPT 发布构成两个转折点，全球北方媒体整体更悲观、南方媒体更乐观。它提供了一幅地缘政治、风险认知与 AI 媒体框架相互作用的经验地图。",
    "10.1080/1369118x.2026.2685121": "这篇文章的贡献不只在于列出 ChatGPT 的错误，而是提出“生成性治理”这一观察角度：平台已经从筛选人类内容转向直接生产用户看到的政治信息。选举期间出现的不准确、遗漏和回避性回答因此被放进平台权力与问责问题中考察，很适合作为生成式 AI 政治传播审计的案例。",
    "10.1080/1369118x.2026.2686314": "论文抓住了平台界面中一种容易被忽略的关系：品牌帖文背后的合作伙伴与发布页面如何结成战略网络。四个二部网络让这些近乎不可见的合作变得可观察，也揭示新闻媒体等传统机构在平台盈利逻辑下如何重新安排角色。它对研究平台化传播中的组织关系很有启发。",
    "10.1080/1369118x.2026.2686315": "文章把帖文的文字与视觉框架、发布者身份和评论反应放进同一模型。结果显示，煽动性和个性化框架会加剧性别对立，外群体身份发布时更明显，而娱乐性框架反而减弱对立。它把“性别冲突为什么会病毒式扩散”落实到了可检验的内容形式与身份线索上。",
    "10.1080/1369118x.2026.2678364": "这篇文章呈现了肥胖管理药物讨论中并存的四组张力：污名、获取不平等、品牌与错误信息混淆，以及对疗效和风险的怀疑。它的价值在于把主题发现与网络社群中的意义协商结合起来，既保留大规模文本的轮廓，也没有把健康传播简化成词频排名。",
    "10.1080/1369118x.2026.2652497": "论文先澄清“把广告投给谁”和“为受众改写什么”是两个不同概念，再用德国大选广告检验二者的实际效果。结果显示政党很少同时使用定位和定制，且只有定位显著影响展示次数。这一操作化框架能帮助后续研究避免把两种数据驱动竞选策略混为一谈。",
    "10.1080/1369118x.2026.2667914": "实验发现参与者有明显的“真实性偏见”，识别真新闻比识别假新闻容易；高可信来源有助于确认真新闻，却可能降低对假新闻的警觉。社会认同线索没有显著作用，而越认为 Twitter 有助于获取信息的人，识别假新闻反而越不准确。它为平台线索如何改变判断提供了直接的因果比较。",
    "10.1080/1369118x.2026.2665253": "文章把数字不平等从接入和技能推进到在线风险暴露。超过一半的参与者在一个月内访问过恶意网站，但暴露高度集中；人口差异在控制浏览强度后明显减弱。这个结果提醒研究者，群体差异可能来自不同的在线活动路径，不能只停在人口属性比较。",
    "10.1080/1369118x.2026.2665252": "这篇研究没有把 Reddit 简单描述成污名传播空间。结构主题模型与质性分析共同显示，平台一面延续对性工作者的负面评价，一面又承担互助、减害和实践知识生产的功能。它适合用来讨论数字社群如何同时产生伤害与支持。",
    "10.1080/1369118x.2026.2659287": "最值得注意的是，极化社群不仅会在一次危机中持续数周，还会在主题完全不同的新危机中重新出现。少数意见领袖、记者、非主流媒体和活动者通过重新激活熟悉叙事维持群体。这个发现把解释重点从单次事件或单一算法移向反复出现的参与者网络。",
    "10.1080/1369118x.2026.2659279": "文章把平台传播与企业的实际绿色行动相连，而不只测量线上声量。结果显示平台曝光与环境投资、披露质量上升相关，原创性、多媒体丰富度和社会影响力会强化这种作用。它提供了一个观察数字平台如何通过政策信号和声誉压力介入组织决策的案例。",
    "10.1080/1369118x.2026.2654668": "这篇论文用内容、来源和关系网络三层证据说明抗议运动如何降低接触激进观点的门槛。Querdenken 的讨论逐渐围绕反精英和阴谋论展开，链接更多替代性、党派性来源，并在网络上靠近极右翼与阴谋论群组。多种分析指向同一过程，使“主流化”不再只是修辞判断。",
    "10.1080/1369118x.2026.2659280": "文章在大规模短视频资料中同时处理身份交叉与表达风格。结果显示有色人种女性科学传播者面临更明显的角色冲突，身份与风格符合受众预期时参与度更高。它能帮助解释 TikTok 的注意力分配为何不仅取决于内容质量，也受身份化期待影响。",
    "10.1080/1369118x.2026.2647352": "论文把 Spotify Wrapped 看作一年一度的算法事件。网络分析显示它制造的是短暂、松散的共同感，而访谈进一步揭示用户会主动“训练”或规避算法，希望年度画像只记录值得展示的音乐。它把算法想象、身份表演和日常数据实践连到了具体的平台仪式。",
    "10.1080/1369118x.2026.2648695": "即使游戏中的互动几乎没有物理运输成本，文化相似性和既有社会联系仍明显影响跨国玩家是否相遇。这个结果用 2,500 万场真实互动说明，数字空间并没有抹平地理和社会边界。它为数字全球化和在线社群研究提供了一个少见的跨国行为数据基准。",
    "10.1080/1369118x.2026.2636138": "这不是只按人工分类汇总的范围综述。BERTopic 与人工编码共同揭示现有研究过度集中于 X、文本内容和检测模型，而视觉、音视频及平台功能受到的关注不足。它既能快速把握仇恨言论研究版图，也清楚指出下一步值得补的经验空白。",
    "10.1080/1369118x.2026.2655885": "文章关注流行文化作品如何触发关于种族历史的公共讨论，并追踪用户在支持、反对和修正立场之间的移动。Reddit 的匿名性既影响意见领袖的形成，也给部分用户改变看法留下空间。计算主题发现与批判性话语分析的组合，让研究能够同时看到总体结构和角色转换。",
    "10.1080/1369118x.2026.2645882": "研究把 Reddit 讨论与一系列政府披露事件对齐，发现透明度举措能够缓解由阴谋叙事引发的信任危机。它的意义在于不只描述阴谋论内容，而是利用事件时序观察官方信息介入后的变化，也提示“披露”本身带有政治含义，不能被当作单纯的信息增量。",
    "10.1080/1369118x.2026.2642851": "这篇文章把虚假信息内容与商业收益直接连在一起。节目中的脏话、音高、音量和特定主题能够预测次日销售，销售变化又会反过来预测后续内容风格。双向关系让“错误信息为何持续生产”从抽象的注意力激励落到可观察的经济反馈。",
    "10.1080/1369118x.2026.2636134": "实验结果对“加一个提示就能改善评论环境”的直觉提出了限制。敌意类型稳定影响情绪、干预意愿和审核偏好，但提示并未提高旁观者干预，有时还增加负面情绪。它为平台助推设计提供了一个重要的无效或反作用案例，也符合本项目纳入在线传播实验的口径。",
    "10.1080/1369118x.2026.2633222": "文章利用五种语言的大规模推文说明，同一个看似普遍的关键词在不同文化中承担不同功能：欧美语言中更常划分政治群体边界，日语和韩语中则更多调节个人偏好与关系。它把跨文化差异与共同的话语规律放在同一设计中，适合用于讨论计算文本研究如何避免英语中心。",
    "10.1080/1369118x.2026.2633216": "文章挑战了“一个平台上榜会带动其他平台同步升温”的直觉。在部分政治议题上，其他平台的热榜反而压低微博对同一新闻的注意力。这个负向溢出结果把跨平台关系从简单协同改写为注意力竞争，也为研究平台生态提供了可检验的框架。",
    "10.1080/1369118x.2026.2624702": "研究发现中文模型在性别评价上表现出更强的男性偏好，中英文模型都偏向高地位职业和年轻群体，其中地位偏差最稳定。它把语言模型中的社会不平等落到可重复的关联测试上，也提醒教育、招聘和内容审核中的模型应用可能复制既有等级。",
    "10.1080/1369118x.2026.2623523": "文章同时检查六类跨版块重叠，发现即便话语内容相近，政治子版块之间的用户、链接和互动交叉仍非常有限。这个长时段结果说明 Reddit 极化不能只归因于内容、身份或平台设计中的单一因素，也直接挑战了把 r/politics 想象成跨立场公共广场的说法。",
    "10.1080/1369118x.2026.2703856": "研究的亮点是把真实收到的政治广告与受众当下的主观评价配对。美国和德国样本都显示，党派一致会增强“这是为我定制的”感受，而这种感受又带来更积极的广告评价。它对理解个性化广告的心理过程很有价值；不过核心分析仍以经验抽样、数据捐赠和一般统计为主，因此在本目录中列为中等相关。",
    "10.1080/1369118x.2026.2669800": "论文提出“自我资产化”来解释玩家如何把经济困境、忠诚和自我剥削包装成可投资价值。这个概念对理解平台化游戏劳动很有启发，但计算工具主要用于整理和抽样 Discord 材料，核心发现来自质性主题分析，所以更适合作为计算传播的边界案例。",
    "10.1080/1369118x.2026.2642840": "结果区分了隐喻“用得多”和“用得新”两种效果：频率只带来约 1%—2% 的参与增长，原创性则使点赞和转发提高约 6%—7%，对回复的影响较弱。它为政治语言如何影响平台参与提供了量化证据；但摘要没有清楚说明隐喻识别是否自动化，因此列为中等相关。",
    "10.1080/1369118x.2026.2631709": "文章从大规模专利中识别车辆通信、机器视觉、网络架构和边缘计算四个主题，借此说明云不只是后台设施，也承载了 AI 规模化的产业想象。计算文本部分扎实，但研究对象更接近基础设施与创新研究，和传播过程的联系较间接，因此作为中等相关论文保留。",
}


PUBLIC_SOURCE_LINKS = {
    "10.1080/1369118x.2026.2695177": (
        "开放材料",
        "OSF：模拟模型、输入数据与输出数据（CC BY 4.0）",
        "https://osf.io/tukv7/overview",
    ),
}


ABSTRACT_OVERRIDES = {
    "10.1080/1369118x.2026.2678364": {
        "source": "R Discovery 公开论文元数据",
        "english": "Recent glucagon-like peptide-1 receptor agonists, commonly known as Obesity Management Medications (OMMs), have disrupted clinical and public conversations about obesity. Although clinical outcomes are well studied, less is known about how online discussions influence acceptance, skepticism, and stigma. Using a netnographic approach with computational text analysis, we analysed Reddit discussions and identified four recurring themes: stigma and social perceptions of being overweight; privilege and access to treatment; online misinformation and brand confusion; and skepticism and perceived risks. We show how portrayals of OMMs as ‘miracle solutions’ coexist with concerns about side effects, relapse, and unequal access, and how these tensions circulate through peer interactions and media narratives. The study contributes to digital health communication by explaining how publics negotiate meaning, assess credibility, and form perceptions in online environments, with implications for governing online health information and reducing harm.",
        "chinese": "新近出现的胰高血糖素样肽-1受体激动剂，通常被称为肥胖管理药物（OMMs），改变了围绕肥胖的临床与公共讨论。尽管临床疗效已得到广泛研究，人们仍不甚了解在线讨论如何影响接受、怀疑与污名。我们采用结合计算文本分析的网络民族志方法，分析 Reddit 讨论并识别出四个反复出现的主题：超重污名与社会认知、治疗特权与获取、在线错误信息与品牌混淆，以及怀疑与感知风险。研究显示，将 OMMs 描绘为“神奇解决方案”的叙事，与对副作用、复胖和获取不平等的担忧同时存在，这些张力通过同伴互动和媒体叙事流通。本研究解释公众如何在在线环境中协商意义、评估可信度并形成认知，从而推进数字健康传播研究，并为在线健康信息治理与减少伤害提供启示。",
    },
    "10.1080/1369118x.2026.2703856": {
        "source": "OpenAlex 公开论文元数据",
        "english": "Online targeted advertising has seen increasing use by political actors. With this technique, messages can be tailored to align with the preferences of individual recipients. In this study, we address the lack of understanding about the factors that influence voters' perceptions of political ads as tailored, as well as the extent to which tailoring perceptions (e.g., the extent to which an individual experiences a message as fitting to them) influence ad evaluation. We first argue that the issue content and the sending parties are the main drivers behind the perception of a political ad as tailored. Subsequently, we hypothesize that perceived tailoring influences ad evaluation positively. Lastly, we test what drives tailoring effects: actual tailoring or perceptions of it. To test these hypotheses, we employ an innovative design to capture ad exposure and responses. We combine the mobile experience sampling method (mESM) with data donations in the form of screenshots of actual ads received by respondents during election campaigns in the United States (2022) and Germany (2021). Our findings confirm that partisan alignment enhances the perception of the ad as tailored. We have mixed results for the effect of issue alignment, which only increases tailoring perceptions in Germany. Furthermore, ads perceived as tailored are evaluated more positively in both the US and Germany. Lastly, we find evidence for partial mediation through tailoring perceptions. Our findings have important implications for understanding the mechanisms of how targeted political ads on social media could impact society.",
        "chinese": "政治行动者越来越多地使用在线定向广告。这项技术可以按照个体接收者的偏好定制信息。本研究关注两个尚未得到充分解释的问题：哪些因素会影响选民将政治广告感知为定制内容，以及这种定制感知会在多大程度上影响广告评价。我们认为，议题内容和广告发送政党是政治广告被感知为定制内容的主要驱动因素，并提出定制感知会正向影响广告评价。研究还检验定制效果究竟来自实际定制，还是来自受众对定制的感知。为检验这些假设，我们采用一种捕捉广告暴露及受众反应的设计，将移动经验抽样法（mESM）与数据捐赠结合起来；受访者捐赠其在美国 2022 年和德国 2021 年选举期间实际收到的广告截图。结果表明，党派一致性会增强广告的定制感知。议题一致性的结果并不一致，它只在德国样本中提高定制感知。在美国和德国，被感知为定制内容的广告都得到更积极的评价。研究还发现定制感知存在部分中介作用。这些发现有助于理解社交媒体定向政治广告影响社会的具体过程。",
    },
}


METADATA_OVERRIDES = {
    "10.1080/1369118x.2026.2631709": {
        "作者": "Alex Gekker, Sam Hind, Gabriel Pereira, Fernando van der Vlist",
    },
}


def parse_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    try:
        return datetime.fromisoformat(str(value)).date()
    except ValueError:
        return None


def marker(cell) -> str:
    fill = cell.fill
    if fill.fill_type != "solid":
        return "无"
    if fill.fgColor.type == "rgb" and fill.fgColor.rgb in {"FFFFFF00", "00FFFF00"}:
        return "黄色"
    return "棕色"


def exclusion_reason(method: str, abstract: str) -> str:
    text = (abstract or "").lower()
    if "survey" in text or method in {"传统量化分析", "量化分析"}:
        return "传统调查、面板或一般统计分析；摘要未显示计算方法或明确的在线实验。"
    if method in {"质性分析", "纯质性分析"}:
        if "manual content" in text or "qualitative content" in text:
            return "人工质性内容分析；未使用计算分析。"
        return "访谈、民族志、批判分析或其他质性方法；未使用计算方法。"
    if method in {"理论分析", "纯理论分析"}:
        return "理论、概念或政策讨论；没有计算性经验方法。"
    if method == "综述":
        return "综述未使用计算文本或计算数据方法。"
    if "digital diary" in text:
        return "数字日记属于数字化资料收集，但摘要未显示计算分析。"
    return "摘要未显示计算方法或与计算传播相关的在线实验。"


def decision_for(doi: str, method: str, abstract: str) -> tuple[str, str, str, str]:
    key = (doi or "").lower().strip()
    if key in INCLUDE:
        evidence, method_group = INCLUDE[key]
        basis = "出版方全文方法明确" if key in FULLTEXT_VERIFIED else "摘要明确"
        return "纳入", f"议题符合，且方法证据显示：{evidence}。", method_group, basis
    if key in EXCLUDE_OVERRIDES:
        return "排除", EXCLUDE_OVERRIDES[key], "不适用", "出版方全文方法明确"
    if key in PENDING:
        return "待全文确认", PENDING[key], "待确认", "摘要不足"
    return "排除", exclusion_reason(method or "", abstract or ""), "不适用", "摘要明确或原表方法明确"


def normalize_xlsx_archive(path: Path) -> None:
    """Make generated XLSX archives stable across repeated runs."""
    with ZipFile(path, "r") as source:
        members = [(item, source.read(item.filename)) for item in source.infolist()]
    temporary = path.with_suffix(path.suffix + ".tmp")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as target:
        for original, payload in sorted(members, key=lambda pair: pair[0].filename):
            if original.filename == "docProps/core.xml":
                payload = re.sub(
                    rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                    rb"\g<1>2026-09-22T00:00:00Z\g<2>",
                    payload,
                )
            item = ZipInfo(original.filename, date_time=(2026, 9, 22, 0, 0, 0))
            item.compress_type = ZIP_DEFLATED
            item.external_attr = original.external_attr
            item.internal_attr = original.internal_attr
            item.create_system = original.create_system
            target.writestr(item, payload)
    os.replace(temporary, path)


def copy_2026_workbook() -> None:
    wb = load_workbook(SOURCE)
    ws = wb.active
    headers = {cell.value: cell.column for cell in ws[1]}
    date_col = headers["发布日期"]
    for row in range(ws.max_row, 1, -1):
        published = parse_date(ws.cell(row, date_col).value)
        if published is None or not (START_DATE <= published <= END_DATE):
            ws.delete_rows(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.properties.description = "仅保留 2026-01-01 至 2026-09-22 的记录；未改动原字段和底色标记。"
    wb.properties.modified = FIXED_XLSX_TIME
    CROPPED.parent.mkdir(parents=True, exist_ok=True)
    wb.save(CROPPED)
    normalize_xlsx_archive(CROPPED)


def load_rows() -> tuple[list[str], list[dict[str, object]]]:
    wb = load_workbook(SOURCE, data_only=False)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    h = {name: idx + 1 for idx, name in enumerate(headers)}
    rows: list[dict[str, object]] = []
    for row_number in range(2, ws.max_row + 1):
        published = parse_date(ws.cell(row_number, h["发布日期"]).value)
        if published is None or not (START_DATE <= published <= END_DATE):
            continue
        record = {name: ws.cell(row_number, col).value for name, col in h.items()}
        record["发布日期"] = published.isoformat()
        record["源文件行号"] = row_number
        record["原标记"] = marker(ws.cell(row_number, h["推荐摘要"]))
        status, reason, method_group, evidence_level = decision_for(
            str(record.get("DOI") or ""),
            str(record.get("研究方法") or ""),
            str(record.get("原文摘要") or ""),
        )
        record["筛选状态"] = status
        record["筛选理由"] = reason
        record["计算方法/实验"] = method_group
        record["判断依据"] = evidence_level
        key = str(record.get("DOI") or "").lower().strip()
        for field, value in METADATA_OVERRIDES.get(key, {}).items():
            record[field] = value
        if key in ABSTRACT_OVERRIDES:
            abstract = ABSTRACT_OVERRIDES[key]
            if not record.get("原文摘要"):
                record["原文摘要"] = abstract["english"]
            if not record.get("中文摘要"):
                record["中文摘要"] = abstract["chinese"]
            record["摘要来源"] = abstract["source"]
        else:
            record["摘要来源"] = "原始工作簿"
        if status == "纳入":
            record["相关级别"] = "中等相关" if key in MEDIUM_RELEVANCE else "明确相关"
            record["边界说明"] = MEDIUM_RELEVANCE.get(key, "")
            record["核心方法"] = METHOD_LABELS.get(key, "")
            record["资料与实施"] = METHOD_OVERRIDES.get(key, "")
            record["复核推荐理由"] = RECOMMENDATION_OVERRIDES.get(
                key,
                compact_text(record.get("推荐理由"), reason),
            )
        else:
            record["相关级别"] = "不适用"
            record["边界说明"] = ""
            record["核心方法"] = ""
            record["资料与实施"] = ""
            record["复核推荐理由"] = ""
        rows.append(record)
    return headers, rows


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            clean = {}
            for key in fields:
                value = row.get(key)
                if value is None:
                    clean[key] = ""
                else:
                    clean[key] = " ".join(part.strip() for part in str(value).splitlines() if part.strip())
            writer.writerow(clean)


def style_sheet(ws, widths: dict[str, int]) -> None:
    navy = "1F4E78"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 34
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for name, width in widths.items():
        for cell in ws[1]:
            if cell.value == name:
                ws.column_dimensions[get_column_letter(cell.column)].width = width
                break


def add_table_sheet(wb: Workbook, name: str, rows: list[dict[str, object]], fields: list[str]) -> None:
    ws = wb.create_sheet(name)
    ws.append(fields)
    for record in rows:
        ws.append([record.get(field, "") for field in fields])
    widths = {
        "源文件行号": 11,
        "原标记": 10,
        "筛选状态": 12,
        "相关级别": 12,
        "边界说明": 52,
        "判断依据": 15,
        "发布日期": 12,
        "英文标题": 48,
        "中文标题": 42,
        "作者": 30,
        "作者机构": 34,
        "研究方法": 16,
        "主题标签": 28,
        "计算方法/实验": 23,
        "核心方法": 34,
        "资料与实施": 80,
        "筛选理由": 52,
        "复核推荐理由": 58,
        "摘要来源": 28,
        "推荐摘要": 55,
        "中文摘要": 55,
        "原文摘要": 65,
        "DOI": 30,
        "原文链接": 42,
        "PDF状态": 12,
        "精读状态": 12,
    }
    style_sheet(ws, widths)
    field_to_col = {cell.value: cell.column for cell in ws[1]}
    for row_idx, record in enumerate(rows, start=2):
        mark_cell = ws.cell(row_idx, field_to_col["原标记"])
        if record["原标记"] == "黄色":
            mark_cell.fill = PatternFill("solid", fgColor="FFF2CC")
        elif record["原标记"] == "棕色":
            mark_cell.fill = PatternFill("solid", fgColor="C9B18F")
        status_cell = ws.cell(row_idx, field_to_col["筛选状态"])
        status_colors = {"纳入": "E2F0D9", "待全文确认": "FFF2CC", "排除": "F2F2F2"}
        status_cell.fill = PatternFill("solid", fgColor=status_colors.get(record["筛选状态"], "FFFFFF"))
        for link_field in ("DOI", "原文链接"):
            if link_field not in field_to_col:
                continue
            link_cell = ws.cell(row_idx, field_to_col[link_field])
            if link_field == "DOI" and link_cell.value:
                link_cell.hyperlink = f"https://doi.org/{link_cell.value}"
            elif link_cell.value:
                link_cell.hyperlink = str(link_cell.value)
            if link_cell.value:
                link_cell.style = "Hyperlink"


def write_curated_workbook(rows: list[dict[str, object]]) -> None:
    high = [dict(row, PDF状态="待获取", 精读状态="未开始") for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "明确相关"]
    medium = [dict(row, PDF状态="待获取", 精读状态="未开始") for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "中等相关"]
    included = high + medium
    pending = [dict(row, PDF状态="待获取", 精读状态="未开始") for row in rows if row["筛选状态"] == "待全文确认"]
    excluded = [row for row in rows if row["筛选状态"] == "排除"]
    selected_fields = [
        "源文件行号", "原标记", "筛选状态", "相关级别", "边界说明", "判断依据", "发布日期", "英文标题", "中文标题", "作者", "作者机构",
        "研究方法", "主题标签", "计算方法/实验", "核心方法", "资料与实施", "筛选理由", "复核推荐理由", "推荐摘要", "中文摘要", "原文摘要", "摘要来源",
        "DOI", "原文链接", "PDF状态", "精读状态",
    ]
    log_fields = ["源文件行号", "原标记", "筛选状态", "相关级别", "边界说明", "判断依据", "筛选理由", "复核推荐理由", "计算方法/实验", "核心方法", "资料与实施"] + [
        key for key in rows[0].keys() if key not in {
            "源文件行号", "原标记", "筛选状态", "相关级别", "边界说明", "判断依据", "筛选理由", "复核推荐理由", "计算方法/实验", "核心方法", "资料与实施"
        }
    ]
    wb = Workbook()
    wb.remove(wb.active)
    info = wb.create_sheet("筛选说明")
    info_rows = [
        ("项目", "Information, Communication & Society 2026 计算传播学初筛"),
        ("时间范围", f"{START_DATE.isoformat()} 至 {END_DATE.isoformat()}"),
        ("明确纳入", len(included)),
        ("其中明确相关", len(high)),
        ("其中中等相关", len(medium)),
        ("待全文确认", len(pending)),
        ("排除", len(excluded)),
        ("纳入条件 1", "研究问题直接涉及 AI、LLM、社交媒体、数字平台、算法传播或在线公共讨论等传播现象。"),
        ("纳入条件 2", "计算步骤必须进入测量、分析、建模或模拟；只用爬虫取数、一般统计或传统调查不算。"),
        ("纳入条件 3", "研究数字传播现象、且在真实或仿真在线环境中实施的实验可以纳入。"),
        ("排除", "纯理论/批判/政策讨论、民族志/访谈/一般质性研究、传统调查或面板研究。"),
        ("方法核查", "35 篇纳入论文均已回到 Taylor & Francis 出版方全文核对研究设计、样本、处理步骤和实际算法或统计模型。全文方法核查不等于已归档 PDF 或完成逐篇精读。"),
        ("原底色", "黄色＝用户原判为计算传播；棕色＝用户原判为不确定；无＝未标记。"),
    ]
    for row in info_rows:
        info.append(row)
    info.column_dimensions["A"].width = 18
    info.column_dimensions["B"].width = 95
    for row in info.iter_rows():
        row[0].font = Font(bold=True)
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    add_table_sheet(wb, "明确纳入", included, selected_fields)
    add_table_sheet(wb, "待全文确认", pending, selected_fields)
    add_table_sheet(wb, "完整筛选记录", rows, log_fields)
    wb.active = wb.sheetnames.index("明确纳入")
    wb.properties.created = FIXED_XLSX_TIME
    wb.properties.modified = FIXED_XLSX_TIME
    CURATED.parent.mkdir(parents=True, exist_ok=True)
    wb.save(CURATED)
    normalize_xlsx_archive(CURATED)

    curated_csv_fields = [
        "筛选状态", "相关级别", "边界说明", "发布日期", "英文标题", "中文标题", "作者", "作者机构", "期刊", "研究方法", "主题标签",
        "计算方法/实验", "核心方法", "资料与实施", "筛选理由", "复核推荐理由", "摘要来源", "原标记", "DOI", "原文链接",
    ]
    write_csv(CURATED_CSV, curated_csv_fields, included)
    write_csv(SCREENING_CSV, log_fields, rows)


def write_index(rows: list[dict[str, object]]) -> None:
    high = [row for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "明确相关"]
    medium = [row for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "中等相关"]
    pending = [row for row in rows if row["筛选状态"] == "待全文确认"]
    lines = [
        "# 文献总台账",
        "",
        f"> 范围：Information, Communication & Society，{START_DATE.isoformat()} 至 {END_DATE.isoformat()}。",
        "> 已完成标题摘要初筛，并在 Taylor & Francis 出版方网页逐篇核查 35 篇纳入论文的方法；尚未归档 PDF，也没有生成精读笔记。",
        "",
        f"## 明确相关（{len(high)} 篇）",
        "",
        "| 日期 | 论文 | DOI | 方法证据 | 全文 | 精读 |",
        "|---|---|---|---|---|---|",
    ]
    for row in high:
        title = str(row["英文标题"]).replace("|", "\\|")
        doi = str(row["DOI"])
        method = str(row["计算方法/实验"]).replace("|", "\\|")
        lines.append(f"| {row['发布日期']} | {title} | [{doi}](https://doi.org/{doi}) | {method} | 待获取 | 未开始 |")
    lines += [
        "",
        f"## 中等相关（{len(medium)} 篇）",
        "",
        "| 日期 | 论文 | DOI | 方法证据 | 边界说明 | 全文 | 精读 |",
        "|---|---|---|---|---|---|---|",
    ]
    for row in medium:
        title = str(row["英文标题"]).replace("|", "\\|")
        doi = str(row["DOI"])
        method = str(row["计算方法/实验"]).replace("|", "\\|")
        boundary = str(row["边界说明"]).replace("|", "\\|")
        lines.append(f"| {row['发布日期']} | {title} | [{doi}](https://doi.org/{doi}) | {method} | {boundary} | 待获取 | 未开始 |")
    lines += [
        "",
        f"## 待全文确认（{len(pending)} 篇）",
        "",
        "| 日期 | 论文 | DOI | 待确认事项 |",
        "|---|---|---|---|",
    ]
    for row in pending:
        title = str(row["英文标题"]).replace("|", "\\|")
        doi = str(row["DOI"])
        reason = str(row["筛选理由"]).replace("|", "\\|")
        lines.append(f"| {row['发布日期']} | {title} | [{doi}](https://doi.org/{doi}) | {reason} |")
    lines += [
        "",
        "## 全文与笔记状态",
        "",
        "- `literature/papers/` 当前为空。只有取得可核验全文后，论文才进入“全文已取得”状态。",
        "- `literature/notes/` 当前只有模板。没有完整原文时，不生成页码证据或精读内容。",
        "- 后续下载的受版权保护 PDF 默认仅保存在本地，不提交到公开仓库。",
        "",
    ]
    INDEX_MD.write_text("\n".join(lines), encoding="utf-8")


def compact_text(value: object, fallback: str = "") -> str:
    if value in (None, ""):
        return fallback
    text = " ".join(part.strip() for part in str(value).splitlines() if part.strip())
    text = re.sub(r"Citation(\d{4})", r"\1", text)
    return text.replace("引用来源：", "")


def write_presentation_markdown(rows: list[dict[str, object]]) -> None:
    """Publish a public-facing, paper-by-paper bilingual catalogue."""
    high = [row for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "明确相关"]
    medium = [row for row in rows if row["筛选状态"] == "纳入" and row["相关级别"] == "中等相关"]
    lines = [
        "# *Information, Communication & Society* 2026 计算传播学论文目录",
        "",
        f"> 时间范围：{START_DATE.isoformat()} 至 {END_DATE.isoformat()}",
        f"> 本目录收录 35 篇，其中明确相关 {len(high)} 篇、中等相关 {len(medium)} 篇。",
        "",
        "## 筛选口径",
        "",
        "本目录采用以下项目筛选口径：",
        "",
        "1. 研究问题直接涉及 AI、LLM、社交媒体、数字平台、算法传播或在线公共讨论等传播现象；",
        "2. 计算步骤进入测量、分析、建模或模拟环节。只用爬虫取数、一般统计或传统调查研究不计为计算方法；",
        "3. 研究数字传播现象、且在真实或仿真在线环境中实施的实验计入。",
        "",
        "“明确相关”表示议题与方法均直接符合上述口径。“中等相关”表示论文仍保留在精选目录中，但议题或方法有一项位于边界。该分级是本项目的操作性判断，不代表统一的学科分类标准。",
        "",
        f"## 明确相关（{len(high)} 篇）",
        "",
    ]

    def add_paper(row: dict[str, object], number: int, show_boundary: bool) -> None:
        english_title = compact_text(row.get("英文标题"))
        chinese_title = compact_text(row.get("中文标题"))
        doi = compact_text(row.get("DOI"))
        lines.extend([f"### {number}. {english_title}", ""])
        metadata = [
            ("中文标题", chinese_title),
            ("作者", compact_text(row.get("作者"))),
            ("作者机构", compact_text(row.get("作者机构"))),
            ("发表日期", compact_text(row.get("发布日期"))),
            ("主题标签", compact_text(row.get("主题标签"))),
        ]
        for label, value in metadata:
            if value:
                lines.extend([f"**{label}：** {value}", ""])
        if doi:
            lines.extend([f"**DOI：** [{doi}](https://doi.org/{doi})", ""])

        chinese_abstract = compact_text(row.get("中文摘要"))
        english_abstract = compact_text(row.get("原文摘要"))
        if chinese_abstract:
            lines.extend(["#### 中文摘要", "", chinese_abstract, ""])

        key = doi.lower()
        method_evidence = METHOD_LABELS.get(key, INCLUDE.get(key, ("", ""))[0])
        method = METHOD_OVERRIDES.get(
            key,
            compact_text(row.get("筛选理由")).removeprefix("议题符合，且方法证据显示："),
        )
        if method_evidence or method:
            lines.extend(["#### 研究设计与方法", ""])
            if method_evidence:
                lines.extend([f"**核心方法：** {method_evidence.rstrip('。')}。", ""])
            if method:
                lines.extend([f"**资料与实施：** {method}", ""])

        if english_abstract:
            lines.extend(["#### English Abstract", "", english_abstract, ""])

        source = PUBLIC_SOURCE_LINKS.get(key)
        if source:
            label, description, url = source
            lines.extend([f"**{label}：** [{description}]({url})", ""])

        if show_boundary:
            lines.extend([
                "#### 中等相关说明",
                "",
                compact_text(row.get("边界说明")),
                "",
            ])
        lines.extend(["---", ""])

    number = 1
    for row in high:
        add_paper(row, number, False)
        number += 1

    lines.extend([
        f"## 中等相关（{len(medium)} 篇）",
        "",
        "以下论文保留在 35 篇精选目录中，但其计算方法或传播议题与本项目口径的衔接相对间接，因此单独列出。",
        "",
    ])
    for row in medium:
        add_paper(row, number, True)
        number += 1

    PRESENTATION_MD.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"Missing source workbook: {SOURCE}")
    copy_2026_workbook()
    _, rows = load_rows()
    write_curated_workbook(rows)
    write_index(rows)
    write_presentation_markdown(rows)
    counts = {status: sum(row["筛选状态"] == status for row in rows) for status in ("纳入", "待全文确认", "排除")}
    if counts != {"纳入": 35, "待全文确认": 0, "排除": 124}:
        raise SystemExit(f"Unexpected counts: {counts}")
    relevance_counts = {level: sum(row["相关级别"] == level for row in rows) for level in ("明确相关", "中等相关")}
    if relevance_counts != {"明确相关": 28, "中等相关": 7}:
        raise SystemExit(f"Unexpected relevance counts: {relevance_counts}")
    print(f"Processed {len(rows)} papers: {counts}")
    print(CROPPED)
    print(CURATED)


if __name__ == "__main__":
    main()
