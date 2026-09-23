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
}


EXCLUDE_OVERRIDES = {
    "10.1080/1369118x.2026.2729550": "全文显示使用九国 Generations and Gender Survey II 与加权多项 Logit；属于传统调查数据分析。",
    "10.1080/1369118x.2026.2713666": "全文显示使用 Ofcom 多年调查与 MAIHDA；没有平台数字痕迹或计算内容分析。",
    "10.1080/1369118x.2026.2663190": "全文显示为四国配额调查、潜在剖面分析与结构模型；属于传统调查研究。",
    "10.1080/1369118x.2026.2713667": "全文摘要显示使用 PATH 纵向调查，研究自然广告暴露；不是在线实验或计算方法。",
    "10.1080/1369118x.2026.2674878": "全文显示以少量 TikTok 视频及评论开展话语分析；爬虫只用于取数，核心分析为质性解释。",
    "10.1080/1369118x.2026.2667921": "全文通过人工汇编和系统编码比较 72 种数据访问工具；研究对象与计算方法有关，但本身未使用计算分析。",
    "10.1080/1369118x.2026.2678366": "全文显示为 23 次半结构化访谈与人工主题分析。",
}


PENDING: dict[str, str] = {}


# These papers remain in the 35-paper collection but sit on one boundary of the
# user-authorized project definition.  The labels are project-specific screening
# judgments, not claims about a universal disciplinary taxonomy.
MEDIUM_RELEVANCE = {
    "10.1080/1369118x.2026.2703856": "议题属于在线政治广告，但研究主要结合移动经验抽样与截图数据捐赠；摘要中计算分析不是核心环节，列为方法边界。",
    "10.1080/1369118x.2026.2669800": "使用 R 和正则表达式处理大规模 Discord 申请，但核心解释来自随机抽样后的主题分析；计算环节主要集中在采集与预处理。",
    "10.1080/1369118x.2026.2642840": "研究社交媒体参与并量化隐喻频率与原创性，但现有摘要没有明确说明自动化文本识别或其他计算分析，列为方法证据边界。",
    "10.1080/1369118x.2026.2631709": "使用主题聚类分析大规模专利文本，但核心问题偏向云基础设施与汽车创新，与传播研究的联系相对间接。",
}


METHOD_OVERRIDES = {
    "10.1080/1369118x.2026.2733507": "研究收集欧洲委员会主席、副主席和委员在 2016 年至 2024 年中期发表的 238 篇演讲，先用主题模型和文本社群检测识别反复出现的话语组合，再结合定性内容分析与时间比较解释框架的变化。",
    "10.1080/1369118x.2026.2715573": "研究分析 2022 年瑞典议会选举期间 Meta 广告库中的 2,395 条 Facebook 和 Instagram 广告，并把广告内容编码与竞选支出、受众人口特征和定向信息结合起来比较各党的付费传播策略。",
    "10.1080/1369118x.2026.2713664": "研究把 758 名匈牙利 Facebook 用户捐赠的数字行为记录与同一批参与者的问卷回答逐一配对，并通过预注册的题序实验比较不同问卷设计对自我报告准确性的影响。",
    "10.1080/1369118x.2026.2716778": "研究以大规模 X 数据追踪美国移民议题中的不文明表达。大型语言模型用于识别用户立场和语言不文明程度，动态提及网络与图机器学习则用于观察敌意言论如何随时间跨越立场边界。",
    "10.1080/1369118x.2026.2695177": "论文采用基于代理模型（ABM），以有界置信意见动力学为基础，在 NetLogo 中设置 2,500 个代理人，并用加泰罗尼亚 2011 年调查数据初始化国家认同。模型采用线上与线下两层网络，交叉操控随机／同质网络和过滤气泡开／关四种条件；每种条件重复模拟 100 次，运行至 400 个时间步，结果在 RStudio 中分析。",
    "10.1080/1369118x.2026.2696929": "研究分析 2021 至 2023 年间 53 个巴西亲博索纳罗 Facebook 群组和页面的 1,200 多万条帖文。作者用大型语言模型分类内容，构建情感极化与参与平衡指标，再以时间序列和回归模型识别重大政治事件前后的变化。",
    "10.1080/1369118x.2026.2699262": "研究汇集中国三起重大性别暴力事件下的大规模在线评论，使用词嵌入计算不同人物称谓与道德范畴之间的语义距离，再比较性别化称谓和性别中立称谓如何引出不同的道德评价。",
    "10.1080/1369118x.2026.2691775": "研究收集欧洲、北美、非洲、亚洲和大洋洲 12 家报纸在 2010 至 2024 年发表的 6,829 篇 AI 新闻，结合主题模型与情感分析比较地区差异，并识别报道框架随重大技术和政治事件发生的转折。",
    "10.1080/1369118x.2026.2685121": "研究在 2024 年美国大选日前对 ChatGPT 生成的政治与选举信息开展混合方法审计，系统辨认事实不准确、回避性回答、虚构内容和关键信息遗漏，并据此评估生成式平台在选举信息环境中的治理责任。",
    "10.1080/1369118x.2026.2686314": "研究以可再生能源相关的品牌化 Facebook 帖文为材料，围绕合作伙伴与发布页面构建四个二部网络，用网络位置和连接关系呈现平台界面中不易直接看见的战略合作结构。",
    "10.1080/1369118x.2026.2686315": "研究分析微博上 22,925 条性暴力相关帖文及其 115 万条评论，对文字、图像等框架方式和发布者性别线索进行编码，并用广义线性混合模型估计它们与评论区性别对立程度的关系。",
    "10.1080/1369118x.2026.2678364": "研究以 Reddit 上 5,789 条肥胖管理药物评论为材料，用 Python 完成数据采集，借助 Top2Vec 发现主题，并由大型语言模型辅助整理主题，最后把计算结果放回网络民族志语境中解释。",
    "10.1080/1369118x.2026.2652497": "研究提出区分政治广告“受众定位”与“内容定制”的测量框架，并将其用于 2021 年德国大选最后四周的 4,760 条 Meta 广告。数据来自人工多模态编码、Meta Ad Targeting 数据集和 Meta 广告库。",
    "10.1080/1369118x.2026.2667914": "研究开展一项 2×2×2 在线实验（N=271），同时操控新闻真假、来源可信度和社会认同线索，并测量参与者对 Twitter 新闻功能的看法，以检验这些因素如何影响真假新闻判断。",
    "10.1080/1369118x.2026.2665253": "研究把一千多名美国受访者一个月的被动域名级浏览记录与恶意网站名单匹配，比较不同群体的暴露差异，并进一步检验差异究竟来自人口特征还是浏览活动强度。",
    "10.1080/1369118x.2026.2665252": "研究分析 r/SexWorkers 和 r/OnlyFans 中的 38,000 多条评论，使用结构主题模型识别 17 个讨论主题，再以定性内容分析解释平台治理、法律、受害经历和经济实践等主题的具体含义。",
    "10.1080/1369118x.2026.2659287": "研究分析德语 X 社群围绕新冠疫情、俄乌战争和加沙战争发布的 400 多万条帖文，把转发网络中的同质互动与语言一致性合成为回音室得分，并结合社群分析追踪极化群体的持续和跨议题重现。",
    "10.1080/1369118x.2026.2659279": "研究把 340 万条微博与 3,851 家上市公司 2007 至 2023 年的面板数据相连，利用分阶段双重差分设计估计国有环境类微博平台曝光对企业环境投资和信息披露的影响，并检验信号与声誉两条解释路径。",
    "10.1080/1369118x.2026.2654668": "研究围绕德国 Querdenken 抗议运动建立多层 Telegram 数据集，联合使用纵向内容分析、URL 来源分析和网络分析，考察激进叙事、替代信息源及群组连接是否共同推动极端观点进入日常讨论。",
    "10.1080/1369118x.2026.2659280": "研究建立包含 10,800 个 TikTok 科学传播视频和 448,002 条评论的多模态数据集，测量传播者的性别、种族与表达风格，并分析这些身份组合和风格匹配如何影响参与度及评论行为。",
    "10.1080/1369118x.2026.2647352": "研究对巴西用户围绕 Spotify Wrapped 发布的 X 帖文进行社会网络分析，并访谈 38 名 Spotify 用户。网络结构用于描述这一年度事件形成的短暂公众，访谈则解释用户如何主动管理算法留下的年度音乐画像。",
    "10.1080/1369118x.2026.2648695": "研究以《帝国时代 II：决定版》超过 2,500 万场多人比赛为个体行为数据，计算国家内部及国家之间的玩家互动量，并以结构引力模型检验文化距离和既有社会联系对在线互动的影响。",
    "10.1080/1369118x.2026.2636138": "研究按 PRISMA-ScR 流程筛得 145 篇全文论文，人工编码研究涉及的平台、内容形式和平台功能，同时使用 BERTopic 识别文献主题，从而绘制针对性少数群体仇恨言论研究的整体版图。",
    "10.1080/1369118x.2026.2655885": "研究对 3,879 条有关《守望者》和《洛夫克拉夫特之乡》的 Reddit 评论进行主题建模，并结合批判性技术文化话语分析，追踪用户如何在支持者、反对者和适应者三种角色之间转换。",
    "10.1080/1369118x.2026.2645882": "研究收集 r/UFO 与 r/UAP 的帖文和评论，建立政府公告、泄密、听证会、新闻报道和国防部公开说明的事件时间线，再用时间序列分析估计官方介入前后阴谋叙事和信任表达的变化。",
    "10.1080/1369118x.2026.2642851": "研究把 2016 至 2018 年 InfoWars 每日销售额与《Alex Jones Show》及网站文章的语言、音频和主题特征按天对齐，通过时间序列模型检验内容能否预测次日销售，以及销售是否反过来影响后续内容。",
    "10.1080/1369118x.2026.2636134": "研究采用预注册 3×3 在线实验，在仿真的 Instagram 帖文中分别操控敌意类型与社会提示方式，观察参与者的可接受度判断、负面情绪、干预意愿和内容审核偏好。",
    "10.1080/1369118x.2026.2633222": "研究分析约 1,500 万条英语、德语、意大利语、日语和韩语推文，先以计算方法比较“价值观”一词出现的主题语境，再结合定性分析解释它在不同语言中的身份表达和边界划分功能。",
    "10.1080/1369118x.2026.2633216": "研究汇集 2023 年中国三个主要数字平台的热榜数据，先测量话题跨平台同时流行的规模，再以以色列—哈马斯战争为案例，估计其他平台上榜对微博同一政治话题注意力的影响。",
    "10.1080/1369118x.2026.2624702": "研究使用填充—掩码关联测试（FMAT）审计四种主要中文预训练语言模型，并以英文模型为参照，分别测量模型在性别、职业地位和年龄维度上的语言态度偏差。",
    "10.1080/1369118x.2026.2623523": "研究追踪 11 个主要政治类 Reddit 子版块长达 16 年的互动，分别计算共同用户、信息共享、跨版发帖、跨版评论、链接关系和话语身份等多种重叠指标，以判断跨立场交流是否真实发生。",
    "10.1080/1369118x.2026.2703856": "研究把美国 2022 年和德国 2021 年选举期间的移动经验抽样与广告截图捐赠结合起来，将参与者实际收到的政治广告与即时评价配对，并检验党派一致、议题一致、定制感知和广告评价之间的关系。",
    "10.1080/1369118x.2026.2669800": "研究从四个 Discord 社群导出约 13,300 份 Axie Infinity“奖学金”申请，用 R 和正则表达式清理、识别材料，再从中随机抽样开展主题分析。计算环节主要服务于大规模材料的整理与抽样，核心解释仍来自质性分析。",
    "10.1080/1369118x.2026.2642840": "研究分析 2023 年土耳其大选两位总统候选人和五个主要政党的 4,131 条推文，以数据驱动框架测量每条推文中隐喻的频率与原创性，并比较它们对点赞、转发和回复的影响。",
    "10.1080/1369118x.2026.2631709": "研究以 69,421 个全球专利家族为语料，对汽车制造商、芯片商、电信与地图服务商等参与者的专利主题进行聚类，再以社会技术想象视角解释云基础设施如何被写入联网自动驾驶的未来。",
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
            record["复核推荐理由"] = RECOMMENDATION_OVERRIDES.get(
                key,
                compact_text(record.get("推荐理由"), reason),
            )
        else:
            record["相关级别"] = "不适用"
            record["边界说明"] = ""
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
        "研究方法", "主题标签", "计算方法/实验", "筛选理由", "复核推荐理由", "推荐摘要", "中文摘要", "原文摘要", "摘要来源",
        "DOI", "原文链接", "PDF状态", "精读状态",
    ]
    log_fields = ["源文件行号", "原标记", "筛选状态", "相关级别", "边界说明", "判断依据", "筛选理由", "复核推荐理由", "计算方法/实验"] + [
        key for key in rows[0].keys() if key not in {
            "源文件行号", "原标记", "筛选状态", "相关级别", "边界说明", "判断依据", "筛选理由", "复核推荐理由", "计算方法/实验"
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
        ("证据边界", "先依据标题与摘要筛选；11 篇待定论文另经 Taylor & Francis 出版方网页人工核查方法。网页核查不等于已归档 PDF 或完成精读。"),
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
        "计算方法/实验", "筛选理由", "复核推荐理由", "摘要来源", "原标记", "DOI", "原文链接",
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
        "> 已完成标题摘要初筛，并在 Taylor & Francis 出版方网页人工核查 11 篇待定论文的方法；尚未归档 PDF，也没有生成精读笔记。",
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
        f"> 本期共筛选 159 篇论文，保留 35 篇，其中明确相关 {len(high)} 篇、中等相关 {len(medium)} 篇。",
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

        key = doi.lower()
        method = METHOD_OVERRIDES.get(
            key,
            compact_text(row.get("筛选理由")).removeprefix("议题符合，且方法证据显示："),
        )
        if method:
            lines.extend(["#### 研究设计与方法", "", method, ""])

        source = PUBLIC_SOURCE_LINKS.get(key)
        if source:
            label, description, url = source
            lines.extend([f"**{label}：** [{description}]({url})", ""])

        recommendation = compact_text(row.get("复核推荐理由"))
        if recommendation:
            lines.extend(["#### 推荐理由", "", recommendation, ""])
        if show_boundary:
            lines.extend([
                "#### 中等相关说明",
                "",
                compact_text(row.get("边界说明")),
                "",
            ])
        if row.get("摘要来源") and row.get("摘要来源") != "原始工作簿":
            lines.extend([
                f"**摘要来源：** {compact_text(row.get('摘要来源'))}",
                "",
            ])
        chinese_abstract = compact_text(row.get("中文摘要"))
        english_abstract = compact_text(row.get("原文摘要"))
        if chinese_abstract:
            lines.extend(["#### 中文摘要", "", chinese_abstract, ""])
        if english_abstract:
            lines.extend(["#### English Abstract", "", english_abstract, ""])
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
    if relevance_counts != {"明确相关": 31, "中等相关": 4}:
        raise SystemExit(f"Unexpected relevance counts: {relevance_counts}")
    print(f"Processed {len(rows)} papers: {counts}")
    print(CROPPED)
    print(CURATED)


if __name__ == "__main__":
    main()
