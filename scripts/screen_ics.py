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


RECOMMENDATION_OVERRIDES = {
    "10.1080/1369118x.2026.2703856": "研究在线政治广告的个性化感知，并结合移动经验抽样与真实广告截图数据捐赠；议题明确属于数字政治传播，但计算分析的中心性较弱，因此列为中等相关。",
    "10.1080/1369118x.2026.2678364": "研究 Reddit 上肥胖管理药物的消费者叙事，使用 Python 抓取、Top2Vec 与 LLM 辅助主题整理，议题与计算文本方法均符合本项目口径。",
    "10.1080/1369118x.2026.2669800": "研究平台化游戏中的数字劳动，使用 Discord 大规模数据、R 清洗和正则识别，但核心解释仍依赖抽样后的主题分析，因此列为中等相关。",
    "10.1080/1369118x.2026.2636134": "研究社交媒体中的反 LGBTQIA+ 敌意与旁观者干预，采用预注册 3×3 在线实验；根据本项目对计算相关在线实验的纳入规则，列为明确相关。",
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


def compact_text(value: object, fallback: str = "原工作簿未提供") -> str:
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
        lines.extend([
            f"### {number}. {english_title}",
            "",
            f"**中文标题：** {chinese_title}",
            "",
            f"**作者：** {compact_text(row.get('作者'))}",
            "",
            f"**作者机构：** {compact_text(row.get('作者机构'))}",
            "",
            f"**发表日期：** {compact_text(row.get('发布日期'))}",
            "",
            f"**DOI：** [{doi}](https://doi.org/{doi})",
            "",
            f"**主题标签：** {compact_text(row.get('主题标签'))}",
            "",
            "#### 研究方法",
            "",
            f"- **原表方法标注：** {compact_text(row.get('研究方法'))}",
            f"- **计算方法或实验类型：** {compact_text(row.get('计算方法/实验'))}",
            f"- **具体方法证据：** {compact_text(row.get('筛选理由')).removeprefix('议题符合，且方法证据显示：')}",
            "",
            "#### 推荐理由",
            "",
            compact_text(row.get("复核推荐理由")),
            "",
        ])
        if show_boundary:
            lines.extend([
                "#### 中等相关说明",
                "",
                compact_text(row.get("边界说明")),
                "",
            ])
        if row.get("摘要来源") != "原始工作簿":
            lines.extend([
                f"**摘要来源：** {compact_text(row.get('摘要来源'))}",
                "",
            ])
        lines.extend([
            "#### 中文摘要",
            "",
            compact_text(row.get("中文摘要")),
            "",
            "#### English Abstract",
            "",
            compact_text(row.get("原文摘要")),
            "",
            "---",
            "",
        ])

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
