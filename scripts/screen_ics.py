#!/usr/bin/env python3
"""Screen the 2026 Information, Communication & Society workbook.

The decisions in this first pass are abstract-level screening decisions.  A paper
is included only when both the research object and the method fit the project's
definition.  Papers without enough method information are kept in a separate
full-text-confirmation queue.
"""

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
        "判断依据": 15,
        "发布日期": 12,
        "英文标题": 48,
        "中文标题": 42,
        "作者": 30,
        "研究方法": 16,
        "主题标签": 28,
        "计算方法/实验": 23,
        "筛选理由": 52,
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
    included = [dict(row, PDF状态="待获取", 精读状态="未开始") for row in rows if row["筛选状态"] == "纳入"]
    pending = [dict(row, PDF状态="待获取", 精读状态="未开始") for row in rows if row["筛选状态"] == "待全文确认"]
    excluded = [row for row in rows if row["筛选状态"] == "排除"]
    selected_fields = [
        "源文件行号", "原标记", "筛选状态", "判断依据", "发布日期", "英文标题", "中文标题", "作者",
        "研究方法", "主题标签", "计算方法/实验", "筛选理由", "推荐摘要", "中文摘要", "原文摘要",
        "DOI", "原文链接", "PDF状态", "精读状态",
    ]
    log_fields = ["源文件行号", "原标记", "筛选状态", "判断依据", "筛选理由", "计算方法/实验"] + [
        key for key in rows[0].keys() if key not in {
            "源文件行号", "原标记", "筛选状态", "判断依据", "筛选理由", "计算方法/实验"
        }
    ]
    wb = Workbook()
    wb.remove(wb.active)
    info = wb.create_sheet("筛选说明")
    info_rows = [
        ("项目", "Information, Communication & Society 2026 计算传播学初筛"),
        ("时间范围", f"{START_DATE.isoformat()} 至 {END_DATE.isoformat()}"),
        ("明确纳入", len(included)),
        ("待全文确认", len(pending)),
        ("排除", len(excluded)),
        ("纳入条件 1", "研究对象属于 AI、LLM、社交媒体、平台、数字传播等计算传播议题。"),
        ("纳入条件 2", "方法使用计算文本、机器学习、网络分析、数字痕迹、大规模平台数据、计算模拟、模型审计等。"),
        ("纳入条件 3", "与上述议题直接相关、且明确在线实施的实验可以纳入。"),
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
        "筛选状态", "发布日期", "英文标题", "中文标题", "作者", "期刊", "研究方法", "主题标签",
        "计算方法/实验", "筛选理由", "原标记", "DOI", "原文链接",
    ]
    write_csv(CURATED_CSV, curated_csv_fields, included)
    write_csv(SCREENING_CSV, log_fields, rows)


def write_index(rows: list[dict[str, object]]) -> None:
    included = [row for row in rows if row["筛选状态"] == "纳入"]
    pending = [row for row in rows if row["筛选状态"] == "待全文确认"]
    lines = [
        "# 文献总台账",
        "",
        f"> 范围：Information, Communication & Society，{START_DATE.isoformat()} 至 {END_DATE.isoformat()}。",
        "> 已完成标题摘要初筛，并在 Taylor & Francis 出版方网页人工核查 11 篇待定论文的方法；尚未归档 PDF，也没有生成精读笔记。",
        "",
        f"## 明确纳入（{len(included)} 篇）",
        "",
        "| 日期 | 论文 | DOI | 方法证据 | 全文 | 精读 |",
        "|---|---|---|---|---|---|",
    ]
    for row in included:
        title = str(row["英文标题"]).replace("|", "\\|")
        doi = str(row["DOI"])
        method = str(row["计算方法/实验"]).replace("|", "\\|")
        lines.append(f"| {row['发布日期']} | {title} | [{doi}](https://doi.org/{doi}) | {method} | 待获取 | 未开始 |")
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


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"Missing source workbook: {SOURCE}")
    copy_2026_workbook()
    _, rows = load_rows()
    write_curated_workbook(rows)
    write_index(rows)
    counts = {status: sum(row["筛选状态"] == status for row in rows) for status in ("纳入", "待全文确认", "排除")}
    if counts != {"纳入": 35, "待全文确认": 0, "排除": 124}:
        raise SystemExit(f"Unexpected counts: {counts}")
    print(f"Processed {len(rows)} papers: {counts}")
    print(CROPPED)
    print(CURATED)


if __name__ == "__main__":
    main()
