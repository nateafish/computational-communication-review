# ICS 2026 待定论文人工全文核查记录

核查日期：2026-09-22  
来源：Taylor & Francis Online 出版方完整网页  
范围：标题摘要初筛后仍待确认的 11 篇论文

本记录用于说明人工筛选所依据的方法证据。核查时逐篇打开出版方论文页，阅读摘要及研究设计、数据或方法部分，再按本项目口径判断。网页可访问不等于本地已经保存全文原件；这些论文在 `reference/INDEX.md` 中仍按 PDF“待获取”处理，也没有据此生成精读笔记或页码证据。

| DOI | 论文 | 出版方页面所见方法证据 | 判定 |
|---|---|---|---|
| [10.1080/1369118X.2026.2729550](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2729550) | Platformed partner markets: online dating and gendered assortative mating in Europe | 使用九国 Generations and Gender Survey II，并做加权多项 Logit 分析。数据和分析属于传统调查研究。 | 排除 |
| [10.1080/1369118X.2026.2713666](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2713666) | Toward inclusive online safety: evaluating the role of intersectionality on online abuse in the United Kingdom | 使用 Ofcom 多年调查与 MAIHDA，没有平台数字痕迹或计算内容分析。 | 排除 |
| [10.1080/1369118X.2026.2663190](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2663190) | Where digital hate finds its home: perpetrators' preferred affordances and their experiences with content moderation and user intervention | 四国配额调查，使用潜在剖面分析和结构模型。核心证据仍来自传统调查。 | 排除 |
| [10.1080/1369118X.2026.2713667](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2713667) | ‘Cool! I wanna try it!’: online exposure to e-cigarette advertisements increases e-cigarette susceptibility among U.S. sexual minority young adults | 使用 PATH 纵向调查研究自然广告暴露，不是在线随机实验，也没有计算方法。 | 排除 |
| [10.1080/1369118X.2026.2703856](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2703856) | Did you make this for me? A two-country study on individuals’ perceptions of tailoring in online political ads | 两国移动经验抽样与广告截图捐赠，结合人工内容编码、主轴因子分析、OLS 和因果中介模型。 | 纳入：中等相关；经验抽样与一般统计为主 |
| [10.1080/1369118X.2026.2674878](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2674878) | Calling out disclosures, protecting followers: the new governors of influencer culture on TikTok | 目的抽样 14 个 TikTok 视频，Zeeschuimer 用于抓取评论和元数据；核心分析是对视频与评论的细读式话语分析。爬虫只承担取数。 | 排除 |
| [10.1080/1369118X.2026.2667921](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2667921) | Price, proficiency, or permission: assessing platform data access for election research | 人工汇编并系统编码 72 种平台数据访问工具。研究对象与计算研究有关，但论文自身没有使用计算分析。 | 排除 |
| [10.1080/1369118X.2026.2678366](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2678366) | Algorithmic ambivalence: LGBTQ+ social media and mental health among Chinese gay men | 23 次半结构化访谈与人工主题分析。 | 排除 |
| [10.1080/1369118X.2026.2678364](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2678364) | Stigma to skepticism: online consumer narratives of obesity management medications | Python 抓取 5,789 条 Reddit 评论，Top2Vec 生成 572 个主题，GPT-4o 仅辅助筛题，最终主题由研究者人工归纳。 | 纳入：明确相关；Top2Vec 进入分析环节 |
| [10.1080/1369118X.2026.2669800](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2669800) | From player to asset: venture labor in axie infinity | 导出 13,329 份 Discord 申请，用 R、正则表达式和随机抽样整理材料，核心解释来自对 1,333 条申请的人工主题分析。 | 纳入：中等相关；计算环节主要是预处理 |
| [10.1080/1369118X.2026.2636134](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2636134) | Nudging against judging? Mitigating anti-LGBTQIA + online hostility by raising bystanders’ awareness and behavioral intentions | 明确写为预注册 3×3 在线实验，操控 Instagram 仿真帖文中的敌意类型和提示方式。 | 纳入：在线实验 |

## 核查结果

11 篇待定论文经人工全文核查后，4 篇纳入、7 篇排除，待全文确认队列清零。其后又对 35 篇纳入论文逐篇复核全文方法；公开目录中的“核心方法”和“资料与实施”采用第二轮结果。

## 排除项假阴性复查

用户指出部分摘要没有交代方法，不能据此直接判断为非计算研究。复查先通读 124 篇排除项的标题、摘要和原表方法字段，再将议题相关但方法缺失、含糊，或可能把实验写成 survey 的论文列为全文候选。下表记录第二轮候选的 Taylor & Francis 全文核查结果。

| DOI | 论文 | 正文方法证据与判断 |
|---|---|---|
| [10.1080/1369118X.2026.2715572](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2715572) | Corporate digital responsibility: what drives public expectations towards companies’ role in the digital society? | 荷兰代表性样本的预注册横截面在线问卷（n=2,196），使用 EFA、CFA、SEM 和调节分析。没有实验操控或计算传播方法，排除。 |
| [10.1080/1369118X.2026.2713669](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2713669) | Replatformization and the expansion of the alt-tech ecosystem on kiwi farms | 用公开帖文、平台文件和新闻材料重建平台迁移过程，属于质性案例分析，排除。 |
| [10.1080/1369118X.2026.2674925](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2674925) | Examining initial trust development and trusting intentions toward humanoid social robots | Prolific 横截面在线问卷（n=375）与 PLS-SEM。没有实验操控，排除。 |
| [10.1080/1369118X.2026.2681885](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2681885) | Media effects, selection effects, or no effects? | 瑞典四波面板问卷（n=3,530），使用 RI-CLPM 与 LGCM，属于传统纵向调查分析，排除。 |
| [10.1080/1369118X.2026.2678363](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2678363) | Mistrust as a precondition for critical thinking and media literacy | 两次相互关联的在线调查；第二次核验任务由人工按量规评分，分析使用 t 检验、相关和 OLS，不是随机实验，排除。 |
| [10.1080/1369118X.2026.2683531](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2683531) | Indigenous digital life | 出版方页面标明为书评，不是经验研究论文，排除。 |
| [10.1080/1369118X.2026.2659281](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2659281) | Reshaping platform governmentality | 从 Stack Exchange Data Dumps 查询 60 个线程，核心分析缩减到 22 个线程并采用溯因式质性分析。查询只承担取数，排除。 |
| [10.1080/1369118X.2026.2652514](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2652514) | Evaluating the potential and limits of trace-based interviews when studying news use | Wakoopa 数字踪迹用于制作个体访谈提示材料；核心证据来自 73 次访谈及反思式主题分析，排除。 |
| [10.1080/1369118X.2026.2645883](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2645883) | Diagnosing multimodal disinformation | 人工核验 407 个短视频，再选取 22 个视频进行人工多模态编码和质性解释，没有自动化多模态分析，排除。 |
| [10.1080/1369118X.2026.2636130](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2636130) | The (Un)desirable shield | 全国在线问卷（n=2,373）展示既有警示标签图片并询问是否见过；只随机题项顺序，没有随机处理条件，排除。 |
| [10.1080/1369118X.2026.2613437](https://www.tandfonline.com/doi/full/10.1080/1369118X.2026.2613437) | ‘It's still abuse’ | 正文为 2×4 随机组间在线情境实验，方法上符合项目实验口径；因涉及 AI 生成性影像虐待和色情网站情境，按用户本轮主题选择排除。 |
