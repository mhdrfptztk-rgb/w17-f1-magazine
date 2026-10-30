# -*- coding: utf-8 -*-
"""Build the W17 editorial microsite confirmation pack (XLSX baseline)."""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

DARK = "0E1012"
GRAPHITE = "1C1F23"
SILVER_TEXT = "E9ECEE"
TEAL = "00A19C"
TEAL_SOFT = "7FE3DD"
LIGHT_HDR = "E9EDEE"
ZEBRA = "F6F7F7"
BORDER = "D5D9DB"
MUTED = "7A8288"
WARN = "B4541E"

FONT = "Microsoft YaHei"
thin = Side(style="thin", color=BORDER)
box = Border(left=thin, right=thin, top=thin, bottom=thin)


def build_sheet(wb, name, tab_color, title, subtitle, headers, rows, widths,
                wrap_cols=(), num_formats=None, freeze=True, row_height=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color
    ncol = len(headers)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(row=1, column=1, value=title)
    c.fill = PatternFill("solid", fgColor=DARK)
    c.font = Font(name=FONT, size=15, bold=True, color=SILVER_TEXT)
    c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[1].height = 36

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    c = ws.cell(row=2, column=2 - 1, value=subtitle)
    c.font = Font(name=FONT, size=9, italic=True, color=MUTED)
    c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[2].height = 22

    hdr_row = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=hdr_row, column=i, value=h)
        c.fill = PatternFill("solid", fgColor=LIGHT_HDR)
        c.font = Font(name=FONT, size=9, bold=True, color=DARK)
        c.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin,
                          bottom=Side(style="medium", color=TEAL))
    ws.row_dimensions[hdr_row].height = 26

    for r, row in enumerate(rows, start=hdr_row + 1):
        for i, val in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=val)
            c.font = Font(name=FONT, size=9, color=DARK)
            c.alignment = Alignment(vertical="top", wrap_text=(i in wrap_cols))
            c.border = box
            if (r - hdr_row) % 2 == 0:
                c.fill = PatternFill("solid", fgColor=ZEBRA)
            if num_formats and i in num_formats:
                c.number_format = num_formats[i]
        if row_height:
            ws.row_dimensions[r].height = row_height

    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    if freeze:
        ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    return ws


wb = Workbook()
wb.remove(wb.active)

# ---------------------------------------------------------------- 00 导读
readme_head = ["项目", "内容"]
readme_rows = [
    ["项目名称", "MERCEDES-AMG PETRONAS F1 — W17 EDITORIAL"],
    ["站点主标语", "THE ART OF SPEED."],
    ["定位", "高端汽车杂志式数字专题站：不是车队官网，也不是数据后台；是一本可滚动、可点击、可阅读的 W17 数字杂志"],
    ["版本", "v0.1 立项确认稿 — 尚未编写任何前端代码"],
    ["数据截止", "2026-10-30。赛果截至 2026-10-25 美国大奖赛；墨西哥站（10-30 至 11-01）正在进行，按“未完成”处理"],
    ["视觉参考", "landonorris.com — 只借鉴：大图留白、非对称版式、超大英文标题、图文交错、克制的导航与滚动节奏；不复制其版式与素材"],
    ["技术栈", "React 19 + TypeScript + Vite + Tailwind CSS v4 + Framer Motion + 自绘 SVG 图表（d3-scale）"],
    ["交付范围", "8 个章节（Hero / Machine / Data / Victories / Battle / Drivers / Details / Season）+ 6 篇分站专题 + 2 篇车手专题 + 赛季时间轴 + 3 组交互图表"],
    ["本稿用途", "作为“内容与数据基线”：先确认定位、结构、数据口径与图片策略，确认后再开始写代码"],
    ["已核实数据", "18 个已完成分站、23 位车手积分、11 支车队积分、28 项 W17 技术参数（逐项注明来源）"],
    ["与原稿的偏差", "3 处必须修改：漏 2 场冠军、4 月巴林/沙特站已取消、W17 并非全胜（6 场被对手拿下）。详见 10_数据核对"],
]
ws = build_sheet(wb, "00_导读", DARK, "W17 EDITORIAL — 立项确认稿",
                 "全部数据均可回溯到官方来源；未获公开来源的数据一律标注“暂无可靠公开数据”，不做估算填充。",
                 readme_head, readme_rows, [22, 108], wrap_cols=(2,), row_height=30)

block_row = len(readme_rows) + 7
ws.merge_cells(start_row=block_row, start_column=1, end_row=block_row, end_column=2)
c = ws.cell(row=block_row, column=1, value="三个必须先定的口径（其余都可后调）")
c.fill = PatternFill("solid", fgColor=TEAL)
c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
ws.row_dimensions[block_row].height = 26

gates = [
    ["① 图片版权", "官方 CDN 热链并降级 / 你提供授权图 / 全部改用 Art-direction 占位（银色轮廓 + 网点 + 大字排版）。此决定影响首屏气质与合规性，必须先定。"],
    ["② 数据诚实", "10 个对比维度里只有 3 个有可靠公开数据（排位单圈差、正赛差距、积分与胜场）。其余 7 项（极速、弯速、下压力、风阻、轮胎退化等）无公开来源 → 明确标注，不做雷达图“填满分”。"],
    ["③ 站点结构", "单页滚动长文 + 分站/车手子路由 / 纯单页锚点。前者更像杂志，后者更轻。"],
]
for i, g in enumerate(gates):
    r = block_row + 1 + i
    c = ws.cell(row=r, column=1, value=g[0])
    c.font = Font(name=FONT, size=9, bold=True, color=DARK)
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left", indent=1)
    c = ws.cell(row=r, column=2, value=g[1])
    c.font = Font(name=FONT, size=9, color=DARK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 46

# ---------------------------------------------------------------- 01 设计语言
design_head = ["类别", "Token / 元素", "值", "用途", "约束"]
design_rows = [
    ["颜色", "--ink", "#08090A", "页面底色、页脚", "占比最大，约 55%"],
    ["颜色", "--carbon", "#121417", "区块交替底色", "与 ink 交替制造段落感"],
    ["颜色", "--graphite", "#1C1F23", "图片说明条、图表背景", "不用于大段文字背景"],
    ["颜色", "--silver-700", "#6F767C", "次级文字、图注", "对比度需 ≥ 4.5:1 才可用于正文"],
    ["颜色", "--silver-300", "#C9CED3", "大标题、细线、页码", "品牌银，标题主色"],
    ["颜色", "--silver-100", "#E9ECEE", "深底上的正文", "不用纯白 #FFF"],
    ["颜色", "--teal", "#00A19C", "数据高亮、当前态、图表主色", "全站合计 ≤ 10% 面积"],
    ["颜色", "--teal-bright", "#00D2BE", "悬停、进度条、赛段标记", "仅交互反馈，静态正文禁用"],
    ["颜色", "--paper", "#F4F5F3", "1–2 个“纸张”浅色章节（如技术参数）", "制造杂志翻页感，最多 2 处"],
    ["字体", "Display — Archivo", "wght 600–800 / wdth 100–125", "H1/H2、数字大字报", "全大写 + 负字距（-0.02em ~ -0.04em）"],
    ["字体", "Editorial — Instrument Serif", "400 italic", "引语（pull quote）、章节序号", "仅用于短句，不超过 2 行"],
    ["字体", "Body — Inter Tight", "400 / 500", "正文、图注", "16–19px，行高 1.55，行长 ≤ 68 字符"],
    ["字体", "Data — JetBrains Mono", "500，tabular-nums", "所有数字、时间、单位", "确保数字对齐抖动为零"],
    ["字体", "中文字族", "'Noto Sans SC' / 'PingFang SC' 回退", "中文正文", "与 Inter Tight 视觉同宽"],
    ["字号", "Hero H1", "clamp(56px, 11vw, 200px)", "THE ART OF SPEED.", "行高 0.86"],
    ["字号", "章节 H2", "clamp(40px, 6vw, 96px)", "THE MACHINE 等", "行高 0.9"],
    ["字号", "大字报数字", "clamp(64px, 12vw, 220px)", "772 / 350 / 15,000", "配 12px 单位标签"],
    ["字号", "正文", "clamp(16px, 1.05vw, 19px)", "叙述段落", "与 8px 基线对齐"],
    ["网格", "12 列 / 边距", "5vw（桌面）/ 20px（移动）", "全站栅格", "gutter 24px"],
    ["网格", "非对称分栏", "7/5、8/4、5/7", "图文交错", "禁止 6/6 对称卡片排布"],
    ["网格", "基线", "8px", "所有纵向节奏", "间距只取 8 的倍数"],
    ["动效", "进入", "opacity 0→1, y 24→0, 600ms", "文字与图片上浮", "cubic-bezier(.16,1,.3,1)，只触发一次"],
    ["动效", "视差", "translateY ≤ 8% 视口高", "Hero 与全屏大图", "移动端关闭"],
    ["动效", "图片缩放", "scale 1.06 → 1.00", "慢速 1.2s", "只在进入视口时"],
    ["动效", "图表", "路径描边 / 数值滚动 900ms", "数据章节", "首次进入视口时播放一次"],
    ["动效", "降级", "prefers-reduced-motion", "全部动画", "关闭视差与缩放，仅保留淡入"],
    ["影像", "全屏大图", "16:9 或 21:9 出血", "Hero / 分站专题", "顶部加 0–40% 黑色渐变遮罩保证标题可读"],
    ["影像", "细节图", "4:5 竖构图", "THE DETAILS 画廊", "不规则网格，允许错位 40px"],
    ["影像", "去饱和", "grayscale(0.35) + contrast(1.05)", "非主角图片", "主角图片保持原色"],
    ["禁用", "玻璃拟态 / 发光 / 大面积渐变", "—", "—", "与原稿要求一致，全站禁用"],
    ["禁用", "卡片网格 / 圆角卡片 / 阴影", "—", "—", "改用出血图 + 细线分隔"],
    ["禁用", "自动播放视频音轨", "—", "—", "视频静音、点击播放"],
]
build_sheet(wb, "01_设计语言", "6F767C", "01 — 设计语言：银色、黑、青色与编辑式排版",
            "先定 token 再写组件；所有数值可直接写入 tailwind @theme。",
            design_head, design_rows, [10, 30, 32, 34, 44], wrap_cols=(3, 5), row_height=34)

# ---------------------------------------------------------------- 02 页面结构
page_head = ["章节", "EN 标题", "中文标题", "叙事目标", "内容模块", "交互", "需要的数据", "关键组件", "验收标准"]
page_rows = [
    ["01", "THE ART OF SPEED.", "Hero 首屏", "3 秒内建立“这是杂志，不是官网”的判断",
     "全屏 W17 侧前方大图（低角度）；叠加超大主标题；副标 MERCEDES-AMG PETRONAS / FORMULA ONE TEAM / 2026 — W17；左下期刊信息（ISSUE 01 · 30 OCT 2026）；右下 SCROLL 提示",
     "滚动视差（图慢速上移）；标题逐行 mask 揭示；SCROLL 提示呼吸动画",
     "无（纯视觉）",
     "HeroStage、EditorialMask、ScrollCue、MarqueeStrip",
     "首屏无横向滚动条；标题不遮挡赛车主体；LCP < 2.0s；移动端标题不换行超过 3 行"],
    ["02", "THE MACHINE.", "认识 W17", "用 8 个“专题跨页”讲清楚这辆车，而不是堆参数",
     "①整车 ②前翼与主动空气动力学 ③侧箱 ④底板与扩散器 ⑤悬挂 ⑥动力单元 M17 ⑦轮胎与制动 ⑧驾驶舱与方向盘；每个模块 = 一张大图 + 一个小标题 + 120–180 字解说 + 一个可展开技术注",
     "左右图文交错进入；点击车身热区（前翼/侧箱/底板/尾翼/驾驶舱）切换右侧解说；模块间细线 + 章节序号",
     "8 个技术子系统的定性说明（官方技术页 + 2026 规则）",
     "SectionSpread、HotspotCar、TechNote（展开/收起）、FigureCaption",
     "8 个模块全部有内容且有图（或占位）；热区可键盘 Tab 聚焦；小屏切换为纵向列表"],
    ["03", "TECHNICAL DATA.", "技术参数", "让数字成为排版主角",
     "3 组参数：动力单元（M17 / 350 kW / 15,000 rpm / 4.0 MJ / 9.0 MJ·lap / 185 kg）、车身（772 kg / 5,505 mm / 1,900 mm / 970 mm / 碳纤维单体壳）、传动与底盘（8 速 / OZ 镁合金 / Brembo / Pirelli）",
     "数字滚动动画（进入视口触发一次）；悬停显示“口径说明”浮层；点击复制数值",
     "28 项官方技术参数（见 05_技术参数）",
     "StatBlock、StatGrid、TooltipNote、FootnoteRail",
     "每个数字都有单位与口径注释；772 kg 必须写明“官方页面口径”；不使用普通表格样式"],
    ["04", "THE VICTORIES.", "经典比赛回顾", "证明竞争力，同时不回避失利",
     "横向时间轴：12 场胜利全部入轴（可横向拖动）；其中 6 场做深度专题（澳大利亚 / 摩纳哥 / 银石 / 蒙扎 / 新加坡 / 美国）",
     "横向 scroll-snap 时间轴；点击进入分站专题页（独立路由）；下一站箭头；键盘左右切换",
     "赛果、名次、差距、圈数、关键节点（需逐站补官方赛报）",
     "VictoryTimeline、RaceCard、RaceDetailPage、GapBar",
     "12 场胜利全部可点；6 篇专题各有独立 URL；未举行的分站不得出现在时间轴"],
    ["05", "THE BATTLE.", "横向性能对比", "这是全站可信度的试金石：只展示能核实的东西",
     "四车队对比（W17 / SF-26 / MCL40 / RB22）；3 组可核实图表 + 7 项“暂无公开数据”的诚实说明；4 条赛道剖面（蒙扎 / 摩纳哥 / 银石 / 新加坡）",
     "车队筛选；赛道切换；图表进入视口时描边；悬停显示数据点与来源链接；“暂无数据”用斜纹占位而非留空",
     "排位单圈差（同场次）、正赛差距、积分与胜场（见 09_性能对比）",
     "GapChart（横向条）、LapDeltaLine、RadarBlockedNotice、CircuitProfile",
     "图表每个数据点可追溯来源；任何估算值必须显示“非官方/估算”标签；不得出现无来源的雷达图满分"],
    ["06", "THE DRIVERS.", "车手专题", "人物杂志式的两个跨页",
     "George Russell（#63）与 Kimi Antonelli（#12）：大幅肖像、姓名与车号、赛季战绩、驾驶特点、代表分站",
     "两个全屏交替跨页；肖像 hover 轻微推移；点击进入车手专题页",
     "积分、胜场、杆位、平均完赛名次（可计算）",
     "DriverSpread、PortraitFrame、DriverStatRail",
     "两位车手数据与 08_积分榜一致；肖像缺图时用大型车号排版占位"],
    ["07", "THE DETAILS.", "细节摄影", "专业汽车摄影杂志的质感段落",
     "不规则网格画廊：前翼 / 扩散器 / 悬挂 / 轮胎与刹车 / 驾驶舱 / 涂装 / 夜赛 / 维修区",
     "点击全屏 lightbox（支持键盘、手势）；网格错位；图片懒加载 + 模糊占位",
     "8–12 张细节图（版权见 11_素材清单）",
     "MasonryGallery、Lightbox、ProgressiveImage",
     "lightbox 可用 Esc 关闭；图片失败时显示带图注的占位块，不留白块"],
    ["08", "THE SEASON.", "赛季时间轴", "让人一眼看懂赛季走势，并明确“数据截至何时”",
     "纵向时间轴：赛季开始 → 首胜 → 连胜期 → 中场调整 → 近期赛果 → 当前积分；顶部固定“DATA AS OF 30 OCT 2026”",
     "滚动进度指示；节点 hover 展开赛果；切换“按轮次 / 按时间”",
     "18 个已完成分站 + 5 个未举行分站（明确标注 TBD）",
     "SeasonRail、RoundNode、StandingsMini、DataFreshnessTag",
     "未举行分站绝不写成已完成；积分与官方一致（Antonelli 365 / Mercedes 607）"],
]
build_sheet(wb, "02_页面结构", "00A19C", "02 — 页面结构：8 个章节，一本可交互的杂志",
            "每章只解决一个问题；章节之间用大图与留白断开，避免连续卡片。",
            page_head, page_rows, [6, 22, 16, 28, 54, 40, 26, 30, 40],
            wrap_cols=(4, 5, 6, 7, 8, 9), row_height=150)

# ---------------------------------------------------------------- 03 技术方案
tech_head = ["层", "选型", "版本", "理由", "备选 / 风险"]
tech_rows = [
    ["框架", "React + TypeScript", "19 / 5.x", "组件化、类型安全、生态成熟", "—"],
    ["构建", "Vite", "7.x", "冷启动与 HMR 快，构建产物干净", "Next.js（若日后需要 SEO 静态化）"],
    ["样式", "Tailwind CSS", "v4 (@theme)", "设计 token 直接进 CSS 变量，便于保持杂志一致性", "CSS Modules（可视化调试更直观但更啰嗦）"],
    ["动画", "Framer Motion (motion)", "12.x", "滚动驱动动画、layout 动画开箱可用", "GSAP + ScrollTrigger（控制力更强，体积更大）"],
    ["平滑滚动", "Lenis", "1.x", "杂志式长页面的滚动质感", "原生 scroll + CSS scroll-snap；移动端建议只用原生"],
    ["图表", "d3-scale + d3-shape + 自绘 SVG", "7.x", "条形/折线/赛道剖面需要精确控制排版与标注，库默认样式会破坏杂志感", "Recharts（快但样式受限）；ECharts（过重）"],
    ["路由", "React Router", "7.x", "分站/车手专题需要独立 URL 与滚动恢复", "Hash 路由（部署更简单）"],
    ["数据层", "本地 JSON + zod 校验", "—", "数据量小（18 分站/23 车手），无需后端；zod 保证字段不缺失", "Headless CMS（后期可迁移，字段结构一致）"],
    ["图片", "sharp 预处理 → AVIF/WebP + srcset", "—", "首屏质量与体积的平衡", "Cloudinary/图片 CDN（若需要动态裁剪）"],
    ["字体", "自托管 woff2 + font-display: swap", "—", "避免外部字体请求阻塞", "Google Fonts（更快落地，但隐私与稳定性略差）"],
    ["质量", "ESLint + Vitest + Playwright 截图", "—", "组件单测 + 关键视口截图回归", "仅手动验收（不推荐）"],
    ["部署", "静态站点（任意静态托管）", "—", "纯前端产物，可直接双击 index.html 预览构建后的 dist", "—"],
]
build_sheet(wb, "03_技术方案", "6F767C", "03 — 技术方案",
            "所有选型服务于“排版控制力”，而不是开发速度。",
            tech_head, tech_rows, [12, 32, 16, 46, 46], wrap_cols=(4, 5), row_height=32)

# ---------------------------------------------------------------- 04 文件结构
tree_head = ["路径", "类型", "作用", "备注"]
tree_rows = [
    ["src/main.tsx", "入口", "挂载 App、全局字体与样式", "—"],
    ["src/App.tsx", "路由", "路由表：/ 及各章节、专题页", "章节既可作为锚点，也可作为独立页面"],
    ["src/theme/tokens.css", "样式", "颜色/字号/间距/动效变量", "与 01_设计语言一一对应"],
    ["src/components/editorial/", "组件", "HeroStage、SectionSpread、StatBlock、FigureCaption、PullQuote", "纯排版组件，不含业务逻辑"],
    ["src/components/charts/", "组件", "GapChart、LapDeltaLine、StandingsBar、CircuitProfile", "基于 d3-scale 的自绘 SVG"],
    ["src/components/interactive/", "组件", "HotspotCar、VictoryTimeline、MasonryGallery、Lightbox", "需要状态与手势"],
    ["src/components/motion/", "组件", "Reveal、Parallax、MaskText、useReducedMotion", "统一动画入口，便于整体降级"],
    ["src/sections/hero|machine|data|victories|battle|drivers|details|season", "区块", "8 个章节，每章一个目录", "章节内再拆跨页模块"],
    ["src/pages/RaceDetail.tsx", "页面", "6 篇分站专题共用一个模板", "数据来自 races.json"],
    ["src/pages/DriverDetail.tsx", "页面", "2 篇车手专题", "数据来自 drivers.json"],
    ["src/data/car.ts", "数据", "W17 技术参数与口径说明", "唯一来源，禁止组件内硬编码"],
    ["src/data/races.ts", "数据", "18 个分站赛果 + 5 个未举行分站", "含 verified / source 字段"],
    ["src/data/standings.ts", "数据", "车手与车队积分（截至 2026-10-25）", "含 cutoff 字段"],
    ["src/data/battle.ts", "数据", "可核实对比指标 + 不可得指标清单", "不可得项必须有 reason 字段"],
    ["src/data/sources.ts", "数据", "来源名称、URL、抓取日期", "图表脚注直接引用"],
    ["public/assets/photos/", "静态资源", "整理后的赛车/车手图片（AVIF/WebP）", "版权与来源见 11_素材清单"],
    ["public/assets/placeholders/", "静态资源", "缺图时的 SVG 占位（轮廓 + 网点 + 大字）", "保证无图也像设计稿"],
    ["docs/content-deck.md", "文档", "全部中英文文案（可直接校对）", "与代码解耦，便于改稿"],
    ["docs/data-method.md", "文档", "每个指标的口径、公式、免责说明", "对比章节的“诚实说明书”"],
]
build_sheet(wb, "04_文件结构", "6F767C", "04 — 工程结构：先有数据层，再有页面",
            "数据与文案全部外置，页面只做呈现；后续换图、改数据不需要动组件。",
            tree_head, tree_rows, [58, 12, 52, 44], wrap_cols=(3, 4), row_height=28)

# ---------------------------------------------------------------- 05 技术参数
spec_head = ["分组", "参数", "值", "单位", "官方口径 / 说明", "状态", "来源"]
spec_rows = [
    ["整车", "车型全称", "Mercedes-AMG F1 W17 E PERFORMANCE", "—", "官方车型名，站点标题统一使用", "已核实", "Mercedes-AMG F1 技术参数页"],
    ["整车", "整车重量", "772", "kg", "官方页面“Overall Weight”。具体统计口径（是否含车手/配重）页面未注明，站点需原样标注", "已核实（口径待补）", "Mercedes-AMG F1 技术参数页"],
    ["整车", "总长", "5505", "mm", "官方写法为 Under 5,505 mm", "已核实", "同上"],
    ["整车", "总宽", "1900", "mm", "2026 规则较上一代收窄", "已核实", "同上"],
    ["整车", "总高", "970", "mm", "—", "已核实", "同上"],
    ["车身", "单体壳", "模压碳纤维 + 蜂窝复合结构", "—", "Moulded carbon fibre and honeycomb composite structure", "已核实", "同上"],
    ["车身", "车身覆盖件", "碳纤维复合材料（引擎盖/侧箱/底板/鼻锥/前后翼）", "—", "官方列举", "已核实", "同上"],
    ["车身", "安全结构", "座舱生存舱、前后碰撞结构、Halo 钛合金保护结构", "—", "含六点式安全带与 HANS", "已核实", "同上"],
    ["悬挂", "前悬挂", "碳纤维叉臂 + 推杆驱动弹簧减震", "—", "push-rod activated", "已核实", "同上"],
    ["悬挂", "后悬挂", "碳纤维叉臂 + 推杆驱动内置弹簧减震", "—", "push-rod activated inboard", "已核实", "同上"],
    ["动力单元", "动力单元型号", "Mercedes-AMG F1 M17 E PERFORMANCE", "—", "PU 最低重量 185 kg；额度：ICE/TC/EXH 各 4 套，MGU-K/ES/CE 各 3 套", "已核实", "同上"],
    ["动力单元", "内燃机", "1.6 L V6 涡轮增压，90° 夹角，24 气门", "—", "单级增压器 + 排气涡轮同轴", "已核实", "同上"],
    ["动力单元", "内燃机最高转速", "15,000", "rpm", "2026 规则上限", "已核实", "同上"],
    ["动力单元", "最大燃油流量", "3000", "MJ/h", "能量流率口径，非升/小时", "已核实", "同上"],
    ["动力单元", "喷油压力", "350", "bar", "高压直喷，每缸一个喷嘴", "已核实", "同上"],
    ["动力单元", "涡轮最高转速", "150,000", "rpm", "排气涡轮", "已核实", "同上"],
    ["混动系统", "MGU-K 最大功率", "350", "kW", "约为内燃机 400 kW 的同等量级——这是“接近 50:50”说法的规则基础", "已核实", "同上"],
    ["混动系统", "MGU-K 最高转速", "60,000", "rpm", "—", "已核实", "同上"],
    ["混动系统", "可用储能", "4.0", "MJ", "Max useable energy storage range on track", "已核实", "同上"],
    ["混动系统", "单圈最大回收", "9.0", "MJ/lap", "官方注明为“视赛道而定”（Circuit Specific），不可当固定值使用", "已核实", "同上"],
    ["传动", "变速箱", "8 前速 + 1 倒挡，碳纤维主箱体", "—", "顺序式半自动液压换挡；碳纤维离合器片", "已核实", "同上"],
    ["底盘", "轮毂 / 轮胎", "OZ 锻造镁合金 / Pirelli", "—", "轮胎规格随分站变化，站点不写死", "已核实", "同上"],
    ["底盘", "制动", "碳/碳刹车盘与刹车片，后轴线控制动", "—", "Brembo 整体式镍铝卡钳（前后）", "已核实", "同上"],
    ["底盘", "转向", "助力齿轮齿条", "—", "方向盘为碳纤维结构", "已核实", "同上"],
    ["底盘", "电子系统", "FIA 标准 ECU + 认证电子电气系统", "—", "仪表来自 Motion Applied Technologies", "已核实", "同上"],
    ["油液", "燃油 / 润滑油", "PETRONAS Primax / PETRONAS Syntium、Tutela", "—", "赞助商品牌，站点可露出", "已核实", "同上"],
    ["规则", "主动空气动力学", "前后翼可动（2026 新规）", "—", "取代 DRS 的 Boost / Overtake 模式；车身更短更窄（轴距 -200 mm、宽 -100 mm）", "已核实（厂商口径）", "Mercedes 2026 赛季发布稿"],
    ["规则", "动力配比", "“接近 50:50”的燃油/电能输出", "—", "属厂商表述；站点必须显示“厂商口径”，不可写成规则原文", "标注口径", "Mercedes 2026 赛季发布稿"],
]
build_sheet(wb, "05_技术参数", "00A19C", "05 — W17 技术参数（28 项，逐项注明来源）",
            "“772 kg”与“接近 50:50”是两处最容易被误写的数据：前者口径未注明，后者是厂商表述。",
            spec_head, spec_rows, [10, 24, 44, 8, 62, 18, 30],
            wrap_cols=(3, 5), row_height=36)

# ---------------------------------------------------------------- 06 赛历赛果
race_head = ["轮次", "日期", "大奖赛", "赛道", "冠军", "车队", "梅赛德斯最佳名次", "备注", "核实状态", "来源"]
race_rows = [
    [1, "2026-03-08", "澳大利亚", "Albert Park", "George Russell", "Mercedes", "P1 + P2（Antonelli +2.974s）", "梅赛德斯包揽冠亚军，58 圈", "已核实", "ESPN 赛果"],
    [2, "2026-03-15", "中国", "上海国际赛车场", "Kimi Antonelli", "Mercedes", "P1", "冲刺赛冠军亦为 Antonelli", "已核实", "ESPN 赛果"],
    [3, "2026-03-29", "日本", "铃鹿", "Kimi Antonelli", "Mercedes", "P1", "—", "已核实", "ESPN 赛果"],
    ["—", "2026-04-12", "巴林（原定）", "—", "未举行", "—", "—", "因中东局势取消", "已核实（取消）", "Kazinform / F1 公告"],
    ["—", "2026-04-19", "沙特阿拉伯（原定）", "—", "未举行", "—", "—", "因中东局势取消", "已核实（取消）", "Kazinform / F1 公告"],
    [4, "2026-05-03", "迈阿密", "Miami International Autodrome", "Kimi Antonelli", "Mercedes", "P1", "冲刺赛冠军为 Norris", "已核实", "ESPN 赛果"],
    [5, "2026-05-24", "加拿大", "Circuit Gilles Villeneuve", "Kimi Antonelli", "Mercedes", "P1（杆位 Russell）", "—", "已核实", "ESPN 赛果"],
    [6, "2026-06-07", "摩纳哥", "Circuit de Monaco", "Kimi Antonelli", "Mercedes", "P1", "78 圈，冠军用时 2:23:31.243", "已核实", "ESPN 赛果"],
    [7, "2026-06-14", "西班牙（巴塞罗那）", "Circuit de Catalunya", "Lewis Hamilton", "Ferrari", "杆位 Russell，正赛未胜", "W17 失手之一", "已核实", "ESPN 赛果"],
    [8, "2026-06-28", "奥地利", "Red Bull Ring", "George Russell", "Mercedes", "P1", "—", "已核实", "ESPN 赛果"],
    [9, "2026-07-05", "英国", "Silverstone", "Charles Leclerc", "Ferrari", "杆位 Antonelli，正赛未胜", "高速弯段落的关键失利", "已核实", "ESPN 赛果"],
    [10, "2026-07-19", "比利时", "Spa-Francorchamps", "Kimi Antonelli", "Mercedes", "P1", "—", "已核实", "ESPN 赛果"],
    [11, "2026-07-26", "匈牙利", "Hungaroring", "Lando Norris", "McLaren", "未胜", "—", "已核实", "ESPN 赛果"],
    [12, "2026-08-23", "荷兰", "Zandvoort", "Lando Norris", "McLaren", "未胜", "—", "已核实", "ESPN 赛果"],
    [13, "2026-09-06", "意大利", "Monza", "Kimi Antonelli", "Mercedes", "P1（杆位为 Alpine 的 Gasly）", "非杆位取胜，含策略价值", "已核实", "ESPN 赛果"],
    [14, "2026-09-13", "西班牙（马德里）", "Madring", "Kimi Antonelli", "Mercedes", "P1", "原稿漏列此胜", "已核实", "ESPN 赛果"],
    [15, "2026-09-26", "阿塞拜疆", "Baku City Circuit", "George Russell", "Mercedes", "P1", "原稿漏列此胜", "已核实", "ESPN 赛果"],
    [16, "2026-10-04", "巴林（改至雪邦）", "Sepang International Circuit", "Max Verstappen", "Red Bull", "冲刺赛冠军 Antonelli", "巴林站改由马来西亚承办", "已核实", "ESPN 赛果 / F1 公告"],
    [17, "2026-10-11", "新加坡", "Marina Bay", "Max Verstappen", "Red Bull", "P3（Antonelli +24.573s）", "高温轮胎管理段落，冲刺周末", "已核实", "ESPN 赛果"],
    [18, "2026-10-25", "美国", "COTA", "Kimi Antonelli", "Mercedes", "P1（杆位，+13.078s）", "Russell 第 8（+54.303s）", "已核实", "ESPN 赛果"],
    [19, "2026-11-01", "墨西哥", "Autódromo Hermanos Rodríguez", "未举行", "—", "—", "本周末进行中（FP1 已开始）", "未举行", "Sky Sports 赛程"],
    [20, "TBD", "巴西", "Interlagos", "未举行", "—", "—", "日期待核实", "未举行", "待抓取"],
    [21, "TBD", "拉斯维加斯", "Las Vegas Strip Circuit", "未举行", "—", "—", "日期待核实", "未举行", "待抓取"],
    [22, "TBD", "卡塔尔", "Lusail", "未举行", "—", "—", "日期待核实", "未举行", "待抓取"],
    [23, "TBD", "阿布扎比", "Yas Marina", "未举行", "—", "—", "日期待核实；收官站", "未举行", "待抓取"],
]
build_sheet(wb, "06_赛历赛果", "00A19C", "06 — 2026 赛历与赛果（截至美国站）",
            "18 个已完成分站 + 5 个未举行分站。未举行的分站以浅色标注，绝不能进入“胜利时间轴”。",
            race_head, race_rows, [6, 13, 20, 26, 18, 12, 32, 36, 14, 22],
            wrap_cols=(4, 7, 8), row_height=30, num_formats={1: "0"})

# ---------------------------------------------------------------- 07 关键分站
key_head = ["编号", "分站", "日期", "赛道 / 特征", "叙事角度", "已核实事实", "要展示的数据", "图片需求", "仍需补充"]
key_rows = [
    ["01", "澳大利亚大奖赛", "2026-03-08", "Albert Park / 半街道，中低速弯多",
     "新规则时代的第一场：W17 用一次 1-2 直接定调",
     "Russell P1、Antonelli P2，差距 2.974s；58 圈",
     "冠亚军差距条；两位车手起步位置；与第三名 Leclerc 的差距（+15.519s）",
     "Russell 与 Antonelli 赛后合影；W17 在 Albert Park 的侧后 3/4 视角",
     "官方赛报：安全车次数、关键超越、轮胎策略"],
    ["02", "摩纳哥大奖赛", "2026-06-07", "Monaco / 极低速街道赛，机械抓地力优先",
     "最不像“新车”的一场胜利：低速段与出弯牵引力的证据",
     "Antonelli P1；78 圈；冠军用时 2:23:31.243",
     "单圈分布提示（若可得）；发车位置；与第二名差距",
     "Antonelli 在 Casino 弯的正面图；W17 通过隧道的动态图",
     "排位赛 Q3 圈速对比；轮胎策略（一停/两停）"],
    ["03", "英国大奖赛", "2026-07-05", "Silverstone / 高速连续弯，对气动效率最敏感",
     "诚实的一章：杆位是 W17 的，冠军是 Ferrari 的",
     "Antonelli 杆位；Leclerc 夺冠；Hamilton 亦在列",
     "杆位到冠军的名次变化；与冠军的完赛差距（需补）",
     "Silverstone 高速弯中的 W17；看台与烟雾的氛围图",
     "正赛各车手完赛差距表；关键进站窗口"],
    ["04", "意大利大奖赛", "2026-09-06", "Monza / 极低阻，直线效率与制动稳定性",
     "非杆位取胜：说明 W17 在“直线 + 重刹”组合上的比赛能力",
     "Antonelli P1；杆位为 Alpine 的 Gasly；53 圈",
     "发车位置 → 冠军的位移；与杆位车手的完赛差距（需补）",
     "Monza 一号弯重刹区；Parabolica 出弯；尾翼细节",
     "测速点极速数据（仅官方公布时使用，否则标注暂无）"],
    ["05", "新加坡大奖赛", "2026-10-11", "Marina Bay / 高温高湿，轮胎退化与体能",
     "短板的证据：同一条赛道，W17 落后 24 秒",
     "Verstappen P1（冲刺周末，+33 分）；Leclerc +7.335s；Antonelli P3 +24.573s；Russell P9（+84.677s，从 P19 起跑）",
     "前三名差距梯度；Russell 的名次位移",
     "夜赛灯光下的 W17；车手在座舱内的特写",
     "轮胎退化曲线所需逐圈数据（公开源有限，若无则仅作定性描述）"],
    ["06", "美国大奖赛", "2026-10-25", "COTA / 高低差 + 长直道，综合考验",
     "赛季最强音：杆位 + 13 秒优势夺冠",
     "Antonelli 杆位并夺冠；Leclerc +13.078s；Verstappen +17.485s；Hamilton +18.515s；Russell P8（+54.303s）",
     "前六名差距阶梯图；梅赛德斯两位车手的分差对比",
     "Antonelli 冲线瞬间；COTA 大直道看台；车队庆祝",
     "赛后官方赛报（首圈事件、策略分化）"],
]
build_sheet(wb, "07_关键分站", "00A19C", "07 — 6 篇分站深度专题（其余 6 场只入时间轴）",
            "每一篇都必须有一个“可被数据支撑的判断”，而不是形容词堆叠。",
            key_head, key_rows, [6, 18, 12, 32, 46, 54, 40, 32, 32],
            wrap_cols=(4, 5, 6, 7, 8, 9), row_height=88)

# ---------------------------------------------------------------- 08 积分榜
ws = wb.create_sheet("08_积分榜")
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = "00A19C"

ws.merge_cells("A1:E1")
c = ws["A1"]
c.value = "08 — 2026 积分榜（截至 2026-10-25 美国站）"
c.fill = PatternFill("solid", fgColor=DARK)
c.font = Font(name=FONT, size=15, bold=True, color=SILVER_TEXT)
c.alignment = Alignment(vertical="center", indent=1)
ws.row_dimensions[1].height = 36

ws.merge_cells("A2:E2")
c = ws["A2"]
c.value = "来源：Mercedes-AMG F1 官方 standings 页（2026）。车手总分 2034、车队总分 2034，双方一致，可作为交叉校验。"
c.font = Font(name=FONT, size=9, italic=True, color=MUTED)
c.alignment = Alignment(vertical="center", indent=1)

drivers = [
    (1, "Kimi Antonelli", "Mercedes-AMG", 365, "ITA"),
    (2, "George Russell", "Mercedes-AMG", 242, "GBR"),
    (3, "Lewis Hamilton", "Ferrari", 241, "GBR"),
    (4, "Max Verstappen", "Red Bull Racing", 236, "NLD"),
    (5, "Charles Leclerc", "Ferrari", 233, "MCO"),
    (6, "Lando Norris", "McLaren", 211, "GBR"),
    (7, "Oscar Piastri", "McLaren", 144, "AUS"),
    (8, "Isack Hadjar", "Red Bull Racing", 105, "FRA"),
    (9, "Liam Lawson", "Racing Bulls", 77, "NZL"),
    (10, "Pierre Gasly", "Alpine", 42, "FRA"),
    (11, "Arvid Lindblad", "Racing Bulls", 38, "GBR"),
    (12, "Franco Colapinto", "Alpine", 27, "ARG"),
    (13, "Oliver Bearman", "TGR Haas", 23, "GBR"),
    (14, "Nico Hülkenberg", "Audi", 13, "DEU"),
    (15, "Gabriel Bortoleto", "Audi", 10, "BRA"),
    (16, "Esteban Ocon", "TGR Haas", 7, "FRA"),
    (17, "Fernando Alonso", "Aston Martin", 7, "ESP"),
    (18, "Carlos Sainz", "Williams", 7, "ESP"),
    (19, "Alexander Albon", "Williams", 5, "THA"),
    (20, "Yuki Tsunoda", "Red Bull Racing", 1, "JPN"),
    (21, "Lance Stroll", "Aston Martin", 0, "CAN"),
    (22, "Valtteri Bottas", "Cadillac", 0, "FIN"),
    (23, "Sergio Pérez", "Cadillac", 0, "MEX"),
]
teams = [
    (1, "Mercedes-AMG PETRONAS F1 Team", 607),
    (2, "Ferrari", 474),
    (3, "Red Bull Racing", 355),
    (4, "McLaren", 355),
    (5, "Racing Bulls", 102),
    (6, "Alpine", 69),
    (7, "TGR Haas", 30),
    (8, "Audi", 23),
    (9, "Williams", 12),
    (10, "Aston Martin", 7),
    (11, "Cadillac", 0),
]


def section_title(ws, row, text, span=5):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = PatternFill("solid", fgColor=TEAL)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 24


def table(ws, start_row, headers, rows, widths_map):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=i, value=h)
        c.fill = PatternFill("solid", fgColor=LIGHT_HDR)
        c.font = Font(name=FONT, size=9, bold=True, color=DARK)
        c.border = Border(left=thin, right=thin, top=thin,
                          bottom=Side(style="medium", color=TEAL))
        c.alignment = Alignment(vertical="center")
    for r, row in enumerate(rows, start=start_row + 1):
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=FONT, size=9, color=DARK)
            c.border = box
            c.alignment = Alignment(vertical="center")
            if (r - start_row) % 2 == 0:
                c.fill = PatternFill("solid", fgColor=ZEBRA)
        ws.row_dimensions[r].height = 18
    for col, w in widths_map.items():
        ws.column_dimensions[col].width = w


section_title(ws, 4, "车手积分榜 / DRIVERS' STANDINGS")
table(ws, 5, ["名次", "车手", "车队", "积分", "国籍"], drivers,
      {"A": 8, "B": 30, "C": 27, "D": 10, "E": 9})
for r in range(6, 6 + len(drivers)):
    ws.cell(row=r, column=4).alignment = Alignment(vertical="center", horizontal="right")
    ws.cell(row=r, column=1).alignment = Alignment(vertical="center", horizontal="right")

section_title(ws, 30, "车队积分榜 / CONSTRUCTORS' STANDINGS")
table(ws, 31, ["名次", "车队", "积分"], teams,
      {"A": 8, "B": 30, "C": 27})
for r in range(32, 32 + len(teams)):
    ws.cell(row=r, column=3).alignment = Alignment(vertical="center", horizontal="right")
    ws.cell(row=r, column=1).alignment = Alignment(vertical="center", horizontal="right")

ws.merge_cells("A45:E46")
ws.row_dimensions[45].height = 16
ws.row_dimensions[46].height = 16
c = ws["A45"]
c.value = ("阅读提示：Antonelli 领先 Russell 123 分、领先非梅赛德斯车手 Hamilton 124 分；车队榜领先 Ferrari 133 分。"
          "墨西哥站（11-01）后必须重算——站点因此需要“数据截止”标签，不能把数字写死。")
c.font = Font(name=FONT, size=9, color=DARK)
c.alignment = Alignment(wrap_text=True, vertical="top")
c.fill = PatternFill("solid", fgColor="F0FAF9")

chart1 = BarChart()
chart1.type = "bar"
chart1.style = 10
chart1.title = "车手积分 Top 10"
chart1.y_axis.majorGridlines = None
chart1.x_axis.scaling.orientation = "maxMin"
chart1.legend = None
data = Reference(ws, min_col=4, min_row=5, max_row=15)
cats = Reference(ws, min_col=2, min_row=6, max_row=15)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 9
chart1.width = 13
chart1.series[0].graphicalProperties.solidFill = TEAL
chart1.series[0].graphicalProperties.line.noFill = True
ws.add_chart(chart1, "G4")

chart2 = BarChart()
chart2.type = "bar"
chart2.style = 10
chart2.title = "车队积分（11 支）"
chart2.y_axis.majorGridlines = None
chart2.x_axis.scaling.orientation = "maxMin"
chart2.legend = None
data2 = Reference(ws, min_col=3, min_row=31, max_row=42)
cats2 = Reference(ws, min_col=2, min_row=32, max_row=42)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 9
chart2.width = 13
chart2.series[0].graphicalProperties.solidFill = "9AA1A7"
chart2.series[0].graphicalProperties.line.noFill = True
ws.add_chart(chart2, "G26")

# ---------------------------------------------------------------- 09 性能对比
battle_head = ["维度", "可核实指标", "公开数据源", "可得性", "图表形式", "口径与免责说明"]
battle_rows = [
    ["1 排位赛单圈", "同一分站 Q3 单圈 / 与杆位差值（%）", "Formula 1 官方排位赛成绩；FIA 计时", "可得（逐站采集）",
     "横向差距条（同场次、同阶段）", "只比同一条赛道的同一节次；跨赛道绝对值无意义"],
    ["2 正赛平均圈速", "逐圈数据 → 平均圈速", "官方逐圈计时需逐站抓取且含进站圈干扰", "部分可得（需清洗）",
     "折线（剔除进站圈后）", "必须剔除安全车、进站圈与雨胎圈；未清洗前不展示"],
    ["3 直道极速", "官方测速点极速（km/h）", "F1 官方在部分分站公布 speed trap", "部分可得",
     "柱状（同分站同测速点）", "同一测速点才能比较；下压力设定不同会直接改变结果"],
    ["4 弯道速度", "弯心最低速 / 弯道段耗时", "无公开遥测", "暂无可靠公开数据",
     "不展示数值（改用赛道剖面 + 定性文字）", "任何“弯速排名”都需遥测，公开源不具备，禁止编造"],
    ["5 加速与能量部署", "出弯加速、能量回收与投放曲线", "无公开数据", "暂无可靠公开数据",
     "示意图（标注为“机制示意，非实测”）", "2026 部署规则复杂，厂队策略不公开"],
    ["6 制动表现", "制动点、制动距离、稳定性", "无公开数据", "暂无可靠公开数据",
     "不展示", "仅可引用车手/工程师公开评论，且须标注为主观表述"],
    ["7 轮胎退化", "长距离 stint 圈速斜率（s/lap）", "可由逐圈数据推算；Pirelli/F1 分析文章偶有结论", "间接可得",
     "折线（标注推算方法与样本量）", "受油量、温度、交通影响；只能同场次同配方比较"],
    ["8 空气动力学效率", "Cd / Cl、气动效率", "车队不公开", "暂无可靠公开数据",
     "不展示（用设计特征对比图替代）", "公开的“下压力数字”几乎都是第三方估算，必须标注非官方"],
    ["9 赛道条件适应", "同场次不同温度/路面下的相对表现", "由成绩差异定性推断", "定性可得",
     "赛道剖面卡片 + 定性结论", "只能描述“在该条件下表现如何”，不能反推整车能力"],
    ["10 整体竞争力", "积分、胜场、平均完赛名次、退赛数", "官方积分榜与赛果", "可得",
     "柱状 + 累积积分折线", "积分受策略与可靠性影响，不等于纯车速"],
    ["补充：赛道剖面", "蒙扎（直线/重刹）、摩纳哥（低速/机械抓地）、银石（高速弯）、新加坡（高温/退化）", "官方赛果 + 赛道数据",
     "可得（定性 + 名次）", "四张赛道剖面卡", "每条赛道说明“这个场地考验什么”，再给该场的名次事实"],
]
build_sheet(wb, "09_性能对比", "B4541E", "09 — THE BATTLE：先定义“能证明什么”，再画图",
            "10 个维度中 3 个有可靠公开数据、3 个部分可得、4 个无公开来源。后 4 项用斜纹占位与文字说明替代图表。",
            battle_head, battle_rows, [20, 34, 40, 18, 34, 54],
            wrap_cols=(2, 3, 5, 6), row_height=64)

# ---------------------------------------------------------------- 10 数据核对
check_head = ["序号", "原稿断言 / 参数", "核实结果", "判定", "处置建议", "证据"]
check_rows = [
    ["R01", "澳大利亚站：Russell 冠军，Mercedes 包揽冠亚军", "Russell P1、Antonelli P2（+2.974s），58 圈；第三名 Leclerc +15.519s", "属实", "可直接使用，补充差距数据", "ESPN 2026 澳大利亚站赛果"],
    ["R02", "中国站：Antonelli 冠军", "Antonelli P1；冲刺赛冠军亦为 Antonelli", "属实", "可使用", "ESPN 2026 中国站赛果"],
    ["R03", "日本站：Antonelli 冠军", "Antonelli P1（杆位）", "属实", "可使用", "ESPN 2026 日本站赛果"],
    ["R04", "迈阿密站：Antonelli 冠军", "Antonelli P1", "属实", "可使用", "ESPN 2026 迈阿密站赛果"],
    ["R05", "加拿大大奖赛：Antonelli 冠军", "Antonelli P1，杆位为 Russell", "属实", "可使用，注明杆位归属", "ESPN 2026 加拿大站赛果"],
    ["R06", "摩纳哥站：Antonelli 冠军", "Antonelli P1，78 圈，2:23:31.243", "属实", "可使用", "ESPN 2026 摩纳哥站赛果"],
    ["R07", "奥地利站：Russell 冠军", "Russell P1", "属实", "可使用", "ESPN 2026 奥地利站赛果"],
    ["R08", "比利时站：Antonelli 冠军", "Antonelli P1", "属实", "可使用", "ESPN 2026 比利时站赛果"],
    ["R09", "意大利站：Antonelli 冠军", "Antonelli P1；杆位为 Alpine 的 Gasly", "属实", "可使用，强调“非杆位取胜”", "ESPN 2026 意大利站赛果"],
    ["R10", "美国站：Antonelli 冠军，领先 Leclerc 13.078s", "Antonelli P1（杆位）；Leclerc +13.078s；Verstappen +17.485s", "属实", "可直接使用", "ESPN 2026 美国站赛果"],
    ["R11", "（原稿未列）西班牙站 09-13", "Antonelli 夺冠（Madring，杆位为 Norris）", "原稿遗漏", "必须补入胜利列表与时间轴", "ESPN 2026 西班牙站赛果"],
    ["R12", "（原稿未列）阿塞拜疆站 09-26", "Russell 夺冠（Baku）", "原稿遗漏", "必须补入", "ESPN 2026 阿塞拜疆站赛果"],
    ["R13", "（原稿未提）4 月巴林 / 沙特站", "两站因中东局势取消；巴林站改由马来西亚雪邦于 10-04 承办，Verstappen 夺冠", "原稿缺项", "赛季时间轴必须体现“取消 + 改址”，否则时间轴不成立", "Kazinform 2026-07-28 报道（引 F1 公告）"],
    ["R14", "（原稿的隐含假设）W17 一路压制", "18 场中 Mercedes 12 胜，但 6 场被对手拿下：Hamilton（巴塞罗那）、Leclerc（银石）、Norris（匈牙利/荷兰）、Verstappen（雪邦/新加坡）", "需要修正叙事", "对比章节必须写短板，否则可信度崩", "ESPN 全季赛果"],
    ["R15", "772 kg 最低重量", "官方技术页 Overall Weight 772 kg，但未注明统计口径", "属实（口径待补）", "站点写“官方页面口径”，不写“最低重量指标”", "Mercedes-AMG F1 W17 技术参数页"],
    ["R16", "MGU-K 350 kW / 15,000 rpm / 8 速", "全部与官方一致；另有 4.0 MJ 储能、9.0 MJ/圈回收、PU 185 kg", "属实", "可扩展为完整参数组", "同上"],
    ["R17", "“接近 50:50 的电能与燃油输出”", "属厂商表述（MGU-K 350 kW vs ICE 约 400 kW 的量级关系）", "厂商标述", "必须标注“厂商口径”，不能写成规则原文", "Mercedes 2026 赛季发布稿"],
    ["R18", "原稿数据截止", "2026-10-30；但墨西哥站（11-01）已在进行，Agenda 需显示“进行中”", "属实", "全站加 DATA AS OF 标签，避免数据过期", "Sky Sports 赛程"],
    ["R19", "原稿参考站 landonorris.com", "仅作版式参考，不可复制素材与代码", "合规要求", "站点使用独立版式语言与自有图片", "法律与品牌规范"],
    ["R20", "图片素材（素材包内 5 个 Mercedes CDN 链接）", "当前环境无法访问外网下载，且为品牌版权图", "待定", "见 11_素材清单：需你确认用图策略", "本地素材包 + 网络探测结果"],
]
build_sheet(wb, "10_数据核对", "B4541E", "10 — 数据核对：原稿 20 条断言逐条验证",
            "结论：10 条赛果全部属实，但漏 2 场冠军、漏 2 站取消与改址、并隐含“全胜”叙事——这三点必须先改。",
            check_head, check_rows, [6, 40, 56, 15, 38, 32],
            wrap_cols=(2, 3, 5, 6), row_height=48)

# ---------------------------------------------------------------- 11 素材清单
asset_head = ["素材", "类型", "来源 / URL", "用于章节", "版权与使用", "降级方案"]
asset_rows = [
    ["W17 主视觉（整车 3/4 侧前）", "照片", "素材包：Mercedes-AMG 官方 W17 页面 CDN（0df76a22…-2704x2028.webp）", "01 Hero", "Mercedes-AMG 版权；公开发布需授权或使用官方媒体库条款", "银色剪影 + 网点 + THE ART OF SPEED 大字排版"],
    ["W17 俯视", "照片", "素材包：2a0ce1f5…-1920x2560.webp", "02 整车/底板", "同上", "线稿结构图（SVG，标注为示意图）"],
    ["W17 维修区", "照片", "素材包：13b5c779…-2704x2028.webp", "08 赛季/07 细节", "同上", "深灰底 + 车队字体排版块"],
    ["W17 正面", "照片", "素材包：873e93d7…-2704x2028.webp", "02 前翼/主动气动", "同上", "同上"],
    ["W17 驾驶舱", "照片", "素材包：051f5974…-2704x2028.webp", "02 驾驶舱/07 细节", "同上", "同上"],
    ["赛道实拍图库", "图集链接", "F1 官方 gallery（2026 赛车首次上赛道）", "04 分站/07 细节", "Formula 1 版权", "改用无人物赛道氛围图或纯排版"],
    ["对手赛车图库", "图集链接", "Motorsport.com：2026 全车阵图集", "05 对比", "Motorsport 版权", "用车型名称 + 车队色块的排版对照"],
    ["车手肖像", "照片", "F1 官方车手图（media.formula1.com 2026 mercedes 路径）", "06 车手", "Formula 1 / 车队版权", "大型车号 63 / 12 的字体海报"],
    ["字体", "字体文件", "Archivo / Instrument Serif / Inter Tight / JetBrains Mono（均为开源）", "全站", "SIL OFL，可商用", "系统字体回退栈"],
    ["图表数据", "数据", "F1 官方成绩页、Mercedes 官方积分页、ESPN 赛果（均已核对）", "04/05/08", "事实数据，注明来源即可", "—"],
]
build_sheet(wb, "11_素材清单", "6F767C", "11 — 素材与版权（最需要你先拍板的一段）",
            "当前运行环境无法访问外网下载图片；若站点要公开，图片授权必须先解决。三条路线见 12_待确认 Q2。",
            asset_head, asset_rows, [26, 10, 56, 20, 40, 40], wrap_cols=(3, 5, 6), row_height=50)

# ---------------------------------------------------------------- 12 待确认
q_head = ["编号", "待确认", "选项", "推荐", "影响", "若你不回复，我按此执行"]
q_rows = [
    ["Q1", "站点语言", "A. 英文标题 + 中文正文（双语混排）／B. 纯英文／C. 纯中文", "A", "决定文案工作量与字体回退方案；英文标题是杂志感的关键", "按 A 执行：英文 Display + 中文正文"],
    ["Q2", "图片策略", "A. 你提供授权图／B. 先热链官方 CDN + 缺图降级／C. 全部改用 Art-direction 占位（轮廓 + 网点 + 大字）", "B + C 组合", "决定首屏气质与合规风险；B 在浏览器里能显示，但版权未清", "按 B + C：能加载则用，失败自动降级为设计占位"],
    ["Q3", "页面结构", "A. 单页滚动 + 分站/车手子路由／B. 纯单页锚点", "A", "决定是否做 8 个独立页面与滚动恢复", "按 A 执行"],
    ["Q4", "对比数据口径", "A. 只展示可核实数据，其余明确写“暂无公开数据”／B. 引入第三方估算并标注“非官方”", "A", "决定 THE BATTLE 章节的可信度", "按 A 执行（与你原稿要求一致）"],
    ["Q5", "赛季数据更新方式", "A. 本地 JSON 手动更新（含 cutoff 字段）／B. 接公开 API 自动更新", "A", "决定“不能把未举行比赛写成已完成”如何落地", "按 A 执行：JSON + 全站 DATA AS OF 标签"],
    ["Q6", "分站专题篇数", "A. 6 篇深度（澳/摩纳哥/银石/蒙扎/新加坡/美国）／B. 10 篇全覆盖", "A", "决定工期与素材量", "按 A 执行，其余 6 场只入时间轴"],
]
build_sheet(wb, "12_待确认", "00A19C", "12 — 待确认事项（6 条，确认后即开工）",
            "前 3 条属于“必须先定”，后 3 条我已有明确默认值，不回复也能推进。",
            q_head, q_rows, [6, 16, 56, 14, 44, 40], wrap_cols=(3, 5, 6), row_height=54)

wb.save("W17_Editorial_确认稿.xlsx")
print("saved")
