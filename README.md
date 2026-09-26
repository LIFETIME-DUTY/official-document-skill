# official-doc-writer（党政机关公文写作技能）

**本skill在LeoYeAI的official-doc-writer的基础上改进而来。原skill：**[official-doc-writer Agent Skill | LeoYeA/openclaw-~07mznsp](https://skillsmp.com/zh/creators/leoyeai/openclaw-master-skills/skills-official-doc-writer)



按《党政机关公文处理工作条例》和 **GB/T 9704—2012《党政机关公文格式》** 起草、排版、校验并导出
15 种法定文种的 Word 公文（.docx），面向 AI 助手 / DSH / Trae / Claude 等支持 skill 的环境。

## 概述

本技能用于生成符合国家标准与条例的党政机关公文，支持 15 种法定文种，并可导出为格式规范的Word 文档：版心、字体字号、层次序数、分隔线、页码、版记全部按国标落地；要素不合法时校验器会直接报错。

## 功能特性

### 支持的公文类型

| 公文类型 | 行文方向 | 结语 | 适用场景 |
| --- | --- | --- | --- |
| 决议 | 下行/会议 | 无 | 会议审议通过的重大事项 |
| 决定 | 下行 | 无 | 决策部署、表彰奖惩 |
| 命令（令） | 下行 | 无 | 公布规章、宣布强制性措施 |
| 公报 | 公开发布 | 无 | 刊登正式文件、人事任免 |
| 公告 | 公开发布 | 特此公告。 | 向社会公布重要事项 |
| 通告 | 公开发布 | 特此通告。 | 公布应遵守事项 |
| 意见 | 上/下/平行 | 无 | 对重要问题提出见解和办法 |
| 通知 | 下行 | 特此通知。 | 部署工作、发布传达事项 |
| 通报 | 下行 | 特此通报。 | 表彰先进、批评错误、告知情况 |
| 报告 | 上行 | 特此报告。 | 汇报工作、反映情况 |
| 请示 | 上行 | 妥否，请批示。 | 请求指示、批准 |
| 批复 | 下行 | 此复。 | 答复下级机关请示 |
| 议案 | 上行/会议 | 请予审议。 | 提请人大审议事项 |
| 函 | 平行 | 请予支持为盼。 | 不相隶属机关商洽工作 |
| 纪要 | 内部 | 无 | 记载会议情况和议定事项 |

### 核心功能

1. **先问后写，一次问全** —— 用户没说清要哪一份公文就先问，**第一问固定是“要 Word、PDF，还是两个都要？”**
   （没问清格式脚本会直接报错、退出码 2，不产出任何文件）；其余必答项为文种、发文机关、主送机关、事由、正文要点、
   成文日期、上行文还要签发人与发文机关署名：**必答项在同一次提问里全部列出，请用户一轮答完；可选项一并提示“有就提供，没有我就不加”**，
   不拆成多轮挤牙膏、不因缺可选项停下、不替用户编造签发人或日期。缺任一必答项时脚本返回退出码 2、**零文件产出**。
   清单与话术见 `references/必答要素与可选项清单.md`，也可用 `scripts/gongwen.py --ask --type 通知` 直接打印。
2. **用户说什么就生成什么** —— `examples/` 与 `examples/generated/` 只是技能自带样例，用来看字段写法和版式效果，
   **不会被拿来套用或补全用户的公文**；一次只交付一份用户明确要求的公文，不加用户没说的附件、抄送、附注等。
3. **交付格式由用户定** —— `--output-format word` 只给 `×××.docx`；`pdf` 只给 `×××.pdf`（先写临时 docx 再转换，
   交付目录里不留 docx）；`both` 给这两份。除用户点名的格式外没有中间文件、规划 JSON、预览图或附赠说明；
   PDF 由脚本调用打包的 LibreOffice Kit 导出（`--pdf`／`--no-pdf` 是旧写法，等价于 `both`／`word`）。
4. **自动格式排版** —— `scripts/gongwen.py` 按 GB/T 9704—2012 落地全部要素：版心 156mm×225mm、每面 22 行每行 28 字、
   标题 2 号小标宋、正文 3 号仿宋、层次序数黑体/楷体/仿宋、版头红色分隔线、版记三条分隔线、版心外页码、附件另面编排。
5. **红头位置精确落位** —— 发文机关标志上边缘至版心上边缘 35mm（命令 20mm、纪要 35mm、信函至上页边 30mm），
   段前距由“目标位置 − 标志行高 − 前面要素占的行数”算出，行高优先取字体文件的真实纵向度量；
   生成报告会直接给出 `layout.issuer_mark.top_from_content_mm` 等实测值。
6. **规范校验（三层）** —— `scripts/gongwen.py --validate` 做要素层（文种、要素完整性、发文字号与成文日期格式、结语匹配、
   层次序数连续性、语言与标点禁忌）＋版式层（页面页边距、固定行距、字体白名单、字号、缩进、页码域）；
   `scripts/gongwen.py --geometry` 做几何层，不渲染就能从 OOXML 量出红头位置、版心、天头、行网格并逐条核对国标。
7. **公文字体支持** —— 优先标准公文字体，缺失时自动回退并给出警告；`scripts/install_fonts.py` 负责检测与安装。

### 对话交互要素

**核心要素（必填）**：公文类型、发文机关、主送机关（公告/通告/纪要等除外）、公文标题、成文日期、正文内容。

**可选要素**：发文字号、密级和保密期限、紧急程度、签发人（上行文必填）、附件说明与附件正文、抄送机关、印发机关和印发日期、附注。

## 快速开始

### 方式一：在 AI 助手里调用

```
帮我写一份关于开展2026年度政务数据资源目录编制工作的通知
```

### 方式二：运行 Python 脚本

```bash
cd official-doc-writer
python scripts/gongwen.py --ask --type 通知            # 一次问全：第一问是交付格式
python scripts/gongwen.py --needs --type 通知          # 只看还缺哪些必答项
python scripts/gongwen.py 我的要素.json -o 通知.docx --output-format both   # 要 Word + PDF
python scripts/gongwen.py 我的要素.json -o 通知.docx --output-format word   # 只要 Word
python scripts/gongwen.py 我的要素.json -o 通知.docx --output-format pdf    # 只要 PDF
python scripts/gongwen.py 我的要素.json --validate 通知.docx                # 要素层 + 版式层 + 几何层
python scripts/gongwen.py --geometry 通知.docx                              # 只量尺寸
```

**交付格式必须显式给**（`--output-format word|pdf|both`，或不写这个参数而在要素 JSON 里写
`"output_format": "both"`）；**两处都没有时脚本拒绝生成**（退出码 2、零文件），因为技能要求先问清用户要哪种。
PDF 用随环境打包的 LibreOffice Kit 转换，不搜索系统 LibreOffice；`--node` / `--libreoffice-cli` 可指定路径，
确需系统 LibreOffice 时加 `--allow-system-soffice --soffice 路径`。

> `examples/*.json` 是技能自带的**样例**，只在你想看字段怎么写时才打开；
> 生成用户公文时不要拿它们的内容去套，缺什么要素就问用户要。

### 方式三：问答式向导（不熟悉字段时用）

```bash
python scripts/prompts.py -o 通知.docx --save 通知要素.json
```


## 目录结构

```
official-doc-writer/
├── SKILL.md                     技能主文件（AI 首先读它：四条铁律、六步流程、红头位置、格式红线）
├── README.md                    本文件
├── LICENSE                      生成文件安全性声明
├── dev.py                       维护：语法自检 + 打包 + 发布（不参与公文生成）
├── scripts/
│   ├── gongwen.py               排版引擎 + 要素/版式校验 + 几何核对 + 提问清单 + 字体检测
│   ├── prompts.py               问答式要素收集 + 写作提示（原 dialog_manager + smart_prompts）
│   └── install_fonts.py         字体检测、复制与安装
├── references/                  按需读取的规范与语料
│   ├── GBT_9704-2012_党政机关公文格式.md     国标全文（唯一的格式依据）
│   ├── 党政机关公文处理工作条例.md            条例全文 42 条
│   ├── 文种写作规范.md                        15 文种模板、开头结语、常见错误、范例
│   ├── 语言文字与数字标点规范.md              语言/数字/标点/语病/自检清单
│   ├── 字段与用法.md                          字段表、命令行、已落地的国标细节、自检怎么读
│   ├── 必答要素与可选项清单.md                一次问全的必答/可选清单与话术
│   ├── SOURCES.md                             参考文件来源与版本说明
│   ├── 政府工作报告/                          14 份国务院政府工作报告、五年规划（官方语料）
│   └── 党代会报告/                            十九大、二十大报告（官方语料）
├── fonts/
│   ├── README.md                字体清单、获取、安装、回退链
│   └── FONTS_LIST.md            字体与字号对照核对表
└── examples/                    仅供查看字段写法的样例，不作模板
    ├── 01-通知.json … 16-信函格式.json        16 份样例要素文件
    ├── generated/                             样例生成的 .docx（版式样板）
    └── preview/                               8 张首页预览图，供目视比对
```

## 安装

技能发现规则是「扫描根目录下的一层」：`<扫描根>/<name>/SKILL.md` 或 `<扫描根>/<name>.md`。
DSH 的扫描根按优先级为：项目 `<项目根>/.dsh/skills` → 项目 `<项目根>/.agents/skills` →
用户 `~/.dsh/skills`（Windows 即 `C:\Users\<用户名>\.dsh\skills`）→ 用户 `~/.agents/skills`。

### 方式 1：解压安装（通用，推荐）

```powershell
# Windows PowerShell：解压后把 official-doc-writer 整个目录放进用户级 skills 目录
Expand-Archive -Path official-doc-writer-2.1.0.zip -DestinationPath $env:TEMP\gw
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.dsh\skills" | Out-Null
Move-Item "$env:TEMP\gw\official-doc-writer" "$env:USERPROFILE\.dsh\skills\" -Force
```

```bash
# macOS / Linux
unzip official-doc-writer-2.1.0.zip -d /tmp/gw
mkdir -p ~/.dsh/skills && mv /tmp/gw/official-doc-writer ~/.dsh/skills/
```

### 方式 2：直接复制目录（不用 zip）

```powershell
Copy-Item -Recurse -Force "D:\Users\Administrator\Desktop\official-doc-writer" "$env:USERPROFILE\.dsh\skills\official-doc-writer"
```

### 方式 3：项目级安装（只对某个项目生效）

```powershell
Copy-Item -Recurse -Force "D:\Users\Administrator\Desktop\official-doc-writer" "<项目根>\.dsh\skills\official-doc-writer"
```

### 方式 4：其他环境（Trae / Claude 等）

把 `official-doc-writer` 目录放到该环境的 skills 目录；若该环境没有 skill 机制，可把 `SKILL.md` 内容
作为系统提示词附加，脚本仍可直接运行。

## 字体安装

```bash
python scripts/install_fonts.py                                    # 检测 + 指引
python scripts/install_fonts.py --copy-from C:\Windows\Fonts       # 复制本机公文字体到 fonts/
python scripts/install_fonts.py --install -y                       # 把 fonts/ 里的字体装到系统
```

| 字体 | 用途 | 来源 |
| --- | --- | --- |
| 方正小标宋_GBK | 发文机关标志、标题 | 商业字体，需授权（缺失时回退华文中宋/宋体） |
| 仿宋_GB2312、黑体、楷体_GB2312、宋体 | 正文、标题层级、页码 | Windows 自带（部分机器字族名为“仿宋/黑体/楷体/宋体”） |

安装完成后请重启 WPS/Word（必要时重启系统）再重新生成。详见 [`fonts/README.md`](fonts/README.md)与 [`fonts/FONTS_LIST.md`](fonts/FONTS_LIST.md)。

## 使用示例


```bash
python scripts/gongwen.py --needs --type 通知        # 先看还缺哪些必填要素
python scripts/gongwen.py 我的要素.json -o 通知.docx --output-format both
python scripts/gongwen.py --list-types
python scripts/gongwen.py 我的要素.json --validate 通知.docx           # 要素层 + 版式层 + 几何层
python scripts/gongwen.py --geometry 通知.docx                          # 红头位置等尺寸核对
python dev.py --syntax                                     # 改动脚本后自查语法
```

## 技术实现

### 依赖库

| 依赖 | 版本 | 用途 |
| --- | --- | --- |
| Python | 3.9+ | 运行环境（在 3.12 上验证） |
| python-docx | 1.2.0 | 生成与读取 .docx |
| fontTools | 可选 | 读取字体真实纵向度量，用于红头精确落位；未安装时用内置度量表 |
| 标准库 | — | argparse、json、zipfile、winreg、subprocess、unicodedata、ast 等 |

> 早期说明中列出的 `python-dateutil` 已不再需要：日期只做格式校验与字符串输出，用标准库 `datetime`/`re` 即可。

### 核心模块

| 模块 | 职责 |
| --- | --- |
| `scripts/gongwen.py` | **一个文件四件事**：①排版引擎（页面/版心、版头、主体、附件、版记、页码，字体探测与回退）②要素层+版式层校验（`--validate`）③几何核对（`--geometry`，解析 OOXML 量毫米）④提问清单（`--ask` / `--needs`）；红头精确落位在 `mark_placement()`，几何自检在 `GongwenBuilder.geometry()` |
| `scripts/prompts.py` | 问答式要素收集（`DialogManager`，上行文才问签发人、校验与兜底、要素预览）+ 15 文种的开头语/结语/结构骨架（`SmartPromptSystem`） |
| `scripts/install_fonts.py` | 字体检测（注册表/字体目录/fc-list）、复制、安装 |
| `dev.py --syntax` | 语法自检：`ast` 解析技能内全部 .py |
| `dev.py --package` | 打包与自检 |
| `dev.py --publish` | 把 zip 与展示指南复制到桌面 |



