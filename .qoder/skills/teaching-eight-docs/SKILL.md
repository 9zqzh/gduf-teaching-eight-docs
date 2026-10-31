---
name: teaching-eight-docs
description: 按广东金融学院模板生成「教学八大件」中的课程简介、课程教学大纲、课程教案、教学进度表，并按需生成课程考核方案、教材与参考书目；教师简介与教学课件由教师提供并原样收录、绝不修改。当用户要做教学文件、教学八大件、课程简介/教学大纲/教案/进度表/考核方案/教材参考书目，或要把课件与教师简介整理成一套可交付 Word 时使用。
---

# 教学八大件生成

把教师提供的课件、教师简介（可选：人才培养方案、考核方案、教材书目）整理成一套对标学校模板的 Word 教学文件，并在交付前做结构与一致性校验。

## 硬性规则

1. **先索要 2 号教师简介和 6 号教学课件**。拿到原件前不要开始生成。
2. **2 号、6 号只读**：复制进交付目录，逐字节不改；不打开改写、不转格式、不重命名内容。
3. 用户没有 2/6 时，先问是否代做：2 号按教师口述信息生成（缺项写「（待补充）」并列出）；6 号只能生成提纲版 PPTX 并标注「提纲版，需教师替换」。
4. **7/8 只在用户提供考核方案/教材书目时生成**；未提供就不生成，并在交付提醒中说明原因。
5. **不编造**课程代码、学分学时、比例、日期、签字人、教材版本；无来源写「（待补充）」。
6. **不把模板原课程内容带入新课程**（数学分析III、常微分方程等仅作版式参考）。
7. 生成前必须让用户确认课程信息、生成范围、2/6 处理方式、输出目录（见 `references/intake.md` 的确认闸门）。

## 目录

```
SKILL.md
references/templates-spec.md   8 份文档的版式、表格、字段规范（生成依据）
references/intake.md           材料收集清单、课程信息确认表、缺失处理
references/qa-checklist.md     成品自检清单与报告格式
scripts/extract_materials.py   从 docx/doc/pptx/ppt/pdf 提取文本与表格
scripts/build_docs.py          按 course.json 生成 1/3/4/5/7/8 并复制 2/6
scripts/validate_docs.py       结构、占位符、学时、比例、跨文档一致性校验
scripts/docx_kit.py            版式工具库（字体、页边距、表格合并）
assets/data/course.example.json  完整数据示例（可直接复制改写）
assets/images/                 封面院徽与标题图（3 号文档封面用）
assets/samples/                2/7/8 号模板样例，供用户对照
```

以下命令中的 `scripts/...` 均相对于本 SKILL.md 所在目录；Windows 下用 `python`，并在命令前加 `PYTHONIOENCODING=utf-8` 避免中文乱码。

## 工作流程

### 第 1 步：收材料（不可跳过）

向用户明确要两样东西：

- **2 号教师简介**（.docx/.doc，含姓名、性别、出生年月、职称、本科/博士院校、参加工作时间、承担课程）
- **6 号教学课件**（PPT/PDF，按章节组织；整套文件夹最好）

同时收集课程基本信息、章节与学时。可选材料：人才培养方案、旧版大纲/进度表、课程考核方案、教材与书目、思政/实践要求。逐项确认表见 `references/intake.md`。

用户说没有 2/6 时，按规则 3 处理并取得明确同意；用户坚持先出稿时，在交付说明顶部标注「草案」。

### 第 2 步：解析材料

```bash
PYTHONIOENCODING=utf-8 python scripts/extract_materials.py \
  --out <工作目录>/materials.json \
  --dir "<材料目录>"
```

也可直接传文件路径。脚本支持 .docx/.doc/.pptx/.ppt/.pdf；.doc/.ppt 通过 Office COM 转换（需本机安装 Word/PowerPoint）。先读 `materials.json` 的 `digest` 字段；需要细节时再按 `files[i].tables` / `slides` / `pages` 查。

从材料中提取：课程名/英文名/代码/性质/专业/学分/学时、章节标题与学时、每章内容/重难点、教师信息、教材书目、考核比例。材料互相冲突时以用户最新说明为准，并记录覆盖关系。

### 第 3 步：组装 course.json 并请用户确认

复制 `assets/data/course.example.json` 为工作文件，按材料填写。最小必需字段：

| 区块 | 关键字段 |
|---|---|
| `meta` | 学期、教师姓名、开课单位、制定人、审核人、审批人、日期 |
| `teacher` | 姓名、性别、出生年月、职称、本科/博士院校、参加工作时间、承担/目前课程（代做 2 号时用） |
| `course` | 名称、英文名、代码、性质、类型、专业、班级、学分、总/理论/实验/实践学时、先修、语言 |
| `course.intro` | background / purpose / content / features / methods / outcomes |
| `course.objectives` | 课程目标列表（可测量） |
| `course.objective_map` | 毕业要求、指标点、描述、对应目标（`1（H）、2（H）`） |
| `course.chapters` | no、title、hours、objectives、content、goals、key_points、difficulties、analysis、preparation、methods、homework、organization、ideology、reflection |
| `course.schedule` | week、chapter、content、hours、method（学时合计必须等于总学时） |
| `course.assessment` | exam_form、regular_ratio、exam_ratio、items[name/requirement/ratio/objectives/group] |
| `course.textbook` | main、references、online |
| `options.docs` | 默认 `[1,3,4,5]`；有考核方案加 7，有书目加 8 |

写作约定：

- 公式用 Unicode 数学符号；上下标用 `^{...}` / `_{...}`，生成时转为 Word 上下标。不要插入公式图片。
- 章节学时用整数，全部章节之和 = `total_hours`；考核子项之和 = 100。
- 没有依据的字段写 `（待补充）`，不要猜。

**确认闸门**：把填好的课程信息表、拟生成清单（1/3/4/5，以及是否 7/8）、2/6 是复制还是代做、输出目录、已知缺口列给用户，得到确认后再生成。

### 第 4 步：生成

交付目录命名 `<学期>-<教师姓名><课程名称>`，例如 `2026-2027-1张三数学分析III`。

用户提供 2/6 时（原样复制）：

```bash
PYTHONIOENCODING=utf-8 python scripts/build_docs.py \
  --data <工作目录>/course.json \
  --out "<交付目录>" \
  --docs 1,3,4,5[,7,8] \
  --teacher-profile "<2 号原件路径>" \
  --courseware "<6 号课件文件夹或文件路径>"
```

用户没有 2 号、同意代做时：把 `2` 加进 `--docs`，不要传 `--teacher-profile`。

- 只生成 1/3/4/5 是默认；7/8 是否加入由材料决定。
- 若 2/6 已在交付目录中，用 `--no-copy` 避免重复复制。
- 脚本会写 `生成记录.json`（生成了什么、复制了什么、哪些字段待补充、原件路径）。
- 源文件在交付目录内时不会自我复制；2/6 的复制使用 `copy2`，保留原始字节与时间戳。

### 第 5 步：校验

```bash
PYTHONIOENCODING=utf-8 python scripts/validate_docs.py \
  --dir "<交付目录>" \
  --data <工作目录>/course.json \
  --report "<交付目录>/校验报告.md"
```

退出码 1 表示有致命问题（缺 2/6、学时或比例不平、缺文档），必须处理；退出码 0 表示只剩需要教师复核的提醒。校验覆盖：

- 文档存在、A4 页面、页边距 2.8/2.6/2.8/2.6
- 各文档必备栏目与表格数量
- 「（待补充）」与模板占位符
- 大纲章节学时、教案次课学时、进度表学时、考核比重
- 课程名/代码/教师/学时/学分在文档间一致

### 第 6 步：交付说明

在聊天里按下面结构提醒用户（由用户决定是否优化）：

```
已生成：1、3、4、5（+7/8）
直接收录未修改：2、6（复制自 <路径>）
未生成：7/8（原因：未提供…）
必须人工补充：<（待补充）字段清单>
警告：<校验报告中的提醒>
建议复核：AI 推导的学情分析、思政设计、作业页码、毕业要求映射
```

同时说明与模板的差异：封面院徽为模板提取图（若学院更换标识需替换 `assets/images/`）；大纲对应关系表按模板原宽度略超版心；模板未提供的内容不要假装齐全。

## 按需处理

- **只有课件、没有大纲/进度表**：按总学时把章节拆成 2 学时/次课，`schedule` 与 `chapters` 学时保持一致；在报告中标注「章节学时为按总学时拆分，请教师确认」。
- **没有人才培养方案**：`objective_map` 用同专业通用表述，标注待确认；不要编造指标点编号。
- **没有考核方案但有考核比例**：只写 3 号文档的考核表与评价标准，不生成 7 号。
- **没有书目**：教材字段写「（待补充）」，不生成 8 号。
- **用户只想要其中几件**：用 `--docs` 控制，仍要先收齐其依赖材料。

## 常见问题

- **中文输出乱码**：命令前加 `PYTHONIOENCODING=utf-8`。
- **.doc/.ppt 无法解析**：需要本机 Word/PowerPoint；或先手动另存为 .docx/.pptx。
- **表格列宽被平均**：`docx_kit.add_table` 已同时设置 `tblGrid`、`tblW`、每个单元格宽度；不要绕开它手写表格。
- **合并单元格报错**：`add_table` 中合并位置用 `None` 占位，`span` 横向合并，`vmerge` 纵向合并。
- **字号/字体不对**：统一用 `docx_kit.set_run_font`；中文字体走 `w:eastAsia`。

版式细节以 `references/templates-spec.md` 为准；交付前对照 `references/qa-checklist.md` 人工抽查。
