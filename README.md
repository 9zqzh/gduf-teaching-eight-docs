# 教学八大件生成 Skill（广东金融学院）

一个 Qoder Agent Skill：把课件、教师简历（可选：人才培养方案、考核要求、教材书目）整理成一套对标学校模板的 Word 教学文件，并在交付前做结构与一致性校验。

## 能生成什么

| 编号 | 文档 | 生成方式 |
|---|---|---|
| 1 | 课程简介 | 按模板生成 |
| 2 | 教师简历 | 教师提供，原样复制、绝不修改 |
| 3 | 课程教学大纲 | 按模板生成 |
| 4 | 课程教案 | 按模板生成 |
| 5 | 教学进度表 | 按模板生成 |
| 6 | 教学课件 | 教师提供，原样复制、绝不修改 |
| 7 | 课程考核方案 | 提供考核要求时才生成 |
| 8 | 教材与参考书目 | 提供教材书目时才生成 |

脚本会写出 `生成记录.json`，记录生成了哪些文件、复制了哪些原件、哪些字段仍待补充。

## 目录结构

```
.qoder/skills/teaching-eight-docs/
├── SKILL.md                技能说明与工作流
├── references/             模板版式规范、材料清单、自检清单
├── scripts/                材料提取、文档生成、成品校验
└── assets/                 示例数据、封面图片、脱敏样表
templates/                  四门课的脱敏 Word 模板（1-5、7、8）
requirements.txt
```

`templates/` 收录四套模板：数学分析III、高等代数上、人工智能、常微分方程，覆盖 `.doc` 与 `.docx` 混合的情况。每套含 1-5、7、8 号文档；6 号课件体积大且含第三方署名，不在仓库中。

## 使用方式

在 Qoder 中调用技能：

```
/teaching-eight-docs
```

按提示先提供 **2 号教师简历** 和 **6 号教学课件**，以及课程名称、代码、学分学时、章节安排。没有 2/6 时可选择代做：2 号按口述信息生成，6 号只生成提纲版 PPTX 并标注「提纲版，需教师替换」。

也可以直接用脚本：

```bash
# 1. 提取材料
python .qoder/skills/teaching-eight-docs/scripts/extract_materials.py \
  --out work/materials.json --dir "<材料目录>"

# 2. 按 course.example.json 组装 course.json 后生成
python .qoder/skills/teaching-eight-docs/scripts/build_docs.py \
  --data work/course.json --out "<交付目录>" \
  --docs 1,3,4,5 --teacher-profile "<2号原件>" --courseware "<6号课件>"

# 3. 校验
python .qoder/skills/teaching-eight-docs/scripts/validate_docs.py \
  --dir "<交付目录>" --data work/course.json --report "<交付目录>/校验报告.md"
```

Windows 下建议在命令前加 `PYTHONIOENCODING=utf-8`，避免中文输出乱码。

## 环境依赖

- Windows + Microsoft Word / PowerPoint（仅用于 `.doc`、`.ppt` 的格式转换）
- Python 3.9+
- 依赖见 `requirements.txt`

## 模板脱敏说明

`templates/` 与 `assets/samples/` 中的文档已做脱敏处理，替换规则：

| 原内容 | 占位符 |
|---|---|
| 教师姓名 | `教师姓名` |
| 审批人姓名 | `审批人姓名` |
| 出生年月 | `（出生年月）` |
| 本科毕业院校 | `（本科毕业院校）` |
| 博士毕业院校 | `（博士毕业院校）` |
| 参加工作时间 | `（参加工作时间）` |

课程名称、章节结构、学时安排、考核比例等教学信息保留，可作为填写参考；7 号文档中的手写签名图片已移除；文档属性中的作者、最后修改者也已改为「模板」。

原始模板目录、教师课件与生成成品含真实个人信息，已通过 `.gitignore` 排除，不进入仓库。

## 使用边界

- 2 号教师简历、6 号教学课件只读：复制进交付目录，不打开改写、不转格式。
- 不编造课程代码、学分学时、考核比例、日期、签字人；无来源的字段写「（待补充）」。
- 7/8 号只在提供考核要求 / 教材书目时生成。
- 模板以学校最新下发的版本为准；本项目为个人整理，非学校官方发布。
