"""Check the generated teaching documents against the template spec and data.

Usage:
    python validate_docs.py --dir "D:/out/2026-2027-1张三高等数学" \
        --data course.json --docs 1,3,4,5,7,8 --report 校验报告.md

Exit code is 1 when a fatal problem is found, 0 otherwise. Warnings never fail
the run: they are the items the teacher must review.

Checks:
  * requested files exist (2/6 are only reported, never generated)
  * A4 page; body margins follow the school template set (2.8/2.6/2.8/2.6 or
    3.175/3.175/2.54/2.54, depending on the document)
  * required headings / tables per document
  * placeholder scan in generated documents
  * chapter hours / total hours / assessment ratios add up
  * course name, code, credits, hours agree across documents
"""

import argparse
import glob
import os
import re
import sys

from docx import Document

FATAL = []
WARN = []
INFO = []

REQUIRED_HEADINGS = {
    "3": [
        "一、课程信息",
        "二、课程简介",
        "三、课程目标",
        "四、课程目标与毕业要求的对应关系",
        "五、教学内容、要求及进度安排",
        "六、考核与评价",
        "七、教材、参考文献与其他教学资源",
        "八、说明",
    ],
    "4": ["一、课程简介", "教学目的与教学要求"],
    "5": ["教学进度表"],
    "7": ["广东金融学院课程考核方案"],
    "8": ["教材与参考书目"],
}

EXPECTED_TABLES = {"1": 1, "2": 1, "3": 10, "5": 1, "7": 1, "8": 0}

PLACEHOLDER_PATTERNS = ["（待补充）", "TODO", "TBD", "XXX", "###"]

DOC_NAMES = {
    "1": "1、课程简介.docx",
    "3": "3、课程教学大纲.docx",
    "4": "4、课程教案.docx",
    "5": "5、教学进度表.docx",
    "7": "7、课程考核方案.docx",
    "8": "8、教材与参考书目.docx",
}


def fatal(msg):
    FATAL.append(msg)


def warn(msg):
    WARN.append(msg)


def info(msg):
    INFO.append(msg)


# ---------------------------------------------------------------- helpers
def doc_text(doc):
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                parts.append(cell.text)
    return "\n".join(parts)


def row_values(row):
    values, seen = [], set()
    for cell in row.cells:
        if id(cell._tc) in seen:
            continue
        seen.add(id(cell._tc))
        values.append(cell.text.strip())
    return values


def find_table(doc, *keywords, cols=None):
    for table in doc.tables:
        if cols is not None and len(table.columns) != cols:
            continue
        header = " ".join(row_values(table.rows[0])) if table.rows else ""
        if all(k in header for k in keywords):
            return table
    return None


def numbers(text):
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", text.replace("％", "%"))]


def check_page_setup(doc, name):
    section = doc.sections[0]
    if abs(section.page_width.cm - 21.0) > 0.15 or abs(section.page_height.cm - 29.7) > 0.15:
        warn("%s 纸张尺寸偏离 A4：%.2f×%.2f" % (name, section.page_width.cm, section.page_height.cm))
    actual = (
        section.left_margin.cm,
        section.right_margin.cm,
        section.top_margin.cm,
        section.bottom_margin.cm,
    )
    allowed = [
        (2.8, 2.6, 2.8, 2.6),
        (3.175, 3.175, 2.54, 2.54),
    ]
    if not any(all(abs(a - e) <= 0.15 for a, e in zip(actual, expected)) for expected in allowed):
        warn(
            "%s 页边距偏离模板：%.2f/%.2f/%.2f/%.2f"
            % (name, actual[0], actual[1], actual[2], actual[3])
        )


def check_placeholders(text, name):
    hits = []
    for pattern in PLACEHOLDER_PATTERNS:
        count = text.count(pattern)
        if count:
            hits.append("%s × %d" % (pattern, count))
    if hits:
        warn("%s 存在待补充内容：%s（需教师确认后补齐）" % (name, "，".join(hits)))


def find_teacher_profile(directory):
    hits = glob.glob(os.path.join(directory, "2、教师简历*"))
    if not hits:
        hits = [p for p in glob.glob(os.path.join(directory, "2、*")) if os.path.isfile(p)]
    return hits


def find_courseware(directory):
    hits = glob.glob(os.path.join(directory, "6、*")) + glob.glob(os.path.join(directory, "6.*"))
    return [h for h in hits]


# ---------------------------------------------------------------- checks
def check_documents(directory, docs, data):
    texts = {}
    for key in docs:
        expected = DOC_NAMES.get(key)
        if not expected:
            continue
        path = os.path.join(directory, expected)
        if not os.path.exists(path):
            fatal("缺少文档：%s" % expected)
            continue
        try:
            doc = Document(path)
        except Exception as exc:  # noqa: BLE001
            fatal("%s 无法打开：%s" % (expected, exc))
            continue

        text = doc_text(doc)
        texts[key] = text
        compact = re.sub(r"\s+", "", text)
        check_page_setup(doc, expected)
        check_placeholders(text, expected)

        for heading in REQUIRED_HEADINGS.get(key, []):
            if re.sub(r"\s+", "", heading) not in compact:
                warn("%s 缺少栏目：%s" % (expected, heading))

        if key in EXPECTED_TABLES:
            expected_count = EXPECTED_TABLES[key]
            actual = len(doc.tables)
            if expected_count and actual != expected_count:
                warn("%s 表格数量 %d，模板为 %d" % (expected, actual, expected_count))

    profiles = find_teacher_profile(directory)
    if not profiles:
        fatal("交付目录缺少 2 号教师简介（应由教师提供，原样收录）")
    courseware = find_courseware(directory)
    if not courseware:
        fatal("交付目录缺少 6 号教学课件（应由教师提供，原样收录）")
    return texts


def check_hours(doc3, doc4, doc5, data):
    total = None
    if data:
        course = data.get("course", {})
        total = course.get("total_hours")
        credits = course.get("credits")
        theory = course.get("theory_hours", total)
        lab = course.get("lab_hours", 0)
        practice = course.get("practice_hours", 0)
        if total and theory is not None:
            if (theory or 0) + (lab or 0) + (practice or 0) != total:
                fatal("学时构成不平：理论 %s + 实验 %s + 实践 %s ≠ 总学时 %s" % (theory, lab, practice, total))
        chapters = course.get("chapters") or []
        if total and chapters:
            chapter_sum = sum(int(ch.get("hours", 0) or 0) for ch in chapters)
            if chapter_sum != total:
                fatal("章节学时合计 %d ≠ 总学时 %s" % (chapter_sum, total))
        info("数据：总学时 %s，学分 %s" % (total, credits))

    if doc3 is not None:
        tables = doc3.tables
        chapter_hours = []
        for table in tables:
            for row in table.rows:
                values = row_values(row)
                if values and re.match(r"第\s*[一二三四五六七八九十百零\d]+\s*章", values[0]):
                    joined = " ".join(values)
                    match = re.search(r"学时[：:]\s*(\d+)", joined)
                    if match:
                        chapter_hours.append(int(match.group(1)))
        if chapter_hours:
            if total and sum(chapter_hours) != total:
                fatal("大纲章节学时合计 %d ≠ 总学时 %s" % (sum(chapter_hours), total))
            info("大纲章节学时：%s（合计 %d）" % ("+".join(map(str, chapter_hours)), sum(chapter_hours)))

        assessment = find_table(doc3, "比重")
        if assessment is not None:
            ratios = []
            for row in assessment.rows[1:]:
                values = row_values(row)
                if len(values) >= 3 and re.fullmatch(r"\d+(?:\.\d+)?", values[2]):
                    ratios.append(float(values[2]))
            if ratios and abs(sum(ratios) - 100) > 0.01:
                fatal("考核比重合计 %.1f ≠ 100" % sum(ratios))
            elif ratios:
                info("考核比重：%s（合计 %.0f）" % ("+".join("%g" % r for r in ratios), sum(ratios)))

    if doc4 is not None:
        pattern = re.compile(r"第\s*[\d一二三四五六七八九十\-—]+\s*次课\s*(\d+)\s*学时")
        lesson_hours = [int(x) for x in pattern.findall(doc_text(doc4))]
        if lesson_hours and total:
            if sum(lesson_hours) != total:
                fatal("教案各次课学时合计 %d ≠ 总学时 %s" % (sum(lesson_hours), total))
            info("教案次课学时合计：%d（%d 次课）" % (sum(lesson_hours), len(lesson_hours)))

    if doc5 is not None:
        table = None
        for candidate in doc5.tables:
            header = " ".join(row_values(candidate.rows[0])) if candidate.rows else ""
            if "周次" in header and ("授课" in header or "学时" in header):
                table = candidate
                break
        if table is None:
            warn("进度表未找到「周次 / 授课学时」表格")
            return
        headers = row_values(table.rows[0])
        hour_col = None
        for i, header in enumerate(headers):
            if "授课" in header or "学时" in header:
                hour_col = i
                break
        if hour_col is None:
            return
        hours = []
        for row in table.rows[1:]:
            values = row_values(row)
            if hour_col < len(values) and re.fullmatch(r"\d+", values[hour_col]):
                hours.append(int(values[hour_col]))
        if hours:
            if total and sum(hours) != total:
                fatal("进度表学时合计 %d ≠ 总学时 %s" % (sum(hours), total))
            info("进度表学时合计：%d" % sum(hours))


def check_consistency(texts, data):
    if not data:
        return
    course = data.get("course", {})
    meta = data.get("meta", {})
    name = str(course.get("name") or "").strip()
    code = str(course.get("code") or "").strip()
    teacher = str(meta.get("author") or meta.get("teacher_name") or "").strip()
    total = course.get("total_hours")
    credits = course.get("credits")

    expect_docs = {
        "课程名称": (name, ["1", "3", "5", "7"]),
        "课程代码": (code, ["1", "3"]),
        "教师姓名": (teacher, ["3", "7"]),
        "总学时": (str(total) if total else "", ["3", "5", "8"]),
        "学分": (str(credits) if credits else "", ["3", "5", "8"]),
    }
    for label, (value, keys) in expect_docs.items():
        if not value:
            continue
        for key in keys:
            text = texts.get(key)
            if text is None:
                continue
            if value not in text:
                warn("%s 中未找到%s「%s」" % (DOC_NAMES.get(key, key), label, value))


# ---------------------------------------------------------------- report
def build_report(directory, docs):
    lines = ["# 教学八大件校验报告", "", "交付目录：`%s`" % directory, "", "## 结论", ""]
    if FATAL:
        lines.append("**不通过**：发现 %d 个致命问题，必须处理后才能交付。" % len(FATAL))
    elif WARN:
        lines.append("**通过（需教师复核）**：无致命问题，有 %d 条提醒。" % len(WARN))
    else:
        lines.append("**通过**：未发现结构性问题。")

    sections = [
        ("致命问题（必须处理）", FATAL),
        ("提醒（建议教师复核）", WARN),
        ("已核对", INFO),
    ]
    for title, items in sections:
        lines.extend(["", "## " + title, ""])
        if not items:
            lines.append("- 无")
        for item in items:
            lines.append("- " + item)

    lines.extend(
        [
            "",
            "## 交付前人工确认",
            "",
            "- 2 号教师简介、6 号教学课件是否与教师提供的原件完全一致（本工具不修改这两项）。",
            "- 「（待补充）」处是否已按教师意见补齐；未补齐的不得当作最终稿。",
            "- 名单、班级、日期、签字人是否与教务安排一致。",
            "- 抽查大纲与教案的章节学时、进度表周次是否互相对应。",
        ]
    )
    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser(description="Validate generated teaching documents")
    parser.add_argument("--dir", required=True, help="delivery folder")
    parser.add_argument("--data", help="course.json used to generate the documents")
    parser.add_argument("--docs", default="1,3,4,5,7,8", help="documents expected in the folder")
    parser.add_argument("--report", help="write the markdown report here")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    data = None
    if args.data:
        import json

        with open(args.data, encoding="utf-8") as fh:
            data = json.load(fh)

    docs = [d.strip() for d in args.docs.split(",") if d.strip()]
    directory = os.path.abspath(args.dir)
    if not os.path.isdir(directory):
        raise SystemExit("delivery folder not found: %s" % directory)

    texts = check_documents(directory, docs, data)

    doc3 = Document(os.path.join(directory, DOC_NAMES["3"])) if os.path.exists(os.path.join(directory, DOC_NAMES["3"])) else None
    doc5 = Document(os.path.join(directory, DOC_NAMES["5"])) if os.path.exists(os.path.join(directory, DOC_NAMES["5"])) else None
    doc4 = Document(os.path.join(directory, DOC_NAMES["4"])) if os.path.exists(os.path.join(directory, DOC_NAMES["4"])) else None
    check_hours(doc3, doc4, doc5, data)
    check_consistency(texts, data)

    report = build_report(directory, docs)
    print(report)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            fh.write(report)
        print("report written: %s" % args.report)
    return 1 if FATAL else 0


if __name__ == "__main__":
    sys.exit(main())
