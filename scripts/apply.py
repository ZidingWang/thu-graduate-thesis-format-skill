"""apply.py — 按 2025 版规范整改格式。只改格式,绝不改文字。

用法:
    检查模式(不写文件,仅输出问题清单):
        python apply.py 论文.docx -m structure_map.json --dry-run -r report.json
    整改模式:
        python apply.py 论文.docx -m structure_map.json -o 论文_整改.docx -r report.json

report.json 记录每一处差异/修改:
    {"paragraph": 序号, "text": 预览, "role": 角色,
     "changes": [{"prop": 属性, "old": 旧值, "new": 新值}]}

设计约束:
- 永不修改 run.text / 段落文本 / 表格单元格文本。
- 字体修改按 中文(eastAsia)/西文(ascii,hAnsi) 分别设置。
- 不碰 run 的加粗/斜体/上下标(标题除外:章节标题统一不强制加粗,黑体本身即粗体效果)。
- 含 OMML 公式的 run 不改字体(只调段落属性),避免破坏公式渲染。
"""
import argparse
import copy
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

sys.path.insert(0, str(Path(__file__).parent))
import spec

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
ALIGN_NAME = {v: k for k, v in ALIGN.items()}

NO_TOUCH_ROLES = {"cover", "skip"}


def run_has_math(run) -> bool:
    return bool(run._r.findall(
        ".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))


def para_has_math(p) -> bool:
    return bool(p._p.findall(
        ".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))


def get_run_fonts(run):
    """返回 (ascii, eastAsia, size_pt)。读取 run 直接格式,不解析样式继承。"""
    rPr = run._r.rPr
    ascii_f = ea_f = None
    if rPr is not None and rPr.rFonts is not None:
        ascii_f = rPr.rFonts.get(qn("w:ascii"))
        ea_f = rPr.rFonts.get(qn("w:eastAsia"))
    size = run.font.size.pt if run.font.size else None
    return ascii_f, ea_f, size


def set_run_fonts(run, ascii_font, ea_font, size_pt, changes, idx):
    old_ascii, old_ea, old_size = get_run_fonts(run)
    if old_ascii != ascii_font or old_ea != ea_font:
        run.font.name = ascii_font  # 设置 ascii + hAnsi
        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), ea_font)
        changes.append({"prop": f"run{idx}.字体",
                        "old": f"西文{old_ascii or '(继承)'}/中文{old_ea or '(继承)'}",
                        "new": f"西文{ascii_font}/中文{ea_font}"})
    if old_size != size_pt:
        run.font.size = Pt(size_pt)
        changes.append({"prop": f"run{idx}.字号",
                        "old": f"{old_size or '(继承)'}pt", "new": f"{size_pt}pt"})


def describe_line(pf):
    rule, ls = pf.line_spacing_rule, pf.line_spacing
    if rule == WD_LINE_SPACING.EXACTLY and ls is not None:
        return ("exact", round(ls.pt, 1))
    if rule == WD_LINE_SPACING.SINGLE:
        return ("single", None)
    if rule == WD_LINE_SPACING.MULTIPLE or isinstance(ls, float):
        return ("multiple", ls)
    if ls is None:
        return (None, None)
    return ("other", str(ls))


def apply_para_format(p, s, changes):
    pf = p.paragraph_format
    # 对齐
    want = ALIGN[s["align"]]
    if pf.alignment != want:
        changes.append({"prop": "对齐",
                        "old": ALIGN_NAME.get(pf.alignment, str(pf.alignment or "(继承)")),
                        "new": s["align"]})
        pf.alignment = want
    # 行距(含公式的段落按规范允许按需调整,正文角色时跳过)
    rule, val = s["line"]
    if not (para_has_math(p) and s is spec.BODY):
        cur = describe_line(pf)
        tgt = (rule, float(val) if val else None)
        if cur != tgt:
            changes.append({"prop": "行距", "old": str(cur), "new": str(tgt)})
            if rule == "exact":
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = Pt(val)
            elif rule == "single":
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            elif rule == "multiple":
                pf.line_spacing = float(val)
    # 段前段后
    for attr, key, label in (("space_before", "before", "段前"),
                             ("space_after", "after", "段后")):
        cur = getattr(pf, attr)
        cur_pt = round(cur.pt, 1) if cur is not None else None
        if cur_pt != float(s[key]):
            changes.append({"prop": label, "old": f"{cur_pt}磅" if cur_pt is not None
                            else "(继承)", "new": f"{s[key]}磅"})
            setattr(pf, attr, Pt(s[key]))
    # 首行缩进(2汉字符) / 悬挂缩进 —— 用 w:ind 的 firstLineChars/hangingChars,随字号自适应
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if "first_indent_chars" in s:
        n = s["first_indent_chars"] * 100
        cur = ind.get(qn("w:firstLineChars")) if ind is not None else None
        if cur != str(n):
            changes.append({"prop": "首行缩进", "old": f"{cur or '(无)'}(百分字符)",
                            "new": f"{s['first_indent_chars']}字符"})
            if ind is None:
                ind = pPr.makeelement(qn("w:ind"), {})
                pPr.append(ind)
            for a in ("w:firstLine", "w:hanging", "w:hangingChars"):
                if ind.get(qn(a)):
                    del ind.attrib[qn(a)]
            ind.set(qn("w:firstLineChars"), str(n))
    elif "hanging_chars" in s:
        n = int(s["hanging_chars"] * 100)
        cur = ind.get(qn("w:hangingChars")) if ind is not None else None
        if cur != str(n):
            changes.append({"prop": "悬挂缩进", "old": f"{cur or '(无)'}(百分字符)",
                            "new": f"{s['hanging_chars']}字符"})
            if ind is None:
                ind = pPr.makeelement(qn("w:ind"), {})
                pPr.append(ind)
            for a in ("w:firstLine", "w:firstLineChars"):
                if ind.get(qn(a)):
                    del ind.attrib[qn(a)]
            ind.set(qn("w:hangingChars"), str(n))
            ind.set(qn("w:hanging"), str(int(s["hanging_chars"] * 240)))
    else:
        # 标题/题注等:清除遗留首行缩进
        if ind is not None and (ind.get(qn("w:firstLineChars")) or ind.get(qn("w:firstLine"))):
            changes.append({"prop": "首行缩进", "old": "有", "new": "无"})
            for a in ("w:firstLine", "w:firstLineChars"):
                if ind.get(qn(a)):
                    del ind.attrib[qn(a)]


def is_cjk(ch):
    return "\u4e00" <= ch <= "\u9fff" or ch in "。,、;:?!“”‘’()《》—…·"


def apply_run_formats(p, s, changes):
    for i, run in enumerate(p.runs):
        if run_has_math(run):
            continue
        t = run.text
        if not t.strip():
            continue
        cjk = sum(1 for c in t if is_cjk(c))
        # 纯西文 run 中文字体也按规范设置,不影响显示;混排 run 按中西文各自字体设置
        set_run_fonts(run, s["en_font"], s["cn_font"], s["size"], changes, i)


def apply_section_setup(doc, report):
    for si, sec in enumerate(doc.sections):
        changes = []
        for attr, key, label in (("top_margin", "margin_top_cm", "上边距"),
                                 ("bottom_margin", "margin_bottom_cm", "下边距"),
                                 ("left_margin", "margin_left_cm", "左边距"),
                                 ("right_margin", "margin_right_cm", "右边距"),
                                 ("gutter", "gutter_cm", "装订线"),
                                 ("header_distance", "header_cm", "页眉距边界"),
                                 ("footer_distance", "footer_cm", "页脚距边界")):
            cur = getattr(sec, attr)
            cur_cm = round(cur.cm, 2) if cur is not None else None
            want = spec.PAGE[key]
            if cur_cm != want:
                changes.append({"prop": label, "old": f"{cur_cm}cm", "new": f"{want}cm"})
                setattr(sec, attr, Cm(want))
        if changes:
            report.append({"paragraph": f"节{si}", "text": "(页面设置)",
                           "role": "section", "changes": changes})


def _border_el(parent_tag_owner, tag, sz_eighth):
    el = parent_tag_owner.makeelement(qn(f"w:{tag}"), {})
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz_eighth))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "000000")
    return el


def apply_three_line_borders(table):
    """三线表:表上下边线1.5磅(sz=12),表头行下线1磅(sz=8),清除其余框线。
    返回修改记录列表。边框 sz 单位为 1/8 磅。"""
    changes = []
    tbl = table._tbl
    tblPr = tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = tblPr.makeelement(qn("w:tblBorders"), {})
    borders.append(_border_el(tblPr, "top", 12))
    borders.append(_border_el(tblPr, "bottom", 12))
    for tag in ("left", "right", "insideH", "insideV"):
        none_el = tblPr.makeelement(qn(f"w:{tag}"), {})
        none_el.set(qn("w:val"), "none")
        none_el.set(qn("w:sz"), "0")
        none_el.set(qn("w:space"), "0")
        borders.append(none_el)
    tblPr.append(borders)
    changes.append({"prop": "表边框", "old": "(原线型)",
                    "new": "三线表:上下1.5磅,内线清除"})
    # 清除单元格级边框,再给表头行下边加 1 磅线
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tb = tcPr.find(qn("w:tcBorders"))
            if tb is not None:
                tcPr.remove(tb)
            if ri == 0:
                tb = tcPr.makeelement(qn("w:tcBorders"), {})
                tb.append(_border_el(tcPr, "bottom", 8))
                tcPr.append(tb)
    changes.append({"prop": "表头下线", "old": "(原线型)", "new": "1磅单线"})
    return changes


def extract_all_text(doc):
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-m", "--map", required=True, help="structure_map.json(已人工复核)")
    ap.add_argument("-o", "--out", help="输出 docx(整改模式必填)")
    ap.add_argument("-r", "--report", default="format_report.json")
    ap.add_argument("--dry-run", action="store_true", help="只检查不写文件")
    ap.add_argument("--skip-page-setup", action="store_true")
    args = ap.parse_args()

    doc = Document(args.docx)
    raw = json.loads(Path(args.map).read_text(encoding="utf-8"))
    smap = raw["paragraphs"] if isinstance(raw, dict) else raw
    table_roles = {t["index"]: t.get("role", "three_line")
                   for t in (raw.get("tables", []) if isinstance(raw, dict) else [])}
    if len(smap) != len(doc.paragraphs):
        sys.exit(f"结构图谱段数({len(smap)})与文档段数({len(doc.paragraphs)})不一致,"
                 "请重新运行 classify.py")

    text_before = extract_all_text(doc)
    report = []

    if not args.skip_page_setup:
        apply_section_setup(doc, report)

    for entry, p in zip(smap, doc.paragraphs):
        role = entry["role"]
        if role in NO_TOUCH_ROLES or role not in spec.ROLE_SPEC:
            continue
        s = spec.ROLE_SPEC[role]
        changes = []
        apply_para_format(p, s, changes)
        apply_run_formats(p, s, changes)
        if changes:
            report.append({"paragraph": entry["index"],
                           "text": entry["text"], "role": role,
                           "changes": changes})

    # 表格:单元格文字格式 + 三线表边框(role=three_line 时)
    s = spec.TABLE_CELL
    for ti, table in enumerate(doc.tables):
        t_role = table_roles.get(ti, "three_line")
        if t_role == "skip":
            continue
        if t_role == "three_line":
            bc = apply_three_line_borders(table)
            if bc:
                report.append({"paragraph": f"表{ti+1}", "text": "(边框)",
                               "role": "table_borders", "changes": bc})
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    changes = []
                    apply_para_format(p, s, changes)
                    apply_run_formats(p, s, changes)
                    if changes:
                        report.append({"paragraph": f"表{ti+1}单元格({ri},{ci})",
                                       "text": p.text.strip()[:40],
                                       "role": "table_cell", "changes": changes})

    text_after = extract_all_text(doc)
    integrity = "PASS" if text_before == text_after else "FAIL"

    out = {"mode": "dry-run" if args.dry_run else "apply",
           "source": args.docx,
           "content_integrity": integrity,
           "total_paragraphs": len(doc.paragraphs),
           "items_changed": len(report),
           "details": report}
    Path(args.report).write_text(json.dumps(out, ensure_ascii=False, indent=1),
                                 encoding="utf-8")
    print(f"[{out['mode']}] 共 {len(report)} 处需要/已经修改 → {args.report}")
    print(f"内容完整性校验: {integrity}")
    if integrity == "FAIL":
        sys.exit("致命错误:文本内容发生了变化,放弃保存!")

    if not args.dry_run:
        if not args.out:
            sys.exit("整改模式必须指定 -o 输出路径")
        doc.save(args.out)
        print(f"已保存 → {args.out}")


if __name__ == "__main__":
    main()
