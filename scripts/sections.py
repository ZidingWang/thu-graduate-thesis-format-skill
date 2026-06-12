"""sections.py — 自动完成分节、篇眉、页码。零人工。

用法:
    python sections.py 输入.docx -m structure_map.json -o 输出.docx -r sections_report.json
        [--cn-cover IDX] [--en-cover IDX]

功能:
1. 在每个"部分"(章级标题段)起始处自动插入分节符:
   - 正文第一章:奇数页分节符(另页右页开始);其余:下一页分节符。
2. 每节篇眉 = 该部分章标题文字,五号(10.5pt),中文宋体/西文 TNR,居中,奇偶相同,首页同。
3. 页脚页码 PAGE 域,居中,TNR 五号,无修饰:
   - 摘要 → 符号和缩略语说明:大写罗马数字,从 Ⅰ 起连续;
   - 正文第一章 → 文末:阿拉伯数字,从 1 起连续。
4. 摘要之前的前置区(封面/名单/授权):篇眉页脚清空。
   --cn-cover/--en-cover 指定封面所在段的"部分序号"时,对该节应用封面专用页边距。
5. settings.xml 写入 updateFields,Word 打开时自动刷新目录/清单等所有域(免按 F9)。

必须在 apply.py 之后运行(二者都不增删段落,structure_map 始终有效)。
"""
import argparse
import copy
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

sys.path.insert(0, str(Path(__file__).parent))
import spec

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

FRONT_ROMAN_TITLES = {"摘要", "Abstract", "目录", "插图和附表清单", "插图清单",
                      "附表清单", "符号和缩略语说明"}
PART_START_ROLES = {"chapter_title", "abstract_en_title"}

CN_COVER_MARGINS = dict(top=6.0, bottom=6.0, left=4.0, right=4.0)
EN_COVER_MARGINS = dict(top=5.5, bottom=5.0, left=3.6, right=3.6)


def norm(t):
    return "".join(t.split()).replace("　", "")


def find_parts(smap_paragraphs):
    """返回 [(start_index, title_text, kind)], kind ∈ cover/front/body/back."""
    parts = []
    seen_abstract = False
    seen_body = False
    for e in smap_paragraphs:
        if e["role"] not in PART_START_ROLES:
            continue
        title = e["text"].strip()
        n = norm(title)
        if n.startswith("第") and "章" in n[:6]:
            kind = "body"
            seen_body = True
        elif n in FRONT_ROMAN_TITLES:
            kind = "front"
            seen_abstract = True
        elif n.startswith("附录"):
            kind = "body"
        elif not seen_abstract:
            kind = "cover"
        else:
            kind = "body" if not seen_body else "back"  # 引言等异名首章兜底为 body
        if n in ("参考文献", "致谢", "声明") or n.startswith("个人简历") \
                or "评语" in n or "决议书" in n:
            kind = "back"
        parts.append({"start": e["index"], "title": title, "kind": kind})
    return parts


def clone_body_sectpr(doc):
    """克隆文档末尾 body 级 sectPr 作为插入模板,剥离页眉页脚引用与页码设置。"""
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    tpl = copy.deepcopy(sectPr)
    for tag in ("w:headerReference", "w:footerReference", "w:pgNumType", "w:type"):
        for el in tpl.findall(qn(tag)):
            tpl.remove(el)
    return tpl


def insert_section_breaks(doc, parts, report):
    """在每个部分起始段的前一段 pPr 中插入 sectPr。"""
    tpl = clone_body_sectpr(doc)
    paragraphs = doc.paragraphs
    for part in parts:
        i = part["start"]
        if i == 0:
            continue  # 文档开头无需分节
        prev = paragraphs[i - 1]
        pPr = prev._p.get_or_add_pPr()
        if pPr.find(qn("w:sectPr")) is not None:
            continue  # 已有分节符
        new_sect = copy.deepcopy(tpl)
        pPr.append(new_sect)
        report.append({"action": "插入分节符", "before_paragraph": i,
                       "part": part["title"]})


def set_pgnumtype(sectPr, fmt=None, start=None):
    el = sectPr.find(qn("w:pgNumType"))
    if el is None:
        el = sectPr.makeelement(qn("w:pgNumType"), {})
        sectPr.append(el)
    if fmt:
        el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), str(start))
    elif el.get(qn("w:start")):
        del el.attrib[qn("w:start")]


def style_header_run(run):
    run.font.size = Pt(spec.HEADER_TEXT["size"])
    run.font.name = spec.F_TNR
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), spec.F_SONG)


def write_header(section, text):
    hdr = section.header
    hdr.is_linked_to_previous = False
    # 清空原有段落内容(只删 run,保留首段)
    for p in list(hdr.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p = hdr.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        run = p.add_run(text)
        style_header_run(run)


def write_footer_pagenum(section, empty=False):
    ftr = section.footer
    ftr.is_linked_to_previous = False
    for p in list(ftr.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p = ftr.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if empty:
        return
    # PAGE 域: fldChar(begin) + instrText + fldChar(end)
    def fld_run(char_type=None, instr=None):
        r = p.add_run()
        r.font.size = Pt(spec.PAGE_NUMBER["size"])
        r.font.name = spec.F_TNR
        if char_type:
            fld = r._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): char_type})
            r._r.append(fld)
        if instr:
            it = r._r.makeelement(qn("w:instrText"), {})
            it.set(qn("xml:space"), "preserve")
            it.text = instr
            r._r.append(it)
        return r
    fld_run(char_type="begin")
    fld_run(instr=" PAGE ")
    fld_run(char_type="end")


def apply_cover_margins(section, margins, report, label):
    section.top_margin = Cm(margins["top"])
    section.bottom_margin = Cm(margins["bottom"])
    section.left_margin = Cm(margins["left"])
    section.right_margin = Cm(margins["right"])
    report.append({"action": f"应用{label}页边距",
                   "values": f"上{margins['top']}/下{margins['bottom']}"
                             f"/左{margins['left']}/右{margins['right']}cm"})


def enable_update_fields(doc, report):
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        el = settings.makeelement(qn("w:updateFields"), {qn("w:val"): "true"})
        settings.append(el)
    report.append({"action": "设置打开时自动更新域(updateFields)",
                   "note": "Word 首次打开会自动刷新目录/插图清单/附表清单页码,无需按F9"})


def disable_even_odd_and_titlepg(doc, sections):
    # 篇眉奇偶相同、首页同:确保未启用 evenAndOddHeaders / titlePg
    settings = doc.settings.element
    for el in settings.findall(qn("w:evenAndOddHeaders")):
        settings.remove(el)
    for sec in sections:
        for el in sec._sectPr.findall(qn("w:titlePg")):
            sec._sectPr.remove(el)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-m", "--map", required=True)
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("-r", "--report", default="sections_report.json")
    ap.add_argument("--cn-cover", type=int, help="中文封面的部分序号(0起)")
    ap.add_argument("--en-cover", type=int, help="英文封面的部分序号(0起)")
    args = ap.parse_args()

    doc = Document(args.docx)
    smap = json.loads(Path(args.map).read_text(encoding="utf-8"))
    paragraphs_map = smap["paragraphs"] if isinstance(smap, dict) else smap
    report = []

    parts = find_parts(paragraphs_map)
    if not parts:
        sys.exit("结构图谱中未发现任何章级标题,无法分节")
    body_first = next((p for p in parts if p["kind"] == "body"), None)
    front_first = next((p for p in parts if p["kind"] == "front"), None)

    insert_section_breaks(doc, parts, report)

    # 分节后,doc.sections 与"前置区(可选)+各部分"一一对应
    sections = doc.sections
    has_preamble = parts[0]["start"] > 0  # 第一个部分之前还有内容(封面区)
    offset = 1 if has_preamble else 0
    if len(sections) != len(parts) + offset:
        sys.exit(f"分节数({len(sections)})与部分数({len(parts)}+{offset})不符,中止")

    if has_preamble:
        write_header(sections[0], "")
        write_footer_pagenum(sections[0], empty=True)
        report.append({"action": "前置区(摘要前)篇眉页脚清空", "section": 0})

    roman_started = False
    arabic_started = False
    for pi, part in enumerate(parts):
        sec = sections[pi + offset]
        sectPr = sec._sectPr
        title_n = norm(part["title"])

        if part["kind"] == "cover":
            write_header(sec, "")
            write_footer_pagenum(sec, empty=True)
            if args.cn_cover == pi:
                apply_cover_margins(sec, CN_COVER_MARGINS, report, "中文封面")
            elif args.en_cover == pi:
                apply_cover_margins(sec, EN_COVER_MARGINS, report, "英文封面")
            report.append({"action": "封面/前置部分,无篇眉页码", "part": part["title"]})
            continue

        # 分节符类型:正文第一章另页右页(奇数页),其余下一页
        sec.start_type = (WD_SECTION.ODD_PAGE if part is body_first
                          else WD_SECTION.NEW_PAGE)
        # 页码
        if part["kind"] == "front":
            set_pgnumtype(sectPr, fmt="upperRoman",
                          start=1 if not roman_started else None)
            roman_started = True
            num_desc = "罗马数字" + ("(从Ⅰ起)" if part is front_first else "(连续)")
        else:
            set_pgnumtype(sectPr, fmt="decimal",
                          start=1 if not arabic_started else None)
            arabic_started = True
            num_desc = "阿拉伯数字" + ("(从1起)" if part is body_first else "(连续)")
        # 篇眉与页脚
        write_header(sec, part["title"])
        write_footer_pagenum(sec)
        report.append({"action": "设置篇眉与页码", "part": part["title"],
                       "篇眉": part["title"], "页码": num_desc,
                       "分节类型": "奇数页" if part is body_first else "下一页"})

    disable_even_odd_and_titlepg(doc, sections)
    enable_update_fields(doc, report)

    doc.save(args.out)
    Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=1),
                                 encoding="utf-8")
    print(f"分节 {len(sections)} 节,篇眉/页码已写入 → {args.out}")
    print(f"操作记录 → {args.report}")


if __name__ == "__main__":
    main()
