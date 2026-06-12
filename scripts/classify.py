"""classify.py — 解析论文 docx,逐段分类,输出结构图谱 structure_map.json.

用法:
    python classify.py 输入.docx [-o structure_map.json]

输出的 JSON 是一个列表,每项:
    {"index": 段落序号, "role": 角色, "text": 前60字预览, "via": 判定依据}

角色定义见 spec.ROLES。分类基于文本特征 + 原文档大纲级别/样式名的启发式,
**可能出错**——使用本 skill 的 Claude 必须人工复核图谱(尤其 equation /
fig_caption / table_source / 各部分边界),修正后再交给 apply.py。
"""
import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).parent))
import spec

RE_CHAPTER = re.compile(r"^第\s*\d+\s*章")
RE_SEC1 = re.compile(r"^\d+\.\d+(\s|　|[^\d.])")
RE_SEC2 = re.compile(r"^\d+\.\d+\.\d+(\s|　|[^\d.])")
RE_SEC3 = re.compile(r"^\d+\.\d+\.\d+\.\d+(\s|　|[^\d.])")
RE_APPENDIX = re.compile(r"^附录\s*[A-Z]")
RE_APPENDIX_SEC = re.compile(r"^[A-Z]\.\d+")
RE_FIG = re.compile(r"^(续?图|Fig\.?|Figure)\s*([A-Z]|\d+)[-.．]\d+")
RE_TABLE = re.compile(r"^(续?表|Table)\s*([A-Z]|\d+)[-.．]\d+")
RE_REF_ENTRY = re.compile(r"^\[\d+\]")
RE_EQ_NUM = re.compile(r"[(（]([A-Z]|\d+)[-.．]\d+[)）]\s*$")
RE_KEYWORDS = re.compile(r"^(关键词|Keywords?)\s*[:：]")
RE_TOC_LINE = re.compile(r"\.{4,}\s*[IVXivx\d]+\s*$|…+\s*\d+\s*$")
RE_SOURCE = re.compile(r"^(资料来源|数据来源|来源)\s*[:：]")

CHAPTER_LEVEL_SET = {t.replace(" ", "") for t in spec.CHAPTER_LEVEL_TITLES}


def has_math(paragraph) -> bool:
    return bool(paragraph._p.findall(
        ".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))


def norm(text: str) -> str:
    return re.sub(r"[\s　]+", "", text)


def classify_paragraph(p, state):
    """返回 (role, via)。state 记录当前所在部分,影响判定。"""
    text = p.text.strip()
    n = norm(text)
    style = (p.style.name or "") if p.style else ""

    if not text and not has_math(p):
        return "skip", "空段落"

    # —— 目录区内:带引导线/行尾页码的行优先判为目录行,防止"第N章…1"误判 ——
    if state["part"] == "toc" and RE_TOC_LINE.search(text):
        if RE_CHAPTER.match(text) or norm(re.split(r"[.…]", text)[0]) in CHAPTER_LEVEL_SET:
            return "toc_chapter", "目录章行"
        return "toc_entry", "目录行"

    # —— 章级标题与部分切换 ——
    if RE_CHAPTER.match(text):
        state["part"] = "body"
        return "chapter_title", "正则:第N章"
    if n in CHAPTER_LEVEL_SET or n.rstrip("(（一二三四五六七八九十)）") in CHAPTER_LEVEL_SET:
        state["part"] = {"摘要": "abstract_cn", "Abstract": "abstract_en",
                         "目录": "toc", "参考文献": "references",
                         "致谢": "back", "声明": "back",
                         }.get(n, "front" if state["part"] in
                               ("front", "abstract_cn", "abstract_en", "toc") else "back")
        if n == "Abstract":
            return "abstract_en_title", "标题:Abstract"
        if n in ("插图和附表清单", "插图清单", "附表清单", "符号和缩略语说明"):
            state["part"] = "lists"
        return "chapter_title", f"章级标题:{n}"
    if RE_APPENDIX.match(text) and len(text) < 60:
        state["part"] = "appendix"
        return "chapter_title", "附录标题"

    # —— 节标题(正文/附录内)——
    if state["part"] in ("body", "appendix"):
        if RE_SEC3.match(text):
            return "section3", "正则:N.M.K.L"
        if RE_SEC2.match(text):
            return "section2", "正则:N.M.K"
        if RE_SEC1.match(text) and len(text) < 60:
            return "section1", "正则:N.M"
        if state["part"] == "appendix" and RE_APPENDIX_SEC.match(text) and len(text) < 60:
            return "section1", "附录节标题"

    # —— 图表题注、来源 ——
    if RE_FIG.match(text):
        return "fig_caption", "正则:图N-M"
    if RE_TABLE.match(text):
        return "table_caption", "正则:表N-M"
    if RE_SOURCE.match(text):
        return "table_source", "正则:资料来源"

    # —— 公式段:含 OMML 或以编号(N-M)结尾且文字极少 ——
    if has_math(p):
        return "equation", "含OMML公式"
    if RE_EQ_NUM.search(text) and len(re.sub(r"[^\u4e00-\u9fff]", "", text)) <= 4:
        return "equation", "行末公式编号"

    # —— 各部分内容 ——
    if state["part"] == "toc" or RE_TOC_LINE.search(text):
        if RE_CHAPTER.match(text) or norm(text.split(".")[0]) in CHAPTER_LEVEL_SET:
            return "toc_chapter", "目录章行"
        return "toc_entry", "目录行"
    if state["part"] == "lists":
        return "toc_entry", "清单/符号说明行"
    if state["part"] == "references":
        if RE_REF_ENTRY.match(text):
            return "reference_entry", "正则:[N]"
        return "reference_entry", "参考文献部分内段落"
    if RE_KEYWORDS.match(text):
        return "keywords", "关键词行"
    if state["part"] == "cover":
        return "cover", "封面/前置页(不动)"

    # 兜底:正文
    via = f"兜底正文(原样式:{style})" if style else "兜底正文"
    return "body", via


PAT_NUM = re.compile(r"(?:续?[图表]|式|[（(])\s*(?:[A-Z]|\d{1,2})([-.．])\d{1,3}")


def classify_document(docx_path: str):
    doc = Document(docx_path)
    state = {"part": "cover"}  # cover → front → body → references → appendix → back
    paragraphs = []
    census = {"hyphen": 0, "dot": 0}
    for i, p in enumerate(doc.paragraphs):
        role, via = classify_paragraph(p, state)
        paragraphs.append({
            "index": i,
            "role": role,
            "text": p.text.strip()[:60],
            "via": via,
        })
        for m in PAT_NUM.finditer(p.text):
            census["hyphen" if m.group(1) == "-" else "dot"] += 1
    # 表格:默认全部按三线表整改;布局表/封面表请人工改为 keep_borders 或 skip
    tables = []
    for ti, t in enumerate(doc.tables):
        first_row = " | ".join(c.text.strip()[:12] for c in t.rows[0].cells[:4]) \
            if t.rows else ""
        tables.append({"index": ti, "rows": len(t.rows),
                       "cols": len(t.columns) if t.rows else 0,
                       "preview": first_row,
                       "role": "three_line"})
        for row in t.rows:
            for cell in row.cells:
                for m in PAT_NUM.finditer(cell.text):
                    census["hyphen" if m.group(1) == "-" else "dot"] += 1
    return {"paragraphs": paragraphs, "tables": tables,
            "numbering_census": census}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default="structure_map.json")
    args = ap.parse_args()
    result = classify_document(args.docx)
    Path(args.out).write_text(
        json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    roles = {}
    for r in result["paragraphs"]:
        roles[r["role"]] = roles.get(r["role"], 0) + 1
    print(f"共 {len(result['paragraphs'])} 段, {len(result['tables'])} 个表格 → {args.out}")
    for k, v in sorted(roles.items(), key=lambda x: -x[1]):
        print(f"  {k:18s} {v}")
    c = result["numbering_census"]
    print(f"编号体例: '-' {c['hyphen']} 处, '.' {c['dot']} 处"
          + (" [警告] 混用,需运行 normalize_numbering.py 统一"
             if c["hyphen"] and c["dot"] else ""))


if __name__ == "__main__":
    main()
