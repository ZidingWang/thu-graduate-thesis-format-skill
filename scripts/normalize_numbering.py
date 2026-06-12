"""normalize_numbering.py — 统一图/表/式编号连接符("-" 与 "."),全文一致。

这是整条流水线中唯一会改动文字的脚本。改动范围被严格限定为编号连接符
这一个字符,每一处替换都记录在 numbering_changes.json,供用户逐条核对。

用法:
    统计(不改):  python normalize_numbering.py 输入.docx --census
    统一(自动选多数派): python normalize_numbering.py 输入.docx --to auto -o 输出.docx
    指定体例:    python normalize_numbering.py 输入.docx --to hyphen|dot -o 输出.docx
"""
import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document

# 仅匹配 图/表/式 引导的编号,以及 式（N-M）/（N.M） 这种括号编号
PAT_LEAD = re.compile(r"(续?[图表]|式)(\s*)([A-Z]|\d{1,2})([-.．])(\d{1,3})")
PAT_PAREN = re.compile(r"([（(])([A-Z]|\d{1,2})([-.．])(\d{1,3})([)）])")

CONNECTOR = {"hyphen": "-", "dot": "."}


def census_text(text, counter):
    for m in PAT_LEAD.finditer(text):
        counter["hyphen" if m.group(4) == "-" else "dot"] += 1
    for m in PAT_PAREN.finditer(text):
        counter["hyphen" if m.group(3) == "-" else "dot"] += 1


def replace_in_runs(p, target, changes, loc):
    """在段落 run 内做替换。编号若被拆进多个 run,先尝试 run 内替换;
    跨 run 的编号在汇总文本里检测到但 run 内替换不了时记入 unresolved。"""
    para_text = p.text
    expected = []
    for m in PAT_LEAD.finditer(para_text):
        if m.group(4) != target:
            expected.append(m.group(0))
    for m in PAT_PAREN.finditer(para_text):
        if m.group(3) != target:
            expected.append(m.group(0))
    if not expected:
        return
    done = 0
    for run in p.runs:
        t = run.text
        if not t:
            continue
        new_t = PAT_LEAD.sub(
            lambda m: m.group(0) if m.group(4) == target
            else f"{m.group(1)}{m.group(2)}{m.group(3)}{target}{m.group(5)}", t)
        new_t = PAT_PAREN.sub(
            lambda m: m.group(0) if m.group(3) == target
            else f"{m.group(1)}{m.group(2)}{target}{m.group(4)}{m.group(5)}", new_t)
        if new_t != t:
            n = sum(1 for a, b in zip(t, new_t) if a != b)
            changes.append({"loc": loc, "old": t.strip()[:50],
                            "new": new_t.strip()[:50], "替换字符数": n})
            run.text = new_t
            done += n
    still = []
    for m in PAT_LEAD.finditer(p.text):
        if m.group(4) != target:
            still.append(m.group(0))
    for m in PAT_PAREN.finditer(p.text):
        if m.group(3) != target:
            still.append(m.group(0))
    if still:
        changes.append({"loc": loc, "unresolved": still,
                        "note": "编号跨run拆分,未能自动替换,需复核"})


def iter_all_paragraphs(doc):
    for i, p in enumerate(doc.paragraphs):
        yield f"段{i}", p
    for ti, t in enumerate(doc.tables):
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    yield f"表{ti+1}({ri},{ci})", p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--census", action="store_true", help="只统计不修改")
    ap.add_argument("--to", choices=["hyphen", "dot", "auto"])
    ap.add_argument("-o", "--out")
    ap.add_argument("-r", "--report", default="numbering_changes.json")
    args = ap.parse_args()

    doc = Document(args.docx)
    counter = {"hyphen": 0, "dot": 0}
    for _, p in iter_all_paragraphs(doc):
        census_text(p.text, counter)
    print(f"编号体例统计: 连字符'-' {counter['hyphen']} 处, 小数点'.' {counter['dot']} 处")

    if args.census:
        mixed = counter["hyphen"] > 0 and counter["dot"] > 0
        print("混用!需要统一。" if mixed else "体例已统一,无需处理。")
        return

    if not args.to or not args.out:
        sys.exit("修改模式需要 --to 和 -o")
    target_key = args.to
    if target_key == "auto":
        if counter["hyphen"] == counter["dot"] == 0:
            print("未发现编号,无需处理。")
            return
        target_key = "hyphen" if counter["hyphen"] >= counter["dot"] else "dot"
        print(f"自动选择多数派: {target_key}({CONNECTOR[target_key]})")
    target = CONNECTOR[target_key]

    changes = []
    for loc, p in iter_all_paragraphs(doc):
        replace_in_runs(p, target, changes, loc)

    out = {"target": target, "census_before": counter, "changes": changes}
    Path(args.report).write_text(json.dumps(out, ensure_ascii=False, indent=1),
                                 encoding="utf-8")
    doc.save(args.out)
    n_changed = sum(1 for c in changes if "old" in c)
    n_unres = sum(1 for c in changes if "unresolved" in c)
    print(f"已替换 {n_changed} 处,未解决 {n_unres} 处 → {args.out}")
    print(f"逐条记录 → {args.report}")
    if n_unres:
        print("[警告] 存在跨run编号,请按报告复核。")


if __name__ == "__main__":
    main()
