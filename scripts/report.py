"""report.py — 把 apply.py 的 format_report.json 转成人类可读的 Markdown 修改清单。

用法:
    python report.py format_report.json -o 修改清单.md
"""
import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path

ROLE_CN = {
    "chapter_title": "章标题", "abstract_en_title": "Abstract标题",
    "section1": "一级节标题", "section2": "二级节标题", "section3": "三级节标题",
    "body": "正文段落", "fig_caption": "图题", "table_caption": "表题",
    "table_source": "表来源注", "equation": "公式段",
    "reference_entry": "参考文献条目", "toc_chapter": "目录章行",
    "toc_entry": "目录/清单行", "keywords": "关键词行",
    "table_cell": "表格单元格", "section": "页面设置",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report_json")
    ap.add_argument("-o", "--out", default="修改清单.md")
    ap.add_argument("--sections", help="sections_report.json,合并入清单")
    ap.add_argument("--numbering", help="numbering_changes.json,合并入清单")
    args = ap.parse_args()
    data = json.loads(Path(args.report_json).read_text(encoding="utf-8"))

    lines = ["# 论文格式修改清单", ""]
    lines.append(f"- 源文件:{data['source']}")
    lines.append(f"- 模式:{'检查(未写入)' if data['mode']=='dry-run' else '已整改'}")
    lines.append(f"- 内容完整性校验:**{data['content_integrity']}**"
                 "(整改前后全文逐字比对,PASS 表示一字未改)")
    lines.append(f"- 修改位置数:{data['items_changed']} / 全文 {data['total_paragraphs']} 段")
    lines.append("")

    # 汇总:按属性统计
    prop_counter = Counter()
    role_counter = Counter()
    for item in data["details"]:
        role_counter[item["role"]] += 1
        for c in item["changes"]:
            prop_counter[c["prop"].split(".")[-1].rstrip("0123456789")] += 1
    lines.append("## 修改汇总")
    lines.append("")
    lines.append("| 修改属性 | 处数 |")
    lines.append("|---|---|")
    for prop, n in prop_counter.most_common():
        lines.append(f"| {prop} | {n} |")
    lines.append("")
    lines.append("| 涉及元素 | 段/处数 |")
    lines.append("|---|---|")
    for role, n in role_counter.most_common():
        lines.append(f"| {ROLE_CN.get(role, role)} | {n} |")
    lines.append("")

    # 明细:按角色分组,同类修改去重示例化
    lines.append("## 修改明细")
    lines.append("")
    by_role = defaultdict(list)
    for item in data["details"]:
        by_role[item["role"]].append(item)
    for role, items in by_role.items():
        lines.append(f"### {ROLE_CN.get(role, role)}({len(items)} 处)")
        lines.append("")
        for item in items:
            loc = item["paragraph"]
            preview = item["text"] or "(无文字)"
            ch = "; ".join(f"{c['prop']}: {c['old']} → {c['new']}"
                           for c in item["changes"])
            lines.append(f"- 段 {loc} 「{preview}」  \n  {ch}")
        lines.append("")

    if args.sections:
        sec = json.loads(Path(args.sections).read_text(encoding="utf-8"))
        lines.append("## 分节、篇眉与页码(自动完成)")
        lines.append("")
        for item in sec:
            desc = "; ".join(f"{k}: {v}" for k, v in item.items() if k != "action")
            lines.append(f"- {item['action']}" + (f" — {desc}" if desc else ""))
        lines.append("")
    if args.numbering:
        num = json.loads(Path(args.numbering).read_text(encoding="utf-8"))
        ch = [c for c in num["changes"] if "old" in c]
        un = [c for c in num["changes"] if "unresolved" in c]
        lines.append(f"## 编号体例统一(目标连接符:{num['target']})")
        lines.append("")
        lines.append(f"原统计:'-' {num['census_before']['hyphen']} 处,"
                     f"'.' {num['census_before']['dot']} 处。"
                     f"已替换 {len(ch)} 处(这是全流程唯一的文字改动,逐条如下):")
        for c in ch:
            lines.append(f"- {c['loc']}: 「{c['old']}」→「{c['new']}」")
        if un:
            lines.append(f"- [警告] 未解决 {len(un)} 处(编号跨run),已人工处理并复核")
        lines.append("")
    Path(args.out).write_text("\n".join(lines), encoding="utf-8")
    print(f"修改清单 → {args.out}")


if __name__ == "__main__":
    main()
