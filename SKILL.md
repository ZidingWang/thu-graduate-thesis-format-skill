---
name: thu-phd-thesis-format
description: 将已完成的清华大学研究生学位论文(博士/硕士)Word 文档整改为符合《清华大学研究生学位论文写作指南》(2025年3月版)的格式,不改动任何文字内容,并输出完整修改清单与内容完整性校验。当用户提到清华博士论文/硕士论文/学位论文格式、论文排版规范、格式检查、格式整改、写作指南、研究生院模板,或上传论文 .docx 要求"改成符合学校格式""检查格式是否合规""按模板调整格式"时,务必使用本 skill——即使用户没有明说"格式整改"四个字,只要任务涉及清华研究生学位论文的字体、字号、行距、页边距、页眉页码、图表公式编号、参考文献格式,都应使用。
---

# 清华大学研究生学位论文格式整改

把用户写好的论文 docx 按 2025 版《研究生学位论文写作指南》整改格式。三条铁律:

1. **内容零改动**——任何文字(包括标点、空格、编号文字)都不许增删改。脚本内置整改前后全文逐字比对,FAIL 即放弃保存。
2. **格式全合规**——以 `references/format-spec.md` 为唯一标准,数值不靠记忆。
3. **修改全记录**——每一处修改记录"位置 + 属性 + 旧值 → 新值",最终生成用户可核对的修改清单。

## 依赖

`pip install python-docx lxml`(若环境无网络,先确认已预装)。
输入必须是 .docx;若用户给 .doc,先转换:
`soffice --headless --convert-to docx 文件.doc --outdir 目录/`

## 工作流(七步,全自动,顺序执行)

### 第 1 步:解析分类

```bash
python scripts/classify.py 论文.docx -o structure_map.json
```

输出 `paragraphs`(逐段角色)、`tables`(表格清单,默认 three_line)、`numbering_census`(编号体例"-"/"."统计)。

### 第 2 步:复核图谱(由 Claude 完成,不需用户动手)

classify.py 是启发式的,必然有误判。Claude 通读 structure_map.json 并修正:

- 各部分边界:摘要、目录、正文第 1 章、参考文献、附录、致谢的起止段。
- `equation`:无编号公式段、图片公式易被判成 body。
- `fig_caption` / `table_caption`:英文对照题注(Fig 2.1 …)、无空格编号的题注。
- `tables`:布局用表格(封面信息表、签名表)改为 `keep_borders`(只统一文字)或 `skip`(完全不动),数据表保持 `three_line`。
- 摘要之前的段一律 `cover`。

修正直接改 JSON 的 `role` 字段。结构特殊(著者-出版年制、人文社科脚注体系)时向用户确认。

### 第 3 步:检查(dry-run)

```bash
python scripts/apply.py 论文.docx -m structure_map.json --dry-run -r before_report.json
```

把"整改前共 N 处不合规"概况告诉用户,然后继续。

### 第 4 步:整改字体/段落/表格

```bash
python scripts/apply.py 论文.docx -m structure_map.json -o step4.docx -r format_report.json
```

覆盖:页边距、各角色的中西文字体/字号/对齐/行距/段前后/缩进、表格单元格文字、**三线表线型(上下 1.5 磅、表头下 1 磅、清内线)**。
结尾必须出现 `内容完整性校验: PASS`,FAIL 时脚本拒绝保存。

### 第 5 步:分节、篇眉、页码(全自动)

```bash
python scripts/sections.py step4.docx -m structure_map.json -o step5.docx -r sections_report.json
```

自动完成:每部分前插分节符(正文第一章为**奇数页**分节符,实现"另页右页开始");每节篇眉=该部分章标题(五号宋体/TNR 居中,奇偶相同、首页同);页脚 PAGE 域居中 TNR 五号;**摘要起罗马数字Ⅰ连续编排,正文第 1 章起阿拉伯数字 1 连续编排**;摘要前的封面区篇眉页脚清空;写入 updateFields——**Word 打开文件时自动刷新目录/清单页码,用户无需按 F9**。
若文档含封面页,用 `--cn-cover 部分序号` / `--en-cover 部分序号` 自动应用封面专用页边距。

### 第 6 步:编号体例统一(检测到混用时执行)

classify 已统计"-"与"."。若混用:

```bash
python scripts/normalize_numbering.py step5.docx --to auto -o step6.docx -r numbering_changes.json
```

这是流水线中**唯一会改文字的脚本**,改动严格限定为编号连接符一个字符,自动统一为全文多数派,每处替换逐条记录。报告中如有 `unresolved`(编号被拆进多个 run),Claude 用 python-docx 定位该段手工合并替换并补记录。未混用则跳过本步。

### 第 7 步:终验、清单、交付

```bash
python scripts/verify.py 论文.docx 最终.docx --allow-numbering
python scripts/report.py format_report.json -o 修改清单.md
```

verify.py 对比原始与最终文档全文:除已授权的编号连接符外,任何文字差异都 FAIL。
交付内容(present_files):最终 docx + 修改清单.md(合并 format/sections/numbering 三份报告的要点)。
建议的自动化质检:`soffice --headless --convert-to pdf 最终.docx`,渲染 PDF 后抽查篇眉、页码、分节是否正确(容器内缺中文字体时文字可能显示异常,但版式检查仍有效)。

## 仍需告知用户的事项(不是操作,只是知会)

1. 封面**内容**(题目/姓名/导师签名)出自学校单独模板,本 skill 整改其页面参数但不生成内容。
2. 双面打印是打印机设置;文档侧的奇偶分页已由奇数页分节符保证。
3. 目录若是手打文本而非 Word 域:格式已统一、updateFields 不会作用于它,建议改用自动目录(可主动提出帮用户把手打目录替换为 TOC 域,经确认后执行)。

## 与规范有出入时的决策原则

- 拿不准某段角色:宁可标 `skip` 不动它并报告,不要猜。
- 规范允许二选一的(公式居中 vs 缩进两字符;编号 "-" vs "."):跟随论文现状中占多数的一种,并在清单中注明。
- 院系另有要求的:指南明确"以院系要求为准",用户提供院系规定时,院系规定优先,差异处在清单中标注。
- 用户论文是英文论文或中英混排特殊情况:先和用户确认适用范围再动手。

## 参考文件

- `references/format-spec.md` — 全部格式参数(字体/字号/行距/磅值/页面),动手前必读,数值以它为准。
- `references/document-structure.md` — 论文组成部分顺序与各部分要点。
- `scripts/spec.py` — format-spec.md 的代码化常量,二者必须一致。

## 验收自查清单(交付前过一遍)

- [ ] 内容完整性校验 PASS
- [ ] 章标题三号黑体居中、段前24磅段后18磅;节标题14pt/13pt黑体居左
- [ ] 正文小四宋体/TNR、固定行距20磅、首行缩进2字符
- [ ] 参考文献五号、固定16磅、悬挂缩进2字符
- [ ] 图题在下(前6后12)、表题在上(前12后6)、均11pt居中
- [ ] 页边距 3.0/3.0/3.0/3.0 cm,页眉页脚距 2.2 cm
- [ ] 分节正确:篇眉=各部分标题;前置罗马页码、正文阿拉伯页码;正文奇数页起
- [ ] 三线表线型:上下 1.5 磅、表头下 1 磅
- [ ] 编号体例全文统一,numbering_changes.json 无 unresolved
- [ ] verify.py 终验 PASS
- [ ] 修改清单完整,知会事项已告知用户
