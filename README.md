# thu-phd-thesis-format-skill

清华大学研究生学位论文(博士/硕士)格式自动整改 [Claude Skill](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview)。

依据《清华大学研究生学位论文写作指南》(研究生院,2025年3月版),在**不改动任何文字内容**的前提下,把已写好的论文 `.docx` 自动整改为符合学校格式规范的文档,并生成逐条修改清单供核对。

> 姊妹项目:[Toflamus/tsinghua_senior_design_word_template_skill](https://github.com/Toflamus/tsinghua_senior_design_word_template_skill) —— 本科生综合论文训练(从零驱动学校模板生成)。本项目面向**已写完、需要整改格式**的研究生学位论文场景,二者可配合使用。

## 它做什么

读取你写好的论文 docx,自动完成:

- **字体字号**:章标题/节标题/正文/图表题注/公式/参考文献等各类段落,按规范分别设置中西文字体与字号
- **段落格式**:对齐方式、行距(单倍/固定值)、段前段后间距、首行缩进、悬挂缩进
- **页面设置**:页边距 3.0/3.0/3.0/3.0 cm、页眉页脚距 2.2 cm
- **分节、篇眉与页码**:自动按各部分插入分节符,篇眉=该部分章标题,摘要起罗马数字页码、正文第1章起阿拉伯数字页码(且从奇数页开始),写入 `updateFields` 使 Word 打开时自动刷新目录
- **三线表**:上下边线 1.5 磅、表头下线 1 磅、清除多余内线
- **图表公式编号体例**:自动统计全文 `图1-1`/`图1.1` 等连接符使用情况,统一为多数派写法(全流程唯一允许的文字改动,逐条记录)
- **终验**:整改前后全文逐字比对,确保除已授权的编号连接符外,内容完全一致

## 它不做什么

- 不生成或修改**封面内容**(题目/姓名/导师签名等)——这部分使用学校单独的封面模板
- 不涉及打印设置(双面打印由打印机完成,文档侧的奇偶分页已通过分节符保证)
- 不重写参考文献内容、不改图表数据——只调整格式

## 从 GitHub 到使用

如果你是在 GitHub 上看到这个项目,按下面方式安装后即可在 Claude 中使用。

### 方式 A:Claude.ai 网页/桌面端

1. 打开项目页面:[ZidingWang/thu-phd-thesis-format-skill](https://github.com/ZidingWang/thu-phd-thesis-format-skill)。
2. 下载 `dist/thu-phd-thesis-format.skill`:
   - 若项目有 Release,优先在 Release 中下载 `.skill` 文件。
   - 若没有 Release,进入仓库的 `dist/` 目录,打开 `thu-phd-thesis-format.skill`,点击下载原始文件。
3. 打开 Claude.ai 或 Claude 桌面端,进入 `Settings` → `Capabilities` → `Skills`。
4. 上传刚下载的 `thu-phd-thesis-format.skill`。
5. 新建对话,上传你的论文 `.docx`,对 Claude 说:

> 帮我把这篇博士论文按清华学校格式整改一下

Claude 会自动调用本 skill,生成整改后的 `.docx`、逐条 `修改清单.md` 和内容完整性校验结果。

### 方式 B:Claude Code

```bash
mkdir -p ~/.claude/skills
git clone https://github.com/ZidingWang/thu-phd-thesis-format-skill.git \
  ~/.claude/skills/thu-phd-thesis-format
pip install -r ~/.claude/skills/thu-phd-thesis-format/requirements.txt
```

安装后,在 Claude Code 对话中上传或指定论文 `.docx`,说类似:

> 帮我把这篇博士论文按清华学校格式整改一下

Claude Code 会自动调用本 skill,依次完成解析分类、格式整改、分节分页、编号统一、终验,并交付:

- 整改后的 `.docx`
- `修改清单.md` —— 逐条记录每一处修改(位置 + 属性 + 旧值 → 新值)
- 内容完整性校验结果

### 依赖说明

脚本依赖 `python-docx` 和 `lxml`。Claude Code 用户可用 `requirements.txt` 安装;Claude.ai/桌面端上传 `.skill` 后,Claude 会在可用环境中按 skill 指令调用这些脚本。

## 目录结构

```
thu-phd-thesis-format/
├── README.md                      # 项目说明与安装方式
├── LICENSE                        # MIT License
├── requirements.txt               # Python 依赖
├── SKILL.md                       # 工作流定义(七步,全自动)
├── references/
│   ├── format-spec.md             # 完整格式规范参数表(字体/字号/行距/磅值/页面)
│   └── document-structure.md      # 论文组成部分顺序与各部分要点
├── scripts/
│   ├── spec.py                    # 规范常量(与 format-spec.md 一致)
│   ├── classify.py                # 解析论文,逐段分类生成结构图谱
│   ├── apply.py                   # 整改字体/段落/表格,内置完整性校验
│   ├── sections.py                # 分节、篇眉、页码自动化
│   ├── normalize_numbering.py     # 图表公式编号体例统一
│   ├── verify.py                  # 整改前后终验
│   └── report.py                  # 生成修改清单 Markdown
├── evals/
│   └── make_sample.py             # 生成测试样本(格式故意不合规的模拟论文)
├── examples/
│   ├── 示例-整改后.docx
│   └── 示例-修改清单.md
└── dist/
    └── thu-phd-thesis-format.skill # Claude.ai 上传包
```

## 开发与验证

生成一份格式故意不合规的测试论文:

```bash
python evals/make_sample.py -o sample_thesis.docx
```

然后可按 `SKILL.md` 中的七步流水线依次运行 `classify.py`、`apply.py`、`sections.py`、`verify.py` 和 `report.py`。

## 适用范围与局限

- 基于学校 2025 年 3 月版《写作指南》提炼,若学校后续更新版本导致规范数值变化,请更新 `references/format-spec.md` 与 `scripts/spec.py`
- `classify.py` 基于正则与启发式规则分类段落,结构特殊的论文(著者-出版年制参考文献、人文社科脚注体系、复杂分图等)可能需要人工修正结构图谱后再继续
- 院系如有比学校规范更严格的额外要求,以院系要求为准,本 skill 仅保证符合研究生院统一规范

## License

MIT,见 [LICENSE](LICENSE)。
