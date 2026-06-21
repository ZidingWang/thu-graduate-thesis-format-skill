"""清华大学研究生学位论文格式规范常量(2025年3月版).

所有数值与 references/format-spec.md 一致。修改前先改 format-spec.md。
单位:字号 pt;间距 pt(磅);页面 cm。
"""

# ---- 字体 ----
F_SONG = "宋体"
F_HEI = "黑体"
F_FANGSONG = "仿宋"
F_TNR = "Times New Roman"
F_ARIAL = "Arial"
F_CAMBRIA_MATH = "Cambria Math"

# ---- 页面设置(摘要起) ----
PAGE = dict(
    margin_top_cm=3.0, margin_bottom_cm=3.0,
    margin_left_cm=3.0, margin_right_cm=3.0,
    gutter_cm=0.0,
    header_cm=2.2, footer_cm=2.2,
)

# ---- 段落规格 ----
# 每条: dict(size, cn_font, en_font, align, line(规则,值), before, after,
#            first_indent_chars, hanging_chars, bold)
# line 规则: ("single", None) 单倍 | ("exact", 20) 固定值20磅 | ("multiple", 1.25)

CHAPTER_TITLE = dict(  # 章标题及同级标题(摘要/目录/参考文献/致谢等)
    size=16, cn_font=F_HEI, en_font=F_ARIAL, align="center",
    line=("single", None), before=24, after=18,
)
ABSTRACT_EN_TITLE = dict(  # "Abstract" 标题
    size=16, cn_font=F_ARIAL, en_font=F_ARIAL, align="center",
    line=("single", None), before=24, after=18,
)
SECTION_1 = dict(  # 一级节标题 N.M
    size=14, cn_font=F_HEI, en_font=F_ARIAL, align="left",
    line=("exact", 20), before=24, after=6,
)
SECTION_2 = dict(  # 二级节标题 N.M.K
    size=13, cn_font=F_HEI, en_font=F_ARIAL, align="left",
    line=("exact", 20), before=12, after=6,
)
SECTION_3 = dict(  # 三级节标题(不建议使用)
    size=12, cn_font=F_HEI, en_font=F_ARIAL, align="left",
    line=("exact", 20), before=12, after=6,
)
BODY = dict(  # 正文段落(也适用于附录/致谢/简历/评语/决议书)
    size=12, cn_font=F_SONG, en_font=F_TNR, align="justify",
    line=("exact", 20), before=0, after=0, first_indent_chars=2,
)
FOOTNOTE = dict(
    size=9, cn_font=F_SONG, en_font=F_TNR, align="justify",
    line=("single", None), before=0, after=0, hanging_chars=1.5,
)
FIG_CAPTION = dict(  # 图序+图题,置于图下
    size=11, cn_font=F_SONG, en_font=F_TNR, align="center",
    line=("single", None), before=6, after=12,
)
TABLE_CAPTION = dict(  # 表序+表题,置于表上
    size=11, cn_font=F_SONG, en_font=F_TNR, align="center",
    line=("single", None), before=12, after=6,
)
TABLE_CELL = dict(
    size=11, cn_font=F_SONG, en_font=F_TNR, align="center",
    line=("single", None), before=3, after=3,
)
TABLE_SOURCE = dict(  # 表下资料来源
    size=10.5, cn_font=F_SONG, en_font=F_TNR, align="left",
    line=("single", None), before=6, after=12,
)
EQUATION = dict(
    size=12, cn_font=F_TNR, en_font=F_TNR, align="center",
    line=("single", None), before=6, after=6,
)
REFERENCE_ENTRY = dict(  # 参考文献条目
    size=10.5, cn_font=F_SONG, en_font=F_TNR, align="justify",
    line=("exact", 16), before=3, after=0, hanging_chars=2,
)
TOC_CHAPTER = dict(  # 目录中的章标题行
    size=12, cn_font=F_HEI, en_font=F_ARIAL, align="left",
    line=("exact", 20), before=0, after=0,
)
TOC_ENTRY = dict(  # 目录其余行 / 插图附表清单 / 符号缩略语内容
    size=12, cn_font=F_SONG, en_font=F_TNR, align="left",
    line=("exact", 20), before=0, after=0,
)
HEADER_TEXT = dict(  # 篇眉
    size=10.5, cn_font=F_SONG, en_font=F_TNR, align="center",
    line=("single", None), before=0, after=0,
)
PAGE_NUMBER = dict(size=10.5, en_font=F_TNR)  # 页脚页码,居中,无修饰线

# 表格三线规格(磅)
TABLE_RULE_OUTER_PT = 1.5  # 上、下边线
TABLE_RULE_INNER_PT = 1.0  # 栏目线(表头下)

# ---- 结构顺序(用于结构核对) ----
DOCUMENT_ORDER = [
    "中文封面", "英文封面", "名单", "授权说明",
    "中文摘要", "Abstract", "目录", "插图和附表清单", "符号和缩略语说明",
    "正文", "参考文献", "附录", "致谢", "声明",
    "个人简历", "指导教师评语", "答辩委员会决议书",
]

# 与章标题同级的前后置部分标题(用于篇眉与分类)
CHAPTER_LEVEL_TITLES = [
    "摘要", "Abstract", "目录", "插图和附表清单", "插图清单", "附表清单",
    "符号和缩略语说明", "参考文献", "致谢", "声明",
    "个人简历、在学期间完成的相关学术成果",
    "指导教师评语", "指导小组评语", "答辩委员会决议书",
]

# 段落分类标签(classify/check/apply 共用)
ROLES = [
    "chapter_title",      # 第N章 X / 同级标题
    "abstract_en_title",  # Abstract
    "section1", "section2", "section3",
    "body",
    "fig_caption", "table_caption", "table_source",
    "equation",
    "reference_entry",
    "toc_chapter", "toc_entry",
    "keywords",           # 关键词行(按正文格式)
    "symbol_entry",       # 符号和缩略语说明条目:有专用版式(tab对齐/悬挂缩进),完全不动
    "cv_entry",           # 个人简历、在学期间完成的相关学术成果:子标题居中/类别标签/
                          # 参考文献式条目混排,整段保留模板版式,完全不动
    "cover", "skip",      # 封面/授权页等不动的部分
]
# 注意:symbol_entry / cv_entry 故意不在 ROLE_SPEC 中——apply.py 对 role not in
# ROLE_SPEC 直接跳过,从而保留模板里的专用版式(符号说明的两列对齐、个人简历的
# 居中子标题与参考文献式成果条目),不被统一成正文格式改乱。

ROLE_SPEC = {
    "chapter_title": CHAPTER_TITLE,
    "abstract_en_title": ABSTRACT_EN_TITLE,
    "section1": SECTION_1,
    "section2": SECTION_2,
    "section3": SECTION_3,
    "body": BODY,
    "fig_caption": FIG_CAPTION,
    "table_caption": TABLE_CAPTION,
    "table_source": TABLE_SOURCE,
    "equation": EQUATION,
    "reference_entry": REFERENCE_ENTRY,
    "toc_chapter": TOC_CHAPTER,
    "toc_entry": TOC_ENTRY,
    "keywords": BODY,
}
