---
description: 交互式学术论文 Word 自动排版——参数前置收集，然后 19 步一键执行
---

# Word Auto Typesetter

交互式收集论文排版要求，通过 Python `win32com.client` 自动修改 Word 文档格式。
**核心原则：只改格式，不改内容。**

## 触发

```
/word-typesetter <目标文件路径.docx>
```

## 阶段一：参数收集（必须执行，不可跳过）

启动后输出以下问卷，等待用户回复。**绝对不在用户回复前执行任何代码。**

```
请填写以下排版参数（括号内为默认值，无需修改的直接回车跳过）：

=== 1. 页面设置 ===
纸张大小（默认 A4）：
上/下/内侧/外侧页边距（厘米，默认 2）：
装订线（厘米，默认 0）：
纸张方向（竖向/横向，默认 竖向）：

=== 2. 页眉页脚位置 ===
页眉顶端/页脚底端距离（厘米，默认 1.5）：

=== 3. 页码格式 ===
正文前页码（罗马大写/罗马小写/阿拉伯，默认 罗马大写）：
正文页码（默认 阿拉伯数字）：

=== 4. [正文] ===
中文字体/西文/字号（默认 宋体/Times New Roman/小四）：
字形/颜色（默认 常规/黑色）：
对齐方式（默认 两端对齐）：
特殊缩进（默认 首行缩进2字符）：
段前/段后/行距（默认 0/0/单倍）：

=== 5. [标题1] ===
中文字体/字号（默认 黑体/小三）：
字形/对齐（默认 加粗/居中）：
段前/段后/行距（默认 0/0/单倍）：

=== 6. [标题2] ===
中文字体/字号（默认 黑体/四号）：
字形/对齐（默认 加粗/左对齐）：

=== 7. [标题3] ===
中文字体/字号（默认 黑体/小四）：
字形/对齐（默认 加粗/左对齐）：

=== 8. [页眉] ===
中文字体/字号（默认 宋体/小五）：
对齐/下划线（默认 居中/单线0.5磅）：

=== 9. [页脚] ===
中文字体/字号/对齐（默认 宋体/小五/居中）：

=== 10. [题注] ===
中文字体/字号/对齐（默认 中宋/五号/居中）：
连接符（默认 1-1）：

=== 11. [图片段落] ===
对齐/段前/段后/行距（默认 居中/0/0/单倍）：

=== 12. [表格正文] ===
中文字体/字号/对齐/行距（默认 宋体/小五/居中/单倍）：

=== 13. [参考文献正文] ===
中文字体/字号/对齐/行距（默认 宋体/小五/左对齐/单倍）：

=== 14. 多级列表 ===
级别1/2/3格式（默认 1 / 1.1 / 1.1.1）：

=== 15. 分节 ===
一级标题节起始（默认 新建页）：

=== 16. 页眉页脚 ===
奇偶页不同（默认 否）：
页脚对齐（默认 居中）：

=== 17. 摘要（中文） ===
《摘要》段应用标题1（默认 否）/ 下方空行（默认 否）/ 正文应用正文样式（默认 是）：
"关键词："加粗（默认 是）/ 分隔符（默认 中文分号）：

=== 18. 英文摘要 ===
《Abstract》段应用标题1（默认 否）/ 下方空行（默认 否）：
"Key Words:"加粗（默认 是）/ 分隔符（默认 英文分号）：

=== 19. 参考文献 ===
《参考文献》段应用标题1（默认 是）：
编号格式（默认 [1]）/ 编号后制表符（默认 是）：
上标去粗（默认 是）：

=== 20. 三线表 ===
表格对齐（默认 居中）/ 行距（默认 单倍）/ 首行缩进（默认 无）：

=== 21. 目录 ===
"目录"字体/字号/对齐（默认 黑体/小二/居中）：
```

## 阶段二：自动化执行（19 步，不中断）

用户确认参数后，生成并执行 Python 脚本。

### Word COM 常量速查表（生成脚本时必须使用以下数值）

```python
# 纸张 — 注意：不是直觉值！
wdPaperA4 = 7          # A4=7，不是 9！
wdPaperA5 = 9          # A5=9
wdOrientPortrait = 0   # 竖向=0，不是 1！
wdOrientLandscape = 1  # 横向=1

# 行距
wdLineSpaceSingle = 0      # 单倍行距
wdLineSpace1pt5 = 1        # 1.5倍行距
wdLineSpaceDouble = 2      # 2倍行距
wdLineSpaceAtLeast = 3     # 最小值
wdLineSpaceExactly = 4     # 固定值（搭配 LineSpacing 磅值）
wdLineSpaceMultiple = 5    # 多倍行距（搭配 LineSpacing 倍数）

# 对齐
wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdAlignParagraphJustify = 3    # 两端对齐

# 页码
wdPageNumberStyleArabic = 0
wdPageNumberStyleUppercaseRoman = 1
wdPageNumberStyleLowercaseRoman = 2

# 其他常用
wdStyleTypeParagraph = 1
wdNumberGallery = 3
wdListNumberStyleArabic = 0
wdListLevelAlignLeft = 0
wdCollapseStart = 1
wdCollapseEnd = 0
wdSectionBreakNextPage = 2
wdHeaderFooterPrimary = 1
wdBorderTop = -1
wdBorderBottom = -3
wdLineStyleNone = 0
wdLineStyleSingle = 1
wdLineWidth025pt = 6
wdLineWidth050pt = 8
wdLineWidth150pt = 12
wdInlineShapePicture = 3
wdReplaceAll = 2
wdTrailingTab = 1
```

### 字号磅值对照

| 小五 | 五号 | 小四 | 四号 | 小三 | 三号 | 小二 | 二号 | 一号 |
|------|------|------|------|------|------|------|------|------|
| 9 | 10.5 | 12 | 14 | 15 | 16 | 18 | 22 | 26 |

### 磅值换算

- 1 厘米 = 28.35 磅
- 1 字符缩进 ≈ 字号磅值 × 字符数
- 1 行段间距 ≈ 字号磅值

---

### 完整代码模板

生成脚本时根据参数填充以下模板。**所有步骤用 try/except 包裹，出错不中断。**

```python
import win32com.client, re, os

TARGET = r"<目标路径>"

# ===== 常量 =====
A4=7; PO=0; SS=0; AL=0; AC=1; AJ=3; SP=1; NG=3; NA=0; LL=0
CS=1; CE=0; SN=2; HP=1; BT=-1; BB=-3; LN=0; LS=1
LW6=6; LW8=8; LW12=12; IP=3; PA=0; PR=1; TT=1; RA=2
M=<边距厘米*28.35>; H=<页眉厘米*28.35>; F=<页脚厘米*28.35>

def it(p):
    try: return p.Range.Information(12) in (True, -1)
    except: return False

# 标记查找 — 兼容 《》 和 <>
def fm(d, t):
    r = d.Range(); f = r.Find; f.Text = t; f.MatchWildcards = False; f.Execute()
    return r.Duplicate if f.Found else None

def fm2(d, n):
    for b in [("《", "》"), ("<", ">")]:
        r = fm(d, b[0] + n + b[1])
        if r: return r
    return None

def ph(p, n):
    for b in [("《", "》"), ("<", ">")]:
        if b[0] + n + b[1] in p.Range.Text: return True
    return False

def sst(st, c, s, bd, al, fl):
    try:
        st.Font.NameFarEast = c; st.Font.Name = "Times New Roman"
        st.Font.Size = s; st.Font.Bold = bd; st.Font.Color = 0
        st.ParagraphFormat.Alignment = al
        st.ParagraphFormat.LeftIndent = 0; st.ParagraphFormat.RightIndent = 0
        st.ParagraphFormat.FirstLineIndent = fl
        st.ParagraphFormat.SpaceBefore = 0; st.ParagraphFormat.SpaceAfter = 0
        st.ParagraphFormat.LineSpacingRule = <行距常量>
    except: pass

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
doc = word.Documents.Open(TARGET)
p0 = doc.Paragraphs.Count

# === Step 1-2: 样式 + 页面 ===
sst(doc.Styles("正文"), "<正文中文字体>", <字号磅值>, <加粗True/False>, <对齐常量>, <首行缩进磅值>)
sst(doc.Styles("标题 1"), "<H1字体>", <H1字号>, <H1加粗>, <H1对齐>, 0)
sst(doc.Styles("标题 2"), "<H2字体>", <H2字号>, <H2加粗>, <H2对齐>, 0)
sst(doc.Styles("标题 3"), "<H3字体>", <H3字号>, <H3加粗>, <H3对齐>, 0)
doc.Styles("页眉").Font.Size = <页眉字号>; doc.Styles("页脚").Font.Size = <页脚字号>

# 自定义样式
for sn, c, s, bd, al in [("题注", "<题注字体>", <题注字号>, False, AC),
                           ("jh图片段落", "<正文字体>", <正文字号>, False, AC),
                           ("jh表格正文", "<表格字体>", <表格字号>, False, AC),
                           ("jh参考文献正文", "<参考文献字体>", <参考文献字号>, False, AL)]:
    try: st = doc.Styles(sn)
    except: st = doc.Styles.Add(sn, SP)
    sst(st, c, s, bd, al, 0)

# 参考文献悬挂缩进
rs = doc.Styles("jh参考文献正文")
rs.ParagraphFormat.LeftIndent = 21; rs.ParagraphFormat.FirstLineIndent = -21

# 页面设置
for sec in doc.Sections:
    try:
        ps = sec.PageSetup; ps.PaperSize = <纸张常量>; ps.Orientation = <方向常量>
        ps.TopMargin = M; ps.BottomMargin = M; ps.LeftMargin = M; ps.RightMargin = M
        ps.Gutter = <装订线磅值>; ps.HeaderDistance = H; ps.FooterDistance = F
    except: pass

# === Step 3: 多级列表 ===
try:
    lt = word.ListGalleries(NG).ListTemplates(1)
    for i, p, sn in [(1, "<L1格式>", "标题 1"), (2, "<L2格式>", "标题 2"), (3, "<L3格式>", "标题 3")]:
        lv = lt.ListLevels(i); lv.NumberFormat = p; lv.NumberStyle = NA
        lv.LinkedStyle = doc.Styles(sn); lv.NumberPosition = 0
        lv.Alignment = LL; lv.TextPosition = i * 0.5 * 28.35
except: pass

# === Step 4: 格式清洗 (逐段，跳过表格) ===
sm = fm2(doc, "正文开始"); em = fm2(doc, "正文结束")
sp = ep = None
if sm:
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        if p.Range.Start <= sm.End <= p.Range.End: sp = i; break
if em:
    for i in range(doc.Paragraphs.Count, 0, -1):
        p = doc.Paragraphs(i)
        if p.Range.Start <= em.Start <= p.Range.End: ep = i; break
if sp and ep and sp < ep:
    for i in range(sp, ep + 1):
        pa = doc.Paragraphs(i)
        if it(pa): continue  # 跳过表格内段落
        try:
            pa.Range.Font.Reset()
            pa.Range.ParagraphFormat.Reset()
            pa.Range.Style = doc.Styles("正文")
        except: pass

# === Step 5: 应用标题样式 (自底向上) ===
for i in range(doc.Paragraphs.Count, 0, -1):
    pa = doc.Paragraphs(i)
    if it(pa): continue
    for t, sn in [("标题1", "标题 1"), ("标题2", "标题 2"),
                   ("标题3", "标题 3"), ("标题4", "标题 4")]:
        if ph(pa, t):
            try: pa.Range.Style = doc.Styles(sn)
            except: pass; break

# === Step 6: 应用题注样式 ===
for i in range(doc.Paragraphs.Count, 0, -1):
    pa = doc.Paragraphs(i)
    if it(pa): continue
    if ph(pa, "题注"):
        try: pa.Range.Style = doc.Styles("题注")
        except: pass

# === Step 7: 分节 (每个《标题1》前插分节符) ===
for i in range(doc.Paragraphs.Count, 0, -1):
    if ph(doc.Paragraphs(i), "标题1"):
        r = doc.Paragraphs(i).Range.Duplicate
        r.Collapse(CS); r.InsertBreak(<分节符常量>)

# === Step 8: 摘要 (中文) ===
# 除标题段落外，全部应用正文样式
ap = aep = None
for i in range(1, doc.Paragraphs.Count + 1):
    if ap is None and ph(doc.Paragraphs(i), "摘要"): ap = i
    if ph(doc.Paragraphs(i), "Abstract"): aep = i; break
if ap and aep:
    for i in range(ap + 1, aep):
        pa = doc.Paragraphs(i)
        if it(pa): continue
        if pa.Range.Text.strip():
            try: pa.Range.Style = doc.Styles("正文")
            except: pass

# === Step 9: 英文摘要 ===
aep2 = nh = None
for i in range(1, doc.Paragraphs.Count + 1):
    if aep2 is None and ph(doc.Paragraphs(i), "Abstract"): aep2 = i
    if aep2 and i > aep2 and ph(doc.Paragraphs(i), "标题1"): nh = i; break
if aep2 and nh:
    for i in range(aep2 + 1, nh):
        pa = doc.Paragraphs(i)
        if it(pa): continue
        if pa.Range.Text.strip():
            try: pa.Range.Style = doc.Styles("正文")
            except: pass

# === Step 10: 参考文献 (Word 原生自动编号) ===
rm = fm2(doc, "参考文献")
if rm:
    rp = rm.Paragraphs(1); rp.Range.Style = doc.Styles("标题 1")
    body = doc.Range(rp.Range.End, doc.Range().End - 1)
    rps = []
    for p in body.Paragraphs:
        ts = p.Range.Text.strip()
        if not ts: continue
        hs = False
        for h in ["标题1", "标题2", "标题3", "标题4"]:
            if ph(p, h): hs = True; break
        if hs: break
        if ts.startswith("致谢") or ts.startswith("附录"): break
        rps.append(p)
    # 清除旧手动编号
    for p in rps:
        t = p.Range.Text.lstrip(); m = re.match(r'\[\d+\]\s*', t)
        if m:
            try: p.Range.Text = t[m.end():]
            except: pass
        try: p.Range.ListFormat.RemoveNumbers()
        except: pass
    # 应用原生自动编号 [1], [2], [3]...
    try:
        lt2 = word.ListGalleries(NG).ListTemplates(1)
        lv = lt2.ListLevels(1)
        lv.NumberFormat = "[%1]"; lv.NumberStyle = NA
        lv.TrailingCharacter = <编号后制表符TT/空格>; lv.NumberPosition = 0
        lv.Alignment = 0; lv.TextPosition = 21
        for p in rps:
            try:
                p.Range.Style = doc.Styles("jh参考文献正文")
                p.Range.ListFormat.ApplyListTemplate(
                    ListTemplate=lt2, ContinuePreviousList=True,
                    ApplyTo=1, DefaultListBehavior=1)
            except: pass
    except: pass

# === Step 11: 正文引用上标去粗体 ===
for pa in doc.Paragraphs:
    if it(pa): continue
    try:
        f = pa.Range.Find; f.Text = r"\[[0-9,\-]+\]"; f.MatchWildcards = True
        while f.Execute():
            if pa.Range.Font.Superscript: pa.Range.Font.Bold = False
    except: pass

# === Step 12-13: 三线表 + 表格正文 (仅正文区域内) ===
bs = sm.End if sm else None; be = em.Start if em else None
for tbl in doc.Tables:
    ts = tbl.Range.Start
    if bs and ts < bs: continue    # 跳过正文前的表格
    if be and ts > be: continue    # 跳过正文后的表格
    # 三线表边框
    tbl.Borders.InsideLineStyle = LN; tbl.Borders.OutsideLineStyle = LN
    for c in tbl.Rows(1).Cells:
        c.Borders(BT).LineStyle = LS; c.Borders(BT).LineWidth = LW12
    for c in tbl.Rows(tbl.Rows.Count).Cells:
        c.Borders(BB).LineStyle = LS; c.Borders(BB).LineWidth = LW12
    for c in tbl.Rows(1).Cells:
        c.Borders(BB).LineStyle = LS; c.Borders(BB).LineWidth = LW8
    tbl.Rows.Alignment = <表格对齐: 0左/1中/2右>  # 默认 1 居中
    # 表格正文样式
    for row in range(1, tbl.Rows.Count + 1):
        for c in tbl.Rows(row).Cells:
            for p in c.Range.Paragraphs:
                try: p.Range.Style = doc.Styles("jh表格正文")
                except: pass

# === Step 14: 图片样式 (仅正文区域内) ===
for shp in doc.InlineShapes:
    if shp.Type != IP: continue
    ip2 = shp.Range.Start
    if bs and ip2 < bs: continue
    if be and ip2 > be: continue
    try: shp.Range.Paragraphs(1).Range.Style = doc.Styles("jh图片段落")
    except: pass

# === Step 15: 页眉页脚 ===
# 找正文起始节
bsec = None
for idx, sec in enumerate(doc.Sections, 1):
    t = sec.Range.Text; hs = False
    for b in [("《", "》"), ("<", ">")]:
        if b[0] + "正文开始" + b[1] in t: hs = True
    if hs: bsec = idx; break
if bsec is None:
    for idx, sec in enumerate(doc.Sections, 1):
        t = sec.Range.Text; h1 = False; ha = False
        for b in [("《", "》"), ("<", ">")]:
            if b[0] + "标题1" + b[1] in t: h1 = True
            if b[0] + "摘要" + b[1] in t or b[0] + "Abstract" + b[1] in t: ha = True
        if h1 and not ha: bsec = idx; break
if bsec is None: bsec = 1

# 断开同前链接
for i in range(bsec, doc.Sections.Count + 1):
    try:
        doc.Sections(i).Headers(HP).LinkToPrevious = False
        doc.Sections(i).Footers(HP).LinkToPrevious = False
    except: pass

# 页眉下划线
for i in range(bsec, doc.Sections.Count + 1):
    try:
        hd = doc.Sections(i).Headers(HP)
        if hd.Exists:
            hd.Range.ParagraphFormat.Borders(BB).LineStyle = LS
            hd.Range.ParagraphFormat.Borders(BB).LineWidth = <下划线磅值常量>
    except: pass

# 页码：正文前罗马，正文阿拉伯从1开始
for i in range(1, bsec):
    try: doc.Sections(i).Footers(HP).PageNumbers.NumberStyle = <罗马常量>
    except: pass
try:
    fb = doc.Sections(bsec).Footers(HP)
    fb.PageNumbers.RestartNumberingAtSection = True
    fb.PageNumbers.StartingNumber = 1
    fb.PageNumbers.NumberStyle = PA
except: pass
for i in range(bsec + 1, doc.Sections.Count + 1):
    try: doc.Sections(i).Footers(HP).PageNumbers.RestartNumberingAtSection = False
    except: pass

# === Step 16: 目录 (在 Abstract/摘要 节之后) ===
ai = None
for idx, sec in enumerate(doc.Sections, 1):
    for b in [("《", "》"), ("<", ">")]:
        if b[0] + "Abstract" + b[1] in sec.Range.Text: ai = idx; break
    if ai: break
if ai is None:
    for idx, sec in enumerate(doc.Sections, 1):
        for b in [("《", "》"), ("<", ">")]:
            if b[0] + "摘要" + b[1] in sec.Range.Text: ai = idx; break
        if ai: break
if ai:
    sr = doc.Sections(ai).Range; sr.Collapse(CE); sr.InsertBreak(SN)
    ns2 = doc.Sections(ai + 1)
    try:
        ns2.Headers(HP).LinkToPrevious = False
        ns2.Footers(HP).LinkToPrevious = False
    except: pass
    nr = ns2.Range; nr.InsertBefore("目录\n")
    tr = doc.Range(nr.Start, nr.Start + 2)
    tr.Font.NameFarEast = "<目录字体>"; tr.Font.Size = <目录字号>
    tr.Font.Bold = True; tr.ParagraphFormat.Alignment = <目录对齐常量>
    toc = doc.TablesOfContents.Add(
        Range=doc.Range(nr.Start + 3, nr.End),
        UseHeadingStyles=True, UpperHeadingLevel=1,
        LowerHeadingLevel=3, UseFields=False, UseHyperlinks=True)
    toc.Update()

# === Step 17: 清除所有标记 (Find/Replace 保留样式) ===
for m in ["正文开始", "正文结束", "标题1", "标题2", "标题3", "标题4",
           "题注", "摘要", "摘要正文", "关键词",
           "Abstract", "英文摘要正文", "Key Words", "参考文献"]:
    for b in [("《", "》"), ("<", ">")]:
        tag = b[0] + m + b[1]
        rng = doc.Range(); f = rng.Find; f.Text = tag; f.Replacement.Text = ""
        f.MatchWildcards = False; f.Forward = True; f.Wrap = 1
        f.Execute(Replace = RA)

# 更新目录
for toc in doc.TablesOfContents:
    try: toc.Update()
    except: pass

# === Step 18: 关键词加粗 (仅标签词+冒号，不整段) ===
def bkw(pa, mk, kw):
    full = pa.Range.Text; idx = full.find(mk)
    if idx < 0: return
    after = full[idx + len(mk):]; ks = 0
    while ks < len(after) and after[ks] in (' ', '\t'): ks += 1
    cp = after.find("：", ks)
    if cp < 0: cp = after.find(":", ks)
    if cp >= 0 and cp <= ks + len(kw) + 2: be = cp + 1
    else: be = ks + len(kw)
    sp = pa.Range.Start + idx + len(mk) + ks
    ep = pa.Range.Start + idx + len(mk) + be
    if sp < ep: doc.Range(sp, ep).Font.Bold = True

for pa in doc.Paragraphs:
    t = pa.Range.Text
    for m, k in [("《关键词》", "关键词"), ("<关键词>", "关键词"),
                  ("《Key Words》", "Key Words"), ("<Key Words>", "Key Words")]:
        if m in t: bkw(pa, m, k); break
```

---

## 安全红线

1. **不删内容**：`Font.Reset()` / `ParagraphFormat.Reset()` 只清格式，不动文字
2. **不建表格**：只对 `doc.Tables` 已有表格改边框和样式
3. **逐段处理**：格式清洗不跨表格边界，`para.Range.Information(12)` 判读并跳过表格内段落
4. **不自动保存**：结果留在 Word 窗口，用户手动 Ctrl+S
5. **自底向上遍历**：插入分节符 / 修改段落时从后往前，避免索引偏移
6. **标记不伤样式**：清除标记用 Find/Replace，不用 `para.Range.Text =` 赋值
7. **双标记兼容**：`《》` 和 `<>` 都识别，生成脚本时用 `fm2()`/`ph()` 辅助函数

## 标记规范

推荐使用 `《》`（全角书名号），也兼容 `<>`（半角尖括号）。

| 标记 | 用途 |
|------|------|
| 《正文开始》/《正文结束》 | 正文区域边界 |
| 《标题1》~《标题4》 | 标题层级 |
| 《摘要》/《Abstract》 | 中英文摘要标题 |
| 《摘要正文》/《英文摘要正文》 | 摘要正文段落 |
| 《关键词》/《Key Words》 | 关键词行 |
| 《参考文献》 | 参考文献标题 |
| 《题注》 | 图表标题 |
