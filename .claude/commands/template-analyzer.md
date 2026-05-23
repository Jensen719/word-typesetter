---
description: 通用子 Agent — 从任意论文模板中全量提取排版要求，模板要求什么就输出什么
---

# Template Analyzer（通用版）

从客户论文模板文件（.doc/.docx）中提取**全部**排版要求。正文、批注、表格、页眉页脚、内置样式定义一律扫描。**模板说什么就提取什么，不预设任何具体格式要求。**

## 触发

用户 @模板文件 说"分析模板""提取排版要求""模板要求"，或被 word-typesetter 调用。

## 执行步骤

### Step 1：全量数据提取

用 win32com 打开模板，把所有可能包含排版信息的来源 dump 出来：

```python
import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
doc = word.Documents.Open(r"<模板路径>")

# 所有段落
for i in range(1, doc.Paragraphs.Count + 1):
    t = doc.Paragraphs(i).Range.Text.strip()
    if t: print("[P%d] %s" % (i, t))

# 所有批注
for i in range(1, doc.Comments.Count + 1):
    c = doc.Comments(i)
    scope = c.Scope.Text.strip()[:200] if c.Scope else ""
    text = c.Range.Text.strip()
    print("[C%d] Scope:%s | Text:%s" % (i, scope, text))

# 内置样式实际定义
for sn in ["正文","标题 1","标题 2","标题 3","标题 4","页眉","页脚",
           "Normal","Heading 1","Heading 2","Heading 3"]:
    try:
        s = doc.Styles(sn)
        print("[STYLE %s] Font_CN=%s Font_EN=%s Size=%.1f Bold=%s Align=%d LineRule=%d LineSpacing=%.1f FirstIndent=%.1f" % (
            sn, s.Font.NameFarEast, s.Font.Name, s.Font.Size, s.Font.Bold,
            s.ParagraphFormat.Alignment, s.ParagraphFormat.LineSpacingRule,
            s.ParagraphFormat.LineSpacing, s.ParagraphFormat.FirstLineIndent))
    except: pass

# 页眉页脚
for idx, sec in enumerate(doc.Sections, 1):
    for ht in [1, 3]:
        try:
            h = sec.Headers(ht)
            if h.Exists and h.Range.Text.strip():
                print("[HEADER S%d T%d] %s" % (idx, ht, h.Range.Text.strip()[:200]))
        except: pass
    try:
        f = sec.Footers(1)
        if f.Exists and f.Range.Text.strip():
            print("[FOOTER S%d] %s" % (idx, f.Range.Text.strip()[:200]))
    except: pass

# 页面设置
for idx, sec in enumerate(doc.Sections, 1):
    try:
        ps = sec.PageSetup
        print("[PAGESETUP S%d] PaperSize=%d Orient=%d Top=%.1f Bottom=%.1f Left=%.1f Right=%.1f HeaderDist=%.1f FooterDist=%.1f" % (
            idx, ps.PaperSize, ps.Orientation, ps.TopMargin, ps.BottomMargin,
            ps.LeftMargin, ps.RightMargin, ps.HeaderDistance, ps.FooterDistance))
    except: pass
```

### Step 2：提取排版要求

分析全部 dump 内容。**以下为检索方向，不设固定分类：**

- 任何关于字体、字号、加粗、倾斜、颜色的描述
- 任何关于行距、缩进、对齐、段间距的描述
- 任何关于纸张大小、边距、页眉页脚距离的描述
- 任何关于标题层级样式（一级/二级/三级等）的描述
- 任何关于摘要、关键词、Abstract 格式的描述
- 任何关于目录格式、显示级别的描述
- 任何关于图表编号、题注格式、图表对齐的描述
- 任何关于公式编号、对齐方式的描述
- 任何关于脚注、尾注格式的描述
- 任何关于参考文献编号、引用上标格式的描述
- 任何关于页码格式（罗马/阿拉伯/起始页）的描述
- 任何关于附录、致谢、科研成果等尾部内容的描述
- 任何关于页眉页脚文字内容、奇偶不同的描述
- **模板中出现的任何其他格式要求**

**提取原则**：
- 模板提到了就提取，没提到的不臆造
- 每条标注来源（批注号/段落号/样式定义）
- 模板没有明确要求的参数不自行推断

### Step 3：输出

输出以下三部分。每部分的内容**完全由当前模板决定**，不预设任何具体格式。

---

## 一、排版要求汇总

按论文结构的自然顺序逐项列出。格式：`要求内容 | 来源`。模板没提到的项目不出现在列表中。

```
【页面设置】
- 纸张大小：XXX | 来源：XXX
- 上/下/内侧/外侧边距：XXX | 来源：XXX
- ...（模板实际提到的页面参数）

【页眉页脚】
- ...（模板实际提到的页眉页脚要求）

【正文段落】
- ...（模板实际提到的正文格式要求）

【各级标题】
- ...（模板实际提到的各级标题格式要求）

【摘要】
- ...（模板实际提到的摘要格式要求）

【目录】
- ...（模板实际提到的目录格式要求）

【图表】
- ...（模板实际提到的图/表格式要求）

【公式】（如模板提及）
- ...（模板实际提到的公式格式要求）

【脚注/尾注】（如模板提及）
- ...（模板实际提到的脚注格式要求）

【参考文献】
- ...（模板实际提到的参考文献格式要求）

【其他】
- ...（模板提及但不在以上分类中的任何其他格式要求）
```

## 二、标记指南

根据模板的论文结构，列出用户需在待排版文档中添加的 `《》` 标记及放置位置。**只列出与排版操作相关的标记。**

## 三、当前 word-typesetter 无法覆盖的项

将第一步提取到的所有要求，与 word-typesetter 当前已有能力逐项比对，列出差异：

- word-typesetter **已有参数可覆盖**的 → 不列在此处，直接传入参数即可
- word-typesetter **没有对应参数/步骤**的 → 列在此处，建议主 Agent 在排版脚本中新增步骤
- word-typesetter **参数值不匹配**的（如模板要求 1.25 倍行距，但 word-typesetter 默认单倍）→ 列在此处，注明参数可覆盖但非默认

**此部分根据实际差异动态生成，不预设任何固定内容。**

---

## 与 word-typesetter 交互

1. 子 Agent 输出以上三部分后，等待用户确认
2. 用户确认无误后，调用 word-typesetter
3. word-typesetter 读取子 Agent 输出的参数，跳过问卷，直接生成排版脚本
4. 子 Agent 列出的"未覆盖项"由 word-typesetter 在脚本中新增对应处理步骤
5. 用户按标记指南在文档中打好标记后，排版脚本执行
