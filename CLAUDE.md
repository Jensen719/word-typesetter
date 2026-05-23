# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个学术论文 Word 排版工作区。通过自定义 Claude Code 命令 `/word-typesetter` 交互式收集排版参数，生成 Python (`win32com.client`) 脚本自动修改 .docx 文档格式。

## 核心命令

`/word-typesetter <文件路径.docx>` — 启动交互式排版流程。引导用户填写 21 类参数后，生成 17 步自动化脚本执行排版。

## 关键常量（Word COM，易错）

生成 Python 排版脚本时，以下常量与直觉值不同，已多次导致 bug：

| 含义 | 正确值 | 错误值 |
|------|--------|--------|
| A4 纸张 | `wdPaperA4 = 7` | ~~9~~ (是 A5) |
| 竖向 | `wdOrientPortrait = 0` | ~~1~~ (是横向) |
| 单倍行距 | `wdLineSpaceSingle = 0` | - |
| 固定行距 | `wdLineSpaceExactly = 4` | - |
| 多倍行距 | `wdLineSpaceMultiple = 5` | ~~4~~ (是固定值) |

完整常量表见 `.claude/commands/word-typesetter.md` 末尾。

## 文档标记规范

排版脚本依赖文档中的标记来定位操作范围。**统一使用 `《》`（全角书名号）**，不用 `<>`（半角尖括号会被 Word Find 当成通配符）。

常用标记：`《正文开始》` `《正文结束》` `《标题1》`-`《标题4》` `《摘要》` `《Abstract》` `《关键词》` `《Key Words》` `《参考文献》` `《题注》`

## 安全约束（排版脚本红线）

1. **不删内容**：`Font.Reset()` / `ParagraphFormat.Reset()` 只清格式
2. **不建表格**：只对 `doc.Tables` 已有表格改边框和样式
3. **逐段处理正文**：不跨表格边界做大 Range 操作，表格内段落用 `para.Range.Information(12)` 判读并跳过
4. **不自动保存**：排版结果留在 Word 窗口，用户手动 Ctrl+S
5. **自底向上遍历**：插入分节符/改段落时从后往前，避免索引偏移
