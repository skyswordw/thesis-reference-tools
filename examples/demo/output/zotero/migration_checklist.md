# Zotero 动态域迁移清单

- Target DOCX: `examples/demo/output/doc/论文_统一编号_demo.docx`
- Unique references: 4
- Citation positions: 4
- Citation paragraphs: 3
- Covered citation occurrences: 7 / 7
- Paragraph coverage OK: True

## 操作规则

- 在 Word 中删除 `Marker` 对应的静态文本，再用 Zotero Add/Edit Citation 插入 `Zotero keys` 中的条目。
- 同一行包含多个 key 时，在同一个 Zotero citation 中连续添加多篇文献。
- 每处理 5-10 个位置保存一个 checkpoint 副本。

| Position | Paragraph | Section | Marker | Zotero keys | Titles |
|---|---:|---|---|---|---|
| P001 | 4 | 1.1 研究背景 | `[1-2]` | `ref001, ref002` | Alpha sensing route; Q.4144: Signalling requirements for cross-operator service orchestration in computing power networks |
| P002 | 6 | 第二章 相关工作 | `[3-4]` | `ref003, ref004` | Beta routing; Gamma compute |
| P003 | 6 | 第二章 相关工作 | `[2]` | `ref002` | Q.4144: Signalling requirements for cross-operator service orchestration in computing power networks |
| P004 | 7 | 第二章 相关工作 | `[4][2]` | `ref004, ref002` | Gamma compute; Q.4144: Signalling requirements for cross-operator service orchestration in computing power networks |
