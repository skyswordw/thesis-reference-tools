# Zotero 接手说明

本阶段交付的是静态统一编号 Word。正文中的 `[N]` 是普通文本，不是 Zotero 动态域。

## 文件

- `output/doc/论文_统一编号.docx`: 可直接提交和人工编辑的 Word。
- `output/zotero/refs.bib`: Zotero 可导入的全文参考文献库，共 4 条唯一文献。
- `output/zotero/refs.csl.json`: CSL JSON 备用导入格式。
- `output/reports/reference_map.xlsx`: 统一编号、原章节编号与题名映射。
- `output/reports/verify_report.md`: 参考文献核查报告。
- `output/zotero/migration_checklist.md`: 静态 `[N]` 到 Zotero 条目的逐位置迁移清单。
- `output/doc/论文_zotero_dynamic_operator.docx`: 官方 Word 插件 UI 兜底操作版，正文引用为唯一 `[[ZOTERO_Pxxx]]` 占位符。
- `output/zotero/group_library_handoff.md`: Zotero Group Library 交接清单。

## 朋友电脑操作（推荐 Group Library）

1. 安装 Zotero 和 Zotero Word 插件。
2. 创建或加入共享 Group Library，把 `refs.bib` 导入该组库。
3. 安装或选择 Chinese Std GB/T 7714-2015 numeric 样式（GB/T 7714-2015 顺序编码制）。
4. 后续新增引用时，从同一个 Group Library 用 Zotero Word 插件 Add/Edit Citation 插入。
5. 如果从个人 My Library 插入，朋友端可能只能看到 orphaned items，不利于协作刷新。

## 切换到 Zotero 全托管

1. 打开 `output/zotero/migration_checklist.md`，按正文段落顺序处理每个 `Position`。
2. 在 Word 中删除该位置静态 `[N]`，用 Zotero 插件插入 `Zotero keys` 对应条目；连续引用如 `[27-29]` 在同一个 citation 里加入多篇文献。
3. 每处理 5-10 个位置保存一个 checkpoint 副本，便于 Word/Zotero 弹窗或误操作后回退。
4. 全部替换后删除静态文末参考文献，使用 Zotero Add/Edit Bibliography 生成。
5. 之后新增、删除、移动引用时，用 Zotero Refresh 统一刷新编号和文末清单。

## 自动化边界

CLI 工具负责生成迁移清单、维护 `refs.bib`/`refs.csl.json`、准备 operator 输入副本，并审计 docx 是否包含真实 Zotero 字段。最终交付优先使用官方 Word + Zotero 插件流程；直接生成 Zotero Word 字段只允许在 `build/`、`output/`、`tmp/` 实验副本中进行，且必须通过 Word 打开无修复、Zotero Refresh、重建 bibliography 和映射审计后才可交付。
