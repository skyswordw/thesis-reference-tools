# Zotero Group Library 交接清单

本清单用于把静态统一编号版本迁移到 Zotero Word 动态引用。推荐使用 Zotero Group Library，而不是个人 My Library，这样朋友打开同一个 Word 后可以从共享组库继续 Add/Edit Citation、Refresh 和更新 bibliography。

## 输入文件

- `output/zotero/refs.bib`: 全文 4 条唯一文献，可导入 Zotero Group Library。
- `output/zotero/refs.csl.json`: 备用导入格式。
- `output/zotero/migration_checklist.md`: 正文静态引用到 Zotero 条目的逐位置迁移清单。
- `output/reports/reference_map.xlsx`: 统一编号、原章节编号与题名映射。

## 准备步骤

1. 在 Zotero 中创建或打开共享 Group Library，并邀请朋友加入。
2. 在该 Group Library 下导入 `refs.bib`，确认条目数为 4。
3. 安装并选择 `Chinese Std GB/T 7714-2015 (numeric)` 或学校认可的 GB/T 7714-2015 顺序编码制 CSL。
4. 打开 `output/doc/论文_统一编号.docx` 的副本，使用 Zotero Word 插件设置 Document Preferences 为 GB/T 7714-2015 numeric。
5. 按 `migration_checklist.md` 逐个替换正文静态 `[N]`；同一位置有多个 key 时，在同一个 Zotero citation 中加入多篇文献。
6. 每处理 5-10 个位置保存一个 checkpoint 副本。
7. 全部正文引用替换后，删除静态文末参考文献，用 Zotero Add/Edit Bibliography 生成动态 bibliography，再运行 Refresh。

## 注意

- 不要直接编辑 `.docx` XML 来伪造 Zotero 字段；字段必须由 Zotero Word 插件创建。
- 如果从个人 My Library 插入引用，朋友电脑上可能变成 orphaned items，不利于后续协作维护。
- Computer Use 只能代替点击和输入，遇到 Zotero 登录、组库权限、Word 宏权限弹窗时需要人工确认。
