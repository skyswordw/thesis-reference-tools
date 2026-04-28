# Thesis Reference Tools

写学位论文时，参考文献很容易在章节里各自编号：第一章小节末尾有一组 `[1]...[n]`，第二章又从 `[1]` 开始，改到后期才发现需要全文统一。这个项目把这件事做得可重复、可检查，并且能接上 Zotero 做后续维护。

它面向**中国学位论文 Word 参考文献整理**场景，默认使用 **GB/T 7714-2015 顺序编码制**。正式输入是一份**单 DOCX**：工具从同一个 Word 文档里识别正文引用、章节局部参考文献块和文末“参考文献”区，再生成统一编号 Word、Zotero 导入文件、编号对照表和检查报告。

## 思路

这个项目分两步走，先把编号理顺，再让引用能长期维护。

| 步骤 | 要解决的问题 | 生成文件 |
| --- | --- | --- |
| 静态整理 | 把章节里分散编号的参考文献抽出来、去重，并按正文首次出现顺序重新编号 | 统一编号 Word、BibTeX、CSL JSON、reference map、检查报告 |
| 可维护引用 | 把整理好的编号接到 Zotero，后续增删文献时可以在 Word 里刷新 | Zotero Group Library、项目 CSL、带 Zotero 字段的 Word、字段检查结果 |

CLI 负责整理、生成和检查；Zotero 负责后续维护、协作修改和 Word 插件刷新。也就是说，脚本先把论文里的参考文献关系梳清楚，Zotero 再负责让这些关系在 Word 里继续活着。

## 快速开始

安装依赖并运行测试：

```bash
uv sync --dev
uv run pytest -q
```

运行合成 demo：

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
```

demo 的输出会写到 `examples/demo/output/`。处理实际论文文档时，可以显式指定输入和输出目录：

```bash
uv run python scripts/run_all.py --input /path/to/thesis.docx --output-root output
```

也可以用环境变量指定输入：

```bash
THESIS_REFS_INPUT_DOCX=/path/to/thesis.docx uv run python scripts/run_all.py
```

## 静态整理：先把编号理顺

静态整理适合先得到一份能打开、能核对、编号统一的 Word。它会：

- 在一个 DOCX 内识别正文 `[n]` 引用和章节局部参考文献块。
- 按正文首次出现顺序生成全文统一编号。
- 跨章节去重，合并重复条目。
- 删除章节末局部参考列表，把统一清单写入文末“参考文献”区。
- 导出 Zotero 可导入的 `refs.bib` 和 `refs.csl.json`。
- 生成 reference map、verification report 和迁移清单，方便核对每个旧编号对应的新编号。

常用命令：

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
uv run python scripts/build_docx.py --input examples/demo/raw/thesis.docx --output-root examples/demo/output
uv run python scripts/generate_zotero_migration_checklist.py --input examples/demo/raw/thesis.docx --output-root examples/demo/output
```

这一步生成的 Word 里，正文引用编号仍是普通文本。它适合作为整理后的稳定版本，但还不能随着 Zotero 自动刷新。

## Zotero 可维护引用：再让 Word 引用可以刷新

如果论文后面还会继续增删文献，或者需要和同门、导师、朋友一起改，建议把整理后的参考文献放进 **Zotero Group Library**，再通过 Word 插件生成可刷新的引用字段。

推荐流程：

1. 在 Zotero 创建或选择一个 **Zotero Group Library**，用于共享同一套引用库。
2. 导入静态整理生成的 `output/zotero/refs.bib`，确认 citekey 与 `ref001`、`ref002` 这类统一编号键对应。
3. 安装项目 CSL：

   ```bash
   uv run python scripts/install_project_csl.py
   ```

   在 Word Zotero Document Preferences 里选择 `NEU Thesis GB/T 7714-2015 Numeric`，它更接近学校样例，并使用 GB/T 7714-2015 顺序编码制。

4. 生成 operator 占位符版，用来确认每个正文引用位置应该连接哪些 Zotero 条目：

   ```bash
   uv run python scripts/prepare_zotero_operator_docx.py --input /path/to/thesis.docx --output-root output
   ```

5. 在 Zotero 本地服务可用、并设置 `ZOTERO_GROUP_ID` 后，可以生成带 Zotero 字段的实验副本：

   ```bash
   ZOTERO_GROUP_ID=<group-id> uv run python scripts/generate_zotero_field_docx.py
   ```

6. 用 Microsoft Word 打开生成的 DOCX，通过 Zotero Word 插件运行 `Refresh`，再用 Add/Edit Bibliography 重建文末参考文献。
7. 保存后检查 Zotero 字段：

   ```bash
   uv run python scripts/audit_zotero_docx.py output/doc/论文_zotero_dynamic.docx --checklist output/zotero/operator_migration_checklist.json
   ```

可维护引用版本至少要满足：Word 打开没有修复提示、Zotero `Refresh` 能完成、bibliography 能重建、字段检查结果能对齐正文引用和 Zotero 条目。

## 生成文件

| 路径 | 用途 |
| --- | --- |
| `output/doc/` | 统一编号 Word、operator 占位符 Word、带 Zotero 字段的 Word |
| `output/zotero/` | `refs.bib`、`refs.csl.json`、Zotero Group Library 说明、迁移清单 |
| `output/reports/` | reference map、核验结果、Zotero 字段检查、metadata 检查 |
| `build/` | 中间 JSON、临时 DOCX、可重复生成的构建文件 |
| `styles/` | 项目自定义 CSL，默认使用 GB/T 7714-2015 顺序编码制 |
| `skills/` | 项目级 Codex skills，记录 Word 格式、Zotero 引用和 metadata 修复流程 |

demo 的同类文件位于 `examples/demo/output/`，可以作为最小可运行样例。

## 使用约定

- 不修改 `raw/` 或任何输入源 DOCX；所有生成文件写入 `build/`、`output/` 或 demo 输出目录。
- 不提交个人论文、密钥、Zotero 账号信息或本机私有配置。
- Zotero Web API key 只放项目本地 `.env.local`，不要写入全局 shell 配置或仓库文件。
- Zotero 字段只能在生成副本中创建；用于实际论文前，必须经过 Word + Zotero Word 插件 `Refresh` 验证。
- 参考文献的语义格式优先通过 Zotero metadata 和 CSL 修复；后处理脚本只负责清理残留、悬挂缩进、行距等版面问题。

## 项目级 skills

- `skills/thesis-word-format/`：Word 论文格式、正文引用上标、作者引导引用改写。
- `skills/thesis-zotero-dynamic/`：静态编号到 Zotero 字段、Group Library 协作、字段检查。
- `skills/thesis-zotero-metadata/`：Zotero 作者、条目类型、语言、DOI/container、报告/标准类型修复。

维护这类流程时，优先阅读对应 skill，再运行项目本地 `uv` 命令验证。
