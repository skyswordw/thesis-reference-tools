# Thesis Reference Tools

面向**中国学位论文 Word 参考文献整理 + Zotero 动态引用交付**的工具包。

它的正式入口是**单 DOCX**：你给一个完整论文 Word 文档，工具从同一份文档里识别正文引用、章节末局部参考文献块和文末“参考文献”区，生成统一编号 Word、Zotero 导入库、迁移清单与审计报告。后续需要能协作刷新、插入或删除参考文献时，再接入 Zotero Group Library 和 Word 插件完成动态引用交付。

## 这个项目解决什么

中国学位论文里常见的问题是：不同章节或小节各自从 `[1]` 开始编号，文末参考文献区为空或不完整，后期又需要按 GB/T 7714-2015 顺序编码制统一排序。这个项目把这件事拆成两层：

| 层级 | 目标 | 主要产物 |
| --- | --- | --- |
| 静态基线 | 从单 DOCX 抽取、去重、按首次出现顺序重编号，并生成可检查的 Word 与 Zotero 导入文件 | 统一编号 Word、BibTeX、CSL JSON、reference map、核验报告 |
| Zotero 全托管 | 把静态编号迁移成 Word 里的 Zotero 动态引用，便于朋友或导师后续刷新、插入、删除和重新排序 | Zotero Group Library、项目 CSL、动态字段候选 Word、Refresh 后审计报告 |

CLI 负责“准备、生成、审计”；Zotero 负责“动态引用、协作同步、Word 插件刷新”。如果目标是可长期维护的 Word 论文，Zotero 不是装饰性增强，而是最终交付流程的一层。

## 快速开始

准备环境：

```bash
uv sync --dev
uv run pytest -q
```

运行合成 demo：

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
```

默认 demo 输出会写入 `examples/demo/output/`。正式文档可以显式指定输入和输出目录：

```bash
uv run python scripts/run_all.py --input /path/to/thesis.docx --output-root output
```

也可以用环境变量指定输入：

```bash
THESIS_REFS_INPUT_DOCX=/path/to/thesis.docx uv run python scripts/run_all.py
```

## 静态基线层

静态基线层适合先把混乱参考文献整理成可交付、可人工检查的版本。它会：

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

静态 Word 里的引用编号仍是普通 Word 文本。它适合作为稳定基线，但不等于 Zotero 动态字段。

## Zotero 全托管层

当论文需要继续增删文献、和朋友协作、或让 Word 自动刷新编号时，进入 Zotero 全托管层。

推荐流程：

1. 在 Zotero 创建或选择一个 **Zotero Group Library**，用于协作共享引用库。
2. 导入静态层生成的 `output/zotero/refs.bib`，确认 citekey 与 `ref001`、`ref002` 这类统一编号键对应。
3. 安装项目 CSL：

   ```bash
   uv run python scripts/install_project_csl.py
   ```

   在 Word Zotero Document Preferences 里选择 `NEU Thesis GB/T 7714-2015 Numeric`，它贴近学校样例并使用 GB/T 7714-2015 顺序编码制。

4. 生成 operator 占位符版，用来确定每个正文引用位置应该连接哪些 Zotero 条目：

   ```bash
   uv run python scripts/prepare_zotero_operator_docx.py --input /path/to/thesis.docx --output-root output
   ```

5. 在 Zotero 本地服务可用、并设置 `ZOTERO_GROUP_ID` 后，可以生成动态字段候选副本：

   ```bash
   ZOTERO_GROUP_ID=<group-id> uv run python scripts/generate_zotero_field_docx.py
   ```

6. 用 Microsoft Word 打开候选 DOCX，通过 Zotero Word 插件运行 `Refresh`，再用 Add/Edit Bibliography 重建文末参考文献。
7. 保存后审计动态字段：

   ```bash
   uv run python scripts/audit_zotero_docx.py output/doc/论文_zotero_dynamic.docx --checklist output/zotero/operator_migration_checklist.json
   ```

动态交付版必须同时满足：Word 打开无修复提示、Zotero `Refresh` 能完成、bibliography 能重建、审计报告能对齐 citation field 与引用条目数量。

## 产物目录

| 路径 | 用途 |
| --- | --- |
| `output/doc/` | 统一编号 Word、operator 占位符 Word、Zotero 动态候选 Word 和最终动态 Word |
| `output/zotero/` | `refs.bib`、`refs.csl.json`、Zotero Group Library 交接说明、迁移清单 |
| `output/reports/` | reference map、核验报告、Zotero 字段审计报告、metadata 审计报告 |
| `build/` | 中间 JSON、临时 DOCX、可重复生成的构建产物 |
| `styles/` | 项目自定义 CSL，默认使用 GB/T 7714-2015 顺序编码制 |
| `skills/` | 项目级 Codex skills，记录 Word 格式、Zotero 动态引用和 metadata 修复流程 |

demo 的同类产物位于 `examples/demo/output/`，可以作为最小可运行样例。

## 边界与安全

- 不修改 `raw/` 或任何输入源 DOCX；所有生成物写入 `build/`、`output/` 或 demo 输出目录。
- 不提交个人论文、密钥、Zotero 账号信息或本机私有配置。
- Zotero Web API key 只放项目本地 `.env.local`，不要写入全局 shell 配置或仓库文件。
- Zotero 动态字段只能在生成副本中创建；正式交付前必须经过 Word + Zotero Word 插件 `Refresh` 验证。
- bibliography 的语义格式优先通过 Zotero metadata 和 CSL 修复；后处理脚本只负责清理残留、悬挂缩进、行距等版面问题。

## 项目级 skills

- `skills/thesis-word-format/`：Word 论文格式、正文引用上标、作者引导引用改写。
- `skills/thesis-zotero-dynamic/`：静态编号到 Zotero 动态字段、Group Library 交接、字段审计。
- `skills/thesis-zotero-metadata/`：Zotero 作者、条目类型、语言、DOI/container、报告/标准类型修复。

开发或维护这类流程时，优先阅读对应 skill，再运行项目本地 `uv` 命令验证。
