---
name: thesis-zotero-metadata
description: Use when auditing or enriching Zotero Group Library metadata for a thesis reference project, preparing metadata patches, handling DOI/identifier gaps, or coordinating metadata cleanup before rebuilding Zotero dynamic thesis DOCX files.
---

# Thesis Zotero Metadata

## Core Rule

Use a verification-first metadata patch route. Audit the existing Zotero Group Library, enrich missing or weak records through public scholarly metadata APIs, produce a reviewable patch, then apply changes to the existing group library before rebuilding and refreshing the dynamic Word deliverable.

## Boundaries

- Do not edit `raw/`; source Word documents are read-only.
- Use the project-local `uv` environment for all Python commands.
- Keep generated audits, patches, reports, and checkpoint DOCX files under `build/`, `output/`, or `tmp/`.
- Store any Zotero Web API key only in project-local `.env.local`; never put it in global shell config, source files, committed docs, or generated reports.
- Do not write to Zotero SQLite. Use the Zotero Web API or exported/generated project artifacts for updates.
- Do not pursue scan-conversion experiments as a project route; they are not part of the accepted metadata or dynamic-citation workflow.

## Metadata Completion Route

1. Audit the configured Zotero Group Library, identified by a project-local setting such as `ZOTERO_GROUP_ID` in `.env.local`: compare item titles, authors, years, Zotero `itemType`, `language`, DOI/ISBN/URL fields, citekeys, and attachment/link state against the project reference artifacts.
2. Enrich incomplete records through Crossref, OpenAlex, and Semantic Scholar. Prefer exact title/DOI matches, record the source used for each field, and leave uncertain matches flagged for manual review.
3. Always patch non-DOI project references too. Reports, white papers, and research reports must be Zotero `report`; ITU-T recommendations and technical standards must be Zotero `standard`; journal and conference papers remain `journalArticle` and `conferencePaper`.
4. Set item language before Word Refresh: English scholarly items use `language=en-US` so multi-author bibliography output uses `et al.`; Chinese items use `language=zh-CN` so Chinese author lists may still use `等`.
5. Generate a reviewable metadata patch before applying anything. The patch should make additions or corrections explicit per Zotero item and keep rejected/uncertain candidates separate.
6. Apply the approved patch to the configured Zotero group through the Zotero Web API using the project-local `.env.local` API key.
7. Rebuild the dynamic DOCX using the project Zotero dynamic-citation workflow.
8. Open the rebuilt document in Microsoft Word, run Zotero Refresh, regenerate or refresh the bibliography, and save a checkpoint copy.
9. Audit the refreshed DOCX against the reference map and migration checklist. Treat the result as handoff-ready only when citation counts, item mappings, bibliography field presence, `itemType`/`language` corrections, and bibliography semantic checks all pass.

## Evidence To Keep

- Group Library audit report before changes.
- Enrichment candidate report with source API and match confidence.
- Approved metadata patch and application log.
- Dynamic DOCX rebuild checkpoint.
- Word Zotero Refresh result and final audit report.

## Coordination

Use `skills/thesis-zotero-dynamic/SKILL.md` for dynamic DOCX generation, Word Refresh validation, Group Library handoff, and citation-field auditing. This skill covers the metadata cleanup pass that should happen before or alongside that workflow.
