---
name: kai-docx-generator
description: >-
  Generate professional .docx documents from Markdown text or fill data into .docx templates.
  Use when: converting markdown to Word, generating reports/contracts/proposals as .docx,
  filling template variables in Word documents, validating .docx structure.
  Triggers: "生成word", "转成docx", "markdown转word", "生成文档", "填充模板",
  "fill template", "generate docx", "convert markdown to word".
version: 1.0.0
metadata:
  openclaw:
    emoji: "📝"
    os:
      - darwin
      - linux
      - windows
    requires:
      bins:
        - python3
    install:
      - id: kai-docx-generator
        kind: pip
        package: kai-docx-generator
        label: "kai-docx-generator (core)"
---

# kai-docx-generator

Generate professional .docx documents from Markdown, or fill data into .docx templates.

## Commands

| Command | What it does |
|---------|-------------|
| `kai-docx-generate <input.md> <output.docx>` | Convert Markdown to .docx |
| `kai-docx-generate --style <style> <input.md> <output.docx>` | Convert with a specific style |
| `kai-docx-fill --data <data.json|frontmatter.md> --output <out.docx> <template.docx>` | Fill template with data |
| `kai-docx-validate <file.docx>` | Validate .docx file structure |

## Markdown → .docx

```bash
kai-docx-generate input.md output.docx
kai-docx-generate --style report input.md output.docx
```

### Supported Markdown Features

| Feature | Syntax |
|---------|--------|
| Headings (H1-H3) | `#`, `##`, `###` |
| Paragraphs | Text blocks |
| Bold / Italic | `**bold**`, `*italic*` |
| Bullet / Numbered lists | `-`, `1.` |
| Tables | Standard GFM tables |
| Code blocks | Triple backticks, with optional language |
| Blockquotes (callout) | `> text` → info box |
| Images | `![caption](path)` |
| Page break | `<!-- pagebreak -->` |
| Horizontal rule | `---` |
| YAML front matter | `---\ntitle: ...\n---` |

### Style Presets

| Style | Use case |
|-------|----------|
| `standard` (default) | General documents |
| `contract` | Contracts & agreements (formal, 12pt) |
| `report` | Business reports (blue headers, 10.5pt) |
| `meeting` | Meeting minutes (compact, 10.5pt) |
| `business_letter` | Business letters (traditional format) |
| `tech_spec` | Technical specifications (monospace, dark theme) |
| `dispatch` | Official dispatchs / 红头文件 (red header) |
| `notice` | Official notices (bold, 14pt) |

### YAML Front Matter

```yaml
---
title: Document Title
author: Author Name
date: 2026-04-10
header:
  text: CONFIDENTIAL
  alignment: center
footer:
  page_number: true
toc:
  max_level: 3
---
```

## Template Filling

```bash
# With JSON data
kai-docx-fill --data data.json --output output.docx template.docx

# With Markdown front matter as data
kai-docx-fill --data frontmatter.md --output output.docx template.docx
```

Data format (JSON):
```json
{
  "company_name": "Acme Corp",
  "amount": "¥1,000,000",
  "date": "2026-04-10"
}
```

Template placeholders are `{{variable_name}}` patterns in the .docx text.

### Input Limits

- Markdown size: 200KB max
- Images: 20 max
- Table rows: 50 max
- Table columns: 20 max

Exceeding limits triggers automatic truncation with no error.

## QA Process

After generating a .docx:

1. **Open and verify** — open the output in Word/WPS/Keynote
2. **Check style** — confirm headings use the correct style preset colors and fonts
3. **Check spacing** — verify callouts have proper spacing from surrounding elements
4. **Check tables** — verify column widths are reasonable and emoji don't overlap
5. **Check code blocks** — verify monospace font and gray background
6. **Check TOC** — if requested, verify TOC field code is present (requires Word to update fields)

## Works with

- **kai-report-creator** — generate .docx reports from markdown output
- Any Markdown content (reports, specs, meeting notes, contracts)
- Existing .docx templates with `{{placeholder}}` variables
