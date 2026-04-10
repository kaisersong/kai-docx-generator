---
name: kai-docx-generator
description: >-
  Generate professional .docx documents from Markdown text or fill data into .docx templates.
  Use when: converting markdown to Word, generating reports/contracts/proposals as .docx,
  filling template variables in Word documents, validating .docx structure.
  Triggers: "生成word", "转成docx", "markdown转word", "生成文档", "填充模板",
  "fill template", "generate docx", "convert markdown to word".
version: 1.0.2
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
      - id: python-docx
        kind: uv
        package: python-docx
        label: "python-docx (core)"
      - id: markdown-it-py
        kind: uv
        package: markdown-it-py
        label: "markdown-it-py (parser)"
      - id: lxml
        kind: uv
        package: lxml
        label: "lxml (XML)"
      - id: Pillow
        kind: uv
        package: Pillow
        label: "Pillow (images)"
      - id: pyyaml
        kind: uv
        package: pyyaml
        label: "pyyaml (YAML)"
---

# kai-docx-generator

Generate professional .docx documents from Markdown, or fill data into .docx templates.

## Commands

| Command | What it does |
|---------|-------------|
| `python3 <skill-path>/scripts/generate.py <input.md> <output.docx>` | Convert Markdown to .docx |
| `python3 <skill-path>/scripts/generate.py --style <style> <input.md> <output.docx>` | Convert with a specific style |
| `python3 <skill-path>/scripts/fill_template.py --data <data.json|frontmatter.md> --output <out.docx> <template.docx>` | Fill template with data |
| `python3 <skill-path>/scripts/validate.py <file.docx>` | Validate .docx file structure |

## Markdown → .docx

```bash
python3 <skill-path>/scripts/generate.py input.md output.docx
python3 <skill-path>/scripts/generate.py --style report input.md output.docx
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
python3 <skill-path>/scripts/fill_template.py --data data.json --output output.docx template.docx

# With Markdown front matter as data
python3 <skill-path>/scripts/fill_template.py --data frontmatter.md --output output.docx template.docx
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
