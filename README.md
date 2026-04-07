# PMLC Document Formatter — Netlify Function

Serverless function that converts markdown content into IO-branded .docx files.

## Deployment

1. Create a new GitHub repo (e.g. `pmlc-docx-formatter`)
2. Push this entire folder to the repo
3. Connect the repo to Netlify (netlify.com → New site from Git)
4. Netlify auto-deploys. Function URL will be: `https://YOUR-SITE.netlify.app/.netlify/functions/format-document`

## API Usage

**POST** `/.netlify/functions/format-document`

### Request Body (JSON)

```json
{
  "content": "# Heading\n\nMarkdown content here...",
  "title": "[STAGE 4] PRD: Project Name",
  "subtitle": "Product Requirements Document",
  "artifact_type": "PRD",
  "project_name": "My Project",
  "product": "Blockfrost",
  "stage": "Stage 4 - Architecture",
  "generated_at": "2026-04-07T12:00:00Z"
}
```

### Response

Binary .docx file (base64 encoded if via API gateway, raw binary via direct call).

### n8n Integration

In the PMLC Update Document subworkflow, add an HTTP Request node after content generation:

1. **HTTP Request** to the Netlify function URL
   - Method: POST
   - Body: JSON with `content`, `title`, `artifact_type`, etc.
   - Response: Binary file

2. **Google Drive Upload** node
   - Upload the binary response to the PMLC Drive folder
   - Set MIME type to `application/vnd.openxmlformats-officedocument.wordprocessingml.document`

## Supported Markdown

- Headings (H1-H4)
- Bold, italic, inline code
- Bullet lists (nested)
- Numbered lists
- Tables (with branded header row)
- Blockquotes (Dawn accent border)
- Code blocks
- Horizontal rules
- Links (displayed as text)

## IO Brand Compliance

- Infrared (#E52321) for titles, H1, H3, table headers
- Dawn (#EC641D) for blockquote accents
- Barlow font throughout (Calibri fallback)
- IOG header on every page
- Page numbers in footer
- Confidential marking
