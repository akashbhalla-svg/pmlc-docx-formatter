const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, convertInchesToTwip, Footer,
  PageNumber, NumberFormat, Header, TabStopPosition, TabStopType,
  TableLayoutType
} = require("docx");

// IO Brand Colors
const BRAND = {
  // Primary
  infrared: "E52321",
  dawn: "EC641D",
  black: "000000",
  white: "FFFFFF",
  // Secondary (inferred)
  ultraviolet: "6B3FA0",
  acidGreen: "A8D83A",
  electricBlue: "1A4FA0",
  voltYellow: "F5D623",
  // Neutrals
  darkGrey: "333333",
  midGrey: "666666",
  lightGrey: "F5F5F5",
  // Document accents — Electric Blue for clean, professional look
  headerAccent: "1E3A5F",
  tableHeader: "1A4FA0",
  quoteAccent: "6B3FA0",
};

// Font config (Barlow as default, fallback to Calibri)
const FONT_HEADING = "Barlow";
const FONT_BODY = "Barlow";
const FONT_FALLBACK = "Calibri";

// Parse markdown content into structured sections
function parseMarkdown(content) {
  const lines = content.split("\n");
  const elements = [];
  let inTable = false;
  let tableRows = [];
  let inCodeBlock = false;
  let codeLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Code blocks
    if (line.trim().startsWith("```")) {
      if (inCodeBlock) {
        elements.push({ type: "code", content: codeLines.join("\n") });
        codeLines = [];
        inCodeBlock = false;
      } else {
        inCodeBlock = true;
      }
      continue;
    }
    if (inCodeBlock) {
      codeLines.push(line);
      continue;
    }

    // Table rows
    if (line.trim().startsWith("|") && line.trim().endsWith("|")) {
      // Skip separator rows
      if (line.replace(/[\s|:-]/g, "").length === 0) continue;
      const cells = line.split("|").filter((c, idx, arr) => idx > 0 && idx < arr.length - 1).map(c => c.trim());
      if (!inTable) {
        inTable = true;
        tableRows = [];
      }
      tableRows.push(cells);
      continue;
    } else if (inTable) {
      elements.push({ type: "table", rows: tableRows });
      tableRows = [];
      inTable = false;
    }

    // Headings
    const h1Match = line.match(/^# (.+)/);
    const h2Match = line.match(/^## (.+)/);
    const h3Match = line.match(/^### (.+)/);
    const h4Match = line.match(/^#### (.+)/);

    if (h1Match) {
      elements.push({ type: "h1", content: h1Match[1].replace(/\*\*/g, "") });
    } else if (h2Match) {
      elements.push({ type: "h2", content: h2Match[1].replace(/\*\*/g, "") });
    } else if (h3Match) {
      elements.push({ type: "h3", content: h3Match[1].replace(/\*\*/g, "") });
    } else if (h4Match) {
      elements.push({ type: "h4", content: h4Match[1].replace(/\*\*/g, "") });
    }
    // Horizontal rule
    else if (line.trim().match(/^[-*_]{3,}$/)) {
      elements.push({ type: "hr" });
    }
    // Bullet points
    else if (line.trim().match(/^[-*] /)) {
      const content = line.trim().replace(/^[-*] /, "");
      const indent = line.match(/^(\s*)/)[1].length;
      elements.push({ type: "bullet", content, indent: Math.floor(indent / 2) });
    }
    // Numbered list
    else if (line.trim().match(/^\d+\. /)) {
      const content = line.trim().replace(/^\d+\. /, "");
      elements.push({ type: "numbered", content });
    }
    // Blockquote
    else if (line.trim().startsWith("> ")) {
      elements.push({ type: "quote", content: line.trim().replace(/^> /, "") });
    }
    // Empty line
    else if (line.trim() === "") {
      elements.push({ type: "empty" });
    }
    // Regular paragraph
    else {
      elements.push({ type: "paragraph", content: line });
    }
  }

  // Flush remaining table
  if (inTable && tableRows.length > 0) {
    elements.push({ type: "table", rows: tableRows });
  }

  return elements;
}

// Parse inline formatting (bold, italic, code, links)
function parseInlineFormatting(text) {
  const runs = [];
  // Regex to match **bold**, *italic*, `code`, and [text](url)
  const regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)|(`(.+?)`)|(\[(.+?)\]\((.+?)\))|([^*`\[]+)/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match[2]) {
      // Bold
      runs.push(new TextRun({ text: match[2], bold: true, font: FONT_BODY, size: 22 }));
    } else if (match[4]) {
      // Italic
      runs.push(new TextRun({ text: match[4], italics: true, font: FONT_BODY, size: 22, color: BRAND.midGrey }));
    } else if (match[6]) {
      // Code
      runs.push(new TextRun({ text: match[6], font: "Consolas", size: 20, shading: { type: ShadingType.SOLID, color: BRAND.lightGrey } }));
    } else if (match[8]) {
      // Link - just show text (docx links are complex)
      runs.push(new TextRun({ text: match[8], font: FONT_BODY, size: 22, color: BRAND.electricBlue, underline: {} }));
    } else if (match[10]) {
      // Plain text
      runs.push(new TextRun({ text: match[10], font: FONT_BODY, size: 22, color: BRAND.darkGrey }));
    }
  }

  if (runs.length === 0) {
    runs.push(new TextRun({ text: text, font: FONT_BODY, size: 22, color: BRAND.darkGrey }));
  }

  return runs;
}

// Build a branded table
function buildTable(rows) {
  if (rows.length === 0) return null;

  const columnCount = rows[0].length;

  const tableRows = rows.map((row, rowIndex) => {
    const isHeader = rowIndex === 0;
    const isAlternate = rowIndex % 2 === 0 && rowIndex > 0;

    const cells = row.map(cellText => {
      return new TableCell({
        children: [
          new Paragraph({
            children: parseInlineFormatting(cellText),
            spacing: { before: 60, after: 60 },
          }),
        ],
        shading: {
          type: ShadingType.SOLID,
          color: isHeader ? BRAND.tableHeader : isAlternate ? BRAND.lightGrey : BRAND.white,
        },
        margins: {
          top: convertInchesToTwip(0.05),
          bottom: convertInchesToTwip(0.05),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1),
        },
      });
    });

    // Pad if row has fewer cells than header
    while (cells.length < columnCount) {
      cells.push(new TableCell({
        children: [new Paragraph({ children: [] })],
      }));
    }

    return new TableRow({
      children: cells,
      tableHeader: isHeader,
    });
  });

  // Override header row text to white
  if (tableRows.length > 0) {
    const headerRow = rows[0];
    const headerCells = headerRow.map(cellText => {
      return new TableCell({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: cellText,
                bold: true,
                font: FONT_HEADING,
                size: 20,
                color: BRAND.white,
              }),
            ],
            spacing: { before: 60, after: 60 },
          }),
        ],
        shading: {
          type: ShadingType.SOLID,
          color: BRAND.tableHeader,
        },
        margins: {
          top: convertInchesToTwip(0.05),
          bottom: convertInchesToTwip(0.05),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1),
        },
      });
    });

    while (headerCells.length < columnCount) {
      headerCells.push(new TableCell({
        children: [new Paragraph({ children: [] })],
        shading: { type: ShadingType.SOLID, color: BRAND.tableHeader },
      }));
    }

    tableRows[0] = new TableRow({
      children: headerCells,
      tableHeader: true,
    });
  }

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
  });
}

// Convert parsed elements to docx paragraphs
function buildDocxChildren(elements) {
  const children = [];

  for (const el of elements) {
    switch (el.type) {
      case "h1":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: el.content,
                bold: true,
                font: FONT_HEADING,
                size: 40,
                color: BRAND.headerAccent,
              }),
            ],
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 3, color: BRAND.headerAccent },
            },
          })
        );
        break;

      case "h2":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: el.content,
                bold: true,
                font: FONT_HEADING,
                size: 32,
                color: BRAND.darkGrey,
              }),
            ],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 360, after: 160 },
          })
        );
        break;

      case "h3":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: el.content,
                bold: true,
                font: FONT_HEADING,
                size: 26,
                color: BRAND.headerAccent,
              }),
            ],
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 280, after: 120 },
          })
        );
        break;

      case "h4":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: el.content,
                bold: true,
                font: FONT_HEADING,
                size: 24,
                color: BRAND.midGrey,
              }),
            ],
            heading: HeadingLevel.HEADING_4,
            spacing: { before: 240, after: 100 },
          })
        );
        break;

      case "paragraph":
        children.push(
          new Paragraph({
            children: parseInlineFormatting(el.content),
            spacing: { before: 80, after: 80, line: 320 },
          })
        );
        break;

      case "bullet":
        children.push(
          new Paragraph({
            children: parseInlineFormatting(el.content),
            bullet: { level: el.indent || 0 },
            spacing: { before: 40, after: 40, line: 320 },
          })
        );
        break;

      case "numbered":
        children.push(
          new Paragraph({
            children: parseInlineFormatting(el.content),
            numbering: { reference: "default-numbering", level: 0 },
            spacing: { before: 40, after: 40, line: 320 },
          })
        );
        break;

      case "quote":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: "  " + el.content,
                italics: true,
                font: FONT_BODY,
                size: 22,
                color: BRAND.midGrey,
              }),
            ],
            indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0) },
            border: {
              left: { style: BorderStyle.SINGLE, size: 6, color: BRAND.quoteAccent, space: 12 },
            },
            spacing: { before: 120, after: 120, line: 320 },
          })
        );
        break;

      case "code":
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: el.content,
                font: "Consolas",
                size: 18,
                color: BRAND.darkGrey,
              }),
            ],
            shading: { type: ShadingType.SOLID, color: BRAND.lightGrey },
            spacing: { before: 120, after: 120 },
            indent: { left: convertInchesToTwip(0.25), right: convertInchesToTwip(0.25) },
          })
        );
        break;

      case "hr":
        children.push(
          new Paragraph({
            children: [],
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 1, color: BRAND.midGrey },
            },
            spacing: { before: 200, after: 200 },
          })
        );
        break;

      case "table":
        const table = buildTable(el.rows);
        if (table) {
          children.push(table);
          children.push(new Paragraph({ children: [], spacing: { before: 120, after: 120 } }));
        }
        break;

      case "empty":
        children.push(new Paragraph({ children: [], spacing: { before: 40, after: 40 } }));
        break;
    }
  }

  return children;
}

exports.handler = async (event) => {
  // CORS headers
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const payload = JSON.parse(event.body);
    const {
      content,
      title = "PMLC Document",
      subtitle = "",
      artifact_type = "",
      project_name = "",
      product = "",
      generated_at = new Date().toISOString(),
      stage = "",
    } = payload;

    if (!content) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: "Missing 'content' field" }),
      };
    }

    // Parse the markdown content
    const elements = parseMarkdown(content);
    const docChildren = buildDocxChildren(elements);

    // Build metadata line
    const metaText = [
      product && `Product: ${product}`,
      project_name && `Project: ${project_name}`,
      stage && `Stage: ${stage}`,
      `Generated: ${new Date(generated_at).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}`,
    ].filter(Boolean).join("  |  ");

    // Build the document
    const doc = new Document({
      numbering: {
        config: [
          {
            reference: "default-numbering",
            levels: [
              {
                level: 0,
                format: NumberFormat.DECIMAL,
                text: "%1.",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
                  },
                },
              },
            ],
          },
        ],
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: convertInchesToTwip(1),
                bottom: convertInchesToTwip(1),
                left: convertInchesToTwip(1.15),
                right: convertInchesToTwip(1.15),
              },
            },
          },
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "IOG  |  ",
                      font: FONT_HEADING,
                      size: 16,
                      color: BRAND.headerAccent,
                      bold: true,
                    }),
                    new TextRun({
                      text: artifact_type || title,
                      font: FONT_HEADING,
                      size: 16,
                      color: BRAND.midGrey,
                    }),
                  ],
                  alignment: AlignmentType.LEFT,
                }),
              ],
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "Confidential  |  ",
                      font: FONT_BODY,
                      size: 14,
                      color: BRAND.midGrey,
                    }),
                    new TextRun({
                      children: [PageNumber.CURRENT],
                      font: FONT_BODY,
                      size: 14,
                      color: BRAND.midGrey,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
          },
          children: [
            // Title block
            new Paragraph({
              children: [
                new TextRun({
                  text: title,
                  bold: true,
                  font: FONT_HEADING,
                  size: 48,
                  color: BRAND.headerAccent,
                }),
              ],
              spacing: { before: 200, after: 80 },
            }),
            // Subtitle
            ...(subtitle
              ? [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: subtitle,
                        font: FONT_HEADING,
                        size: 28,
                        color: BRAND.midGrey,
                      }),
                    ],
                    spacing: { after: 120 },
                  }),
                ]
              : []),
            // Metadata line
            new Paragraph({
              children: [
                new TextRun({
                  text: metaText,
                  font: FONT_BODY,
                  size: 18,
                  color: BRAND.midGrey,
                  italics: true,
                }),
              ],
              spacing: { after: 80 },
            }),
            // Divider
            new Paragraph({
              children: [],
              border: {
                bottom: { style: BorderStyle.SINGLE, size: 3, color: BRAND.headerAccent },
              },
              spacing: { after: 400 },
            }),
            // Document content
            ...docChildren,
          ],
        },
      ],
    });

    // Generate the docx buffer
    const buffer = await Packer.toBuffer(doc);

    return {
      statusCode: 200,
      headers: {
        ...headers,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${title.replace(/[^a-zA-Z0-9 ]/g, "")}.docx"`,
      },
      body: buffer.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (error) {
    console.error("Error generating document:", error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: error.message }),
    };
  }
};
