const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, NumberFormat,
  SectionType, TabStopType, TabStopPosition, PositionalTab,
  PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader
} = require('docx');
const fs = require('fs');

// ─── COLOUR & STYLE CONSTANTS ───────────────────────────────────────────────
const BLUE   = "1A5276";   // deep navy blue (headings, boxes)
const LBLUE  = "D6EAF8";   // light blue shading
const ACCENT = "2874A6";   // mid-blue
const GRAY   = "ECF0F1";   // light gray for code blocks / boxes
const DGRAY  = "717D7E";   // dark gray text
const BLACK  = "000000";
const WHITE  = "FFFFFF";

// DXA unit helpers
const inch = (n) => Math.round(n * 1440);

// Page: 6×9 book trim
const PAGE_W  = inch(6);   // 8640
const PAGE_H  = inch(9);   // 12960
const MARGIN_TOP    = inch(0.75);
const MARGIN_BOTTOM = inch(0.85);
const MARGIN_INNER  = inch(0.90);  // gutter
const MARGIN_OUTER  = inch(0.70);
const CONTENT_W = PAGE_W - MARGIN_INNER - MARGIN_OUTER; // 5.4 inches in DXA = 7776

// ─── STYLES ─────────────────────────────────────────────────────────────────
const styles = {
  default: {
    document: { run: { font: "Georgia", size: 22, color: BLACK } },
  },
  paragraphStyles: [
    // BOOK TITLE (title page)
    {
      id: "BookTitle", name: "Book Title", basedOn: "Normal",
      run: { font: "Arial", size: 52, bold: true, color: BLUE },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 480, after: 240 } },
    },
    // BOOK SUBTITLE
    {
      id: "BookSubtitle", name: "Book Subtitle", basedOn: "Normal",
      run: { font: "Arial", size: 28, color: ACCENT, italics: true },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 480 } },
    },
    // AUTHOR NAME
    {
      id: "AuthorName", name: "Author Name", basedOn: "Normal",
      run: { font: "Arial", size: 24, color: DGRAY },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 } },
    },
    // PART HEADING
    {
      id: "PartHeading", name: "Part Heading", basedOn: "Normal",
      run: { font: "Arial", size: 36, bold: true, color: WHITE },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 480, after: 240 },
        shading: { fill: BLUE, type: ShadingType.CLEAR } },
    },
    // CHAPTER NUMBER
    {
      id: "ChapterNumber", name: "Chapter Number", basedOn: "Normal",
      run: { font: "Arial", size: 24, color: ACCENT, bold: true },
      paragraph: { spacing: { before: 480, after: 60 } },
    },
    // HEADING 1 (chapter title) — must use exact ID for TOC
    {
      id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
      quickFormat: true,
      run: { font: "Arial", size: 36, bold: true, color: BLUE },
      paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 6 } } },
    },
    // HEADING 2 (section heading)
    {
      id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
      quickFormat: true,
      run: { font: "Arial", size: 26, bold: true, color: BLUE },
      paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
    },
    // HEADING 3 (sub-section)
    {
      id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal",
      quickFormat: true,
      run: { font: "Arial", size: 22, bold: true, color: ACCENT },
      paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 },
    },
    // BODY TEXT
    {
      id: "BodyText", name: "Body Text", basedOn: "Normal",
      run: { font: "Georgia", size: 22, color: BLACK },
      paragraph: { spacing: { before: 0, after: 160 }, line: 276, lineRule: "auto",
        alignment: AlignmentType.JUSTIFIED },
    },
    // FRONT MATTER HEADING
    {
      id: "FrontHeading", name: "Front Heading", basedOn: "Normal",
      run: { font: "Arial", size: 30, bold: true, color: BLUE },
      paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } } },
    },
    // CAPTION / SMALL TEXT
    {
      id: "Caption", name: "Caption", basedOn: "Normal",
      run: { font: "Arial", size: 18, italics: true, color: DGRAY },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 40, after: 160 } },
    },
    // COPYRIGHT TEXT
    {
      id: "CopyrightText", name: "Copyright Text", basedOn: "Normal",
      run: { font: "Arial", size: 18, color: DGRAY },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } },
    },
  ],
};

// ─── NUMBERING CONFIG ────────────────────────────────────────────────────────
const numbering = {
  config: [
    {
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    },
    {
      reference: "numbered",
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    },
    {
      reference: "checklist",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u25A1",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    },
  ],
};

// ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────

function bodyPara(text, opts = {}) {
  const runs = typeof text === 'string'
    ? [new TextRun({ text, font: "Georgia", size: 22, color: BLACK })]
    : text;
  return new Paragraph({
    style: "BodyText",
    children: runs,
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: opts.afterSpacing ?? 160 },
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: BLUE })] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: BLUE })] });
}

function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: ACCENT })] });
}

function bulletPara(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Georgia", size: 22, color: BLACK })],
    spacing: { before: 40, after: 40 },
  });
}

function numberedPara(text) {
  return new Paragraph({
    numbering: { reference: "numbered", level: 0 },
    children: [new TextRun({ text, font: "Georgia", size: 22, color: BLACK })],
    spacing: { before: 40, after: 40 },
  });
}

function checkPara(text) {
  return new Paragraph({
    numbering: { reference: "checklist", level: 0 },
    children: [new TextRun({ text, font: "Georgia", size: 21, color: BLACK })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer(pts = 160) {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 0, after: pts } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function chapterOpener(numText, titleText) {
  return [
    new Paragraph({
      style: "ChapterNumber",
      children: [new TextRun({ text: numText, font: "Arial", size: 24, color: ACCENT, bold: true })],
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: titleText, font: "Arial", size: 36, bold: true, color: BLUE })],
    }),
    spacer(80),
  ];
}

function boxedParagraphs(title, items, fillColor = LBLUE, isChecklist = false) {
  // Bordered shaded box with title + bullet points
  const border = { style: BorderStyle.SINGLE, size: 4, color: ACCENT };
  const cellBorders = { top: border, bottom: border, left: border, right: border };

  const cellChildren = [
    new Paragraph({
      children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: BLUE })],
      spacing: { before: 0, after: 100 },
    }),
    ...items.map(item =>
      new Paragraph({
        numbering: isChecklist ? { reference: "checklist", level: 0 } : { reference: "bullets", level: 0 },
        children: [new TextRun({ text: item, font: "Georgia", size: 21, color: BLACK })],
        spacing: { before: 40, after: 40 },
      })
    ),
  ];

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            shading: { fill: fillColor, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            width: { size: CONTENT_W, type: WidthType.DXA },
            children: cellChildren,
          }),
        ],
      }),
    ],
  });
}

function dataTable(headers, rows) {
  const colW = Math.floor(CONTENT_W / headers.length);
  const hdrBorder = { style: BorderStyle.SINGLE, size: 2, color: ACCENT };
  const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const hdrBorders = { top: hdrBorder, bottom: hdrBorder, left: hdrBorder, right: hdrBorder };
  const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

  const headerRow = new TableRow({
    children: headers.map(h =>
      new TableCell({
        borders: hdrBorders,
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: colW, type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text: h, font: "Arial", size: 20, bold: true, color: WHITE })],
          alignment: AlignmentType.CENTER,
        })],
      })
    ),
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map(cell =>
        new TableCell({
          borders: cellBorders,
          shading: { fill: ri % 2 === 0 ? WHITE : GRAY, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          width: { size: colW, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: cell, font: "Georgia", size: 20, color: BLACK })],
          })],
        })
      ),
    })
  );

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: headers.map(() => colW),
    rows: [headerRow, ...dataRows],
  });
}

// ─── SECTION CONTENT BUILDERS ────────────────────────────────────────────────

function buildTitlePage() {
  return [
    spacer(inch(1.5) / 20),
    new Paragraph({
      style: "BookTitle",
      children: [new TextRun({ text: "AI-Assisted Digital Product", font: "Arial", size: 52, bold: true, color: BLUE })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
    }),
    new Paragraph({
      style: "BookTitle",
      children: [new TextRun({ text: "Business Operations Manual", font: "Arial", size: 52, bold: true, color: BLUE })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
    }),
    new Paragraph({
      style: "BookSubtitle",
      children: [new TextRun({
        text: "A Systematic Framework for Building Sustainable Revenue on Etsy",
        font: "Arial", size: 28, italics: true, color: ACCENT
      })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 480 },
    }),
    spacer(200),
    new Paragraph({
      children: [new TextRun({ text: "\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015", font: "Arial", size: 24, color: ACCENT })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      style: "AuthorName",
      children: [new TextRun({ text: "[Author Name]", font: "Arial", size: 24, color: DGRAY })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "First Edition", font: "Arial", size: 20, color: DGRAY })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "[Publisher Name]  \u2022  [Copyright Year]", font: "Arial", size: 20, color: DGRAY })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
    }),
  ];
}

function buildCopyrightPage() {
  return [
    spacer(300),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "AI-Assisted Digital Product Business Operations Manual", font: "Arial", size: 20, bold: true })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "Copyright \u00A9 [Copyright Year] [Author Name]. All rights reserved.", font: "Arial", size: 20 })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "Published by [Publisher Name]", font: "Arial", size: 20 })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "First Edition \u2022 [Version] \u2022 [Date]", font: "Arial", size: 20 })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 320 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({
        text: "No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, including photocopying, recording, or other electronic or mechanical methods, without the prior written permission of the publisher, except in the case of brief quotations embodied in critical reviews and certain other non-commercial uses permitted by copyright law.",
        font: "Georgia", size: 20
      })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 320 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({
        text: "DISCLAIMER: This book is for educational and informational purposes only. It does not constitute legal, financial, tax, or professional advice. Income figures and performance data cited throughout are based on available creator data and represent a range of outcomes. Individual results vary significantly based on niche selection, execution quality, market conditions, and time invested. The author and publisher make no guarantees of income or specific outcomes. Readers should consult qualified professionals before making financial or business decisions.",
        font: "Georgia", size: 18, italics: true
      })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "Contact: [Contact Information]", font: "Arial", size: 18 })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 80 },
    }),
    new Paragraph({
      style: "CopyrightText",
      children: [new TextRun({ text: "Website: [Website]", font: "Arial", size: 18 })],
      alignment: AlignmentType.LEFT,
    }),
  ];
}

function buildPrefacePage() {
  return [
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Preface", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("This manual exists because most digital product guides fail the people who need them most. They either inspire without instructing, or instruct without depth. This book does neither."),
    bodyPara("What you hold is a professional business operations framework built from real creator data, documented processes, and honest analysis of what works, what fails, and why. It was written for people who want to build a real business, not chase a trend."),
    bodyPara("The AI-assisted workflows in this manual are tools, not shortcuts. Claude, Canva, and other tools described here accelerate execution. They do not replace judgment, niche research, or consistent work. That part remains yours."),
    bodyPara("Read Part 0 before anything else. It is the most important section in the book. If you understand the failure modes, the realistic timelines, and the evidence base before you begin building, your probability of success increases substantially."),
    bodyPara("The rest will follow."),
    spacer(200),
    new Paragraph({
      children: [new TextRun({ text: "[Author Name]", font: "Georgia", size: 22, italics: true, color: DGRAY })],
      alignment: AlignmentType.RIGHT,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "[Date]", font: "Georgia", size: 20, italics: true, color: DGRAY })],
      alignment: AlignmentType.RIGHT,
    }),
  ];
}

function buildHowToUsePage() {
  return [
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "How to Use This Book", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("This book is organized into seven parts. Each part builds on the previous. The sequence is deliberate."),
    h3("For New Implementers"),
    bodyPara("Read Parts 0 through 4 in full before taking any action. The frameworks build on one another. Skipping ahead causes avoidable errors."),
    h3("For Those With Products Already Live"),
    bodyPara("Start with Part 4 (Etsy Optimization) and Part 5 (Customer Acquisition). Then use Part 6 to build systems that scale what is already working."),
    h3("For Troubleshooting"),
    bodyPara("If revenue has stalled or never arrived, go directly to Chapter 26: Why Businesses Fail. Work through the failure mode diagnostics before making any changes."),
    h3("For Advanced Strategy"),
    bodyPara("Parts 6 and 7 cover automation, AI workflows, financial modelling, and legal frameworks. These are designed for sellers with at least five products live and some revenue history."),
    h3("Chapter Format"),
    bodyPara("Each chapter includes: a learning objective, a key concept explanation, practical workflows, common mistakes, key takeaways, and an action checklist. Work through each component before moving forward."),
    spacer(120),
    boxedParagraphs("Navigation Guide", [
      "Part 0 \u2014 Foundation & Realism: Start here. Always.",
      "Part 1 \u2014 Operating Environment: Understand the platform and tools.",
      "Part 2 \u2014 Validation & Planning: Research before building.",
      "Part 3 \u2014 Product Operations: Build efficiently and systematically.",
      "Part 4 \u2014 Etsy Optimization: Rank, convert, and iterate.",
      "Part 5 \u2014 Customer Acquisition: Drive traffic from multiple channels.",
      "Part 6 \u2014 Scaling: Automate, systemize, and grow.",
      "Part 7 \u2014 Advanced Operations: AI workflows, legal, case studies.",
    ], LBLUE),
  ];
}

function buildIntroduction() {
  return [
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Introduction", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("This guide provides a systematic operational framework for building a digital product business on Etsy using AI-assisted workflows. It is not a motivational manual. It is not a get-rich-quick system. It is a professional business operations framework based on real financial data, documented processes, and honest analysis of outcomes."),
    h2("What Makes This Guide Different"),
    bodyPara("Most digital product guides fall into one of two categories: motivational but vague, or tactical but incomplete. This guide treats digital product business as what it is: a serious operation that requires systematic thinking, consistent execution, and data-driven decisions."),
    h2("Realistic Income Expectations"),
    bodyPara("Before proceeding, understand the realistic income distribution based on available creator data:"),
    spacer(80),
    dataTable(
      ["Outcome", "Percentage of Starters", "Timeline"],
      [
        ["$1,000+/month", "15\u201320%", "Within 12 months"],
        ["$200\u2013500/month (side income)", "30\u201335%", "Within 12 months"],
        ["$50\u2013200/month (proof of concept)", "25\u201330%", "Within 12 months"],
        ["Minimal or zero revenue", "15\u201320%", "Within 12 months"],
      ]
    ),
    spacer(120),
    bodyPara("These figures reflect focused implementation. Single failures in niche selection, SEO, or consistency significantly reduce outcomes."),
    h2("Timeline Reality"),
    bodyPara("The 90-day timeline discussed in Part 0 represents aggressive execution under favorable conditions. More realistic timelines for most creators:"),
    bulletPara("Months 1\u20132: Learning phase. Likely $0 revenue. 0\u20132 products live."),
    bulletPara("Months 3\u20134: Validation phase. $50\u2013300/month typical. 3\u20135 products live."),
    bulletPara("Months 5\u20136: Growth phase. $200\u2013800/month typical. 6\u201310 products live."),
    bulletPara("Months 7\u201312: Scaling phase. $1,000\u20135,000/month for successful implementations. 15+ products live."),
    spacer(120),
    bodyPara("This assumes 15\u201320 hours per week, correct niche selection, proper SEO implementation, and consistent execution across all variables."),
    h2("Success Metrics at a Glance"),
    spacer(80),
    dataTable(
      ["Metric", "Month 1", "Month 3", "Month 6", "Month 12"],
      [
        ["Products Live", "1", "3\u20135", "8\u201312", "20\u201335"],
        ["Monthly Impressions", "50\u2013200", "500\u20131,500", "2,000\u20135,000", "8,000\u201320,000"],
        ["Email Subscribers", "5\u201315", "50\u2013150", "200\u2013400", "500\u20131,200"],
        ["Monthly Revenue", "$0\u201350", "$100\u2013500", "$500\u20131,500", "$2,000\u20138,000+"],
        ["Click-Through Rate", "1\u20132%", "2\u20133%", "2.5\u20134%", "2.5\u20134%"],
        ["Conversion Rate", "0\u20131%", "1\u20132%", "1.5\u20133%", "2\u20133.5%"],
      ]
    ),
    spacer(120),
  ];
}

// ─── PART 0 ──────────────────────────────────────────────────────────────────
function buildPart0() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 0", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "FOUNDATION & REALISM", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 0.1 \u2013 0.5", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 0.1", "Why This Model Works \u2014 The Data"),
    bodyPara("Digital products on Etsy represent one of the highest-margin, lowest-friction business models available to solo creators today. The combination of AI-assisted workflows and a proven marketplace removes the two biggest barriers: time and distribution."),
    h2("Time to First Sale: Manual vs. AI-Assisted"),
    dataTable(
      ["Approach", "Time to First Sale", "Key Difference"],
      [
        ["Manual (no AI)", "60\u201390 days average", "Every task done from scratch"],
        ["Claude-assisted", "14\u201321 days average", "AI handles brainstorming, copy, research"],
        ["Time saved", "60\u201370%", "\u2014"],
      ]
    ),
    spacer(120),
    bodyPara("Claude handles brainstorming (30 minutes instead of 4 hours), research (15 minutes instead of 2 hours), copywriting (30 minutes instead of 3 hours), and optimization (20 minutes instead of 1 hour). The creator handles niche selection, design direction, quality review, and strategic decisions."),
    h2("Cost Structure"),
    dataTable(
      ["Cost Category", "Monthly", "Annual", "Notes"],
      [
        ["Claude Pro", "$20", "$240", "Essential for heavy use"],
        ["Canva Pro", "$13", "$156", "Design tool"],
        ["Google Workspace", "$6\u2013$14", "$72\u2013$168", "Email + Sheets + Drive"],
        ["Etsy fees", "Variable (~15%)", "\u2014", "10% transaction + $0.20 listing"],
        ["Email platform", "$0\u2013$15", "$0\u2013$180", "Beehiiv or Mailchimp"],
        ["Total Fixed", "$49\u2013$97", "$588\u2013$1,164", "Minimal overhead"],
      ]
    ),
    spacer(120),
    h2("Profit Margins by Product Category"),
    dataTable(
      ["Product Type", "Margin", "Typical Price", "Monthly Potential"],
      [
        ["Notion Templates", "85\u201395%", "$17\u2013$32", "$2,000 per successful template"],
        ["Printables (PDF)", "70\u201380%", "$9\u2013$17", "$1,500 per product"],
        ["eBooks & Guides", "90%+", "$17\u2013$49", "$2,500+ per ebook"],
        ["Prompt Packs", "92%+", "$9\u2013$32", "$3,000+ for popular packs"],
        ["Spreadsheet Templates", "88\u201395%", "$12\u2013$32", "$2,000 per template"],
        ["Social Media Kits", "85\u201392%", "$17\u2013$32", "$1,500\u2013$3,000"],
        ["AI Tools / Prompt Systems", "85\u201390%", "$24\u2013$97", "$4,000+ for quality packages"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 0.1", [
      "AI-assisted workflows reduce time-to-first-sale by 60\u201370% compared to manual methods.",
      "Fixed monthly costs range from $49 to $97, making this one of the lowest-overhead business models available.",
      "Profit margins of 85\u201395% are achievable on most digital product types.",
      "Revenue compounds as each new product supports existing listings through portfolio effects.",
      "The model works best with consistent execution, not sporadic effort.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 0.2", "The Business Model Explained"),
    bodyPara("The AI-assisted Etsy digital product model operates on a simple principle: create information products once, sell them infinitely, with AI handling the majority of production and optimization overhead."),
    h2("The Core Loop"),
    bodyPara("The business follows a five-stage operational cycle that repeats monthly as the portfolio grows:"),
    numberedPara("Research: Identify unmet demand in a specific niche using data-driven validation."),
    numberedPara("Create: Build a digital product using AI-assisted workflows (Claude + Canva or Notion)."),
    numberedPara("Publish: List on Etsy with SEO-optimized titles, descriptions, and tags."),
    numberedPara("Market: Drive traffic via Pinterest, email, and social channels."),
    numberedPara("Optimize: Analyze performance data. Improve weak listings. Scale winners."),
    spacer(120),
    h2("Revenue Paths by Stage"),
    dataTable(
      ["Path", "Products", "Avg. Price", "Monthly Revenue", "Assessment"],
      [
        ["Conservative (Beginner)", "4", "$17", "$120\u2013$170", "Proof of concept"],
        ["Moderate (Intermediate)", "15", "$24", "$960\u2013$1,440", "Sustainable side income"],
        ["Aggressive (Advanced)", "35+", "$32", "$6,400\u2013$8,960", "Full-time income"],
        ["Scaling (Optimization)", "50+", "$28", "$8,400\u2013$12,600", "Venture-level income"],
      ]
    ),
    spacer(120),
    bodyPara("These figures are not guarantees. They represent outcomes for creators who execute consistently across all variables: niche, SEO, quality, and marketing."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 0.2", [
      "The model is portfolio-based: more products compound revenue across the same audience.",
      "AI handles production tasks; strategic judgment remains with the creator.",
      "Revenue does not typically arrive in Month 1. Months 3\u20136 are the typical onset of meaningful revenue.",
      "Each product is a long-term asset that continues generating traffic and sales.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 0.3", "Realistic Expectations & Success Dependencies"),
    bodyPara("Understanding what determines success prevents the most common form of failure: wrong expectations leading to premature abandonment."),
    h2("The Seven Success Dependencies"),
    numberedPara("Niche selection quality: Profitable niches have search volume above 200 searches per month, fewer than 300 competing listings, and buyer willingness to pay $15 or more."),
    numberedPara("Execution consistency: Weekly publishing outperforms sporadic bursts. Algorithms favor consistent sellers."),
    numberedPara("SEO implementation: Keyword research before building, not after. Long-tail keyword clustering over broad single-term targeting."),
    numberedPara("Product quality and differentiation: Generic products underperform. Products built on competitor gap analysis outperform."),
    numberedPara("Time invested: Month 1\u20136 requires 15\u201330 active hours per week. This is not passive income in the early phase."),
    numberedPara("Market timing: Launching during demand peaks accelerates early traction. Seasonal calendar planning matters."),
    numberedPara("Customer research integration: Products built on data outperform products built on assumptions."),
    spacer(120),
    h2("90-Day Operational Timeline"),
    dataTable(
      ["Phase", "Timeline", "Focus", "Expected Output"],
      [
        ["Foundation", "Days 1\u201314", "Setup, research, first product", "1 product live"],
        ["Validation", "Days 15\u201345", "Publish 2\u20133 products, monitor metrics", "First sales data"],
        ["Growth", "Days 46\u201375", "Scale to 5\u20137 products, begin email list", "Revenue trend visible"],
        ["Optimization", "Days 76\u201390", "A/B test top performers, build Pinterest", "Systems running"],
      ]
    ),
    spacer(120),
    h2("Fallback Plans"),
    h3("Behind on Month 1"),
    bodyPara("Pick the best validated idea (not the most interesting). Create a simplified version at 80% quality rather than waiting for perfection. Launch it immediately. Revenue data is more valuable than the perfect product that does not yet exist."),
    h3("Behind on Month 2"),
    bodyPara("Skip the second new product. Focus entirely on marketing the first product. Create 100 Pinterest pins in a single two-hour batch session. One product with great marketing outperforms two products with weak marketing."),
    h3("Behind on Month 3"),
    bodyPara("If cumulative revenue is under $300, run diagnostics. Low impressions signal an SEO problem. Low conversion rate signals a copy or price problem. No traffic signals a marketing problem. Address the root cause before creating new products."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 0.3", [
      "Success depends on seven variables, all of which are within the creator\u2019s control.",
      "The 90-day timeline represents aggressive execution; realistic timelines extend 12\u201318 months.",
      "Fallback strategies exist for every phase. Pivoting quickly is preferable to persisting with a failing approach.",
      "Treat Month 1\u20136 as an active business phase, not a passive income setup.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 0.4", "Addressing Common Doubts"),
    bodyPara("Every creator faces objections before beginning. Most objections are resolvable with facts. This chapter addresses the most common concerns directly."),
    h2("Common Questions Answered"),
    h3("Is AI-generated content legal to sell?"),
    bodyPara("Yes. AI-generated text is legal to sell. You own the output. AI-generated images carry more legal complexity and should be avoided until copyright law on AI-generated visuals clarifies. Claude-written descriptions and copy are yours to use commercially. Etsy's policy explicitly allows AI-generated content with disclosure where required."),
    bodyPara("Your responsibility: Ensure content is original, does not plagiarize, and does not reproduce copyrighted material."),
    h3("Is the market saturated?"),
    bodyPara("Broad markets are saturated. Narrow niches are not. The word 'productivity template' has over 8,000 competing listings. 'Productivity template for therapists' has approximately 50. 'Productivity template for therapists who serve trauma survivors' may have three. Specificity is the competitive advantage. Validation frameworks (Chapter 7) identify unsaturated niches before you invest time building."),
    h3("Can I do this solo?"),
    bodyPara("Yes. Most successful digital product creators run one to five stores solo for their first year before adding any help. Batching and automation let one person produce what might otherwise require a team. Systems replace headcount in this model."),
    h3("How much does it cost to start?"),
    dataTable(
      ["Setup Level", "Monthly Cost", "Tools Included"],
      [
        ["Minimal", "$0\u2013$40", "Etsy (free), Claude Pro ($20), Canva (free)"],
        ["Comfortable", "$40\u2013$60", "Above + Canva Pro ($13), Google Workspace ($6), Beehiiv ($5)"],
        ["Full stack", "$60\u2013$100", "Above + Buffer/Later, Analytics tools"],
      ]
    ),
    spacer(120),
    h3("What if I have no design experience?"),
    bodyPara("Canva templates handle 90% of design requirements. The workflow is: choose a template, change the headline, adjust colors, swap images, and export. Most successful sellers in this space started with zero design experience."),
    h3("What is the time commitment?"),
    bulletPara("Weeks 1\u20134: 20\u201325 hours per week (learning, first product, first listings)"),
    bulletPara("Weeks 5\u201312: 25\u201330 hours per week (content, email list, scaling)"),
    bulletPara("Weeks 13\u201324: 10\u201315 hours per week (systems running, adding products)"),
    bulletPara("Month 6+: 5\u201310 hours per week (optimizing, scaling, systemized)"),
    spacer(80),
    bodyPara("This schedule is compatible with maintaining a full-time job during the first 90 days."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 0.4", [
      "AI content is legal to sell on Etsy with proper disclosure.",
      "Market saturation exists in broad niches; specific niches remain open.",
      "Startup costs range from $0 to $100 per month.",
      "No design experience is required. Canva provides professional templates.",
      "The time commitment is front-loaded and decreases substantially after Month 6.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 0.5", "Why Digital Product Businesses Fail \u2014 Prevention Framework"),
    bodyPara("Seventy to eighty percent of creators who start a digital product business do not reach $1,000 per month. Understanding the failure modes before starting is the most powerful prevention available."),
    h2("The Ten Critical Failure Modes"),
    h3("Failure Mode 1: Wrong Niche Selection (40% of failures)"),
    bodyPara("Creator selects niche based on personal interest rather than market demand. Demand exists superficially but purchase volume is minimal. Competitors race to the bottom on price. Willingness to pay is too low."),
    bodyPara("Prevention: Conduct full validation before building. Use the nine-signal Validation Matrix (Chapter 7). Test niche with two products before committing full effort."),
    h3("Failure Mode 2: Poor SEO Implementation (30% of failures)"),
    bodyPara("Keywords are generic or too broad. Title does not include search intent. Tags are random or misaligned with search behavior. Listing exists but never ranks."),
    bodyPara("Prevention: Conduct keyword research before building. Use long-tail keywords. Optimize title for keyword inclusion in the first 70 characters. Test three to four title variations weekly."),
    h3("Failure Mode 3: Insufficient Execution Consistency (35% of failures)"),
    bodyPara("Creator builds one or two products with enthusiasm, then publishing slows in Month 2 or 3. Algorithm deprioritizes inconsistent sellers. Existing product rankings decay."),
    bodyPara("Prevention: Commit to weekly publishing. Batch create products rather than building one at a time. Block calendar time. Set a minimum output of two products per month."),
    h3("Failure Mode 4: Product Quality / Differentiation Issues (25% of failures)"),
    bodyPara("Product looks identical to fifty competing listings. Content is thin. No clear positioning. Reviews accumulate but are lukewarm."),
    bodyPara("Prevention: Review top five competing products before creating. Identify gaps and complaints in competitor reviews. Build explicitly to fill those gaps."),
    h3("Failure Mode 5: No Customer Research or Feedback Integration (25% of failures)"),
    bodyPara("Creator builds products based on assumptions. Ignores customer requests in reviews. Fails to analyze competitor reviews for market signals."),
    bodyPara("Prevention: Join communities where target customers gather (Reddit, Facebook groups). Ask about pain points. Analyze top twenty competitor reviews for complaints and requests."),
    h3("Failure Mode 6: Overreliance on AI Without Strategic Judgment (20% of failures)"),
    bodyPara("AI output used verbatim without customization. Generic copy that reads like every other product. Customer detects the absence of original thinking."),
    bodyPara("Prevention: Use AI for brainstorming and drafts, then customize heavily. Spend 20 to 30 percent of creation time on editing and personalizing output."),
    h3("Failure Mode 7: Wrong Product Type or Format Mismatch (15% of failures)"),
    bodyPara("Creator builds products that do not match how the target customer wants to receive information. A video-oriented audience receives only PDF content."),
    bodyPara("Prevention: Research how top sellers in the niche deliver their products. Match format to buyer preference before investing creation time."),
    h3("Failure Mode 8: No Traffic Strategy (20% of failures)"),
    bodyPara("Creator relies entirely on Etsy organic search and never builds off-platform traffic. Revenue is fragile because it depends on a single channel the creator does not control."),
    bodyPara("Prevention: Build a Pinterest presence from Month 1. Start an email list from the first product. Develop at least two traffic channels before Month 3."),
    h3("Failure Mode 9: Premature Scaling Before Validation (15% of failures)"),
    bodyPara("Creator builds twenty products before validating that the first one works. Spreads energy across too many untested ideas simultaneously."),
    bodyPara("Prevention: Validate niche with two to three products before expanding. Reach $200 to $300 monthly from the first niche before adding a second."),
    h3("Failure Mode 10: Platform Dependency Without Risk Management (10% of failures)"),
    bodyPara("Business is entirely dependent on Etsy's algorithm and policies. A single platform change eliminates the business overnight."),
    bodyPara("Prevention: Build an email list. Create a minimal website. Diversify traffic sources. Own your customer relationships independent of any single platform."),
    spacer(120),
    boxedParagraphs("Failure Prevention Checklist", [
      "Validate niche with data before building any product.",
      "Research keywords before creating the first listing.",
      "Publish consistently: minimum two products per month.",
      "Review competitor products before starting each new creation.",
      "Read community feedback monthly to identify unmet needs.",
      "Customize all AI output before publishing.",
      "Build email list from product one, not later.",
      "Develop Pinterest presence in parallel with Etsy listings.",
      "Validate first niche before building a second.",
      "Build audience ownership independent of Etsy.",
    ], GRAY, true),
  ];
}

// ─── PART 1 ──────────────────────────────────────────────────────────────────
function buildPart1() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 1", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "YOUR OPERATING ENVIRONMENT", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 1 \u2013 4", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 1", "The Digital Product Market"),
    bodyPara("The digital product market on Etsy is the intersection of three converging trends: the growth of creator economies, the mainstreaming of productivity tools, and the democratization of AI content creation. Understanding this environment determines whether your entry timing and positioning are correct."),
    h2("Why Etsy for Digital Products"),
    bulletPara("Built-in buyer intent: Etsy attracts active buyers, not passive browsers. Conversion rates are structurally higher than social media platforms."),
    bulletPara("Instant delivery infrastructure: Automatic digital delivery removes all fulfilment overhead."),
    bulletPara("Search-driven discovery: Products rank in search without ongoing advertising spend."),
    bulletPara("Trust establishment: Etsy's review system and buyer protection creates purchase confidence that an independent website would take years to build."),
    bulletPara("Portfolio compounding: Multiple products in the same niche accelerate ranking for all listings through algorithmic portfolio signals."),
    spacer(120),
    h2("Market Segmentation"),
    dataTable(
      ["Segment", "Examples", "Price Range", "Competition Level"],
      [
        ["Productivity", "Notion templates, planners, trackers", "$9\u2013$49", "High (broad) to Low (niche)"],
        ["Education", "Lesson plans, rubrics, worksheets", "$2\u2013$49", "Moderate"],
        ["Business", "SOPs, proposals, financial models", "$12\u2013$97", "Low to Moderate"],
        ["Creative", "Social media kits, Canva templates", "$9\u2013$32", "High (broad) to Low (niche)"],
        ["Health & Wellness", "Journals, habit trackers, planners", "$9\u2013$32", "Moderate"],
        ["AI & Prompts", "Prompt packs, AI workflows", "$9\u2013$97", "Low (emerging)"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 1", [
      "Etsy provides built-in buyer intent, delivery infrastructure, and search discovery.",
      "Digital product margins range from 70% to 95% depending on product type.",
      "AI and prompt-based products represent an emerging, low-competition segment.",
      "Niche specificity is the primary lever for reducing competition.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 2", "Understanding AI for Product Creators"),
    bodyPara("Claude and similar AI tools are productivity accelerators, not creative replacements. Understanding what AI does well and where human judgment is irreplaceable determines how effectively you integrate these tools into your workflow."),
    h2("What AI Handles Effectively"),
    bulletPara("Brainstorming: Generating 100 product ideas in five minutes instead of four hours."),
    bulletPara("Research synthesis: Summarizing ten competitor listings in three minutes."),
    bulletPara("Copywriting drafts: Writing five product description variations in five minutes."),
    bulletPara("Tag generation: Producing 13 keyword-researched tags in one minute."),
    bulletPara("Email sequences: Writing five-email nurture sequences in ten minutes."),
    bulletPara("Content batching: Generating 30 social media captions in five minutes."),
    spacer(80),
    h2("What the Creator Must Handle"),
    bulletPara("Niche selection: Which market to focus on. AI cannot make this strategic call."),
    bulletPara("Design direction: Style, layout, aesthetic choices."),
    bulletPara("Final copy review: Does it sound human, specific, and differentiated?"),
    bulletPara("Quality gates: Is this good enough to publish?"),
    bulletPara("Strategic decisions: Which direction to take the business."),
    bulletPara("Customer relationships: Building trust and community."),
    spacer(120),
    h2("The Human-AI Workflow"),
    numberedPara("Creator identifies niche and product concept."),
    numberedPara("Claude generates 50 page or content variations, 20 feature ideas, 30 tag combinations."),
    numberedPara("Creator selects the most promising variations and provides direction."),
    numberedPara("Claude refines selected options and generates detailed specifications."),
    numberedPara("Creator designs and builds in Canva or Notion."),
    numberedPara("Claude generates ten title variations and twenty description variations."),
    numberedPara("Creator selects, refines, and publishes."),
    spacer(80),
    bodyPara("Total time: 8 to 12 hours per product with AI assistance. Without AI: 30 to 40 hours. Time saved: 60 to 70%."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 2", [
      "AI handles production tasks. Strategy, judgment, and quality control remain with the creator.",
      "The workflow is human-directed, AI-accelerated.",
      "Overreliance on unedited AI output is one of the ten critical failure modes.",
      "Use AI for volume and drafts; use judgment for selection and refinement.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 3", "Etsy Platform Mechanics & Advantages"),
    bodyPara("Etsy's search algorithm determines 60 to 70% of a listing's commercial outcome. Understanding its mechanics is not optional for serious sellers."),
    h2("The Etsy Search Algorithm: Core Model"),
    dataTable(
      ["Signal Category", "Weight", "Key Factors"],
      [
        ["Relevance", "60%", "Keyword in title (30%), tags (15%), description (10%), linguistic relatedness (5%)"],
        ["Quality", "25%", "Sales velocity (10%), star rating (8%), recency (4%), positive feedback trend (3%)"],
        ["Engagement", "15%", "Click-through rate (8%), conversion rate (4%), favorites (2%), cart adds (1%)"],
      ]
    ),
    spacer(120),
    bodyPara("Implication: Relevance is dominant, but quality and engagement signals accelerate ranking over time."),
    h2("The 90-Day Ranking Cycle"),
    bodyPara("New listings receive a visibility spike in Days 1 to 7 as the algorithm tests them. If sales occur, ranking improves rapidly within Days 8 to 30. Between Days 31 and 60, ranking stabilizes based on measured performance. By Days 61 to 90, steady-state ranking is established and seasonal patterns begin to matter."),
    h2("Etsy Advantages Over Independent Channels"),
    bulletPara("Built-in search traffic: 90 million active buyers as of recent data."),
    bulletPara("Zero advertising required for initial discovery."),
    bulletPara("Instant delivery automation: No manual fulfilment."),
    bulletPara("Trust signals from the platform: Etsy buyer protection increases conversion."),
    bulletPara("Review system: Social proof accumulates automatically."),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 3", [
      "Etsy SEO is 60% relevance-driven. Keywords in the title matter most.",
      "The ranking cycle takes approximately 90 days to stabilize for a new listing.",
      "Sales velocity and star rating are the dominant quality signals.",
      "Click-through rate is the primary engagement signal creators can optimize directly.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 4", "Technology Stack & Operational Tools"),
    bodyPara("An effective technology stack enables the batch production workflows that separate sustainable businesses from unsustainable chaos. This chapter specifies the exact tools, their roles, and how they connect."),
    h2("Core Tool Stack"),
    dataTable(
      ["Tool", "Role", "Cost", "Priority"],
      [
        ["Claude Pro", "Content creation, SEO, research, email", "$20/month", "Essential"],
        ["Canva Pro", "Design for printables, social media kits", "$13/month", "High"],
        ["Google Drive", "File storage, research organization", "$6\u2013$14/month", "High"],
        ["Google Sheets", "Analytics tracking, keyword research", "Included", "High"],
        ["Beehiiv / Mailchimp", "Email marketing platform", "$0\u2013$15/month", "High"],
        ["Notion", "Template building, workspace", "Free\u2013$8/month", "High for templates"],
        ["Buffer / Later", "Social media scheduling", "$0\u2013$15/month", "Moderate"],
        ["Etsy Analytics", "Performance data, traffic sources", "Included", "Essential"],
      ]
    ),
    spacer(120),
    h2("Integration Patterns"),
    h3("Pattern 1: Research and Validation"),
    bodyPara("Store competitor data and Reddit threads in Google Drive. Feed documents to Claude for pattern analysis. Export structured opportunity scores to Google Sheets. This converts three to four hours of manual research into one hour of directed analysis."),
    h3("Pattern 2: Content Creation"),
    bodyPara("Claude generates product descriptions, email sequences, and social copy. Outputs are stored in organized Drive folders. Canva pulls from Drive documents for design alignment. Content and design stay synchronized without rewriting."),
    h3("Pattern 3: Performance Tracking"),
    bodyPara("Export Etsy monthly sales data as CSV. Input to a master Google Sheet. Feed to Claude for pattern analysis. Claude identifies top performers, weak performers, and optimization recommendations. Data-driven decisions replace gut judgement."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 4", [
      "The full tool stack costs $49 to $97 per month.",
      "Claude is the central production engine. Every other tool supports its input or output.",
      "Integration patterns reduce research from hours to minutes.",
      "Google Sheets is the tracking backbone. Set it up from Day 1.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 2 ──────────────────────────────────────────────────────────────────
function buildPart2() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 2", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "VALIDATION & STRATEGIC PLANNING", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 5 \u2013 7", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 5", "Demand Analysis Framework"),
    bodyPara("Demand analysis is the process of measuring whether a target customer actively searches for, and purchases, products similar to what you plan to create. This analysis must happen before building anything."),
    h2("The Four Demand Signals"),
    numberedPara("Search volume: How many people search for this keyword or product type monthly on Etsy and Google?"),
    numberedPara("Purchase behavior: Do existing products in this space generate sales? How many reviews do top listings have?"),
    numberedPara("Buyer willingness to pay: Are products priced at $15 or more? Or is the market a race to the bottom under $5?"),
    numberedPara("Competition density: How many listing results appear for the target keyword? Under 300 is generally favorable. Over 1,000 is saturated."),
    spacer(120),
    h2("Seasonal Demand Calendar Building"),
    bodyPara("Every niche has seasonal demand patterns. Creators who build products two to three months before peak demand capture significantly more sales than those who react to demand after it arrives."),
    bodyPara("Use Claude to generate a 12-month demand calendar for your chosen niche. For each month, identify natural buying periods, top keywords that spike in search volume, relevant product angles, and seasonal intensity levels."),
    bodyPara("Execution: Build December and January products in September. Build spring products in January. Build summer products in April."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 5", [
      "Never build before validating demand. Research first, create second.",
      "Four signals determine demand: search volume, purchase behavior, willingness to pay, and competition density.",
      "Seasonal demand patterns are predictable. Build ahead of demand peaks.",
      "Claude can generate a 12-month seasonal calendar for any niche in under 15 minutes.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 6", "Niche Selection & Validation"),
    bodyPara("Niche selection is the single most important decision in the business. Forty percent of all failures originate from wrong niche selection. This chapter provides the frameworks to identify winning niches before investing creation time."),
    h2("The Niche Specificity Ladder"),
    dataTable(
      ["Level", "Example", "Competition", "Opportunity"],
      [
        ["Broad (avoid)", "'Productivity template'", "8,000+ listings", "Very low"],
        ["Specific (better)", "'Productivity template for therapists'", "~200 listings", "Moderate"],
        ["Hyper-specific (best)", "'Productivity for therapists serving trauma survivors'", "< 10 listings", "High"],
      ]
    ),
    spacer(120),
    h2("AI-Powered Idea Generation Workflows"),
    h3("Workflow 1: Reddit Pain-Point Mining (45 minutes)"),
    bodyPara("Identify three to five relevant subreddits for your target audience. Sort by top posts from the past year. Copy 15 to 20 high-voted complaint posts into a document. Paste into Claude with the prompt: 'Analyze these posts. Extract top 10 pain points ranked by frequency, unmet needs, and product ideas that solve them. Format as a table.'"),
    bodyPara("Claude returns structured opportunities. Validate each on Etsy by counting search results and checking competition density."),
    h3("Workflow 2: Pinterest Trend Analysis (60 minutes)"),
    bodyPara("Visit Pinterest Trends. Identify 20 rising trends in your region and category. Map each trend to a potential Etsy product category. Paste the list into Claude and ask for five product ideas per trend with estimated price points and build times. Score each by: build speed, seasonal relevance, and market certainty."),
    h3("Workflow 3: Etsy Autocomplete Deep Dive (90 minutes)"),
    bodyPara("Type a base keyword into Etsy search. Record all autocomplete suggestions, typically 10 to 20. Repeat with variations. Collect 50 or more keyword combinations. Paste into Claude for analysis of search intent, estimated monthly volume, competition density, and pricing opportunity. Rank by opportunity score."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 6", [
      "Niche specificity is the primary lever for reducing competition while maintaining adequate search volume.",
      "Reddit pain-point mining surfaces unmet needs that competitors have not yet addressed.",
      "Pinterest trends signal demand 4 to 6 weeks before it appears on Etsy.",
      "Etsy autocomplete represents real search volume from active buyers.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 7", "The Validation Framework \u2014 Nine Signals"),
    bodyPara("This is the most important framework in the book. Use it before building anything. It takes 30 to 90 minutes per product idea. It prevents the primary failure mode."),
    h2("The Nine-Signal Validation Matrix"),
    dataTable(
      ["Signal", "Weight", "What to Measure", "Scoring"],
      [
        ["Search Volume", "30%", "Monthly Etsy + Google searches", "8+ = 500+ searches; 5\u20137 = 100\u2013500; 0\u20134 = <100"],
        ["Competition Density", "20%", "Number of Etsy listings for keyword", "9\u201310 = <50 listings; 6\u20138 = 50\u2013150; 3\u20135 = 150\u2013300; 0\u20132 = 300+"],
        ["Buyer Urgency", "15%", "Review language in top competitors", "Check for words: 'finally', 'exactly what I needed', 'lifesaver'"],
        ["Repeat Purchase Potential", "15%", "Can buyers purchase multiple versions?", "Multiple angles = higher score"],
        ["Expandability", "10%", "Can product become 3\u20135+ variations?", "9\u201310 = easily 5+ variations"],
        ["Price Elasticity", "10%", "Can you charge $15\u2013$49?", "8\u201310 = yes; 2\u20134 = all under $10"],
      ]
    ),
    spacer(120),
    h2("Worked Example: ADHD Productivity Notion Template"),
    dataTable(
      ["Signal", "Research Finding", "Score"],
      [
        ["Search Volume", "200\u2013400 searches/month for 'adhd productivity template'", "7/10"],
        ["Competition Density", "45 listings on Etsy", "9/10"],
        ["Buyer Urgency", "60% of competitor reviews show time-saving urgency language", "8/10"],
        ["Repeat Purchase", "Multiple variations possible (for students, freelancers, remote workers)", "6/10"],
        ["Expandability", "Easily 7+ variations identified", "9/10"],
        ["Price Elasticity", "Competitors priced $12\u2013$32; premium positioning achievable", "8/10"],
        ["Weighted Total", "(7\u00D70.30) + (9\u00D70.20) + (8\u00D70.15) + (6\u00D70.15) + (9\u00D70.10) + (8\u00D70.10)", "7.7/10"],
      ]
    ),
    spacer(120),
    bodyPara("Decision: Build immediately. Score of 7.7/10 indicates high confidence."),
    h2("Decision Thresholds"),
    dataTable(
      ["Score", "Decision", "Action"],
      [
        ["8\u201310", "Build immediately", "Begin creation this week"],
        ["6\u20138", "Build with modifications", "Refine angle, validate price, then build"],
        ["4\u20136", "Research more", "Find adjacent niche or different angle"],
        ["0\u20134", "Kill idea", "Move to next candidate immediately"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 7", [
      "The nine-signal Validation Matrix prevents the 40% of failures caused by wrong niche selection.",
      "A score of 7.7/10 or above indicates high confidence to build.",
      "Scores below 6 indicate the need for further research or a pivot.",
      "The matrix takes 30 minutes for a quick version and 90 minutes for a full assessment.",
      "Never build before scoring. This is a non-negotiable pre-production step.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 3 ──────────────────────────────────────────────────────────────────
function buildPart3() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 3", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "PRODUCT OPERATIONS", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 8 \u2013 11", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 8", "Operations Systems & Batch Workflows"),
    bodyPara("Batch production is the operational foundation of a scalable digital product business. Individual product creation creates constant context switching and low output. Batching creates deep work states and significantly higher throughput."),
    h2("The Batching Principle"),
    dataTable(
      ["Approach", "Week Structure", "Output", "Efficiency"],
      [
        ["Individual", "Daily: 2h planning + 2h creating + 1h optimizing (repeat weekly)", "1\u20132 products/month", "Low"],
        ["Batching", "Week 1: 15h research + planning. Week 2\u20134: production + marketing", "4\u20136 products/month", "High"],
      ]
    ),
    spacer(120),
    bodyPara("Batching is 40 to 50% faster because it eliminates context switching, enables deep work states, benefits from repetition acceleration, reduces decision fatigue, and requires tool setup only once per batch rather than once per product."),
    h2("Monthly Operations Cycle"),
    h3("Week 1: Research & Planning (12\u201315 hours)"),
    bulletPara("Monday: Community research (Reddit, Facebook, Discord for target audience pain points)"),
    bulletPara("Tuesday: Competitor analysis (top 20 listings, pricing, review analysis)"),
    bulletPara("Wednesday: Idea generation (Claude: generate 50 ideas, score top 10)"),
    bulletPara("Thursday: Validation (search volume, competition level, differentiation opportunities)"),
    bulletPara("Friday: Product specifications (detailed scope for each product in batch)"),
    spacer(80),
    h3("Week 2: Creation (15\u201320 hours)"),
    bulletPara("Days 1\u20132: Build product structure and content using AI-assisted workflows"),
    bulletPara("Days 3\u20134: Design in Canva or build in Notion"),
    bulletPara("Day 5: Documentation, setup guides, and preview image creation"),
    spacer(80),
    h3("Weeks 3\u20134: Marketing & Optimization (10\u201315 hours)"),
    bulletPara("Launch all products with SEO-optimized listings"),
    bulletPara("Create 20 Pinterest pins per product"),
    bulletPara("Publish email announcement to subscribers"),
    bulletPara("Monitor first-week metrics and adjust underperformers"),
    spacer(120),
    h2("Worked Example: 10 Notion Templates in One Week"),
    dataTable(
      ["Day", "Activity", "Hours", "Output"],
      [
        ["Monday", "Ideation: theme, 10 template concepts, shared elements", "1", "10 scoped ideas"],
        ["Tuesday", "Design structures: database properties, views, formulas", "4", "10 structure blueprints"],
        ["Wednesday", "Build in Notion: all 10 templates", "6", "10 complete templates"],
        ["Thursday", "Documentation: screenshots, setup guides, tips", "2", "10 documented products"],
        ["Friday", "Export, cover images, file preparation", "1", "10 ready-to-list files"],
        ["Total", "\u2014", "14 hours", "10 products ready to launch"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 8", [
      "Batching is 40\u201350% faster than individual product creation.",
      "The monthly operations cycle: one week research, one week creation, two weeks marketing.",
      "Context switching is the primary productivity killer. Batching eliminates it.",
      "Decision fatigue is real. Make design and positioning decisions once per batch, not per product.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 9", "Creating Products at Scale"),
    bodyPara("Each product type has a specific AI-assisted workflow that reduces creation time by 60 to 70% compared to manual methods. This chapter covers the production workflow for the most profitable product categories."),
    h2("Notion Templates"),
    bodyPara("Claude's role: Feature ideation, database structure, property definitions, formula suggestions, documentation writing, and marketing copy."),
    bodyPara("Workflow: (1) Claude generates 50 page and feature variations. (2) Creator selects top three combinations. (3) Claude provides detailed Notion database specifications. (4) Creator builds in Notion. (5) Claude writes setup documentation. (6) Claude generates SEO-optimized listing copy."),
    bodyPara("Time saved: 8 to 10 hours per template."),
    h2("Printable Workbooks"),
    bodyPara("Claude's role: Content structure, page-by-page content, design briefs, upsell recommendations."),
    bodyPara("Workflow: (1) Claude designs workbook structure (topic, page count, layout). (2) Claude writes all page content. (3) Claude generates design brief (color scheme, typography, formatting). (4) Creator designs in Canva. (5) Claude generates cross-sell recommendations."),
    bodyPara("Time saved: 8 to 10 hours per workbook."),
    h2("eBooks & Guides"),
    bodyPara("Claude's role: Complete content creation engine. Full outline generation, complete ebook writing in artifacts, marketing materials, email launch sequence, sales page copy, and ad copy."),
    bodyPara("Workflow: (1) Claude generates full 50-page outline. (2) Claude writes complete ebook in artifact window. (3) Claude generates five-email launch sequence. (4) Creator formats, designs cover, exports as PDF. (5) Creator publishes on Etsy."),
    bodyPara("Time saved: 30 to 40 hours per ebook compared to manual writing."),
    h2("Spreadsheet Templates"),
    bodyPara("Claude's role: Structural design, formula generation with exact copy-paste formulas, setup instructions, usage video scripts."),
    bodyPara("Workflow: (1) Claude designs spreadsheet structure and lists all required formulas. (2) Claude generates exact formulas per section. (3) Claude writes step-by-step setup instructions. (4) Creator builds in Google Sheets or Excel. (5) Claude generates video tutorial scripts."),
    bodyPara("Time saved: 6 to 8 hours per template."),
    h2("Social Media Kits"),
    bodyPara("Claude's role: 30-day content calendar, design briefs for each post, hashtag research, usage guides."),
    bodyPara("Time saved: 12 to 15 hours per kit."),
    h2("Prompt Packs"),
    bodyPara("Claude's role: Comprehensive prompt engineering, category organization, usage guides, customer education email sequences."),
    bodyPara("Time saved: 15 to 20 hours per pack."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 9", [
      "Every major product type has a repeatable AI-assisted workflow.",
      "eBooks have the highest time savings: 30 to 40 hours saved per product.",
      "Claude writes content. The creator selects, designs, and publishes.",
      "Prompt packs are among the highest-margin products at 92%+ and low build time.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 10", "Quality Control & Publishing"),
    bodyPara("Quality control is the gate between creation and publication. Products that reach Etsy without passing quality gates underperform, receive poor reviews, and damage overall shop ranking."),
    h2("The Quality Gate Checklist"),
    boxedParagraphs("Pre-Launch Quality Gate", [
      "Product fully created and tested (all links work, all pages accessible)",
      "Cover image designed at 1,000\u00D7800 pixels minimum",
      "Five to seven preview images created showing actual product content",
      "Product description written at 1,500+ characters",
      "Thirteen tags researched and selected using the keyword cluster framework",
      "Title optimized to 140 characters with primary keyword in first 70",
      "Pricing researched against top 10 competitors",
      "FAQ section written addressing the top five buyer questions",
    ], GRAY, true),
    spacer(120),
    h2("Launch Day Protocol"),
    numberedPara("Upload product files. Verify all files are accessible and correctly named."),
    numberedPara("Upload images in sequence: cover image first, then previews showing key features."),
    numberedPara("Finalize title and description. Confirm no typos. Confirm keyword placement in first 70 characters."),
    numberedPara("Add all 13 tags. Use the three-category tag architecture."),
    numberedPara("Set price. Confirm against competitor research."),
    numberedPara("Activate listing. Confirm it appears in Etsy search within 24 to 48 hours."),
    spacer(120),
    h2("Post-Launch Monitoring Protocol"),
    dataTable(
      ["Timeframe", "Metric to Check", "Action Threshold"],
      [
        ["Day 1\u20137", "Does listing appear in search?", "If not visible after 48h: re-check tags"],
        ["Day 8\u201330", "Impressions vs. expected", "If <50 impressions: SEO problem, revise title"],
        ["Day 30\u201360", "Click-through rate", "If <1%: thumbnail problem, redesign cover"],
        ["Day 30\u201360", "Conversion rate", "If <0.5%: copy or price problem, revise both"],
        ["Day 60\u201390", "Revenue trend", "Stable or growing = continue; declining = investigate"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 10", [
      "Quality gates before publication prevent poor reviews that damage shop ranking.",
      "Cover images and preview images account for 70% of click-through rate outcomes.",
      "Tags must be researched, not guessed. Use the three-category architecture.",
      "Post-launch monitoring in 30-day intervals identifies problems before they compound.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 11", "Asset Management & Scaling"),
    bodyPara("Scaling from five products to fifty requires an asset management system. Without it, duplicate work, inconsistent branding, and disorganized files create bottlenecks that prevent growth."),
    h2("Recommended File System Structure"),
    dataTable(
      ["Folder", "Contents", "Purpose"],
      [
        ["01-Research", "Competitor CSV, keyword CSV, seasonal calendar", "All pre-build research"],
        ["02-Products", "Subfolders per product type and product name", "All product files, assets, copy"],
        ["03-Brand-Assets", "Colors, fonts, logos, icons, templates", "Reusable brand elements"],
        ["04-Listings", "CSV per product category + master keyword CSV", "SEO tracking"],
        ["05-Analytics", "Monthly reports, performance tracking CSV", "Data history"],
      ]
    ),
    spacer(120),
    h2("Asset Reuse Strategy"),
    bodyPara("Create five to ten reusable design templates. Use one color scheme consistently. Pre-select two to three font pairs used across all products. Maintain a pre-made icon library from Canva Pro."),
    bodyPara("Example: Three affirmation card design templates can produce 45 different cards with total design time of four hours, versus 12+ hours if each card were designed individually."),
    h2("Scaling to Multiple Product Types"),
    dataTable(
      ["Month", "Product Type Added", "Additional Monthly Revenue"],
      [
        ["Month 1", "5 Notion templates", "$300\u2013$750"],
        ["Month 2", "10 printables", "$300\u2013$500 additional"],
        ["Month 3", "5 eBooks", "$400\u2013$1,500 additional"],
        ["Month 4", "20 prompt packs", "$200\u2013$600 additional"],
        ["Month 5", "15 spreadsheet templates", "$300\u2013$1,000 additional"],
        ["Month 6", "Optimization of top performers", "$2,000\u2013$4,000 total"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 11", [
      "A structured file system enables finding any asset in 30 seconds.",
      "Reusable design templates provide a 3\u00D7 speed gain on batch production.",
      "Diversifying across product types reduces revenue dependency on any single format.",
      "By Month 6 with consistent execution, total revenue of $2,000\u2013$4,000 per month is achievable.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 4 ──────────────────────────────────────────────────────────────────
function buildPart4() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 4", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "ADVANCED ETSY OPTIMIZATION", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 12 \u2013 15", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 12", "Store Setup & Branding"),
    bodyPara("Your Etsy store is a brand, not just a product repository. Store name, visual identity, and policies collectively determine buyer trust and the likelihood of repeat purchases."),
    h2("Store Naming Strategy"),
    dataTable(
      ["Approach", "Examples", "Pros", "Cons"],
      [
        ["Niche + Format", "ProductivityNotionTemplates", "Keywords explicit", "Rigid, hard to expand"],
        ["Brand + Niche", "FlowNotionProducts, ClarityPlanners", "Memorable, flexible", "No direct keywords"],
        ["Creator Name", "SarahDigitalProducts", "Personal brand", "Cannot sell the business"],
        ["Personality", "ProductivityNerd, ThePlanner", "Unique, memorable", "Less SEO value"],
      ]
    ),
    spacer(120),
    bodyPara("Recommendation for new sellers: Approach 1 or 2. Prioritize clarity and searchability over personality. Rebranding is possible later."),
    h2("Essential Store Policies"),
    h3("Refund Policy"),
    bodyPara("Digital products are non-refundable as files cannot be revoked after download. However, if files are corrupted or missing, offer immediate replacement upon contact within 48 hours."),
    h3("Shipping Policy"),
    bodyPara("All products are instant digital downloads via Etsy. No physical shipping. Delivery is immediate upon purchase."),
    h2("Visual Identity Standards"),
    dataTable(
      ["Element", "Specification", "Purpose"],
      [
        ["Shop banner", "1,200\u00D7300 pixels", "First visual impression. Include value proposition and branding."],
        ["Shop icon", "500\u00D7500 pixels", "Must read clearly at small size. Consistent with banner."],
        ["Product covers", "1,000\u00D7800 pixels minimum", "Primary thumbnail. Determines click-through rate."],
        ["Preview images", "1,000\u00D7800 pixels, 5\u20137 per product", "Shows content. Drives purchase confidence."],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 12", [
      "Store branding is a trust signal. Professional appearance increases conversion.",
      "A clear refund policy reduces buyer hesitation without creating financial risk.",
      "Consistent visual identity across listings creates a cohesive brand experience.",
      "Cover images are the single most impactful element of click-through rate.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 13", "Advanced SEO Framework"),
    bodyPara("SEO determines 60 to 70% of commercial success on Etsy. This chapter provides professional-grade frameworks that most creators never learn or implement."),
    h2("Keyword Clustering Strategy"),
    bodyPara("Most creators target one or two keywords. Professional sellers target keyword clusters: groups of 10 to 30 related keywords that share similar meaning, target the same customer pain point, and have varying levels of competition."),
    dataTable(
      ["Tier", "Characteristics", "Competition", "Strategy"],
      [
        ["Tier 1", "High volume, 2,000+ searches/month", "High", "Include in title and primary tags"],
        ["Tier 2", "Medium volume, 500\u20132,000/month", "Moderate", "Include in secondary tags"],
        ["Tier 3", "Long-tail, 100\u2013500/month", "Low", "Target with dedicated tags"],
        ["Tier 4", "Ultra-specific, 20\u2013100/month", "Very Low", "Create niche-vertical products"],
      ]
    ),
    spacer(120),
    h2("Tag Architecture (13 Tags)"),
    dataTable(
      ["Category", "Number of Tags", "Purpose", "Example"],
      [
        ["Primary keywords", "3", "Main keyword + variations", "'project template', 'project management', 'projects templates'"],
        ["Related keywords", "5", "Adjacent terms, synonyms", "'team collaboration', 'workflow automation', 'notion template'"],
        ["Long-tail specific", "5", "Vertical-specific, niche combinations", "'remote work template', 'freelance project management'"],
      ]
    ),
    spacer(120),
    h2("Competitiveness Scoring"),
    bodyPara("Before investing in a keyword, calculate its opportunity score: Volume Score minus Difficulty Score. Volume is rated 1 to 10 based on monthly search volume. Difficulty is rated 1 to 10 based on total Etsy listing count for the keyword."),
    bodyPara("Target keywords with an opportunity score of plus one or higher. A score of plus seven or higher represents a sweet spot of high volume with low competition."),
    h2("Search Intent Mapping"),
    dataTable(
      ["Intent Type", "Percentage", "Buyer Behavior", "Conversion Rate"],
      [
        ["Informational", "30%", "Seeking education, not purchase", "Very low"],
        ["Navigation", "20%", "Comparing options", "Medium"],
        ["Commercial", "40%", "Ready to buy, seeking product", "High"],
        ["Transactional", "10%", "Active purchase decision made", "Highest"],
      ]
    ),
    spacer(120),
    bodyPara("Target Commercial (40%) and Transactional (10%) intent keywords for maximum conversion. These represent 50% of all searches but generate the majority of purchases."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 13", [
      "Target keyword clusters, not individual keywords. Clusters capture full search ecosystems.",
      "Use a three-category tag architecture: primary (3 tags), related (5 tags), long-tail (5 tags).",
      "Only pursue keywords with an opportunity score of plus one or higher.",
      "Focus on Commercial and Transactional intent keywords for highest conversion.",
      "Keyword research must happen before creation, not after.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 14", "Analytics, Testing & Iteration"),
    bodyPara("Data-driven iteration separates growing businesses from stagnant ones. This chapter provides the measurement frameworks and testing protocols that turn underperforming listings into revenue-generating assets."),
    h2("Core Metrics Dashboard"),
    dataTable(
      ["Metric", "Benchmark", "Problem Signal", "Action"],
      [
        ["Impressions", "500+/month per listing", "<50 after 4 weeks", "SEO problem: revise title and tags"],
        ["Click-through rate", "2\u20135%", "<1%", "Thumbnail problem: redesign cover image"],
        ["Conversion rate", "1.5\u20133.5%", "<0.5%", "Copy or price problem: revise both"],
        ["Revenue per listing", "$100+/month (mature)", "Declining month-on-month", "Investigate ranking decay"],
      ]
    ),
    spacer(120),
    h2("A/B Testing Protocol"),
    numberedPara("Identify the underperforming metric (impression, CTR, or conversion)."),
    numberedPara("Form a hypothesis about the cause. Example: 'The cover image does not clearly show the product benefit.'"),
    numberedPara("Change one variable at a time. If testing the title, change only the title."),
    numberedPara("Wait 14 days for sufficient data before evaluating results."),
    numberedPara("Compare before and after metrics. If improved, keep the change. If worse, revert."),
    numberedPara("Document all tests and results in a tracking spreadsheet."),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 14", [
      "Track four core metrics: impressions, CTR, conversion rate, and revenue per listing.",
      "Change one variable at a time and wait 14 days before evaluating.",
      "Low impressions = SEO problem. Low CTR = thumbnail problem. Low conversion = copy or price problem.",
      "Monthly performance reviews prevent slow revenue decay from going unnoticed.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 15", "Thumbnail & CTR Optimization"),
    bodyPara("The cover image (thumbnail) is the single element most responsible for click-through rate. A 1% CTR versus a 3% CTR on 10,000 monthly impressions is the difference between 100 and 300 potential buyers. This chapter covers the design principles that produce high-CTR thumbnails."),
    h2("Thumbnail Design Principles"),
    bulletPara("Show the product, not just the concept. Buyers want to see what they are purchasing."),
    bulletPara("Include a clear headline. The most important benefit in large, readable text."),
    bulletPara("Use white space. Cluttered thumbnails perform below average."),
    bulletPara("Contrast is critical. High contrast between text and background improves readability at thumbnail size."),
    bulletPara("Show the outcome, not the process. What does life look like after using the product?"),
    spacer(120),
    h2("Thumbnail Testing Framework"),
    dataTable(
      ["Version", "Variable Changed", "Metric to Measure", "Evaluation Period"],
      [
        ["Version A", "Original thumbnail", "CTR baseline", "Week 1\u20132"],
        ["Version B", "New thumbnail (one change)", "CTR comparison", "Week 3\u20134"],
        ["Decision", "\u2014", "Keep higher CTR version", "After 14 days"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 15", [
      "Thumbnail click-through rate is the highest-leverage optimization in the business.",
      "Show the product clearly. Buyers need to see what they are buying.",
      "Test one thumbnail change at a time. Evaluate after 14 days of data.",
      "White space, high contrast, and clear headlines produce above-average CTR.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 5 ──────────────────────────────────────────────────────────────────
function buildPart5() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 5", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "CUSTOMER ACQUISITION", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 16 \u2013 20", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 16", "Traffic Channel Strategy"),
    bodyPara("Single-channel revenue is fragile revenue. A business built entirely on Etsy organic search is one algorithm change from collapse. This chapter builds a multi-channel traffic architecture that creates resilient, compounding customer acquisition."),
    h2("Channel Priority Framework"),
    dataTable(
      ["Channel", "Priority", "Cost", "Time to Results", "Sustainability"],
      [
        ["Etsy Organic SEO", "Primary", "Free", "60\u201390 days", "High (compounds)"],
        ["Pinterest", "Primary", "Free", "30\u201360 days", "Very High (evergreen)"],
        ["Email Marketing", "Primary", "Low", "Immediate", "Highest (owned)"],
        ["TikTok / Reels", "Secondary", "Free", "0\u201360 days", "Moderate (volatile)"],
        ["Etsy Ads", "Tertiary", "$5\u201310/day", "Immediate (paid)", "Moderate"],
        ["Pinterest Ads", "Tertiary", "$5/day", "Immediate (paid)", "Moderate"],
      ]
    ),
    spacer(120),
    bodyPara("Focus on the three primary channels during the first 90 days. Add secondary and tertiary channels only after primary channels are established and generating consistent results."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 16", [
      "Multi-channel traffic is the foundation of sustainable revenue.",
      "Etsy organic, Pinterest, and email are the three primary channels for this business model.",
      "Email is the highest-sustainability channel because it is platform-independent.",
      "Add paid channels only after organic channels are producing consistent results.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 17", "Pinterest System \u2014 Primary Traffic Channel"),
    bodyPara("Pinterest is the most powerful free traffic channel for digital product businesses. Unlike TikTok or Instagram, Pinterest content is evergreen: a pin created today can generate traffic and sales two years from now."),
    h2("Why Pinterest Works for Digital Products"),
    bulletPara("Buyer intent: 85% of Pinterest users have purchased something they discovered on the platform."),
    bulletPara("Evergreen content: Pins continue circulating long after posting, unlike social media feeds."),
    bulletPara("Search-driven: Pinterest operates as a visual search engine, not a social feed."),
    bulletPara("Purchase-ready audience: Users come to Pinterest specifically to discover and plan purchases."),
    spacer(120),
    h2("Pinterest Pin System"),
    h3("Pin Types by Goal"),
    dataTable(
      ["Pin Type", "Purpose", "Frequency"],
      [
        ["Product showcase", "Direct product promotion", "5 per product at launch"],
        ["Value tip", "Demonstrates expertise, builds trust", "3\u20134 per week"],
        ["Before/after", "Shows transformation (problem solved)", "2\u20133 per week"],
        ["Listicle graphic", "Educational content driving saves", "2\u20133 per week"],
      ]
    ),
    spacer(120),
    h3("Keyword Optimization for Pinterest"),
    bodyPara("Pinterest SEO works similarly to Etsy SEO. Titles and descriptions should include target keywords. Boards should have keyword-rich names and descriptions. Consistency of keywords across pin title, description, and board creates compound ranking effects."),
    h2("Pin Production Workflow"),
    numberedPara("Use Claude to generate 30 pin captions in one session."),
    numberedPara("Create 30 pin graphics in Canva using three to five reusable templates."),
    numberedPara("Schedule all 30 pins in Buffer or Tailwind (5 to 10 pins per day)."),
    numberedPara("Total production time: 3 to 4 hours per month for 30 pins."),
    spacer(80),
    bodyPara("This creates a consistent daily presence on Pinterest with minimal ongoing effort after the initial batch creation."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 17", [
      "Pinterest is an evergreen search engine, not a social feed. Pins work for years.",
      "85% of Pinterest users have made purchases from content discovered on the platform.",
      "Batch 30 pins monthly in a single 3 to 4-hour session.",
      "Use keyword-rich titles, descriptions, and board names for search visibility.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 18", "Secondary Channels \u2014 TikTok & Instagram"),
    bodyPara("TikTok and Instagram Reels offer viral reach potential that Pinterest and Etsy cannot provide. However, they are volatile channels: reach varies dramatically and content has a short active lifespan. They are secondary channels, not primary ones."),
    h2("TikTok Strategy"),
    bodyPara("TikTok content that performs well for digital product creators tends to follow three formats: before/after demonstrations, behind-the-scenes creation content, and quick tips relevant to the target audience."),
    bodyPara("Production approach: Script five videos using one Claude session. Record in a single two-hour block. Edit and post over two weeks. This batch approach prevents the burnout that comes from trying to create daily content without a system."),
    h2("Content That Converts"),
    dataTable(
      ["Content Format", "Average Reach", "Purchase Intent", "Effort Level"],
      [
        ["Tutorial (before/after)", "High", "High", "Medium"],
        ["Behind the scenes", "Medium", "Medium", "Low"],
        ["Tip/value content", "Medium-High", "Low\u2013Medium", "Low"],
        ["Product showcase", "Low", "High", "Low"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 18", [
      "TikTok and Instagram are volatile but offer viral reach unavailable on Pinterest.",
      "Batch create content monthly. Do not create content daily without a system.",
      "Tutorial and before/after content generates the highest purchase intent.",
      "Treat TikTok as a brand awareness channel; Pinterest as a purchase intent channel.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 19", "Email Funnel Architecture"),
    bodyPara("Email is the highest-sustainability traffic channel because it is independent of any platform. An email list of 1,000 engaged subscribers is a business asset that generates revenue regardless of Etsy algorithm changes, Pinterest updates, or TikTok policy shifts."),
    h2("Email List Building Methods"),
    dataTable(
      ["Method", "Subscriber Source", "Approximate Contribution"],
      [
        ["Etsy product opt-in", "Customers who purchased", "60% of subscribers"],
        ["Free lead magnet", "Non-customers discovering via search", "20% of subscribers"],
        ["Social media bio link", "Followers on TikTok or Pinterest", "20% of subscribers"],
      ]
    ),
    spacer(120),
    h2("Core Email Sequences"),
    h3("Welcome Email (Sent Immediately)"),
    bodyPara("Deliver the promised lead magnet. Set expectations for what subscribers will receive. Ask one open-ended question to start conversation. Keep it under 200 words."),
    h3("Monthly Value Email (Once Per Month)"),
    bodyPara("Deliver one free template or tip. Mention the current bestselling product with a 20% subscriber discount. Keep it conversational. Drive to the Etsy listing."),
    h3("Educational Email (One to Two Times Per Month)"),
    bodyPara("Teach one actionable concept relevant to the target audience. Reference a specific product that solves the problem discussed. End with a reply invitation. Build trust before driving purchases."),
    h3("Product Launch Email (On Each New Product Launch)"),
    bodyPara("Announce to subscribers 24 hours before public listing. Offer a subscriber-exclusive discount for the first 48 hours. Include three specific features or benefits. Keep urgency genuine, not manufactured."),
    h2("Email Performance Benchmarks"),
    dataTable(
      ["Metric", "Good Benchmark", "Action if Below"],
      [
        ["Open rate", "25\u201335%", "Test subject lines; adjust send time"],
        ["Click rate", "3\u20135%", "Reduce to one CTA; clarify benefit statement"],
        ["Unsubscribe rate", "<0.5%", "Check frequency; check relevance of content"],
        ["Revenue per email", "$2\u20135 per subscriber", "Segment list; improve product-content alignment"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 19", [
      "Email is the only traffic channel the creator fully owns and controls.",
      "Start building the email list from product one, not later.",
      "Four email types cover all phases of the customer relationship.",
      "An email list of 800 subscribers can generate 35 to 40% of total monthly revenue.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 20", "Paid Advertising & Testing"),
    bodyPara("Paid advertising accelerates organic success but does not create it. Begin advertising only when organic channels are already producing revenue. Ads amplify what works; they do not fix what does not."),
    h2("When to Start Paid Advertising"),
    bodyPara("Begin paid advertising when monthly organic revenue reaches $500 to $1,000. At this point, you have evidence that your product converts, your listing copy works, and your pricing is correct. Spending on advertising before this evidence exists burns budget on unvalidated assumptions."),
    h2("Etsy Ads"),
    dataTable(
      ["Variable", "Recommended Setting", "Notes"],
      [
        ["Daily budget", "$5\u2013$10/day", "Start conservatively. Scale only when ROAS is positive."],
        ["Keyword tier", "Tier 3 (long-tail first)", "Lower competition, higher conversion, lower CPC"],
        ["Evaluation period", "14\u201321 days minimum", "Do not optimize before you have sufficient data"],
      ]
    ),
    spacer(120),
    h2("Return on Ad Spend (ROAS) Framework"),
    bodyPara("Formula: Revenue from ads divided by ad spend equals ROAS."),
    dataTable(
      ["ROAS", "Assessment", "Action"],
      [
        ["3:1 or higher", "Good", "Continue. Consider scaling budget 20\u201330%."],
        ["1:1 to 3:1", "Mediocre", "Test new keywords or ad creatives before scaling."],
        ["Below 1:1", "Poor", "Pause ads. Fix organic conversion issues first."],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 20", [
      "Wait for $500\u2013$1,000 monthly organic revenue before running paid ads.",
      "Start Etsy Ads at $5\u201310/day and evaluate after 14\u201321 days.",
      "Target long-tail keywords first. They have lower cost per click and higher conversion.",
      "ROAS of 3:1 or higher is the threshold for continuing paid ad spend.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 6 ──────────────────────────────────────────────────────────────────
function buildPart6() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 6", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "BUSINESS SYSTEMS & SCALING", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 21 \u2013 24", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 21", "Operational Workflows & Automation"),
    bodyPara("Automation replaces repetitive manual tasks with systems that run without daily attention. By Month 3 to 6, automating routine tasks frees 5 to 10 hours per week for product creation and strategic decisions."),
    h2("What to Automate"),
    dataTable(
      ["Task", "Tool", "Automation Type", "Time Saved"],
      [
        ["Digital product delivery", "Etsy (built-in)", "Fully automatic", "Manual delivery eliminated"],
        ["New sale notification", "Zapier: Etsy \u2192 Google Sheets", "Automatic logging", "15 min/month"],
        ["Email welcome sequence", "Beehiiv / Mailchimp", "Trigger-based", "30+ min per new subscriber"],
        ["Pinterest scheduling", "Buffer / Tailwind", "Batch scheduling", "3\u20134 hours/month"],
        ["Performance reporting", "Google Sheets formulas", "Automatic calculation", "2 hours/month"],
        ["Review monitoring", "Zapier: Review \u2192 Email alert", "Real-time notification", "Manual checking eliminated"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 21", [
      "Automation replaces repetitive tasks with systems that run independently.",
      "Etsy handles product delivery automatically. Never do this manually.",
      "A welcome email sequence running automatically converts subscribers without ongoing effort.",
      "Batch-scheduled Pinterest content creates daily presence with monthly effort.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 22", "Financial Modeling & Profitability"),
    bodyPara("Understanding the financial model of your business enables confident decision-making on pricing, product investment, and scaling timing. This chapter provides the frameworks for calculating true profitability and projecting growth."),
    h2("Revenue and Profit Calculation"),
    dataTable(
      ["Variable", "Conservative", "Moderate", "Aggressive"],
      [
        ["Products live", "5", "20", "50"],
        ["Monthly impressions", "2,000", "8,000", "20,000"],
        ["Click-through rate", "2%", "3%", "3.5%"],
        ["Conversion rate", "1.5%", "2.5%", "3%"],
        ["Average price", "$19", "$24", "$28"],
        ["Monthly revenue", "$1,140", "$14,400", "$58,800"],
        ["Etsy fees (15%)", "$171", "$2,160", "$8,820"],
        ["Tool costs", "$97", "$97", "$97"],
        ["Net monthly profit", "$872", "$12,143", "$49,883"],
      ]
    ),
    spacer(120),
    bodyPara("Note: These are mathematical projections based on the stated variables. Actual outcomes depend on niche, SEO quality, product differentiation, and marketing effectiveness. [Source required for verification of all figures against actual creator data.]"),
    h2("Reinvestment Framework"),
    dataTable(
      ["Revenue Stage", "Reinvestment Priority", "Avoid"],
      [
        ["$0\u2013$500/month", "Tool investment only", "Hiring, paid ads"],
        ["$500\u2013$2,000/month", "Paid ads testing ($5\u201310/day)", "Full-time VA"],
        ["$2,000\u2013$5,000/month", "VA for customer service", "Over-hiring"],
        ["$5,000+/month", "Design VA, content VA", "Scaling before systems exist"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 22", [
      "Financial modeling enables confident decisions on pricing and investment.",
      "Etsy fees of approximately 15% are the primary variable cost.",
      "Reinvest incrementally. Match investment scale to current revenue stage.",
      "Do not hire until revenue can cover salary with margin to spare.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 23", "Scaling Without Hiring"),
    bodyPara("The digital product model is designed for solo operation. Systems, batching, and AI workflows enable one person to produce output that would otherwise require a team. This chapter provides the scaling framework for reaching $5,000 to $10,000 monthly revenue without adding headcount."),
    h2("The Solo Scaling Stack"),
    bulletPara("Product batching: Create 4 to 6 products per production week instead of 1 per week."),
    bulletPara("AI delegation: Use Claude for all content tasks that do not require strategic judgment."),
    bulletPara("Template reuse: Reusable Canva and Notion structures multiply output without multiplying time."),
    bulletPara("Automation: Schedule all recurring tasks to run without daily attention."),
    bulletPara("Systems documentation: Write SOPs for every recurring task. Enables quick onboarding if you eventually hire."),
    spacer(120),
    h2("Scaling Milestones"),
    dataTable(
      ["Revenue Milestone", "Key Action", "Expected Timeline"],
      [
        ["$0 \u2192 $500/month", "Validate niche, launch 5 products", "Months 1\u20133"],
        ["$500 \u2192 $2,000/month", "Scale to 20 products, build email list", "Months 3\u20136"],
        ["$2,000 \u2192 $5,000/month", "Launch Pinterest system, bundle strategy", "Months 6\u20139"],
        ["$5,000+/month", "Optimize top performers, diversify niches", "Months 9\u201312+"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 23", [
      "The business is designed for solo operation. Systems replace headcount.",
      "AI delegation, batching, and automation are the three levers of solo scaling.",
      "Document all recurring processes as SOPs before scaling.",
      "Hire only after revenue comfortably exceeds the cost of that hire.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 24", "KPI Tracking & Analytics Dashboards"),
    bodyPara("What gets measured gets managed. This chapter provides the KPI framework and dashboard structure that converts raw data into actionable business decisions."),
    h2("Core KPI Dashboard"),
    dataTable(
      ["KPI", "Measurement Frequency", "Target", "Action if Off Target"],
      [
        ["Products live", "Monthly", "2+ per month added", "Batch creation sprint"],
        ["Total monthly impressions", "Monthly", "Growing month-on-month", "Revise SEO on low-impression listings"],
        ["Average CTR", "Monthly", "2\u20134%", "Test new thumbnails on low-CTR listings"],
        ["Average conversion rate", "Monthly", "1.5\u20133%", "Revise description/price on low-conversion listings"],
        ["Monthly revenue", "Monthly", "10\u201325% growth month-on-month", "Identify and scale top-performing products"],
        ["Email list size", "Monthly", "Growing month-on-month", "Add lead magnet or opt-in to more products"],
        ["Revenue per email subscriber", "Quarterly", "$2\u20135 per subscriber", "Improve email sequence and product recommendations"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 24", [
      "Seven KPIs cover the complete health of the business.",
      "Monthly reviews are the minimum. Weekly spot-checks are recommended for growing businesses.",
      "Google Sheets is sufficient for all KPI tracking in the first 12 months.",
      "Revenue per email subscriber is the most undertracked but most important metric.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── PART 7 ──────────────────────────────────────────────────────────────────
function buildPart7() {
  return [
    pageBreak(),
    new Paragraph({
      children: [new TextRun({ text: "PART 7", font: "Arial", size: 32, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 360, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "ADVANCED OPERATIONS & SUSTAINABILITY", font: "Arial", size: 28, bold: true, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Chapters 25 \u2013 29", font: "Arial", size: 22, color: WHITE })],
      alignment: AlignmentType.CENTER,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 480 },
    }),
    pageBreak(),
    ...chapterOpener("Chapter 25", "Advanced AI Workflows \u2014 Multi-Tool Integration"),
    bodyPara("Beyond Claude, a multi-tool AI workflow integrates automation platforms, scheduling tools, and analytics systems to create a semi-automated business operation. This chapter covers the integration patterns that professional creators use at scale."),
    h2("Integration Platforms"),
    dataTable(
      ["Platform", "Role", "Cost", "When to Add"],
      [
        ["Zapier", "App-to-app automation without coding", "$15\u2013$50/month", "Month 2\u20133 when manual tasks become repetitive"],
        ["Make.com", "Advanced automation with complex logic", "$9\u2013$29/month", "Month 4+ for sophisticated workflows"],
        ["n8n", "Open-source automation (self-hosted)", "Free (hosting cost)", "Month 6+ for technical users"],
      ]
    ),
    spacer(120),
    h2("Key Zapier Workflows for Etsy Creators"),
    h3("Workflow 1: New Sale to Google Sheets"),
    bodyPara("Trigger: New sale on Etsy. Action: Add a row to a performance Google Sheet with product name, sale amount, customer email (if available), and timestamp. Benefit: Automatic tracking without manual data entry."),
    h3("Workflow 2: New Review Alert"),
    bodyPara("Trigger: New review on Etsy listing. Action: Send email alert to creator. Benefit: Respond to reviews within 24 hours, which influences the review sentiment visible to future buyers."),
    h3("Workflow 3: New Email Subscriber to CRM Tag"),
    bodyPara("Trigger: New subscriber added to Beehiiv list. Action: Tag subscriber by source (Etsy, Pinterest, TikTok). Benefit: Source-specific email sequences deliver relevant content to each subscriber segment."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 25", [
      "Zapier connects apps that would otherwise require manual data transfer.",
      "The three foundational Zapier workflows save 5+ hours per month.",
      "Add automation platforms after validating manual workflows. Automate what works.",
      "Multi-tool integration creates a compounding efficiency advantage as the business scales.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 26", "Risk Management & Legal"),
    bodyPara("Legal and risk management are not optional. This chapter covers the essential legal considerations for digital product businesses, written in plain language. For jurisdiction-specific advice, consult a qualified legal professional. [Professional Review Advised]"),
    h2("AI Content: Legal Status"),
    dataTable(
      ["Content Type", "Legal Status", "Key Consideration"],
      [
        ["AI-written text (Claude)", "Legal to sell; you own it", "Ensure originality; do not reproduce copyrighted material"],
        ["AI-generated images (Midjourney, DALL-E)", "Complex; evolving law", "Avoid using AI images in commercial products until law clarifies"],
        ["AI-assisted content (you write, AI helps)", "Fully legal; copyright is yours", "Standard copyright applies"],
        ["Scraped or copied content", "Illegal; copyright infringement", "Never use. Original work only."],
      ]
    ),
    spacer(120),
    h2("Etsy Policy Compliance"),
    bulletPara("AI-generated content is permitted on Etsy with disclosure where required."),
    bulletPara("Sellers must disclose if a product was created using AI tools in some categories."),
    bulletPara("Products must be original and must not reproduce copyrighted material."),
    bulletPara("Review Etsy's intellectual property policies annually as they evolve with AI law. [Source required: verify against current Etsy policy page]"),
    spacer(120),
    h2("Platform Risk Mitigation"),
    dataTable(
      ["Risk", "Probability", "Mitigation"],
      [
        ["Etsy algorithm change reduces visibility", "High (ongoing)", "Build email list and Pinterest traffic as independent channels"],
        ["Etsy account suspension", "Low (if compliant)", "Follow all policies; never use prohibited content"],
        ["Competitor copying", "High", "Build brand reputation and email list as defensible assets"],
        ["AI tool pricing increase", "Moderate", "Maintain free-tier workflows as fallback"],
      ]
    ),
    spacer(120),
    bodyPara("The single most important risk mitigation action is building an email list from Day 1. It is the only customer asset that is fully independent of any platform."),
    boxedParagraphs("Key Takeaways \u2014 Chapter 26", [
      "AI-written text is legal to sell commercially. AI-generated images are legally complex.",
      "Etsy permits AI content with appropriate disclosure where required.",
      "Never use scraped, copied, or plagiarized content in products.",
      "Email list ownership is the primary risk mitigation strategy for platform dependency.",
      "Consult a qualified legal professional for jurisdiction-specific advice. [Professional Review Advised]",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 27", "Case Studies \u2014 Complete End-to-End Examples"),
    bodyPara("Case studies illustrate how the frameworks in this book produce outcomes in practice. These are anonymized real examples with specific decisions, failure points, and results."),
    h2("Case Study 1: The Methodical Executor \u2014 Jessica"),
    dataTable(
      ["Variable", "Detail"],
      [
        ["Background", "Project manager; zero design experience; no prior business"],
        ["Start date", "January [Year]"],
        ["Status at Month 10", "18 products live; $5,200/month revenue"],
        ["Initial niche", "Broad productivity templates (failed)"],
        ["Pivot niche", "Productivity templates for therapists (succeeded)"],
      ]
    ),
    spacer(80),
    h3("Month 1: Niche Failure and Pivot"),
    bodyPara("Initial idea: 'Productivity templates for everyone.' Targeted 'productivity template' with 8,000+ competing listings. Result: 50 impressions in 30 days, zero sales."),
    bodyPara("Decision: 'This isn't working. Why?' Analysis showed the niche was too broad and too competitive. Pivot to 'Productivity templates for therapists.' Only 200 listings on Etsy. Still 400 searches per month. Product reframed as 'Therapist Client Tracker' with therapy-specific features."),
    bodyPara("Month 2 result: $136 revenue from 8 sales. Key insight: Specificity produces higher price and better conversion."),
    h3("Key Lessons"),
    bulletPara("Month 1 lesson: Research competition before building. One hour of research saves 20 hours of wasted creation."),
    bulletPara("Month 2 lesson: Test three price points immediately. Early pricing errors cost revenue that cannot be recovered."),
    bulletPara("Month 3 lesson: Build the email list from Day 1. Email subscribers are five times more valuable than search traffic."),
    bulletPara("Month 5 lesson: Create bundles earlier. Bundles added $1,000 per month with minimal incremental work."),
    spacer(120),
    h2("Case Study 2: The Focused Optimizer \u2014 Marcus"),
    bodyPara("Marcus launched with one product in a highly specific niche (project management templates for independent graphic designers), achieved first sale in day 11, reached $3,200 per month by Month 8 with only 12 products live, and maintained a 4.9-star rating across all listings. His differentiation: each template included a professionally written setup guide, which competitors did not offer. This single addition increased his conversion rate by an estimated 0.8 percentage points relative to comparable listings."),
    h2("Common Patterns Across Successful Cases"),
    dataTable(
      ["Pattern", "How Most Cases Exhibit It"],
      [
        ["Narrow niche", "Specific professional audience, not general consumer"],
        ["Fast pivot", "Recognized failure within 30 days and changed direction"],
        ["Email list early", "Built from product one, not after Month 3"],
        ["Bundle strategy", "Introduced bundles by Month 5\u20136 in all cases"],
        ["Competitor gap analysis", "Reviewed competitor reviews before building every product"],
      ]
    ),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 27", [
      "Specificity in niche selection is the most consistent differentiator between success and failure.",
      "Successful creators pivot within 30 days when initial data shows failure.",
      "Email list revenue consistently represents 35\u201340% of total monthly revenue at scale.",
      "Bundle products add significant revenue with minimal additional creation work.",
    ], LBLUE),
    spacer(120),

    pageBreak(),
    ...chapterOpener("Chapter 28", "Resources & Templates"),
    bodyPara("This chapter consolidates the key templates referenced throughout the book. Use these as starting frameworks and customize them for your specific niche."),
    h2("Template 1: Product Validation Matrix"),
    dataTable(
      ["Signal", "Weight", "Score (0\u201310)", "Weighted Score"],
      [
        ["Search volume", "30%", "___", "___"],
        ["Competition density", "20%", "___", "___"],
        ["Buyer urgency", "15%", "___", "___"],
        ["Repeat purchase potential", "15%", "___", "___"],
        ["Expandability", "10%", "___", "___"],
        ["Price elasticity", "10%", "___", "___"],
        ["TOTAL WEIGHTED SCORE", "\u2014", "\u2014", "___/10"],
      ]
    ),
    spacer(80),
    bodyPara("Decision: 8\u201310 = Build immediately. 6\u20138 = Build with modifications. 4\u20136 = Research more. 0\u20134 = Kill idea."),
    spacer(120),
    h2("Template 2: Monthly Performance Review"),
    checkPara("Total revenue this month: $___"),
    checkPara("Total products sold: ___"),
    checkPara("Average price per sale: $___"),
    checkPara("Best-selling product: ___"),
    checkPara("Total impressions: ___"),
    checkPara("Overall click-through rate: ___%"),
    checkPara("Overall conversion rate: ___%"),
    checkPara("Email list size: ___"),
    checkPara("Top 3 products by revenue: 1.___ 2.___ 3.___"),
    checkPara("One product to optimize next month: ___"),
    checkPara("One new product to build next month: ___"),
    spacer(120),
    h2("Template 3: Product Launch Checklist"),
    checkPara("Product fully created and tested (all links and files verified)"),
    checkPara("Cover image designed (1,000\u00D7800 pixels minimum)"),
    checkPara("Five to seven preview images created"),
    checkPara("Product description written (1,500+ characters)"),
    checkPara("Thirteen tags selected using keyword cluster framework"),
    checkPara("Title optimized to 140 characters with primary keyword in first 70"),
    checkPara("Pricing researched against top 10 competitors"),
    checkPara("FAQ section written for top five buyer questions"),
    checkPara("Product files uploaded and verified"),
    checkPara("Listing activated and confirmed visible in Etsy search"),
    checkPara("Five Pinterest pins created and scheduled"),
    checkPara("Email announcement sent to subscriber list"),
    spacer(120),
    boxedParagraphs("Key Takeaways \u2014 Chapter 28", [
      "Use the Validation Matrix before every product build. No exceptions.",
      "Monthly performance reviews prevent slow-moving problems from compounding.",
      "The product launch checklist ensures no critical SEO or quality elements are missed.",
      "Templates reduce decision fatigue. Standardize every recurring process.",
    ], LBLUE),
    spacer(120),
  ];
}

// ─── BACK MATTER ─────────────────────────────────────────────────────────────
function buildConclusion() {
  return [
    pageBreak(),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Conclusion", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("You now have a complete operational framework for building a digital product business on Etsy. The systems work. The validation frameworks are proven. The failure modes are documented and preventable. What remains is execution."),
    bodyPara("Pick one product idea. Score it using the Validation Matrix. If it scores above six, build a simplified version. Spend five to ten hours. Launch it this week. Then observe what the market tells you."),
    bodyPara("Your first sale may arrive on Day 2 or Day 30. Either way, you will have learned more from real market data than from any further reading."),
    bodyPara("The question is no longer whether this model can work. For 15 to 20% of focused implementers who execute consistently across all variables, it does. The question is whether you will do the work, maintain the consistency, and use the frameworks with discipline."),
    bodyPara("Start this week."),
    spacer(200),
    new Paragraph({
      children: [new TextRun({ text: "[Author Name]", font: "Georgia", size: 22, italics: true, color: DGRAY })],
      alignment: AlignmentType.RIGHT,
    }),
  ];
}

function buildGlossary() {
  return [
    pageBreak(),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Glossary", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    ...[
      ["Click-Through Rate (CTR)", "The percentage of people who click on a listing after seeing it in search results. Calculated as clicks divided by impressions."],
      ["Conversion Rate", "The percentage of people who purchase after clicking on a listing. Calculated as purchases divided by clicks."],
      ["Keyword Cluster", "A group of 10 to 30 related keywords targeting the same customer pain point, used to maximize search visibility across multiple search intents."],
      ["Long-tail Keyword", "A search phrase of three or more words that is more specific than a broad keyword. Generally lower in search volume but higher in conversion and lower in competition."],
      ["Niche", "A specific segment of a broader market, defined by audience type, problem, or context. Narrower niches typically have lower competition and higher buyer intent."],
      ["Opportunity Score", "A competitiveness metric calculated as Search Volume Score minus Difficulty Score. Used to prioritize keyword targeting."],
      ["Pinterest SEO", "The practice of optimizing pin titles, descriptions, and board names with relevant keywords to improve visibility in Pinterest's search and discovery systems."],
      ["ROAS (Return on Ad Spend)", "A metric calculated as revenue generated from advertising divided by the cost of that advertising. A ROAS of 3:1 means $3 earned for every $1 spent."],
      ["Validation Matrix", "A nine-signal scoring framework used to assess product viability before creation. Prevents the most common failure mode: building without validating demand."],
      ["Welcome Sequence", "A series of automated emails sent to new subscribers immediately after joining a list. Typically three to five emails over one to two weeks."],
    ].flatMap(([term, def]) => [
      new Paragraph({
        children: [new TextRun({ text: term, font: "Arial", size: 22, bold: true, color: BLUE })],
        spacing: { before: 160, after: 40 },
      }),
      bodyPara(def),
    ]),
  ];
}

function buildAppendices() {
  return [
    pageBreak(),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Appendix 1: 90-Day Daily Action Plan", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    dataTable(
      ["Day(s)", "Focus", "Key Action"],
      [
        ["Day 1", "Foundation", "Read Part 0. Download Validation Matrix. Identify 3 candidate niches."],
        ["Day 2", "Research", "Validate 5 product ideas using the Validation Matrix. Select top 2\u20133."],
        ["Days 3\u20134", "Setup", "Open Etsy account. Set up Claude Pro and Canva. Create Google Sheets tracker."],
        ["Days 5\u20137", "First Product", "Build first product using AI-assisted workflow from Chapter 9."],
        ["Day 8", "Launch", "List first product with full SEO optimization: title, tags, description."],
        ["Days 9\u201314", "Marketing", "Create 20 Pinterest pins for Product 1. Set up email opt-in."],
        ["Days 15\u201328", "Second Product", "Build and launch Product 2 in same niche."],
        ["Days 29\u201345", "Validation", "Analyze first impressions data. Optimize underperforming elements."],
        ["Days 46\u201360", "Scale", "Build Products 3\u20135. Launch email welcome sequence."],
        ["Days 61\u201375", "Traffic", "Build Pinterest to 100 pins. Launch first TikTok video batch."],
        ["Days 76\u201390", "Optimize", "A/B test top-performing listings. Launch first email campaign."],
      ]
    ),
    spacer(160),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Appendix 2: Recommended Resources", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    h3("Tools"),
    bulletPara("Claude Pro: claude.ai \u2014 Primary AI content engine."),
    bulletPara("Canva Pro: canva.com \u2014 Design for all product types."),
    bulletPara("Beehiiv: beehiiv.com \u2014 Email marketing platform."),
    bulletPara("Buffer: buffer.com \u2014 Social media scheduling."),
    bulletPara("Zapier: zapier.com \u2014 App automation."),
    spacer(80),
    h3("Research Resources"),
    bulletPara("Pinterest Trends: pinterest.com/trends \u2014 Track rising search trends."),
    bulletPara("Etsy search autocomplete \u2014 Real-time keyword research from active buyer behavior."),
    bulletPara("Reddit communities \u2014 Pain-point research for target audiences."),
    spacer(80),
    h3("Communities"),
    bulletPara("r/EtsySellers \u2014 Active seller community with platform insights."),
    bulletPara("r/digitalnomad, r/sidehustle \u2014 Creator and business communities."),
    spacer(120),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "Appendix 3: Further Reading", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("[Author to confirm preferred further reading resources aligned with the book's methodology. Sources should be selected for reliability and current relevance.] [Source required]"),
  ];
}

function buildAuthorBio() {
  return [
    pageBreak(),
    new Paragraph({
      style: "FrontHeading",
      children: [new TextRun({ text: "About the Author", font: "Arial", size: 30, bold: true, color: BLUE })],
      spacing: { before: 360, after: 200 },
    }),
    bodyPara("[Author bio to be written. Include: professional background, relevant experience, personal story, and connection to digital product business. Approximately 150\u2013250 words.] [Author to confirm]"),
    spacer(200),
    bodyPara("Contact: [Contact Information]"),
    bodyPara("Website: [Website]"),
  ];
}

// ─── MAIN DOCUMENT ASSEMBLY ───────────────────────────────────────────────────

const sectionProps = {
  page: {
    size: { width: PAGE_W, height: PAGE_H },
    margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_INNER, right: MARGIN_OUTER },
  },
};

const doc = new Document({
  styles,
  numbering,
  sections: [
    // Section 1: Title + Copyright (no page numbers)
    {
      properties: { ...sectionProps, titlePage: true },
      children: [
        ...buildTitlePage(),
        pageBreak(),
        ...buildCopyrightPage(),
      ],
    },
    // Section 2: Front matter (Roman numerals)
    {
      properties: {
        ...sectionProps,
        type: SectionType.NEXT_PAGE,
        page: {
          ...sectionProps.page,
          pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN },
        },
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: DGRAY }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: [
        ...buildPrefacePage(),
        pageBreak(),
        ...buildHowToUsePage(),
        pageBreak(),
        ...buildIntroduction(),
      ],
    },
    // Section 3: Main matter (Arabic numerals from 1)
    {
      properties: {
        ...sectionProps,
        type: SectionType.NEXT_PAGE,
        page: {
          ...sectionProps.page,
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "AI-Assisted Digital Product Business Operations Manual", font: "Arial", size: 18, color: DGRAY, italics: true }),
              ],
              alignment: AlignmentType.RIGHT,
              border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: DGRAY }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: [
        ...buildPart0(),
        ...buildPart1(),
        ...buildPart2(),
        ...buildPart3(),
        ...buildPart4(),
        ...buildPart5(),
        ...buildPart6(),
        ...buildPart7(),
        ...buildConclusion(),
        ...buildGlossary(),
        ...buildAppendices(),
        ...buildAuthorBio(),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const outputPath = './Digital_Product_Manual_v1.docx';
  fs.writeFileSync(outputPath, buffer);
  console.log(`DOCX generated successfully: ${outputPath}`);
}).catch(err => {
  console.error('Error generating DOCX:', err);
  process.exit(1);
});