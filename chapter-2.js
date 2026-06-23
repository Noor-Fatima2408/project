const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, 
        HeadingLevel, BorderStyle, WidthType, ShadingType } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "0F0E0F" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "1F2937" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 }
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "374151" },
        paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: "bullet",
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      },
      {
        reference: "numbers",
        levels: [
          {
            level: 0,
            format: "decimal",
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("Understanding Claude AI")]
        }),
        new Paragraph({ spacing: { after: 240 }, children: [new TextRun("")] }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("If you're building a digital product business, your choice of AI tool matters more than you might think. The wrong tool will slow you down, produce mediocre output, and cost you money. The right one becomes your unfair advantage—a co-founder who works at machine speed while you focus on strategy and execution.")]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("This chapter isn't about which AI is best in a vacuum. It's about which AI is best for your business. And if you're running a digital product business—especially on platforms like Etsy—the answer is Claude.")]
        }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("What Makes Claude Different")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude is built by Anthropic with a specific philosophy: prioritize accuracy, reasoning, and usefulness for business and technical work. That's not marketing speak. That's the difference you'll feel in every interaction.")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 1: Claude vs ChatGPT Comparison Dashboard]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1920x1080px professional dashboard image here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("The comparison below captures the key differences. But here's what matters most: while ChatGPT is a conversational tool that's great for exploration, Claude is built for business output. It's the difference between a helpful chatbot and a business partner.")]
        }),
        createComparisonTable(),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Why Claude Is Objectively Better for Digital Products")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 160 },
          children: [new TextRun("Longer context window = paste entire competitive landscape for analysis. With 200,000 tokens, you can dump your entire competitor research, all your product data, and months of customer feedback into a single conversation. Claude holds all of it. ChatGPT's 128,000 tokens looks good on paper until you realize you're splicing data across multiple conversations.")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 160 },
          children: [new TextRun("Better structured output = organized results ready to use. You ask for JSON, you get clean JSON. You ask for a bulleted list, you get formatting that imports cleanly to spreadsheets. This might sound minor until you're copy-pasting data from chatbot responses into your production systems daily.")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 160 },
          children: [new TextRun("Business focus = Claude understands ROI, conversion rates, customer lifetime value, and revenue metrics. It's not just an API question—it's a philosophy embedded in how Claude responds. When you ask about pricing strategy, Claude doesn't just offer options. It reasons through your margins, competitive positioning, and customer psychology.")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Deeper reasoning = multi-step problems get solved in one conversation. Complex problems—like 'why are my customers leaving?' or 'how should I structure my marketing funnel?'—require chaining logical steps. Claude's reasoning capability means fewer back-and-forths and faster problem resolution.")]
        }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Four Core Strengths That Drive Your Digital Product Business")]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Understanding Claude's technical capabilities is interesting. Using them to grow your revenue is what matters. Here are the four strengths that create the biggest ROI for digital product entrepreneurs.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("1. Human-Quality Content Generation at Scale")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude produces publishable content across every format your business needs. Not almost publishable. Publishable. This is different from earlier-generation AI tools that generated content that sounded like it was written by a robot that learned English from Wikipedia.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("The formats you'll use most frequently:")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Product descriptions (SEO-optimized, conversion-focused)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Email sequences (personality-driven, audience-specific)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Blog posts (10,000+ words without degradation in quality)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Social media variations (multiple versions for A/B testing)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Ebook content (full guides that read like professional publications)")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Real Example: From Blank Page to 10 Variations")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You're selling a Notion project template for small teams. You need 10 title variations optimized for Etsy search. Normally, you'd spend an hour researching keywords, testing combinations, and validating against search trends.")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You ask Claude: 'Generate 10 title variations optimized for Etsy search for a project management Notion template.'")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Two minutes later, Claude delivers:")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("\"Project Management Notion Template | Team Collaboration System\"")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("\"Notion Project Manager for Small Teams | Task Tracking Template\"")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("\"Team Project Tracker in Notion | Workflow Automation Template\"")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("... (7 more variations, each researched and test-ready)")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 2: Content Generation Speed & Quality Timeline]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1600x900px timeline infographic here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Quality: Professional, tested variations. Time taken: 2 minutes. Your time saved: 30 minutes. This is the pattern you'll repeat hundreds of times. Content generation that would normally require outsourcing or significant time investment becomes a 2-3 minute task.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("2. Research & Competitive Analysis")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Understanding your market is non-negotiable. But research at scale—analyzing 50 competitor reviews, mapping feature gaps, identifying market trends—typically requires a team or weeks of individual effort.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude compresses this into hours. Here's what it synthesizes:")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Review synthesis (summarize 50 competitor reviews into 5 key themes)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Competitor feature mapping (what they offer, what they're missing)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Trend analysis (seasonal patterns, emerging customer needs)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Data processing (organize raw customer feedback into actionable categories)")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Real Example: From Complaints to Product Improvements")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You paste 20 Etsy product reviews from a competitor's listing. You ask Claude: 'What are the top 5 complaints? Top 5 wishes? What would make customers genuinely happy?'")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 3: Competitor Analysis Process Flow]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1600x1000px competitor analysis flow diagram here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude returns:")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Top Complaints")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Confusing setup (5 mentions)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Missing templates (7 mentions)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("No video tutorial (4 mentions)")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Customer Wishes")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Mobile version (6 mentions)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("More templates and examples (9 mentions)")]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Now you have your product roadmap. You don't need to guess. You have market validation. Your time saved: 45 minutes of reading and analysis.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("3. Structured Output for Systems Integration")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Digital product businesses run on data. Your Etsy listings need tags. Your email campaigns need audience segments. Your marketing needs keyword research. Claude outputs data in exactly the format you need: JSON, CSV, spreadsheet-ready tables, structured queries.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You ask: 'I need 50 Etsy tags for my Notion template. Format as JSON with: tag, search_volume_estimate, competition_level.'")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 5: JSON Output Structure Example with Etsy Tags]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1600x900px JSON code editor mockup here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude delivers perfectly formatted JSON that you import directly into your spreadsheet, your marketing automation tool, or your product database. No reformatting. No wasted time. Just data that works.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("4. Multi-Step Problem-Solving")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("The hardest problems don't have one-sentence answers. Why aren't your products converting? How should you price your template? What's preventing customer retention? These problems require reasoning across multiple connected ideas.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude's deeper reasoning capability means you get diagnostic thinking, not just suggestions.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Real Example: Why Your Notion Template Isn't Selling")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You're frustrated. Your Notion template is good, but it's not selling. You provide Claude with your actual data: 500 impressions, 15 clicks, 0 sales. Your price: $32. Competitors: $17–24. Your description: 50 words.")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 6: Problem-Solving Analysis Dashboard]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1800x1100px diagnostic dashboard visualization here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude analyzes:")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("CTR is 3%—that's actually good. Means your Etsy listing is visible and people are clicking.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Conversion is 0%—that's the problem. People click but don't buy.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Likely cause: Price too high + weak description (customers click expecting detail, see 50 words, bounce).")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("The Diagnosis")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Lower price to $19 (mid-market positioning, competitive)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Expand description to 200+ words (show value, reduce buyer hesitation)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Add feature list (explicit about what they're getting)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Test for 2 weeks (let Etsy's algorithm pick it up)")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Track conversion improvement (run metrics again, validate solution)")]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("This is strategic thinking. You could have done it yourself in an hour. But Claude compressed it into 2 minutes. Your time saved: 1 hour of strategic thinking.")]
        }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Context Windows & Why They Matter")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("A context window is how much conversation history an AI can remember in a single session. Claude's 200,000-token window is the game-changer most people don't immediately appreciate.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("To put this in perspective: 200,000 tokens equals approximately 150,000 words—roughly the length of a complete novel. ChatGPT's 128,000 tokens equals roughly 96,000 words.")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 4: Context Window Size Comparison with Real-World Examples]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1400x1200px context window comparison visualization here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("What This Means for Your Digital Product Business")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Paste your entire competitive landscape. All 40+ competitor product listings? Paste them. All 50 customer reviews? Paste them. All your product analytics? Paste them. Claude holds it all.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Ask Claude to analyze and identify gaps. Now Claude can see the full picture: what your competitors offer, what customers complain about, where opportunities exist.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Keep context for follow-up questions. You don't lose context mid-project. Ask a follow-up question, and Claude still remembers the competitive analysis from 10 messages ago.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Build on previous conversations. One conversation becomes the foundation for the next. Your entire product strategy lives in one evolving document, not fragmented across multiple chat windows.")]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("With ChatGPT's smaller window, you're constantly splicing data across conversations, losing context, and repeating yourself. With Claude, you work with the complete picture.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Claude Artifacts: Dedicated Long-Form Content")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude Artifacts is a feature that opens long-form content in a dedicated, editable window. Instead of reading your ebook or email sequence inside a chat bubble, it appears in a professional editor where you can read it, edit it, and export it immediately.")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Here's why this matters:")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Write full ebooks (10,000+ words) in one shot, visible in a professional editor")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Generate complete templates (fully structured, ready to customize)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Create detailed design briefs (with visual specifications and examples)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Write email sequences (5+ emails, all in one artifact, iterable)")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Real Workflow: From Blank Page to Finished Ebook")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("You: 'Write an ebook: \"Productivity for Remote Teams\" (50 pages, Notion-focused, includes templates and case studies).'")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Claude: Generates full ebook in Artifact window. You can see the structure, formatting, and all content immediately.")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("You: 'Make chapter 3 more practical. Add a step-by-step setup guide.'")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [new TextRun("Claude: Edits just that chapter in the same artifact. You see the changes in real-time.")]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("You: 'Export as PDF.' The ebook downloads, production-ready, 2-hour turnaround from concept to finished product.")]
        }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Claude Projects: Persistent Context for Your Business")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude Projects is an advanced feature that changes how you work with AI. Instead of starting from scratch every conversation, you create a project with persistent instructions, context, and system rules. Claude remembers them. Every conversation stays consistent.")]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("What Projects Do")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Keep context persistent. Your instructions stay saved. You don't re-explain your rules every conversation.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Maintain voice consistency. Claude remembers your brand voice, writing style, and formatting preferences.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Build custom workflows. Special instructions for SEO, content generation, competitor analysis, pricing strategy—whatever you need.")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("Organize your work. Related conversations stay grouped. Your entire marketing strategy, product roadmap, or customer analysis lives in one project.")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 7: Claude Projects Feature - Persistent Context System]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1800x1000px Claude Projects before/after comparison here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun("Real Example: The Etsy SEO Specialist Project")]
        }),
        new Paragraph({
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You create a project called 'Etsy SEO Specialist.' You give it specific instructions:")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[VISUAL 8: Project Instructions Repository - Real-World Example]",
            italics: true,
            color: "666666"
          })]
        }),
        new Paragraph({
          spacing: { after: 280 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "[Insert 1400x1600px detailed system instructions card/document here]",
            color: "999999",
            size: 20
          })]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Now, every time you talk to this project:")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Claude remembers your rules (keywords first, search volume prioritized, specific format requirements)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Claude applies your brand voice (understands you sell Notion templates, your audience, your price point)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Claude formats output consistently (always JSON, always includes search volume estimate)")]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 240 },
          children: [new TextRun("You skip the explaining. You just ask for what you need.")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Instead of: 'Hey Claude, remember I need SEO-optimized titles, 140 characters, keyword in first 40, formatted as JSON, and you should know I'm selling Notion templates for $19 to remote workers...'")]
        }),
        new Paragraph({
          spacing: { after: 240 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("You just: 'Generate 10 title variations for my new collaboration template.'")]
        }),
        new Paragraph({
          spacing: { after: 320 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude already knows all your preferences. It's like having a specialized team member who knows your business inside and out.")]
        }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun("")] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Why This Matters for Your Bottom Line")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("The features we've covered—deeper reasoning, longer context, structured output, persistent projects—aren't interesting because they're technically impressive. They're interesting because they directly impact your revenue.")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Faster content generation = more product listings, more marketing materials, faster testing cycles. Deeper problem-solving = smarter pricing, better product positioning, clearer value propositions. Structured output = seamless integration with your tools, less manual work, more automation.")]
        }),
        new Paragraph({
          spacing: { after: 200 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("In the next chapter, we'll move from understanding Claude to actually using it. You'll learn the prompting techniques that separate mediocre output from exceptional output. You'll see how to turn Claude into your competitive advantage.")]
        }),
        new Paragraph({
          spacing: { after: 100 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("But here's the key insight from this chapter: The AI tool you choose compounds over time. Every day you use the right tool is a day you're building your business faster than your competitors. Every day you use the wrong tool is a day you're at a disadvantage.")]
        }),
        new Paragraph({
          spacing: { after: 0 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun("Claude isn't perfect. But for digital product entrepreneurs, it's the right choice. Now let's learn how to use it.")]
        })
      ]
    }
  ]
});

function createComparisonTable() {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const headerShading = { fill: "2E75B6", type: ShadingType.CLEAR };

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2340, 3510, 3510],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            shading: headerShading,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: "Dimension", color: "FFFFFF", bold: true })] })]
          }),
          new TableCell({
            borders,
            shading: headerShading,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: "Claude", color: "FFFFFF", bold: true })] })]
          }),
          new TableCell({
            borders,
            shading: headerShading,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: "ChatGPT", color: "FFFFFF", bold: true })] })]
          })
        ]
      }),
      createTableDataRow("Context Window", "200,000 tokens", "128,000 tokens", borders),
      createTableDataRow("Reasoning Depth", "Deeper, analytical", "Good, conversational", borders),
      createTableDataRow("Long-form Quality", "Excellent", "Good", borders),
      createTableDataRow("Business Focus", "Intentional", "Secondary", borders),
      createTableDataRow("Output Structure", "Highly organized", "More conversational", borders),
      createTableDataRow("Consistency", "Higher", "Variable", borders),
      createTableDataRow("Code Quality", "Excellent", "Good", borders),
      createTableDataRow("Learning Curve", "Steeper", "Shallow", borders)
    ]
  });
}

function createTableDataRow(dimension, claude, chatgpt, borders) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: dimension, bold: true })] })]
      }),
      new TableCell({
        borders,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun(claude)] })]
      }),
      new TableCell({
        borders,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun(chatgpt)] })]
      })
    ]
  });
}

Packer.toBuffer(doc).then(buffer => {
  const outputPath = "./Chapter_2_Understanding_Claude_AI.docx";
  fs.writeFileSync(outputPath, buffer);
  console.log(`✅ Chapter 2 created successfully: ${outputPath}`);
});