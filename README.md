# Reno-Platform-solution
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, ExternalHyperlink
} = require("docx");
const fs = require("fs");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: "1a1a2e" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: "2d2d7a" })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: "444444" })]
  });
}
function p(...runs) {
  return new Paragraph({
    spacing: { before: 40, after: 80 },
    children: runs.map(r => {
      if (typeof r === "string") return new TextRun({ text: r, font: "Arial", size: 22 });
      return r;
    })
  });
}
function bold(text) {
  return new TextRun({ text, bold: true, font: "Arial", size: 22 });
}
function code(text) {
  return new TextRun({ text, font: "Courier New", size: 20, color: "1a1a2e" });
}
function codeBlock(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    shading: { fill: "f4f4f8", type: ShadingType.CLEAR },
    indent: { left: 360, right: 360 },
    border: {
      top: border, bottom: border, left: { style: BorderStyle.SINGLE, size: 6, color: "2d2d7a" }, right: border
    },
    children: [new TextRun({ text, font: "Courier New", size: 19, color: "1a1a2e" })]
  });
}
function bullet(text, subtext) {
  const runs = [new TextRun({ text, font: "Arial", size: 22 })];
  if (subtext) runs.push(new TextRun({ text: `  ${subtext}`, font: "Arial", size: 20, color: "666666" }));
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    children: runs
  });
}
function sp() { return new Paragraph({ children: [new TextRun("")] }); }
function hr() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCDD", space: 1 } },
    spacing: { before: 120, after: 120 },
    children: [new TextRun("")]
  });
}

function makeTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a,b)=>a+b,0);
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: "2d2d7a", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })]
        }))
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, i) => new TableCell({
          borders,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: ri % 2 === 0 ? "f7f7ff" : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 20 })] })]
        }))
      }))
    ]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1a1a2e" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2d2d7a" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "checks",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2610", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // ── TITLE ──────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Reno Notice Board", bold: true, size: 48, font: "Arial", color: "1a1a2e" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Web Development Internship Assignment", size: 28, font: "Arial", color: "2d2d7a" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: "Live URL: ", bold: true, size: 22, font: "Arial" }),
          new ExternalHyperlink({
            link: "https://reno-noticeboard.vercel.app",
            children: [new TextRun({ text: "reno-noticeboard.vercel.app", color: "2d2d7a", underline: {}, size: 22, font: "Arial" })]
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: "GitHub: ", bold: true, size: 22, font: "Arial" }),
          new ExternalHyperlink({
            link: "https://github.com/vijay0778",
            children: [new TextRun({ text: "github.com/vijay0778", color: "2d2d7a", underline: {}, size: 22, font: "Arial" })]
          })
        ]
      }),
      hr(),

      // ── ABOUT ──────────────────────────────────────────────────────────
      h1("About the Project"),
      p("A full-stack ", bold("Notice Board"), " application built for the Reno Platforms Web Development Internship. It supports full Create, Read, Update, and Delete (CRUD) operations. Urgent notices always appear at the top with a red badge. The app is deployed on Vercel and persists data in a hosted MySQL-compatible database via Prisma."),
      sp(),

      // ── FEATURES ──────────────────────────────────────────────────────
      h1("Features"),
      bullet("Create, Read, Update, Delete notices end-to-end"),
      bullet("Urgent notices always sorted first (done in database query, not browser)"),
      bullet("Red 'Urgent' badge on urgent notices"),
      bullet("Responsive card layout — works on mobile and desktop"),
      bullet("Delete requires confirmation dialog before executing"),
      bullet("Server-side input validation on all API routes (400 on bad input)"),
      bullet("Add/Edit use a single shared form component"),
      bullet("Optional image URL per notice (bonus feature)"),
      sp(),
      hr(),

      // ── TECH STACK ────────────────────────────────────────────────────
      h1("Tech Stack"),
      sp(),
      makeTable(
        ["Layer", "Technology", "Details"],
        [
          ["Framework", "Next.js 14 — Pages Router", "pages/ directory (NOT app/)"],
          ["Database ORM", "Prisma", "Type-safe queries, schema, migrations"],
          ["Database", "TiDB Cloud (MySQL-compatible)", "Free Serverless tier, hosted"],
          ["Styling", "Tailwind CSS", "Responsive, utility-first"],
          ["Hosting", "Vercel (Hobby tier)", "Free, auto-deploy from GitHub"],
        ],
        [2500, 3500, 3360]
      ),
      sp(),
      hr(),

      // ── LOCAL SETUP ───────────────────────────────────────────────────
      h1("How to Run Locally"),
      sp(),
      h2("Prerequisites"),
      bullet("Node.js 18 or higher"),
      bullet("npm or yarn"),
      bullet("A free TiDB Cloud account at tidbcloud.com (no credit card required)"),
      sp(),

      h2("Step 1 — Clone the Repository"),
      codeBlock("git clone https://github.com/vijay0778/reno-noticeboard.git"),
      codeBlock("cd reno-noticeboard"),
      sp(),

      h2("Step 2 — Install Dependencies"),
      codeBlock("npm install"),
      sp(),

      h2("Step 3 — Set Up Environment Variables"),
      p("Copy the example file and fill in your TiDB Cloud connection string:"),
      codeBlock("cp .env.example .env"),
      sp(),
      p("Edit ", code(".env"), " and set:"),
      codeBlock('DATABASE_URL="mysql://USER:PASSWORD@HOST:4000/noticeboard?sslaccept=strict"'),
      sp(),
      p("To get your connection string:"),
      bullet("Go to tidbcloud.com and create a free Serverless cluster"),
      bullet("Click Connect → General → Prisma"),
      bullet("Copy the connection string and paste it in .env"),
      sp(),

      h2("Step 4 — Push Database Schema"),
      codeBlock("npx prisma db push"),
      sp(),

      h2("Step 5 — Start Development Server"),
      codeBlock("npm run dev"),
      sp(),
      p("Open ", code("http://localhost:3000"), " in your browser."),
      sp(),
      hr(),

      // ── PROJECT STRUCTURE ─────────────────────────────────────────────
      h1("Project Structure"),
      sp(),
      codeBlock(
`reno-noticeboard/
├── prisma/
│   └── schema.prisma          # Notice model, enums, DB config
├── pages/
│   ├── api/
│   │   └── notices/
│   │       ├── index.ts       # GET (list) + POST (create)
│   │       └── [id].ts        # PUT (update) + DELETE
│   ├── notices/
│   │   ├── new.tsx            # Add Notice page
│   │   └── [id]/edit.tsx      # Edit Notice page
│   ├── _app.tsx
│   └── index.tsx              # Notice listing (home)
├── components/
│   ├── NoticeCard.tsx         # Card with Edit & Delete buttons
│   ├── NoticeForm.tsx         # Shared Add/Edit form
│   └── ConfirmDialog.tsx      # Delete confirmation modal
├── lib/
│   └── prisma.ts              # Prisma client singleton
├── styles/
│   └── globals.css
├── .env.example               # Safe-to-commit env template
└── README.md`
      ),
      sp(),
      hr(),

      // ── API ──────────────────────────────────────────────────────────
      h1("API Routes"),
      sp(),
      makeTable(
        ["Method", "Route", "Description"],
        [
          ["GET",    "/api/notices",        "Fetch all notices — Urgent first, then by date"],
          ["POST",   "/api/notices",        "Create a new notice — validates on server"],
          ["PUT",    "/api/notices/[id]",   "Update an existing notice — validates on server"],
          ["DELETE", "/api/notices/[id]",   "Delete a notice by ID"],
        ],
        [1800, 3200, 4360]
      ),
      sp(),
      p("All write operations return ", code("400"), " with an error array if validation fails, and ", code("404"), " if the notice is not found."),
      sp(),
      hr(),

      // ── WHAT I WOULD IMPROVE ─────────────────────────────────────────
      h1("One Thing I Would Improve With More Time"),
      sp(),
      p("I would replace the optional image URL field with a proper ", bold("direct file upload"), " feature using ", bold("Vercel Blob"), " or ", bold("Cloudinary"), ". Currently users must paste an external image URL, which is fragile — links can break or become unavailable. A drag-and-drop upload would store images reliably and give a much better user experience."),
      sp(),
      p("I would also add ", bold("SWR or React Query"), " for automatic background refresh of the notice list, so users see new notices without manually refreshing the page."),
      sp(),
      hr(),

      // ── AI USAGE ─────────────────────────────────────────────────────
      h1("AI Usage"),
      sp(),
      p("I used ", bold("Claude (Anthropic)"), " during this assignment in the following ways:"),
      sp(),
      bullet("Generated initial boilerplate for API route server-side validation logic"),
      bullet("Helped debug the Prisma orderBy configuration to correctly sort Urgent notices before Normal notices at the database level"),
      bullet("Suggested Tailwind CSS class combinations for the responsive card grid layout"),
      bullet("Helped draft sections of this README"),
      sp(),
      p(bold("What I did myself:"), " All architecture decisions, database schema design, component structure planning, and integration of all parts. I reviewed and understood every line of generated code before including it in the project. The final implementation, debugging of deployment issues, and testing were all done by me."),
      sp(),
      hr(),

      // ── DEPLOYMENT CHECKLIST ─────────────────────────────────────────
      h1("Deployment Checklist"),
      sp(),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Live Vercel URL opens without logging in", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  GitHub repository is public", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Multiple meaningful commits in history", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Create, Edit, Delete all work on production", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Urgent notices appear first with red badge", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Delete asks for confirmation first", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Pages Router used (pages/ not app/)", font: "Arial", size: 22 })] }),
      new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { before: 30, after: 30 },
        children: [new TextRun({ text: "  Free tier only — no credit card used", font: "Arial", size: 22 })] }),
      sp(),
      hr(),

      // ── FOOTER ──────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 0 },
        children: [new TextRun({ text: "Reno Platforms · Web Development Internship · vijay0778", size: 18, font: "Arial", color: "999999" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/README_GitHub.docx", buf);
  console.log("Done!");
});
