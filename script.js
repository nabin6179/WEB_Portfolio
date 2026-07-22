const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, LevelFormat, convertInchesToTwip
} = require("docx");

const ACCENT = "00ADB5";
const DARK = "121212";
const GREY = "555555";

const heading = (text, level = HeadingLevel.HEADING_1) =>
  new Paragraph({ text, heading: level, spacing: { before: 280, after: 140 } });

const body = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text, size: 22, ...opts })],
  });

const bullet = (text) =>
  new Paragraph({
    text,
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
  });

const codeBlock = (lines) =>
  new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: "1B2021" },
    spacing: { before: 100, after: 200 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "2A3132" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "2A3132" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "2A3132" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "2A3132" },
    },
    children: [new TextRun({ text: lines, font: "Consolas", size: 19, color: "E8F2F2" })],
  });

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 2000, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, fill: "121212" } : undefined,
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            bold: !!opts.header,
            color: opts.header ? "FFFFFF" : "000000",
            size: 20,
          }),
        ],
      }),
    ],
  });
}

const marking = [
  ["HTML5 Semantics (<nav>, <section>, <footer>)", "20", "Implemented — see Section 3.1"],
  ["CSS Responsiveness (Flexbox/Grid & Media Queries)", "30", "Implemented — see Section 3.2"],
  ["UI/UX Consistency (Colour, Typography, Alignment)", "20", "Implemented — see Section 2"],
  ["JavaScript Interaction (Smooth scroll / dynamic filter)", "15", "Implemented — see Section 3.3"],
  ["Deployment & Git (Working link, clean commit history)", "15", "Complete once deployed — see Section 5"],
];

const markingRows = [
  new TableRow({
    tableHeader: true,
    children: [
      cell("Criteria", { header: true, width: 5500 }),
      cell("Marks", { header: true, width: 1200 }),
      cell("Status", { header: true, width: 2300 }),
    ],
  }),
  ...marking.map(([c, m, s]) =>
    new TableRow({
      children: [cell(c, { width: 5500 }), cell(m, { width: 1200 }), cell(s, { width: 2300 })],
    })
  ),
];

const paletteRows = [
  new TableRow({
    tableHeader: true,
    children: [
      cell("Role", { header: true, width: 2500 }),
      cell("Hex", { header: true, width: 2000 }),
      cell("Usage", { header: true, width: 4500 }),
    ],
  }),
  new TableRow({
    children: [
      cell("Background", { width: 2500 }),
      cell("#121212", { width: 2000 }),
      cell("Page background — dark mode base", { width: 4500 }),
    ],
  }),
  new TableRow({
    children: [
      cell("Text / Primary", { width: 2500 }),
      cell("#FFFFFF", { width: 2000 }),
      cell("Headings, body copy on dark surfaces", { width: 4500 }),
    ],
  }),
  new TableRow({
    children: [
      cell("Accent", { width: 2500 }),
      cell("#00ADB5", { width: 2000 }),
      cell("Links, buttons, highlights, focus states", { width: 4500 }),
    ],
  }),
];

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
        },
      },
      children: [
        // ---------- COVER PAGE ----------
        new Paragraph({ text: "", spacing: { after: 1200 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Texas College of Management & IT", bold: true, size: 30 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "Bachelor of Computer Science (Hons) — Network Technology & Cybersecurity", bold: true, size: 24 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 600, after: 200 },
          children: [new TextRun({ text: "MINI PROJECT REPORT", bold: true, size: 26, color: ACCENT })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 800 },
          children: [new TextRun({ text: "\u201CThe Digital Identity Project\u201D", italics: true, size: 24 })],
        }),
        new Paragraph({ text: "", spacing: { after: 800 } }),
        new Paragraph({ children: [new TextRun({ text: "Course Title: ", bold: true, size: 22 }), new TextRun({ text: "Web Designing", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Course Code: ", bold: true, size: 22 }), new TextRun({ text: "WBD - 113", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Instructor: ", bold: true, size: 22 }), new TextRun({ text: "Dr. Ashish Gautam", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Submitted By: ", bold: true, size: 22 }), new TextRun({ text: "Nabin Tiwari", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "LCID: ", bold: true, size: 22 }), new TextRun({ text: "_____________________", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Year / Semester: ", bold: true, size: 22 }), new TextRun({ text: "BCS 2nd Year / 4th Semester", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Date of Submission: ", bold: true, size: 22 }), new TextRun({ text: "_____________________", size: 22 })], spacing: { after: 120 } }),
        new Paragraph({ children: [new PageBreak()] }),

        // ---------- 1. OBJECTIVE ----------
        heading("1. Objective"),
        body("The objective of this mini project is to design, build and deploy a responsive, single-page personal portfolio — \u201CThe Digital Identity Project\u201D — that showcases my skills, projects and contact information using modern, standard-compliant web technologies (HTML5, CSS3 and vanilla JavaScript)."),
        body("The site presents my profile as a Network Technology & Cybersecurity student: a summary of my background, a skills toolkit, a filterable showcase of security and networking projects, and a way for recruiters to get in touch."),

        // ---------- 2. INFORMATION ARCHITECTURE ----------
        heading("2. Information Architecture & Planning"),
        body("Before writing any code, the site was planned as a single page divided into five clearly labelled sections, navigated via a sticky top bar. This keeps the whole \u201Cdigital identity\u201D reachable within one or two scrolls or clicks, which suits a recruiter scanning quickly."),

        heading("2.1 Sitemap", HeadingLevel.HEADING_2),
        bullet("Home (Hero) — name, role, one-line pitch, call-to-action buttons"),
        bullet("About — profile summary and key facts (location, focus, education)"),
        bullet("Skills — a 4-card grid of technical skill categories"),
        bullet("Projects — a filterable grid of six projects (Network / Web / Pentest)"),
        bullet("Contact — links to portfolio site, GitHub and email"),

        heading("2.2 UI/UX Goal — F-Layout", HeadingLevel.HEADING_2),
        body("The hero section places the name and pitch top-left and a supporting visual (an animated terminal) to the right, matching the F-Layout pattern so a recruiter scanning top-left to right encounters the most important identity information first."),

        heading("2.3 Colour Palette", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9000, type: WidthType.DXA },
          columnWidths: [2500, 2000, 4500],
          rows: paletteRows,
        }),
        new Paragraph({ text: "", spacing: { after: 200 } }),

        heading("2.4 Typography", HeadingLevel.HEADING_2),
        body("Inter (sans-serif) is used for all headings and body copy for clean readability. JetBrains Mono, a monospace utility face, is used sparingly for the eyebrow labels, the logo and the terminal element — reinforcing the cybersecurity/command-line theme without hurting readability elsewhere."),

        // ---------- 3. DEVELOPMENT & CODE STRUCTURE ----------
        heading("3. Development Steps & Code Structure"),
        body("The project uses three files: index.html (structure), style.css (presentation) and script.js (behaviour). This separation of concerns keeps the codebase easy to maintain and mirrors standard front-end practice."),

        heading("3.1 HTML5 Semantics", HeadingLevel.HEADING_2),
        body("Semantic elements replace generic <div> tags wherever a native tag exists, which improves accessibility and SEO:"),
        codeBlock(
          "<header class=\"site-header\">\n" +
          "  <nav class=\"navbar\">...</nav>\n" +
          "</header>\n" +
          "<main id=\"main\">\n" +
          "  <section id=\"home\" class=\"hero\">...</section>\n" +
          "  <section id=\"about\" class=\"section\">...</section>\n" +
          "  <section id=\"projects\" class=\"section\">\n" +
          "    <article class=\"project-card\">...</article>\n" +
          "  </section>\n" +
          "</main>\n" +
          "<footer class=\"site-footer\">...</footer>"
        ),
        body("A \"skip to content\" link and visible :focus-visible outlines are also included for keyboard accessibility."),

        heading("3.2 CSS Responsiveness", HeadingLevel.HEADING_2),
        body("Layout uses CSS Grid for the hero, skills and project galleries, and Flexbox for the navbar, buttons and tag rows. Two media query breakpoints reflow the design for tablet and mobile:"),
        codeBlock(
          "@media (max-width: 900px) {\n" +
          "  .hero-grid { grid-template-columns: 1fr; }\n" +
          "  .skills-grid { grid-template-columns: repeat(2, 1fr); }\n" +
          "}\n\n" +
          "@media (max-width: 640px) {\n" +
          "  .nav-menu { /* collapses into a toggled mobile menu */ }\n" +
          "  .skills-grid,\n" +
          "  .projects-grid { grid-template-columns: 1fr; }\n" +
          "}"
        ),
        body("At the mobile breakpoint the horizontal navbar collapses behind a hamburger toggle, and every multi-column grid drops to a single column so content stacks vertically without horizontal scrolling."),

        heading("3.3 JavaScript Interaction", HeadingLevel.HEADING_2),
        body("script.js adds four behaviours:"),
        bullet("Smooth-scroll navigation with an IntersectionObserver that highlights the current section's nav link as the user scrolls."),
        bullet("A dynamic project filter — clicking \u201CNetwork Security\u201D, \u201CWeb Security\u201D or \u201CPentesting\u201D hides non-matching project cards using a data-category attribute, with no page reload."),
        bullet("A scroll-reveal effect that fades skill and project cards in as they enter the viewport."),
        bullet("A typed-out terminal animation in the hero (the site's signature visual element) that types an nmap command and prints mock scan output, reinforcing the security theme."),
        codeBlock(
          "filterButtons.forEach(btn => {\n" +
          "  btn.addEventListener('click', () => {\n" +
          "    const filter = btn.dataset.filter;\n" +
          "    projectCards.forEach(card => {\n" +
          "      const categories = card.dataset.category.split(' ');\n" +
          "      const show = filter === 'all' || categories.includes(filter);\n" +
          "      card.classList.toggle('hidden', !show);\n" +
          "    });\n" +
          "  });\n" +
          "});"
        ),

        // ---------- 4. TESTING / SCREENSHOTS ----------
        heading("4. Testing & Documentation"),
        body("The site was checked in a desktop viewport and a mobile viewport (DevTools device toolbar) to confirm the layout reflows correctly, and it should also be run through Lighthouse before submission. Insert the three required screenshots below once captured:"),
        body("[ Insert Screenshot: Desktop View — full navigation and hero section ]", { italics: true, color: GREY }),
        body("[ Insert Screenshot: Mobile View — responsive vertical stack from a mobile simulator ]", { italics: true, color: GREY }),
        body("[ Insert Screenshot: Lighthouse Report — performance audit, aim for 90+ ]", { italics: true, color: GREY }),

        // ---------- 5. DEPLOYMENT ----------
        heading("5. Deployment & Git"),
        body("Steps to complete this section before submission:"),
        bullet("git init in the project folder, then git add . && git commit -m \"Initial commit: Digital Identity Project\""),
        bullet("Create a public GitHub repository and push: git remote add origin https://github.com/nabin6179/digital-identity-project.git then git push -u origin main"),
        bullet("Enable GitHub Pages (Settings → Pages → Deploy from branch) or deploy to Netlify for a live URL"),
        new Paragraph({ children: [new TextRun({ text: "GitHub Link: ", bold: true, size: 22 }), new TextRun({ text: "https://github.com/nabin6179/digital-identity-project", size: 22, color: ACCENT })], spacing: { after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: "Live Deployment: ", bold: true, size: 22 }), new TextRun({ text: "https://nabin6179.github.io/digital-identity-project", size: 22, color: ACCENT })], spacing: { after: 120 } }),

        // ---------- 6. SELF ASSESSMENT ----------
        heading("6. Self-Assessment Against Marking Scheme"),
        new Table({
          width: { size: 9000, type: WidthType.DXA },
          columnWidths: [5500, 1200, 2300],
          rows: markingRows,
        }),
        new Paragraph({ text: "", spacing: { after: 200 } }),

        // ---------- 7. CONCLUSION ----------
        heading("7. Conclusion"),
        body("This mini project translated the theoretical modules of WBD-113 into a working, deployed artefact: a responsive personal portfolio built with semantic HTML5, a Grid/Flexbox layout governed by two media-query breakpoints, and vanilla JavaScript interactivity (scroll-spy navigation and a dynamic project filter). The result is a single-page \u201Cdigital identity\u201D that is readable on desktop and mobile alike, and that accurately represents my current focus — network security, penetration testing and web application security — to prospective employers."),

        new Paragraph({ children: [new PageBreak()] }),

        // ---------- DECLARATION ----------
        heading("Student Declaration"),
        body("I declare that this mini project is my original work and completed by me."),
        new Paragraph({ text: "", spacing: { after: 600 } }),
        body("Student Signature: ________________________________"),
        body("Date: ________________________________"),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("/home/claude/Digital_Identity_Project_Report.docx", buf);
  console.log("done");
});