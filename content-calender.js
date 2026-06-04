const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition,
  ExternalHyperlink, Header, Footer, UnderlineType
} = require('docx');
const fs = require('fs');

// ─── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  navy:    "1B3A6B",
  blue:    "2563EB",
  lightBlue: "DBEAFE",
  teal:    "0D9488",
  lightTeal: "CCFBF1",
  gold:    "D97706",
  lightGold: "FEF3C7",
  green:   "16A34A",
  lightGreen: "DCFCE7",
  red:     "DC2626",
  lightRed: "FEE2E2",
  purple:  "7C3AED",
  lightPurple: "EDE9FE",
  gray:    "6B7280",
  lightGray: "F3F4F6",
  darkGray: "374151",
  white:   "FFFFFF",
  black:   "111827",
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const br = () => new Paragraph({ children: [new TextRun("")], spacing: { after: 80 } });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const h1 = (text, color = C.navy) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, color, bold: true, size: 36, font: "Arial" })],
  spacing: { before: 360, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.blue, space: 4 } }
});

const h2 = (text, color = C.navy) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, color, bold: true, size: 28, font: "Arial" })],
  spacing: { before: 280, after: 160 }
});

const h3 = (text, color = C.darkGray) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, color, bold: true, size: 24, font: "Arial" })],
  spacing: { before: 200, after: 120 }
});

const p = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, font: "Arial", size: 22, color: C.darkGray, ...opts })],
  spacing: { after: 120, line: 276 }
});

const bold = (text, color = C.black) => new TextRun({ text, bold: true, font: "Arial", size: 22, color });

const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullets", level },
  children: [new TextRun({ text, font: "Arial", size: 22, color: C.darkGray })],
  spacing: { after: 80 }
});

const numbered = (text, level = 0) => new Paragraph({
  numbering: { reference: "numbers", level },
  children: [new TextRun({ text, font: "Arial", size: 22, color: C.darkGray })],
  spacing: { after: 80 }
});

const colorBullet = (text, color) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text, font: "Arial", size: 22, color })],
  spacing: { after: 80 }
});

// Colored callout box (single-row, single-cell table)
const callout = (label, text, bgColor, labelColor = C.white) => {
  const border = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      width: { size: 9360, type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: label + " ", bold: true, font: "Arial", size: 22, color: labelColor }),
        new TextRun({ text, font: "Arial", size: 22, color: labelColor })
      ], spacing: { after: 0 } })]
    })]})],
  });
};

// Two-column table
const twoCol = (rows, hdr1 = "", hdr2 = "", col1W = 3000, col2W = 6360) => {
  const bdr = { style: BorderStyle.SINGLE, size: 1, color: "D1D5DB" };
  const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
  const hdrBdr = { style: BorderStyle.SINGLE, size: 2, color: C.blue };
  const hdrBorders = { top: hdrBdr, bottom: hdrBdr, left: hdrBdr, right: hdrBdr };

  const tableRows = [];
  if (hdr1) {
    tableRows.push(new TableRow({ children: [
      new TableCell({ borders: hdrBorders, shading: { fill: C.navy, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, width: { size: col1W, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: hdr1, bold: true, font: "Arial", size: 20, color: C.white })] })] }),
      new TableCell({ borders: hdrBorders, shading: { fill: C.navy, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, width: { size: col2W, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: hdr2, bold: true, font: "Arial", size: 20, color: C.white })] })] }),
    ]}));
  }

  rows.forEach(([a, b], i) => {
    const fill = i % 2 === 0 ? C.white : "F9FAFB";
    tableRows.push(new TableRow({ children: [
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, width: { size: col1W, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: a, bold: true, font: "Arial", size: 20, color: C.navy })] })] }),
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, width: { size: col2W, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: b, font: "Arial", size: 20, color: C.darkGray })] })] }),
    ]}));
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [col1W, col2W], rows: tableRows });
};

// ─── COVER PAGE ──────────────────────────────────────────────────────────────
const coverPage = () => [
  br(), br(), br(),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "🏗", size: 80 })], spacing: { after: 200 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "100-DAY FULL-STACK DEVELOPER", bold: true, font: "Arial", size: 52, color: C.navy })], spacing: { after: 60 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ROADMAP", bold: true, font: "Arial", size: 52, color: C.blue })], spacing: { after: 160 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Real Estate Tech Specialization", font: "Arial", size: 30, color: C.teal, bold: true })], spacing: { after: 80 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "From Zero to First Paying Freelance Client", font: "Arial", size: 26, color: C.gray })], spacing: { after: 400 } }),
  br(),
  new Table({
    width: { size: 7000, type: WidthType.DXA },
    columnWidths: [1750, 1750, 1750, 1750],
    rows: [new TableRow({ children: [
      ...["9 Parts","100 Days","6 Projects","$100k Path"].map(t =>
        new TableCell({ borders: { top: { style: BorderStyle.NONE, size: 0, color: "fff" }, bottom: { style: BorderStyle.NONE, size: 0, color: "fff" }, left: { style: BorderStyle.NONE, size: 0, color: "fff" }, right: { style: BorderStyle.NONE, size: 0, color: "fff" } },
          shading: { fill: C.navy, type: ShadingType.CLEAR }, margins: { top: 160, bottom: 160, left: 80, right: 80 }, width: { size: 1750, type: WidthType.DXA },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, bold: true, font: "Arial", size: 22, color: C.white })] })] })
      )
    ]})],
  }),
  br(), br(), br(), br(),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Prepared for: Full-Stack Freelance Developer", font: "Arial", size: 22, color: C.gray })], spacing: { after: 60 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Specialization: Real Estate Technology", font: "Arial", size: 22, color: C.gray })], spacing: { after: 60 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Goal: First Paying Client in 100 Days", font: "Arial", size: 22, color: C.teal, bold: true })], spacing: { after: 60 } }),
  pageBreak()
];

// ─── PART 1: TECHNICAL ROADMAP ────────────────────────────────────────────────
const part1 = () => {
  const days = [
    // Phase 1: Web Foundations (Days 1-14)
    { range:"Days 1–2", topic:"HTML5 Fundamentals", time:"3 hrs/day", topics:"Semantic HTML, structure, forms, accessibility, meta tags, SEO basics", resources:"freeCodeCamp HTML (free), MDN Web Docs (free)", exercise:"Build a real estate agent bio page with semantic HTML", deliverable:"Deployed HTML page on GitHub Pages" },
    { range:"Days 3–5", topic:"CSS3 Mastery", time:"3 hrs/day", topics:"Box model, flexbox, grid, animations, responsive design, CSS variables, media queries", resources:"Kevin Powell YouTube (free), CSS Tricks (free), Scrimba CSS", exercise:"Style the agent bio page — mobile-first, card-based layout", deliverable:"Fully responsive styled agent page" },
    { range:"Days 6–8", topic:"Tailwind CSS", time:"3 hrs/day", topics:"Utility-first workflow, config, responsive prefixes, dark mode, component patterns", resources:"Tailwind docs (free), Tailwind UI (paid), Scrimba Tailwind", exercise:"Rebuild the agent bio page using Tailwind only", deliverable:"Same page, zero custom CSS, deployed" },
    { range:"Days 9–14", topic:"JavaScript ES6+", time:"4 hrs/day", topics:"Variables, functions, arrays, objects, DOM, fetch, async/await, modules, destructuring, spread, promises", resources:"javascript.info (free), The Odin Project JS (free), Eloquent JS (free)", exercise:"Build a property search filter (JS arrays + DOM)", deliverable:"Working JS property search widget" },
    // Phase 2: React (Days 15-30)
    { range:"Days 15–17", topic:"TypeScript Basics", time:"3 hrs/day", topics:"Types, interfaces, generics, enums, type inference, strict mode", resources:"TypeScript Handbook (free), Matt Pocock YouTube (free)", exercise:"Convert JS property filter to TypeScript", deliverable:"TS version of property filter" },
    { range:"Days 18–24", topic:"React.js Core", time:"4 hrs/day", topics:"Components, JSX, props, state, hooks (useState, useEffect, useContext, useRef), event handling, forms, lists, lifting state", resources:"Official React docs (free), Scrimba React, Jack Herrington YouTube (free)", exercise:"Build a real estate property listing app with filters", deliverable:"Deployed React property listing app" },
    { range:"Days 25–30", topic:"Next.js Foundations", time:"4 hrs/day", topics:"App router, file-based routing, SSR, SSG, ISR, API routes, metadata API, image optimization, layouts", resources:"Next.js official tutorial (free), Lee Robinson YouTube (free)", exercise:"Convert React property app to Next.js with SSG", deliverable:"Next.js real estate listing site on Vercel" },
    // Phase 3: Backend (Days 31-55)
    { range:"Days 31–36", topic:"Node.js & Express", time:"4 hrs/day", topics:"Node runtime, npm, modules, Express routing, middleware, REST principles, request/response cycle, error handling", resources:"The Odin Project backend (free), Traversy Media Node crash course (free)", exercise:"Build a properties CRUD REST API", deliverable:"Live REST API on Railway/Render" },
    { range:"Days 37–42", topic:"PostgreSQL & SQL", time:"4 hrs/day", topics:"Relational databases, tables, CRUD, JOINs, indexes, constraints, normalization, Prisma ORM", resources:"SQLZoo (free), Supabase docs (free), Prisma docs (free)", exercise:"Design and build a real estate database schema", deliverable:"Working DB with properties, agents, users tables" },
    { range:"Days 43–47", topic:"MongoDB & Mongoose", time:"3 hrs/day", topics:"Document model, collections, CRUD, aggregation, indexing, schema design, population", resources:"MongoDB University (free), Mongoose docs (free)", exercise:"Build a property leads/CRM document store", deliverable:"MongoDB-backed leads API" },
    { range:"Days 48–53", topic:"Authentication & Authorization", time:"4 hrs/day", topics:"JWT tokens, bcrypt, sessions, NextAuth.js, role-based access (admin, agent, client), OAuth (Google)", resources:"NextAuth docs (free), JWT.io (free), Traversy Media auth videos (free)", exercise:"Add auth to the real estate API — agent login, admin panel", deliverable:"Full auth flow with protected routes" },
    { range:"Days 54–55", topic:"REST API Design & Best Practices", time:"3 hrs/day", topics:"API versioning, pagination, filtering, sorting, rate limiting, error codes, documentation (Swagger)", resources:"REST API Design Guidelines (free), Swagger docs (free)", exercise:"Document your properties API with Swagger", deliverable:"Swagger API documentation page" },
    // Phase 4: Full Stack Integration (Days 56-72)
    { range:"Days 56–60", topic:"Git, GitHub & Workflow", time:"3 hrs/day", topics:"Git flow, branching, PRs, .gitignore, README writing, GitHub Actions (basic CI), conventional commits", resources:"GitHub Learning Lab (free), Atlassian Git tutorial (free)", exercise:"Set up a professional GitHub profile with pinned repos and READMEs", deliverable:"GitHub profile ready to show clients" },
    { range:"Days 61–66", topic:"Full-Stack Project Integration", time:"5 hrs/day", topics:"Connect Next.js frontend to Express/Node backend, environment variables, CORS, API client (Axios/fetch), state management (Zustand)", resources:"Traversy Media full-stack tutorials (free), Next.js docs", exercise:"Build a full-stack real estate listing site (frontend + backend + DB)", deliverable:"Deployed full-stack real estate app — Project #1" },
    { range:"Days 67–70", topic:"Deployment & DevOps Basics", time:"3 hrs/day", topics:"Vercel (Next.js), Railway/Render (Node.js), environment management, domain setup, HTTPS, CI/CD basics", resources:"Vercel docs (free), Railway docs (free), Render docs (free)", exercise:"Deploy full-stack app with custom domain", deliverable:"Live app at custom domain" },
    { range:"Days 71–72", topic:"Performance Optimization", time:"3 hrs/day", topics:"Lighthouse audits, Core Web Vitals, lazy loading, image optimization, code splitting, caching headers, CDN basics", resources:"web.dev (free), Next.js performance docs (free)", exercise:"Audit and optimize your real estate app to 90+ Lighthouse score", deliverable:"Performance-optimized deployed app" },
    // Phase 5: Advanced (Days 73-88)
    { range:"Days 73–76", topic:"SEO for Developers", time:"3 hrs/day", topics:"Technical SEO, meta tags, Open Graph, structured data (JSON-LD for real estate), sitemaps, robots.txt, canonical URLs", resources:"Ahrefs SEO basics (free), Google Search Console (free)", exercise:"Add full SEO to real estate app including JSON-LD property schema", deliverable:"SEO-optimized app with structured data" },
    { range:"Days 77–79", topic:"Testing Fundamentals", time:"3 hrs/day", topics:"Unit testing (Vitest/Jest), React Testing Library, integration tests, API testing (Supertest), test-driven basics", resources:"Vitest docs (free), Testing Library docs (free)", exercise:"Write tests for your properties API and React components", deliverable:"Test suite with 80%+ coverage" },
    { range:"Days 80–84", topic:"AI Integration", time:"4 hrs/day", topics:"OpenAI/Anthropic API, prompt engineering, streaming responses, AI-powered features (property descriptions, chatbot, valuation assistant)", resources:"OpenAI docs (free tier), Anthropic docs (free tier), Vercel AI SDK (free)", exercise:"Build an AI property description generator for real estate listings", deliverable:"AI feature integrated into your real estate app" },
    { range:"Days 85–88", topic:"Full-Stack Architecture & Patterns", time:"4 hrs/day", topics:"MVC pattern, repository pattern, service layer, clean architecture, microservices intro, system design basics", resources:"Architecture patterns (free articles), ByteByteGo YouTube (free)", exercise:"Refactor your codebase to use clean service/repository pattern", deliverable:"Refactored codebase with clear separation of concerns" },
    // Phase 6: Freelance Ready (Days 89-100)
    { range:"Days 89–92", topic:"Real Estate Project #2 Build", time:"6 hrs/day", topics:"Build a property management dashboard with auth, listings, inquiries, and analytics — your flagship portfolio piece", resources:"Your accumulated knowledge + Figma for UI reference", exercise:"Full build sprint", deliverable:"Deployed property management dashboard" },
    { range:"Days 93–96", topic:"Portfolio Website Build", time:"5 hrs/day", topics:"Personal brand site, project showcases, case studies, contact form, testimonials section, SEO", resources:"Awwwards for inspiration, Tailwind CSS + Next.js", exercise:"Build and deploy yourname.dev", deliverable:"Live personal portfolio website" },
    { range:"Days 97–100", topic:"Client Launch Sprint", time:"5 hrs/day", topics:"Upwork profile, LinkedIn optimization, cold outreach, proposal templates, first client pitch", resources:"Upwork Academy (free), LinkedIn Learning", exercise:"Send 10 proposals and 20 outreach messages", deliverable:"First client conversation booked" },
  ];

  const dayRows = days.map(d => [d.range, `${d.topic} (${d.time})\nTopics: ${d.topics}\nResources: ${d.resources}\nExercise: ${d.exercise}\nDeliverable: ${d.deliverable}`]);

  return [
    h1("PART 1: 100-Day Technical Learning Roadmap"),
    p("This day-by-day roadmap takes you from complete beginner to a job-ready full-stack developer specializing in real estate applications. Designed for 3–5 hours of focused daily study."),
    br(),
    callout("🎯 GOAL:", "By Day 100 you will have built 2+ full-stack apps, a live portfolio, and sent your first client proposals.", C.navy),
    br(),
    h2("The 6 Learning Phases"),
    twoCol([
      ["Phase 1 | Days 1–14", "Web Foundations — HTML5, CSS3, Tailwind CSS, JavaScript ES6+"],
      ["Phase 2 | Days 15–30", "Frontend Framework — TypeScript, React.js, Next.js"],
      ["Phase 3 | Days 31–55", "Backend & Databases — Node.js, Express, PostgreSQL, MongoDB, Auth"],
      ["Phase 4 | Days 56–72", "Full-Stack Integration — Git, Deployment, Performance"],
      ["Phase 5 | Days 73–88", "Advanced Skills — SEO, Testing, AI Integration, Architecture"],
      ["Phase 6 | Days 89–100", "Freelance Launch — Portfolio, Projects, Client Acquisition"],
    ], "Phase", "What You Learn"),
    br(),
    h2("Day-by-Day Learning Plan"),
    ...days.flatMap(d => [
      new Paragraph({ children: [
        new TextRun({ text: `${d.range}: `, bold: true, font: "Arial", size: 22, color: C.blue }),
        new TextRun({ text: d.topic, bold: true, font: "Arial", size: 22, color: C.navy }),
        new TextRun({ text: ` (${d.time})`, font: "Arial", size: 20, color: C.gray }),
      ], spacing: { before: 160, after: 60 } }),
      bullet(`Topics: ${d.topics}`),
      bullet(`Resources: ${d.resources}`),
      bullet(`Exercise: ${d.exercise}`),
      colorBullet(`✅ Deliverable: ${d.deliverable}`, C.green),
      br(),
    ]),
    callout("💡 DAILY ROUTINE:", "1hr theory (video/docs) → 30min hands-on exercise → 1.5hrs build project → 30min review & commit to GitHub. Consistency beats intensity.", C.lightBlue, C.navy),
  ];
};

// ─── PART 2: REAL ESTATE NICHE ───────────────────────────────────────────────
const part2 = () => [
  pageBreak(),
  h1("PART 2: Real Estate Niche Specialization"),
  p("Real estate is one of the most profitable niches for freelance developers. Here's everything you need to dominate it."),
  br(),
  h2("Why Real Estate is a Gold Mine for Developers"),
  twoCol([
    ["Market Size", "$4.3 trillion industry in the US alone — even 0.001% is life-changing revenue"],
    ["Tech Gap", "Most real estate businesses use outdated 2010-era websites or none at all"],
    ["High Budgets", "Real estate agents earn $80k–$300k+ and invest heavily in marketing tools"],
    ["Clear ROI", "One extra property sale = $5,000–$30,000 commission. Your $3,000 site pays for itself immediately"],
    ["Referral Culture", "Real estate runs on referrals — one happy client sends 3–5 more"],
    ["Recurring Revenue", "MLS integrations, lead management, CRM tools = monthly retainer opportunities"],
  ], "Factor", "Why It Matters"),
  br(),
  h2("Real Estate Business Types You Can Serve"),
  ...["Individual Real Estate Agents — personal brand sites, listing showcases, lead capture forms",
     "Real Estate Teams & Brokerages — team sites, agent directories, performance dashboards",
     "Property Management Companies — tenant portals, maintenance request systems, rent collection",
     "Real Estate Investors — portfolio sites, deal calculators, market analysis tools",
     "Real Estate Developers — project showcase sites, pre-sale landing pages, virtual tour integration",
     "Commercial Real Estate Firms — property search portals, lease calculators, inquiry systems",
     "Real Estate Marketing Agencies — campaign landing pages, A/B testing tools, lead funnels",
     "PropTech Startups — MVP development, investor pitch apps, market data dashboards"
  ].map(b => bullet(b)),
  br(),
  h2("Problems Real Estate Businesses Face (That You Solve)"),
  twoCol([
    ["Slow/outdated website", "Build fast, modern Next.js sites with 95+ Lighthouse scores"],
    ["No lead capture", "Build contact forms, chatbots, property inquiry forms with email automation"],
    ["Manual listing updates", "Build CMS-powered listing pages agents update themselves"],
    ["No mobile experience", "Build mobile-first responsive designs — 60% of searches are mobile"],
    ["No analytics", "Add Google Analytics 4, heatmaps, conversion tracking"],
    ["Poor SEO visibility", "Add local SEO, structured data, Google Business optimization"],
    ["Manual follow-up", "Build automated email sequences and lead nurturing flows"],
    ["No virtual tours", "Integrate Matterport, iGUIDE, or YouTube virtual tour embeds"],
  ], "Their Problem", "Your Solution"),
  br(),
  h2("Software Solutions & Apps You Can Build"),
  ...["Property listing websites with advanced search and filtering (price, beds, baths, location)",
     "Agent CRM systems for managing leads, follow-ups, and deal pipelines",
     "Booking/showing request systems with calendar integration (Calendly/Cal.com)",
     "Mortgage and ROI calculator tools embedded in property pages",
     "Market statistics dashboards pulling live data for agents and investors",
     "Tenant management portals with maintenance requests and rent tracking",
     "Virtual tour integration platforms connecting listings to 3D tour services",
     "Email marketing landing pages with lead magnet delivery automation",
     "Property comparison tools (side-by-side comparison of 2–4 properties)",
     "Neighborhood insights pages with walkability scores, schools, demographics",
  ].map(b => bullet(b)),
  br(),
  h2("Positioning: Become a Real Estate Tech Specialist"),
  callout("🏆 YOUR POSITIONING STATEMENT:", '"I build high-performance websites and digital tools exclusively for real estate professionals — helping agents generate more leads, close more deals, and grow their brand online."', C.navy),
  br(),
  h2("Pricing Strategy for Real Estate Clients"),
  twoCol([
    ["Agent Personal Site", "$800 – $2,500 (starter package)"],
    ["Brokerage/Team Site", "$3,000 – $8,000"],
    ["Property Listing Platform", "$5,000 – $15,000"],
    ["Custom CRM Tool", "$4,000 – $12,000"],
    ["Monthly Maintenance Retainer", "$300 – $800/month"],
    ["SEO Package Add-on", "$500 – $1,500/month"],
    ["Lead Funnel + Landing Page", "$1,200 – $3,500"],
    ["AI Integration Feature", "$2,000 – $6,000 add-on"],
  ], "Service", "Price Range"),
];

// ─── PART 3: PORTFOLIO PROJECTS ───────────────────────────────────────────────
const part3 = () => {
  const projects = [
    {
      num: "01", name: "Elite Real Estate Agent Website",
      overview: "A high-converting personal brand website for a real estate agent with property listings, lead capture, and Google Maps integration.",
      features: ["Hero with CTA and lead capture form","Dynamic listings page (filter by price, beds, location)","Individual property detail pages","Agent bio + testimonials","Blog/market updates section","Google Maps property pins","Contact form with email notification (Resend)","Mobile-first responsive design"],
      stack: "Next.js, Tailwind CSS, Sanity CMS, PostgreSQL, Resend, Vercel",
      inspiration: "Luxury Presence, AgentFire, Placester — clean white with gold accents",
      dbStructure: "Tables: properties (id, title, price, beds, baths, sqft, location, images, status), agents (id, name, bio, photo, contact), testimonials",
      apis: "Google Maps API, Resend (email), Sanity (CMS), Cloudinary (image hosting)",
      advanced: "IDX/MLS feed integration (Spark API), mortgage calculator widget, virtual tour embed, live chat (Crisp)",
      deployment: "Vercel for frontend + CMS, Cloudinary for images, custom domain (agentname.com)",
      caseStudy: "Problem: Agent had outdated HTML site losing leads. Solution: Modern Next.js site with lead capture. Result: Showcase 40% increase in inquiry form submissions."
    },
    {
      num: "02", name: "Property Listing & Search Platform",
      overview: "A full-stack property search platform similar to Zillow — users search, filter, and save properties. Agents manage their own listings via a dashboard.",
      features: ["Advanced search with 10+ filters (price, location, property type, features)","Interactive map view using Mapbox or Google Maps","User registration and saved properties","Agent dashboard — add/edit/delete listings","Property detail page with image gallery","Mortgage calculator on each listing page","Inquiry form connecting buyer to agent","Admin panel for platform management"],
      stack: "Next.js, TypeScript, PostgreSQL, Prisma, NextAuth, Mapbox GL, Tailwind CSS, Railway",
      inspiration: "Zillow, Realtor.com — clean search UX with map/list split view",
      dbStructure: "Tables: users, agents, properties, property_images, saved_properties, inquiries, agent_profiles",
      apis: "Mapbox or Google Maps API, Cloudinary, Resend, Walk Score API (neighborhood data)",
      advanced: "AI-powered property recommendation engine, price history chart (Recharts), school district overlay, commute time calculator",
      deployment: "Next.js on Vercel, PostgreSQL on Railway, images on Cloudinary",
      caseStudy: "Problem: Small brokerage had no digital listing presence. Solution: Custom search platform with agent dashboards. Result: Agents now manage listings independently — saved 10hrs/week admin time."
    },
    {
      num: "03", name: "Property Management Tenant Portal",
      overview: "A full-stack multi-tenant portal where landlords manage properties and tenants submit maintenance requests, pay rent, and communicate with property managers.",
      features: ["Landlord dashboard — manage properties, units, tenants","Tenant portal — view lease, submit maintenance requests","Maintenance request workflow (submit → assigned → resolved)","Rent payment tracking and reminders","Document storage (lease agreements, inspection reports)","Notification system (email + in-app)","Role-based access: admin, landlord, tenant, maintenance","Financial summary — rent collected vs. outstanding"],
      stack: "Next.js, TypeScript, PostgreSQL, Prisma, NextAuth, Stripe (payments), Tailwind CSS",
      inspiration: "Buildium, AppFolio — clean dashboard UI with sidebar navigation",
      dbStructure: "Tables: users (role), properties, units, tenancies, maintenance_requests, payments, documents, notifications",
      apis: "Stripe (rent payments), Resend (notifications), Cloudinary (documents), Google Calendar API",
      advanced: "Stripe recurring rent payments, AI maintenance triage (priority classification), occupancy analytics dashboard",
      deployment: "Vercel + Railway (DB) + Stripe webhooks",
      caseStudy: "Problem: Landlord managing 20 units using WhatsApp and spreadsheets. Solution: Custom portal reducing admin time by 60%. Result: Expanded to 40 units without hiring staff."
    },
    {
      num: "04", name: "Real Estate Investor Portfolio Dashboard",
      overview: "A private dashboard for real estate investors to track their property portfolio performance — ROI, cash flow, appreciation, equity, and deal pipeline.",
      features: ["Portfolio overview — total properties, equity, monthly cash flow","Per-property analytics — purchase price, current value, ROI, cap rate","Deal pipeline (prospecting → under contract → closed)","Expense tracking per property (repairs, taxes, insurance)","Mortgage tracker with amortization schedule","Market comparison — property value vs. local comps","Document vault per property (deeds, insurance, tax records)","Export reports to PDF"],
      stack: "Next.js, TypeScript, PostgreSQL, Prisma, NextAuth, Recharts, Tailwind CSS, Vercel",
      inspiration: "Stessa, Baselane — dark-mode-friendly dashboard with data visualization",
      dbStructure: "Tables: users, properties, transactions, expenses, mortgages, documents, deals, market_comps",
      apis: "Attom Data API (property valuations), Plaid (bank account sync — optional), Recharts (charts)",
      advanced: "AI deal analyzer (paste address → get investment analysis), automated comp pulling, tax report generation",
      deployment: "Vercel + Supabase (PostgreSQL) + PDF generation with Puppeteer",
      caseStudy: "Problem: Investor tracking 12 properties in Excel. Solution: Custom dashboard with automated ROI calculations. Result: Identified 2 underperforming properties and increased overall portfolio ROI by 18%."
    },
    {
      num: "05", name: "Real Estate Agent CRM",
      overview: "A lightweight CRM built specifically for real estate agents to manage leads, track deal stages, schedule follow-ups, and monitor pipeline performance.",
      features: ["Lead capture with source tracking (website, referral, Zillow, social)","Deal pipeline Kanban board (New → Nurturing → Active → Closed)","Lead detail page with notes, activity timeline, documents","Follow-up scheduler with email/SMS reminders","Email integration — log sent emails, template library","Performance dashboard — leads by source, conversion rate, revenue","Contact import/export (CSV)","Mobile-responsive for on-the-go agents"],
      stack: "Next.js, TypeScript, PostgreSQL, Prisma, NextAuth, React DnD (kanban), Tailwind CSS, Resend",
      inspiration: "HubSpot, Pipedrive — clean kanban with sidebar stats panel",
      dbStructure: "Tables: users, leads, deals, activities, notes, email_templates, follow_ups, pipeline_stages",
      apis: "Resend (email), Twilio (SMS reminders), Google Calendar API, CSV parsing",
      advanced: "AI lead scoring model, email sequence automation, predictive close date, WhatsApp integration",
      deployment: "Vercel + Railway + Twilio webhooks",
      caseStudy: "Problem: Agent losing track of leads across sticky notes and spreadsheets. Solution: Custom CRM with automated follow-ups. Result: Lead response time dropped from 48hrs to 2hrs — closed 3 additional deals in first quarter."
    },
    {
      num: "06", name: "AI-Powered Property Description Generator",
      overview: "A SaaS tool that uses AI to generate professional property listing descriptions, social media captions, and marketing copy — saving agents hours per listing.",
      features: ["Input form: property type, beds, baths, features, neighborhood, price","AI generates 3 listing description variations (standard, luxury, brief)","Social media caption generator (Instagram, Facebook, LinkedIn)","Email newsletter copy for the listing","One-click copy to clipboard","Saved generations history","Credit-based system (10 free, then subscription)","Agent branding — add agent name/brokerage to output"],
      stack: "Next.js, TypeScript, PostgreSQL, Prisma, NextAuth, Stripe, OpenAI/Anthropic API, Tailwind CSS",
      inspiration: "Copy.ai, Jasper — clean two-panel UI (input left, output right)",
      dbStructure: "Tables: users, generations, credits, subscriptions, templates, saved_outputs",
      apis: "OpenAI GPT-4o or Anthropic Claude API, Stripe (subscriptions), Resend",
      advanced: "Image-based generation (upload listing photos, AI describes them), bulk generation for entire listing portfolios, custom tone/style training per agent",
      deployment: "Vercel + Supabase + Stripe webhooks + OpenAI API",
      caseStudy: "Problem: Agents spend 45min writing each listing description. Solution: AI tool reduces this to 3 minutes. Result: 30+ agents paying $29/month = $870 MRR from a single project."
    },
  ];

  return [
    pageBreak(),
    h1("PART 3: Portfolio Projects for Real Estate Niche"),
    p("These 6 projects are specifically designed to impress real estate clients and demonstrate every skill needed for full-stack freelance work. Build them in order — each one builds on the last."),
    br(),
    callout("💼 PORTFOLIO STRATEGY:", "Build Projects 1 & 2 first (client-facing sites). Then Project 6 (AI tool) — it's the highest-value differentiator. Use each as a case study with before/after framing.", C.navy),
    br(),
    ...projects.flatMap(proj => [
      h2(`Project ${proj.num}: ${proj.name}`),
      p(proj.overview),
      h3("Key Features"),
      ...proj.features.map(f => bullet(f)),
      br(),
      twoCol([
        ["Tech Stack", proj.stack],
        ["UI Inspiration", proj.inspiration],
        ["Database Design", proj.dbStructure],
        ["APIs to Integrate", proj.apis],
        ["Advanced Features", proj.advanced],
        ["Deployment Plan", proj.deployment],
      ], "Area", "Details"),
      br(),
      callout("📋 CASE STUDY ANGLE:", proj.caseStudy, C.lightGreen, C.darkGray),
      br(),
    ]),
  ];
};

// ─── PART 4: CLIENT ACQUISITION ────────────────────────────────────────────────
const part4 = () => [
  pageBreak(),
  h1("PART 4: Freelance Client Acquisition Plan"),
  p("A step-by-step system for landing your first real estate tech client within 100 days."),
  br(),
  h2("Platform Setup Checklist"),
  h3("LinkedIn Profile Optimization"),
  ...["Headline: 'Full-Stack Developer for Real Estate Businesses | Next.js · React · Node.js'",
     "Banner: Screenshot of your best real estate project (1584x396px)",
     "About: Start with who you help, what you build, and one result/story. End with CTA.",
     "Featured: Pin your 3 best portfolio projects with live links",
     "Skills: Next.js, React, Node.js, PostgreSQL, TypeScript, REST APIs, Tailwind CSS",
     "Experience: List each portfolio project as 'Freelance Full-Stack Developer' with results",
     "Recommendations: Ask 2–3 people (bootcamp instructors, peers) to endorse your skills",
  ].map(b => bullet(b)),
  br(),
  h3("GitHub Profile Optimization"),
  ...["Profile README: Add a banner, bio, skills icons, and links to live projects",
     "Pin 3–4 best repos — real estate projects with live demo links in description",
     "Every repo needs: professional README, live demo link, clear description, good commit history",
     "Consistent daily commits show activity — even documentation updates count",
     "Add topics/tags to each repo (real-estate, nextjs, typescript, fullstack)",
  ].map(b => bullet(b)),
  br(),
  h3("Portfolio Website (yourname.dev)"),
  ...["Hero: Name, title ('Full-Stack Developer for Real Estate'), and 2 CTAs (View Work / Contact)",
     "Projects section: 3 best projects with live link, GitHub link, and 2-line description",
     "About: Short story — who you are, why real estate, what makes you different",
     "Testimonials: Even spec/free work testimonials count at the start",
     "Contact: Simple form + direct email + LinkedIn link",
     "Tech: Build it with Next.js + Tailwind — it IS a portfolio project in itself",
  ].map(b => bullet(b)),
  br(),
  h2("Outreach Strategy"),
  h3("Cold Email Template — Real Estate Agents"),
  callout("EMAIL TEMPLATE:", `Subject: Quick idea for [Agent Name]'s website

Hi [Name],

I came across your listings on [Platform] and noticed your website [specific observation — slow load / no mobile / outdated design].

I build websites specifically for real estate agents — focused on generating leads and showcasing listings professionally. I recently built [Project Name] which [specific result].

Would you be open to a 15-minute call this week? I have a few ideas that could help [specific goal — more leads / faster listing updates / better mobile experience].

[Your Name]
[Portfolio link]`, C.lightGray, C.darkGray),
  br(),
  h3("Upwork Strategy"),
  ...["Complete profile to 100% before applying to any jobs",
     "Headline: 'Real Estate Website & App Developer | Next.js · React · Full-Stack'",
     "Hourly rate: Start at $35–45/hr to build reviews — raise 30% after first 3 five-star reviews",
     "Apply to 5–8 jobs daily. Jobs posted in last 12 hours get highest response rates",
     "Proposal formula: First line references their specific project → 2–3 sentences on your approach → 1 question to invite a reply. Keep under 150 words.",
     "Target: real estate website jobs, property management app, MLS integration, agent CRM",
     "After 5 reviews, switch to fixed-price proposals for higher value projects",
  ].map(b => bullet(b)),
  br(),
  h3("Fiverr Strategy"),
  ...["Create 3–4 gigs specifically for real estate: 'Real Estate Agent Website', 'Property Listing App', 'Real Estate Landing Page'",
     "Use keyword-rich gig titles and descriptions targeting real estate buyer search terms",
     "Gig packages: Basic ($300 — 3-page site), Standard ($800 — 5-page + CMS), Premium ($2,000 — full listing platform)",
     "Add portfolio screenshots as gig gallery images — visual proof is everything on Fiverr",
     "Respond to ALL messages within 1 hour for first 90 days — response time affects ranking",
  ].map(b => bullet(b)),
  br(),
  h2("Weekly Lead Generation Goals"),
  twoCol([
    ["Week 1–2", "5 cold emails/day · 3 LinkedIn connections/day · 1 Upwork proposal/day"],
    ["Week 3–4", "8 cold emails/day · 5 LinkedIn connections/day · 3 Upwork proposals/day"],
    ["Week 5–8", "10 cold emails/day · 7 LinkedIn connections/day · 5 Upwork proposals/day"],
    ["Week 9–14", "10 cold emails/day · 5 LinkedIn connections/day · 5 proposals/day + 2 follow-ups on previous leads"],
  ], "Timeframe", "Daily Targets"),
];

// ─── PART 5: PERSONAL BRANDING ────────────────────────────────────────────────
const part5 = () => [
  pageBreak(),
  h1("PART 5: Personal Branding Strategy"),
  br(),
  callout("🏆 BRAND STATEMENT:", '"I am [Your Name] — a Full-Stack Developer who builds websites, apps, and digital tools exclusively for real estate businesses. I help agents generate more leads, brokerages scale their operations, and investors track their portfolios — through clean, fast, modern technology."', C.navy),
  br(),
  h2("Content Pillars (The 5 Topics You Own)"),
  twoCol([
    ["Pillar 1: My Learning Journey", "Daily progress updates, wins, struggles — humanizes you and builds trust"],
    ["Pillar 2: Real Estate Tech Education", "Teach real estate pros about technology — SEO, websites, tools, automation"],
    ["Pillar 3: Project Showcases", "Show your builds — before/after, demos, walkthroughs, tech breakdowns"],
    ["Pillar 4: Developer Tips & Tutorials", "Next.js, React, Tailwind tips — attracts developer peers and positions expertise"],
    ["Pillar 5: Freelance Business", "Proposals, pricing, client stories, lessons learned — builds authority and attracts clients"],
  ], "Content Pillar", "Why It Works"),
  br(),
  h2("Platform-by-Platform Strategy"),
  h3("LinkedIn (Primary Platform — Highest Client ROI)"),
  ...["Post format: Text posts with a strong hook (no external links in post body)",
     "Frequency: 3–4 posts/week",
     "Best content: Project launches, client results, developer lessons, real estate tech takes",
     "Connection strategy: Connect with real estate agents, brokers, property managers daily",
     "Engagement: Comment thoughtfully on 5 real estate professional posts daily",
  ].map(b => bullet(b)),
  h3("X / Twitter (Developer Community + Thought Leadership)"),
  ...["Post format: Threads for education, single tweets for quick takes, reply to industry conversations",
     "Frequency: 1–2 posts/day",
     "Best content: Code snippets, build-in-public updates, real estate tech opinions",
     "Follow: @levelsio (build in public), real estate influencers, PropTech accounts",
  ].map(b => bullet(b)),
  h3("Instagram + TikTok (Brand Awareness + Reach)"),
  ...["Content: Short-form videos showing your builds, coding time-lapses, before/after site transformations",
     "Frequency: 3–4 Reels/TikToks per week",
     "Hook formula: 'This $500 website is costing real estate agents thousands in lost leads' + show solution",
     "Trending sounds + developer/real estate niche hashtags",
  ].map(b => bullet(b)),
  h3("YouTube (Long-form Authority Building)"),
  ...["Content: Full project walkthroughs, real estate website tutorials, 100-day journey vlogs",
     "Frequency: 1 video/week (10–20 mins)",
     "SEO: 'how to build a real estate website', 'Next.js property listing tutorial', 'freelance developer journey'",
     "Long-form YouTube builds the deepest trust with potential clients",
  ].map(b => bullet(b)),
];

// ─── PART 6: CONTENT CALENDAR (SAMPLE 30 DAYS) ──────────────────────────────
const part6 = () => {
  const calendar = [
    { day:"1", platform:"LinkedIn", format:"Text Post", idea:"Why I'm spending 100 days becoming a real estate tech developer", hook:"I just made a decision that could change everything for my career.", cta:"Follow along — I'm documenting every day.", time:"8AM WAT" },
    { day:"2", platform:"X (Twitter)", format:"Thread", idea:"10 reasons real estate agents need better websites in 2026", hook:"Your website is losing you listings. Here's why:", cta:"RT if you know a real estate agent who needs this.", time:"7AM WAT" },
    { day:"3", platform:"Instagram/TikTok", format:"Reel", idea:"Day 3 of learning HTML — show your first webpage", hook:"Day 3 learning to code. Here's what I built today:", cta:"Follow for the full 100-day journey.", time:"6PM WAT" },
    { day:"4", platform:"LinkedIn", format:"Text Post", idea:"The problem with most real estate websites (and how I'll fix it)", hook:"I looked at 50 real estate websites yesterday. 43 of them had this same problem.", cta:"What's the #1 problem you see with real estate websites?", time:"8AM WAT" },
    { day:"5", platform:"X", format:"Single Tweet", idea:"Daily coding update — share what you built", hook:"Day 5 update: [what you built today]", cta:"See my GitHub for the code.", time:"9PM WAT" },
    { day:"6", platform:"LinkedIn", format:"Carousel", idea:"5 websites every real estate agent should have (and why)", hook:"5 pages every real estate agent's website needs — most have 0 of them.", cta:"Save this for your next website meeting.", time:"8AM WAT" },
    { day:"7", platform:"YouTube", format:"Vlog", idea:"Week 1 of my 100-day developer journey — what I learned", hook:"I spent 21 hours coding this week. Here's everything that happened.", cta:"Subscribe for weekly updates.", time:"12PM WAT" },
    { day:"8", platform:"LinkedIn", format:"Text Post", idea:"What I earned vs. what I want to earn — the freelance income gap", hook:"Right now I earn $0 as a developer. In 93 days I want my first client. This is my plan.", cta:"Drop a comment if you're on a similar journey.", time:"8AM WAT" },
    { day:"9", platform:"TikTok/Reels", format:"Short Video", idea:"Real estate website transformation — show a before/after concept", hook:"This is what a bad real estate website looks like. And this is what I'll build instead.", cta:"Follow to see the full build.", time:"6PM WAT" },
    { day:"10", platform:"LinkedIn", format:"Text Post", idea:"The math behind why real estate is the best niche for developers", hook:"One extra house sale = $15,000 commission. Your $2,000 website paid for itself on day one.", cta:"Which niche do YOU think is best for developers?", time:"8AM WAT" },
    { day:"14", platform:"LinkedIn", format:"Project Post", idea:"HTML/CSS complete — here's what I built in 2 weeks", hook:"2 weeks ago I wrote my first line of HTML. Here's what I can build today.", cta:"Like if you think progress like this is possible for anyone.", time:"8AM WAT" },
    { day:"21", platform:"LinkedIn + GitHub", format:"Portfolio Update", idea:"JavaScript complete — property search filter demo", hook:"I just built a working property search filter from scratch. No tutorials, just code.", cta:"Live demo in the comments. What should I build next?", time:"8AM WAT" },
    { day:"30", platform:"LinkedIn + YouTube", format:"Milestone Post", idea:"30 days in — React complete + first full project deployed", hook:"30 days. 1 full React project. 1 deployed app. Here's every lesson I learned.", cta:"Subscribe/follow — the backend journey starts tomorrow.", time:"8AM WAT" },
    { day:"60", platform:"All platforms", format:"Big Milestone Post", idea:"60 days in — full-stack app deployed, portfolio taking shape", hook:"60 days ago I didn't know what a server was. Today I deployed my first full-stack real estate app.", cta:"DM me 'site' if you're a real estate agent who wants a free website audit.", time:"8AM WAT" },
    { day:"90", platform:"All platforms", format:"Case Study Launch", idea:"My real estate property management dashboard — full breakdown", hook:"I spent 30 days building this. It solves a $50,000/year problem for property managers.", cta:"Full breakdown in my portfolio. Link in bio.", time:"8AM WAT" },
    { day:"100", platform:"All platforms", format:"Celebration + CTA", idea:"100 days done — here's everything I built and what comes next", hook:"100 days ago I started with nothing. Today I have a portfolio, a niche, and my first client conversations.", cta:"If you're a real estate professional who needs a website or app — let's talk. Link in bio.", time:"8AM WAT" },
  ];

  return [
    pageBreak(),
    h1("PART 6: 100-Day Content Calendar"),
    p("A strategic content plan to build authority, grow your audience, and attract real estate clients simultaneously with your technical learning journey."),
    br(),
    callout("📣 CONTENT STRATEGY:", "Post about your JOURNEY, not just your expertise. People connect with a developer learning in public. Your struggles and wins build more trust than polished expertise posts.", C.lightTeal, C.darkGray),
    br(),
    h2("Content Schedule Overview"),
    twoCol([
      ["LinkedIn", "3–4x/week — primary client acquisition channel"],
      ["X (Twitter)", "1–2x/day — developer community + build in public"],
      ["Instagram Reels", "3–4x/week — visual project showcases"],
      ["TikTok", "3–4x/week — same content as Reels (repurpose)"],
      ["YouTube", "1x/week — longer walkthroughs and vlogs"],
      ["Facebook/Threads", "Repurpose LinkedIn content — 1–2x/week"],
    ], "Platform", "Frequency & Purpose"),
    br(),
    h2("Sample Content Calendar — Key Milestones"),
    ...calendar.flatMap(item => [
      new Paragraph({ children: [
        new TextRun({ text: `Day ${item.day} — `, bold: true, font: "Arial", size: 22, color: C.blue }),
        new TextRun({ text: item.platform, bold: true, font: "Arial", size: 22, color: C.teal }),
        new TextRun({ text: ` | ${item.format}`, font: "Arial", size: 20, color: C.gray }),
      ], spacing: { before: 160, after: 60 } }),
      bullet(`Idea: ${item.idea}`),
      bullet(`Hook: "${item.hook}"`),
      bullet(`CTA: ${item.cta}`),
      bullet(`Best Time: ${item.time}`),
      br(),
    ]),
    callout("♻ REPURPOSING RULE:", "One LinkedIn post = 1 Twitter thread = 3 tweets = 1 Instagram carousel = 1 TikTok script. Create once, publish everywhere. Spend 20% of content time creating, 80% repurposing.", C.lightGold, C.darkGray),
  ];
};

// ─── PART 7: VIRAL CONTENT FRAMEWORKS ────────────────────────────────────────
const part7 = () => {
  const hooks = [
    "I looked at 100 real estate websites. 87 of them had this problem.",
    "Real estate agents are losing leads every day because of this website mistake.",
    "This $1,500 website generates $30,000/year in commission for a real estate agent.",
    "I spent 100 days learning to code. Here's what nobody tells you.",
    "The real estate tech industry is worth $4 trillion and most agents have a $200 website.",
    "This one page on your real estate website is losing you 40% of potential leads.",
    "I turned a real estate agent's website from 10 leads/month to 47 leads/month. Here's exactly how.",
    "Most real estate websites fail the 3-second test. Does yours?",
    "Why I quit applying for jobs and started building websites for real estate agents instead.",
    "The mortgage calculator that made a real estate developer $200k in pre-sales.",
  ];

  const reContentIdeas = [
    "5 PropTech tools every real estate agent should be using in 2026",
    "How AI is changing the way real estate listings are written",
    "The real estate website checklist: 10 things your site needs to generate leads",
    "Virtual tours vs. photos: what the data says about property engagement",
    "How real estate agents can use automation to follow up with leads faster",
    "The SEO strategy that puts real estate websites on the first page of Google",
    "Why 60% of real estate searches happen on mobile (and what that means for your site)",
    "CRM tools for real estate agents: free vs. paid breakdown",
    "How to use Google Analytics to track which listings get the most interest",
    "The landing page formula that converts property visitors into leads",
  ];

  const devIdeas = [
    "Build a mortgage calculator in React — full tutorial",
    "How to integrate Google Maps into a Next.js real estate app",
    "Building a property search filter with URL query params (Next.js)",
    "PostgreSQL database schema for a real estate platform — my approach",
    "How I use Tailwind CSS to build real estate UI components in 30 minutes",
    "TypeScript tips I wish I knew when I started building web apps",
    "Deploying a full-stack Next.js app to Vercel + Railway — step by step",
    "How I built an AI property description generator using the Claude API",
    "The tech stack I use for every real estate client project in 2026",
    "Building a real estate listing CMS with Sanity — beginner tutorial",
  ];

  return [
    pageBreak(),
    h1("PART 7: Viral Content Frameworks"),
    br(),
    h2("10 High-Performing Hooks (Real Estate Tech)"),
    ...hooks.map((h, i) => new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: `"${h}"`, font: "Arial", size: 22, color: C.darkGray })], spacing: { after: 80 } })),
    br(),
    h2("10 Real Estate Tech Content Ideas"),
    ...reContentIdeas.map((idea, i) => new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: idea, font: "Arial", size: 22, color: C.darkGray })], spacing: { after: 80 } })),
    br(),
    h2("10 Developer Content Ideas"),
    ...devIdeas.map((idea, i) => new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: idea, font: "Arial", size: 22, color: C.darkGray })], spacing: { after: 80 } })),
    br(),
    h2("Storytelling Template: The Transformation Story"),
    callout("TEMPLATE:", `Hook: [Before state — problem or pain point]
Bridge: I [action you took or discovered]
Story: [What happened — the journey with specific details]
Lesson: Here's what I learned: [3 bullet insights]
CTA: [Question to spark comments or action to take]

EXAMPLE: "This agent was getting 3 leads/month from their website. I rebuilt it in 3 weeks. Now they get 22/month. Here's the 5 changes that made the difference: [1, 2, 3, 4, 5]"`, C.lightGray, C.darkGray),
    br(),
    h2("Client Attraction Post Framework"),
    callout("TEMPLATE:", `Hook: Are you a real estate [agent/investor/broker] who [problem they have]?
Problem: Most [type] are losing [specific outcome] because [root cause].
Solution: I build [specific solution] that [specific result].
Proof: [Portfolio link or specific result]
CTA: DM me "[keyword]" and I'll send you a free [audit/guide/review] for your [website/system].`, C.lightBlue, C.navy),
  ];
};

// ─── PART 8: REVENUE ROADMAP ──────────────────────────────────────────────────
const part8 = () => [
  pageBreak(),
  h1("PART 8: Revenue Roadmap — $0 to $100k"),
  br(),
  h2("Stage 1: $0 → First Client (Days 1–100)"),
  twoCol([
    ["Services", "Simple agent websites, landing pages, property listing pages"],
    ["Pricing", "$500–$2,500 fixed price (discounted to build reviews)"],
    ["How to get it", "Cold email to local agents + Upwork + Fiverr + LinkedIn outreach"],
    ["Goal", "1 paid project. Not about the money — about the testimonial."],
    ["Key Action", "Offer first project at 50% rate in exchange for a written case study"],
  ], "Element", "Details"),
  br(),
  h2("Stage 2: $0 → $1,000/month (Month 2–4)"),
  twoCol([
    ["Services", "Agent sites + CMS ($800–$2,000) + monthly maintenance retainer ($300–$500/mo)"],
    ["Math", "2 project clients ($800 avg) + 1 retainer ($400) = $2,000 month"],
    ["Priority Action", "After every project, pitch a maintenance retainer immediately"],
    ["Raise Rate", "After 3 positive reviews: raise Upwork rate from $35 to $50/hr"],
    ["Goal", "2 active clients, 1 retainer, $1,000+ MRR"],
  ], "Element", "Details"),
  br(),
  h2("Stage 3: $1,000 → $3,000/month (Month 4–7)"),
  twoCol([
    ["Services", "Full listing platforms ($3,000–$5,000) + CRM builds + AI feature add-ons"],
    ["Math", "1 large project ($3,500) + 2 retainers ($600/mo each) = $4,700 month"],
    ["Priority Action", "Specialize messaging: 'I build tech for real estate businesses' — start turning down non-RE work"],
    ["Raise Rate", "New clients: $65–$80/hr. Productize with packages (Starter/Growth/Premium)"],
    ["Goal", "3 retainer clients, taking on premium projects only, $3,000+ MRR"],
  ], "Element", "Details"),
  br(),
  h2("Stage 4: $3,000 → $5,000+/month (Month 7–12)"),
  twoCol([
    ["Services", "Full-stack platforms ($6,000–$15,000) + AI SaaS tools + team/brokerage sites"],
    ["Math", "1 flagship project ($8,000) + 4 retainers ($750/mo each) = $11,000 month"],
    ["Priority Action", "Launch AI property description tool as SaaS — passive MRR on top of client work"],
    ["Raise Rate", "$85–$125/hr. Offer annual maintenance contracts (10% discount for upfront)"],
    ["Goal", "$5,000+ MRR, referral pipeline flowing, SaaS product generating $500–$2,000/mo"],
  ], "Element", "Details"),
  br(),
  h2("Packaging Strategy"),
  twoCol([
    ["STARTER — $1,500", "5-page Next.js site + CMS + contact form + basic SEO + 1 month support"],
    ["GROWTH — $3,500", "All Starter + property listings + inquiry system + Google Analytics + 3 months support"],
    ["PREMIUM — $7,500", "All Growth + custom features + AI integration + full SEO setup + 6 months retainer"],
    ["RETAINER ADD-ON", "$400–$800/month — updates, new pages, performance monitoring, bug fixes"],
    ["AI UPSELL", "$2,000–$4,000 add-on — integrate AI description generator or chatbot into any project"],
  ], "Package", "What's Included"),
];

// ─── PART 9: ACCOUNTABILITY DASHBOARD ─────────────────────────────────────────
const part9 = () => [
  pageBreak(),
  h1("PART 9: Weekly Accountability Dashboard"),
  br(),
  h2("Daily KPIs — Track Every Day"),
  twoCol([
    ["Hours coded", "Target: 3–5 hours minimum"],
    ["GitHub commits", "Target: 1+ commit per day"],
    ["Cold emails/DMs sent", "Target: 5–10 per day (from Day 60+)"],
    ["Content posted", "Target: 1 post on at least 1 platform"],
    ["LinkedIn connections", "Target: 5–10 new real estate professionals"],
    ["Upwork proposals", "Target: 3–5 per day (from Day 80+)"],
  ], "KPI", "Daily Target"),
  br(),
  h2("30-60-90-100 Day Milestones"),
  twoCol([
    ["Day 30 Checkpoint", "HTML, CSS, Tailwind, JS complete ✅ | First React app deployed ✅ | GitHub profile live ✅ | 500+ LinkedIn connections ✅"],
    ["Day 60 Checkpoint", "Next.js + Node.js + PostgreSQL complete ✅ | Full-stack app #1 deployed ✅ | Portfolio site live ✅ | 1,000+ LinkedIn connections ✅ | 10+ Upwork proposals sent ✅"],
    ["Day 90 Checkpoint", "All 6 skills complete ✅ | 3+ portfolio projects live ✅ | 50+ outreach emails sent ✅ | At least 3 client conversations started ✅ | First proposal submitted ✅"],
    ["Day 100 GOAL", "First paying client secured ✅ | $500–$2,500 project agreed ✅ | Testimonial earned ✅ | $1,000/month MRR path clear ✅"],
  ], "Milestone", "Success Criteria"),
  br(),
  h2("Weekly Review Template"),
  callout("WEEKLY CHECK-IN QUESTIONS:", `1. How many hours did I code this week? (Target: 21+ hrs)
2. What did I ship or deploy? (Target: 1 deliverable/week)
3. How many outreach messages did I send? (Target: 50+/week from Day 60)
4. What content did I post? (Target: 4+ pieces/week)
5. What was my biggest win this week?
6. What held me back? How do I fix it next week?
7. Am I on track for my Day 30/60/90/100 milestone?`, C.lightGray, C.darkGray),
  br(),
  callout("🚀 FINAL MESSAGE:", "The difference between developers who land clients in 100 days and those who don't is simple: daily action on ALL three tracks simultaneously — learning, building, and outreach. Don't wait until you're 'ready'. Start reaching out at Day 60. The confidence comes from doing, not from waiting.", C.navy),
];

// ─── BUILD DOCUMENT ──────────────────────────────────────────────────────────
async function buildDoc() {
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 22, color: C.darkGray } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: "Arial", color: C.navy },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: C.navy },
          paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: C.darkGray },
          paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
      ]
    },
    numbering: {
      config: [
        { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
        { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: { default: new Header({ children: [new Paragraph({ children: [
        new TextRun({ text: "100-Day Full-Stack Developer Roadmap | Real Estate Tech Specialization", font: "Arial", size: 18, color: C.gray })
      ], border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.lightGray, space: 4 } } })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ children: [
        new TextRun({ text: "© 100-Day Roadmap  |  ", font: "Arial", size: 18, color: C.gray }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: C.gray }),
      ], alignment: AlignmentType.CENTER })] }) },
      children: [
        ...coverPage(),
        ...part1(),
        ...part2(),
        ...part3(),
        ...part4(),
        ...part5(),
        ...part6(),
        ...part7(),
        ...part8(),
        ...part9(),
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("/mnt/user-data/outputs/100-Day-FullStack-Roadmap.docx", buffer);
  console.log("✅ Document written successfully");
}

buildDoc().catch(console.error);
