const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, TabStopType, BorderStyle,
  LevelFormat, ExternalHyperlink,
} = require('docx');

const OUT = '/home/user/CV/Kasib';

// ---------- Shared content ----------
const C = {
  name: 'Mohamed Kasib Farhan Cassim',
  location: 'United Arab Emirates',
  email: 'kasibfarhan@gmail.com',
  phone: '+971 XX XXX XXXX',
  linkedin: 'linkedin.com/in/mohamed-kasib-farhan-cassim-7b2a4222b',
  headline: 'AI Go-to-Market Lead  |  Business Development  |  B2B Enterprise Sales',
  summary:
    'AI go-to-market lead with B2B sales and business development experience across the Middle East and Southeast Asia. ' +
    'Built a MYR 6.2M+ qualified pipeline from 100+ enterprise accounts and now leads UAE market entry for Axsona, an AI platform ' +
    'that tests products and AI agents on synthetic populations before launch. Builds AI automations that make sales teams faster. ' +
    'MBA (CGPA 3.62); fluent in English, Hindi, Tamil and Urdu.',
  highlights: [
    ['MYR 6.2M+ qualified pipeline', ' built from 100+ high-profile enterprise accounts, from cold outreach to close.'],
    ['Cross-border expansion', ' across the Middle East and East Asia, managing business development teams in multiple countries.'],
    ['800+ SMEs', ' supported to build greener, more sustainable value chains through national sustainability programmes.'],
    ['AI-automated sales process', ': personally built an internal billing system and an AI help-desk agent (Slack + Jira).'],
  ],
  jobs: [
    {
      org: 'Axsona (Pantas Software Sdn. Bhd.)',
      loc: 'United Arab Emirates',
      roles: [{
        title: 'AI Go-to-Market Lead, Middle East',
        dates: 'Sep 2026 – Present',
        bullets: [
          'Lead Axsona’s end-to-end market entry into the UAE and wider Middle East, from ICP and segment prioritisation to positioning, partner model and pipeline targets.',
          'Building an AI-automated GTM engine (account research, lead enrichment, personalised multi-channel outreach and CRM updates) so a lean team can run outbound across several segments at once.',
          'Developing a two-track route to market: direct enterprise sales into banking, telecom, fintech, retail and real estate, plus a partner channel of agencies, consultancies and AI implementers.',
          'Turning buyer conversations and regional industry events into account-specific proposals, sales collateral and team enablement.',
          'Setting up the regional sales operating rhythm (pipeline stages, weekly activity targets, leadership reporting) to move from first pilots to repeatable revenue.',
        ],
      }],
    },
    {
      org: 'Pantas Software Sdn. Bhd.',
      loc: 'Kuala Lumpur, Malaysia',
      roles: [{
        title: 'Business Development & Go-to-Market Specialist',
        dates: 'Aug 2025 – Aug 2026',
        bullets: [
          'Owned the full sales cycle, building a MYR 6.2M+ (≈ USD 1.4M) qualified pipeline from 100+ high-profile enterprise accounts, from cold outreach to close.',
          'Ran product demos, led RFP/RFQ processes and scoped proposals that moved high-value prospects through to close.',
          'Automated the sales process with AI from prospecting to close; personally built an internal billing system and an AI help-desk agent (Slack + Jira).',
          'Drove sustainability enablement for 800+ SMEs, building greener value chains with the SFIA, GHB and GVC programmes.',
          'Shaped go-to-market strategy through strategic partnerships and outbound campaigns, opening new channels and high-value accounts.',
        ],
      }],
    },
    {
      org: 'BackwardsLA',
      loc: 'Remote (Indiana, USA)',
      roles: [
        {
          title: 'Strategy and Operations Lead',
          dates: 'Jun 2025 – Oct 2025',
          bullets: [
            'Led client acquisition and business development, including high-impact pitches across key verticals.',
            'Directed market expansion across the Middle East and East Asia, managing cross-border teams and hiring and onboarding business development staff.',
            'Product Owner for a Growth Program giving students real-world industry experience through hands-on projects and mentorship.',
            'Streamlined workflows between representatives and prospects; served as the link between regional operations and executive leadership.',
          ],
        },
        {
          title: 'Business Development Manager',
          dates: 'Mar 2025 – Jul 2025',
          bullets: [
            'Qualified and delivered high-quality leads to the sales team, consistently exceeding performance targets.',
            'Streamlined the lead hand-off with sales and marketing to improve client conversion.',
          ],
        },
      ],
    },
  ],
  education: [
    {
      org: 'Taylor’s University', loc: 'Subang Jaya, Malaysia',
      degree: 'Master of Business Administration (MBA)', dates: 'Sep 2025 – Aug 2026',
      notes: ['CGPA: 3.62 / 4.00'],
    },
    {
      org: 'Asia Pacific University of Technology & Innovation (APU)', short: 'Asia Pacific University (APU)', loc: 'Kuala Lumpur, Malaysia',
      degree: 'Bachelor of Business Management, E-Business', dates: 'Nov 2021 – 2024',
      notes: ['Activities: APU Muslim Association'],
    },
  ],
  skills: [
    ['Go-to-Market & Sales', 'GTM strategy, market entry, B2B enterprise sales, full-cycle sales, pipeline generation, account research, product demos, RFP/RFQ, proposals, strategic and channel partnerships, sales enablement, CRM'],
    ['AI & GTM Engineering', 'GTM engineering, outbound and sales process automation, lead enrichment, AI agents, generative AI (Claude, Gemini), Slack and Jira integrations'],
    ['Leadership', 'Cross-border team management, hiring and onboarding, stakeholder management, product ownership'],
  ],
  languages: 'English (native/bilingual); Hindi, Tamil, Urdu (full professional); Sinhala (professional working); Arabic (limited working)',
};

// ---------- Helpers ----------
const A4 = { width: 11906, height: 16838 };

function numbering(bulletIndent = 300) {
  return {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: bulletIndent, hanging: bulletIndent - 60 } } },
      }],
    }],
  };
}

function build(spec) {
  const margin = spec.margin;
  const textWidth = A4.width - margin.left - margin.right;
  const T = spec.theme;
  const run = (text, o = {}) => new TextRun({ text, font: T.font, size: T.size, color: T.ink, ...o });
  const rightTab = [{ type: TabStopType.RIGHT, position: textWidth }];

  const lineLR = (left, right, lo = {}, ro = {}, pOpts = {}) => new Paragraph({
    tabStops: rightTab, keepNext: true, ...pOpts,
    children: [run(left, lo), run('\t' + right, ro)],
  });
  const bullet = (children) => new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { after: T.bulletAfter ?? 20, line: T.line },
    children: typeof children === 'string' ? [run(children)] : children,
  });
  const heading = (text) => new Paragraph({
    keepNext: true,
    spacing: { before: T.headBefore, after: T.headAfter },
    alignment: T.headAlign || AlignmentType.LEFT,
    border: T.headRule === false ? undefined
      : { bottom: { style: BorderStyle.SINGLE, size: T.ruleSize || 6, color: T.ruleColor || T.accent, space: 1 } },
    children: [run(T.headCaps === false ? text : text.toUpperCase(), {
      bold: true, size: T.headSize, color: T.accent, font: T.headFont || T.font,
      smallCaps: !!T.smallCaps, characterSpacing: T.headSpacing || 0,
    })],
  });
  const para = (children, o = {}) => new Paragraph({ spacing: { after: 40, line: T.line }, ...o, children });
  const link = (text, url, o = {}) => new ExternalHyperlink({ link: url, children: [run(text, { color: T.ink, ...o })] });

  return { textWidth, run, lineLR, bullet, heading, para, link, rightTab, T };
}

function contactLines(h, sep, pOpts = {}) {
  const { run, link } = h;
  return [
    new Paragraph({ ...pOpts, spacing: { after: 0 }, children: [
      run(C.location), run(sep),
      link(C.email, 'mailto:' + C.email), run(sep),
      run(C.phone, { highlight: 'yellow' }),
    ] }),
    new Paragraph({ ...pOpts, spacing: { after: 40 }, children: [link(C.linkedin, 'https://www.' + C.linkedin)] }),
  ];
}

async function save(doc, file) {
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(path.join(OUT, file), buf);
  console.log('wrote', file);
}

function makeDoc(spec, children) {
  return new Document({
    creator: C.name,
    title: `${C.name} – CV`,
    styles: { default: { document: { run: { font: spec.theme.font, size: spec.theme.size } } } },
    numbering: numbering(spec.bulletIndent),
    sections: [{
      properties: { page: { size: A4, margin: spec.margin } },
      children,
    }],
  });
}

// ---------- 1. Harvard ----------
async function harvard() {
  const spec = {
    margin: { top: 850, bottom: 720, left: 900, right: 900 },
    bulletIndent: 320,
    theme: {
      font: 'Times New Roman', size: 21, ink: '000000', accent: '000000', line: 245,
      headSize: 22, headBefore: 150, headAfter: 70, ruleColor: '000000', ruleSize: 6, headAlign: AlignmentType.LEFT,
      bulletAfter: 10,
    },
  };
  const h = build(spec);
  const { run, lineLR, bullet, heading, para } = h;
  const ch = [];

  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 30 },
    children: [run(C.name.toUpperCase(), { bold: true, size: 32 })] }));
  ch.push(...contactLines(h, '  •  ', { alignment: AlignmentType.CENTER }));

  ch.push(heading('Education'));
  C.education.forEach((e, i) => {
    ch.push(lineLR(e.org, e.loc, { bold: true }, {}, { spacing: { before: i ? 80 : 0 } }));
    ch.push(lineLR(e.degree, e.dates, { italics: true }, {}));
    e.notes.forEach(n => ch.push(bullet(n)));
  });

  ch.push(heading('Experience'));
  C.jobs.forEach((j, i) => {
    ch.push(lineLR(j.org, j.loc, { bold: true }, {}, { spacing: { before: i ? 100 : 0 } }));
    j.roles.forEach((r, k) => {
      ch.push(lineLR(r.title, r.dates, { italics: true }, {}, { spacing: { before: k ? 50 : 0 } }));
      r.bullets.forEach(b => ch.push(bullet(b)));
    });
  });

  ch.push(heading('Skills, Languages & Interests'));
  C.skills.forEach(([k, v]) => ch.push(para([run(k + ': ', { bold: true }), run(v)], { spacing: { after: 20, line: 245 } })));
  ch.push(para([run('Languages: ', { bold: true }), run(C.languages)], { spacing: { after: 20, line: 245 } }));
  ch.push(para([run('Interests: ', { bold: true }), run('Applied AI and agentic systems, sustainability, startup ecosystems in the GCC and Southeast Asia')], { spacing: { after: 0, line: 245 } }));

  await save(makeDoc(spec, ch), 'Kasib_Cassim_CV_Harvard.docx');
}

// ---------- 2. Stanford ----------
async function stanford() {
  const CARDINAL = '8C1515';
  const spec = {
    margin: { top: 720, bottom: 640, left: 860, right: 860 },
    bulletIndent: 300,
    theme: {
      font: 'Cambria', size: 20, ink: '222222', accent: CARDINAL, line: 240,
      headSize: 21, headBefore: 130, headAfter: 60, ruleColor: CARDINAL, ruleSize: 8, smallCaps: false,
      headSpacing: 20, bulletAfter: 15,
    },
  };
  const h = build(spec);
  const { run, lineLR, bullet, heading, para } = h;
  const ch = [];

  ch.push(new Paragraph({ spacing: { after: 20 }, children: [run(C.name, { bold: true, size: 36, color: CARDINAL })] }));
  ch.push(new Paragraph({ spacing: { after: 30 }, children: [run('AI Go-to-Market Lead, Middle East', { italics: true, size: 20, color: '555555' })] }));
  ch.push(...contactLines(h, '   |   '));

  ch.push(heading('Summary'));
  ch.push(para([run(C.summary)], { spacing: { after: 20, line: 250 }, alignment: AlignmentType.JUSTIFIED }));

  ch.push(heading('Experience'));
  C.jobs.forEach((j, i) => {
    ch.push(lineLR(j.org.toUpperCase(), j.loc, { bold: true, size: 19, characterSpacing: 10 }, { italics: true, color: '555555' }, { spacing: { before: i ? 110 : 0 } }));
    j.roles.forEach((r, k) => {
      ch.push(lineLR(r.title, r.dates, { bold: true, color: CARDINAL }, { color: '555555' }, { spacing: { before: k ? 50 : 10, after: 10 } }));
      r.bullets.forEach(b => ch.push(bullet(b)));
    });
  });

  ch.push(heading('Education'));
  C.education.forEach((e, i) => {
    ch.push(lineLR(e.org.toUpperCase(), e.loc, { bold: true, characterSpacing: 10 }, { italics: true, color: '555555' }, { spacing: { before: i ? 80 : 0 } }));
    ch.push(lineLR(e.degree + '  ·  ' + e.notes[0].replace('Activities: ', ''), e.dates, {}, { color: '555555' }));
  });

  ch.push(heading('Skills & Languages'));
  C.skills.forEach(([k, v]) => ch.push(para([run(k + ': ', { bold: true }), run(v)], { spacing: { after: 20, line: 250 } })));
  ch.push(para([run('Languages: ', { bold: true }), run(C.languages)], { spacing: { after: 0, line: 250 } }));

  await save(makeDoc(spec, ch), 'Kasib_Cassim_CV_Stanford.docx');
}

// ---------- 3. Modern ATS (recommended for UAE / tech roles) ----------
async function modern() {
  const NAVY = '1F3A5F';
  const spec = {
    margin: { top: 680, bottom: 600, left: 780, right: 780 },
    bulletIndent: 280,
    theme: {
      font: 'Calibri', size: 20, ink: '1A1A1A', accent: NAVY, line: 240,
      headSize: 22, headBefore: 100, headAfter: 50, ruleColor: NAVY, ruleSize: 10, headSpacing: 30,
      bulletAfter: 15,
    },
  };
  const h = build(spec);
  const { run, lineLR, bullet, heading, para, textWidth } = h;
  const ch = [];

  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 10 },
    children: [run(C.name.toUpperCase(), { bold: true, size: 34, color: NAVY, characterSpacing: 30 })] }));
  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [run(C.headline, { bold: true, size: 21, color: '444444' })] }));
  ch.push(...contactLines(h, '  |  ', { alignment: AlignmentType.CENTER }));

  ch.push(heading('Professional Summary'));
  ch.push(para([run(C.summary)], { spacing: { after: 20, line: 250 }, alignment: AlignmentType.JUSTIFIED }));

  ch.push(heading('Key Achievements'));
  C.highlights.forEach(([b, rest]) => ch.push(bullet([run(b, { bold: true }), run(rest)])));

  ch.push(heading('Core Skills'));
  [...C.skills, ['Languages', C.languages]].forEach(([k, v]) => ch.push(para([run(k + ': ', { bold: true, color: NAVY }), run(v)], { spacing: { after: 20, line: 240 } })));

  ch.push(heading('Professional Experience'));
  C.jobs.forEach((j, i) => {
    j.roles.forEach((r, k) => {
      const first = k === 0;
      ch.push(lineLR(r.title, r.dates, { bold: true, size: 21, color: NAVY }, { bold: true },
        { spacing: { before: first ? (i ? 120 : 0) : 60 } }));
      if (first) ch.push(lineLR(j.org, j.loc, { bold: true, color: '444444' }, { italics: true, color: '444444' }, { spacing: { after: 20 } }));
      // Key Achievements already carries the headline metrics, so keep each role to its strongest points
      const skip = ['Turning buyer conversations', 'Shaped go-to-market', 'Streamlined workflows'];
      r.bullets.filter(b => !skip.some(x => b.startsWith(x))).forEach(b => ch.push(bullet(b)));
    });
  });

  ch.push(heading('Education'));
  C.education.forEach((e, i) => {
    ch.push(new Paragraph({ tabStops: h.rightTab, spacing: { before: i ? 40 : 0 }, children: [
      run(e.degree, { bold: true, color: NAVY }), run('  |  ' + (e.short || e.org) + ', ' + e.loc.split(',')[0]),
      run(i ? '' : '  |  ' + e.notes[0]), run('\t' + e.dates, { bold: true }),
    ] }));
  });

  await save(makeDoc(spec, ch), 'Kasib_Cassim_CV_Modern_ATS.docx');
}

(async () => { await harvard(); await stanford(); await modern(); })();
