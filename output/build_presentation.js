#!/usr/bin/env node
/**
 * build_presentation.js — Statistika Mistráky 2026
 *
 * Builds the full club presentation:
 *   - Season 2026 overview + per-team PLACEHOLDER slides (to be filled after last match)
 *   - Historical section: loyalty, leaderboard, milestones, records, Tříška comeback, closing
 *
 * Run: node build_presentation.js
 * Output: /home/claude/output/Statistika_mistraky_2026.pptx
 */

const PptxGenJS = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pres = new PptxGenJS();
pres.layout  = "LAYOUT_WIDE";   // 13.3" × 7.5"
pres.title   = "Statistika Mistráky 2026";
pres.author  = "TK Sport Kolovraty";

// ─── Color palette (NO # prefix!) ─────────────────────────────────────────
const C = {
  orange:       "E97132",
  orangeLight:  "FAE2D5",
  navy:         "0E2841",
  green:        "196B24",
  white:        "FFFFFF",
  black:        "000000",
  charcoal:     "1A1A1A",
  lgray:        "E8E8E8",
  mgray:        "999999",
  clay:         "C44B1A",      // deeper clay-court orange-red for accents
};

// ─── Image paths ────────────────────────────────────────────────────────────
const MEDIA = "/home/claude/pres_media/";
const logo       = MEDIA + "image16.jpg";   // club logo
const dronePic   = MEDIA + "image15.jpg";   // aerial court photo
const clayBall   = MEDIA + "image9.jpg";    // tennis ball on clay
const playerPic  = MEDIA + "image12.jpg";   // player near court

const OUT = "/home/claude/output/Statistika_mistraky_2026.pptx";
fs.mkdirSync("/home/claude/output", { recursive: true });

// ─── History data (career totals through 2025 + any scraped 2026) ──────────
const top15 = [
  { name: "Roušar Petr",          nick: "Péťa",           vazeno: 120.0,  seasons: 18, celkem: 187 },
  { name: "Uhlíř Vít",            nick: "Víťa",           vazeno: 105.0,  seasons: 15, celkem: 172 },
  { name: "Nováčková Karolína",   nick: "Kája",           vazeno: 60.0,   seasons: 10, celkem: 104 },
  { name: "Tříška Martin",        nick: "Tříška",         vazeno: 49.0,   seasons: 7,  celkem: 85  },
  { name: "Válek Milan",          nick: "Milan",          vazeno: 42.5,   seasons: 7,  celkem: 81  },
  { name: "Volf Jakub",           nick: "Kuba",           vazeno: 32.5,   seasons: 8,  celkem: 66  },
  { name: "Břicháček Bedřich",    nick: "Béďa",           vazeno: 30.5,   seasons: 5,  celkem: 63  },
  { name: "Kubín Marian",         nick: "Marian",         vazeno: 21.0,   seasons: 6,  celkem: 63  },
  { name: "Bubeník Jan",          nick: "Honza Bubeník",  vazeno: 19.5,   seasons: 9,  celkem: 69  },
  { name: "Uhlířová Veronika",    nick: "Verča",          vazeno: 15.0,   seasons: 3,  celkem: 30  },
  { name: "Suvorov Maksym",       nick: "Maks",           vazeno: 12.0,   seasons: 2,  celkem: 16  },
  { name: "Hýbl Jan st.",         nick: "Hýbl st.",       vazeno: 7.5,    seasons: 2,  celkem: 13  },
  { name: "Tichá Lenka",          nick: "Lenka T.",       vazeno: 6.5,    seasons: 2,  celkem: 24  },
  { name: "Pikula Tomáš",         nick: "Tomáš P.",       vazeno: 6.0,    seasons: 2,  celkem: 10  },
  { name: "Vydlák Šimon",         nick: "Šimon",          vazeno: 5.5,    seasons: 7,  celkem: 29  },
];

const loyalty = [
  { name: "Roušar Petr",        nick: "Péťa",        seasons: 18, celkem: 187, team: "A, C" },
  { name: "Uhlíř Vít",          nick: "Víťa",        seasons: 15, celkem: 172, team: "A, C" },
  { name: "Nováčková Karolína", nick: "Kája",        seasons: 10, celkem: 104, team: "A"    },
  { name: "Bubeník Jan",        nick: "Honza",       seasons: 9,  celkem: 69,  team: "C"    },
  { name: "Volf Jakub",         nick: "Kuba",        seasons: 8,  celkem: 66,  team: "A, C" },
];

// ─── Shared helpers ─────────────────────────────────────────────────────────

function sectionDivider(title, subtitle, bgImage) {
  const slide = pres.addSlide();
  // Dark overlay shape
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: "100%", h: "100%",
    fill: { color: C.navy },
  });
  // Background image semi-transparent
  slide.addImage({
    path: bgImage, x: 0, y: 0, w: 13.3, h: 7.5,
    transparency: 65,
  });
  // Title
  slide.addText(title, {
    x: 0.8, y: 2.4, w: 11.7, h: 1.8,
    fontSize: 54, bold: true, color: C.white,
    fontFace: "Cambria", align: "center", valign: "middle",
  });
  // Subtitle
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.8, y: 4.2, w: 11.7, h: 0.7,
      fontSize: 22, color: C.orange,
      fontFace: "Calibri", align: "center", charSpacing: 4,
    });
  }
  // Logo bottom-right
  slide.addImage({ path: logo, x: 11.4, y: 6.1, w: 1.6, h: 1.2 });
  return slide;
}

function placeholder(title, teamColor, body) {
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // Top banner shape
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.3, h: 1.2,
    fill: { color: teamColor || C.orange }, line: { color: teamColor || C.orange },
  });
  slide.addText(title, {
    x: 0.4, y: 0, w: 11.5, h: 1.2,
    fontSize: 26, bold: true, color: C.white,
    fontFace: "Cambria", valign: "middle",
  });
  // Logo in banner
  slide.addImage({ path: logo, x: 11.8, y: 0.05, w: 1.3, h: 1.05 });

  // Placeholder content
  slide.addShape(pres.ShapeType.roundRect, {
    x: 1.5, y: 1.9, w: 10.3, h: 4.5,
    fill: { color: C.lgray }, line: { color: "CCCCCC", width: 1 },
    rectRadius: 0.15,
  });
  slide.addText([
    { text: "⏳  DATA BUDE DOPLNĚNA", options: { bold: true, fontSize: 22, color: C.charcoal, breakLine: true } },
    { text: "po posledním kole sezóny 2026", options: { fontSize: 16, color: C.mgray, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: body || "", options: { fontSize: 14, color: C.charcoal } },
  ], {
    x: 1.5, y: 1.9, w: 10.3, h: 4.5,
    align: "center", valign: "middle",
  });
  return slide;
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  // Left: big title text
  s.addText([
    { text: "Statistika", options: { breakLine: true } },
    { text: "Mistráky", options: { breakLine: true } },
    { text: "2026", options: { color: C.orange } },
  ], {
    x: 0.7, y: 1.4, w: 6.0, h: 4.2,
    fontSize: 58, bold: true, color: C.charcoal,
    fontFace: "Cambria", valign: "middle",
  });

  // Subtitle
  s.addText("TK Sport Kolovraty", {
    x: 0.7, y: 5.8, w: 6.0, h: 0.7,
    fontSize: 18, color: C.mgray, fontFace: "Calibri", charSpacing: 3,
  });

  // Right: logo + clay background image
  s.addImage({ path: dronePic, x: 6.9, y: 0, w: 6.4, h: 7.5 });
  // Dark overlay on right side for contrast
  s.addShape(pres.ShapeType.rect, {
    x: 6.9, y: 0, w: 6.4, h: 7.5,
    fill: { color: C.navy, transparency: 30 }, line: { color: C.navy, transparency: 30 },
  });
  s.addImage({ path: logo, x: 8.9, y: 1.4, w: 2.5, h: 2.0 });
  s.addText("TK SPORT KOLOVRATY", {
    x: 7.2, y: 3.5, w: 5.7, h: 0.6,
    fontSize: 12, bold: true, color: C.white, align: "center", charSpacing: 4, fontFace: "Calibri",
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 2 — SEASON 2026 OVERVIEW (placeholder)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  // Left photo
  s.addImage({ path: clayBall, x: 0, y: 0, w: 5.2, h: 7.5 });

  // Right content
  s.addText("Shrnutí sezóny 2026", {
    x: 5.6, y: 0.6, w: 7.3, h: 0.9,
    fontSize: 32, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addShape(pres.ShapeType.line, {
    x: 5.6, y: 1.55, w: 6.8, h: 0, line: { color: C.orange, width: 3 },
  });

  const teams2026 = [
    { label: "Tým A", comp: "Pražská divize",   result: "[VÝSLEDEK — doplnit po posledním kole]" },
    { label: "Tým B", comp: "2. třída C",        result: "[VÝSLEDEK — doplnit po posledním kole]" },
    { label: "Tým C", comp: "3. třída B",        result: "[VÝSLEDEK — doplnit po posledním kole]" },
    { label: "Tým D", comp: "4. třída A",        result: "[VÝSLEDEK — doplnit po posledním kole]" },
    { label: "Dorost",   comp: "2. třída A",     result: "[VÝSLEDEK — doplnit po posledním kole]" },
    { label: "Mládež",   comp: "A + B",          result: "[VÝSLEDEK — doplnit po posledním kole]" },
  ];

  let yy = 1.85;
  for (const t of teams2026) {
    s.addShape(pres.ShapeType.roundRect, {
      x: 5.6, y: yy, w: 7.3, h: 0.72,
      fill: { color: C.orangeLight }, line: { color: C.orange, width: 0.5 }, rectRadius: 0.08,
    });
    s.addText([
      { text: `${t.label}  `, options: { bold: true, color: C.orange } },
      { text: `(${t.comp})  `, options: { color: C.charcoal, fontSize: 13 } },
      { text: t.result, options: { color: C.mgray, italic: true, fontSize: 12 } },
    ], { x: 5.75, y: yy, w: 7.1, h: 0.72, valign: "middle", fontSize: 14 });
    yy += 0.82;
  }
  // Logo
  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDES 3–6 — PER-TEAM PLACEHOLDERS (Dospělí A, B, C, D)
// ═══════════════════════════════════════════════════════════════════════════
const teamColors = ["156082", "E97132", "196B24", "6B196B"];
const teamDefs = [
  { label: "Tým A — Pražská divize",   note: "Ocenění, statistiky a callouts budou doplněny po skončení sezóny." },
  { label: "Tým B — 2. třída C",       note: "Ocenění, statistiky a callouts budou doplněny po skončení sezóny." },
  { label: "Tým C — 3. třída B",       note: "Ocenění, statistiky a callouts budou doplněny po skončení sezóny." },
  { label: "Tým D — 4. třída A",       note: "Ocenění, statistiky a callouts budou doplněny po skončení sezóny." },
];
teamDefs.forEach((t, i) => placeholder(t.label, teamColors[i], t.note));

// SLIDE 7 — YOUTH PLACEHOLDER
placeholder("Mládež 2026 — Dorost · Mladší žáci · Babytenis", "467886",
  "Výsledky a statistiky všech mládežnických kategorií budou doplněny po posledním kole.");

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 8 — SECTION DIVIDER: HISTORICKÝ PŘEHLED
// ═══════════════════════════════════════════════════════════════════════════
sectionDivider(
  "Historický přehled",
  "TK Sport Kolovraty · dospělí",
  dronePic
);

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 — VĚRNOST KLUBU (Loyalty)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Věrnost klubu", {
    x: 0.6, y: 0.3, w: 8.0, h: 0.85,
    fontSize: 38, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addText("Hráči s nejvíce sezónami v dospělé soutěži za TK Sport Kolovraty", {
    x: 0.6, y: 1.1, w: 9.0, h: 0.5,
    fontSize: 15, color: C.mgray, fontFace: "Calibri",
  });
  s.addShape(pres.ShapeType.line, {
    x: 0.6, y: 1.6, w: 12.0, h: 0, line: { color: C.lgray, width: 1 },
  });

  // 5 loyalty cards
  const CARD_W = 2.34, CARD_GAP = 0.13, START_X = 0.6;
  const cardColors = ["E97132", "E97132", "C44B1A", "C44B1A", "8B3A10"];
  const medalEmoji = ["🥇", "🥈", "🥉", "4.", "5."];

  loyalty.forEach((p, i) => {
    const cx = START_X + i * (CARD_W + CARD_GAP);

    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 1.85, w: CARD_W, h: 5.0,
      fill: { color: cardColors[i] }, line: { color: cardColors[i] },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.18 },
      rectRadius: 0.15,
    });

    // Rank
    s.addText(medalEmoji[i], {
      x: cx, y: 2.0, w: CARD_W, h: 0.75,
      fontSize: 28, align: "center", color: C.white, fontFace: "Calibri",
    });

    // Seasons number — big
    s.addText(String(p.seasons + 1), {   // +1 for 2026
      x: cx, y: 2.75, w: CARD_W, h: 1.5,
      fontSize: 72, bold: true, align: "center", color: C.white, fontFace: "Cambria", valign: "middle",
    });

    s.addText("sezón", {
      x: cx, y: 4.3, w: CARD_W, h: 0.45,
      fontSize: 14, align: "center", color: C.white, fontFace: "Calibri",
    });

    // Name
    s.addText(p.nick, {
      x: cx, y: 4.75, w: CARD_W, h: 0.6,
      fontSize: 18, bold: true, align: "center", color: C.white, fontFace: "Cambria",
    });

    // Team + matches
    s.addText(`${p.celkem}+ zápasů · Tým ${p.team}`, {
      x: cx, y: 5.4, w: CARD_W, h: 0.45,
      fontSize: 11, align: "center", color: "FFE0CC", fontFace: "Calibri",
    });
  });

  s.addText("* počet sezón včetně sezóny 2026", {
    x: 0.6, y: 7.1, w: 10, h: 0.35,
    fontSize: 10, color: C.mgray, italic: true,
  });
  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 10 — HISTORICKÝ ŽEBŘÍČEK (native bar chart)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Historický žebříček", {
    x: 0.5, y: 0.2, w: 8.5, h: 0.75,
    fontSize: 34, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addText("Celkem bodů pro klub · dospělí · kariéra", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 14, color: C.mgray, fontFace: "Calibri",
  });

  // Top 12 for chart readability
  const chartData = top15.slice(0, 12);

  s.addChart(pres.charts.BAR, [{
    name: "Body Váženo",
    labels: chartData.map(p => p.nick),
    values: chartData.map(p => p.vazeno),
  }], {
    x: 0.3, y: 1.3, w: 8.8, h: 5.8, barDir: "bar",
    chartColors: [C.orange],
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    plotArea: { fill: { color: C.white } },
    catAxisLabelColor: C.charcoal,
    valAxisLabelColor: C.mgray,
    catAxisLabelFontSize: 13,
    valAxisLabelFontSize: 11,
    valGridLine: { color: C.lgray, size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelColor: C.charcoal,
    dataLabelFontSize: 11,
    showLegend: false,
    barGapWidthPct: 55,
  });

  // Note
  s.addText("Hodnoty zahrnují sezóny dospělých · Váženo: singl = 1 bod, debl = 0,5 bodu", {
    x: 0.3, y: 7.1, w: 10, h: 0.35,
    fontSize: 10, color: C.mgray, italic: true,
  });

  // Right info cards
  const cards = [
    { label: "1. Roušar Petr", value: "120 bodů", sub: "87 singl + 66×½ debl" },
    { label: "2. Uhlíř Vít",   value: "105 bodů", sub: "72 singl + 66×½ debl" },
    { label: "3. Nováčková K.", value: "60 bodů",  sub: "47 singl + 26×½ debl" },
  ];

  let cy = 1.6;
  for (const c of cards) {
    s.addShape(pres.ShapeType.roundRect, {
      x: 9.4, y: cy, w: 3.6, h: 1.35,
      fill: { color: C.orangeLight }, line: { color: C.orange, width: 1 }, rectRadius: 0.1,
    });
    s.addText([
      { text: c.label + "\n", options: { bold: true, fontSize: 13, color: C.charcoal } },
      { text: c.value + "\n", options: { bold: true, fontSize: 22, color: C.orange } },
      { text: c.sub, options: { fontSize: 11, color: C.mgray } },
    ], { x: 9.55, y: cy + 0.1, w: 3.3, h: 1.15, valign: "middle" });
    cy += 1.5;
  }

  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 11 — PÉŤA 200 ZÁPASŮ MILESTONE
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  // Dark clay background
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: C.navy }, line: { color: C.navy },
  });

  // Player photo on right
  s.addImage({ path: playerPic, x: 7.8, y: 0, w: 5.5, h: 7.5 });
  s.addShape(pres.ShapeType.rect, {
    x: 7.8, y: 0, w: 5.5, h: 7.5,
    fill: { color: C.navy, transparency: 55 }, line: { color: C.navy, transparency: 55 },
  });

  // Left content
  s.addText("MILNÍK", {
    x: 0.7, y: 0.5, w: 7.0, h: 0.6,
    fontSize: 14, bold: true, color: C.orange, fontFace: "Calibri", charSpacing: 8,
  });

  s.addText("200+", {
    x: 0.5, y: 0.9, w: 7.0, h: 2.6,
    fontSize: 110, bold: true, color: C.white, fontFace: "Cambria", valign: "middle",
  });

  s.addText("odehraných zápasů", {
    x: 0.7, y: 3.4, w: 7.0, h: 0.75,
    fontSize: 26, color: C.orange, fontFace: "Cambria",
  });

  s.addText("Roušar Petr", {
    x: 0.7, y: 4.15, w: 7.0, h: 0.65,
    fontSize: 24, bold: true, color: C.white, fontFace: "Cambria",
  });

  s.addText(
    "Péťa překročí hranici 200 odehraných zápasů v dospělé soutěži " +
    "za klub v posledním kole sezóny 2026 — první hráč v historii " +
    "TK Sport Kolovraty, kdo toho dosáhl.",
    {
      x: 0.7, y: 4.85, w: 7.0, h: 1.4,
      fontSize: 14, color: "AACCEE", fontFace: "Calibri", align: "left",
    }
  );

  // Career stats bottom
  const stats = [
    ["18 + sezón", "v dresu klubu"],
    ["153", "bodů pro klub"],
    ["90,6 %", "úspěšnost v singlu"],
  ];
  let sx = 0.7;
  for (const [val, lbl] of stats) {
    s.addText(val, {
      x: sx, y: 6.2, w: 2.1, h: 0.6,
      fontSize: 20, bold: true, color: C.orange, fontFace: "Cambria", align: "center",
    });
    s.addText(lbl, {
      x: sx, y: 6.8, w: 2.1, h: 0.4,
      fontSize: 11, color: "AACCEE", fontFace: "Calibri", align: "center",
    });
    sx += 2.3;
  }

  s.addImage({ path: logo, x: 11.6, y: 6.4, w: 1.4, h: 1.1 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 12 — SEZÓNA 2026 · SPECIÁLNÍ REKORDY
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Sezóna 2026 · Speciální záznamy", {
    x: 0.6, y: 0.25, w: 12.0, h: 0.8,
    fontSize: 34, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addShape(pres.ShapeType.line, {
    x: 0.6, y: 1.1, w: 12.0, h: 0, line: { color: C.lgray, width: 1 },
  });

  const records = [
    {
      icon: "💯", title: "Téměř dokonalá sezóna",
      who: "Suvorov Maksym (Maks)",
      text: "100% úspěšnost v singlu a pouze jedna ztráta v deblu. Nejlepší poměr výher sezóny 2026.",
    },
    {
      icon: "🏆", title: "Nejvíce zápasů v sezóně",
      who: "[DOPLNIT — TBD]",
      text: "Hráč s absolutně největším počtem odehraných zápasů v sezóně 2026 (dospělí, všechny týmy).",
    },
    {
      icon: "🎯", title: "Bezporážkový v sezóně",
      who: "[DOPLNIT — TBD]",
      text: "Hráč/hráčka, kteří odehráli v sezóně 2026 zápasy bez jediné prohry.",
    },
    {
      icon: "🔄", title: "Návrat legendy",
      who: "Tříška Martin",
      text: "Po 13 letech pauzy se vrátil na soupisce a opět hraje v dresu TK Sport Kolovraty.",
    },
  ];

  const cols = [[0, 1], [2, 3]];
  const COL_X = [0.6, 7.0];
  const ROW_Y = [1.35, 4.1];

  for (let ci = 0; ci < 2; ci++) {
    for (let ri = 0; ri < 2; ri++) {
      const rec = records[ci * 2 + ri];
      const cx = COL_X[ci], cy = ROW_Y[ri];

      s.addShape(pres.ShapeType.roundRect, {
        x: cx, y: cy, w: 6.1, h: 2.5,
        fill: { color: C.orangeLight }, line: { color: C.orange, width: 1 }, rectRadius: 0.12,
        shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.1 },
      });
      s.addText([
        { text: rec.icon + "  " + rec.title + "\n", options: { bold: true, fontSize: 16, color: C.orange } },
        { text: rec.who + "\n", options: { bold: true, fontSize: 14, color: C.charcoal } },
        { text: rec.text, options: { fontSize: 12, color: C.charcoal } },
      ], { x: cx + 0.2, y: cy + 0.15, w: 5.75, h: 2.2, valign: "top" });
    }
  }

  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 13 — KARIÉRNÍ ZÁZNAMY
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Kariérní záznamy", {
    x: 0.6, y: 0.25, w: 9.0, h: 0.8,
    fontSize: 34, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addText("Všechny záznamy týkají se dospělé soutěže — dohromady za všechny týmy", {
    x: 0.6, y: 1.0, w: 12.0, h: 0.45,
    fontSize: 13, color: C.mgray, fontFace: "Calibri",
  });

  const recItems = [
    { label: "Nejvíce odehraných zápasů",     value: "187+",    who: "Roušar Petr",     sub: "(překročí 200 v posledním kole 2026)" },
    { label: "Nejvíce sezón",                  value: "18+",     who: "Roušar Petr",     sub: "(s aktuální sezónou)" },
    { label: "Nejvíce singl výher v kariéře",  value: "87",      who: "Roušar Petr",     sub: "(z 96 odehraných singlů)" },
    { label: "Nejlepší kariérní singl %",      value: "90,6 %",  who: "Roušar Petr",     sub: "(min. 50 odehraných singlů)" },
    { label: "Nejvíce debl výher v kariéře",   value: "66",      who: "Roušar Petr  &  Uhlíř Vít", sub: "(oba shodně)" },
    { label: "Nejlepší kariérní debl %",       value: "78,6 %",  who: "Uhlíř Vít",       sub: "(min. 50 odehraných deblů)" },
    { label: "Ø bodů za sezónu (min. 5 sez.)", value: "7,0",    who: "Uhlíř Vít  &  Tříška Martin", sub: "(shodně; Péťa = 6,7)" },
    { label: "Nejlepší kariérní debl %",       value: "90,4 %",  who: "Nováčková Karolína", sub: "(singl, min. 50 odehraných)" },
  ];

  // Two-column table
  const colW = 6.0;
  const colX = [0.6, 6.9];
  let rr = 0;
  for (let ci = 0; ci < 2; ci++) {
    let yy = 1.6;
    for (let ri = 0; ri < 4; ri++) {
      const item = recItems[ci * 4 + ri];
      if (!item) break;
      const bg = ri % 2 === 0 ? C.white : "F7F7F7";
      s.addShape(pres.ShapeType.rect, {
        x: colX[ci], y: yy, w: colW, h: 1.1,
        fill: { color: bg }, line: { color: C.lgray, width: 0.5 },
      });
      s.addText([
        { text: item.value + "  ", options: { bold: true, fontSize: 20, color: C.orange, fontFace: "Cambria" } },
        { text: item.who + "\n", options: { bold: true, fontSize: 13, color: C.charcoal } },
        { text: item.label + "   ", options: { fontSize: 11, color: C.mgray } },
        { text: item.sub, options: { fontSize: 10, color: C.mgray, italic: true } },
      ], { x: colX[ci] + 0.1, y: yy + 0.05, w: colW - 0.2, h: 1.0, valign: "middle" });
      yy += 1.18;
    }
  }

  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 14 — TŘÍŠKA MARTIN — NÁVRAT LEGENDY
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: "F7F3EE" };  // warm cream inspired by clay

  // Left dark panel
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 5.3, h: 7.5,
    fill: { color: C.navy }, line: { color: C.navy },
  });

  // Timeline on dark panel
  s.addText("KARIÉRNÍ PŘEHLED", {
    x: 0.25, y: 0.4, w: 4.8, h: 0.55,
    fontSize: 13, bold: true, color: C.orange, fontFace: "Calibri",
    charSpacing: 5, align: "center",
  });

  const timeline = [
    { year: "2006",  text: "První sezóna za klub", highlight: false },
    { year: "2007",  text: "7/7 singl výher — perfektní sezóna", highlight: true },
    { year: "2008 – 2012", text: "5 dalších sezón, celkem 7 let", highlight: false },
    { year: "2013 – 2025", text: "13 let pauzy", highlight: false, pause: true },
    { year: "2026 ↩",  text: "Návrat! — Tým A · TK Sport Kolovraty", highlight: true },
  ];

  let ty = 1.15;
  for (const item of timeline) {
    const isGap = item.pause;
    s.addShape(pres.ShapeType.ellipse, {
      x: 0.55, y: ty + 0.02, w: 0.22, h: 0.22,
      fill: { color: isGap ? "445566" : C.orange }, line: { color: "FFFFFF" },
    });
    s.addText(item.year, {
      x: 0.9, y: ty, w: 1.5, h: 0.32,
      fontSize: 11, bold: true, color: isGap ? "AABBCC" : C.orange, fontFace: "Calibri",
    });
    s.addText(item.text, {
      x: 0.9, y: ty + 0.28, w: 4.15, h: 0.42,
      fontSize: 11, color: isGap ? "778899" : "CCE0F0", fontFace: "Calibri",
      italic: isGap,
    });
    ty += isGap ? 0.95 : 0.82;
    if (!isGap) {
      s.addShape(pres.ShapeType.line, {
        x: 0.65, y: ty - 0.35, w: 0, h: 0.28,
        line: { color: "334466", width: 1, dashType: "dot" },
      });
    }
  }

  // Right content
  s.addText("Tříška Martin", {
    x: 5.6, y: 0.5, w: 7.4, h: 0.9,
    fontSize: 40, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addText("Návrat legendy", {
    x: 5.6, y: 1.35, w: 7.4, h: 0.65,
    fontSize: 22, color: C.orange, fontFace: "Cambria",
  });
  s.addShape(pres.ShapeType.line, {
    x: 5.6, y: 2.1, w: 7.2, h: 0, line: { color: C.lgray, width: 1 },
  });

  s.addText(
    "Po 13 letech pauzy se Tříška Martin vrátil na kurty TK Sport Kolovraty " +
    "a nastoupil v sezóně 2026 znovu v dospělé soutěži.\n\n" +
    "V letech 2006–2012 odehrál 85 zápasů za klub, z toho ve své nejlepší " +
    "sezóně 2007 vyhrál všechny 7 singlů (100%).",
    {
      x: 5.6, y: 2.3, w: 7.1, h: 2.0,
      fontSize: 15, color: C.charcoal, fontFace: "Calibri",
    }
  );

  // Career stat cards
  const cstats = [
    ["85", "kariérních zápasů", "2006–2012"],
    ["7", "sezón", "za klub"],
    ["65", "kariérních výher", "(33 singl + 32 debl)"],
  ];
  let csx = 5.6;
  for (const [val, lbl, sub] of cstats) {
    s.addShape(pres.ShapeType.roundRect, {
      x: csx, y: 4.55, w: 2.25, h: 2.3,
      fill: { color: C.orangeLight }, line: { color: C.orange, width: 1 }, rectRadius: 0.12,
    });
    s.addText(val, {
      x: csx, y: 4.75, w: 2.25, h: 1.0,
      fontSize: 44, bold: true, color: C.orange, fontFace: "Cambria", align: "center", valign: "middle",
    });
    s.addText(lbl, {
      x: csx, y: 5.8, w: 2.25, h: 0.4,
      fontSize: 13, bold: true, color: C.charcoal, align: "center",
    });
    s.addText(sub, {
      x: csx, y: 6.2, w: 2.25, h: 0.35,
      fontSize: 11, color: C.mgray, align: "center", italic: true,
    });
    csx += 2.45;
  }

  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 15 — PRŮMĚRNÁ VÝKONNOST (Efficiency per season)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Průměrná výkonnost za sezónu", {
    x: 0.5, y: 0.2, w: 9.5, h: 0.8,
    fontSize: 32, bold: true, color: C.charcoal, fontFace: "Cambria",
  });
  s.addText("Ø Body Váženo za sezónu · hráči s alespoň 5 sezónami · dospělí", {
    x: 0.5, y: 0.95, w: 12.0, h: 0.45,
    fontSize: 13, color: C.mgray, fontFace: "Calibri",
  });

  const effData = [
    { nick: "Víťa",   avg: 7.00, sez: 15 },
    { nick: "Tříška", avg: 7.00, sez: 7 },
    { nick: "Péťa",   avg: 6.67, sez: 18 },
    { nick: "Béďa",   avg: 6.10, sez: 5 },
    { nick: "Milan",  avg: 6.07, sez: 7 },
    { nick: "Kája",   avg: 6.00, sez: 10 },
    { nick: "Kuba",   avg: 4.06, sez: 8 },
    { nick: "Marian", avg: 3.50, sez: 6 },
    { nick: "Bubeník", avg: 2.17, sez: 9 },
    { nick: "Šimon",  avg: 0.79, sez: 7 },
  ];

  s.addChart(pres.charts.BAR, [{
    name: "Ø Váženo / sezóna",
    labels: effData.map(p => p.nick),
    values: effData.map(p => p.avg),
  }], {
    x: 0.3, y: 1.4, w: 8.2, h: 5.7, barDir: "bar",
    chartColors: [C.clay],
    chartArea: { fill: { color: C.white } },
    catAxisLabelColor: C.charcoal, catAxisLabelFontSize: 13,
    valAxisLabelColor: C.mgray, valAxisLabelFontSize: 11,
    valGridLine: { color: C.lgray, size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelFontSize: 11, dataLabelColor: C.charcoal,
    showLegend: false, barGapWidthPct: 55,
  });

  // Callout — interesting fact
  s.addShape(pres.ShapeType.roundRect, {
    x: 8.9, y: 1.5, w: 4.1, h: 3.2,
    fill: { color: C.orangeLight }, line: { color: C.orange, width: 1 }, rectRadius: 0.12,
  });
  s.addText([
    { text: "💡  Zaujímavost\n\n", options: { bold: true, fontSize: 15, color: C.orange } },
    { text: "Víťa a Tříška Martin mají shodný průměr 7,0 bodů za sezónu — nejvyšší v historii klubu.\n\n", options: { fontSize: 13, color: C.charcoal } },
    { text: "Tříška přitom odehrál všechny sezóny v jediné éře (2006–2012).", options: { fontSize: 12, color: C.charcoal } },
  ], { x: 9.1, y: 1.7, w: 3.7, h: 2.8, valign: "top" });

  s.addShape(pres.ShapeType.roundRect, {
    x: 8.9, y: 4.9, w: 4.1, h: 2.3,
    fill: { color: "EBF3FF" }, line: { color: "4A90D9", width: 1 }, rectRadius: 0.12,
  });
  s.addText([
    { text: "📊  Konzistence\n\n", options: { bold: true, fontSize: 15, color: "1A6FBF" } },
    { text: "Péťa má nejvyšší CELKOVÝ počet bodů (120) díky rekordním 18 sezónám.", options: { fontSize: 12, color: C.charcoal } },
  ], { x: 9.1, y: 5.1, w: 3.7, h: 2.0, valign: "top" });

  s.addImage({ path: logo, x: 11.7, y: 6.5, w: 1.35, h: 1.0 });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 16 — CLOSING (full-bleed drone photo)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.addImage({ path: dronePic, x: 0, y: 0, w: 13.3, h: 7.5 });
  // Overlay for readability
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.3, h: 7.5,
    fill: { color: C.navy, transparency: 45 }, line: { color: C.navy, transparency: 45 },
  });
  s.addText("Díky za sezónu 2026!", {
    x: 1.0, y: 2.0, w: 11.3, h: 2.0,
    fontSize: 56, bold: true, color: C.white, fontFace: "Cambria", align: "center", valign: "middle",
  });
  s.addText("TK Sport Kolovraty · Praha 22 · Kolovraty", {
    x: 1.0, y: 4.2, w: 11.3, h: 0.65,
    fontSize: 18, color: C.orange, fontFace: "Calibri", align: "center", charSpacing: 4,
  });
  s.addText("⟳ Vizualizace a historická data: club_history_interactive.html", {
    x: 1.0, y: 6.6, w: 11.3, h: 0.45,
    fontSize: 12, color: "AACCEE", fontFace: "Calibri", align: "center", italic: true,
  });
  s.addImage({ path: logo, x: 5.9, y: 5.0, w: 1.6, h: 1.25, transparency: 15 });
}

// ─── Write ──────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: OUT }).then(() => {
  console.log(`✅  Written: ${OUT}`);
}).catch(err => {
  console.error("❌  Error:", err);
  process.exit(1);
});
