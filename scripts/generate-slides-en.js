const pptxgen = require("pptxgenjs");
const data = require("./weekly-data-2026-06-01");

const C = {
  navy: "1A2744",
  darkNavy: "0F1A2E",
  gold: "C9A96E",
  white: "FFFFFF",
  offWhite: "F5F6F8",
  lightGray: "E8EBF0",
  medGray: "8C95A6",
  bodyText: "2D3748",
  red: "C53030",
  green: "2F855A",
  paleGold: "FFF7E8",
};

const makeShadow = () => ({
  type: "outer",
  blur: 4,
  offset: 2,
  angle: 135,
  color: "000000",
  opacity: 0.12,
});

function addFooter(pres, slide, text) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 5.25,
    w: 10,
    h: 0.375,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText(text, {
    x: 0.5,
    y: 5.25,
    w: 9,
    h: 0.375,
    fontSize: 8,
    fontFace: "Arial",
    color: C.medGray,
    valign: "middle",
    margin: 0,
  });
}

function addHeader(pres, slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 10,
    h: 0.7,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText(title, {
    x: 0.6,
    y: 0,
    w: 9,
    h: 0.7,
    fontSize: 20,
    fontFace: "Arial",
    bold: true,
    color: C.white,
    valign: "middle",
    margin: 0,
  });
}

function pctColor(text) {
  return text.startsWith("-") ? C.red : C.green;
}

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.company = "42";
  pres.subject = "Crypto Weekly Digest";
  pres.title = data.en.deckTitle;
  pres.lang = "en-US";

  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 1.1, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
  s1.addText(data.en.deckTitle, {
    x: 0.8, y: 1.8, w: 8.2, h: 0.8,
    fontSize: 30, fontFace: "Arial", bold: true, color: C.white, margin: 0,
  });
  s1.addText(data.weekEn, {
    x: 0.8, y: 2.55, w: 8.2, h: 0.4,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0,
  });
  s1.addText("Sources: Messari · The Block · CoinGecko · CoinDesk · Crypto.news", {
    x: 0.8, y: 3.2, w: 8.5, h: 0.35,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0,
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold }, line: { color: C.gold } });
  s1.addText("CONFIDENTIAL — FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true,
  });
  s1.addNotes(data.notes[0]);

  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  addHeader(pres, s2, "📊  Market Overview");
  const cards = [
    { label: "BTC", value: data.marketCards.btc.price, change: data.marketCards.btc.change, accent: C.gold },
    { label: "ETH", value: data.marketCards.eth.price, change: data.marketCards.eth.change, accent: C.gold },
    { label: "Fear & Greed", value: data.marketCards.fng.value, change: data.marketCards.fng.label, accent: C.red },
  ];
  cards.forEach((card, i) => {
    const x = 0.6 + i * 3.1;
    s2.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.8, h: 1.05, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s2.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 0.06, h: 1.05, fill: { color: card.accent }, line: { color: card.accent } });
    s2.addText(card.label, { x: x + 0.24, y: 1.05, w: 2.2, h: 0.2, fontSize: 12, color: C.medGray, bold: true, margin: 0 });
    s2.addText(card.value, { x: x + 0.24, y: 1.32, w: 2.2, h: 0.28, fontSize: 25, color: C.bodyText, bold: true, margin: 0 });
    s2.addText(card.change, { x: x + 0.24, y: 1.67, w: 2.2, h: 0.2, fontSize: 13, color: pctColor(card.change), bold: true, margin: 0 });
  });
  const table = [[
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, align: "center" } },
    { text: "Price", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, align: "center" } },
    { text: "7d Change", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, align: "center" } },
    { text: "Notes", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, align: "center" } },
  ]];
  data.marketTableEn.forEach((row) => {
    table.push([
      { text: row.asset, options: { fontSize: 10, bold: true, align: "center", color: C.bodyText } },
      { text: row.price, options: { fontSize: 10, align: "right", color: C.bodyText } },
      { text: row.change, options: { fontSize: 10, align: "center", bold: true, color: pctColor(row.change) } },
      { text: row.note, options: { fontSize: 9, color: C.medGray } },
    ]);
  });
  s2.addTable(table, {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.2, 1.8, 1.4, 4.4],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.31, 0.31, 0.31, 0.31, 0.31, 0.31, 0.31],
    autoPage: false,
  });
  s2.addText(data.marketSummaryEn, {
    x: 0.6, y: 4.55, w: 8.8, h: 0.45,
    fontSize: 10, color: C.medGray, italic: true, margin: 0,
  });
  addFooter(pres, s2, `Crypto Market Weekly | ${data.weekShort}`);
  s2.addNotes(data.notes[1]);

  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  addHeader(pres, s3, "🔥  Top Stories & Industry");
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s3.addText("TOP 5 STORIES", { x: 0.6, y: 0.95, w: 4, h: 0.3, fontSize: 13, bold: true, color: C.navy, margin: 0 });
  data.topStoriesEn.forEach((story, i) => {
    const y = 1.36 + i * 0.72;
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y, w: 0.3, h: 0.3, fill: { color: C.gold }, line: { color: C.gold } });
    s3.addText(String(i + 1), { x: 0.65, y, w: 0.3, h: 0.3, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(story.title, { x: 1.1, y: y - 0.02, w: 3.5, h: 0.23, fontSize: 10.2, bold: true, color: C.bodyText, margin: 0 });
    s3.addText(story.desc, { x: 1.1, y: y + 0.22, w: 3.5, h: 0.23, fontSize: 8.5, color: C.medGray, margin: 0 });
  });
  [["Layer 1 & Layer 2", data.l1ItemsEn, 0.9], ["DeFi & Perp DEXes", data.defiItemsEn, 3.05]].forEach(([title, items, boxY]) => {
    s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: boxY, w: 4.5, h: 1.95, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: boxY, w: 4.5, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
    s3.addText(title, { x: 5.3, y: boxY, w: 4, h: 0.35, fontSize: 11, bold: true, color: C.white, valign: "middle", margin: 0 });
    s3.addText(items.map((text, idx) => ({
      text,
      options: { bullet: true, breakLine: idx < items.length - 1, fontSize: 9.3, color: C.bodyText },
    })), { x: 5.3, y: boxY + 0.42, w: 4.05, h: 1.4, margin: 0, paraSpaceAfter: 4 });
  });
  addFooter(pres, s3, `Crypto Market Weekly | ${data.weekShort}`);
  s3.addNotes(data.notes[2]);

  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  addHeader(pres, s4, "⚖️  Regulation & Fundraising");
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s4.addText("Regulation & Policy", { x: 0.7, y: 0.95, w: 4, h: 0.3, fontSize: 14, bold: true, color: C.navy, margin: 0 });
  data.regItemsEn.forEach((item, i) => {
    const y = 1.43 + i * 0.65;
    s4.addText(item.title, { x: 0.7, y, w: 4, h: 0.22, fontSize: 10.3, bold: true, color: C.bodyText, margin: 0 });
    s4.addText(item.desc, { x: 0.7, y: y + 0.23, w: 4, h: 0.24, fontSize: 8.9, color: C.medGray, margin: 0 });
  });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s4.addText("Fundraising & M&A", { x: 5.4, y: 0.95, w: 4, h: 0.3, fontSize: 14, bold: true, color: C.navy, margin: 0 });
  const fundTable = [[
    { text: "Project", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5 } },
    { text: "Amount", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5, align: "right" } },
    { text: "Notes", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5 } },
  ]];
  data.fundRowsEn.forEach((row) => {
    fundTable.push([
      { text: row.project, options: { fontSize: 9.0, bold: true, color: C.bodyText } },
      { text: row.amount, options: { fontSize: 9.0, align: "right", color: C.bodyText } },
      { text: row.note, options: { fontSize: 8.3, color: C.medGray } },
    ]);
  });
  s4.addTable(fundTable, {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.95, 1.85],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.29, 0.29, 0.29, 0.29, 0.29],
    autoPage: false,
  });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.45, w: 4.1, h: 0.78, fill: { color: C.paleGold }, line: { color: C.paleGold }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.45, w: 0.06, h: 0.78, fill: { color: C.gold }, line: { color: C.gold } });
  s4.addText(data.fundCalloutEn, { x: 5.55, y: 3.55, w: 3.6, h: 0.5, fontSize: 9.2, color: C.bodyText, margin: 0 });
  addFooter(pres, s4, `Crypto Market Weekly | ${data.weekShort}`);
  s4.addNotes(data.notes[3]);

  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  addHeader(pres, s5, "🔐  Security & Sector Highlights");
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.3, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red }, line: { color: C.red } });
  s5.addText(`Security Incidents | Weekly Loss ${data.securityLoss}`, { x: 0.6, y: 0.9, w: 4.1, h: 0.35, fontSize: 11, bold: true, color: C.white, valign: "middle", margin: 0 });
  const secTable = [[
    { text: "Incident", options: { bold: true, fontSize: 9, fill: { color: C.lightGray }, color: C.bodyText } },
    { text: "Loss", options: { bold: true, fontSize: 9, fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
    { text: "Vector", options: { bold: true, fontSize: 9, fill: { color: C.lightGray }, color: C.bodyText } },
  ]];
  data.securityRowsEn.forEach((row) => {
    secTable.push([
      { text: row.name, options: { fontSize: 8.7 } },
      { text: row.loss, options: { fontSize: 8.7, align: "right", color: C.red } },
      { text: row.vector, options: { fontSize: 8.2, color: C.medGray } },
    ]);
  });
  s5.addTable(secTable, {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.35, 1.1, 1.75],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });
  [["Stablecoins", data.stablecoinItemsEn, 3.45], ["AI x Crypto", data.aiItemsEn, 0.9], ["Prediction Markets", data.predictionItemsEn, 2.95]].forEach(([title, items, boxY]) => {
    const x = title === "Stablecoins" ? 0.4 : 5.1;
    const w = 4.5;
    const h = title === "Stablecoins" ? 1.55 : 1.85;
    s5.addShape(pres.shapes.RECTANGLE, { x, y: boxY, w, h, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s5.addShape(pres.shapes.RECTANGLE, { x, y: boxY, w, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
    s5.addText(title, { x: x + 0.2, y: boxY, w: 4, h: 0.35, fontSize: 11, bold: true, color: C.white, valign: "middle", margin: 0 });
    s5.addText(items.map((text, i) => ({
      text,
      options: { bullet: true, breakLine: i < items.length - 1, fontSize: 9.0, color: C.bodyText },
    })), { x: x + 0.2, y: boxY + 0.42, w: 4.05, h: h - 0.55, margin: 0, paraSpaceAfter: 4 });
  });
  addFooter(pres, s5, `Crypto Market Weekly | ${data.weekShort}`);
  s5.addNotes(data.notes[4]);

  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  addHeader(pres, s6, "🔮  Catalysts & Outlook");
  const catTable = [[
    { text: "Date", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5, align: "center" } },
    { text: "Event", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5 } },
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5, align: "center" } },
    { text: "Expected Impact", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9.5 } },
  ]];
  data.catalystsEn.forEach((row) => {
    const fill = row.highlight ? { color: C.paleGold } : undefined;
    catTable.push([
      { text: row.date, options: { fontSize: 9.0, align: "center", color: C.bodyText, fill } },
      { text: row.event, options: { fontSize: 9.0, color: C.bodyText, bold: row.highlight, fill } },
      { text: row.asset, options: { fontSize: 9.0, align: "center", color: C.navy, bold: true, fill } },
      { text: row.impact, options: { fontSize: 8.3, color: C.medGray, fill } },
    ]);
  });
  s6.addTable(catTable, {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.1, 1.1, 3.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.55, w: 9, h: 1.35, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.55, w: 0.06, h: 1.35, fill: { color: C.gold }, line: { color: C.gold } });
  s6.addText("Weekly Summary", { x: 0.8, y: 3.68, w: 8.2, h: 0.25, fontSize: 13, bold: true, color: C.navy, margin: 0 });
  s6.addText(data.outlookSummaryEn, { x: 0.8, y: 4.0, w: 8.3, h: 0.7, fontSize: 9.6, color: C.bodyText, margin: 0 });
  addFooter(pres, s6, `Crypto Market Weekly | ${data.weekShort} | CONFIDENTIAL`);
  s6.addNotes(data.notes[5]);

  await pres.writeFile({ fileName: data.outEn });
  console.log(`Written to ${data.outEn}`);
}

generate().catch((err) => {
  console.error(err);
  process.exit(1);
});
