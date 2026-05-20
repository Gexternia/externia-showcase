// Externia · Showcase Deck (PPTX)
// 11 slides, 4:3 (10 x 7.5"), dark Externia aesthetic, clickable links.

const PptxGenJS = require("pptxgenjs");
const path = require("path");

const ASSETS = path.resolve(__dirname, "..", "assets");
const A = (name) => path.join(ASSETS, name);

// --- Externia brand palette ---
const C = {
  bg:        "0A0A0B",
  void2:     "131314",
  surface:   "1C1B1C",
  surfaceHi: "201F20",
  text:      "E5E2E3",
  dim:       "A09D9E",
  faint:     "6B6869",
  outline:   "2A292A",
  outline2:  "3A3839",
  magenta:   "FF3A8A",
  orange:    "FF571A",
  orangeLt:  "FF8A3D",
  yellow:    "F1C100",
  rose:      "C0288F",
  green:     "6FE3A8",
  greenBg:   "1A3A28",
  greenBd:   "2B6A48",
};

const F = {
  head: "Calibri",
  body: "Calibri Light",
  mono: "Consolas",
};

const pptx = new PptxGenJS();
// 4:3 = 10 x 7.5 inches — more document-like, less "TV"
pptx.defineLayout({ name: "DECK_43", width: 10, height: 7.5 });
pptx.layout = "DECK_43";
pptx.title = "Externia · Showcase 2026";
pptx.author = "Externia";
pptx.company = "Externia";
pptx.subject = "Showcase comercial";

const W = 10;
const H = 7.5;
const TOTAL_SLIDES = 11;

// --- Helpers --------------------------------------------------
function darkBg(slide) {
  slide.background = { color: C.bg };
}
function eyebrow(slide, text, opts = {}) {
  const { x = 0.5, y = 0.4, color = C.magenta } = opts;
  slide.addText("▲ " + text, {
    x, y, w: 8, h: 0.3,
    fontFace: F.mono, fontSize: 10, bold: true,
    color, charSpacing: 4,
  });
}
function pageNum(slide, n) {
  slide.addText(`${String(n).padStart(2, "0")} / ${String(TOTAL_SLIDES).padStart(2, "0")}`, {
    x: W - 1.3, y: H - 0.45, w: 1, h: 0.25,
    fontFace: F.mono, fontSize: 8, color: C.faint, align: "right", charSpacing: 3,
  });
}
function brand(slide) {
  slide.addText("EXTERNIA", {
    x: W - 1.6, y: 0.4, w: 1.2, h: 0.25,
    fontFace: F.mono, fontSize: 9, bold: true, color: C.dim, charSpacing: 6, align: "right",
  });
}
function softCard(slide, x, y, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: 0.12,
    fill: { color: C.void2 },
    line: { color: C.outline, width: 1 },
    ...opts,
  });
}
// Adds a tiny "↗ link" line that's clickable
function linkLine(slide, x, y, w, label, url, opts = {}) {
  const { color = C.magenta, fontSize = 9, align = "left" } = opts;
  slide.addText("↗  " + label, {
    x, y, w, h: 0.28,
    fontFace: F.mono, fontSize, color,
    bold: true, charSpacing: 1, align, valign: "middle",
    hyperlink: { url, tooltip: label },
  });
}

// ============================================================
// SLIDE 1 — COVER
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  // Ambient circles
  s.addShape("ellipse", { x: -2, y: -1.5, w: 4.5, h: 4.5, fill: { color: C.magenta, transparency: 88 }, line: { type: "none" } });
  s.addShape("ellipse", { x: W - 2.5, y: H - 2.5, w: 4, h: 4, fill: { color: C.yellow, transparency: 92 }, line: { type: "none" } });

  s.addImage({ path: A("externia-icon-512.png"), x: W/2 - 0.85, y: H/2 - 2.0, w: 1.7, h: 1.7 });

  s.addText("externia", {
    x: 0, y: H/2 + 0.0, w: W, h: 1.3,
    fontFace: F.head, fontSize: 84, bold: true, color: C.magenta,
    align: "center", charSpacing: -2,
  });

  s.addText("Showcase · 2026", {
    x: 0, y: H/2 + 1.3, w: W, h: 0.4,
    fontFace: F.mono, fontSize: 12, color: C.dim, align: "center", charSpacing: 8,
  });

  s.addText("Producto  ·  Activaciones  ·  Formaciones  ·  I+D", {
    x: 0.5, y: H - 0.7, w: 6, h: 0.3,
    fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 4,
  });

  // Email clickable
  s.addText("g.prado@externia.ai", {
    x: W - 3.5, y: H - 0.7, w: 3, h: 0.3,
    fontFace: F.mono, fontSize: 9, color: C.faint, align: "right", charSpacing: 3,
    hyperlink: { url: "mailto:g.prado@externia.ai", tooltip: "Escríbenos" },
  });
}

// ============================================================
// SLIDE 2 — HOOK / STATS
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "No es un deck. Es lo que ya opera");

  s.addText([
    { text: "Lo que estamos construyendo\n", options: { color: C.text } },
    { text: "en ", options: { color: C.text } },
    { text: "Externia", options: { color: C.magenta, bold: true } },
    { text: ", ahora mismo", options: { color: C.text } },
  ], {
    x: 0.5, y: 1.1, w: W - 1.0, h: 2.0,
    fontFace: F.head, fontSize: 44, bold: true,
    align: "center", charSpacing: -1, paraSpaceAfter: 0,
  });

  s.addText(
    "Producto en producción, activaciones para marcas, formaciones de IA y la investigación que define el próximo salto. Todo está vivo. Todo se puede tocar.",
    {
      x: 1.5, y: 3.4, w: W - 3, h: 1.0,
      fontFace: F.body, fontSize: 13, color: C.dim, align: "center",
    }
  );

  // Stats strip
  const stats = [
    { k: "8+",   l: "demos públicas" },
    { k: "10",   l: "meses operando" },
    { k: "100%", l: "hecho por Externia" },
    { k: "∞",    l: "blog autónomo" },
  ];
  const colW = 1.9;
  const rowX = (W - colW * 4) / 2;
  const rowY = 5.0;
  s.addShape("line", { x: rowX, y: rowY, w: colW*4, h: 0, line: { color: C.outline2, width: 1 } });
  s.addShape("line", { x: rowX, y: rowY + 1.4, w: colW*4, h: 0, line: { color: C.outline2, width: 1 } });
  stats.forEach((st, i) => {
    const x = rowX + i * colW;
    if (i > 0) s.addShape("line", { x, y: rowY + 0.2, w: 0, h: 1, line: { color: C.outline2, width: 1 } });
    s.addText(st.k, {
      x, y: rowY + 0.18, w: colW, h: 0.7,
      fontFace: F.head, fontSize: 36, bold: true, color: C.magenta, align: "center",
    });
    s.addText(st.l, {
      x, y: rowY + 0.9, w: colW, h: 0.35,
      fontFace: F.mono, fontSize: 9, color: C.faint, align: "center", charSpacing: 3,
    });
  });

  pageNum(s, 2);
}

// ============================================================
// SLIDE 3 — EVENTOPLUS
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "A · Cliente recurrente en producción");

  // Logos row
  s.addImage({ path: A("eventoplus-logo.webp"), x: 0.5, y: 0.85, w: 0.6, h: 0.6 });
  s.addText("×", { x: 1.18, y: 0.85, w: 0.3, h: 0.6, fontFace: F.body, fontSize: 22, color: C.faint, align: "center", valign: "middle" });
  s.addImage({ path: A("finderia-logo.webp"), x: 1.55, y: 0.97, w: 1.6, h: 0.32 });

  // Live chip
  s.addShape("roundRect", {
    x: 3.4, y: 1.0, w: 2.8, h: 0.3, rectRadius: 0.15,
    fill: { color: C.greenBg }, line: { color: C.greenBd, width: 1 },
  });
  s.addText("● Live · Cliente activo desde feb 2026", {
    x: 3.4, y: 1.0, w: 2.8, h: 0.3,
    fontFace: F.mono, fontSize: 9, color: C.green, align: "center", valign: "middle", charSpacing: 1, bold: true,
  });

  // Title
  s.addText([
    { text: "El stack conversacional\nde ", options: { color: C.text } },
    { text: "Eventoplus", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.5, y: 1.6, w: W - 1.0, h: 1.4,
    fontFace: F.head, fontSize: 32, bold: true, charSpacing: -1,
  });

  // Description
  s.addText(
    "Chatbot de búsqueda + dashboard analítico + informe ejecutivo semanal. Tres piezas que conviven en un loop de producto, datos y consultoría — no es software entregado, es un servicio recurrente que evoluciona cada semana.",
    {
      x: 0.5, y: 3.05, w: W - 1.0, h: 0.85,
      fontFace: F.body, fontSize: 12, color: C.dim,
    }
  );

  // 3 KPI cards
  const kpis = [
    { k: "32",   l: "sesiones · sem 19" },
    { k: "116",  l: "mensajes procesados" },
    { k: "+14%", l: "vs sem. anterior" },
  ];
  const kY = 4.0;
  const kpW = 2.95;
  const kpStart = 0.5;
  kpis.forEach((kp, i) => {
    const x = kpStart + i * (kpW + 0.05);
    softCard(s, x, kY, kpW, 1.0);
    s.addText(kp.k, {
      x: x + 0.2, y: kY + 0.15, w: kpW - 0.4, h: 0.45,
      fontFace: F.head, fontSize: 28, bold: true, color: C.magenta,
    });
    s.addText(kp.l, {
      x: x + 0.2, y: kY + 0.65, w: kpW - 0.4, h: 0.3,
      fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 3,
    });
  });

  // Service buttons (clickable)
  const services = [
    { code: "01", name: "Chatbot FinderAI",          url: "https://naranja-bot.onrender.com",          host: "naranja-bot.onrender.com" },
    { code: "02", name: "Dashboard de estadísticas", url: "https://admin-frontend-b33o.onrender.com",  host: "admin-frontend-b33o.onrender.com" },
  ];
  services.forEach((srv, i) => {
    const y = 5.2 + i * 0.55;
    const sx = 0.5;
    const sw = W - 1.0;
    s.addShape("roundRect", {
      x: sx, y, w: sw, h: 0.45, rectRadius: 0.08,
      fill: { color: i === 0 ? C.magenta : C.surfaceHi },
      line: { color: i === 0 ? C.magenta : C.outline2, width: 1 },
      hyperlink: { url: srv.url, tooltip: srv.host },
    });
    s.addText(srv.code, {
      x: sx + 0.18, y, w: 0.4, h: 0.45,
      fontFace: F.mono, fontSize: 10, color: i === 0 ? "1A0008" : C.magenta, valign: "middle", bold: true, charSpacing: 2,
    });
    s.addText(srv.name, {
      x: sx + 0.7, y, w: 4.5, h: 0.45,
      fontFace: F.head, fontSize: 12, color: i === 0 ? "1A0008" : C.text, valign: "middle", bold: true,
      hyperlink: { url: srv.url, tooltip: srv.host },
    });
    s.addText("↗ " + srv.host, {
      x: sx + sw - 4.0, y, w: 3.8, h: 0.45,
      fontFace: F.mono, fontSize: 9, color: i === 0 ? "1A0008" : C.faint, valign: "middle", align: "right", charSpacing: 1,
      hyperlink: { url: srv.url, tooltip: srv.host },
    });
  });

  // Footer note
  s.addText("Demanda · Madrid 44%  ·  Andalucía 13%  ·  Barcelona 16%  ·  Valencia · Alicante 9%  ·  1 lead cerrado en Valencia", {
    x: 0.5, y: 6.45, w: W - 1.0, h: 0.3,
    fontFace: F.mono, fontSize: 9, color: C.faint, align: "center", charSpacing: 2,
  });

  pageNum(s, 3);
}

// ============================================================
// SLIDE 4 — INFORME EJECUTIVO
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "A · El entregable que cierra el valor", { color: C.yellow });

  s.addText("Informe Ejecutivo Semanal", {
    x: 0.5, y: 0.85, w: W - 1.0, h: 0.7,
    fontFace: F.head, fontSize: 36, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText("PDF · 4 páginas · entregado cada lunes a Eventoplus", {
    x: 0.5, y: 1.55, w: W - 1.0, h: 0.3,
    fontFace: F.mono, fontSize: 11, color: C.yellow, charSpacing: 3,
  });

  s.addText("Cada lunes, el equipo de Eventoplus recibe un análisis ejecutivo entregable con:", {
    x: 0.5, y: 2.2, w: W - 1.0, h: 0.4,
    fontFace: F.body, fontSize: 13, color: C.dim,
  });

  const bullets = [
    "Síntesis ejecutiva + KPIs vs semana anterior",
    "Demanda geográfica por regiones",
    "Tipología de consultas y hallazgos",
    "Conversiones y momentos destacados",
    "Incidencias técnicas y áreas de mejora",
    "Recomendaciones accionables próxima semana",
  ];
  bullets.forEach((b, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55;
    const y = 2.8 + row * 0.55;
    s.addText("▸", {
      x, y, w: 0.25, h: 0.45,
      fontFace: F.head, fontSize: 14, color: C.yellow, bold: true, valign: "middle",
    });
    s.addText(b, {
      x: x + 0.25, y, w: 4.2, h: 0.45,
      fontFace: F.body, fontSize: 12, color: C.text, valign: "middle",
    });
  });

  s.addShape("line", { x: 0.5, y: 4.85, w: W - 1.0, h: 0, line: { color: C.outline2, width: 1, dashType: "dash" } });
  s.addText([
    { text: "Es la diferencia entre vender un chatbot y vender ", options: { color: C.text, italic: true } },
    { text: "inteligencia operativa continua", options: { color: C.magenta, italic: false, bold: true } },
  ], {
    x: 0.5, y: 5.1, w: W - 1.0, h: 0.9,
    fontFace: F.head, fontSize: 18, italic: true,
  });

  // PDF link
  s.addShape("roundRect", {
    x: W/2 - 2.4, y: 6.3, w: 4.8, h: 0.55, rectRadius: 0.1,
    fill: { color: C.surfaceHi }, line: { color: C.outline2, width: 1 },
    hyperlink: { url: "https://externia-showcase.onrender.com/assets/informe-eventoplus-semana-19.pdf", tooltip: "Ver muestra del informe" },
  });
  s.addText("↗  Ver muestra · Informe Ejecutivo · semana 19 (PDF)", {
    x: W/2 - 2.4, y: 6.3, w: 4.8, h: 0.55,
    fontFace: F.head, fontSize: 12, color: C.text, bold: true, align: "center", valign: "middle",
    hyperlink: { url: "https://externia-showcase.onrender.com/assets/informe-eventoplus-semana-19.pdf", tooltip: "Ver muestra del informe" },
  });

  pageNum(s, 4);
}

// ============================================================
// SLIDE 5 — TIPOS DE TRABAJO  (NEW!)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "Tipos de trabajo que hacemos");

  s.addText([
    { text: "Cuatro formatos, ", options: { color: C.text } },
    { text: "una misma metodología", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.5, y: 0.85, w: W - 1.0, h: 0.8,
    fontFace: F.head, fontSize: 30, bold: true, charSpacing: -1,
  });
  s.addText("Lo que producimos para cada cliente, según lo que necesita.", {
    x: 0.5, y: 1.7, w: W - 1.0, h: 0.35,
    fontFace: F.body, fontSize: 12, color: C.dim,
  });

  // 2x2 grid of categories
  const types = [
    {
      tag: "01 · ACTIVACIÓN EXPERIENCIAL",
      title: "Activaciones de marca en evento",
      desc: "Producto generativo en vivo para asistentes de un evento corporativo. Se llevan algo personalizado, gancho real, retención alta. La marca queda asociada a la experiencia.",
      audience: "Para el asistente final",
      examples: "KPMG · Clinique · Talento Joven",
      color: C.magenta,
    },
    {
      tag: "02 · FORMACIÓN / TALLER IA",
      title: "Talleres prácticos de IA",
      desc: "Sesiones donde el equipo del cliente trabaja con herramientas reales: caricaturas, edición selectiva, generación visual. Aprenden usando, no escuchando. Salen con criterio.",
      audience: "Para el equipo del cliente",
      examples: "Taller IA · ING",
      color: C.orangeLt,
    },
    {
      tag: "03 · PRODUCTO RECURRENTE",
      title: "Servicio operando todos los días",
      desc: "Bot conversacional, dashboard analítico e informe consultivo semanal. No es entrega puntual: es un sistema que vive con el cliente y evoluciona con su producto cada semana.",
      audience: "Para la operación del cliente",
      examples: "Eventoplus · FinderAI",
      color: C.yellow,
    },
    {
      tag: "04 · CONSULTORÍA CON AGENTES",
      title: "Sistemas de agentes IA",
      desc: "Agentes entrenados con el histórico del cliente para acelerar concursos, propuestas, decks, investigación. La inteligencia interna del equipo se vuelve sistema replicable.",
      audience: "Para los equipos creativos y comerciales",
      examples: "Turespaña · DeckCraft · Externia interno",
      color: C.rose,
    },
  ];

  const cardW = 4.45;
  const cardH = 2.45;
  const gap = 0.1;
  const gridX = 0.5;
  const gridY = 2.2;

  types.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = gridX + col * (cardW + gap);
    const y = gridY + row * (cardH + gap);

    softCard(s, x, y, cardW, cardH, { line: { color: t.color, width: 1 } });

    // Top accent bar
    s.addShape("rect", { x: x + 0.15, y: y + 0.15, w: 0.4, h: 0.04, fill: { color: t.color }, line: { type: "none" } });
    // Tag
    s.addText(t.tag, {
      x: x + 0.15, y: y + 0.22, w: cardW - 0.3, h: 0.3,
      fontFace: F.mono, fontSize: 9, color: t.color, charSpacing: 3, bold: true,
    });
    // Title
    s.addText(t.title, {
      x: x + 0.15, y: y + 0.55, w: cardW - 0.3, h: 0.5,
      fontFace: F.head, fontSize: 17, bold: true, color: C.text, charSpacing: -0.5,
    });
    // Desc
    s.addText(t.desc, {
      x: x + 0.15, y: y + 1.1, w: cardW - 0.3, h: 0.9,
      fontFace: F.body, fontSize: 10.5, color: C.dim,
    });
    // Footer line
    s.addShape("line", {
      x: x + 0.15, y: y + 2.0, w: cardW - 0.3, h: 0,
      line: { color: C.outline, width: 1, dashType: "dash" },
    });
    // Audience + examples
    s.addText(t.audience, {
      x: x + 0.15, y: y + 2.05, w: cardW - 0.3, h: 0.18,
      fontFace: F.mono, fontSize: 8, color: C.faint, charSpacing: 2,
    });
    s.addText(t.examples, {
      x: x + 0.15, y: y + 2.22, w: cardW - 0.3, h: 0.2,
      fontFace: F.head, fontSize: 10, color: t.color, bold: true,
    });
  });

  pageNum(s, 5);
}

// ============================================================
// SLIDE 6 — ACTIVACIONES (grid 3+2)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "B · Activaciones públicas · accesibles ahora");

  s.addText([
    { text: "Producto creado para marcas finales, ", options: { color: C.text } },
    { text: "en vivo", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.5, y: 0.85, w: W - 1.0, h: 0.7,
    fontFace: F.head, fontSize: 26, bold: true, charSpacing: -1,
  });
  s.addText("Cada URL abre la demo real que se usó en el evento del cliente.", {
    x: 0.5, y: 1.55, w: W - 1.0, h: 0.3,
    fontFace: F.body, fontSize: 11, color: C.dim,
  });

  const tiles = [
    { client: "KPMG",               title: "ComicGen",                  blurb: "Cómic generado con IA en directo. Asistentes son protagonistas.", url: "https://comic-ai-s1ca.onrender.com",   host: "comic-ai-s1ca.onrender.com",   image: A("comicgen-kpmg.png") },
    { client: "Clinique · INHOUSE", title: "Cartas personalizadas",     blurb: "Cartas de marca con IA, narrativa coherente, listas para imprimir.", url: "https://clinique-bium.onrender.com",  host: "clinique-bium.onrender.com",   image: A("clinique-letter.png") },
    { client: "Taller IA · ING",    title: "SmartBrush + caricaturas",  blurb: "Caricatura en vivo, edición selectiva por pincel, retratos generativos.", url: "https://taller-ia-ing.onrender.com",  host: "taller-ia-ing.onrender.com",   image: A("smartbrush-caricature.png") },
    { client: "Talento Joven",      title: "Retro Game",                blurb: "Videojuego retro generado para evento. Sin instalación, jugable en móvil.", url: "https://videogame-witg.onrender.com", host: "videogame-witg.onrender.com", image: null },
    { client: "Externia · Interno", title: "DeckCraft",                 blurb: "Generador de presentaciones desde brief + sistema de diseño en MD.", url: "https://deckcraft-3u04.onrender.com", host: "deckcraft-3u04.onrender.com", image: A("deckcraft-sample.png") },
  ];

  const tileW = 2.95;
  const tileH = 2.35;
  const gap = 0.1;
  const startX = 0.55;
  const topY = 1.95;
  const bottomY = topY + tileH + gap;
  const bottomTileW = 4.5;
  const bottomStartX = (W - (bottomTileW * 2 + gap)) / 2;

  tiles.forEach((t, i) => {
    const isTop = i < 3;
    const x = isTop ? startX + i * (tileW + gap) : bottomStartX + (i - 3) * (bottomTileW + gap);
    const y = isTop ? topY : bottomY;
    const w = isTop ? tileW : bottomTileW;

    softCard(s, x, y, w, tileH, {
      hyperlink: { url: t.url, tooltip: t.host },
    });

    const thumbH = tileH * 0.50;
    s.addShape("roundRect", {
      x: x + 0.08, y: y + 0.08, w: w - 0.16, h: thumbH - 0.08,
      rectRadius: 0.08,
      fill: { color: "0f0e10" }, line: { color: C.outline, width: 1 },
    });
    if (t.image) {
      s.addImage({ path: t.image, x: x + 0.1, y: y + 0.1, w: w - 0.2, h: thumbH - 0.12, sizing: { type: "cover", w: w - 0.2, h: thumbH - 0.12 }, hyperlink: { url: t.url, tooltip: t.host } });
    } else {
      // Retro Game pixel placeholder
      const pxColors = [C.yellow, C.magenta, C.yellow, C.orange, C.magenta, C.yellow];
      const startPxX = x + w/2 - (pxColors.length * 0.22 + (pxColors.length-1) * 0.1) / 2;
      pxColors.forEach((col, pi) => {
        s.addShape("rect", {
          x: startPxX + pi * 0.32, y: y + thumbH/2 - 0.11, w: 0.22, h: 0.22,
          fill: { color: col }, line: { type: "none" },
        });
      });
    }
    s.addText("● " + t.client, {
      x: x + 0.18, y: y + 0.15, w: w - 0.4, h: 0.22,
      fontFace: F.mono, fontSize: 7, color: C.text, charSpacing: 2, bold: true,
    });
    s.addText(t.title, {
      x: x + 0.18, y: y + thumbH + 0.02, w: w - 0.36, h: 0.32,
      fontFace: F.head, fontSize: 14, bold: true, color: C.text, charSpacing: -0.5,
      hyperlink: { url: t.url, tooltip: t.host },
    });
    s.addText(t.blurb, {
      x: x + 0.18, y: y + thumbH + 0.35, w: w - 0.36, h: 0.55,
      fontFace: F.body, fontSize: 9, color: C.dim,
    });
    s.addText("↗  " + t.host, {
      x: x + 0.18, y: y + tileH - 0.32, w: w - 0.36, h: 0.25,
      fontFace: F.mono, fontSize: 8, color: C.magenta, charSpacing: 1, bold: true,
      hyperlink: { url: t.url, tooltip: t.host },
    });
  });

  pageNum(s, 6);
}

// ============================================================
// SLIDE 7 — TURESPAÑA: Reto + Externa vs Externia
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "C · Caso público tier-1");

  s.addText([
    { text: "V Convención de Turespaña 2025\n", options: { color: C.text } },
    { text: "Cómo ", options: { color: C.text } },
    { text: "Externa", options: { color: C.magenta, bold: true } },
    { text: " la ganó", options: { color: C.text } },
  ], {
    x: 0.5, y: 0.85, w: W - 1.0, h: 1.5,
    fontFace: F.head, fontSize: 30, bold: true, charSpacing: -1, paraSpaceAfter: 0,
  });

  // Reto card
  softCard(s, 0.5, 2.5, W - 1.0, 1.15);
  s.addText("▼ EL RETO", {
    x: 0.7, y: 2.62, w: 4, h: 0.28,
    fontFace: F.mono, fontSize: 9, color: C.yellow, bold: true, charSpacing: 3,
  });
  s.addText(
    "Concurso público tier-1, criterios técnicos estrictos, competencia de agencias top. Memoria densa, escenografía ambiciosa, narrativa coherente con el «por qué» de Turespaña.",
    {
      x: 0.7, y: 2.9, w: W - 1.4, h: 0.7,
      fontFace: F.body, fontSize: 11.5, color: C.text,
    }
  );

  // Two columns
  const colW = (W - 1.1) / 2;
  const colY = 3.8;
  const colH = 3.0;

  // Externa col
  softCard(s, 0.5, colY, colW, colH);
  s.addText("Externa", {
    x: 0.7, y: colY + 0.18, w: colW - 0.4, h: 0.35,
    fontFace: F.head, fontSize: 17, bold: true, color: C.text,
  });
  s.addText("Lo que aportó", {
    x: 0.7, y: colY + 0.5, w: colW - 0.4, h: 0.25,
    fontFace: F.mono, fontSize: 8, color: C.faint, charSpacing: 3,
  });
  const externaItems = [
    "Cartera tier-1 — Amazon, ING, Caixabank, KPMG, Cámara de Comercio.",
    "Criterio del equipo sobre qué funciona en cada bloque.",
    "Relación de confianza con Turespaña y dominio del cliente institucional.",
  ];
  externaItems.forEach((it, i) => {
    const y = colY + 0.9 + i * 0.62;
    s.addText("▸", { x: 0.7, y, w: 0.2, h: 0.5, fontFace: F.head, fontSize: 11, color: C.faint, bold: true, valign: "top" });
    s.addText(it, { x: 0.9, y, w: colW - 0.55, h: 0.6, fontFace: F.body, fontSize: 10.5, color: C.text });
  });

  // Externia col
  const x2 = 0.5 + colW + 0.1;
  softCard(s, x2, colY, colW, colH, { line: { color: C.magenta, width: 1 } });
  s.addText("Externia", {
    x: x2 + 0.2, y: colY + 0.18, w: colW - 0.4, h: 0.35,
    fontFace: F.head, fontSize: 17, bold: true, color: C.magenta,
  });
  s.addText("Lo que aportó", {
    x: x2 + 0.2, y: colY + 0.5, w: colW - 0.4, h: 0.25,
    fontFace: F.mono, fontSize: 8, color: C.faint, charSpacing: 3,
  });
  const externiaItems = [
    "Agente entrenado con histórico Turespaña + feedback público.",
    "Inteligencia interna de Externa hecha sistema replicable.",
    "Investigación profunda automática del «porqué» del cliente.",
    "Renders de escenografía con IA — propuesta visualmente rica.",
  ];
  externiaItems.forEach((it, i) => {
    const y = colY + 0.9 + i * 0.47;
    s.addText("▸", { x: x2 + 0.2, y, w: 0.2, h: 0.45, fontFace: F.head, fontSize: 11, color: C.magenta, bold: true, valign: "top" });
    s.addText(it, { x: x2 + 0.4, y, w: colW - 0.55, h: 0.45, fontFace: F.body, fontSize: 10.5, color: C.text });
  });

  pageNum(s, 7);
}

// ============================================================
// SLIDE 8 — TURESPAÑA: Quote
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  s.addShape("ellipse", {
    x: W/2 - 4, y: -1.5, w: 8, h: 8,
    fill: { color: C.magenta, transparency: 92 }, line: { type: "none" },
  });

  brand(s);
  eyebrow(s, "C · Cita del caso");

  s.addText("“", {
    x: 0.5, y: 1.2, w: 1.8, h: 1.8,
    fontFace: F.head, fontSize: 160, color: C.magenta, bold: true,
  });

  s.addText([
    { text: "La IA no ganó la convención.\n", options: { color: C.text } },
    { text: "La ganó ", options: { color: C.text } },
    { text: "Externa", options: { color: C.magenta, bold: true } },
  ], {
    x: 1.2, y: 2.4, w: W - 2.4, h: 2.3,
    fontFace: F.head, fontSize: 42, bold: true, align: "center", charSpacing: -1.5,
  });

  s.addText(
    "Pero la IA hizo posible una propuesta más rica de la que habríamos podido permitirnos meter en cada concurso.",
    {
      x: 1.5, y: 4.95, w: W - 3, h: 1.0,
      fontFace: F.body, fontSize: 15, color: C.dim, align: "center", italic: true,
    }
  );

  s.addShape("roundRect", {
    x: W/2 - 2.3, y: 6.3, w: 4.6, h: 0.55, rectRadius: 0.28,
    fill: { color: C.greenBg }, line: { color: C.greenBd, width: 1 },
  });
  s.addText("✓  Externa ganó la V Convención", {
    x: W/2 - 2.3, y: 6.3, w: 4.6, h: 0.55,
    fontFace: F.head, fontSize: 14, color: C.green, bold: true, align: "center", valign: "middle",
  });

  pageNum(s, 8);
}

// ============================================================
// SLIDE 9 — METODOLOGÍA
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "D · Cómo trabajamos");

  s.addText("Cómo entramos en una agencia", {
    x: 0.5, y: 0.85, w: W - 1.0, h: 0.7,
    fontFace: F.head, fontSize: 32, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText(
    "Lo primero no es construir. Es observar, mapear y decidir qué construir primero.",
    {
      x: 0.5, y: 1.6, w: W - 1.0, h: 0.4,
      fontFace: F.body, fontSize: 12, color: C.dim,
    }
  );

  // Timeline
  const tlY = 2.55;
  const tlPad = 0.5;
  const tlW = W - tlPad * 2;
  s.addShape("rect", { x: tlPad, y: tlY - 0.025, w: tlW, h: 0.05, fill: { color: C.magenta, transparency: 30 }, line: { type: "none" } });

  const phases = [
    { n: "01", when: "Sem 1-3", name: "Inmersión", hot: true, desc: "Dentro del equipo. Mapa de tareas que consumen tiempo, quick wins identificados." },
    { n: "02", when: "Mes 1-2", name: "Capa rápida", desc: "Ganancias visibles en semanas. Agente redactor, transcripción, generador de decks." },
    { n: "03", when: "Mes 2-4", name: "Capa profunda", desc: "Agentes conectados a datos. Histórico de propuestas consultable, pliegos estructurados." },
    { n: "04", when: "Mes 4-9", name: "Escala", desc: "Sistema que mejora con cada concurso. Métricas de uso, roadmap consolidado." },
  ];
  const phW = tlW / 4;
  phases.forEach((p, i) => {
    const x = tlPad + i * phW;
    s.addShape("ellipse", {
      x: x + phW/2 - 0.27, y: tlY - 0.27, w: 0.54, h: 0.54,
      fill: { color: p.hot ? C.magenta : C.surfaceHi },
      line: { color: p.hot ? C.magenta : C.outline2, width: 1.5 },
    });
    s.addText(p.n, {
      x: x + phW/2 - 0.27, y: tlY - 0.27, w: 0.54, h: 0.54,
      fontFace: F.mono, fontSize: 10, color: p.hot ? "1A0008" : C.dim, bold: true, align: "center", valign: "middle",
    });
    s.addText(p.when, {
      x: x + 0.1, y: tlY + 0.45, w: phW - 0.2, h: 0.28,
      fontFace: F.mono, fontSize: 9, color: C.faint, align: "center", charSpacing: 3,
    });
    s.addText(p.name, {
      x: x + 0.1, y: tlY + 0.78, w: phW - 0.2, h: 0.45,
      fontFace: F.head, fontSize: 16, bold: true,
      color: p.hot ? C.magenta : C.text, align: "center", charSpacing: -0.5,
    });
    s.addText(p.desc, {
      x: x + 0.1, y: tlY + 1.32, w: phW - 0.2, h: 1.5,
      fontFace: F.body, fontSize: 9.5, color: C.dim, align: "center",
    });
  });

  // Felt block
  const fY = 5.75;
  s.addShape("line", { x: 0.5, y: fY - 0.12, w: W - 1.0, h: 0, line: { color: C.outline, width: 1, dashType: "dash" } });
  s.addText("▼ DESDE EL PRIMER MES", {
    x: 0.5, y: fY, w: 4, h: 0.28,
    fontFace: F.mono, fontSize: 9, color: C.yellow, bold: true, charSpacing: 3,
  });
  s.addText([
    { text: "No sienten que les han metido un proyecto encima. Sienten que tienen ", options: { color: C.text } },
    { text: "herramientas nuevas", options: { color: C.magenta, bold: true } },
    { text: " que están haciendo el trabajo aburrido por ellos.", options: { color: C.text } },
  ], {
    x: 0.5, y: fY + 0.3, w: W - 1.0, h: 0.9,
    fontFace: F.head, fontSize: 15, italic: true,
  });

  pageNum(s, 9);
}

// ============================================================
// SLIDE 10 — I+D ABIERTO
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "E · I+D abierto · Q2 2026", { color: C.yellow });

  s.addText("Plataforma 3D · World API", {
    x: 0.5, y: 0.85, w: W - 1.0, h: 0.7,
    fontFace: F.head, fontSize: 32, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText([
    { text: "Investigamos lo que va a ", options: { color: C.dim } },
    { text: "redefinir el sector", options: { color: C.magenta, bold: true } },
    { text: " en los próximos 18 meses.", options: { color: C.dim } },
  ], {
    x: 0.5, y: 1.55, w: W - 1.0, h: 0.4,
    fontFace: F.body, fontSize: 13,
  });

  // 4 KPIs row
  const kY = 2.25;
  const kpis = [
    { v: "Validada",     l: "Tecnología" },
    { v: "En formación", l: "Mercado" },
    { v: "12-18 m",      l: "Ventana competitiva" },
    { v: "< $100",       l: "Coste de validación" },
  ];
  const kpiW = (W - 1.15) / 4;
  kpis.forEach((kp, i) => {
    const x = 0.5 + i * (kpiW + 0.05);
    softCard(s, x, kY, kpiW, 0.95);
    s.addText(kp.v, {
      x: x + 0.15, y: kY + 0.15, w: kpiW - 0.3, h: 0.35,
      fontFace: F.head, fontSize: 16, color: C.magenta, bold: true,
    });
    s.addText(kp.l, {
      x: x + 0.15, y: kY + 0.55, w: kpiW - 0.3, h: 0.3,
      fontFace: F.mono, fontSize: 8, color: C.faint, charSpacing: 3,
    });
  });

  // Backers strip
  const bY = 3.45;
  s.addShape("line", { x: 0.5, y: bY, w: W - 1.0, h: 0, line: { color: C.outline, width: 1 } });
  s.addText("BACKERS", {
    x: 0.5, y: bY + 0.15, w: 1.2, h: 0.3,
    fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 4,
  });
  s.addText("NVIDIA  ·  AMD  ·  Autodesk  ·  $1.290M  ·  World API pública desde enero 2026", {
    x: 1.7, y: bY + 0.15, w: W - 2.2, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.text, charSpacing: 2, bold: true,
  });

  // 3 puntos
  const pY = 4.15;
  const points = [
    { n: "01", t: "World API",          d: "El motor de World Labs. Una sola foto → mundo 3D navegable persistente." },
    { n: "02", t: "Gaussian Splatting", d: "Renderizado que alcanzó madurez en 2025. La curva de coste ya cruzó." },
    { n: "03", t: "First mover MICE",   d: "Nadie ha verticalizado esta tecnología para eventos. Ventana real." },
  ];
  const ptW = (W - 1.2) / 3;
  points.forEach((p, i) => {
    const x = 0.5 + i * (ptW + 0.1);
    softCard(s, x, pY, ptW, 1.5);
    s.addText(p.n + " · " + p.t.toUpperCase(), {
      x: x + 0.15, y: pY + 0.18, w: ptW - 0.3, h: 0.3,
      fontFace: F.mono, fontSize: 9, color: C.magenta, charSpacing: 2, bold: true,
    });
    s.addText(p.d, {
      x: x + 0.15, y: pY + 0.55, w: ptW - 0.3, h: 0.85,
      fontFace: F.body, fontSize: 10, color: C.text,
    });
  });

  // Punch
  s.addText([
    { text: "Cliente que entre ahora tiene ", options: { color: C.text } },
    { text: "18 meses de ventaja competitiva real", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.5, y: 6.0, w: W - 1.0, h: 0.7,
    fontFace: F.head, fontSize: 16, italic: true, align: "center",
  });

  pageNum(s, 10);
}

// ============================================================
// SLIDE 11 — CIERRE
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  s.addShape("ellipse", { x: -2.5, y: H - 3, w: 6, h: 6, fill: { color: C.magenta, transparency: 92 }, line: { type: "none" } });
  s.addShape("ellipse", { x: W - 3, y: -2.5, w: 5.5, h: 5.5, fill: { color: C.yellow, transparency: 94 }, line: { type: "none" } });

  brand(s);
  eyebrow(s, "F · Pipeline editorial autónomo", { color: C.yellow });

  // Blog proof card — clickable
  softCard(s, 0.5, 0.9, W - 1.0, 1.7, {
    hyperlink: { url: "https://externia.ai/blog", tooltip: "Abrir el blog autónomo" },
  });
  s.addText("externia.ai/blog", {
    x: 0.7, y: 1.05, w: W - 1.4, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.faint, charSpacing: 4,
    hyperlink: { url: "https://externia.ai/blog" },
  });
  s.addText([
    { text: "El blog de Externia se ", options: { color: C.text } },
    { text: "escribe solo", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.7, y: 1.4, w: W - 1.4, h: 0.55,
    fontFace: F.head, fontSize: 22, bold: true, charSpacing: -1,
    hyperlink: { url: "https://externia.ai/blog" },
  });
  s.addText(
    "Posts publicados sin intervención humana en la cadena editorial. Agentes de Externia operando en producción desde hace meses.",
    {
      x: 0.7, y: 2.0, w: W - 1.4, h: 0.5,
      fontFace: F.body, fontSize: 11, color: C.dim,
    }
  );

  // Closing pitch
  s.addText([
    { text: "Esto opera hoy\n", options: { color: C.text } },
    { text: "Hablemos de ", options: { color: C.text } },
    { text: "lo siguiente", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.5, y: 3.1, w: W - 1.0, h: 1.7,
    fontFace: F.head, fontSize: 46, bold: true, align: "center", charSpacing: -1.5,
  });

  // CTA buttons (clickable)
  const ctaY = 5.4;
  // Primary
  s.addShape("roundRect", {
    x: 0.6, y: ctaY, w: 4.3, h: 0.9, rectRadius: 0.1,
    fill: { color: C.magenta }, line: { type: "none" },
    hyperlink: { url: "mailto:g.prado@externia.ai", tooltip: "Escríbenos" },
  });
  s.addText("✉  g.prado@externia.ai", {
    x: 0.6, y: ctaY + 0.05, w: 4.3, h: 0.5,
    fontFace: F.head, fontSize: 15, color: "1A0008", bold: true, align: "center", valign: "middle",
    hyperlink: { url: "mailto:g.prado@externia.ai", tooltip: "Escríbenos" },
  });
  s.addText("Una conversación, no un formulario", {
    x: 0.6, y: ctaY + 0.5, w: 4.3, h: 0.35,
    fontFace: F.mono, fontSize: 9, color: "5C0018", align: "center", valign: "middle", charSpacing: 2,
    hyperlink: { url: "mailto:g.prado@externia.ai" },
  });

  // Secondary
  s.addShape("roundRect", {
    x: 5.1, y: ctaY, w: 4.3, h: 0.9, rectRadius: 0.1,
    fill: { color: C.surfaceHi }, line: { color: C.outline2, width: 1 },
    hyperlink: { url: "https://externia.ai", tooltip: "Web corporativa" },
  });
  s.addText("▲  externia.ai", {
    x: 5.1, y: ctaY + 0.05, w: 4.3, h: 0.5,
    fontFace: F.head, fontSize: 15, color: C.text, bold: true, align: "center", valign: "middle",
    hyperlink: { url: "https://externia.ai", tooltip: "Web corporativa" },
  });
  s.addText("Web corporativa · activaciones · blog", {
    x: 5.1, y: ctaY + 0.5, w: 4.3, h: 0.35,
    fontFace: F.mono, fontSize: 9, color: C.faint, align: "center", valign: "middle", charSpacing: 2,
    hyperlink: { url: "https://externia.ai" },
  });

  // Showcase link (the landing itself)
  s.addText("↗  externia-showcase.onrender.com — explora cada activación en vivo", {
    x: 0.5, y: 6.7, w: W - 1.0, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.magenta, align: "center", charSpacing: 2, bold: true,
    hyperlink: { url: "https://externia-showcase.onrender.com", tooltip: "El showcase completo" },
  });

  pageNum(s, 11);
}

// SAVE
const outFile = path.resolve(__dirname, "..", "Externia-Showcase-Deck.pptx");
pptx.writeFile({ fileName: outFile }).then(fn => {
  console.log("✓ Generated:", fn);
});
