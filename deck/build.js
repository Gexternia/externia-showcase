// Externia · Showcase Deck (PPTX)
// Generates a 10-slide commercial pitch deck mirroring the Showcase landing.

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
};

// Fonts (PowerPoint will substitute if not installed — Calibri is universal fallback)
const F = {
  head: "Calibri",      // Headlines
  body: "Calibri Light", // Body text
  mono: "Consolas",     // Meta / monospace
};

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 inches (16:9)
pptx.title = "Externia · Showcase 2026";
pptx.author = "Externia";
pptx.company = "Externia";
pptx.subject = "Showcase comercial";

const W = 13.333; // slide width
const H = 7.5;    // slide height

// === Reusable helpers ============================================
function darkBg(slide) {
  slide.background = { color: C.bg };
}
function eyebrow(slide, text, opts = {}) {
  const { x = 0.6, y = 0.45, color = C.magenta } = opts;
  slide.addText("▲ " + text, {
    x, y, w: 8, h: 0.3,
    fontFace: F.mono, fontSize: 11, bold: true,
    color, charSpacing: 4,
  });
}
function pageNum(slide, n, total = 10) {
  slide.addText(`${String(n).padStart(2, "0")} / ${String(total).padStart(2, "0")}`, {
    x: W - 1.4, y: H - 0.45, w: 1, h: 0.25,
    fontFace: F.mono, fontSize: 9, color: C.faint, align: "right", charSpacing: 3,
  });
}
function brand(slide, opts = {}) {
  const { x = W - 1.6, y = 0.4 } = opts;
  slide.addText("EXTERNIA", {
    x, y, w: 1.2, h: 0.25,
    fontFace: F.mono, fontSize: 10, bold: true, color: C.dim, charSpacing: 6, align: "right",
  });
}
function gradAccent(slide, x, y, w = 0.6, h = 0.06) {
  // Single magenta bar as the brand accent (PPTX gradient text isn't portable)
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: C.magenta }, line: { type: "none" },
  });
}
function softCard(slide, x, y, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: 0.15,
    fill: { color: C.void2 },
    line: { color: C.outline, width: 1 },
    ...opts,
  });
}

// ============================================================
// SLIDE 1 — COVER
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  // Decorative ambient circles (replace light leaks)
  s.addShape("ellipse", {
    x: -2.5, y: -2, w: 6, h: 6,
    fill: { color: C.magenta, transparency: 88 }, line: { type: "none" },
  });
  s.addShape("ellipse", {
    x: W - 3, y: H - 3, w: 5, h: 5,
    fill: { color: C.yellow, transparency: 92 }, line: { type: "none" },
  });

  // Icon
  s.addImage({ path: A("externia-icon-512.png"), x: W/2 - 0.9, y: H/2 - 1.9, w: 1.8, h: 1.8 });

  // Wordmark
  s.addText("externia", {
    x: 0, y: H/2 + 0.1, w: W, h: 1.4,
    fontFace: F.head, fontSize: 96, bold: true, color: C.magenta,
    align: "center", charSpacing: -2,
  });

  // Subtitle
  s.addText("Showcase · 2026", {
    x: 0, y: H/2 + 1.5, w: W, h: 0.4,
    fontFace: F.mono, fontSize: 14, color: C.dim, align: "center", charSpacing: 8,
  });

  // Bottom-left tagline
  s.addText("Producto  ·  Activaciones  ·  I+D", {
    x: 0.6, y: H - 0.7, w: 6, h: 0.3,
    fontFace: F.mono, fontSize: 11, color: C.faint, charSpacing: 4,
  });
  s.addText("g.prado@externia.ai", {
    x: W - 4.6, y: H - 0.7, w: 4, h: 0.3,
    fontFace: F.mono, fontSize: 11, color: C.faint, align: "right", charSpacing: 3,
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
    x: 0.6, y: 1.3, w: W - 1.2, h: 2.4,
    fontFace: F.head, fontSize: 64, bold: true,
    align: "center", charSpacing: -1, paraSpaceAfter: 0,
  });

  s.addText(
    "Producto en producción, activaciones para marcas y la investigación que define el próximo salto. Todo está vivo. Todo se puede tocar.",
    {
      x: 2.5, y: 3.9, w: W - 5, h: 1,
      fontFace: F.body, fontSize: 16, color: C.dim, align: "center",
    }
  );

  // Stats strip
  const stats = [
    { k: "8+", l: "demos públicas" },
    { k: "10", l: "meses operando" },
    { k: "100%", l: "hecho por Externia" },
    { k: "∞",   l: "blog autónomo" },
  ];
  const colW = 2.4;
  const rowX = (W - colW * 4) / 2;
  const rowY = 5.4;
  // Top + bottom rules
  s.addShape("line", { x: rowX, y: rowY, w: colW*4, h: 0, line: { color: C.outline2, width: 1 } });
  s.addShape("line", { x: rowX, y: rowY + 1.4, w: colW*4, h: 0, line: { color: C.outline2, width: 1 } });
  stats.forEach((st, i) => {
    const x = rowX + i * colW;
    if (i > 0) s.addShape("line", { x, y: rowY + 0.2, w: 0, h: 1, line: { color: C.outline2, width: 1 } });
    s.addText(st.k, {
      x, y: rowY + 0.18, w: colW, h: 0.7,
      fontFace: F.head, fontSize: 40, bold: true, color: C.magenta, align: "center",
    });
    s.addText(st.l, {
      x, y: rowY + 0.95, w: colW, h: 0.35,
      fontFace: F.mono, fontSize: 10, color: C.faint, align: "center", charSpacing: 3,
    });
  });

  pageNum(s, 2);
}

// ============================================================
// SLIDE 3 — EVENTOPLUS (cliente recurrente)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "A · Cliente recurrente en producción");

  // Logos row
  s.addImage({ path: A("eventoplus-logo.webp"), x: 0.6, y: 1.0, w: 0.7, h: 0.7 });
  s.addText("×", { x: 1.45, y: 1.0, w: 0.4, h: 0.7, fontFace: F.body, fontSize: 28, color: C.faint, align: "center", valign: "middle" });
  s.addImage({ path: A("finderia-logo.webp"), x: 1.85, y: 1.15, w: 2.0, h: 0.4 });

  // Live chip
  s.addShape("roundRect", {
    x: 4.1, y: 1.18, w: 3.3, h: 0.36, rectRadius: 0.2,
    fill: { color: "1A3A28" }, line: { color: "2B6A48", width: 1 },
  });
  s.addText("● Live · Cliente activo desde feb 2026", {
    x: 4.1, y: 1.18, w: 3.3, h: 0.36,
    fontFace: F.mono, fontSize: 10, color: C.green, align: "center", valign: "middle", charSpacing: 2, bold: true,
  });

  // Title
  s.addText([
    { text: "El stack conversacional\nde ", options: { color: C.text } },
    { text: "Eventoplus", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.6, y: 1.9, w: W/2 + 0.4, h: 1.7,
    fontFace: F.head, fontSize: 38, bold: true, charSpacing: -1,
  });

  // Description
  s.addText(
    "Chatbot de búsqueda + dashboard analítico + informe ejecutivo semanal. Tres piezas que conviven en un loop de producto, datos y consultoría — no es software entregado, es un servicio recurrente que evoluciona cada semana.",
    {
      x: 0.6, y: 3.6, w: W/2 + 0.4, h: 1.5,
      fontFace: F.body, fontSize: 14, color: C.dim, paraSpaceAfter: 6,
    }
  );

  // Services list
  const services = [
    { code: "01", name: "Chatbot FinderAI", url: "naranja-bot.onrender.com" },
    { code: "02", name: "Dashboard de estadísticas", url: "admin-frontend-b33o.onrender.com" },
    { code: "03", name: "Informe ejecutivo · cada lunes", url: "PDF · entregable consultivo" },
  ];
  services.forEach((srv, i) => {
    const y = 5.3 + i * 0.55;
    s.addShape("roundRect", {
      x: 0.6, y, w: W/2 + 0.2, h: 0.45, rectRadius: 0.08,
      fill: { color: C.surfaceHi }, line: { color: C.outline2, width: 1 },
    });
    s.addText(srv.code, {
      x: 0.78, y, w: 0.45, h: 0.45,
      fontFace: F.mono, fontSize: 10, color: C.magenta, valign: "middle", bold: true, charSpacing: 2,
    });
    s.addText(srv.name, {
      x: 1.3, y, w: 4.5, h: 0.45,
      fontFace: F.head, fontSize: 12, color: C.text, valign: "middle", bold: true,
    });
    s.addText(srv.url, {
      x: 1.3, y, w: 5, h: 0.45,
      fontFace: F.mono, fontSize: 9, color: C.faint, valign: "middle", align: "right", charSpacing: 1,
    });
  });

  // Right side: KPI panel
  const px = W/2 + 1.4;
  const py = 1.9;
  const pw = W - px - 0.6;
  const ph = H - py - 0.9;
  softCard(s, px, py, pw, ph);

  s.addText("SEMANA 19  ·  3 - 10 MAYO 2026", {
    x: px + 0.3, y: py + 0.2, w: pw - 0.6, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.faint, charSpacing: 3,
  });
  s.addShape("roundRect", {
    x: px + pw - 1.4, y: py + 0.2, w: 1.1, h: 0.3, rectRadius: 0.15,
    fill: { color: "1A3A28" }, line: { color: "2B6A48", width: 1 },
  });
  s.addText("+14% WoW", {
    x: px + pw - 1.4, y: py + 0.2, w: 1.1, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.green, align: "center", valign: "middle", bold: true, charSpacing: 1,
  });

  // KPI rows
  const rows = [
    { name: "Sesiones totales",    a: "32",  b: "28",   v: "+14%",     hot: false },
    { name: "Mensajes totales",    a: "116", b: "104",  v: "+12%",     hot: false },
    { name: "Mensajes / sesión",   a: "3,6", b: "3,7",  v: "≈ Estable", hot: false },
    { name: "Solicitudes cerradas",a: "1",   b: "0",    v: "+1",       hot: true  },
    { name: "Errores n8n",         a: "1",   b: "3",    v: "−67%",     hot: false },
  ];
  const rowH = 0.65;
  const tableY = py + 0.7;
  rows.forEach((r, i) => {
    const ry = tableY + i * (rowH + 0.05);
    s.addShape("roundRect", {
      x: px + 0.25, y: ry, w: pw - 0.5, h: rowH, rectRadius: 0.08,
      fill: { color: r.hot ? "2A1614" : "151415" },
      line: { color: r.hot ? "884226" : C.outline, width: 1 },
    });
    s.addText(r.name, {
      x: px + 0.4, y: ry, w: 2.2, h: rowH,
      fontFace: F.head, fontSize: 12, color: C.text, valign: "middle", bold: true,
    });
    s.addText(r.a, {
      x: px + 2.6, y: ry, w: 1, h: rowH,
      fontFace: F.head, fontSize: 18, color: C.magenta, valign: "middle", align: "center", bold: true,
    });
    s.addText(r.b, {
      x: px + 3.6, y: ry, w: 1, h: rowH,
      fontFace: F.mono, fontSize: 12, color: C.faint, valign: "middle", align: "center",
    });
    s.addShape("roundRect", {
      x: px + 4.7, y: ry + 0.18, w: pw - 5.0, h: rowH - 0.36, rectRadius: 0.15,
      fill: { color: "1A3A28" }, line: { color: "2B6A48", width: 1 },
    });
    s.addText(r.v, {
      x: px + 4.7, y: ry + 0.18, w: pw - 5.0, h: rowH - 0.36,
      fontFace: F.mono, fontSize: 10, color: C.green, valign: "middle", align: "center", bold: true,
    });
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

  // Title
  s.addText("Informe Ejecutivo Semanal", {
    x: 0.6, y: 1.0, w: W - 1.2, h: 0.9,
    fontFace: F.head, fontSize: 44, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText("PDF · 4 páginas · entregado cada lunes a Eventoplus", {
    x: 0.6, y: 1.95, w: W - 1.2, h: 0.35,
    fontFace: F.mono, fontSize: 12, color: C.yellow, charSpacing: 3,
  });

  // Lead
  s.addText("Cada lunes, el equipo de Eventoplus recibe un análisis ejecutivo entregable con:", {
    x: 0.6, y: 2.6, w: W - 1.2, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.dim,
  });

  // Bullets (2 columns)
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
    const x = 0.6 + col * 5.8;
    const y = 3.2 + row * 0.55;
    s.addText("▸", {
      x, y, w: 0.3, h: 0.45,
      fontFace: F.head, fontSize: 16, color: C.yellow, bold: true, valign: "middle",
    });
    s.addText(b, {
      x: x + 0.3, y, w: 5.3, h: 0.45,
      fontFace: F.body, fontSize: 14, color: C.text, valign: "middle",
    });
  });

  // Closer (italic)
  s.addShape("line", { x: 0.6, y: 5.6, w: W - 1.2, h: 0, line: { color: C.outline2, width: 1, dashType: "dash" } });
  s.addText([
    { text: "Es la diferencia entre vender un chatbot y vender ", options: { color: C.text, italic: true } },
    { text: "inteligencia operativa continua", options: { color: C.magenta, italic: false, bold: true } },
    { text: ".", options: { color: C.text, italic: true } },
  ], {
    x: 0.6, y: 5.85, w: W - 1.2, h: 0.9,
    fontFace: F.head, fontSize: 22, italic: true,
  });

  pageNum(s, 4);
}

// ============================================================
// SLIDE 5 — ACTIVACIONES (grid)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "B · Activaciones para marcas");

  s.addText([
    { text: "Producto creado para marcas finales, ", options: { color: C.text } },
    { text: "en vivo", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.6, y: 0.95, w: W - 1.2, h: 0.85,
    fontFace: F.head, fontSize: 32, bold: true, charSpacing: -1,
  });

  // 5 tiles: 3 top + 2 bottom (centered)
  const tiles = [
    { client: "KPMG",         title: "ComicGen",                  blurb: "Cómic generado con IA en directo para Encuentro Alumni — los asistentes son protagonistas.", url: "comic-ai-s1ca.onrender.com",   image: A("comicgen-kpmg.png") },
    { client: "Clinique · INHOUSE", title: "Cartas personalizadas",  blurb: "Generador de cartas de marca con narrativa coherente, listas para imprimir y entregar.",       url: "clinique-bium.onrender.com",  image: A("clinique-letter.png") },
    { client: "ING · Taller IA",    title: "SmartBrush + caricaturas", blurb: "Suite visual con IA: caricaturización en vivo, edición selectiva por pincel, retratos.",     url: "taller-ia-ing.onrender.com",  image: A("smartbrush-caricature.png") },
    { client: "Talento Joven",      title: "Retro Game",              blurb: "Videojuego retro generado para evento corporativo de jóvenes talento. Sin instalación.",     url: "videogame-witg.onrender.com", image: null },
    { client: "Externia · Interno", title: "DeckCraft",               blurb: "Generador de presentaciones desde brief + sistema de diseño en Markdown. Producción 10×.",   url: "deckcraft-3u04.onrender.com", image: A("deckcraft-sample.png") },
  ];

  const tileW = 4.0;
  const tileH = 2.4;
  const gap = 0.2;
  const startX = (W - (tileW * 3 + gap * 2)) / 2;
  const topY = 2.05;
  const bottomY = topY + tileH + gap;
  const bottomStartX = (W - (tileW * 2 + gap)) / 2;

  tiles.forEach((t, i) => {
    const isTop = i < 3;
    const x = isTop ? startX + i * (tileW + gap) : bottomStartX + (i - 3) * (tileW + gap);
    const y = isTop ? topY : bottomY;
    // Card
    softCard(s, x, y, tileW, tileH);
    // Thumb area (top 55%)
    const thumbH = tileH * 0.55;
    s.addShape("roundRect", {
      x: x + 0.08, y: y + 0.08, w: tileW - 0.16, h: thumbH - 0.08,
      rectRadius: 0.08,
      fill: { color: "0f0e10" }, line: { color: C.outline, width: 1 },
    });
    if (t.image) {
      s.addImage({ path: t.image, x: x + 0.1, y: y + 0.1, w: tileW - 0.2, h: thumbH - 0.12, sizing: { type: "cover", w: tileW - 0.2, h: thumbH - 0.12 } });
    } else {
      // Retro Game placeholder — pixel mock
      const pxColors = [C.yellow, C.magenta, C.yellow, C.orange, C.magenta, C.yellow];
      pxColors.forEach((col, pi) => {
        s.addShape("rect", {
          x: x + 0.7 + pi * 0.4, y: y + 0.6, w: 0.22, h: 0.22,
          fill: { color: col }, line: { type: "none" },
        });
      });
    }
    // Client tag (corner overlay)
    s.addText("● " + t.client, {
      x: x + 0.2, y: y + 0.18, w: tileW - 0.4, h: 0.25,
      fontFace: F.mono, fontSize: 8, color: C.text, charSpacing: 2, bold: true,
    });
    // Title
    s.addText(t.title, {
      x: x + 0.2, y: y + thumbH + 0.05, w: tileW - 0.4, h: 0.35,
      fontFace: F.head, fontSize: 16, bold: true, color: C.text, charSpacing: -0.5,
    });
    // Blurb
    s.addText(t.blurb, {
      x: x + 0.2, y: y + thumbH + 0.42, w: tileW - 0.4, h: 0.5,
      fontFace: F.body, fontSize: 9.5, color: C.dim,
    });
    // URL
    s.addText("↗ " + t.url, {
      x: x + 0.2, y: y + tileH - 0.35, w: tileW - 0.4, h: 0.25,
      fontFace: F.mono, fontSize: 8, color: C.magenta, charSpacing: 1, bold: true,
    });
  });

  pageNum(s, 5);
}

// ============================================================
// SLIDE 6 — TURESPAÑA: Reto + Externa vs Externia
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
    x: 0.6, y: 0.95, w: W - 1.2, h: 1.7,
    fontFace: F.head, fontSize: 38, bold: true, charSpacing: -1, paraSpaceAfter: 0,
  });

  // Reto card
  softCard(s, 0.6, 2.8, W - 1.2, 1.2);
  s.addText("▼ EL RETO", {
    x: 0.85, y: 2.95, w: 4, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.yellow, bold: true, charSpacing: 3,
  });
  s.addText(
    "Concurso público tier-1, criterios técnicos estrictos, competencia de agencias top. Memoria densa, escenografía ambiciosa, narrativa coherente con el «por qué» de Turespaña.",
    {
      x: 0.85, y: 3.25, w: W - 1.7, h: 0.7,
      fontFace: F.body, fontSize: 13, color: C.text,
    }
  );

  // Two columns
  const colW = (W - 1.4) / 2;
  const colY = 4.25;
  const colH = 2.6;

  // Externa col
  softCard(s, 0.6, colY, colW - 0.05, colH);
  s.addText("Externa", {
    x: 0.85, y: colY + 0.2, w: colW - 0.5, h: 0.4,
    fontFace: F.head, fontSize: 20, bold: true, color: C.text,
  });
  s.addText("Lo que aportó", {
    x: 0.85, y: colY + 0.6, w: colW - 0.5, h: 0.25,
    fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 3,
  });
  const externaItems = [
    "Cartera tier-1 — Amazon, ING, Caixabank, KPMG, Cámara de Comercio.",
    "Criterio del equipo sobre qué funciona en cada bloque de memoria.",
    "Relación de confianza con Turespaña y dominio del cliente institucional.",
  ];
  externaItems.forEach((it, i) => {
    const y = colY + 1.05 + i * 0.55;
    s.addText("▸", { x: 0.85, y, w: 0.25, h: 0.4, fontFace: F.head, fontSize: 12, color: C.faint, bold: true, valign: "top" });
    s.addText(it, { x: 1.10, y, w: colW - 0.7, h: 0.5, fontFace: F.body, fontSize: 12, color: C.text });
  });

  // Externia col
  const x2 = 0.6 + colW + 0.1;
  softCard(s, x2, colY, colW - 0.05, colH, { line: { color: C.magenta, width: 1 } });
  s.addText("Externia", {
    x: x2 + 0.25, y: colY + 0.2, w: colW - 0.5, h: 0.4,
    fontFace: F.head, fontSize: 20, bold: true, color: C.magenta,
  });
  s.addText("Lo que aportó", {
    x: x2 + 0.25, y: colY + 0.6, w: colW - 0.5, h: 0.25,
    fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 3,
  });
  const externiaItems = [
    "Agente entrenado con histórico de concursos Turespaña y feedback público.",
    "Inteligencia interna de Externa sobre qué funciona por bloque, hecha sistema.",
    "Investigación profunda automática del «porqué» temático del cliente.",
    "Renders de escenografía generados con IA — propuesta visualmente rica.",
  ];
  externiaItems.forEach((it, i) => {
    const y = colY + 1.05 + i * 0.42;
    s.addText("▸", { x: x2 + 0.25, y, w: 0.25, h: 0.4, fontFace: F.head, fontSize: 12, color: C.magenta, bold: true, valign: "top" });
    s.addText(it, { x: x2 + 0.50, y, w: colW - 0.7, h: 0.42, fontFace: F.body, fontSize: 11.5, color: C.text });
  });

  pageNum(s, 6);
}

// ============================================================
// SLIDE 7 — TURESPAÑA: Quote (impact)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  // Ambient circles
  s.addShape("ellipse", {
    x: W/2 - 5, y: -2, w: 10, h: 10,
    fill: { color: C.magenta, transparency: 92 }, line: { type: "none" },
  });

  brand(s);
  eyebrow(s, "C · Cita del cliente");

  // Massive opening quote mark
  s.addText("“", {
    x: 0.6, y: 1.4, w: 2, h: 2,
    fontFace: F.head, fontSize: 200, color: C.magenta, bold: true,
  });

  // Quote
  s.addText([
    { text: "La IA no ganó la convención.\n", options: { color: C.text } },
    { text: "La ganó ", options: { color: C.text } },
    { text: "Externa", options: { color: C.magenta, bold: true } },
    { text: ".", options: { color: C.text } },
  ], {
    x: 1.5, y: 2.6, w: W - 3, h: 2.2,
    fontFace: F.head, fontSize: 56, bold: true, align: "center", charSpacing: -2,
  });

  // Subquote
  s.addText(
    "Pero la IA hizo posible una propuesta más rica de la que habríamos podido permitirnos meter en cada concurso.",
    {
      x: 2, y: 5.0, w: W - 4, h: 1.0,
      fontFace: F.body, fontSize: 18, color: C.dim, align: "center", italic: true,
    }
  );

  // Result chip
  s.addShape("roundRect", {
    x: W/2 - 2.4, y: 6.4, w: 4.8, h: 0.6, rectRadius: 0.3,
    fill: { color: "1A3A28" }, line: { color: "2B6A48", width: 1 },
  });
  s.addText("✓  Externa ganó la V Convención", {
    x: W/2 - 2.4, y: 6.4, w: 4.8, h: 0.6,
    fontFace: F.head, fontSize: 16, color: C.green, bold: true, align: "center", valign: "middle",
  });

  pageNum(s, 7);
}

// ============================================================
// SLIDE 8 — METODOLOGÍA (timeline)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "D · Cómo trabajamos");

  s.addText("Cómo entramos en una agencia", {
    x: 0.6, y: 0.95, w: W - 1.2, h: 0.9,
    fontFace: F.head, fontSize: 40, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText(
    "Lo primero que hacemos no es construir. Es observar, mapear y decidir qué construir primero.",
    {
      x: 0.6, y: 1.85, w: W - 1.2, h: 0.5,
      fontFace: F.body, fontSize: 15, color: C.dim,
    }
  );

  // Timeline line (horizontal)
  const tlY = 3.0;
  const tlPad = 0.6;
  const tlW = W - tlPad * 2;
  s.addShape("line", { x: tlPad, y: tlY, w: tlW, h: 0, line: { color: C.outline2, width: 2 } });
  // Gradient overlay (magenta-orange-yellow) — approximated as a magenta bar full width
  s.addShape("rect", { x: tlPad, y: tlY - 0.04, w: tlW, h: 0.08, fill: { color: C.magenta, transparency: 50 }, line: { type: "none" } });

  const phases = [
    { n: "01", when: "Semanas 1-3", name: "Inmersión", hot: true,
      desc: "Dentro del equipo en cada departamento. Mapa de tareas que consumen tiempo, las que bloquean valor, quick wins ya identificados." },
    { n: "02", when: "Meses 1-2", name: "Capa rápida",
      desc: "Ganancias visibles en semanas. Agente redactor, transcripción a informe, generador de decks." },
    { n: "03", when: "Meses 2-4", name: "Capa profunda",
      desc: "Agentes conectados a datos históricos. 20 años de propuestas consultables, pliegos estructurados, feedback público digerido." },
    { n: "04", when: "Meses 4-9", name: "Escala",
      desc: "Sistema operando que mejora con cada concurso. Modelo de uso interno, casos abiertos del cliente, roadmap consolidado." },
  ];
  const phW = tlW / 4;
  phases.forEach((p, i) => {
    const x = tlPad + i * phW;
    // Number circle
    s.addShape("ellipse", {
      x: x + phW/2 - 0.3, y: tlY - 0.3, w: 0.6, h: 0.6,
      fill: { color: p.hot ? C.magenta : C.surfaceHi },
      line: { color: p.hot ? C.magenta : C.outline2, width: 1.5 },
    });
    s.addText(p.n, {
      x: x + phW/2 - 0.3, y: tlY - 0.3, w: 0.6, h: 0.6,
      fontFace: F.mono, fontSize: 12, color: p.hot ? "1A0008" : C.dim, bold: true, align: "center", valign: "middle",
    });
    // When
    s.addText(p.when, {
      x: x + 0.2, y: tlY + 0.55, w: phW - 0.4, h: 0.3,
      fontFace: F.mono, fontSize: 10, color: C.faint, align: "center", charSpacing: 3,
    });
    // Name
    s.addText(p.name, {
      x: x + 0.2, y: tlY + 0.95, w: phW - 0.4, h: 0.5,
      fontFace: F.head, fontSize: 20, bold: true,
      color: p.hot ? C.magenta : C.text, align: "center", charSpacing: -0.5,
    });
    // Desc
    s.addText(p.desc, {
      x: x + 0.2, y: tlY + 1.6, w: phW - 0.4, h: 1.6,
      fontFace: F.body, fontSize: 11, color: C.dim, align: "center",
    });
  });

  // Felt block
  const fY = 6.0;
  s.addShape("line", { x: 0.6, y: fY - 0.15, w: W - 1.2, h: 0, line: { color: C.outline, width: 1, dashType: "dash" } });
  s.addText("▼ DESDE EL PRIMER MES", {
    x: 0.6, y: fY, w: 4, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.yellow, bold: true, charSpacing: 3,
  });
  s.addText([
    { text: "No sienten que les han metido un proyecto encima. Sienten que tienen ", options: { color: C.text } },
    { text: "herramientas nuevas", options: { color: C.magenta, bold: true } },
    { text: " que están haciendo el trabajo aburrido por ellos.", options: { color: C.text } },
  ], {
    x: 0.6, y: fY + 0.35, w: W - 1.2, h: 1.0,
    fontFace: F.head, fontSize: 18, italic: true,
  });

  pageNum(s, 8);
}

// ============================================================
// SLIDE 9 — I+D ABIERTO (3D Worlds)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);
  brand(s);
  eyebrow(s, "E · I+D abierto · Q2 2026", { color: C.yellow });

  s.addText("Plataforma 3D · World API", {
    x: 0.6, y: 0.95, w: W - 1.2, h: 0.9,
    fontFace: F.head, fontSize: 40, bold: true, color: C.text, charSpacing: -1,
  });
  s.addText([
    { text: "No solo construimos lo que se compra hoy. Investigamos lo que va a ", options: { color: C.dim } },
    { text: "redefinir el sector", options: { color: C.magenta, bold: true } },
    { text: " en los próximos 18 meses.", options: { color: C.dim } },
  ], {
    x: 0.6, y: 1.95, w: W - 1.2, h: 0.5,
    fontFace: F.body, fontSize: 15,
  });

  // 4 KPIs row
  const kY = 2.7;
  const kpis = [
    { v: "Validada",     l: "Tecnología" },
    { v: "En formación", l: "Mercado" },
    { v: "12-18 m",      l: "Ventana competitiva" },
    { v: "< $100",       l: "Coste de validación" },
  ];
  const kpiW = (W - 1.4) / 4;
  kpis.forEach((kp, i) => {
    const x = 0.6 + i * (kpiW + 0.05);
    softCard(s, x, kY, kpiW - 0.1, 1.1);
    s.addText(kp.v, {
      x: x + 0.2, y: kY + 0.18, w: kpiW - 0.5, h: 0.4,
      fontFace: F.head, fontSize: 18, color: C.magenta, bold: true,
    });
    s.addText(kp.l, {
      x: x + 0.2, y: kY + 0.65, w: kpiW - 0.5, h: 0.3,
      fontFace: F.mono, fontSize: 9, color: C.faint, charSpacing: 3,
    });
  });

  // Backers strip
  const bY = 4.05;
  s.addShape("line", { x: 0.6, y: bY, w: W - 1.2, h: 0, line: { color: C.outline, width: 1 } });
  s.addText("BACKERS", {
    x: 0.6, y: bY + 0.15, w: 1.5, h: 0.3,
    fontFace: F.mono, fontSize: 10, color: C.faint, charSpacing: 4,
  });
  s.addText("NVIDIA  ·  AMD  ·  Autodesk  ·  $1.290M en financiación · World API pública desde enero 2026", {
    x: 2.1, y: bY + 0.15, w: W - 2.7, h: 0.3,
    fontFace: F.mono, fontSize: 11, color: C.text, charSpacing: 2, bold: true,
  });

  // 3 puntos
  const pY = 4.85;
  const points = [
    { n: "01", t: "World API",         d: "El motor de World Labs. Una sola foto → mundo 3D navegable persistente." },
    { n: "02", t: "Gaussian Splatting", d: "Renderizado que alcanzó madurez de producción en 2025. Curva de coste ya cruzó." },
    { n: "03", t: "First mover MICE",   d: "Nadie ha verticalizado esta tecnología para sector eventos. Ventana real." },
  ];
  const ptW = (W - 1.4) / 3;
  points.forEach((p, i) => {
    const x = 0.6 + i * (ptW + 0.1);
    softCard(s, x, pY, ptW - 0.1, 1.5);
    s.addText(p.n + " · " + p.t.toUpperCase(), {
      x: x + 0.2, y: pY + 0.2, w: ptW - 0.5, h: 0.3,
      fontFace: F.mono, fontSize: 10, color: C.magenta, charSpacing: 2, bold: true,
    });
    s.addText(p.d, {
      x: x + 0.2, y: pY + 0.6, w: ptW - 0.5, h: 0.85,
      fontFace: F.body, fontSize: 12, color: C.text,
    });
  });

  // Punch line
  s.addText([
    { text: "Cliente que entre ahora en esta tecnología tiene ", options: { color: C.text } },
    { text: "18 meses de ventaja competitiva real", options: { color: C.magenta, bold: true } },
    { text: ".", options: { color: C.text } },
  ], {
    x: 0.6, y: 6.55, w: W - 1.2, h: 0.5,
    fontFace: F.head, fontSize: 17, bold: false, italic: true, align: "center",
  });

  pageNum(s, 9);
}

// ============================================================
// SLIDE 10 — CIERRE (Blog + Contacto)
// ============================================================
{
  const s = pptx.addSlide();
  darkBg(s);

  // Ambient circles
  s.addShape("ellipse", {
    x: -3, y: H - 4, w: 8, h: 8,
    fill: { color: C.magenta, transparency: 92 }, line: { type: "none" },
  });
  s.addShape("ellipse", {
    x: W - 4, y: -3, w: 7, h: 7,
    fill: { color: C.yellow, transparency: 94 }, line: { type: "none" },
  });

  brand(s);
  eyebrow(s, "F · Pipeline editorial autónomo", { color: C.yellow });

  // Blog proof
  softCard(s, 0.6, 1.0, W - 1.2, 1.8);
  s.addText("externia.ai/blog", {
    x: 0.85, y: 1.2, w: W - 1.7, h: 0.3,
    fontFace: F.mono, fontSize: 11, color: C.faint, charSpacing: 4,
  });
  s.addText([
    { text: "El blog de Externia se ", options: { color: C.text } },
    { text: "escribe solo", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.85, y: 1.55, w: W - 1.7, h: 0.7,
    fontFace: F.head, fontSize: 28, bold: true, charSpacing: -1,
  });
  s.addText(
    "Posts publicados sin intervención humana. Investigación, redacción, optimización SEO, formato, publicación — todo lo opera un sistema de agentes Externia desde hace meses.",
    {
      x: 0.85, y: 2.25, w: W - 1.7, h: 0.5,
      fontFace: F.body, fontSize: 12, color: C.dim,
    }
  );

  // The closing pitch
  s.addText([
    { text: "Esto opera hoy\n", options: { color: C.text } },
    { text: "Hablemos de ", options: { color: C.text } },
    { text: "lo siguiente", options: { color: C.magenta, bold: true } },
  ], {
    x: 0.6, y: 3.4, w: W - 1.2, h: 1.8,
    fontFace: F.head, fontSize: 64, bold: true, align: "center", charSpacing: -2,
  });

  // Contact CTA buttons
  const ctaY = 5.7;

  // Primary CTA
  s.addShape("roundRect", {
    x: W/2 - 4.6, y: ctaY, w: 4.4, h: 0.85, rectRadius: 0.12,
    fill: { color: C.magenta }, line: { type: "none" },
  });
  s.addText("✉  g.prado@externia.ai", {
    x: W/2 - 4.6, y: ctaY, w: 4.4, h: 0.5,
    fontFace: F.head, fontSize: 17, color: "1A0008", bold: true, align: "center", valign: "middle",
  });
  s.addText("Una conversación, no un formulario", {
    x: W/2 - 4.6, y: ctaY + 0.45, w: 4.4, h: 0.35,
    fontFace: F.mono, fontSize: 10, color: "5C0018", align: "center", valign: "middle", charSpacing: 2,
  });

  // Secondary CTA
  s.addShape("roundRect", {
    x: W/2 + 0.2, y: ctaY, w: 4.4, h: 0.85, rectRadius: 0.12,
    fill: { color: C.surfaceHi }, line: { color: C.outline2, width: 1 },
  });
  s.addText("▲  externia.ai", {
    x: W/2 + 0.2, y: ctaY, w: 4.4, h: 0.5,
    fontFace: F.head, fontSize: 17, color: C.text, bold: true, align: "center", valign: "middle",
  });
  s.addText("Web corporativa · activaciones · blog", {
    x: W/2 + 0.2, y: ctaY + 0.45, w: 4.4, h: 0.35,
    fontFace: F.mono, fontSize: 10, color: C.faint, align: "center", valign: "middle", charSpacing: 2,
  });

  pageNum(s, 10);
}

// ============================================================
// SAVE
// ============================================================
const outFile = path.resolve(__dirname, "..", "Externia-Showcase-Deck.pptx");
pptx.writeFile({ fileName: outFile }).then(fn => {
  console.log("✓ Generated:", fn);
});
