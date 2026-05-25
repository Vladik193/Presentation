const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Корпоративная культура";

// ─── Palette: warm cream + deep teal + gold accent ───────────────────────────
const C = {
  bg:        "FAFAF7",   // warm cream - main background
  bgAlt:     "F0F4F2",   // light sage tint
  bgDark:    "1C3A35",   // deep forest - title/accent slides
  teal:      "2A7B6F",   // primary teal
  tealLight: "E6F4F1",   // light teal
  tealMid:   "4A9E8E",   // mid teal
  gold:      "C9961A",   // gold accent
  goldLight: "FDF5DC",   // light gold
  coral:     "D4634A",   // warm coral
  coralLight:"FDF0EC",
  sage:      "7BAE97",   // sage green
  sageLight: "EBF5F0",
  textDark:  "1A2E2A",
  textMid:   "4A6560",
  textLight: "8AA89F",
  border:    "C8DDD8",
};

// ─── Helpers ─────────────────────────────────────────────────────────────────
function rect(s, x, y, w, h, fill, lineColor, lineW = 0) {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: lineColor || fill, width: lineW }
  });
}
function oval(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } });
}
function txt(s, text, x, y, w, h, opts = {}) {
  s.addText(text, { x, y, w, h, fontFace: "Calibri", margin: 0, ...opts });
}

// Slide label header (teal bar left + title)
function slideHeader(s, title) {
  rect(s, 0, 0, 0.18, 5.63, C.teal);
  rect(s, 0.18, 0, 9.82, 0.78, C.bgAlt);
  txt(s, title, 0.38, 0.12, 9.3, 0.56, {
    fontSize: 21, bold: true, color: C.textDark, align: "left", valign: "middle"
  });
  rect(s, 0.18, 0.78, 9.82, 0.04, C.border);
}

// Rounded badge
function badge(s, text, x, y, w, h, bgColor, textColor) {
  rect(s, x, y, w, h, bgColor, bgColor, 0);
  txt(s, text, x, y, w, h, { fontSize: 11, bold: true, color: textColor, align: "center", valign: "middle" });
}

// Card with left accent bar
function accentCard(s, x, y, w, h, barColor, bgColor, title, body, titleSize = 12.5, bodySize = 11.5) {
  rect(s, x, y, w, h, bgColor, C.border, 0.6);
  rect(s, x, y, 0.12, h, barColor, barColor, 0);
  if (title) txt(s, title, x + 0.22, y + 0.1, w - 0.32, 0.32, { fontSize: titleSize, bold: true, color: barColor, align: "left" });
  if (body) txt(s, body, x + 0.22, y + (title ? 0.44 : 0.12), w - 0.32, h - (title ? 0.54 : 0.22), { fontSize: bodySize, color: C.textDark, align: "left" });
}

// ─── SLIDE 1: Title ──────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  // Decorative geometry
  rect(s, 0, 0, 0.22, 5.63, C.teal);
  rect(s, 0.22, 4.8, 9.78, 0.83, "162E29");
  oval(s, 7.2, -0.8, 3.5, 3.5, "243F3A");
  oval(s, 6.8, 3.8, 2.2, 2.2, "1A302C");

  // Gold top line
  rect(s, 0.22, 0, 9.78, 0.06, C.gold);

  // Main text block
  txt(s, "КОРПОРАТИВНАЯ", {
    x: 0.55, y: 0.55, w: 9, h: 0.95,
    fontSize: 50, bold: true, color: "FFFFFF", align: "left", charSpacing: 2
  });
  txt(s, "КУЛЬТУРА", {
    x: 0.55, y: 1.42, w: 9, h: 0.9,
    fontSize: 50, bold: true, color: C.gold, align: "left", charSpacing: 2
  });

  rect(s, 0.55, 2.42, 5.5, 0.05, C.tealMid);

  txt(s, "Дисциплина: Управление человеческими ресурсами", {
    x: 0.55, y: 2.62, w: 8.5, h: 0.42,
    fontSize: 14, color: "9BBFB8", align: "left"
  });
  txt(s, "Ценности, нормы и поведение как основа эффективной организации", {
    x: 0.55, y: 3.1, w: 8.5, h: 0.48,
    fontSize: 13, color: "6A8F89", align: "left", italic: true
  });

  // Bottom tag
  txt(s, "УЧР  ·  Корпоративная культура", {
    x: 0.55, y: 5.1, w: 5, h: 0.35,
    fontSize: 11, color: "4A706A", align: "left"
  });
}

// ─── SLIDE 2: Contents ───────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "СОДЕРЖАНИЕ");

  const items = [
    ["01", "Понятие корпоративной культуры",       C.teal],
    ["02", "Структура и элементы культуры",         C.gold],
    ["03", "Типология корпоративных культур",       C.coral],
    ["04", "Функции корпоративной культуры",        C.sage],
    ["05", "Уровни культуры по Э. Шейну",          C.teal],
    ["06", "Влияние на персонал и мотивацию",       C.gold],
    ["07", "Формирование и управление культурой",   C.coral],
    ["08", "Диагностика и оценка",                 C.sage],
    ["09", "Трансформация и изменение культуры",   C.teal],
    ["10", "Выводы",                               C.textMid],
  ];

  items.forEach(([num, label, color], i) => {
    const col = i < 5 ? 0 : 1;
    const row = i % 5;
    const x = col === 0 ? 0.38 : 5.25;
    const y = 0.98 + row * 0.92;

    rect(s, x, y, 4.65, 0.78, i % 2 === 0 ? C.tealLight : C.bg, C.border, 0.5);
    // Number circle
    oval(s, x + 0.12, y + 0.12, 0.52, 0.52, color);
    txt(s, num, x + 0.12, y + 0.12, 0.52, 0.52, { fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, label, x + 0.76, y + 0.14, 3.78, 0.5, { fontSize: 12.5, color: C.textDark, align: "left", valign: "middle" });
  });
}

// ─── SLIDE 3: Definition ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ПОНЯТИЕ КОРПОРАТИВНОЙ КУЛЬТУРЫ");

  // Big definition box
  rect(s, 0.38, 0.98, 9.24, 1.3, C.bgDark, C.bgDark, 0);
  rect(s, 0.38, 0.98, 0.18, 1.3, C.gold, C.gold, 0);
  txt(s, "«Корпоративная культура — это система разделяемых ценностей, убеждений, норм поведения и традиций, которые определяют способ мышления и действия сотрудников организации»", {
    x: 0.7, y: 1.04, w: 8.7, h: 1.18,
    fontSize: 13, color: "FFFFFF", italic: true, align: "left", valign: "middle"
  });

  // Three columns: что это, зачем, кто формирует
  const cols = [
    { title: "Что это такое", color: C.teal, bg: C.tealLight, items: ["Неписаные правила игры", "«Дух» компании", "Коллективная идентичность", "Поведенческий код"] },
    { title: "Из чего состоит", color: C.gold, bg: C.goldLight, items: ["Ценности и миссия", "Традиции и ритуалы", "Язык и символы", "Истории и герои"] },
    { title: "Кто формирует", color: C.coral, bg: C.coralLight, items: ["Основатели компании", "Топ-менеджмент", "HR и внутренние лидеры", "Сами сотрудники"] },
  ];

  cols.forEach((col, i) => {
    const x = 0.38 + i * 3.12;
    rect(s, x, 2.48, 2.94, 2.88, col.bg, col.color, 0.8);
    rect(s, x, 2.48, 2.94, 0.4, col.color, col.color, 0);
    txt(s, col.title, x + 0.12, 2.5, 2.7, 0.36, { fontSize: 13, bold: true, color: "FFFFFF", align: "center" });
    col.items.forEach((item, j) => {
      txt(s, [{ text: "▸  ", options: { color: col.color, bold: true } }, { text: item, options: { color: C.textDark } }], {
        x: x + 0.18, y: 2.98 + j * 0.55, w: 2.6, h: 0.5,
        fontSize: 12, align: "left"
      });
    });
  });
}

// ─── SLIDE 4: Structure & Elements ───────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "СТРУКТУРА И ЭЛЕМЕНТЫ КОРПОРАТИВНОЙ КУЛЬТУРЫ");

  // Central "Culture" circle
  oval(s, 4.05, 2.0, 1.9, 1.9, C.bgDark);
  txt(s, "КУЛЬТУРА", 4.05, 2.0, 1.9, 1.9, { fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle" });

  // 6 satellite elements
  const elements = [
    { label: "Ценности\nи миссия",    x: 0.38, y: 1.0,  color: C.teal,  bg: C.tealLight },
    { label: "Нормы\nповедения",      x: 0.38, y: 2.7,  color: C.teal,  bg: C.tealLight },
    { label: "Ритуалы\nи традиции",   x: 2.05, y: 0.92, color: C.gold,  bg: C.goldLight },
    { label: "Символы\nи артефакты",  x: 2.05, y: 3.62, color: C.gold,  bg: C.goldLight },
    { label: "Герои\nи лидеры",       x: 6.38, y: 0.92, color: C.coral, bg: C.coralLight },
    { label: "Коммуни-\nкации",       x: 6.38, y: 3.62, color: C.coral, bg: C.coralLight },
    { label: "Климат\nи среда",       x: 7.82, y: 1.0,  color: C.sage,  bg: C.sageLight },
    { label: "Истории\nи мифы",       x: 7.82, y: 2.7,  color: C.sage,  bg: C.sageLight },
  ];

  elements.forEach(el => {
    rect(s, el.x, el.y, 1.58, 0.82, el.bg, el.color, 0.8);
    txt(s, el.label, el.x + 0.08, el.y + 0.04, 1.42, 0.74, { fontSize: 11, bold: true, color: el.color, align: "center", valign: "middle" });
  });

  // Lines from center to elements (connecting dots)
  const centerX = 5.0, centerY = 2.95;
  const dots = [
    [1.96, 1.41], [1.96, 3.11], [3.63, 1.41], [3.63, 3.73],
    [6.38, 1.41], [6.38, 3.73], [7.82, 1.41], [7.82, 3.11]
  ];
  dots.forEach(([dx, dy]) => {
    oval(s, dx - 0.06, dy - 0.06, 0.12, 0.12, C.border);
  });
}

// ─── SLIDE 5: Typology (Cameron & Quinn) ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ТИПОЛОГИЯ КОРПОРАТИВНЫХ КУЛЬТУР (КАМЕРОН И КУИНН)");

  // Axis labels
  txt(s, "Гибкость и дискретность  ↑", 1.5, 0.92, 7, 0.32, { fontSize: 11, color: C.textLight, align: "center" });
  txt(s, "↓  Стабильность и контроль", 1.5, 5.15, 7, 0.32, { fontSize: 11, color: C.textLight, align: "center" });
  txt(s, "←  Внутренний фокус", 0.2, 2.8, 1.4, 0.32, { fontSize: 10, color: C.textLight, align: "center", rotate: 90 });
  txt(s, "Внешний фокус  →", 8.5, 2.8, 1.2, 0.32, { fontSize: 10, color: C.textLight, align: "center", rotate: 90 });

  const quadrants = [
    { name: "КЛАН", sub: "Семейная культура", x: 0.38, y: 1.28,
      color: C.teal, bg: C.tealLight,
      items: ["Командная работа", "Наставничество", "Вовлечённость", "Лояльность сотрудников"] },
    { name: "АДХОКРАТИЯ", sub: "Культура творчества", x: 5.12, y: 1.28,
      color: C.gold, bg: C.goldLight,
      items: ["Инновации и риск", "Предприимчивость", "Гибкость", "Рост и развитие"] },
    { name: "ИЕРАРХИЯ", sub: "Культура контроля", x: 0.38, y: 3.28,
      color: C.coral, bg: C.coralLight,
      items: ["Процессы и правила", "Эффективность", "Стабильность", "Стандартизация"] },
    { name: "РЫНОК", sub: "Конкурентная культура", x: 5.12, y: 3.28,
      color: C.sage, bg: C.sageLight,
      items: ["Результат и победа", "Конкуренция", "Достижение целей", "Репутация лидера"] },
  ];

  quadrants.forEach(q => {
    rect(s, q.x, q.y, 4.55, 1.82, q.bg, q.color, 0.8);
    rect(s, q.x, q.y, 4.55, 0.44, q.color, q.color, 0);
    txt(s, q.name, q.x + 0.15, q.y + 0.04, 2.5, 0.36, { fontSize: 14, bold: true, color: "FFFFFF", align: "left" });
    txt(s, q.sub, q.x + 2.7, q.y + 0.06, 1.7, 0.32, { fontSize: 10, color: "FFFFFF", italic: true, align: "right" });
    q.items.forEach((item, j) => {
      txt(s, [{ text: "◆  ", options: { color: q.color, bold: true } }, { text: item, options: { color: C.textDark } }], {
        x: q.x + 0.18, y: q.y + 0.52 + j * 0.31, w: 4.2, h: 0.28,
        fontSize: 11.5, align: "left"
      });
    });
  });

  // Center cross lines
  rect(s, 4.75, 1.28, 0.06, 3.82, C.border, C.border, 0);
  rect(s, 0.38, 3.18, 9.24, 0.06, C.border, C.border, 0);
  oval(s, 4.68, 3.1, 0.22, 0.22, C.textLight);
}

// ─── SLIDE 6: Schein's model ──────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "УРОВНИ КОРПОРАТИВНОЙ КУЛЬТУРЫ ПО Э. ШЕЙНУ");

  // Pyramid layers (trapezoids via wide→narrow rects)
  const layers = [
    { label: "УРОВЕНЬ 1", name: "Артефакты", sub: "Видимые и осязаемые элементы", w: 9.24, x: 0.38, y: 1.05, h: 0.88, color: C.teal, bg: C.tealLight,
      examples: "Офис и дизайн  ·  Дресс-код  ·  Логотип  ·  Ритуалы  ·  Язык общения" },
    { label: "УРОВЕНЬ 2", name: "Провозглашаемые ценности", sub: "Декларируемые нормы и стратегии", w: 7.2, x: 1.02, y: 2.05, h: 0.88, color: C.gold, bg: C.goldLight,
      examples: "Миссия и vision  ·  Корпоративный кодекс  ·  Лозунги  ·  Цели компании" },
    { label: "УРОВЕНЬ 3", name: "Базовые предположения", sub: "Глубинные убеждения и установки", w: 5.1, x: 2.07, y: 3.05, h: 0.88, color: C.coral, bg: C.coralLight,
      examples: "Отношение ко времени  ·  Природа человека  ·  Стиль власти  ·  Неосознанные паттерны" },
  ];

  layers.forEach(l => {
    rect(s, l.x, l.y, l.w, l.h, l.bg, l.color, 0.8);
    rect(s, l.x, l.y, l.w, 0.4, l.color, l.color, 0);
    txt(s, l.label, l.x + 0.15, l.y + 0.04, 1.4, 0.32, { fontSize: 10, color: "FFFFFF", align: "left" });
    txt(s, l.name, l.x + 1.6, l.y + 0.04, l.w - 1.75, 0.32, { fontSize: 14, bold: true, color: "FFFFFF", align: "center" });
    txt(s, l.sub, l.x + 0.15, l.y + 0.46, l.w - 0.3, 0.22, { fontSize: 11, color: l.color, italic: true, align: "center" });
    txt(s, l.examples, l.x + 0.15, l.y + 0.62, l.w - 0.3, 0.22, { fontSize: 10.5, color: C.textMid, align: "center" });
  });

  // Arrows showing depth
  txt(s, "↑  Легко наблюдать", 0.38, 1.05, 1.5, 0.88, { fontSize: 10, color: C.teal, align: "center", valign: "middle", rotate: 90 });

  // Key insight box
  rect(s, 0.38, 4.15, 9.24, 1.28, C.bgDark, C.bgDark, 0);
  rect(s, 0.38, 4.15, 0.18, 1.28, C.gold, C.gold, 0);
  txt(s, "Ключевая идея Шейна:", 0.7, 4.22, 2.5, 0.35, { fontSize: 12, bold: true, color: C.gold, align: "left" });
  txt(s, "Видимые артефакты — лишь «верхушка айсберга». Настоящая культура скрыта в базовых предположениях, которые сотрудники не осознают, но которые управляют их поведением каждый день.", {
    x: 0.7, y: 4.6, w: 8.7, h: 0.72,
    fontSize: 12, color: "C8D8D4", align: "left", italic: true
  });
}

// ─── SLIDE 7: Functions ───────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ФУНКЦИИ КОРПОРАТИВНОЙ КУЛЬТУРЫ");

  const functions = [
    { icon: "🧭", title: "Ориентирующая", desc: "Задаёт систему координат: что хорошо, что плохо, как принимать решения в нестандартных ситуациях без обращения к руководству.", color: C.teal, bg: C.tealLight },
    { icon: "🤝", title: "Интегрирующая", desc: "Создаёт ощущение «мы». Объединяет сотрудников вокруг общих ценностей, снижает конфликтность и формирует командный дух.", color: C.gold, bg: C.goldLight },
    { icon: "🛡️", title: "Охранная", desc: "Защищает организацию от нежелательных внешних влияний. Формирует иммунитет к разрушительным практикам и конкурентам.", color: C.coral, bg: C.coralLight },
    { icon: "⚡", title: "Мотивирующая", desc: "Сильная культура сама по себе мотивирует: сотрудник разделяет смысл своей работы и чувствует принадлежность к чему-то большему.", color: C.sage, bg: C.sageLight },
    { icon: "📡", title: "Коммуникативная", desc: "Формирует общий язык, нормы общения, стиль взаимодействия внутри и вне компании — с клиентами, партнёрами.", color: C.teal, bg: C.tealLight },
    { icon: "🔄", title: "Адаптивная", desc: "Помогает организации и сотрудникам адаптироваться к изменениям: гибкие культуры легче переживают кризисы и трансформации.", color: C.gold, bg: C.goldLight },
  ];

  functions.forEach((fn, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.38 + col * 3.12;
    const y = 1.05 + row * 2.22;

    rect(s, x, y, 2.94, 2.02, fn.bg, fn.color, 0.8);
    rect(s, x, y, 2.94, 0.42, fn.color, fn.color, 0);
    txt(s, fn.icon + "  " + fn.title, x + 0.12, y + 0.04, 2.7, 0.34, { fontSize: 12.5, bold: true, color: "FFFFFF", align: "left" });
    txt(s, fn.desc, x + 0.15, y + 0.52, 2.64, 1.42, { fontSize: 11, color: C.textDark, align: "left" });
  });
}

// ─── SLIDE 8: Influence on staff & motivation ────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ВЛИЯНИЕ НА ПЕРСОНАЛ И МОТИВАЦИЮ");

  // Stats row
  const stats = [
    { val: "+30%", label: "вовлечённость\nв сильной культуре", color: C.teal },
    { val: "−40%", label: "текучесть кадров\nпри ценностном совпадении", color: C.gold },
    { val: "×3",   label: "выше инициатива\nу сотрудников-«носителей»", color: C.coral },
    { val: "72%",  label: "сотрудников готовы\nменять ради культуры зарплату", color: C.sage },
  ];

  stats.forEach((st, i) => {
    const x = 0.38 + i * 2.32;
    rect(s, x, 1.05, 2.14, 1.28, i % 2 === 0 ? C.tealLight : C.goldLight, st.color, 0.8);
    txt(s, st.val, x, 1.1, 2.14, 0.68, { fontSize: 32, bold: true, color: st.color, align: "center" });
    txt(s, st.label, x + 0.1, 1.78, 1.94, 0.48, { fontSize: 10.5, color: C.textMid, align: "center" });
  });

  // Two columns: positive / negative
  rect(s, 0.38, 2.52, 4.55, 0.38, C.teal, C.teal, 0);
  txt(s, "Сильная культура → мотивирует", 0.38, 2.54, 4.55, 0.34, { fontSize: 12.5, bold: true, color: "FFFFFF", align: "center" });

  rect(s, 5.07, 2.52, 4.55, 0.38, C.coral, C.coral, 0);
  txt(s, "Слабая культура → демотивирует", 5.07, 2.54, 4.55, 0.34, { fontSize: 12.5, bold: true, color: "FFFFFF", align: "center" });

  const posItems = [
    "Сотрудники понимают смысл работы",
    "Высокая психологическая безопасность",
    "Самостоятельность и ответственность",
    "Гордость за принадлежность к компании",
    "Органичная взаимопомощь в команде",
  ];
  const negItems = [
    "Непонимание целей и приоритетов",
    "Страх ошибки и публичной критики",
    "Высокий уровень стресса и выгорания",
    "Разрыв между словами и делами",
    "«Токсичные» неформальные нормы",
  ];

  posItems.forEach((item, i) => {
    rect(s, 0.38, 3.0 + i * 0.47, 4.55, 0.42, i % 2 === 0 ? C.tealLight : C.bg, C.border, 0);
    txt(s, [{ text: "✓  ", options: { color: C.teal, bold: true } }, { text: item, options: { color: C.textDark } }], {
      x: 0.55, y: 3.03 + i * 0.47, w: 4.22, h: 0.36, fontSize: 11.5, align: "left", valign: "middle"
    });
  });

  negItems.forEach((item, i) => {
    rect(s, 5.07, 3.0 + i * 0.47, 4.55, 0.42, i % 2 === 0 ? C.coralLight : C.bg, C.border, 0);
    txt(s, [{ text: "✗  ", options: { color: C.coral, bold: true } }, { text: item, options: { color: C.textDark } }], {
      x: 5.22, y: 3.03 + i * 0.47, w: 4.22, h: 0.36, fontSize: 11.5, align: "left", valign: "middle"
    });
  });
}

// ─── SLIDE 9: Formation & Management ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ФОРМИРОВАНИЕ И УПРАВЛЕНИЕ КУЛЬТУРОЙ");

  // Left: who forms it
  txt(s, "Кто формирует культуру", 0.38, 1.05, 4.4, 0.36, { fontSize: 14, bold: true, color: C.teal, align: "left" });

  const actors = [
    { role: "Основатели", desc: "Закладывают базовые ценности и «ДНК» компании через личный пример и ключевые решения на старте", color: C.teal },
    { role: "Топ-менеджмент", desc: "Транслируют и поддерживают культуру через стратегию, риторику, символические действия и решения о людях", color: C.gold },
    { role: "HR-отдел", desc: "Операционализирует культуру: найм по ценностям, онбординг, обучение, ритуалы признания", color: C.coral },
    { role: "Сотрудники", desc: "Со-создатели и хранители культуры — неформальные лидеры влияют на неё сильнее, чем регламенты", color: C.sage },
  ];

  actors.forEach((a, i) => {
    const y = 1.52 + i * 0.98;
    rect(s, 0.38, y, 4.4, 0.86, i % 2 === 0 ? C.tealLight : C.bg, a.color, 0.6);
    oval(s, 0.48, y + 0.18, 0.52, 0.52, a.color);
    txt(s, (i + 1).toString(), 0.48, y + 0.18, 0.52, 0.52, { fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, a.role, 1.12, y + 0.08, 3.5, 0.3, { fontSize: 13, bold: true, color: a.color, align: "left" });
    txt(s, a.desc, 1.12, y + 0.42, 3.5, 0.4, { fontSize: 10.5, color: C.textMid, align: "left" });
  });

  // Right: HR tools
  txt(s, "Инструменты HR для культуры", 5.05, 1.05, 4.55, 0.36, { fontSize: 14, bold: true, color: C.teal, align: "left" });

  const tools = [
    { icon: "🔍", name: "Найм по ценностям", desc: "Культурно-совместимые кандидаты" },
    { icon: "🚀", name: "Онбординг", desc: "Погружение в культуру с первых дней" },
    { icon: "🏆", name: "Система признания", desc: "Публичное поощрение нужного поведения" },
    { icon: "📚", name: "Обучение и развитие", desc: "Трансляция ценностей через знания" },
    { icon: "🎉", name: "Ритуалы и традиции", desc: "Корпоративы, церемонии, события" },
    { icon: "📊", name: "Оценка и обратная связь", desc: "360°, pulse-surveys, eNPS" },
  ];

  tools.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 5.05 + col * 2.28;
    const y = 1.52 + row * 1.38;
    rect(s, x, y, 2.1, 1.24, C.bg, C.border, 0.6);
    txt(s, t.icon, x + 0.15, y + 0.1, 0.42, 0.42, { fontSize: 20, align: "center" });
    txt(s, t.name, x + 0.65, y + 0.08, 1.35, 0.3, { fontSize: 11.5, bold: true, color: C.teal, align: "left" });
    txt(s, t.desc, x + 0.65, y + 0.4, 1.35, 0.72, { fontSize: 10.5, color: C.textMid, align: "left" });
  });
}

// ─── SLIDE 10: Diagnostics ────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ДИАГНОСТИКА И ОЦЕНКА КОРПОРАТИВНОЙ КУЛЬТУРЫ");

  // Methods table
  const methods = [
    { name: "OCAI (Камерон–Куинн)", type: "Анкетирование", what: "Определяет тип культуры по 4 квадрантам. Выявляет разрыв между текущей и желаемой культурой", color: C.teal, bg: C.tealLight },
    { name: "Денисонский опросник", type: "Анкетирование", what: "Измеряет 4 параметра: вовлечённость, согласованность, адаптивность, миссию. Связывает культуру с финансовыми результатами", color: C.gold, bg: C.goldLight },
    { name: "eNPS / Pulse-опросы", type: "Регулярный мониторинг", what: "Быстрые короткие опросы для отслеживания динамики. Показывают готовность рекомендовать компанию как работодателя", color: C.coral, bg: C.coralLight },
    { name: "Глубинные интервью", type: "Качественный метод", what: "1:1 с сотрудниками разных уровней. Выявляют неформальные нормы, «теневую» культуру и болевые точки", color: C.sage, bg: C.sageLight },
    { name: "Наблюдение и анализ", type: "Полевое исследование", what: "Наблюдение за встречами, анализ коммуникаций, изучение артефактов (офис, переписка, решения)", color: C.teal, bg: C.tealLight },
  ];

  methods.forEach((m, i) => {
    const y = 1.05 + i * 0.9;
    rect(s, 0.38, y, 9.24, 0.78, m.bg, m.color, 0.6);
    rect(s, 0.38, y, 0.14, 0.78, m.color, m.color, 0);
    txt(s, m.name, 0.65, y + 0.08, 2.8, 0.3, { fontSize: 13, bold: true, color: m.color, align: "left" });
    txt(s, m.type, 0.65, y + 0.44, 2.8, 0.26, { fontSize: 10.5, color: C.textLight, italic: true, align: "left" });
    txt(s, m.what, 3.6, y + 0.08, 5.8, 0.62, { fontSize: 11.5, color: C.textDark, align: "left", valign: "middle" });
  });

  // Bottom note
  rect(s, 0.38, 5.58, 9.24, 0.0, C.tealLight, C.tealLight, 0);
}

// ─── SLIDE 11: Transformation ────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "ТРАНСФОРМАЦИЯ И ИЗМЕНЕНИЕ КОРПОРАТИВНОЙ КУЛЬТУРЫ");

  // Quote
  rect(s, 0.38, 1.05, 9.24, 0.82, C.bgDark, C.bgDark, 0);
  rect(s, 0.38, 1.05, 0.18, 0.82, C.gold, C.gold, 0);
  txt(s, "«Культуру нельзя изменить приказом. Но её можно трансформировать — через поведение лидеров, символические решения и терпение»", {
    x: 0.7, y: 1.1, w: 8.7, h: 0.72,
    fontSize: 12.5, color: "FFFFFF", italic: true, align: "left", valign: "middle"
  });

  // 5-step roadmap
  txt(s, "Дорожная карта трансформации культуры", 0.38, 2.05, 9.24, 0.35, { fontSize: 14, bold: true, color: C.textDark, align: "left" });

  const steps = [
    { n: "1", title: "Диагностика", desc: "Где мы\nсейчас?", color: C.teal },
    { n: "2", title: "Видение",     desc: "Где хотим\nбыть?",  color: C.gold },
    { n: "3", title: "Gap-анализ",  desc: "Что мешает\nдвигаться?", color: C.coral },
    { n: "4", title: "Внедрение",   desc: "Пилоты,\nритуалы, лидеры", color: C.sage },
    { n: "5", title: "Мониторинг",  desc: "Метрики\nи корректировка", color: C.teal },
  ];

  steps.forEach((st, i) => {
    const x = 0.38 + i * 1.88;
    rect(s, x, 2.52, 1.72, 1.72, i % 2 === 0 ? C.tealLight : C.bg, st.color, 0.8);
    oval(s, x + 0.6, 2.62, 0.52, 0.52, st.color);
    txt(s, st.n, x + 0.6, 2.62, 0.52, 0.52, { fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, st.title, x + 0.12, 3.25, 1.48, 0.32, { fontSize: 12, bold: true, color: st.color, align: "center" });
    txt(s, st.desc, x + 0.12, 3.6, 1.48, 0.58, { fontSize: 10.5, color: C.textMid, align: "center" });
    if (i < 4) {
      txt(s, "→", x + 1.72, 3.1, 0.16, 0.36, { fontSize: 16, color: C.textLight, align: "center", valign: "middle" });
    }
  });

  // Barriers & success factors
  const left = ["Сопротивление сотрудников", "Формальное внедрение без смысла", "Несоответствие поведения лидеров", "Нетерпение (культура меняется 3–7 лет)"];
  const right = ["Личный пример топ-менеджмента", "Вовлечение сотрудников в процесс", "Символические «быстрые победы»", "Регулярное измерение и признание"];

  rect(s, 0.38, 4.42, 4.55, 0.36, C.coral, C.coral, 0);
  txt(s, "⚠  Барьеры трансформации", 0.38, 4.44, 4.55, 0.32, { fontSize: 12.5, bold: true, color: "FFFFFF", align: "center" });

  rect(s, 5.07, 4.42, 4.55, 0.36, C.teal, C.teal, 0);
  txt(s, "✓  Факторы успеха", 5.07, 4.44, 4.55, 0.32, { fontSize: 12.5, bold: true, color: "FFFFFF", align: "center" });

  left.forEach((item, i) => {
    rect(s, 0.38, 4.88 + i * 0.18, 4.55, 0.17, i % 2 === 0 ? C.coralLight : C.bg, C.border, 0);
    txt(s, "▸  " + item, 0.55, 4.9 + i * 0.18, 4.22, 0.14, { fontSize: 10.5, color: C.textDark, align: "left" });
  });

  right.forEach((item, i) => {
    rect(s, 5.07, 4.88 + i * 0.18, 4.55, 0.17, i % 2 === 0 ? C.tealLight : C.bg, C.border, 0);
    txt(s, "▸  " + item, 5.22, 4.9 + i * 0.18, 4.22, 0.14, { fontSize: 10.5, color: C.textDark, align: "left" });
  });
}

// ─── SLIDE 12: Conclusions ────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "КЛЮЧЕВЫЕ ВЫВОДЫ");

  const conclusions = [
    { num: "1", text: "Корпоративная культура — нематериальный актив, оказывающий прямое влияние на эффективность, мотивацию и удержание персонала.", color: C.teal },
    { num: "2", text: "Культура формируется на трёх уровнях (по Шейну): видимые артефакты, декларируемые ценности и неосознанные базовые убеждения.", color: C.gold },
    { num: "3", text: "Сильная культура повышает вовлечённость на 30%, снижает текучесть на 40% и служит конкурентным преимуществом компании.", color: C.coral },
    { num: "4", text: "HR-специалисты являются архитекторами культуры: через найм, онбординг, признание и обучение они транслируют и поддерживают ценности.", color: C.sage },
    { num: "5", text: "Трансформация культуры — долгосрочный процесс (3–7 лет), требующий личного примера лидеров, системного подхода и регулярного измерения.", color: C.teal },
  ];

  conclusions.forEach((c, i) => {
    const rowH = 0.75;
    const circleSize = 0.52;
    const y = 1.05 + i * 0.9;
    const circleY = y + (rowH - circleSize) / 2;
    rect(s, 0.38, y, 9.24, rowH, i % 2 === 0 ? C.tealLight : C.bg, C.border, 0.6);
    oval(s, 0.5, circleY, circleSize, circleSize, c.color);
    txt(s, c.num, 0.5, circleY, circleSize, circleSize, { fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, c.text, 1.18, y + 0.05, 8.3, rowH - 0.1, { fontSize: 12.5, color: C.textDark, align: "left", valign: "middle" });
  });
}

// ─── SLIDE 13: Final ──────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  rect(s, 0, 0, 0.22, 5.63, C.gold);
  rect(s, 0, 5.38, 10, 0.25, "162E29");
  oval(s, 7.0, -0.6, 3.2, 3.2, "243F3A");
  oval(s, 6.5, 3.9, 2.0, 2.0, "1A302C");

  txt(s, "СПАСИБО", { x: 0.55, y: 0.95, w: 9, h: 1.1, fontSize: 54, bold: true, color: "FFFFFF", align: "left", charSpacing: 4 });
  txt(s, "ЗА ВНИМАНИЕ", { x: 0.55, y: 1.95, w: 9, h: 0.85, fontSize: 36, bold: false, color: C.gold, align: "left", charSpacing: 3 });

  rect(s, 0.55, 2.95, 5.5, 0.05, C.tealMid);

  txt(s, "Корпоративная культура", { x: 0.55, y: 3.15, w: 8, h: 0.42, fontSize: 16, color: "8BBFB8", align: "left" });
  txt(s, "Управление человеческими ресурсами", { x: 0.55, y: 3.6, w: 8, h: 0.38, fontSize: 13, color: "5A8882", italic: true, align: "left" });

  rect(s, 0.55, 4.3, 5.2, 0.75, C.teal, C.teal, 0);
  txt(s, "Готов ответить на ваши вопросы", { x: 0.55, y: 4.3, w: 5.2, h: 0.75, fontSize: 15, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
}

// ─── Write ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/mnt/user-data/outputs/Корпоративная_культура_УЧР.pptx" })
  .then(() => console.log("Done!"))
  .catch(err => console.error(err));
