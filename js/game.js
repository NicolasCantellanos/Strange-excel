/* ============================================================
   STRANGE EXCEL — Vallenoche, 1986
   Motor de misiones con progresión de menos a más,
   basada en los módulos del curso de Excel:
     Ep1: celdas y fórmulas básicas, SUMA        (Módulo 1)
     Ep2: PROMEDIO, MAX, MIN                     (Módulo 1/8)
     Ep3: CONTAR.SI, SUMAR.SI (condicionales)    (Módulo 2/8)
     Futuro: BUSCARV (M4), tablas dinámicas (M5), Power Query/BI (M10)
   ============================================================ */

/* ---------- 0. Funciones de Excel en español ----------
   La librería solo trae los nombres en inglés. Alias en español: */

function _flat(args) {
  const out = [];
  const walk = (v) => { if (Array.isArray(v)) v.forEach(walk); else out.push(v); };
  Array.prototype.slice.call(args).forEach(walk);
  return out;
}

function _nums(args) {
  return _flat(args).map(parseFloat).filter((n) => !isNaN(n));
}

function _matches(value, crit) {
  const c = String(crit).trim();
  const op = c.match(/^(>=|<=|<>|>|<)(.*)$/);
  if (op) {
    const n = parseFloat(value);
    const target = parseFloat(op[2]);
    if (isNaN(n) || isNaN(target)) return false;
    switch (op[1]) {
      case ">": return n > target;
      case "<": return n < target;
      case ">=": return n >= target;
      case "<=": return n <= target;
      case "<>": return n !== target;
    }
  }
  return String(value).trim().toLowerCase() === c.toLowerCase();
}

window.SUMA = function () {
  return _nums(arguments).reduce((a, b) => a + b, 0);
};

window.PROMEDIO = function () {
  const v = _nums(arguments);
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0;
};

/* MAX/MIN nativas de la librería devuelven 0 con autoCasting:false
   (ver nota más abajo) — las reemplazamos por versiones propias. */
window.MAX = function () {
  const v = _nums(arguments);
  return v.length ? Math.max(...v) : 0;
};
window.MIN = function () {
  const v = _nums(arguments);
  return v.length ? Math.min(...v) : 0;
};

/* Funciones lógicas y de texto (Módulo 2). Nota técnica: la hoja se crea
   con autoCasting:false (ver startPuzzle/startTutorial) — sin eso, la
   librería corrompe cualquier texto que termine en números (p.ej.
   "CAS0012" llegaba como el número 12) al pasarlo a una función. */

window.SI = function (cond, siTrue, siFalse) { return cond ? siTrue : siFalse; };
window.SI.ERROR = function (valor, alternativa) {
  const s = String(valor);
  if (s.startsWith("#") || s === "NaN" || s === "Infinity" || s === "-Infinity" || s === "undefined") return alternativa;
  return valor;
};
window.Y = function () { return Array.prototype.slice.call(arguments).every(Boolean); };
window.O = function () { return Array.prototype.slice.call(arguments).some(Boolean); };

window.MAYUSC = (v) => String(v).toUpperCase();
window.MINUSC = (v) => String(v).toLowerCase();
window.ESPACIOS = (v) => String(v).trim().replace(/\s+/g, " ");
window.IZQUIERDA = (v, n) => String(v).slice(0, parseInt(n, 10));
window.DERECHA = (v, n) => { const s = String(v); return s.slice(Math.max(0, s.length - parseInt(n, 10))); };
window.CONCATENAR = function () { return Array.prototype.slice.call(arguments).join(""); };

/* Fechas en formato AAAA-MM-DD (evitamos Date() por líos de huso horario) */
function _fecha(v) {
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? { y: parseInt(m[1], 10), mo: parseInt(m[2], 10), d: parseInt(m[3], 10) } : null;
}
window.AÑO = (v) => { const f = _fecha(v); return f ? f.y : "#VALOR!"; };
window.MES = (v) => { const f = _fecha(v); return f ? f.mo : "#VALOR!"; };
window.DIA = (v) => { const f = _fecha(v); return f ? f.d : "#VALOR!"; };

/* BUSCARX: como BUSCARV pero sin restricción de dirección — puede buscar
   en una columna y devolver otra columna A SU IZQUIERDA (algo que BUSCARV
   no puede hacer). Cada misión que la usa define mission.lookupX con los
   offsets de columna (dentro de mission.lookup) a buscar y a devolver. */
window.BUSCARX = function (valorBuscado) {
  if (!spreadsheet || !mission || !mission.lookup || !mission.lookupX) return "#N/D";
  const L = mission.lookup, X = mission.lookupX;
  const valor = String(valorBuscado).trim().toLowerCase();
  for (let r = 0; r < L.rows; r++) {
    const key = String(spreadsheet.getValueFromCoords(L.x0 + X.searchCol, L.y0 + r, true) || "").trim().toLowerCase();
    if (key === valor) {
      return spreadsheet.getValueFromCoords(L.x0 + X.returnCol, L.y0 + r, true);
    }
  }
  return "#N/D";
};

/* CONTAR.SI y SUMAR.SI: el punto funciona como acceso a propiedad.
   Ojo: el parser de la librería a veces desparrama los rangos como
   argumentos sueltos (B1, B2, ..., B8, "Cassette") en vez de pasar
   una lista — soportamos los dos formatos. */
window.CONTAR = {
  SI: function () {
    const args = Array.prototype.slice.call(arguments);
    let range, crit;
    if (Array.isArray(args[0])) {
      range = _flat([args[0]]);
      crit = args[1];
    } else {
      crit = args[args.length - 1];
      range = args.slice(0, -1);
    }
    return range.filter((v) => _matches(v, crit)).length;
  },
};

window.SUMAR = {
  SI: function () {
    const args = Array.prototype.slice.call(arguments);
    let vals, crit, sums;
    if (Array.isArray(args[0])) {
      vals = _flat([args[0]]);
      crit = args[1];
      sums = args[2] !== undefined ? _flat([args[2]]) : vals;
    } else if ((args.length - 1) % 2 === 0 && args.length >= 3) {
      const n = (args.length - 1) / 2;
      vals = args.slice(0, n);
      crit = args[n];
      sums = args.slice(n + 1);
    } else {
      crit = args[args.length - 1];
      vals = args.slice(0, -1);
      sums = vals;
    }
    let total = 0;
    for (let i = 0; i < vals.length; i++) {
      if (_matches(vals[i], crit)) {
        const n = parseFloat(sums[i]);
        if (!isNaN(n)) total += n;
      }
    }
    return total;
  },
};

/* BUSCARV: el parser de la librería rompe los textos al pasar tablas 2D,
   así que la función lee la tabla directamente de la hoja de la misión
   (region definida en mission.lookup). El jugador escribe la sintaxis
   real de Excel: =BUSCARV("código", A1:C8, columna) */
window.FALSO = false;
window.VERDADERO = true;

window.BUSCARV = function () {
  const args = Array.prototype.slice.call(arguments);
  if (!spreadsheet || !mission || !mission.lookup) return "#N/D";
  const last = args[args.length - 1];
  let colIdx = typeof last === "boolean" ? args[args.length - 2] : last;
  colIdx = parseInt(colIdx, 10);
  const valor = String(args[0]).trim().toLowerCase();
  const L = mission.lookup;
  if (isNaN(colIdx) || colIdx < 1 || colIdx > L.cols) return "#N/D";
  for (let r = 0; r < L.rows; r++) {
    const key = String(spreadsheet.getValueFromCoords(L.x0, L.y0 + r, true) || "").trim().toLowerCase();
    if (key === valor) {
      return spreadsheet.getValueFromCoords(L.x0 + colIdx - 1, L.y0 + r, true);
    }
  }
  return "#N/D";
};

/* ---------- 1. Arte pixel ---------- */

function pixelSVG(rows, palette, cell) {
  const w = rows[0].length;
  const h = rows.length;
  let rects = "";
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const ch = rows[y][x];
      if (ch === ".") continue;
      rects += `<rect x="${x * cell}" y="${y * cell}" width="${cell}" height="${cell}" fill="${palette[ch]}"/>`;
    }
  }
  return `<svg viewBox="0 0 ${w * cell} ${h * cell}" width="${w * 8}" height="${h * 8}" xmlns="http://www.w3.org/2000/svg">${rects}</svg>`;
}

const MARTA_ROWS = [
  "................",
  "....HHHHHHHH....",
  "...HHHHHHHHHH...",
  "..HHHHHHHHHHHH..",
  "..HHSSSSSSSSHH..",
  "..HSSSSSSSSSSH..",
  "..HSSEESSEESSH..",
  "..HSSSSSSSSSSH..",
  "..HSSSOOOOSSSH..",
  "..HHSSSSSSSSHH..",
  "...HHSSSSSSHH...",
  "....HHSSSSHH....",
  "...RRRRRRRRRR...",
  "..RRRRRRRRRRRR..",
  ".RRRRRRRRRRRRRR.",
  "RRRRRRRRRRRRRRRR",
  "RRRRCCCCCCCCRRRR",
  "RRRRRRRRRRRRRRRR",
  "RRRRRRRRRRRRRRRR",
  "................",
];

/* Variante "preocupada": mismas filas, cejas fruncidas agregadas a la altura de los ojos */
const MARTA_ROWS_WORRIED = MARTA_ROWS.map((row, i) =>
  i === 5 ? "..HSSBBSSBBSSH.." : row
);

const MARTA_PALETTE = {
  H: "#3a2350", S: "#e8c39e", E: "#141414",
  O: "#7a4a2c", R: "#b5293b", C: "#7c1a28", B: "#8a1f2f",
};

function showPortrait(mood) {
  const rows = mood === "worried" ? MARTA_ROWS_WORRIED : MARTA_ROWS;
  document.getElementById("portrait").innerHTML = pixelSVG(rows, MARTA_PALETTE, 10);
}
function hidePortrait() {
  document.getElementById("portrait").innerHTML = "";
}

/* ---------- 1b. Escenarios de fondo (generados por código) ---------- */

function svgWrap(w, h, inner, cls) {
  return `<svg class="${cls || ""}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="xMidYMax slice" xmlns="http://www.w3.org/2000/svg" shape-rendering="crispEdges">${inner}</svg>`;
}

function buildTownBackdrop() {
  const w = 480, h = 160;
  const buildings = [
    { x: 0, w: 60, h: 70 }, { x: 55, w: 40, h: 100 }, { x: 90, w: 70, h: 60 },
    { x: 155, w: 50, h: 90 }, { x: 200, w: 80, h: 50 }, { x: 275, w: 45, h: 110 },
    { x: 315, w: 60, h: 75 }, { x: 370, w: 50, h: 95 }, { x: 415, w: 70, h: 65 },
  ];
  let s = buildings.map((b) => `<rect x="${b.x}" y="${h - b.h}" width="${b.w}" height="${b.h}" fill="#150a22"/>`).join("");
  /* la ventana iluminada de la tienda de discos, en el segundo edificio */
  s += `<rect x="70" y="${h - 70}" width="10" height="12" fill="#ffcc4d">
      <animate attributeName="opacity" values="1;0.6;1" dur="4s" repeatCount="indefinite"/>
    </rect>`;
  return svgWrap(w, h, s, "backdrop-town");
}

function buildShopBackdrop() {
  const w = 480, h = 160;
  const colors = ["#b5293b", "#37f0d8", "#ffcc4d", "#7c1a28", "#3a2350"];
  let s = `<rect x="0" y="0" width="${w}" height="${h}" fill="#2a1408"/>`;
  [40, 75, 110].forEach((y) => { s += `<rect x="20" y="${y}" width="440" height="6" fill="#5c2f12"/>`; });
  [24, 59, 94].forEach((y, row) => {
    for (let i = 0; i < 18; i++) {
      s += `<rect x="${24 + i * 24}" y="${y}" width="14" height="14" fill="${colors[(i + row) % colors.length]}"/>`;
    }
  });
  s += `<rect x="0" y="${h - 32}" width="${w}" height="4" fill="#5c2f12"/>`;
  s += `<rect x="0" y="${h - 28}" width="${w}" height="28" fill="#3a1c0c"/>`;
  return svgWrap(w, h, s, "backdrop-shop");
}

function buildBasementBackdrop() {
  const w = 480, h = 160;
  const cx = w / 2;
  let s = `<rect x="0" y="0" width="${w}" height="${h}" fill="#0a0612"/>`;
  s += `<rect x="0" y="20" width="${w}" height="6" fill="#1a1024"/>`;
  s += `<rect x="0" y="130" width="${w}" height="6" fill="#1a1024"/>`;
  [60, 180, 300, 420].forEach((x) => { s += `<rect x="${x}" y="20" width="8" height="116" fill="#241834"/>`; });
  s += `<rect x="${cx - 1}" y="26" width="2" height="30" fill="#3a2350"/>`;
  s += `<circle cx="${cx}" cy="60" r="24" fill="#ffcc4d" opacity="0.15">
      <animate attributeName="opacity" values="0.15;0.02;0.15;0.15;0.05;0.15" dur="2.6s" repeatCount="indefinite"/>
    </circle>`;
  s += `<circle cx="${cx}" cy="60" r="10" fill="#ffcc4d">
      <animate attributeName="opacity" values="1;0.15;1;1;0.3;1" dur="2.6s" repeatCount="indefinite"/>
    </circle>`;
  return svgWrap(w, h, s, "backdrop-basement");
}

const BACKDROPS = { town: buildTownBackdrop, shop: buildShopBackdrop, basement: buildBasementBackdrop };

function setBackdrop(name) {
  const el2 = document.getElementById("scene-backdrop");
  el2.innerHTML = name && BACKDROPS[name] ? BACKDROPS[name]() : "";
}

/* ---------- 2. Progreso ---------- */

const SAVE_KEY = "strangeExcelProgress";

function loadProgress() {
  try { return JSON.parse(localStorage.getItem(SAVE_KEY)) || { completed: [] }; }
  catch { return { completed: [] }; }
}
function saveProgress(p) {
  try { localStorage.setItem(SAVE_KEY, JSON.stringify(p)); } catch {}
}
function isDone(id) { return loadProgress().completed.includes(id); }
function markDone(id) {
  const p = loadProgress();
  if (!p.completed.includes(id)) { p.completed.push(id); saveProgress(p); }
}

/* ---------- 3. Misiones ---------- */

const MISSIONS = [
  {
    id: 1,
    name: "EP.1 — EL REGISTRO QUE NO CIERRA",
    concepts: "Celdas · fórmulas · multiplicar · SUMA",
    offerTutorial: true,
    intro: [
      { speaker: "", text: "Vallenoche. Un pueblo tranquilo... casi siempre.", portrait: false, bg: "town" },
      { speaker: "", text: "En la trastienda de «Discos Cassette Vallenoche» hay una computadora vieja que nadie enciende de día.", portrait: false, bg: "shop" },
      { speaker: "MARTA", text: "¡Por fin llegás! Necesito ayuda con esto antes de que cierre la tienda.", portrait: true },
      { speaker: "MARTA", text: "Cada mes hago las cuentas a mano... pero esta vez algo no cierra. Y cuando el total no cierra, la luz del sótano empieza a parpadear.", portrait: true, mood: "worried", bg: "basement" },
    ],
    briefing: [
      { speaker: "MARTA", text: "Acá está la planilla de ventas. Cada FILA es un día — un registro. Cada COLUMNA es un dato de esa venta: día, producto, cantidad, precio.", portrait: true },
      { speaker: "MARTA", text: "Para cada día necesito el Total vendido. Pensá: ¿qué operación conecta la Cantidad con el Precio Unitario?", portrait: true },
      { speaker: "MARTA", text: "Si te trabás, apretá el botón 💡 PISTA. Para eso está.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "¡Cerró perfecto! Mirá — la luz del sótano dejó de parpadear.", portrait: true, bg: "basement" },
      { speaker: "MARTA", text: "Pero esto era solo el principio. Volvé mañana... los números de este pueblo esconden más secretos.", portrait: true },
    ],
    title: "EL REGISTRO QUE NO CIERRA",
    sheet: {
      data: [
        ["Lunes", "Cassette - Blondie", 3, 800, ""],
        ["Martes", "Vinilo - Bowie", 2, 1500, ""],
        ["Miércoles", "Cassette - Queen", 5, 800, ""],
        ["Jueves", "Vinilo - Talking Heads", 1, 2000, ""],
        ["Viernes", "Cassette - Bowie", 4, 800, ""],
        ["", "TOTAL GENERAL", "", "", ""],
      ],
      columns: [
        { type: "text", title: "Día", width: 100, readOnly: true },
        { type: "text", title: "Producto", width: 190, readOnly: true },
        { type: "numeric", title: "Cantidad", width: 100, readOnly: true },
        { type: "numeric", title: "Precio Unit.", width: 110, readOnly: true },
        { type: "numeric", title: "Total", width: 120 },
      ],
    },
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          Completá la columna <b>Total</b> para cada día (fila). El Total de cada fila depende de dos columnas de esa misma fila.<br>
          <span style="color:var(--text-dim);">Doble clic en una celda de Total para escribir. Toda fórmula empieza con <b>=</b> — y mientras la escribís, podés hacer clic en otras celdas para agregarlas. ¿Trabado? Botón 💡 PISTA.</span>`,
        hints: [
          "Pensá en la fila del Lunes: se vendieron 3 cassettes a $800 cada uno. ¿Qué cuenta harías con la calculadora?",
          "Hay que MULTIPLICAR la Cantidad por el Precio Unitario. En Excel, el símbolo de multiplicar es el asterisco (*). Y toda fórmula empieza con =",
          "Hacé doble clic en el Total del Lunes (columna E, fila 1) y escribí: =C1*D1 — después repetí el patrón en cada fila (=C2*D2, =C3*D3...).",
        ],
        check(s) {
          const expected = [2400, 3000, 4000, 2000, 3200];
          let ok = true, formulas = true;
          for (let r = 0; r < 5; r++) {
            if (Math.abs(numAt(s, 4, r) - expected[r]) > 0.01 || isNaN(numAt(s, 4, r))) ok = false;
            if (!rawAt(s, 4, r).startsWith("=")) formulas = false;
          }
          return { ok, feedback: ok
            ? (formulas ? "¡Perfecto! Usaste fórmulas — si cambian los números, se recalcula solo."
                        : "¡Los números están bien! (Tip: la próxima probá con la fórmula =C1*D1 en vez del número a mano — así se recalcula sola).")
            : "Todavía no cierra. Revisá la columna Total. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Parte 1 completa — los totales por día están bien.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Ahora Marta necesita el total de toda la semana junta. Escribí el <b>Total General</b> en la última fila de la columna Total.<br>
          <span style="color:var(--text-dim);">Sumar los cinco números uno por uno es lento... ¿no habrá una forma más rápida?</span>`,
        hints: [
          "Podrías sumar celda por celda... pero Excel tiene una FUNCIÓN que suma un rango entero de una sola vez.",
          "La función se llama SUMA. Un rango se escribe con dos puntos: E1:E5 significa «desde E1 hasta E5».",
          "En la celda de Total General escribí: =SUMA(E1:E5)",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 4, 5) - 14600) < 0.01;
          return { ok, feedback: ok
            ? (rawAt(s, 4, 5).startsWith("=") ? "¡Cerró! El sótano dejó de parpadear."
                                              : "¡Cerró! (Tip: la próxima probá =SUMA(E1:E5) — suma un rango entero de una vez).")
            : "Ese no es el total de la semana todavía. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 2,
    name: "EP.2 — LA SEMANA FANTASMA",
    concepts: "PROMEDIO · MAX · MIN",
    briefing: [
      { speaker: "MARTA", text: "Volviste. Bien. Anoche encontré este cuaderno viejo en el sótano... tiene las ventas de una semana de 1983 que nadie recuerda.", portrait: true, bg: "basement" },
      { speaker: "MARTA", text: "Quiero entender esa semana: cuánto se vendía en un día típico, cuál fue el mejor día y cuál el peor.", portrait: true },
      { speaker: "MARTA", text: "Sumar ya sabés. Pero para esto hay funciones nuevas... vas a tener que descubrirlas. El botón 💡 PISTA sigue ahí.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Un promedio de $3.300... el mejor día $5.200 y el peor $2.000. Igual que la semana en que desapareció el dueño anterior...", portrait: true },
      { speaker: "MARTA", text: "No importa. Ya sabés resumir datos: promedio, máximo, mínimo. Eso en análisis de datos se usa TODOS los días.", portrait: true },
    ],
    title: "LA SEMANA FANTASMA",
    sheet: {
      data: [
        ["Lunes", 2400],
        ["Martes", 3000],
        ["Miércoles", 4000],
        ["Jueves", 2000],
        ["Viernes", 3200],
        ["Sábado", 5200],
        ["", ""],
        ["PROMEDIO semanal", ""],
        ["MEJOR día (máximo)", ""],
        ["PEOR día (mínimo)", ""],
      ],
      columns: [
        { type: "text", title: "Día", width: 220, readOnly: true },
        { type: "numeric", title: "Total vendido", width: 160 },
      ],
    },
    protect: [
      { x: 1, y: 0, v: 2400 }, { x: 1, y: 1, v: 3000 }, { x: 1, y: 2, v: 4000 },
      { x: 1, y: 3, v: 2000 }, { x: 1, y: 4, v: 3200 }, { x: 1, y: 5, v: 5200 },
    ],
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          ¿Cuánto vendía la tienda en un día típico de esa semana? Calculá el <b>PROMEDIO semanal</b> en la celda B8.<br>
          <span style="color:var(--text-dim);">Los totales de la semana están en B1:B6. Existe una función que hace el promedio de un rango entero...</span>`,
        hints: [
          "El promedio es: sumar todo y dividir por la cantidad de días. Pero Excel tiene una función que lo hace de una.",
          "La función se llama PROMEDIO y se usa igual que SUMA: =PROMEDIO(rango).",
          "En la celda B8 escribí: =PROMEDIO(B1:B6)",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 1, 7) - 3300) < 0.01;
          return { ok, feedback: ok
            ? "¡Eso! Un día típico de esa semana: $3.300."
            : "Ese no es el promedio de la semana. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Promedio listo: $3.300 por día.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Ahora encontrá el <b>mejor día</b> (el valor más alto, en B9) y el <b>peor día</b> (el más bajo, en B10).<br>
          <span style="color:var(--text-dim);">Podrías mirarlos a ojo... pero con 6 filas. ¿Y si fueran 60.000? Hay funciones para esto.</span>`,
        hints: [
          "Excel tiene una función que devuelve el número MÁS GRANDE de un rango, y otra para el MÁS CHICO.",
          "Se llaman MAX y MIN (así, en inglés corto). Se usan como SUMA: =MAX(rango).",
          "En B9 escribí =MAX(B1:B6) y en B10 escribí =MIN(B1:B6)",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 1, 8) - 5200) < 0.01 && Math.abs(numAt(s, 1, 9) - 2000) < 0.01;
          return { ok, feedback: ok
            ? "¡Máximo y mínimo encontrados!"
            : "Todavía falta: B9 debe tener el valor más alto y B10 el más bajo. (¿Una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 3,
    name: "EP.3 — CASSETTES CONTRA VINILOS",
    concepts: "CONTAR.SI · SUMAR.SI (condiciones)",
    briefing: [
      { speaker: "MARTA", text: "Esto se pone raro. El proveedor dice que le debemos plata por los VINILOS... pero solo por los vinilos. Y yo tengo todo mezclado en una sola lista.", portrait: true },
      { speaker: "MARTA", text: "Necesito dos cosas: cuántas ventas fueron de Cassette, y cuánta plata entró SOLO por los Vinilos.", portrait: true },
      { speaker: "MARTA", text: "Contar y sumar ya sabés... pero acá hay que contar y sumar SOLO LO QUE CUMPLE UNA CONDICIÓN. Eso es nuevo.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "4 cassettes vendidos y $7.500 de vinilos. Exactamente lo que decía el proveedor. Misterio resuelto.", portrait: true },
      { speaker: "MARTA", text: "Lo que acabás de hacer — filtrar con una condición — es la base de TODO el análisis de datos. Lo que viene ahora... buscar datos entre miles de filas. Pero eso es otra noche.", portrait: true },
    ],
    title: "CASSETTES CONTRA VINILOS",
    sheet: {
      data: [
        ["Blondie - Parallel Lines", "Cassette", 800],
        ["Bowie - Let's Dance", "Vinilo", 1500],
        ["Queen - Greatest Hits", "Cassette", 800],
        ["Talking Heads - 77", "Vinilo", 2000],
        ["Bowie - Heroes", "Cassette", 800],
        ["Soda Stereo - Nada Personal", "Vinilo", 1800],
        ["Charly - Clics Modernos", "Cassette", 900],
        ["Spinetta - Artaud", "Vinilo", 2200],
        ["", "", ""],
        ["¿Cuántas ventas de Cassette?", "", ""],
        ["¿Cuánta plata en Vinilos?", "", ""],
      ],
      columns: [
        { type: "text", title: "Producto", width: 260, readOnly: true },
        { type: "text", title: "Tipo", width: 110, readOnly: true },
        { type: "numeric", title: "Monto", width: 130 },
      ],
    },
    protect: [
      { x: 2, y: 0, v: 800 }, { x: 2, y: 1, v: 1500 }, { x: 2, y: 2, v: 800 },
      { x: 2, y: 3, v: 2000 }, { x: 2, y: 4, v: 800 }, { x: 2, y: 5, v: 1800 },
      { x: 2, y: 6, v: 900 }, { x: 2, y: 7, v: 2200 },
    ],
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          ¿Cuántas de estas ventas fueron de <b>Cassette</b>? Poné el resultado en la celda C10.<br>
          <span style="color:var(--text-dim);">Los tipos están en B1:B8. Podrías contarlos con el dedo... pero la idea es que Excel cuente solo las filas que cumplen la condición.</span>`,
        hints: [
          "Necesitás CONTAR, pero solo las celdas que dicen «Cassette». Es contar CON UNA CONDICIÓN (por eso la función termina en .SI).",
          "La función es CONTAR.SI y se usa así: =CONTAR.SI(rango; condición). La condición va entre comillas: \"Cassette\". Ojo: acá el separador es coma.",
          "En C10 escribí: =CONTAR.SI(B1:B8,\"Cassette\")",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 2, 9) - 4) < 0.01;
          return { ok, feedback: ok
            ? "¡4 cassettes! Contaste solo lo que cumplía la condición."
            : "Ese no es el número de ventas de Cassette. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ 4 ventas de Cassette contadas.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Ahora la plata: ¿cuánto entró <b>solo por los Vinilos</b>? Poné el resultado en la celda C11.<br>
          <span style="color:var(--text-dim);">Es como la parte 1, pero en vez de contar filas hay que SUMAR los montos (C1:C8) de las filas que son Vinilo.</span>`,
        hints: [
          "Es una SUMA con condición: sumar los montos solo donde el tipo es «Vinilo». La función hermana de CONTAR.SI...",
          "Se llama SUMAR.SI y lleva tres partes: =SUMAR.SI(rango_condición, condición, rango_a_sumar).",
          "En C11 escribí: =SUMAR.SI(B1:B8,\"Vinilo\",C1:C8)",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 2, 10) - 7500) < 0.01;
          return { ok, feedback: ok
            ? "¡$7.500 en vinilos! Sumaste solo lo que cumplía la condición."
            : "Ese no es el total de los Vinilos. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 4,
    name: "EP.4 — EL CLIENTE DE MEDIANOCHE",
    concepts: "BUSCARV: buscar en una tabla",
    briefing: [
      { speaker: "MARTA", text: "Anoche alguien deslizó una nota bajo la puerta. Solo dice: «VIN-004. Medianoche. Que esté listo.»", portrait: true },
      { speaker: "MARTA", text: "VIN-004 es un código de nuestro catálogo. Necesito saber YA qué producto es y cuánto cuesta.", portrait: true },
      { speaker: "MARTA", text: "Con 8 filas lo buscás a ojo... pero el catálogo del depósito tiene 8.000. Hay una función para esto, y es la más importante que vas a aprender.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "«Talking Heads - 77»... a $2.000. Preparalo. Y no preguntes quién lo viene a buscar.", portrait: true },
      { speaker: "MARTA", text: "Acordate de esta función: BUSCARV. En cualquier trabajo con datos la vas a usar todas las semanas. (En el Excel real se le agrega FALSO al final para búsqueda exacta: =BUSCARV(\"VIN-004\",A1:C8,3,FALSO)).", portrait: true },
    ],
    title: "EL CLIENTE DE MEDIANOCHE",
    lookup: { x0: 0, y0: 0, cols: 3, rows: 8 },
    sheet: {
      data: [
        ["CAS-001", "Blondie - Parallel Lines", 800],
        ["VIN-002", "Bowie - Let's Dance", 1500],
        ["CAS-003", "Queen - Greatest Hits", 800],
        ["VIN-004", "Talking Heads - 77", 2000],
        ["CAS-005", "Bowie - Heroes", 800],
        ["VIN-006", "Soda Stereo - Nada Personal", 1800],
        ["CAS-007", "Charly - Clics Modernos", 900],
        ["VIN-008", "Spinetta - Artaud", 2200],
        ["", "", ""],
        ["CÓDIGO MISTERIOSO:", "VIN-004", ""],
        ["¿Qué producto es?", "", ""],
        ["¿Cuánto cuesta?", "", ""],
      ],
      columns: [
        { type: "text", title: "Código", width: 150, readOnly: true },
        { type: "text", title: "Producto", width: 240, readOnly: true },
        { type: "numeric", title: "Precio", width: 130 },
      ],
    },
    protect: [
      { x: 2, y: 0, v: 800 }, { x: 2, y: 1, v: 1500 }, { x: 2, y: 2, v: 800 },
      { x: 2, y: 3, v: 2000 }, { x: 2, y: 4, v: 800 }, { x: 2, y: 5, v: 1800 },
      { x: 2, y: 6, v: 900 }, { x: 2, y: 7, v: 2200 },
    ],
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          ¿Qué producto es el código <b>VIN-004</b>? Respondé en la celda C11 — pero no vale copiarlo a mano: la idea es que <b>Excel lo busque</b> en la tabla A1:C8.<br>
          <span style="color:var(--text-dim);">Imaginate que el catálogo tuviera 8.000 filas. Hay una función que busca un valor en la primera columna de una tabla y devuelve otro dato de esa misma fila.</span>`,
        hints: [
          "La función busca VERTICALMENTE (fila por fila) en la primera columna de una tabla. Por eso su nombre termina en V.",
          "Se llama BUSCARV: =BUSCARV(valor_buscado, tabla, número_de_columna). El código va entre comillas: \"VIN-004\". La tabla es A1:C8. El producto es la columna 2 de esa tabla.",
          "En C11 escribí: =BUSCARV(\"VIN-004\",A1:C8,2)",
        ],
        check(s) {
          const val = String(s.getValueFromCoords(2, 10, true) || "").toLowerCase();
          const raw = rawAt(s, 2, 10).toUpperCase();
          const rightValue = val.includes("talking");
          if (rightValue && !raw.includes("BUSCARV")) {
            return { ok: false, feedback: "El producto es ese... ¡pero lo copiaste a mano! Con 8.000 filas no podrías. Usá la función BUSCARV. (¿Una 💡 PISTA?)" };
          }
          return { ok: rightValue, feedback: rightValue
            ? "¡«Talking Heads - 77»! BUSCARV lo encontró solo."
            : "Ese no es el producto del código VIN-004. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Producto encontrado: Talking Heads - 77.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Ahora el precio: ¿cuánto cuesta el <b>VIN-004</b>? Respondé en C12, otra vez con BUSCARV.<br>
          <span style="color:var(--text-dim);">Es la misma búsqueda... lo único que cambia es QUÉ columna de la tabla querés que te devuelva.</span>`,
        hints: [
          "Ya tenés la fórmula de la parte 1. Pensá qué número le decía a Excel «devolveme el producto»...",
          "El precio está en la columna 3 de la tabla A1:C8. Solo cambia el último número de la fórmula.",
          "En C12 escribí: =BUSCARV(\"VIN-004\",A1:C8,3)",
        ],
        check(s) {
          const ok = Math.abs(numAt(s, 2, 11) - 2000) < 0.01;
          const raw = rawAt(s, 2, 11).toUpperCase();
          if (ok && !raw.includes("BUSCARV")) {
            return { ok: false, feedback: "El precio es ese, pero copiado a mano no cuenta — usá BUSCARV. (¿Una 💡 PISTA?)" };
          }
          return { ok, feedback: ok
            ? "¡$2.000! Búsqueda completa."
            : "Ese no es el precio del VIN-004. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 5,
    name: "EP.5 — LA TABLA QUE SE ARMA SOLA",
    concepts: "Tablas dinámicas: agrupar y resumir",
    briefing: [
      { speaker: "MARTA", text: "El contador del pueblo me pide «el resumen por tipo de producto»: cuántas ventas y cuánta plata por cada tipo. TODOS los meses lo mismo.", portrait: true },
      { speaker: "MARTA", text: "En el Excel real esto lo hace una TABLA DINÁMICA en dos clics. Pero para usarla bien, primero hay que entender QUÉ hace por dentro.", portrait: true },
      { speaker: "MARTA", text: "Así que hoy la armamos a mano: agrupar por tipo, y resumir cada grupo. Ya tenés todas las herramientas... solo hay que combinarlas.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Esto que acabás de armar ES una tabla dinámica: agrupar filas por una categoría y resumir cada grupo (contar, sumar, promediar).", portrait: true },
      { speaker: "MARTA", text: "En el Excel real: Insertar → Tabla dinámica, arrastrás «Tipo» a filas y «Monto» a valores... y esta tabla se arma sola. Ahora cuando la uses, vas a saber exactamente qué está haciendo.", portrait: true },
    ],
    title: "LA TABLA QUE SE ARMA SOLA",
    sheet: {
      data: [
        ["Cassette", 800, ""],
        ["Vinilo", 1500, ""],
        ["CD", 1200, ""],
        ["Cassette", 900, ""],
        ["Vinilo", 2000, ""],
        ["Cassette", 800, ""],
        ["CD", 1100, ""],
        ["Vinilo", 1800, ""],
        ["Cassette", 700, ""],
        ["CD", 1300, ""],
        ["", "", ""],
        ["RESUMEN POR TIPO", "Cantidad", "Total $"],
        ["Cassette", "", ""],
        ["Vinilo", "", ""],
        ["CD", "", ""],
      ],
      columns: [
        { type: "text", title: "Tipo", width: 180, readOnly: true },
        { type: "numeric", title: "Monto", width: 140 },
        { type: "numeric", title: "", width: 140 },
      ],
    },
    protect: [
      { x: 1, y: 0, v: 800 }, { x: 1, y: 1, v: 1500 }, { x: 1, y: 2, v: 1200 },
      { x: 1, y: 3, v: 900 }, { x: 1, y: 4, v: 2000 }, { x: 1, y: 5, v: 800 },
      { x: 1, y: 6, v: 1100 }, { x: 1, y: 7, v: 1800 }, { x: 1, y: 8, v: 700 },
      { x: 1, y: 9, v: 1300 },
      { x: 1, y: 11, v: "Cantidad" }, { x: 2, y: 11, v: "Total $" },
    ],
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          Primera columna del resumen: la <b>Cantidad</b> de ventas de cada tipo. Completá B13 (Cassette), B14 (Vinilo) y B15 (CD).<br>
          <span style="color:var(--text-dim);">Los tipos están en A1:A10. Esto ya lo sabés hacer de la noche de los cassettes contra los vinilos...</span>`,
        hints: [
          "Contar solo las filas que cumplen una condición... ¿te suena de una misión anterior?",
          "Es CONTAR.SI: =CONTAR.SI(rango, condición). El rango de tipos es A1:A10.",
          "B13: =CONTAR.SI(A1:A10,\"Cassette\") — B14: =CONTAR.SI(A1:A10,\"Vinilo\") — B15: =CONTAR.SI(A1:A10,\"CD\")",
        ],
        check(s) {
          const ok = numAt(s, 1, 12) === 4 && numAt(s, 1, 13) === 3 && numAt(s, 1, 14) === 3;
          const usedFn = [12, 13, 14].every((y) => rawAt(s, 1, y).toUpperCase().includes("CONTAR"));
          if (ok && !usedFn) {
            return { ok: false, feedback: "Los números están bien... pero contados a mano. La gracia es que Excel cuente solo: usá CONTAR.SI. (¿Una 💡 PISTA?)" };
          }
          return { ok, feedback: ok
            ? "¡Grupo contado! 4 cassettes, 3 vinilos, 3 CDs."
            : "Todavía no están las tres cantidades correctas. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Cantidades listas: 4 / 3 / 3.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Segunda columna: el <b>Total $</b> de cada tipo. Completá C13, C14 y C15.<br>
          <span style="color:var(--text-dim);">Ahora no es contar filas: es sumar los montos (B1:B10) de las filas de cada tipo.</span>`,
        hints: [
          "Sumar con condición... la hermana de CONTAR.SI.",
          "SUMAR.SI lleva tres partes: =SUMAR.SI(rango_condición, condición, rango_a_sumar). Condiciones en A1:A10, montos en B1:B10.",
          "C13: =SUMAR.SI(A1:A10,\"Cassette\",B1:B10) — y lo mismo con \"Vinilo\" y \"CD\".",
        ],
        check(s) {
          const ok = numAt(s, 2, 12) === 3200 && numAt(s, 2, 13) === 5300 && numAt(s, 2, 14) === 3600;
          const usedFn = [12, 13, 14].every((y) => rawAt(s, 2, y).toUpperCase().includes("SUMAR"));
          if (ok && !usedFn) {
            return { ok: false, feedback: "Los totales están bien pero sumados a mano — usá SUMAR.SI para que Excel lo haga solo. (¿Una 💡 PISTA?)" };
          }
          return { ok, feedback: ok
            ? "¡Resumen completo! Acabás de armar tu primera tabla dinámica (a mano)."
            : "Todavía no están los tres totales correctos. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 6,
    name: "EP.6 — EL INFORME FINAL",
    concepts: "Proyecto final: todo junto",
    briefing: [
      { speaker: "MARTA", text: "Llegó el momento. El banco del pueblo quiere EL INFORME COMPLETO de la tienda para renovarnos el crédito. Todo lo que aprendiste, junto.", portrait: true },
      { speaker: "MARTA", text: "Total, promedio, la venta más alta, cuántos vinilos, y el dato de un código puntual. Sin ayuda esta vez... bueno, las pistas siguen ahí. Pero intentalo solo.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Informe completo. El crédito está renovado... y la luz del sótano no parpadeó ni una vez este mes.", portrait: true, bg: "basement" },
      { speaker: "MARTA", text: "Ya no te necesito... pero el pueblo sí. Dicen que en la biblioteca hay planillas de hace 40 años que nadie pudo abrir. Power Query, les dicen. Pero esa... esa es otra historia.", portrait: true },
    ],
    title: "EL INFORME FINAL",
    lookup: { x0: 0, y0: 0, cols: 3, rows: 8 },
    sheet: {
      data: [
        ["CAS-001", "Cassette", 800],
        ["VIN-002", "Vinilo", 1500],
        ["CAS-003", "Cassette", 800],
        ["VIN-004", "Vinilo", 2000],
        ["CAS-005", "Cassette", 800],
        ["VIN-006", "Vinilo", 1800],
        ["CAS-007", "Cassette", 900],
        ["VIN-008", "Vinilo", 2200],
        ["", "", ""],
        ["TOTAL GENERAL", "", ""],
        ["PROMEDIO POR VENTA", "", ""],
        ["VENTA MÁS ALTA", "", ""],
        ["CANT. VENTAS DE VINILO", "", ""],
        ["MONTO DEL CÓDIGO CAS-007", "", ""],
      ],
      columns: [
        { type: "text", title: "Código", width: 200, readOnly: true },
        { type: "text", title: "Tipo", width: 120, readOnly: true },
        { type: "numeric", title: "Monto", width: 140 },
      ],
    },
    protect: [
      { x: 2, y: 0, v: 800 }, { x: 2, y: 1, v: 1500 }, { x: 2, y: 2, v: 800 },
      { x: 2, y: 3, v: 2000 }, { x: 2, y: 4, v: 800 }, { x: 2, y: 5, v: 1800 },
      { x: 2, y: 6, v: 900 }, { x: 2, y: 7, v: 2200 },
    ],
    stages: [
      {
        instructions: `<span class="stage-label">INFORME — PARTE 1 de 2</span>
          Los tres números grandes del informe, en la columna C:<br>
          <b>C10</b>: total general de ventas · <b>C11</b>: promedio por venta · <b>C12</b>: la venta más alta.<br>
          <span style="color:var(--text-dim);">Los montos están en C1:C8. Todo esto ya lo hiciste antes — confía.</span>`,
        hints: [
          "Tres funciones que ya conocés: una suma un rango, otra saca el promedio, otra encuentra el máximo.",
          "SUMA, PROMEDIO y MAX — las tres se usan igual: =FUNCIÓN(C1:C8).",
          "C10: =SUMA(C1:C8) — C11: =PROMEDIO(C1:C8) — C12: =MAX(C1:C8)",
        ],
        check(s) {
          const ok = numAt(s, 2, 9) === 10800 && numAt(s, 2, 10) === 1350 && numAt(s, 2, 11) === 2200;
          return { ok, feedback: ok
            ? "¡Totales del informe listos!"
            : "Revisá: C10 el total, C11 el promedio, C12 el máximo. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Totales listos: $10.800 · $1.350 · $2.200.</span>
          <span class="stage-label" style="margin-top:10px;">INFORME — PARTE 2 de 2</span>
          Los dos datos finos:<br>
          <b>C13</b>: cuántas ventas fueron de Vinilo · <b>C14</b>: el monto del código CAS-007 (buscado con fórmula, no a ojo).<br>
          <span style="color:var(--text-dim);">Una es contar con condición; la otra es buscar en la tabla A1:C8.</span>`,
        hints: [
          "Para C13: contar con una condición (los tipos están en B1:B8). Para C14: buscar un código en una tabla... las dos las usaste en misiones anteriores.",
          "C13 usa CONTAR.SI sobre B1:B8 con \"Vinilo\". C14 usa BUSCARV sobre A1:C8 — el monto es la columna 3 de la tabla.",
          "C13: =CONTAR.SI(B1:B8,\"Vinilo\") — C14: =BUSCARV(\"CAS-007\",A1:C8,3)",
        ],
        check(s) {
          const okCount = numAt(s, 2, 12) === 4;
          const okLookup = numAt(s, 2, 13) === 900;
          const rawLookup = rawAt(s, 2, 13).toUpperCase();
          if (okCount && okLookup && !rawLookup.includes("BUSCARV")) {
            return { ok: false, feedback: "El 900 está bien pero copiado a ojo — el banco exige fórmulas: usá BUSCARV en C14. (¿Una 💡 PISTA?)" };
          }
          const ok = okCount && okLookup;
          return { ok, feedback: ok
            ? "¡INFORME COMPLETO! Usaste todo lo aprendido."
            : "Falta algo: C13 la cantidad de vinilos, C14 el monto del CAS-007. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 7,
    name: "EP.7 — EL SEMÁFORO DE STOCK",
    concepts: "SI · Y · O (decisiones)",
    intro: [
      { speaker: "", text: "Pasaron unos meses. El negocio funciona mejor... pero ahora hay decisiones más finas que tomar.", portrait: false, bg: "town" },
    ],
    briefing: [
      { speaker: "MARTA", text: "Mirá esta lista de stock. Necesito que la planilla me diga sola cuáles hay que reponer... sin que yo tenga que revisarlas una por una.", portrait: true },
      { speaker: "MARTA", text: "Para eso existen las funciones que TOMAN DECISIONES. La más importante se llama SI.", portrait: true },
      { speaker: "MARTA", text: "Después vamos a afinar más: no todo lo que falta reponer es urgente. Eso es para la parte 2.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Ahora la planilla decide sola. Bueno... vos le decís CÓMO decidir, y ella lo hace en cada fila.", portrait: true },
      { speaker: "MARTA", text: "Esto que hiciste — pedir que se cumplan DOS condiciones con Y — también existe al revés: la función O, que alcanza con que se cumpla UNA sola. La vas a necesitar tarde o temprano.", portrait: true },
    ],
    title: "EL SEMÁFORO DE STOCK",
    sheet: {
      data: [
        ["Cassette - Blondie", 8, 5, "", ""],
        ["Vinilo - Bowie", 2, 5, "", ""],
        ["CD - Queen", 5, 5, "", ""],
        ["Vinilo - Talking Heads", 1, 3, "", ""],
        ["Cassette - Charly", 0, 4, "", ""],
      ],
      columns: [
        { type: "text", title: "Producto", width: 190, readOnly: true },
        { type: "numeric", title: "Stock actual", width: 110, readOnly: true },
        { type: "numeric", title: "Stock mínimo", width: 110, readOnly: true },
        { type: "text", title: "Estado", width: 130 },
        { type: "text", title: "¿Urgente?", width: 120 },
      ],
    },
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          Completá la columna <b>Estado</b>: "REPONER" si el Stock actual es menor al Stock mínimo, o "OK" si no.<br>
          <span style="color:var(--text-dim);">Esto es una DECISIÓN: según se cumpla o no una condición, un resultado u otro.</span>`,
        hints: [
          "Para cada producto, comparás el Stock actual contra el Stock mínimo. Si el actual es menor... hay que reponer.",
          'La función que decide entre dos resultados según una condición se llama SI: =SI(condición, resultado_si_true, resultado_si_false).',
          'En Estado de la fila 1 escribí: =SI(B1<C1,"REPONER","OK") y repetí el patrón en las demás filas.',
        ],
        check(s) {
          const expected = ["OK", "REPONER", "OK", "REPONER", "REPONER"];
          let ok = true, usedFn = true;
          for (let r = 0; r < 5; r++) {
            if (String(s.getValueFromCoords(3, r, true) || "").trim().toUpperCase() !== expected[r]) ok = false;
            if (!rawAt(s, 3, r).toUpperCase().includes("SI(")) usedFn = false;
          }
          if (ok && !usedFn) return { ok: false, feedback: "Los resultados están bien, pero escritos a mano — usá la función SI para que decida sola. (¿Una 💡 PISTA?)" };
          return { ok, feedback: ok ? "¡La planilla decide sola qué reponer!" : "Todavía no está bien la columna Estado. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Estado listo.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Completá <b>¿Urgente?</b>: "URGENTE" solo si el producto necesita reposición <b>Y</b> le queda 1 unidad o menos. Si no, "-".<br>
          <span style="color:var(--text-dim);">No alcanza con una condición: acá hacen falta las DOS a la vez.</span>`,
        hints: [
          "No alcanza con que falte stock: acá buscamos los casos REALMENTE urgentes — los que además tienen casi nada (1 unidad o menos).",
          "Cuando necesitás que se cumplan DOS condiciones a la vez, existe la función Y: =Y(condición1,condición2) da verdadero solo si las dos son ciertas.",
          'En ¿Urgente? de la fila 1 escribí: =SI(Y(B1<C1,B1<=1),"URGENTE","-") y repetí el patrón.',
        ],
        check(s) {
          const expected = ["-", "-", "-", "URGENTE", "URGENTE"];
          let ok = true, usedFn = true;
          for (let r = 0; r < 5; r++) {
            if (String(s.getValueFromCoords(4, r, true) || "").trim().toUpperCase() !== expected[r]) ok = false;
            const raw = rawAt(s, 4, r).toUpperCase();
            if (!raw.includes("SI(") || !raw.includes("Y(")) usedFn = false;
          }
          if (ok && !usedFn) return { ok: false, feedback: "Los resultados están bien, pero probá combinando SI con Y para que sea automático. (¿Una 💡 PISTA?)" };
          return { ok, feedback: ok ? "¡Encontraste los casos realmente urgentes!" : "Todavía no está bien la columna ¿Urgente?. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 8,
    name: "EP.8 — LOS MENSAJES CIFRADOS",
    concepts: "MAYUSC · ESPACIOS · IZQUIERDA · DERECHA",
    briefing: [
      { speaker: "MARTA", text: "El sistema viejo del depósito exportó esta lista de socios... y quedó todo desprolijo. Espacios de más, mayúsculas mezcladas con minúsculas.", portrait: true },
      { speaker: "MARTA", text: "No podemos mandarle esto al banco así. Hay que limpiarlo con fórmulas — a mano tardaríamos toda la noche.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Ahora sí, una lista prolija. Esto que aprendiste — limpiar texto con fórmulas — te va a ahorrar horas cada vez que alguien te pase una planilla desordenada. Que va a ser siempre.", portrait: true },
      { speaker: "MARTA", text: "Los códigos de producto también estaban ilegibles. Ya sabés separar un texto en partes y volver a unirlas como quieras.", portrait: true },
    ],
    title: "LOS MENSAJES CIFRADOS",
    sheet: {
      data: [
        ["  juan perez  ", "", "CAS0012", ""],
        ["MARIA  lopez", "", "VIN0045", ""],
        ["carlos GOMEZ  ", "", "CDX0100", ""],
        ["  ana torres", "", "CAS0099", ""],
        ["Diego  Ramirez  ", "", "VIN0007", ""],
      ],
      columns: [
        { type: "text", title: "Nombre crudo", width: 170, readOnly: true },
        { type: "text", title: "Nombre VIP", width: 190 },
        { type: "text", title: "Código crudo", width: 130, readOnly: true },
        { type: "text", title: "Código legible", width: 150 },
      ],
    },
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          Completá <b>Nombre VIP</b>: el nombre sin espacios de más, todo en MAYÚSCULAS.<br>
          <span style="color:var(--text-dim);">Necesitás dos funciones juntas: una que limpie los espacios y otra que ponga mayúsculas — una DENTRO de la otra.</span>`,
        hints: [
          "Estos nombres tienen espacios de más al principio, en el medio o al final. Antes que nada, hay que LIMPIARLOS.",
          "La función que saca los espacios sobrantes se llama ESPACIOS. Y para pasar todo a mayúsculas existe MAYUSC. Podés meter una función DENTRO de la otra.",
          "En Nombre VIP de la fila 1 escribí: =MAYUSC(ESPACIOS(A1)) y repetí el patrón.",
        ],
        check(s) {
          const expected = ["JUAN PEREZ", "MARIA LOPEZ", "CARLOS GOMEZ", "ANA TORRES", "DIEGO RAMIREZ"];
          let ok = true, usedFn = true;
          for (let r = 0; r < 5; r++) {
            if (String(s.getValueFromCoords(1, r, true) || "").trim().toUpperCase() !== expected[r]) ok = false;
            const raw = rawAt(s, 1, r).toUpperCase();
            if (!raw.includes("ESPACIOS(") || !raw.includes("MAYUSC(")) usedFn = false;
          }
          if (ok && !usedFn) return { ok: false, feedback: "Los nombres quedaron bien... ¿pero los limpiaste a mano? Usá ESPACIOS y MAYUSC combinadas. (¿Una 💡 PISTA?)" };
          return { ok, feedback: ok ? "¡Lista de socios prolija!" : "Todavía hay nombres mal. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Nombres limpios.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          Completá <b>Código legible</b>: separá el código en categoría y número con un guion (ej: "CAS0012" → "CAS-0012").<br>
          <span style="color:var(--text-dim);">El código tiene 3 letras de categoría y 4 números de ID. Hay funciones para agarrar "la izquierda" y "la derecha" de un texto.</span>`,
        hints: [
          "El código tiene dos partes: las primeras 3 letras (categoría) y los últimos 4 números (ID). Hay funciones que agarran «la izquierda» y «la derecha» de un texto.",
          "Se llaman IZQUIERDA y DERECHA: =IZQUIERDA(texto,cantidad) y =DERECHA(texto,cantidad). Para juntarlas con un guion podés usar & o CONCATENAR.",
          'En Código legible de la fila 1 escribí: =CONCATENAR(IZQUIERDA(C1,3),"-",DERECHA(C1,4)) y repetí el patrón.',
        ],
        check(s) {
          const expected = ["CAS-0012", "VIN-0045", "CDX-0100", "CAS-0099", "VIN-0007"];
          let ok = true, usedFn = true;
          for (let r = 0; r < 5; r++) {
            if (String(s.getValueFromCoords(3, r, true) || "").trim().toUpperCase() !== expected[r]) ok = false;
            const raw = rawAt(s, 3, r).toUpperCase();
            if (!raw.includes("IZQUIERDA(") || !raw.includes("DERECHA(")) usedFn = false;
          }
          if (ok && !usedFn) return { ok: false, feedback: "Los códigos están bien, pero armados a mano — usá IZQUIERDA y DERECHA. (¿Una 💡 PISTA?)" };
          return { ok, feedback: ok ? "¡Códigos legibles!" : "Todavía hay códigos mal. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 9,
    name: "EP.9 — EL EXPEDIENTE ANTIGUO",
    concepts: "BUSCARX · AÑO (fechas)",
    briefing: [
      { speaker: "MARTA", text: "Encontré una caja en el sótano: fichas de compras viejas, de antes de que yo llevara la tienda. Están ordenadas por código, no por nombre.", portrait: true, bg: "basement" },
      { speaker: "MARTA", text: "Con BUSCARV esto no se puede: el nombre del producto está a la IZQUIERDA del código, y esa función solo mira hacia la derecha.", portrait: true },
      { speaker: "MARTA", text: "Por suerte existe una versión más moderna, sin esa restricción.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "BUSCARX no le importa de qué lado está la columna que buscás. Es la que se usa hoy en día — quedate con esa.", portrait: true },
      { speaker: "MARTA", text: "Y ya sabés sacarle el año a una fecha. Con eso alcanza para armar cualquier reporte por período.", portrait: true },
    ],
    title: "EL EXPEDIENTE ANTIGUO",
    lookup: { x0: 0, y0: 0, cols: 4, rows: 8 },
    lookupX: { searchCol: 1, returnCol: 0 },
    sheet: {
      data: [
        ["Blondie - Parallel Lines", "CAS-001", 800, "2024-03-10"],
        ["Bowie - Let's Dance", "VIN-002", 1500, "2024-06-01"],
        ["Queen - Greatest Hits", "CAS-003", 800, "2024-01-15"],
        ["Talking Heads - 77", "VIN-004", 2000, "2023-11-20"],
        ["Bowie - Heroes", "CAS-005", 800, "2024-02-05"],
        ["Soda Stereo - Nada Personal", "VIN-006", 1800, "2024-07-22"],
        ["Charly - Clics Modernos", "CAS-007", 900, "2023-12-30"],
        ["Spinetta - Artaud", "VIN-008", 2200, "2024-04-18"],
        ["", "", "", ""],
        ["¿Qué producto es el código VIN-004?", "", "", ""],
        ["¿En qué año se compró el VIN-004?", "", "", ""],
      ],
      columns: [
        { type: "text", title: "Producto", width: 220, readOnly: true },
        { type: "text", title: "Código", width: 100, readOnly: true },
        { type: "numeric", title: "Precio", width: 100, readOnly: true },
        { type: "text", title: "Fecha compra", width: 130, readOnly: true },
        { type: "text", title: "Respuesta", width: 200 },
      ],
    },
    stages: [
      {
        instructions: `<span class="stage-label">MISIÓN — PARTE 1 de 2</span>
          ¿Qué producto es el código <b>VIN-004</b>? Respondé en E10 buscando en la columna Código y devolviendo Producto — que está a la IZQUIERDA.<br>
          <span style="color:var(--text-dim);">Esto es justo lo que BUSCARV no puede hacer.</span>`,
        hints: [
          "Necesitás buscar «VIN-004» en la columna Código (B) y traer el dato de la columna Producto (A), que está antes, a la izquierda.",
          "La función es BUSCARX: =BUSCARX(valor_buscado, columna_donde_buscar, columna_a_devolver).",
          'En E10 escribí: =BUSCARX("VIN-004",B1:B8,A1:A8)',
        ],
        check(s) {
          const val = String(s.getValueFromCoords(4, 9, true) || "").toLowerCase();
          const raw = rawAt(s, 4, 9).toUpperCase();
          const rightValue = val.includes("talking");
          if (rightValue && !raw.includes("BUSCARX")) {
            return { ok: false, feedback: "El producto es ese... ¡pero copiado a mano! Usá BUSCARX. (¿Una 💡 PISTA?)" };
          }
          return { ok: rightValue, feedback: rightValue ? "¡«Talking Heads - 77»! BUSCARX buscó hacia la izquierda, algo que BUSCARV no puede." : "Ese no es el producto del VIN-004. (¿Necesitás una 💡 PISTA?)" };
        },
      },
      {
        instructions: `<span class="success">✔ Producto encontrado: Talking Heads - 77.</span>
          <span class="stage-label" style="margin-top:10px;">MISIÓN — PARTE 2 de 2</span>
          ¿En qué <b>año</b> se compró? La fecha del VIN-004 está en D4. Respondé en E11.<br>
          <span style="color:var(--text-dim);">Hay una función que saca solo el año de una fecha completa.</span>`,
        hints: [
          "La fecha completa es «2023-11-20». De ahí necesitás solo la primera parte: el año.",
          "La función se llama AÑO y se usa así: =AÑO(celda_con_fecha).",
          "En E11 escribí: =AÑO(D4)",
        ],
        check(s) {
          const ok = numAt(s, 4, 10) === 2023;
          const raw = rawAt(s, 4, 10).toUpperCase();
          if (ok && !raw.includes("AÑO(") && !raw.includes("ANO(")) {
            return { ok: false, feedback: "2023 está bien, pero escrito a mano — usá la función AÑO. (¿Una 💡 PISTA?)" };
          }
          return { ok, feedback: ok ? "¡2023! Expediente completo." : "Ese no es el año de compra del VIN-004. (¿Necesitás una 💡 PISTA?)" };
        },
      },
    ],
  },

  {
    id: 10,
    type: "query",
    name: "EP.10 — EL ARCHIVO DEL SISTEMA VIEJO",
    concepts: "Power Query: limpiar datos en pasos, en orden",
    briefing: [
      { speaker: "MARTA", text: "El sistema viejo del depósito por fin arrancó... y escupió esto.", portrait: true, mood: "worried" },
      { speaker: "MARTA", text: "Es la lista de productos de hace años, exportada mil veces por gente distinta: el mismo producto escrito de formas distintas cada vez.", portrait: true },
      { speaker: "MARTA", text: "Esto ya no se arregla con una fórmula suelta. Hay que armar una CONSULTA: una serie de PASOS, uno atrás del otro, que limpien los datos solos.", portrait: true },
      { speaker: "MARTA", text: "Lo bueno de una consulta: el día que llegue otro archivo así de desprolijo, aplicás los mismos pasos y listo — no hay que rehacer nada a mano.", portrait: true },
    ],
    outro: [
      { speaker: "MARTA", text: "Lista limpia, sin repetidos, con el precio de cada producto. Y si mañana llega otro archivo desprolijo, aplicás la misma consulta de nuevo.", portrait: true },
      { speaker: "MARTA", text: "Eso que armaste es, en chiquito, lo que hace Power Query en el Excel de verdad: automatizar la limpieza de datos, paso por paso.", portrait: true },
      { speaker: "MARTA", text: "Lo que sigue ya no te lo puedo enseñar acá adentro: con estos datos limpios, se arman paneles — Business Intelligence — gráficos y números que se actualizan solos cada vez que cambia la planilla. Power BI, Excel real, todo conectado.", portrait: true },
      { speaker: "MARTA", text: "Bien. Llegaste hasta acá. El pueblo ya no tiene ningún número que no cierre... por ahora.", portrait: true },
      { speaker: "", text: "FIN DE LA TEMPORADA 1 — GRACIAS POR JUGAR", portrait: false, bg: "town" },
    ],
    title: "EL ARCHIVO DEL SISTEMA VIEJO",
    query: {
      raw: [
        " cassette blondie ",
        "Cassette Blondie",
        "VINILO BOWIE",
        "cassette QUEEN",
        " Vinilo Bowie",
        "CASSETTE queen ",
      ],
      priceTable: { "Cassette Blondie": 800, "Vinilo Bowie": 1500, "Cassette Queen": 800 },
      target: [
        { producto: "Cassette Blondie", precio: 800 },
        { producto: "Vinilo Bowie", precio: 1500 },
        { producto: "Cassette Queen", precio: 800 },
      ],
      steps: [
        {
          id: "trim",
          label: "Quitar espacios extra",
          fn: (rows) => rows.map((r) => ({ ...r, producto: r.producto.replace(/^\s+|\s+$/g, "") })),
        },
        {
          id: "titlecase",
          label: "Formato Título (cassette BLONDIE → Cassette Blondie)",
          fn: (rows) => rows.map((r) => ({ ...r, producto: r.producto.toLowerCase().replace(/\b\p{L}/gu, (c) => c.toUpperCase()) })),
        },
        {
          id: "dedupe",
          label: "Quitar filas duplicadas",
          fn: (rows) => {
            const seen = new Set();
            return rows.filter((r) => {
              if (seen.has(r.producto)) return false;
              seen.add(r.producto);
              return true;
            });
          },
        },
        {
          id: "merge",
          label: "Combinar con Tabla de Precios",
          fn: (rows) => rows.map((r) => ({ ...r, precio: mission.query.priceTable[r.producto] ?? null })),
        },
      ],
      stages: [
        {
          instructions: `<span class="stage-label">CONSULTA — PARTE 1 de 2</span>
            Los productos están repetidos y escritos de formas distintas. Agregá pasos (en el orden que quieras) hasta que la vista previa muestre <b>3 productos únicos, bien escritos</b>, sin la columna de precio todavía.<br>
            <span style="color:var(--text-dim);">Mirá la vista previa después de cada paso — te muestra si algo quedó raro. Podés deshacer el último paso si te equivocás.</span>`,
          hints: [
            "Antes de sacar los duplicados, los textos tienen que quedar IGUALES entre sí (mismos espacios, mismas mayúsculas) — si no, Excel no los reconoce como el mismo producto.",
            "Empezá por 'Quitar espacios extra' y 'Formato Título'. Recién ahí 'Quitar filas duplicadas' va a funcionar bien.",
            "Orden sugerido: Quitar espacios extra → Formato Título → Quitar filas duplicadas.",
          ],
          check(rows) {
            const ok = rows.length === 3 && rows.every((r) => ["Cassette Blondie", "Vinilo Bowie", "Cassette Queen"].includes(r.producto));
            return { ok, feedback: ok
              ? "¡3 productos limpios y sin repetidos!"
              : `Todavía hay ${rows.length} filas en la vista previa (debería haber 3, bien escritas). (¿Necesitás una 💡 PISTA?)` };
          },
        },
        {
          instructions: `<span class="success">✔ Lista limpia: 3 productos únicos.</span>
            <span class="stage-label" style="margin-top:10px;">CONSULTA — PARTE 2 de 2</span>
            Ahora agregá el último paso: sumar el <b>precio</b> de cada producto desde la Tabla de Precios.`,
          hints: [
            "Falta un paso que agregue la columna de precio, buscando cada producto en la Tabla de Precios.",
            "Es el paso 'Combinar con Tabla de Precios' — funciona porque los nombres ya están limpios y coinciden.",
            "Agregá el paso 'Combinar con Tabla de Precios' al final de la lista.",
          ],
          check(rows) {
            const ok = rows.length === 3 && rows.every((r) => r.precio === mission.query.priceTable[r.producto]);
            return { ok, feedback: ok
              ? "¡Consulta completa! Datos limpios, sin repetidos, con precio."
              : "Falta (o está mal) el precio de algún producto. (¿Necesitás una 💡 PISTA?)" };
          },
        },
      ],
    },
  },
];

const COMING_SOON = [];

/* ---------- 4. Referencias de pantalla ---------- */

const el = {
  menuScreen: document.getElementById("menu-screen"),
  mapScreen: document.getElementById("map-screen"),
  sceneScreen: document.getElementById("scene-screen"),
  tutorialScreen: document.getElementById("tutorial-screen"),
  puzzleScreen: document.getElementById("puzzle-screen"),
  queryScreen: document.getElementById("query-screen"),
  mapMissions: document.getElementById("map-missions"),
  speaker: document.getElementById("speaker-name"),
  text: document.getElementById("dialogue-text"),
  btnContinue: document.getElementById("btn-continue"),
  puzzleTitle: document.getElementById("puzzle-title"),
  instructions: document.getElementById("puzzle-instructions"),
  btnVerify: document.getElementById("btn-verify"),
  btnHint: document.getElementById("btn-hint"),
  hintBox: document.getElementById("hint-box"),
  feedback: document.getElementById("feedback"),
  tutProgress: document.getElementById("tutorial-progress"),
  tutText: document.getElementById("tutorial-text"),
  tutNote: document.getElementById("tutorial-note"),
  btnTutDone: document.getElementById("btn-tutorial-done"),
  btnSkipTut: document.getElementById("btn-skip-tutorial"),
};

function showScreen(which) {
  [el.menuScreen, el.mapScreen, el.sceneScreen, el.tutorialScreen, el.puzzleScreen, el.queryScreen]
    .forEach((s) => s.classList.add("hidden"));
  which.classList.remove("hidden");
}

/* ---------- 3b. Menú de inicio ---------- */

function renderMenu() {
  showScreen(el.menuScreen);
  hidePortrait();
  document.getElementById("menu-backdrop").innerHTML = buildTownBackdrop();
}

document.getElementById("btn-menu-play").addEventListener("click", renderMap);
document.getElementById("btn-map-menu").addEventListener("click", renderMenu);
document.getElementById("btn-back-tutorial").addEventListener("click", renderMap);
document.getElementById("btn-back-puzzle").addEventListener("click", renderMap);
document.getElementById("btn-back-query").addEventListener("click", renderMap);

/* ---------- 5. Motor de diálogo ---------- */

let currentQueue = [];
let queueIndex = 0;
let typing = false;
let typeTimer = null;
let onQueueEnd = null;

function typeText(str) {
  typing = true;
  el.text.textContent = "";
  let i = 0;
  clearInterval(typeTimer);
  typeTimer = setInterval(() => {
    el.text.textContent += str[i];
    i++;
    if (i >= str.length) { clearInterval(typeTimer); typing = false; }
  }, 18);
}

function renderScene() {
  const scene = currentQueue[queueIndex];
  el.speaker.textContent = scene.speaker || "";
  if (scene.portrait) showPortrait(scene.mood); else hidePortrait();
  setBackdrop(scene.bg || (scene.portrait ? "shop" : null));
  typeText(scene.text);
}

function advanceDialogue() {
  if (typing) {
    clearInterval(typeTimer);
    el.text.textContent = currentQueue[queueIndex].text;
    typing = false;
    return;
  }
  queueIndex++;
  if (queueIndex >= currentQueue.length) {
    if (onQueueEnd) onQueueEnd();
    return;
  }
  renderScene();
}

function playScenes(scenes, endCallback) {
  currentQueue = scenes;
  queueIndex = 0;
  onQueueEnd = endCallback;
  showScreen(el.sceneScreen);
  el.btnContinue.classList.remove("hidden");
  renderScene();
}

el.btnContinue.addEventListener("click", advanceDialogue);
document.addEventListener("keydown", (e) => {
  if ((e.key === "Enter" || e.key === " ") && !el.sceneScreen.classList.contains("hidden") && !document.getElementById("choice-wrap")) {
    e.preventDefault();
    advanceDialogue();
  }
});

/* ---------- 6. Clic en celda inserta la referencia (como Excel) ---------- */

let lastClickWasRef = false;
let refDrag = null; /* arrastre para seleccionar un rango dentro de una fórmula */

const REF_AL_FINAL = /[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/;

function getOpenEditorInput(containerId) {
  return document.querySelector(`#${containerId} td.editor input, #${containerId} td.editor textarea`);
}

function cellName(x, y) {
  return jspreadsheet.getColumnNameFromId([x, y]);
}

function tdCoords(e) {
  const td = e.target.closest("td[data-x]");
  if (!td || td.classList.contains("editor")) return null;
  return { x: parseInt(td.getAttribute("data-x"), 10), y: parseInt(td.getAttribute("data-y"), 10) };
}

function attachFormulaClickHelper(containerId) {
  const container = document.getElementById(containerId);

  container.addEventListener(
    "mousedown",
    (e) => {
      const input = getOpenEditorInput(containerId);
      if (!input) return;
      const val = String(input.value || "");
      if (!val.startsWith("=")) return;
      const c = tdCoords(e);
      if (!c) return;
      e.preventDefault();
      e.stopPropagation();
      const ref = cellName(c.x, c.y);
      if (lastClickWasRef && REF_AL_FINAL.test(val)) {
        input.value = val.replace(REF_AL_FINAL, ref);
      } else {
        input.value = val + ref;
      }
      lastClickWasRef = true;
      refDrag = { containerId, ax: c.x, ay: c.y };
      input.focus();
      updateFormulaHelper(containerId);
      mirrorToBar(containerId);
    },
    true
  );

  /* El ayudante de fórmulas y la barra siguen lo que se escribe */
  container.addEventListener("input", () => { updateFormulaHelper(containerId); mirrorToBar(containerId); });
  container.addEventListener("dblclick", () => setTimeout(() => { updateFormulaHelper(containerId); mirrorToBar(containerId); }, 80));

  /* Arrastrar con el mouse apretado agranda la referencia a un rango (E1:E5) */
  container.addEventListener(
    "mousemove",
    (e) => {
      if (!refDrag || refDrag.containerId !== containerId) return;
      const input = getOpenEditorInput(containerId);
      if (!input) { refDrag = null; return; }
      const c = tdCoords(e);
      if (!c) return;
      e.preventDefault();
      e.stopPropagation();
      const x1 = Math.min(refDrag.ax, c.x), x2 = Math.max(refDrag.ax, c.x);
      const y1 = Math.min(refDrag.ay, c.y), y2 = Math.max(refDrag.ay, c.y);
      const ref = (x1 === x2 && y1 === y2)
        ? cellName(x1, y1)
        : cellName(x1, y1) + ":" + cellName(x2, y2);
      input.value = String(input.value).replace(REF_AL_FINAL, ref);
      mirrorToBar(containerId);
    },
    true
  );
}

document.addEventListener("mouseup", () => {
  if (refDrag) {
    const input = getOpenEditorInput(refDrag.containerId);
    if (input) input.focus();
    refDrag = null;
  }
}, true);

document.addEventListener("keydown", () => { lastClickWasRef = false; refDrag = null; }, true);

/* ---------- 6a. Barra de fórmulas y copiar/pegar ---------- */

function getActiveSheetContext() {
  if (!el.tutorialScreen.classList.contains("hidden")) {
    return tutSheet ? { sheet: tutSheet, containerId: "tutorial-sheet", barId: "fbar-tutorial" } : null;
  }
  if (!el.puzzleScreen.classList.contains("hidden")) {
    return spreadsheet ? { sheet: spreadsheet, containerId: "spreadsheet", barId: "fbar-mission" } : null;
  }
  return null;
}

function barFor(containerId) {
  const id = containerId === "tutorial-sheet" ? "fbar-tutorial" : "fbar-mission";
  const bar = document.getElementById(id);
  return { bar, cellEl: bar.querySelector(".fb-cell"), input: bar.querySelector(".fb-input") };
}

const lastSelection = { "spreadsheet": { x: 0, y: 0 }, "tutorial-sheet": { x: 0, y: 0 } };

/* La hoja avisa qué celda está seleccionada -> la barra muestra su contenido crudo */
function onSheetSelection(containerId, sheet, x, y) {
  lastSelection[containerId] = { x, y };
  const { cellEl, input } = barFor(containerId);
  cellEl.textContent = cellName(x, y);
  if (!getOpenEditorInput(containerId) && document.activeElement !== input) {
    input.value = String(sheet.getValueFromCoords(x, y, false) || "");
  }
}

/* Espejo: lo que se escribe en la celda aparece en la barra */
function mirrorToBar(containerId) {
  const editorInput = getOpenEditorInput(containerId);
  if (!editorInput) return;
  const { input } = barFor(containerId);
  if (document.activeElement !== input) input.value = editorInput.value;
}

/* Y al revés: escribir en la barra edita la celda */
function setupFormulaBar(containerId) {
  const { input } = barFor(containerId);

  input.addEventListener("focus", () => {
    const ctx = getActiveSheetContext();
    if (!ctx || ctx.containerId !== containerId) return;
    if (!getOpenEditorInput(containerId)) {
      const sel = lastSelection[containerId];
      const td = document.querySelector(`#${containerId} td[data-x="${sel.x}"][data-y="${sel.y}"]`);
      if (td && !td.classList.contains("readonly")) {
        ctx.sheet.openEditor(td, false);
        const editorInput = getOpenEditorInput(containerId);
        if (editorInput) input.value = editorInput.value;
        input.focus();
      }
    }
  });

  input.addEventListener("input", () => {
    const editorInput = getOpenEditorInput(containerId);
    if (editorInput) {
      editorInput.value = input.value;
      updateFormulaHelper(containerId);
    }
  });

  input.addEventListener("keydown", (e) => {
    const ctx = getActiveSheetContext();
    if (!ctx) return;
    const td = document.querySelector(`#${containerId} td.editor`);
    if (e.key === "Enter" && td) {
      e.preventDefault();
      ctx.sheet.closeEditor(td, true);
      input.blur();
    } else if (e.key === "Escape" && td) {
      e.preventDefault();
      ctx.sheet.closeEditor(td, false);
      input.blur();
    }
  });
}

/* --- Copiar y pegar con ajuste de referencias (como Excel) --- */

let gameClipboard = null;

function shiftRef(letters, num, dx, dy) {
  let x = 0;
  for (let i = 0; i < letters.length; i++) x = x * 26 + (letters.charCodeAt(i) - 64);
  x = x - 1 + dx;
  const y = num - 1 + dy;
  if (x < 0 || y < 0) return letters + num;
  return cellName(x, y);
}

function adjustRefs(raw, dx, dy) {
  if (!String(raw).startsWith("=")) return raw;
  return String(raw).split('"').map((part, i) => {
    if (i % 2 === 1) return part; /* adentro de comillas no se toca */
    return part.replace(/([A-Z]+)([0-9]+)/g, (m, letters, num) => shiftRef(letters, parseInt(num, 10), dx, dy));
  }).join('"');
}

function colIsReadOnly(sheet, x) {
  const col = sheet.options.columns && sheet.options.columns[x];
  return !!(col && col.readOnly);
}

document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const key = e.key.toLowerCase();
  if (key !== "c" && key !== "v") return;
  const ctx = getActiveSheetContext();
  if (!ctx) return;
  if (getOpenEditorInput(ctx.containerId)) return; /* editando adentro de la celda: copia de texto normal */
  const active = document.activeElement;
  if (active && active.classList && active.classList.contains("fb-input")) return;
  const sel = ctx.sheet.selectedCell;
  if (!sel) return;
  const x1 = Math.min(parseInt(sel[0], 10), parseInt(sel[2], 10));
  const y1 = Math.min(parseInt(sel[1], 10), parseInt(sel[3], 10));
  const x2 = Math.max(parseInt(sel[0], 10), parseInt(sel[2], 10));
  const y2 = Math.max(parseInt(sel[1], 10), parseInt(sel[3], 10));

  if (key === "c") {
    gameClipboard = { raw: String(ctx.sheet.getValueFromCoords(x1, y1, false) || ""), x: x1, y: y1 };
    return; /* dejamos que el copiado nativo siga su curso */
  }

  /* pegar */
  if (!gameClipboard) return;
  e.preventDefault();
  e.stopPropagation();
  for (let ty = y1; ty <= y2; ty++) {
    for (let tx = x1; tx <= x2; tx++) {
      if (colIsReadOnly(ctx.sheet, tx)) continue;
      const v = adjustRefs(gameClipboard.raw, tx - gameClipboard.x, ty - gameClipboard.y);
      ctx.sheet.setValueFromCoords(tx, ty, v);
    }
  }
}, true);

/* ---------- 6b. Ayudante de fórmulas (como el de Excel) ---------- */

const FORMULA_DOCS = [
  { name: "SUMA", args: ["rango"], desc: "Suma todos los números de un rango.", ej: "=SUMA(E1:E5)" },
  { name: "PROMEDIO", args: ["rango"], desc: "El promedio de los números de un rango.", ej: "=PROMEDIO(B1:B6)" },
  { name: "MAX", args: ["rango"], desc: "El número más grande del rango.", ej: "=MAX(B1:B6)" },
  { name: "MIN", args: ["rango"], desc: "El número más chico del rango.", ej: "=MIN(B1:B6)" },
  { name: "CONTAR.SI", args: ["rango", "condición"], desc: "Cuenta las celdas que cumplen la condición.", ej: '=CONTAR.SI(B1:B8,"Cassette")' },
  { name: "SUMAR.SI", args: ["rango_condición", "condición", "rango_a_sumar"], desc: "Suma solo las filas que cumplen la condición.", ej: '=SUMAR.SI(B1:B8,"Vinilo",C1:C8)' },
  { name: "BUSCARV", args: ["valor_buscado", "tabla", "columna"], desc: "Busca el valor en la 1ª columna de la tabla y devuelve la columna pedida.", ej: '=BUSCARV("VIN-004",A1:C8,2)' },
  { name: "BUSCARX", args: ["valor_buscado", "matriz_buscada", "matriz_devuelta"], desc: "Como BUSCARV, pero sin restricción de dirección: puede devolver una columna a la izquierda.", ej: '=BUSCARX("VIN-004",B1:B8,A1:A8)' },
  { name: "SI", args: ["condición", "si_verdadero", "si_falso"], desc: "Devuelve un resultado u otro según se cumpla la condición.", ej: '=SI(B1<C1,"REPONER","OK")' },
  { name: "SI.ERROR", args: ["valor", "valor_si_error"], desc: "Si la fórmula da error, devuelve otra cosa en su lugar.", ej: '=SI.ERROR(A1/B1,"—")' },
  { name: "Y", args: ["condición1", "condición2", "..."], desc: "VERDADERO solo si TODAS las condiciones se cumplen.", ej: "=Y(B1<C1,B1<=1)" },
  { name: "O", args: ["condición1", "condición2", "..."], desc: "VERDADERO si AL MENOS UNA condición se cumple.", ej: "=O(B1<C1,D1<C1)" },
  { name: "MAYUSC", args: ["texto"], desc: "Convierte el texto a MAYÚSCULAS.", ej: "=MAYUSC(A1)" },
  { name: "MINUSC", args: ["texto"], desc: "Convierte el texto a minúsculas.", ej: "=MINUSC(A1)" },
  { name: "ESPACIOS", args: ["texto"], desc: "Quita espacios sobrantes al principio, final y en el medio.", ej: "=ESPACIOS(A1)" },
  { name: "IZQUIERDA", args: ["texto", "cantidad"], desc: "Los primeros N caracteres del texto, desde la izquierda.", ej: "=IZQUIERDA(A1,3)" },
  { name: "DERECHA", args: ["texto", "cantidad"], desc: "Los últimos N caracteres del texto, desde la derecha.", ej: "=DERECHA(A1,4)" },
  { name: "CONCATENAR", args: ["texto1", "texto2", "..."], desc: "Une varios textos en uno solo.", ej: '=CONCATENAR(A1,"-",B1)' },
  { name: "AÑO", args: ["fecha"], desc: "El año de una fecha (formato AAAA-MM-DD).", ej: "=AÑO(D1)" },
  { name: "MES", args: ["fecha"], desc: "El mes de una fecha.", ej: "=MES(D1)" },
  { name: "DIA", args: ["fecha"], desc: "El día de una fecha.", ej: "=DIA(D1)" },
];

const helperBox = document.createElement("div");
helperBox.id = "formula-helper";
document.body.appendChild(helperBox);

let helperContainer = null;

function hideFormulaHelper() {
  helperBox.classList.remove("visible");
  helperBox.innerHTML = "";
}

function positionHelper(containerId) {
  const td = document.querySelector(`#${containerId} td.editor`);
  if (!td) return false;
  const r = td.getBoundingClientRect();
  helperBox.style.left = Math.min(r.left, window.innerWidth - 440) + "px";
  helperBox.style.top = (r.bottom + 6) + "px";
  return true;
}

/* Encuentra la función abierta más interna y en qué argumento está el cursor */
function innerOpenFunction(val) {
  const stack = [];
  let inQuotes = false;
  for (let i = 0; i < val.length; i++) {
    const ch = val[i];
    if (ch === '"') { inQuotes = !inQuotes; continue; }
    if (inQuotes) continue;
    if (ch === "(") {
      const m = val.slice(0, i).match(/([A-Za-z][A-Za-z.]*)$/);
      stack.push({ name: m ? m[1].toUpperCase() : null, argIndex: 0 });
    } else if (ch === ")") {
      stack.pop();
    } else if (ch === "," && stack.length) {
      stack[stack.length - 1].argIndex++;
    }
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    if (stack[i].name) return stack[i];
  }
  return null;
}

function updateFormulaHelper(containerId) {
  const input = getOpenEditorInput(containerId);
  if (!input) { hideFormulaHelper(); return; }
  const val = String(input.value || "");
  if (!val.startsWith("=")) { hideFormulaHelper(); return; }
  helperContainer = containerId;

  /* Modo 1: escribiendo un nombre de función (tras =, paréntesis, coma u operador).
     Solo si hay letras escritas, o si recién se puso el "=" — si no, gana la firma. */
  const partial = val.match(/(?:^=|\(|,|\+|\-|\*|\/)\s*([A-Za-z][A-Za-z.]*)?$/);
  const dentroDeComillas = (val.split('"').length - 1) % 2 === 1;
  if (partial && !dentroDeComillas && ((partial[1] || "") !== "" || val.replace(/\s/g, "") === "=")) {
    const pref = (partial[1] || "").toUpperCase();
    const matches = FORMULA_DOCS.filter((d) => d.name.startsWith(pref));
    if (matches.length) {
      helperBox.innerHTML = matches.map((d, i) =>
        `<div class="fh-item" data-fn="${d.name}">
           <span class="fh-name">${d.name}</span><span class="fh-sig-args">(${d.args.join(", ")})</span><br>
           <span class="fh-desc">${d.desc}</span>
         </div>`).join("");
      helperBox.querySelectorAll(".fh-item").forEach((item) => {
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          e.stopPropagation();
          const inp = getOpenEditorInput(containerId);
          if (!inp) return;
          inp.value = String(inp.value).replace(/[A-Za-z.]*$/, item.getAttribute("data-fn") + "(");
          inp.focus();
          updateFormulaHelper(containerId);
        });
      });
      if (positionHelper(containerId)) helperBox.classList.add("visible");
      return;
    }
  }

  /* Modo 2: adentro de una función conocida — mostrar firma con el argumento activo */
  const open = innerOpenFunction(val);
  if (open) {
    const doc = FORMULA_DOCS.find((d) => d.name === open.name);
    if (doc) {
      const args = doc.args.map((a, i) =>
        i === Math.min(open.argIndex, doc.args.length - 1)
          ? `<span class="fh-arg-activo">${a}</span>` : a).join(", ");
      helperBox.innerHTML = `<div class="fh-signature">
          <span class="fh-name">${doc.name}</span>(${args})<br>
          <span class="fh-desc">${doc.desc}</span>
          <span class="fh-ej">Ej: ${doc.ej}</span>
        </div>`;
      if (positionHelper(containerId)) helperBox.classList.add("visible");
      return;
    }
  }
  hideFormulaHelper();
}

/* Si el editor se cerró (Enter, Escape, clic afuera), ocultar el ayudante */
function syncFormulaHelper() {
  setTimeout(() => {
    if (!helperContainer || !getOpenEditorInput(helperContainer)) hideFormulaHelper();
  }, 80);
}
document.addEventListener("keydown", syncFormulaHelper, true);
document.addEventListener("mousedown", (e) => {
  if (!helperBox.contains(e.target)) syncFormulaHelper();
}, true);
attachFormulaClickHelper("spreadsheet");
attachFormulaClickHelper("tutorial-sheet");
setupFormulaBar("spreadsheet");
setupFormulaBar("tutorial-sheet");

/* ---------- 7. Mapa de misiones ---------- */

function renderMap() {
  showScreen(el.mapScreen);
  hidePortrait();
  document.getElementById("map-backdrop").innerHTML = buildTownBackdrop();
  el.mapMissions.innerHTML = "";

  const tutCard = document.createElement("div");
  tutCard.className = "mission-card";
  tutCard.innerHTML = `<div class="mission-status">🎓</div>
    <div class="mission-info">
      <div class="mission-name">ENTRENAMIENTO</div>
      <div class="mission-concepts">Celdas, filas, columnas y tu primera fórmula (repetible)</div>
    </div>`;
  tutCard.onclick = () => startTutorial(renderMap);
  el.mapMissions.appendChild(tutCard);

  MISSIONS.forEach((m, i) => {
    const done = isDone(m.id);
    const unlocked = i === 0 || isDone(MISSIONS[i - 1].id);
    const card = document.createElement("div");
    card.className = "mission-card" + (done ? " done" : "") + (!unlocked ? " locked" : "");
    card.innerHTML = `<div class="mission-status">${done ? "✔" : unlocked ? "▶" : "🔒"}</div>
      <div class="mission-info">
        <div class="mission-name">${m.name}</div>
        <div class="mission-concepts">${m.concepts}</div>
      </div>`;
    if (unlocked) card.onclick = () => startMission(m);
    el.mapMissions.appendChild(card);
  });

  COMING_SOON.forEach((c) => {
    const card = document.createElement("div");
    card.className = "mission-card locked";
    card.innerHTML = `<div class="mission-status">🔒</div>
      <div class="mission-info">
        <div class="mission-name">${c.name}</div>
        <div class="mission-concepts">${c.concepts} — próximamente</div>
      </div>`;
    el.mapMissions.appendChild(card);
  });
}

/* ---------- 8. Tutorial interactivo ---------- */

let tutSheet = null;
let tutStep = 0;
let tutExitTo = null;

const TUTORIAL_STEPS = [
  {
    text: "Esto de acá abajo es una HOJA DE CÁLCULO. Las FILAS son los números (1, 2, 3...) y las COLUMNAS son las letras (A, B, C...). Cada casillero se llama CELDA, y tiene nombre: la celda B2 está en la columna B, fila 2.",
    task: "👉 Hacé clic en la celda B2.",
  },
  {
    text: "¡Esa es! Ahora vamos a escribir. Para escribir en una celda: doble clic sobre ella, escribís, y apretás Enter.",
    task: "👉 Hacé doble clic en la celda A1, escribí tu nombre y apretá Enter.",
  },
  {
    text: "Las celdas guardan distintos TIPOS DE DATOS: texto (como tu nombre) o números. Con los números se pueden hacer cuentas.",
    task: "👉 Escribí el número 5 en la celda B1, y el número 3 en la celda B2.",
  },
  {
    text: "Ahora la magia: las FÓRMULAS. Una fórmula siempre empieza con el signo = y hace cuentas por vos.",
    task: "👉 En la celda B3 escribí =B1+B2 y apretá Enter. (Tip: escribí = y después hacé clic en B1, escribí +, clic en B2, Enter).",
  },
  {
    text: "Mirá: B3 dice 8. Pero lo importante no es el 8... es que la celda guarda la FÓRMULA. Probalo:",
    task: "👉 Cambiá el 5 de B1 por un 10 (doble clic, borrá, escribí 10, Enter). Mirá lo que pasa con B3.",
  },
];

function renderTutorialStep(note, good) {
  const s = TUTORIAL_STEPS[tutStep];
  el.tutProgress.textContent = `PASO ${tutStep + 1} DE ${TUTORIAL_STEPS.length}`;
  el.tutText.innerHTML = `${s.text}<br><b>${s.task}</b>`;
  el.tutNote.textContent = note || "";
  el.tutNote.className = good ? "good" : "";
}

function tutAdvance(msg) {
  tutStep++;
  if (tutStep >= TUTORIAL_STEPS.length) {
    el.tutProgress.textContent = "ENTRENAMIENTO COMPLETO ✔";
    el.tutText.innerHTML = "¿Viste? B3 cambió solo a 13. La fórmula se recalcula sola cuando cambian los datos — <b>por eso en Excel usamos fórmulas y no cuentas a mano</b>.";
    el.tutNote.textContent = "";
    el.btnTutDone.classList.remove("hidden");
    el.btnSkipTut.classList.add("hidden");
    return;
  }
  renderTutorialStep(msg, true);
}

function tutNum(x, y) {
  const v = parseFloat(tutSheet.getValueFromCoords(x, y, true));
  return isNaN(v) ? null : v;
}

function evaluateTutorial() {
  if (!tutSheet) return;
  if (tutStep === 1) {
    const a1 = String(tutSheet.getValueFromCoords(0, 0, true) || "").trim();
    if (a1 !== "") tutAdvance(`¡Hola, ${a1}! Eso que escribiste es un dato de tipo TEXTO.`);
  } else if (tutStep === 2) {
    if (tutNum(1, 0) === 5 && tutNum(1, 1) === 3) tutAdvance("¡Bien! Dos números listos para hacer cuentas.");
  } else if (tutStep === 3) {
    const raw = String(tutSheet.getValueFromCoords(1, 2, false) || "");
    if (tutNum(1, 2) === 8) {
      if (raw.startsWith("=")) tutAdvance("¡Fórmula perfecta! =B1+B2 sumó las dos celdas.");
      else el.tutNote.textContent = "El 8 está bien... pero lo escribiste a mano. Borralo y probá con la fórmula: =B1+B2";
    }
  } else if (tutStep === 4) {
    if (tutNum(1, 0) === 10 && tutNum(1, 2) === 13) tutAdvance();
  }
}

function startTutorial(exitTo) {
  tutExitTo = exitTo;
  showScreen(el.tutorialScreen);
  tutStep = 0;
  el.btnTutDone.classList.add("hidden");
  el.btnSkipTut.classList.remove("hidden");
  renderTutorialStep();

  document.getElementById("tutorial-sheet").innerHTML = "";
  tutSheet = jspreadsheet(document.getElementById("tutorial-sheet"), {
    data: [["", "", "", ""], ["", "", "", ""], ["", "", "", ""], ["", "", "", ""]],
    minDimensions: [4, 4],
    defaultColWidth: 120,
    allowInsertRow: false, allowInsertColumn: false,
    allowDeleteRow: false, allowDeleteColumn: false,
    columnSorting: false,
    autoCasting: false,
    onselection: (instance, x1, y1) => {
      onSheetSelection("tutorial-sheet", tutSheet, parseInt(x1, 10), parseInt(y1, 10));
      if (tutStep === 0 && parseInt(x1, 10) === 1 && parseInt(y1, 10) === 1) {
        tutAdvance("¡Justo ahí! Columna B, fila 2 = celda B2.");
      }
    },
    onchange: () => { setTimeout(evaluateTutorial, 50); },
  });
}

el.btnTutDone.addEventListener("click", () => { if (tutExitTo) tutExitTo(); });
el.btnSkipTut.addEventListener("click", () => { if (tutExitTo) tutExitTo(); });

/* ---------- 9. Motor de misión ---------- */

let spreadsheet = null;
let mission = null;
let stageIndex = 0;
let hintIndex = 0;

function numAt(s, x, y) {
  const raw = s.getValueFromCoords(x, y, true);
  return parseFloat(String(raw).replace(/[^\d.\-]/g, ""));
}
function rawAt(s, x, y) {
  return String(s.getValueFromCoords(x, y, false) || "");
}

function startMission(m) {
  mission = m;
  const startWhatever = m.type === "query" ? () => startQueryMission(m) : () => startPuzzle(m);
  const goBriefing = () => playScenes(m.briefing, startWhatever);
  if (m.intro && !isDone(m.id)) {
    playScenes(m.intro, () => {
      if (m.offerTutorial) showTutorialChoice(goBriefing);
      else goBriefing();
    });
  } else {
    goBriefing();
  }
}

function showTutorialChoice(next) {
  el.speaker.textContent = "MARTA";
  showPortrait();
  setBackdrop("shop");
  el.text.textContent = "Antes de empezar... ¿usaste alguna vez una planilla de cálculo?";
  el.btnContinue.classList.add("hidden");

  const wrap = document.createElement("div");
  wrap.id = "choice-wrap";
  wrap.style.cssText = "display:flex;gap:14px;margin-top:14px;flex-wrap:wrap;";

  const btnNo = document.createElement("button");
  btnNo.className = "pixel-btn";
  btnNo.textContent = "nunca — enseñame";
  btnNo.onclick = () => { wrap.remove(); startTutorial(next); };

  const btnYes = document.createElement("button");
  btnYes.className = "pixel-btn ghost";
  btnYes.textContent = "ya sé lo básico";
  btnYes.onclick = () => { wrap.remove(); next(); };

  wrap.appendChild(btnNo);
  wrap.appendChild(btnYes);
  document.getElementById("dialogue-box").appendChild(wrap);
}

function startPuzzle(m) {
  showScreen(el.puzzleScreen);
  stageIndex = 0;
  resetHints();
  el.feedback.textContent = "";
  el.feedback.className = "";
  el.btnVerify.disabled = false;
  el.puzzleTitle.textContent = m.title;
  el.instructions.innerHTML = m.stages[0].instructions;

  document.getElementById("spreadsheet").innerHTML = "";
  spreadsheet = jspreadsheet(document.getElementById("spreadsheet"), {
    data: m.sheet.data.map((r) => r.slice()),
    columns: m.sheet.columns,
    minDimensions: [m.sheet.columns.length, m.sheet.data.length],
    allowInsertRow: false, allowInsertColumn: false,
    allowDeleteRow: false, allowDeleteColumn: false,
    columnSorting: false,
    autoCasting: false,
    onselection: (instance, x1, y1) => {
      onSheetSelection("spreadsheet", spreadsheet, parseInt(x1, 10), parseInt(y1, 10));
    },
  });
}

function restoreProtected() {
  if (!mission.protect) return;
  mission.protect.forEach((p) => {
    const current = String(spreadsheet.getValueFromCoords(p.x, p.y, true) || "").trim();
    if (current !== String(p.v)) spreadsheet.setValueFromCoords(p.x, p.y, p.v);
  });
}

function currentStage() { return mission.stages[stageIndex]; }

function resetHints() {
  hintIndex = 0;
  el.hintBox.classList.add("hidden");
  el.hintBox.innerHTML = "";
}

function showNextHint() {
  const hints = currentStage().hints;
  if (hintIndex < hints.length) hintIndex++;
  const shown = hints.slice(0, hintIndex)
    .map((h) => `<div style="margin-bottom:8px;">💡 ${h}</div>`).join("");
  el.hintBox.innerHTML = `<span class="hint-count">PISTA ${hintIndex} DE ${hints.length}</span>${shown}`;
  el.hintBox.classList.remove("hidden");
}

el.btnHint.addEventListener("click", showNextHint);

el.btnVerify.addEventListener("click", () => {
  restoreProtected();
  const result = currentStage().check(spreadsheet);
  el.feedback.textContent = result.feedback;
  el.feedback.className = result.ok ? "ok" : "bad";

  if (!result.ok) return;

  if (stageIndex < mission.stages.length - 1) {
    stageIndex++;
    resetHints();
    el.instructions.innerHTML = currentStage().instructions;
  } else {
    el.btnVerify.disabled = true;
    markDone(mission.id);
    setTimeout(() => { playScenes(mission.outro, renderMap); }, 1600);
  }
});

/* ---------- 9b. Motor de misión tipo "consulta" (Power Query simulado) ---------- */

let queryStageIndex = 0;
let queryAppliedSteps = [];
let queryHintIndex = 0;

function qEl(id) { return document.getElementById(id); }

function rowsToRaw(strings) {
  return strings.map((s) => ({ producto: s }));
}

function computeQueryRows() {
  let rows = rowsToRaw(mission.query.raw);
  queryAppliedSteps.forEach((stepId) => {
    const step = mission.query.steps.find((s) => s.id === stepId);
    if (step) rows = step.fn(rows);
  });
  return rows;
}

function renderQueryRawTable() {
  qEl("query-raw-table").innerHTML =
    `<tr><th>Producto (crudo)</th></tr>` +
    mission.query.raw.map((r) => `<tr><td>"${r}"</td></tr>`).join("");
}

function renderQueryPreview() {
  const rows = computeQueryRows();
  const hasPrice = rows.some((r) => r.precio !== undefined);
  const header = hasPrice ? "<tr><th>Producto</th><th>Precio</th></tr>" : "<tr><th>Producto</th></tr>";
  const body = rows.length
    ? rows.map((r) => hasPrice
        ? `<tr><td>${r.producto}</td><td>${r.precio === null || r.precio === undefined ? "?" : "$" + r.precio}</td></tr>`
        : `<tr><td>${r.producto}</td></tr>`).join("")
    : `<tr><td class="query-empty">(sin filas)</td></tr>`;
  qEl("query-preview-table").innerHTML = header + body;
}

function renderQueryAvailableList() {
  const wrap = qEl("query-available-list");
  wrap.innerHTML = "";
  mission.query.steps.forEach((step) => {
    const btn = document.createElement("button");
    btn.className = "query-step-btn";
    btn.textContent = "+ " + step.label;
    btn.onclick = () => addQueryStep(step.id);
    wrap.appendChild(btn);
  });
}

function renderQueryAppliedList() {
  const wrap = qEl("query-applied-list");
  wrap.innerHTML = "";
  if (!queryAppliedSteps.length) {
    wrap.innerHTML = `<span style="color:var(--text-dim);">(todavía ningún paso)</span>`;
    return;
  }
  queryAppliedSteps.forEach((stepId, i) => {
    const step = mission.query.steps.find((s) => s.id === stepId);
    const div = document.createElement("div");
    div.className = "query-step-applied";
    div.innerHTML = `<span class="query-step-num">${i + 1}</span>${step ? step.label : stepId}`;
    wrap.appendChild(div);
  });
}

function addQueryStep(stepId) {
  queryAppliedSteps.push(stepId);
  renderQueryAppliedList();
  renderQueryPreview();
}

function undoQueryStep() {
  queryAppliedSteps.pop();
  renderQueryAppliedList();
  renderQueryPreview();
}

document.getElementById("btn-query-undo").addEventListener("click", undoQueryStep);

function currentQueryStage() { return mission.query.stages[queryStageIndex]; }

function resetQueryHints() {
  queryHintIndex = 0;
  qEl("query-hint-box").classList.add("hidden");
  qEl("query-hint-box").innerHTML = "";
}

function showNextQueryHint() {
  const hints = currentQueryStage().hints;
  if (queryHintIndex < hints.length) queryHintIndex++;
  const shown = hints.slice(0, queryHintIndex)
    .map((h) => `<div style="margin-bottom:8px;">💡 ${h}</div>`).join("");
  qEl("query-hint-box").innerHTML = `<span class="hint-count">PISTA ${queryHintIndex} DE ${hints.length}</span>${shown}`;
  qEl("query-hint-box").classList.remove("hidden");
}

document.getElementById("btn-query-hint").addEventListener("click", showNextQueryHint);

function startQueryMission(m) {
  showScreen(el.queryScreen);
  queryStageIndex = 0;
  queryAppliedSteps = [];
  resetQueryHints();
  qEl("query-feedback").textContent = "";
  qEl("query-feedback").className = "";
  qEl("btn-query-verify").disabled = false;
  qEl("query-title").textContent = m.title;
  qEl("query-instructions").innerHTML = m.query.stages[0].instructions;
  renderQueryRawTable();
  renderQueryAvailableList();
  renderQueryAppliedList();
  renderQueryPreview();
}

document.getElementById("btn-query-verify").addEventListener("click", () => {
  const rows = computeQueryRows();
  const result = currentQueryStage().check(rows);
  const fb = qEl("query-feedback");
  fb.textContent = result.feedback;
  fb.className = result.ok ? "ok" : "bad";
  if (!result.ok) return;

  if (queryStageIndex < mission.query.stages.length - 1) {
    queryStageIndex++;
    resetQueryHints();
    qEl("query-instructions").innerHTML = currentQueryStage().instructions;
  } else {
    qEl("btn-query-verify").disabled = true;
    markDone(mission.id);
    setTimeout(() => { playScenes(mission.outro, renderMap); }, 1600);
  }
});

/* ---------- 10. Arranque ---------- */

renderMenu();
