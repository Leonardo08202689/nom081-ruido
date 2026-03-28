/**
 * generar_memoria.js
 * Genera el Word de Memoria de Cálculo NOM-081 fiel al estilo del documento original.
 * Uso: node generar_memoria.js calc_result.json output.docx
 */

const fs   = require("fs");
const path = require("path");
const docx = require("/home/claude/.npm-global/lib/node_modules/docx");

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak,
  LevelFormat, UnderlineType,
} = docx;

// ── Cargar datos ──────────────────────────────────────────────────────────────
const data   = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const M      = data.metadata;
const PF     = data.fuente_stats;   // {A:{N50,sigma,N10,Neq,suma}, ...}
const PB     = data.fondo_stats;
const PROM   = data.promedios;
const CORR   = data.correcciones;
const RES    = data.resultado;
const FD     = data.fuente_data;    // raw readings per period
const BD     = data.fondo_data;

// ── Constantes de formato ─────────────────────────────────────────────────────
const FONT        = "Arial Nova Light";
const FONT_FALL   = "Arial";
const SZ          = 20;            // 10pt en half-points
const SZ_TITLE    = 24;            // 12pt
const SZ_SECTION  = 22;            // 11pt
const COLOR_BLACK  = "000000";
const COLOR_GREEN  = "E2EFD9";     // fondo título principal
const COLOR_HEADER = "C6E0B4";     // fondo encabezados de tabla
const COLOR_BLUE_H = "D9E1F2";     // azul claro cabecera secundaria
const COLOR_GREY   = "F2F2F2";
const COLOR_ACCENT = "375623";     // verde oscuro texto sección
const PAGE_W       = 12240;        // Letter 8.5"
const PAGE_H       = 15840;        // Letter 11"
const MARGIN       = 1080;         // 0.75" margins
const CONTENT_W    = PAGE_W - 2 * MARGIN; // 10080 DXA

// ── Helpers ───────────────────────────────────────────────────────────────────

function run(text, opts = {}) {
  return new TextRun({
    text,
    font:  { name: FONT },
    size:  opts.size  || SZ,
    bold:  opts.bold  || false,
    color: opts.color || COLOR_BLACK,
    italics: opts.italic || false,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 0,
               line: opts.line || 240, lineRule: "auto" },
    indent: opts.indent,
    children: Array.isArray(children) ? children : [children],
  });
}

function emptyPara(h = 60) {
  return new Paragraph({
    spacing: { before: h, after: 0 },
    children: [run("")],
  });
}

const BORDER_THIN   = { style: BorderStyle.SINGLE, size: 4,  color: "000000" };
const BORDER_THICK  = { style: BorderStyle.SINGLE, size: 8,  color: "000000" };
const BORDER_NIL    = { style: BorderStyle.NIL };
const borders_all   = { top: BORDER_THIN, bottom: BORDER_THIN,
                        left: BORDER_THIN, right: BORDER_THIN };
const borders_thick = { top: BORDER_THICK, bottom: BORDER_THICK,
                        left: BORDER_THICK, right: BORDER_THICK };

function shade(fill) {
  return { fill, type: ShadingType.CLEAR, color: "auto" };
}

function cell(content, opts = {}) {
  const children = Array.isArray(content) ? content
    : [para(
        Array.isArray(content) ? content : [run(String(content), {
          bold: opts.bold, size: opts.size || SZ, color: opts.color || COLOR_BLACK,
        })],
        { align: opts.align || AlignmentType.CENTER,
          before: 60, after: 0 }
      )];
  return new TableCell({
    width:      opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    columnSpan: opts.span,
    rowSpan:    opts.rowSpan,
    verticalAlign: opts.vAlign || VerticalAlign.CENTER,
    borders:    opts.borders || borders_all,
    shading:    opts.fill ? shade(opts.fill) : undefined,
    margins:    { top: 60, bottom: 60, left: 100, right: 100 },
    children,
  });
}

function labelCell(text, w, bold=true) {
  return cell([run(text, { bold, size: SZ })],
    { width: w, align: AlignmentType.RIGHT,
      borders: borders_all, fill: COLOR_GREY });
}
function valueCell(text, w, fill) {
  return cell([run(String(text), { size: SZ })],
    { width: w, align: AlignmentType.LEFT, fill });
}

// ── TÍTULO PRINCIPAL ──────────────────────────────────────────────────────────
function titlePara() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    shading: shade(COLOR_GREEN),
    spacing: { before: 60, after: 60, line: 240 },
    children: [run("CÁLCULOS PARA ZONA CRITICA 1", { bold: true, size: SZ_TITLE })],
  });
}

// ── TABLA ENCABEZADO (metadata) ────────────────────────────────────────────────
function headerTable() {
  const W = CONTENT_W;
  const c1 = 1700, c2 = 5600, c3 = W - c1 - c2;  // ~2780

  function hCell(texts, w, opts={}) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      columnSpan: opts.span,
      rowSpan:    opts.rowSpan,
      verticalAlign: VerticalAlign.CENTER,
      borders: opts.borders || borders_all,
      shading: opts.fill ? shade(opts.fill) : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: texts,
    });
  }

  function labelP(text) {
    return para([run(text, { bold: true, size: SZ })],
      { align: AlignmentType.RIGHT, before: 40, after: 0 });
  }
  function valP(text) {
    return para([run(text, { size: SZ })],
      { align: AlignmentType.LEFT, before: 40, after: 0 });
  }

  const evaluadoresLines = M.evaluadores.split("\n");

  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [c1, c2, c3],
    rows: [
      // Row 1 – Compañía | value | Fecha muestreo
      new TableRow({ children: [
        hCell([labelP("Compañía:")], c1),
        hCell([valP(M.compania)],    c2),
        hCell([para([run(" Fecha de muestreo: ", { bold: true, size: SZ })],
               { align: AlignmentType.LEFT, before: 40 })],
               c3, { rowSpan: 2, borders: borders_all }),
      ]}),
      // Row 2 – Ubicación
      new TableRow({ children: [
        hCell([labelP("Ubicación:")], c1,
          { borders: { top: BORDER_NIL, bottom: BORDER_NIL,
                       left: BORDER_THIN, right: BORDER_NIL }}),
        hCell([valP(M.ubicacion)],   c2,
          { borders: { top: BORDER_NIL, bottom: BORDER_NIL,
                       left: BORDER_NIL, right: BORDER_NIL }}),
      ]}),
      // Row 3 – Evaluadores + Fecha valor
      new TableRow({ children: [
        hCell([labelP("Evaluadores:")], c1),
        hCell(evaluadoresLines.map(t => valP(t)), c2),
        hCell([para([run(M.fecha, { size: SZ })],
               { align: AlignmentType.CENTER, before: 40 })], c3),
      ]}),
      // Row 4 – empty separator
      new TableRow({ children: [
        hCell([para([run("")], { before: 20 })], c1,
          { borders: { top: BORDER_NIL, bottom: BORDER_THIN,
                       left: BORDER_THIN, right: BORDER_NIL }}),
        hCell([para([run("")], { before: 20 })], c2,
          { borders: { top: BORDER_NIL, bottom: BORDER_THIN,
                       left: BORDER_NIL, right: BORDER_NIL }}),
        hCell([para([run("")], { before: 20 })], c3,
          { borders: { top: BORDER_NIL, bottom: BORDER_THIN,
                       left: BORDER_NIL, right: BORDER_THIN }}),
      ]}),
      // Row 5 – Zona / Evaluación / Hora inicio
      new TableRow({ children: [
        hCell([para([run("Zona Crítica", { bold: true, size: SZ })],
               { align: AlignmentType.LEFT, before: 40 })],
               c1, { rowSpan: 2 }),
        hCell([para([run(M.zona, { size: SZ })],
               { align: AlignmentType.LEFT, before: 40 })], c2,
               { rowSpan: 2 }),
        hCell([
          para([run("Evaluación: ", { bold: true, size: SZ }),
                run(M.evaluacion, { size: SZ })],
               { before: 40 }),
          para([run("Hora inicio: ", { bold: true, size: SZ }),
                run(M.hora_inicio, { size: SZ })],
               { before: 20 }),
          para([run("Hora final:  ", { bold: true, size: SZ }),
                run(M.hora_final,  { size: SZ })],
               { before: 20, after: 40 }),
        ], c3),
      ]}),
    ],
  });
}

// ── TABLA DE DATOS DE CAMPO ────────────────────────────────────────────────────
function datosTable(tipo, periodos, rawData, stats) {
  const W   = CONTENT_W;
  const cL  = 1100;  // col "Número de lectura"
  const cP  = Math.floor((W - cL) / periodos.length);  // col por periodo
  const cR  = W - cL - cP * (periodos.length - 1);

  const colWidths = [cL, ...periodos.map((_, i) =>
    i === periodos.length - 1 ? cR : cP)];

  const rows = [];

  // ── Título bloque ──
  rows.push(new TableRow({ children: [
    new TableCell({
      columnSpan: periodos.length + 1,
      width: { size: W, type: WidthType.DXA },
      shading: shade(tipo === "FUENTE" ? COLOR_GREEN : COLOR_BLUE_H),
      borders: borders_thick,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run(`DATOS DE CAMPO — RUIDO DE ${tipo}`, { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
  ]}));

  // ── Encabezado periodos ──
  rows.push(new TableRow({ children: [
    new TableCell({
      width: { size: cL, type: WidthType.DXA },
      rowSpan: 2, verticalAlign: VerticalAlign.CENTER,
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run("N° de\nlectura", { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
    new TableCell({
      columnSpan: periodos.length,
      width: { size: W - cL, type: WidthType.DXA },
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run("Periodo", { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
  ]}));

  // ── Sub-encabezado letras/números de periodo ──
  rows.push(new TableRow({ children: periodos.map((p, i) =>
    new TableCell({
      width: { size: colWidths[i + 1], type: WidthType.DXA },
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run(p, { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    })
  )}));

  // ── 35 filas de lecturas ──
  for (let i = 0; i < 35; i++) {
    const rowFill = i % 2 === 0 ? "FFFFFF" : COLOR_GREY;
    rows.push(new TableRow({ children: [
      new TableCell({
        width: { size: cL, type: WidthType.DXA },
        shading: shade(COLOR_HEADER), borders: borders_all,
        margins: { top: 40, bottom: 40, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [para([run(String(i + 1), { size: SZ })],
                   { align: AlignmentType.CENTER })],
      }),
      ...periodos.map((p, pi) => new TableCell({
        width: { size: colWidths[pi + 1], type: WidthType.DXA },
        shading: shade(rowFill), borders: borders_all,
        margins: { top: 40, bottom: 40, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [para([run(String(rawData[p][i]), { size: SZ })],
                   { align: AlignmentType.CENTER })],
      })),
    ]}));
  }

  // ── Fila sumatoria ──
  rows.push(new TableRow({ children: [
    new TableCell({
      width: { size: cL, type: WidthType.DXA },
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run("Sumatoria", { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
    ...periodos.map((p, pi) => new TableCell({
      width: { size: colWidths[pi + 1], type: WidthType.DXA },
      shading: shade(COLOR_GREEN), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run(String(stats[p].suma.toFixed(2)), { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    })),
  ]}));

  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows,
  });
}

// ── TABLA DE RESULTADOS POR PERÍODO ──────────────────────────────────────────
function resultadosPeriodoTable(tipo, periodos, stats) {
  const W  = CONTENT_W;
  const c0 = 1800;
  const cP = Math.floor((W - c0) / periodos.length);
  const cL = W - c0 - cP * (periodos.length - 1);
  const colWidths = [c0, ...periodos.map((_,i) => i===periodos.length-1?cL:cP)];

  const indicadores = [
    ["N₅₀ (dB)",  "N50"],
    ["σ (dB)",    "sigma"],
    ["N₁₀ (dB)", "N10"],
    ["Neq (dB)",  "Neq"],
  ];

  const rows = [];

  // Título
  rows.push(new TableRow({ children: [
    new TableCell({
      columnSpan: periodos.length + 1,
      shading: shade(tipo === "FUENTE" ? COLOR_GREEN : COLOR_BLUE_H),
      borders: borders_thick,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run(`RESULTADOS POR PERÍODO — RUIDO DE ${tipo}`,
                       { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
  ]}));

  // Encabezado
  rows.push(new TableRow({ children: [
    new TableCell({
      width: { size: c0, type: WidthType.DXA },
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run("Indicador", { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    }),
    ...periodos.map((p, pi) => new TableCell({
      width: { size: colWidths[pi+1], type: WidthType.DXA },
      shading: shade(COLOR_HEADER), borders: borders_all,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run(p, { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })],
    })),
  ]}));

  // Filas de indicadores
  indicadores.forEach(([label, key], idx) => {
    const fill = idx % 2 === 0 ? "FFFFFF" : COLOR_GREY;
    rows.push(new TableRow({ children: [
      new TableCell({
        width: { size: c0, type: WidthType.DXA },
        shading: shade(COLOR_HEADER), borders: borders_all,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [para([run(label, { bold: true, size: SZ })],
                   { align: AlignmentType.LEFT })],
      }),
      ...periodos.map((p, pi) => new TableCell({
        width: { size: colWidths[pi+1], type: WidthType.DXA },
        shading: shade(fill), borders: borders_all,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [para([run(stats[p][key].toFixed(2), { size: SZ })],
                   { align: AlignmentType.CENTER })],
      })),
    ]}));
  });

  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows,
  });
}

// ── SECCIÓN: Fórmulas explicadas ──────────────────────────────────────────────
function seccionFormulas() {
  function sectionTitle(letter, title) {
    return new Paragraph({
      spacing: { before: 200, after: 80, line: 240 },
      children: [
        run(`${letter}.  `, { bold: true, size: SZ_SECTION }),
        run(title,          { bold: true, size: SZ_SECTION }),
      ],
    });
  }

  function formula(text) {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
      children: [run(text, { size: SZ, italic: true })],
    });
  }

  function explanation(text) {
    return new Paragraph({
      spacing: { before: 60, after: 40 },
      indent: { left: 360 },
      children: [run(text, { size: SZ })],
    });
  }

  const perF = ["A","B","C","D","E"];
  const perB = ["I","II","III","IV","V"];

  return [
    // ── A. Fórmulas aplicables ──────────────────────────────────
    sectionTitle("A", "Fórmulas aplicables"),
    para([run("Para cada periodo de 35 lecturas (Nᵢ) en dB se calculan:", { size: SZ })],
         { before: 60 }),
    formula("N₅₀ = (1/n) × Σ Nᵢ"),
    explanation("N₅₀: Nivel sonoro medio aritmético de las 35 lecturas del periodo."),
    formula("σ = √[ Σ(Nᵢ − N₅₀)² / (n−1) ]"),
    explanation("σ: Desviación estándar muestral (ddof = 1, n = 35)."),
    formula("N₁₀ = N₅₀ + 1.2817 × σ"),
    explanation("N₁₀: Nivel superado el 10 % del tiempo (percentil 90). El factor 1.2817 es el valor z de la distribución normal estándar para P = 90 %."),
    formula("Neq = 10 × log₁₀[ (1/n) × Σ 10^(Nᵢ/10) ]"),
    explanation("Neq: Nivel de presión sonora equivalente (promedio energético de las 35 lecturas)."),
    emptyPara(60),
    para([run("Promedios globales de los 5 periodos:", { size: SZ, bold: true })],
         { before: 60 }),
    formula("N₅₀_prom = (1/5) × Σ N₅₀ᵢ"),
    formula("σ_prom   = (1/5) × Σ σᵢ"),
    formula("N₁₀_prom = N₅₀_prom + 1.2817 × σ_prom"),
    formula("(Neq)eq  = 10 × log₁₀[ (1/5) × Σ 10^(Neqᵢ/10) ]"),
    explanation("(Neq)eq: promedio energético de los Neq de los 5 periodos."),

    // ── B. Sustitución ──────────────────────────────────────────
    sectionTitle("B", "Sustitución"),
    para([run("Aplicando las fórmulas a los datos medidos:", { size: SZ })],
         { before: 60 }),
    emptyPara(40),

    // Tablas de resultados por período
    resultadosPeriodoTable("FUENTE", perF, PF),
    emptyPara(120),
    resultadosPeriodoTable("FONDO",  perB, PB),
    emptyPara(120),

    // Tabla de promedios
    promediosTable(),
    emptyPara(80),

    // ── C. Cálculo de correcciones ──────────────────────────────
    sectionTitle("C", "Cálculo de correcciones"),
    para([run("Una vez obtenidos los promedios, se aplican las correcciones de la norma:", { size: SZ })],
         { before: 60 }),
    emptyPara(40),

    para([
      run("Ce = 0.9023 × σ_prom(fuente)  = 0.9023 × " +
          PROM.fuente.sigma.toFixed(2) + " = ", { size: SZ }),
      run(CORR.Ce.toFixed(2) + " dB", { size: SZ, bold: true }),
    ], { before: 60, indent: { left: 360 } }),
    explanation("Ce: Corrección por la presencia de valores extremos en la distribución de niveles."),

    para([
      run("Δ₅₀ = N₅₀_fuente − N₅₀_fondo  = " +
          PROM.fuente.N50.toFixed(2) + " − " + PROM.fondo.N50.toFixed(2) + " = ", { size: SZ }),
      run(CORR.delta50.toFixed(2) + " dB", { size: SZ, bold: true }),
    ], { before: 60, indent: { left: 360 } }),
    explanation("Δ₅₀: Diferencia entre el nivel medio de la fuente y el fondo. Determina si aplica la corrección por ruido de fondo."),

    para([
      run("N'₅₀ = N₅₀_fuente + Ce  = " +
          PROM.fuente.N50.toFixed(2) + " + " + CORR.Ce.toFixed(2) + " = ", { size: SZ }),
      run(CORR.N50_corr.toFixed(2) + " dB", { size: SZ, bold: true }),
    ], { before: 60, indent: { left: 360 } }),

    ...corrTabla(),

    // ── D. Determinación de Nff ──────────────────────────────────
    sectionTitle("D", "Determinación del nivel de fuente fija (Nff)"),
    para([run("El N_ff es el mayor entre N'₅₀ y (Neq)eq:", { size: SZ })],
         { before: 60 }),
    formula("Nff = max( N'₅₀ , (Neq)eq )  =  max( " +
            CORR.N50_corr.toFixed(2) + " , " +
            PROM.fuente.Neq.toFixed(2) + " )  =  " +
            RES.Nff.toFixed(2) + " dB"),
    emptyPara(60),

    ...cfAplicacion(),

    emptyPara(80),
    resultadoFinalTable(),
  ];
}

function promediosTable() {
  const W  = CONTENT_W;
  const c0 = 2600, c1 = Math.floor((W-c0)/2), c2 = W-c0-c1;
  const rows = [];
  rows.push(new TableRow({ children: [
    new TableCell({ columnSpan: 3, shading: shade(COLOR_GREEN), borders: borders_thick,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [para([run("PROMEDIOS GLOBALES", { bold: true, size: SZ })],
                 { align: AlignmentType.CENTER })]}),
  ]}));
  rows.push(new TableRow({ children: [
    new TableCell({ width: {size:c0,type:WidthType.DXA}, shading: shade(COLOR_HEADER),
      borders: borders_all, margins: {top:60,bottom:60,left:100,right:100},
      children: [para([run("Indicador",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
    new TableCell({ width: {size:c1,type:WidthType.DXA}, shading: shade(COLOR_HEADER),
      borders: borders_all, margins: {top:60,bottom:60,left:100,right:100},
      children: [para([run("FUENTE (A–E)",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
    new TableCell({ width: {size:c2,type:WidthType.DXA}, shading: shade(COLOR_BLUE_H),
      borders: borders_all, margins: {top:60,bottom:60,left:100,right:100},
      children: [para([run("FONDO (I–V)",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
  ]}));

  const indicadores = [
    ["N₅₀ (dB)", "N50"], ["σ (dB)","sigma"], ["N₁₀ (dB)","N10"], ["(Neq)eq (dB)","Neq"]
  ];
  indicadores.forEach(([label, key], i) => {
    const fill = i%2===0?"FFFFFF":COLOR_GREY;
    rows.push(new TableRow({ children: [
      new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
        borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
        children:[para([run(label,{bold:true,size:SZ})],{align:AlignmentType.LEFT})]}),
      new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade(fill),
        borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
        children:[para([run(PROM.fuente[key].toFixed(2),{size:SZ})],{align:AlignmentType.CENTER})]}),
      new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade(fill),
        borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
        children:[para([run(PROM.fondo[key].toFixed(2),{size:SZ})],{align:AlignmentType.CENTER})]}),
    ]}));
  });
  return new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[c0,c1,c2], rows });
}

function corrTabla() {
  const W = CONTENT_W;
  const c0=3800, c1=2000, c2=W-c0-c1;
  const cfVal = CORR.Cf_aplica ? CORR.Cf.toFixed(2) : "No Aplica";

  const tbl = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[c0,c1,c2],
    rows: [
      new TableRow({ children: [
        new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Parámetro",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Valor",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Unidad",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
      ]}),
      ...[ ["Ce =", CORR.Ce.toFixed(2), "dB"],
           ["Δ₅₀ =", Math.abs(CORR.delta50).toFixed(2), "dB"],
           ["Cf =", cfVal, "dB"],
           ["N'₅₀ =", CORR.Cf_aplica ? CORR.N50_corr.toFixed(2) : "No Aplica", "dB"],
      ].map(([p,v,u], i) => new TableRow({ children: [
        new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade(i%2===0?"FFFFFF":COLOR_GREY),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run(p,{size:SZ})],{align:AlignmentType.LEFT})]}),
        new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade(i%2===0?"FFFFFF":COLOR_GREY),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run(v,{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade(i%2===0?"FFFFFF":COLOR_GREY),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run(u,{size:SZ})],{align:AlignmentType.CENTER})]}),
      ]})),
    ],
  });

  return [emptyPara(60), tbl, emptyPara(60)];
}

function cfAplicacion() {
  if (!CORR.Cf_aplica) {
    return [
      new Paragraph({
        spacing: { before: 80, after: 80 },
        indent:  { left: 360 },
        children: [
          run("Como Δ₅₀ = " + CORR.delta50.toFixed(2) +
              " dB < 0.75 dB, la corrección por ruido de fondo ",
              { size: SZ }),
          run("no aplica", { size: SZ, bold: true }),
          run(". Por lo tanto:", { size: SZ }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 80 },
        children: [run("(N')ff = Nff = " + RES.Nff_corr.toFixed(2) + " dB",
                   { size: SZ, italic: true })],
      }),
    ];
  }
  return [
    new Paragraph({
      spacing: { before: 80, after: 40 },
      indent: { left: 360 },
      children: [run("Como Δ₅₀ = " + CORR.delta50.toFixed(2) +
                     " dB ≥ 0.75 dB, aplica la corrección por ruido de fondo:", { size: SZ })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
      children: [run("Cf = −(Δ₅₀+9) + 3√(4·Δ₅₀−3) = " + CORR.Cf.toFixed(2) + " dB",
                 { size: SZ, italic: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
      children: [run("(N')ff = Nff + Cf = " + RES.Nff.toFixed(2) +
                     " + " + CORR.Cf.toFixed(2) + " = " + RES.Nff_corr.toFixed(2) + " dB",
                 { size: SZ, italic: true })],
    }),
  ];
}

function resultadoFinalTable() {
  const W = CONTENT_W;
  const c0=3800, c1=2000, c2=W-c0-c1;
  const excede = RES.Nff_corr > M.limite;
  const verdict = excede
    ? `EXCEDE el límite (${M.limite} dB)`
    : `CUMPLE el límite (${M.limite} dB)`;
  const fillResult = excede ? "FFCCCC" : "C6EFCE";

  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[c0,c1,c2],
    rows: [
      new TableRow({ children: [
        new TableCell({ columnSpan:3, shading:shade(COLOR_GREEN), borders:borders_thick,
          margins:{top:80,bottom:80,left:100,right:100},
          children:[para([run("NIVEL DE FUENTE FIJA — RESULTADO FINAL",{bold:true,size:SZ})],
                   {align:AlignmentType.CENTER})]}),
      ]}),
      new TableRow({ children: [
        new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Parámetro",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Valor",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade(COLOR_HEADER),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("Unidad",{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
      ]}),
      new TableRow({ children: [
        new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade("FFFFFF"),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("N_ff =",{size:SZ})],{align:AlignmentType.LEFT})]}),
        new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade("FFFFFF"),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run(RES.Nff.toFixed(2),{bold:true,size:SZ})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade("FFFFFF"),
          borders:borders_all, margins:{top:60,bottom:60,left:100,right:100},
          children:[para([run("dB",{size:SZ})],{align:AlignmentType.CENTER})]}),
      ]}),
      new TableRow({ children: [
        new TableCell({ width:{size:c0,type:WidthType.DXA}, shading:shade(fillResult),
          borders:borders_thick, margins:{top:80,bottom:80,left:100,right:100},
          children:[para([run("(N')ff =",{bold:true,size:SZ_SECTION})],{align:AlignmentType.LEFT})]}),
        new TableCell({ width:{size:c1,type:WidthType.DXA}, shading:shade(fillResult),
          borders:borders_thick, margins:{top:80,bottom:80,left:100,right:100},
          children:[para([run(RES.Nff_corr.toFixed(2),{bold:true,size:SZ_SECTION})],{align:AlignmentType.CENTER})]}),
        new TableCell({ width:{size:c2,type:WidthType.DXA}, shading:shade(fillResult),
          borders:borders_thick, margins:{top:80,bottom:80,left:100,right:100},
          children:[para([run("dB",{bold:true,size:SZ_SECTION})],{align:AlignmentType.CENTER})]}),
      ]}),
      new TableRow({ children: [
        new TableCell({ columnSpan:3, shading:shade(fillResult), borders:borders_thick,
          margins:{top:80,bottom:80,left:100,right:100},
          children:[para([run(verdict, {bold:true, size:SZ_SECTION,
                         color: excede ? "C00000" : "375623"})],
                   {align:AlignmentType.CENTER})]}),
      ]}),
    ],
  });
}

// ── FOOTER ────────────────────────────────────────────────────────────────────
function footerPara() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "375623", space: 4 } },
        spacing: { before: 60, after: 0 },
        children: [
          run("NOM-081-SEMARNAT-1994  |  " + M.compania + "  |  " + M.fecha,
              { size: 16, color: "666666" }),
        ],
      }),
    ],
  });
}

// ── ARMAR DOCUMENTO ────────────────────────────────────────────────────────────
const children = [
  titlePara(),
  emptyPara(80),
  headerTable(),
  emptyPara(160),

  // Datos de campo fuente
  datosTable("FUENTE", ["A","B","C","D","E"], FD, PF),
  emptyPara(160),

  // Datos de campo fondo
  datosTable("FONDO", ["I","II","III","IV","V"], BD, PB),
  emptyPara(160),

  // Secciones A–D con explicaciones y resultados
  ...seccionFormulas(),
];

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size:   { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    footers: { default: footerPara() },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[3], buf);
  console.log("✅ Word generado:", process.argv[3]);
}).catch(err => {
  console.error("❌ Error:", err.message);
  process.exit(1);
});
