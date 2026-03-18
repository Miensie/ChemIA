/**
 * ================================================================
 * chemCharts.js — Moteur de graphiques SVG pour la chimiométrie
 *
 * Graphiques implémentés :
 *  - Score Plot (PCA/PLS)
 *  - Loading Plot
 *  - Scree Plot
 *  - Biplot
 *  - VIP Scores (PLS)
 *  - Prédit vs Réel (PLS)
 *  - Dendrogramme
 *  - Heatmap
 *  - Graphique de contrôle Mahalanobis / T²
 *  - Méthode Elbow
 * ================================================================
 */
"use strict";

// ─── Palette et utilitaires ───────────────────────────────────────────────────

const CLUSTER_COLORS = [
  "#00E5FF","#00E676","#FF9800","#CE93D8","#FF5252",
  "#FFEB3B","#40C4FF","#FF80AB","#69F0AE","#EA80FC",
];

const BG    = "#080F1A";
const BG2   = "#0D1825";
const BG3   = "#122035";
const GRID  = "#1A2E47";
const TEXT  = "#7B9DB8";
const CYAN  = "#00E5FF";
const GREEN = "#00E676";
const ORANGE= "#FF9800";
const RED   = "#FF5252";

function svgOpen(w, h, extraStyle = "") {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}"
    style="width:100%;height:auto;display:block;background:${BG};border-radius:6px;${extraStyle}">`;
}
function svgClose() { return '</svg>'; }

function fmt(v, dec = 2) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  if (Math.abs(v) >= 1000) return (v / 1000).toFixed(1) + 'k';
  return Number(v).toFixed(dec);
}

function linearScale(domain, range) {
  const [d0, d1] = domain, [r0, r1] = range;
  return v => r0 + (v - d0) / (d1 - d0 + 1e-12) * (r1 - r0);
}

function gridLines(xScale, yScale, ML, MT, CW, CH, nX = 5, nY = 5) {
  let g = '';
  for (let i = 0; i <= nX; i++) {
    const x = ML + i * CW / nX;
    g += `<line x1="${x}" y1="${MT}" x2="${x}" y2="${MT+CH}" stroke="${GRID}" stroke-width="0.6"/>`;
  }
  for (let i = 0; i <= nY; i++) {
    const y = MT + i * CH / nY;
    g += `<line x1="${ML}" y1="${y}" x2="${ML+CW}" y2="${y}" stroke="${GRID}" stroke-width="0.6"/>`;
  }
  return g;
}

// ─── Score Plot ──────────────────────────────────────────────────────────────

/**
 * Score Plot PCA ou PLS
 * @param {number[][]} scores  — m×nComp
 * @param {string[]}   labels  — noms des échantillons
 * @param {number}     pcX     — index composante X (0-based)
 * @param {number}     pcY     — index composante Y
 * @param {number[]}   explVar — variance expliquée en %
 * @param {number[]}   clusterLabels — facultatif, colore par cluster
 */
function buildScorePlot(scores, sampleNames, pcX, pcY, explVar, clusterLabels) {
  const W = 560, H = 360;
  const ML = 52, MR = 24, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const xs = scores.map(s => s[pcX]);
  const ys = scores.map(s => s[pcY]);
  const xExt = [Math.min(...xs), Math.max(...xs)];
  const yExt = [Math.min(...ys), Math.max(...ys)];
  const pad = v => [v[0] - (v[1]-v[0])*.12, v[1] + (v[1]-v[0])*.12];
  const [xLo, xHi] = pad(xExt), [yLo, yHi] = pad(yExt);
  const sx = linearScale([xLo, xHi], [ML, ML+CW]);
  const sy = linearScale([yLo, yHi], [MT+CH, MT]);

  let svg = svgOpen(W, H);
  svg += `<rect width="${W}" height="${H}" fill="${BG}"/>`;
  svg += gridLines(sx, sy, ML, MT, CW, CH);

  // Axes zéro
  if (xLo < 0 && xHi > 0) {
    const x0 = sx(0).toFixed(1);
    svg += `<line x1="${x0}" y1="${MT}" x2="${x0}" y2="${MT+CH}" stroke="${GRID}" stroke-width="1.2" stroke-dasharray="4,3"/>`;
  }
  if (yLo < 0 && yHi > 0) {
    const y0 = sy(0).toFixed(1);
    svg += `<line x1="${ML}" y1="${y0}" x2="${ML+CW}" y2="${y0}" stroke="${GRID}" stroke-width="1.2" stroke-dasharray="4,3"/>`;
  }

  // Ellipse de confiance 95% (approximation)
  const mux = xs.reduce((s,v)=>s+v,0)/xs.length;
  const muy = ys.reduce((s,v)=>s+v,0)/ys.length;
  const sdx = Math.sqrt(xs.reduce((s,v)=>s+(v-mux)**2,0)/(xs.length-1));
  const sdy = Math.sqrt(ys.reduce((s,v)=>s+(v-muy)**2,0)/(ys.length-1));
  const rx = sdx * 2.45 * CW / (xHi - xLo);
  const ry = sdy * 2.45 * CH / (yHi - yLo);
  svg += `<ellipse cx="${sx(mux).toFixed(1)}" cy="${sy(muy).toFixed(1)}"
    rx="${rx.toFixed(1)}" ry="${ry.toFixed(1)}"
    fill="none" stroke="${CYAN}" stroke-width="1" stroke-dasharray="5,3" opacity="0.4"/>`;

  // Points
  scores.forEach((s, i) => {
    const cx = sx(xs[i]).toFixed(1), cy = sy(ys[i]).toFixed(1);
    const col = clusterLabels ? CLUSTER_COLORS[clusterLabels[i] % CLUSTER_COLORS.length] : CYAN;
    svg += `<circle cx="${cx}" cy="${cy}" r="5" fill="${col}" stroke="${BG}" stroke-width="1.2" opacity="0.85"/>`;
    // Label
    const name = (sampleNames && sampleNames[i]) ? sampleNames[i].substring(0,8) : `${i+1}`;
    svg += `<text x="${+cx+7}" y="${+cy+3}" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${name}</text>`;
  });

  // Axes labels
  const xLbl = `PC${pcX+1}${explVar ? ` (${explVar[pcX]?.toFixed(1)}%)` : ''}`;
  const yLbl = `PC${pcY+1}${explVar ? ` (${explVar[pcY]?.toFixed(1)}%)` : ''}`;
  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">${xLbl}</text>`;
  svg += `<text transform="rotate(-90,14,${MT+CH/2})" x="14" y="${MT+CH/2+4}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">${yLbl}</text>`;
  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Score Plot</text>`;

  // Ticks X
  for (let i = 0; i <= 4; i++) {
    const v = xLo + i*(xHi-xLo)/4;
    const xp = sx(v).toFixed(1);
    svg += `<text x="${xp}" y="${MT+CH+14}" text-anchor="middle" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,2)}</text>`;
  }
  // Ticks Y
  for (let i = 0; i <= 4; i++) {
    const v = yLo + i*(yHi-yLo)/4;
    const yp = sy(v).toFixed(1);
    svg += `<text x="${ML-5}" y="${+yp+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,2)}</text>`;
  }

  svg += svgClose();
  return svg;
}

// ─── Loading Plot ─────────────────────────────────────────────────────────────

function buildLoadingPlot(loadings, varNames, pcX, pcY, explVar) {
  const W = 560, H = 360;
  const ML = 52, MR = 24, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const xs = loadings[pcX];
  const ys = loadings[pcY];
  const absMax = Math.max(Math.max(...xs.map(Math.abs)), Math.max(...ys.map(Math.abs))) * 1.2 || 1;
  const sx = linearScale([-absMax, absMax], [ML, ML+CW]);
  const sy = linearScale([-absMax, absMax], [MT+CH, MT]);

  let svg = svgOpen(W, H);
  svg += gridLines(sx, sy, ML, MT, CW, CH);

  // Axes zéro
  const x0 = sx(0).toFixed(1), y0 = sy(0).toFixed(1);
  svg += `<line x1="${x0}" y1="${MT}" x2="${x0}" y2="${MT+CH}" stroke="${GRID}" stroke-width="1.2" stroke-dasharray="3,3"/>`;
  svg += `<line x1="${ML}" y1="${y0}" x2="${ML+CW}" y2="${y0}" stroke="${GRID}" stroke-width="1.2" stroke-dasharray="3,3"/>`;

  // Cercle unitaire
  const cx0 = sx(0), cy0 = sy(0);
  const rUnit = Math.min(CW, CH) / 2 * 0.85;
  svg += `<circle cx="${cx0.toFixed(1)}" cy="${cy0.toFixed(1)}" r="${rUnit.toFixed(1)}" fill="none" stroke="${GRID}" stroke-width="0.8" opacity="0.5"/>`;

  // Vecteurs (flèches)
  xs.forEach((x, i) => {
    const tx = sx(x).toFixed(1), ty = sy(ys[i]).toFixed(1);
    const len = Math.sqrt(x*x + ys[i]*ys[i]) / absMax;
    const col = len > 0.6 ? ORANGE : len > 0.3 ? CYAN : TEXT;
    svg += `<line x1="${x0}" y1="${y0}" x2="${tx}" y2="${ty}" stroke="${col}" stroke-width="1.5" opacity="0.8"/>`;
    svg += `<circle cx="${tx}" cy="${ty}" r="3" fill="${col}" opacity="0.9"/>`;
    const name = varNames[i]?.substring(0, 10) || `V${i+1}`;
    const offset = +tx > +x0 ? 5 : -5;
    const anchor = +tx > +x0 ? 'start' : 'end';
    svg += `<text x="${+tx+offset}" y="${+ty+3}" text-anchor="${anchor}" font-size="8" fill="${col}" font-family="Space Mono,monospace">${name}</text>`;
  });

  const xLbl = `PC${pcX+1}${explVar ? ` (${explVar[pcX]?.toFixed(1)}%)` : ''}`;
  const yLbl = `PC${pcY+1}${explVar ? ` (${explVar[pcY]?.toFixed(1)}%)` : ''}`;
  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">${xLbl}</text>`;
  svg += `<text transform="rotate(-90,14,${MT+CH/2})" x="14" y="${MT+CH/2+4}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">${yLbl}</text>`;
  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Loading Plot</text>`;
  svg += svgClose();
  return svg;
}

// ─── Scree Plot ───────────────────────────────────────────────────────────────

function buildScreePlot(explainedVar, cumulativeVar) {
  const W = 480, H = 300;
  const ML = 52, MR = 60, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const n = explainedVar.length;
  const barW = CW / n * 0.6;
  const sx = i => ML + (i + 0.5) * CW / n;
  const sy = v => MT + CH - v / 100 * CH;

  let svg = svgOpen(W, H);
  svg += gridLines(null, null, ML, MT, CW, CH, 0, 5);

  // Barres de variance individuelle
  explainedVar.forEach((v, i) => {
    const x = sx(i) - barW/2, y = sy(v);
    svg += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${(MT+CH-y).toFixed(1)}"
      fill="${CYAN}" opacity="0.7" rx="2"/>`;
    svg += `<text x="${sx(i).toFixed(1)}" y="${MT+CH+14}" text-anchor="middle" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">PC${i+1}</text>`;
    svg += `<text x="${sx(i).toFixed(1)}" y="${(y-3).toFixed(1)}" text-anchor="middle" font-size="8" fill="${CYAN}" font-family="Space Mono,monospace">${v.toFixed(1)}%</text>`;
  });

  // Courbe de variance cumulée
  const curvePts = cumulativeVar.map((v, i) => `${sx(i).toFixed(1)},${sy(v).toFixed(1)}`).join(' ');
  svg += `<polyline points="${curvePts}" fill="none" stroke="${GREEN}" stroke-width="2"/>`;
  cumulativeVar.forEach((v, i) => {
    svg += `<circle cx="${sx(i).toFixed(1)}" cy="${sy(v).toFixed(1)}" r="4" fill="${GREEN}" stroke="${BG}" stroke-width="1.5"/>`;
  });

  // Ligne 95%
  const y95 = sy(95).toFixed(1);
  svg += `<line x1="${ML}" y1="${y95}" x2="${ML+CW}" y2="${y95}" stroke="${ORANGE}" stroke-width="1" stroke-dasharray="4,3"/>`;
  svg += `<text x="${ML+CW+4}" y="${+y95+3}" font-size="8" fill="${ORANGE}" font-family="Space Mono,monospace">95%</text>`;

  // Ticks Y
  for (let i = 0; i <= 5; i++) {
    const v = i * 20;
    svg += `<text x="${ML-5}" y="${sy(v)+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${v}%</text>`;
  }

  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Composante principale</text>`;
  svg += `<text transform="rotate(-90,14,${MT+CH/2})" x="14" y="${MT+CH/2+4}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Variance (%)</text>`;
  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Scree Plot</text>`;

  // Légende
  svg += `<line x1="${ML}" y1="${H-18}" x2="${ML+16}" y2="${H-18}" stroke="${CYAN}" stroke-width="4"/>`;
  svg += `<text x="${ML+20}" y="${H-15}" font-size="8" fill="${CYAN}" font-family="Space Mono,monospace">Variance individuelle</text>`;
  svg += `<line x1="${ML+140}" y1="${H-18}" x2="${ML+156}" y2="${H-18}" stroke="${GREEN}" stroke-width="2"/>`;
  svg += `<text x="${ML+160}" y="${H-15}" font-size="8" fill="${GREEN}" font-family="Space Mono,monospace">Variance cumulée</text>`;

  svg += svgClose();
  return svg;
}

// ─── PLS : Prédit vs Réel ────────────────────────────────────────────────────

function buildPredVsReal(yReal, yPred, r2, rmse) {
  const W = 440, H = 360;
  const ML = 52, MR = 20, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const all = [...yReal, ...yPred];
  const lo = Math.min(...all), hi = Math.max(...all);
  const pad = (hi - lo) * 0.1;
  const sx = linearScale([lo-pad, hi+pad], [ML, ML+CW]);
  const sy = linearScale([lo-pad, hi+pad], [MT+CH, MT]);

  let svg = svgOpen(W, H);
  svg += gridLines(null, null, ML, MT, CW, CH);

  // Ligne parfaite y=x
  svg += `<line x1="${ML}" y1="${MT+CH}" x2="${ML+CW}" y2="${MT}" stroke="${GRID}" stroke-width="1.5" stroke-dasharray="5,3"/>`;

  // Points
  yReal.forEach((y, i) => {
    const cx = sx(y).toFixed(1), cy = sy(yPred[i]).toFixed(1);
    svg += `<circle cx="${cx}" cy="${cy}" r="4.5" fill="${CYAN}" stroke="${BG}" stroke-width="1" opacity="0.8"/>`;
  });

  // Stats box
  svg += `<rect x="${ML+CW-90}" y="${MT+6}" width="88" height="38" rx="4" fill="${BG3}" stroke="${GRID}"/>`;
  svg += `<text x="${ML+CW-46}" y="${MT+20}" text-anchor="middle" font-size="9" fill="${CYAN}" font-family="Space Mono,monospace">R² = ${fmt(r2,4)}</text>`;
  svg += `<text x="${ML+CW-46}" y="${MT+34}" text-anchor="middle" font-size="9" fill="${GREEN}" font-family="Space Mono,monospace">RMSE = ${fmt(rmse,4)}</text>`;

  // Ticks
  for (let i = 0; i <= 4; i++) {
    const v = lo-pad + i*(hi+pad-lo+pad)/4;
    const xp = sx(v).toFixed(1);
    svg += `<text x="${xp}" y="${MT+CH+14}" text-anchor="middle" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,2)}</text>`;
    const yp = sy(v).toFixed(1);
    svg += `<text x="${ML-5}" y="${+yp+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,2)}</text>`;
  }

  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Valeurs réelles Y</text>`;
  svg += `<text transform="rotate(-90,14,${MT+CH/2})" x="14" y="${MT+CH/2+4}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Valeurs prédites Ŷ</text>`;
  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Prédit vs Réel</text>`;
  svg += svgClose();
  return svg;
}

// ─── VIP Scores ───────────────────────────────────────────────────────────────

function buildVIPChart(vip, varNames) {
  const W = 560, H = Math.max(260, varNames.length * 18 + 80);
  const ML = 130, MR = 60, MT = 28, MB = 30;
  const CW = W - ML - MR, CH = H - MT - MB;
  const barH = Math.min(14, CH / varNames.length - 2);

  const maxVIP = Math.max(...vip, 1.2) * 1.1;
  const sx = v => ML + v / maxVIP * CW;

  // Trier par VIP décroissant
  const order = vip.map((v, i) => ({ v, i })).sort((a, b) => b.v - a.v);

  let svg = svgOpen(W, H);

  // Ligne VIP = 1
  const x1 = sx(1).toFixed(1);
  svg += `<line x1="${x1}" y1="${MT}" x2="${x1}" y2="${MT+CH}" stroke="${ORANGE}" stroke-width="1.2" stroke-dasharray="4,3"/>`;
  svg += `<text x="${x1}" y="${MT-5}" text-anchor="middle" font-size="8" fill="${ORANGE}" font-family="Space Mono,monospace">VIP=1</text>`;

  order.forEach(({ v, i }, rank) => {
    const y = MT + rank * (barH + 2);
    const w = sx(v) - ML;
    const col = v >= 1.0 ? ORANGE : v >= 0.8 ? CYAN : TEXT;
    svg += `<rect x="${ML}" y="${y}" width="${w.toFixed(1)}" height="${barH}" fill="${col}" opacity="0.8" rx="2"/>`;
    const name = (varNames[i] || `V${i+1}`).substring(0, 16);
    svg += `<text x="${ML-5}" y="${y+barH-2}" text-anchor="end" font-size="9" fill="${TEXT}" font-family="Space Mono,monospace">${name}</text>`;
    svg += `<text x="${(ML+w+4).toFixed(1)}" y="${y+barH-2}" font-size="9" fill="${col}" font-family="Space Mono,monospace">${v.toFixed(2)}</text>`;
  });

  svg += `<text x="${ML+CW/2}" y="${MT-14}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">VIP Scores</text>`;
  svg += svgClose();
  return svg;
}

// ─── Heatmap ──────────────────────────────────────────────────────────────────

function buildHeatmap(X, rowNames, colNames) {
  const nRows = Math.min(X.length, 50);
  const nCols = X[0].length;
  const cellW = Math.min(40, 480 / nCols);
  const cellH = Math.min(20, 400 / nRows);
  const ML = 100, MT = 80, MR = 60, MB = 20;
  const W = ML + nCols * cellW + MR;
  const H = MT + nRows * cellH + MB;

  const flat = X.slice(0, nRows).flat().filter(v => isFinite(v));
  const lo = Math.min(...flat), hi = Math.max(...flat);

  function heatColor(v) {
    const t = (v - lo) / (hi - lo + 1e-12);
    const stops = [
      [0,   [8,  15, 26]],   // dark bg
      [0.25,[0,  60, 150]],
      [0.5, [0, 200, 220]],
      [0.75,[0, 200, 100]],
      [1.0, [255,140, 0]],
    ];
    let s = stops[0], e = stops[stops.length-1];
    for (let i = 0; i < stops.length-1; i++) {
      if (t >= stops[i][0] && t <= stops[i+1][0]) { s = stops[i]; e = stops[i+1]; break; }
    }
    const u = (t - s[0]) / (e[0] - s[0] + 1e-12);
    const r = Math.round(s[1][0] + u*(e[1][0]-s[1][0]));
    const g = Math.round(s[1][1] + u*(e[1][1]-s[1][1]));
    const b = Math.round(s[1][2] + u*(e[1][2]-s[1][2]));
    return `rgb(${r},${g},${b})`;
  }

  let svg = svgOpen(W, H);

  // En-têtes colonnes
  colNames.forEach((name, j) => {
    svg += `<text x="${ML + j*cellW + cellW/2}" y="${MT-4}"
      text-anchor="end" transform="rotate(-45,${ML+j*cellW+cellW/2},${MT-4})"
      font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${name.substring(0,10)}</text>`;
  });

  // Cellules
  X.slice(0, nRows).forEach((row, i) => {
    const y = MT + i * cellH;
    const rName = (rowNames && rowNames[i]) ? rowNames[i].substring(0,14) : `${i+1}`;
    svg += `<text x="${ML-4}" y="${y+cellH-2}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${rName}</text>`;
    row.forEach((v, j) => {
      const x = ML + j * cellW;
      svg += `<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${heatColor(v)}" stroke="${BG}" stroke-width="0.3"/>`;
    });
  });

  // Colorbar
  const cbX = ML + nCols * cellW + 8, cbH = Math.min(120, H - MT - MB);
  for (let i = 0; i <= 30; i++) {
    const t = i / 30;
    const v = lo + t * (hi - lo);
    const yc = MT + (1 - t) * cbH;
    svg += `<rect x="${cbX}" y="${yc.toFixed(1)}" width="10" height="${(cbH/30+0.5).toFixed(1)}" fill="${heatColor(v)}"/>`;
  }
  svg += `<text x="${cbX+5}" y="${MT-4}" text-anchor="middle" font-size="7" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(hi,2)}</text>`;
  svg += `<text x="${cbX+5}" y="${MT+cbH+12}" text-anchor="middle" font-size="7" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(lo,2)}</text>`;

  svg += `<text x="${ML+nCols*cellW/2}" y="16" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Heatmap</text>`;
  svg += svgClose();
  return svg;
}

// ─── Dendrogramme (simplifié) ────────────────────────────────────────────────

function buildDendrogram(linkageMatrix, sampleNames, k) {
  const n = sampleNames ? sampleNames.length : (linkageMatrix.length + 1);
  const W = Math.max(500, n * 16 + 100);
  const H = 360;
  const ML = 60, MR = 20, MT = 20, MB = 60;
  const CW = W - ML - MR, CH = H - MT - MB;

  // Positions X des feuilles
  const leafX = Array.from({length: n}, (_, i) => ML + (i + 0.5) * CW / n);
  const maxDist = Math.max(...linkageMatrix.map(l => l[2])) || 1;
  const sy = d => MT + CH - (d / maxDist) * CH;

  // Position X des nœuds internes
  const nodeX = [...leafX];
  const nodeY = new Array(n).fill(MT + CH);

  let svg = svgOpen(W, H);
  svg += gridLines(null, null, ML, MT, CW, CH, 0, 4);

  // Couleurs de coupe pour k clusters
  const cutDist = k > 1 && linkageMatrix.length >= k - 1
    ? linkageMatrix[linkageMatrix.length - k + 1]?.[2] ?? 0 : 0;

  if (cutDist > 0) {
    const yc = sy(cutDist).toFixed(1);
    svg += `<line x1="${ML}" y1="${yc}" x2="${ML+CW}" y2="${yc}" stroke="${RED}" stroke-width="1.2" stroke-dasharray="5,3"/>`;
    svg += `<text x="${ML+4}" y="${+yc-3}" font-size="8" fill="${RED}" font-family="Space Mono,monospace">k=${k}</text>`;
  }

  // Lignes du dendrogramme
  linkageMatrix.forEach(([a, b, dist, size], m) => {
    const newId = n + m;
    const xa = nodeX[a], xb = nodeX[b];
    const ya = nodeY[a], yb = nodeY[b];
    const yNew = sy(dist);
    const xNew = (xa + xb) / 2;

    const col = dist > cutDist ? TEXT : ORANGE;
    svg += `<line x1="${xa.toFixed(1)}" y1="${ya.toFixed(1)}" x2="${xa.toFixed(1)}" y2="${yNew.toFixed(1)}" stroke="${col}" stroke-width="1.4"/>`;
    svg += `<line x1="${xb.toFixed(1)}" y1="${yb.toFixed(1)}" x2="${xb.toFixed(1)}" y2="${yNew.toFixed(1)}" stroke="${col}" stroke-width="1.4"/>`;
    svg += `<line x1="${xa.toFixed(1)}" y1="${yNew.toFixed(1)}" x2="${xb.toFixed(1)}" y2="${yNew.toFixed(1)}" stroke="${col}" stroke-width="1.4"/>`;

    nodeX[newId] = xNew;
    nodeY[newId] = yNew;
  });

  // Labels feuilles
  leafX.forEach((x, i) => {
    const name = (sampleNames && sampleNames[i]) ? sampleNames[i].substring(0,8) : `${i+1}`;
    svg += `<text x="${x.toFixed(1)}" y="${MT+CH+14}" text-anchor="end"
      transform="rotate(-45,${x.toFixed(1)},${MT+CH+14})"
      font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${name}</text>`;
  });

  // Ticks Y (distances)
  for (let i = 0; i <= 4; i++) {
    const d = i * maxDist / 4;
    const yp = sy(d).toFixed(1);
    svg += `<text x="${ML-5}" y="${+yp+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(d,2)}</text>`;
  }

  svg += `<text transform="rotate(-90,14,${MT+CH/2})" x="14" y="${MT+CH/2+4}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Distance</text>`;
  svg += `<text x="${ML+CW/2}" y="16" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Dendrogramme</text>`;
  svg += svgClose();
  return svg;
}

// ─── Graphique de contrôle Mahalanobis ───────────────────────────────────────

function buildControlChart(scores, threshold, sampleNames, title) {
  const W = 560, H = 280;
  const ML = 52, MR = 20, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const m = scores.length;
  const maxY = Math.max(Math.max(...scores), threshold) * 1.15;
  const sx = i => ML + (i + 0.5) * CW / m;
  const sy = v => MT + CH - v / maxY * CH;

  let svg = svgOpen(W, H);
  svg += gridLines(null, null, ML, MT, CW, CH, 0, 4);

  // Ligne seuil
  const yThresh = sy(threshold).toFixed(1);
  svg += `<line x1="${ML}" y1="${yThresh}" x2="${ML+CW}" y2="${yThresh}" stroke="${RED}" stroke-width="1.2" stroke-dasharray="5,3"/>`;
  svg += `<text x="${ML+CW+4}" y="${+yThresh+3}" font-size="8" fill="${RED}" font-family="Space Mono,monospace">Seuil</text>`;

  // Barres
  scores.forEach((v, i) => {
    const x = sx(i), y = sy(v);
    const h = MT + CH - y;
    const col = v > threshold ? RED : CYAN;
    svg += `<rect x="${(x-3).toFixed(1)}" y="${y.toFixed(1)}" width="6" height="${h.toFixed(1)}" fill="${col}" opacity="0.8"/>`;
    svg += `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="${col}" stroke="${BG}" stroke-width="1"/>`;
  });

  // Labels X (tous les n/10)
  const step = Math.max(1, Math.floor(m / 10));
  for (let i = 0; i < m; i += step) {
    const name = (sampleNames && sampleNames[i]) ? sampleNames[i].substring(0,6) : `${i+1}`;
    svg += `<text x="${sx(i).toFixed(1)}" y="${MT+CH+14}" text-anchor="middle" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${name}</text>`;
  }

  // Ticks Y
  for (let i = 0; i <= 4; i++) {
    const v = i * maxY / 4;
    svg += `<text x="${ML-5}" y="${sy(v)+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,2)}</text>`;
  }

  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">${title || 'Graphique de contrôle'}</text>`;
  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Échantillon</text>`;
  svg += svgClose();
  return svg;
}

// ─── Méthode Elbow ────────────────────────────────────────────────────────────

function buildElbowChart(elbowData) {
  const W = 400, H = 280;
  const ML = 52, MR = 20, MT = 28, MB = 44;
  const CW = W - ML - MR, CH = H - MT - MB;

  const ks = elbowData.map(d => d.k);
  const inertias = elbowData.map(d => d.inertia);
  const kMin = Math.min(...ks), kMax = Math.max(...ks);
  const iMax = Math.max(...inertias);
  const sx = k => ML + (k - kMin) / (kMax - kMin + 1e-12) * CW;
  const sy = v => MT + CH - v / iMax * CH;

  let svg = svgOpen(W, H);
  svg += gridLines(null, null, ML, MT, CW, CH);

  let path = '';
  elbowData.forEach((d, i) => {
    const x = sx(d.k).toFixed(1), y = sy(d.inertia).toFixed(1);
    path += (i === 0 ? `M${x},${y}` : `L${x},${y}`);
    svg += `<circle cx="${x}" cy="${y}" r="5" fill="${CYAN}" stroke="${BG}" stroke-width="1.5"/>`;
    svg += `<text x="${x}" y="${+y-8}" text-anchor="middle" font-size="8" fill="${CYAN}" font-family="Space Mono,monospace">k=${d.k}</text>`;
  });
  svg += `<path d="${path}" fill="none" stroke="${CYAN}" stroke-width="2"/>`;

  // Ticks
  ks.forEach(k => {
    svg += `<text x="${sx(k).toFixed(1)}" y="${MT+CH+14}" text-anchor="middle" font-size="9" fill="${TEXT}" font-family="Space Mono,monospace">${k}</text>`;
  });
  for (let i = 0; i <= 4; i++) {
    const v = i * iMax / 4;
    svg += `<text x="${ML-5}" y="${sy(v)+3}" text-anchor="end" font-size="8" fill="${TEXT}" font-family="Space Mono,monospace">${fmt(v,0)}</text>`;
  }

  svg += `<text x="${ML+CW/2}" y="${MT-10}" text-anchor="middle" font-size="11" fill="${CYAN}" font-family="Syne,sans-serif" font-weight="700">Méthode Elbow</text>`;
  svg += `<text x="${ML+CW/2}" y="${H-6}" text-anchor="middle" font-size="10" fill="${TEXT}" font-family="Space Mono,monospace">Nombre de clusters k</text>`;
  svg += svgClose();
  return svg;
}

// ─── Résidus PLS ─────────────────────────────────────────────────────────────

function buildResidualPlot(yPred, yReal, sampleNames) {
  const residuals = yReal.map((y, i) => yPred[i] - y);
  return buildControlChart(residuals.map(Math.abs), 0, sampleNames, 'Résidus |ŷ − y|');
}

// ─── Export ───────────────────────────────────────────────────────────────────

window.ChemCharts = {
  buildScorePlot, buildLoadingPlot, buildScreePlot,
  buildPredVsReal, buildVIPChart, buildHeatmap,
  buildDendrogram, buildControlChart, buildElbowChart, buildResidualPlot,
};
