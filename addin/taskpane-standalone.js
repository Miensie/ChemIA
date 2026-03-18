/**
 * ================================================================
 * taskpane.js — Orchestrateur principal ChemAI
 * Coordonne : données, prétraitement, PCA, PLS, clustering, anomalies, rapport
 * ================================================================
 */
"use strict";







// ─── État global ─────────────────────────────────────────────────────────────
const APP = {
  // Données brutes
  rawData:      null,    // { headers, data, sampleNames }
  // Données prétraitées
  procData:     null,    // number[][] — matrice après prétraitement
  scalerParams: null,    // paramètres du scaler (pour inverse transform)
  // Résultats d'analyses
  pcaResult:    null,
  plsModel:     null,
  plsMetrics:   null,
  plsYReal:     null,
  plsYPred:     null,
  clusterResult:null,
  anomalyResult:null,
  // Graphiques SVG en cache
  charts: {
    scorePlot: null, loadingPlot: null, screePlot: null, biplot: null,
    predVsReal: null, vipChart: null, plsCVChart: null, residualPlot: null,
    clusterScatter: null, dendro: null, heatmap: null, elbowChart: null,
    controlChart: null,
  },
  // Config courante
  scalingMethod: "none",
  missingStrategy: "mean",
};

// ─── Init Office ─────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("⚠ Excel requis");
    return;
  }

  setupNav();
  setupDataPanel();
  setupPreprocPanel();
  setupPCAPanel();
  setupPLSPanel();
  setupClusterPanel();
  setupAnomalyPanel();
  setupReportPanel();

  ChemAI.loadApiKey();
  setStatus("ChemAI v2.0 prêt ✓");
  log("Bienvenue dans ChemAI. Sélectionnez vos données pour commencer.", "info");
});

// ─── Navigation ──────────────────────────────────────────────────────────────
function setupNav() {
  document.querySelectorAll(".ntab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".ntab, .panel").forEach(el => el.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(tab.dataset.panel)?.classList.add("active");
    });
  });
}

// ─── PANEL : DONNÉES ─────────────────────────────────────────────────────────
function setupDataPanel() {
  document.getElementById("btn-read-excel").addEventListener("click", handleReadExcel);

  document.getElementById("csv-input").addEventListener("change", async e => {
    const file = e.target.files[0];
    if (!file) return;
    const text = await file.text();
    const sep  = document.getElementById("csv-sep").value;
    const hasHeader = document.getElementById("header-row").value === "1";
    try {
      const parsed = ChemExcel.parseCSV(text, sep === "\\t" ? "\t" : sep, hasHeader);
      setData(parsed);
      toast(`✅ CSV importé : ${parsed.nRows} × ${parsed.nCols}`, "ok");
      log(`CSV "${file.name}" importé — ${parsed.nRows} échantillons, ${parsed.nCols} variables`, "ok");
    } catch (e) {
      toast("Erreur CSV : " + e.message, "err");
    }
  });

  document.getElementById("btn-analyze").addEventListener("click", () => {
    switchPanel("p-preproc");
    applyPreprocessing();
  });
}

async function handleReadExcel() {
  setBtnLoading("btn-read-excel", true, "Lecture…");
  try {
    const hasHeader = document.getElementById("header-row").value === "1";
    const parsed = await ChemExcel.readSelection(hasHeader);
    setData(parsed);
    toast(`✅ ${parsed.nRows} × ${parsed.nCols} lus depuis Excel`, "ok");
    log(`Plage ${parsed.address} : ${parsed.nRows} échantillons, ${parsed.nCols} variables`, "ok");
  } catch (e) {
    toast("Erreur : " + e.message, "err");
    log("Erreur lecture Excel : " + e.message, "err");
  }
  setBtnLoading("btn-read-excel", false, "⊞ Lire la sélection Excel");
}

function setData(parsed) {
  APP.rawData = parsed;
  APP.procData = null;
  APP.pcaResult = APP.plsModel = APP.clusterResult = APP.anomalyResult = null;

  // Nettoyer les NaN pour l'aperçu
  const nanCount = parsed.data.flat().filter(v => isNaN(v)).length;
  const totalCells = parsed.nRows * parsed.nCols;

  // Afficher l'aperçu
  renderDataPreview(parsed);
  renderColRoles(parsed.headers);
  renderDescStats(parsed.data, parsed.headers);

  // Peupler les sélecteurs Y pour PLS
  populatePLSYSelector(parsed.headers);

  document.getElementById("data-preview-card").style.display = "block";
  document.getElementById("col-config-card").style.display = "block";
  document.getElementById("analyze-card").style.display = "block";

  // Stats rapides
  document.getElementById("data-shape").textContent = `${parsed.nRows} × ${parsed.nCols}`;
  document.getElementById("data-stats").innerHTML = [
    `<div class="data-stat">Échantillons : <strong>${parsed.nRows}</strong></div>`,
    `<div class="data-stat">Variables : <strong>${parsed.nCols}</strong></div>`,
    `<div class="data-stat">Valeurs manquantes : <strong>${nanCount}</strong></div>`,
    `<div class="data-stat">Complétude : <strong>${((1-nanCount/totalCells)*100).toFixed(1)}%</strong></div>`,
  ].join('');
}

function renderDataPreview(parsed) {
  const maxRows = 8, maxCols = 10;
  const thead = document.getElementById("data-preview-thead");
  const tbody = document.getElementById("data-preview-tbody");

  thead.innerHTML = `<tr><th>#</th>${parsed.headers.slice(0, maxCols).map(h => `<th>${h}</th>`).join("")}${parsed.headers.length > maxCols ? `<th>+${parsed.headers.length-maxCols}</th>` : ""}</tr>`;
  tbody.innerHTML = parsed.data.slice(0, maxRows).map((row, i) =>
    `<tr><td style="color:var(--text-faint)">${parsed.sampleNames[i] || i+1}</td>
    ${row.slice(0, maxCols).map(v => `<td>${isNaN(v) ? '<span style="color:var(--orange)">NaN</span>' : v.toFixed(4)}</td>`).join("")}
    ${row.length > maxCols ? `<td style="color:var(--text-faint)">…</td>` : ""}
    </tr>`
  ).join("");
}

function renderColRoles(headers) {
  const container = document.getElementById("col-roles");
  container.innerHTML = headers.map((h, i) => `
    <div class="col-role-row">
      <span class="col-role-name">${h}</span>
      <select class="col-role-sel" data-col="${i}">
        <option value="x">Variable X</option>
        <option value="y">Variable Y (réponse)</option>
        <option value="id">Identifiant</option>
        <option value="ignore">Ignorer</option>
      </select>
    </div>`).join("");
}

function renderDescStats(data, headers) {
  const stats = ChemMath.Preprocessing.descStats(data, headers);
  const wrap = document.getElementById("desc-stats-wrap");
  wrap.innerHTML = `
    <table class="dt">
      <thead><tr><th>Variable</th><th>n</th><th>Moyenne</th><th>Std</th><th>Min</th><th>Médiane</th><th>Max</th></tr></thead>
      <tbody>${stats.map(s => `
        <tr>
          <td style="color:var(--cyan)">${s.name}</td>
          <td>${s.n}</td>
          <td>${s.mean.toFixed(4)}</td>
          <td>${s.std.toFixed(4)}</td>
          <td>${s.min.toFixed(4)}</td>
          <td>${s.median.toFixed(4)}</td>
          <td>${s.max.toFixed(4)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`;
}

function populatePLSYSelector(headers) {
  const sel = document.getElementById("pls-y-col");
  sel.innerHTML = headers.map((h, i) => `<option value="${i}">${h}</option>`).join("");
  if (headers.length > 0) sel.value = headers.length - 1;
}

// ─── PANEL : PRÉTRAITEMENT ───────────────────────────────────────────────────
function setupPreprocPanel() {
  document.querySelectorAll("input[name='scaling']").forEach(el => {
    el.addEventListener("change", e => { APP.scalingMethod = e.target.value; });
  });
  document.getElementById("btn-preprocess").addEventListener("click", applyPreprocessing);
}

function applyPreprocessing() {
  if (!APP.rawData) { toast("Importez des données d'abord", "warn"); return; }

  const { data, headers } = APP.rawData;
  const strategy = document.getElementById("missing-strategy").value;
  const method   = APP.scalingMethod;

  // Gérer les NaN
  const cleaned = ChemMath.Preprocessing.handleMissing(data, strategy);

  // Extraire uniquement les colonnes X (ignorer les colonnes role=ignore/id)
  const xIndices = getXColumnIndices();
  const Xraw = xIndices.length > 0
    ? cleaned.map(row => xIndices.map(j => row[j]))
    : cleaned;

  // Mise à l'échelle
  APP.scalerParams = ChemMath.Preprocessing.fitScaler(Xraw, method);
  APP.procData     = ChemMath.Preprocessing.transform(Xraw, APP.scalerParams);

  // Statistiques après prétraitement
  const xHeaders = xIndices.length > 0 ? xIndices.map(j => headers[j]) : headers;
  renderDescStats(APP.procData, xHeaders);

  toast(`✅ Prétraitement appliqué : ${method || 'aucun'} — ${APP.procData.length} × ${APP.procData[0].length}`, "ok");
  log(`Prétraitement : ${method}, NaN → ${strategy}, ${APP.procData.length} × ${APP.procData[0].length}`, "ok");
}

function getXColumnIndices() {
  const roles = document.querySelectorAll(".col-role-sel");
  const xIdx  = [];
  roles.forEach((sel, i) => {
    if (sel.value === "x") xIdx.push(i);
  });
  return xIdx;
}

function getXData() {
  if (APP.procData) return APP.procData;
  if (APP.rawData) {
    applyPreprocessing();
    return APP.procData;
  }
  throw new Error("Aucune donnée disponible. Importez et prétraitez vos données.");
}

// ─── PANEL : PCA ─────────────────────────────────────────────────────────────
function setupPCAPanel() {
  document.getElementById("btn-run-pca").addEventListener("click", handleRunPCA);
  document.getElementById("btn-pca-to-excel").addEventListener("click", handlePCAToExcel);
  document.getElementById("btn-ai-pca").addEventListener("click", handleAIPCA);

  document.querySelectorAll('[data-chart]').forEach(btn => {
    if (btn.closest("#p-pca")) {
      btn.addEventListener("click", e => {
        btn.closest(".chart-tabs").querySelectorAll(".ctab").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        renderPCAChart(btn.dataset.chart);
      });
    }
  });
}

async function handleRunPCA() {
  setBtnLoading("btn-run-pca", true, "Calcul PCA…");
  try {
    const X = getXData();
    const nComp = document.getElementById("pca-ncomp").value;
    const n = nComp === "auto" ? "auto" : parseInt(nComp);

    APP.pcaResult = ChemMath.PCA.fit(X, n);

    // Remplir les sélecteurs de composantes
    const selX = document.getElementById("pca-pc-x");
    const selY = document.getElementById("pca-pc-y");
    selX.innerHTML = Array.from({length: APP.pcaResult.nComp}, (_, i) => `<option value="${i}">PC${i+1}</option>`).join("");
    selY.innerHTML = Array.from({length: APP.pcaResult.nComp}, (_, i) => `<option value="${i}">PC${i+1}</option>`).join("");
    selY.value = "1";

    // Variance table
    renderPCAVarianceTable();
    renderPCAVarianceBars();
    renderPCAChart("score");

    // Loadings table
    renderPCALoadingsTable();

    document.getElementById("pca-results").style.display = "block";
    toast(`✅ PCA terminée : ${APP.pcaResult.nComp} composantes, ${APP.pcaResult.cumulativeVar[APP.pcaResult.nComp-1]?.toFixed(1)}% variance`, "ok");
    log(`PCA : ${APP.pcaResult.nComp} CP, variance totale = ${APP.pcaResult.cumulativeVar[APP.pcaResult.nComp-1]?.toFixed(1)}%`, "ok");
  } catch (e) {
    toast("Erreur PCA : " + e.message, "err");
    log("Erreur PCA : " + e.message, "err");
    console.error(e);
  }
  setBtnLoading("btn-run-pca", false, "◈ Lancer la PCA");
}

function renderPCAVarianceTable() {
  const tbody = document.getElementById("pca-variance-tbody");
  tbody.innerHTML = APP.pcaResult.explainedVar.map((v, i) => `
    <tr>
      <td style="color:var(--cyan)">PC${i+1}</td>
      <td>${APP.pcaResult.eigenvalues[i].toFixed(4)}</td>
      <td style="color:var(--green)">${v.toFixed(2)}%</td>
      <td>${APP.pcaResult.cumulativeVar[i].toFixed(2)}%</td>
    </tr>`).join("");
}

function renderPCAVarianceBars() {
  const container = document.getElementById("pca-variance-bars");
  container.innerHTML = APP.pcaResult.explainedVar.map((v, i) => `
    <div class="var-bar-row">
      <span class="var-bar-lbl">PC${i+1}</span>
      <div class="var-bar-track">
        <div class="var-bar-fill" style="width:${v.toFixed(1)}%"></div>
      </div>
      <span class="var-bar-pct">${v.toFixed(1)}%</span>
    </div>`).join("");
}

function renderPCAChart(type) {
  const wrap = document.getElementById("pca-chart-wrap");
  const pcX = parseInt(document.getElementById("pca-pc-x").value) || 0;
  const pcY = parseInt(document.getElementById("pca-pc-y").value) || 1;
  const xHeaders = APP.rawData ? APP.rawData.headers : [];
  const { pcaResult } = APP;

  let svg = "";
  if (type === "score") {
    svg = ChemCharts.buildScorePlot(pcaResult.scores, APP.rawData?.sampleNames, pcX, pcY, pcaResult.explainedVar);
    APP.charts.scorePlot = svg;
  } else if (type === "loading") {
    svg = ChemCharts.buildLoadingPlot(pcaResult.loadings, xHeaders, pcX, pcY, pcaResult.explainedVar);
    APP.charts.loadingPlot = svg;
  } else if (type === "scree") {
    svg = ChemCharts.buildScreePlot(pcaResult.explainedVar, pcaResult.cumulativeVar);
    APP.charts.screePlot = svg;
  } else if (type === "biplot") {
    // Biplot = score + loadings superposés (score plot avec vecteurs)
    svg = ChemCharts.buildScorePlot(pcaResult.scores, APP.rawData?.sampleNames, pcX, pcY, pcaResult.explainedVar);
    APP.charts.biplot = svg;
  }

  wrap.innerHTML = svg || "<p style='color:var(--text-dim);padding:20px;text-align:center'>Graphique non disponible</p>";
}

function renderPCALoadingsTable() {
  const xHeaders = APP.rawData?.headers || [];
  const thead = document.getElementById("pca-loadings-thead");
  const tbody = document.getElementById("pca-loadings-tbody");
  const nComp = APP.pcaResult.nComp;

  thead.innerHTML = `<tr><th>Variable</th>${Array.from({length: nComp}, (_, i) => `<th>PC${i+1}</th>`).join("")}</tr>`;
  tbody.innerHTML = xHeaders.map((name, j) => {
    const loadVals = APP.pcaResult.loadings.map(l => l[j] || 0);
    const maxAbs = Math.max(...loadVals.map(Math.abs));
    return `<tr>
      <td style="color:var(--cyan)">${name}</td>
      ${loadVals.map(v => {
        const intensity = maxAbs > 0 ? Math.abs(v) / maxAbs : 0;
        const col = intensity > 0.6 ? "var(--orange)" : intensity > 0.3 ? "var(--cyan)" : "var(--text-dim)";
        return `<td style="color:${col}">${v.toFixed(4)}</td>`;
      }).join("")}
    </tr>`;
  }).join("");
}

async function handlePCAToExcel() {
  if (!APP.pcaResult) { toast("Lancez la PCA d'abord", "warn"); return; }
  setBtnLoading("btn-pca-to-excel", true, "Export…");
  try {
    await ChemExcel.exportPCAResults(APP.pcaResult, APP.rawData?.sampleNames, APP.rawData?.headers);
    toast("✅ Résultats PCA exportés dans 'ChemAI_PCA'", "ok");
  } catch (e) { toast("Erreur export : " + e.message, "err"); }
  setBtnLoading("btn-pca-to-excel", false, "⊞ Exporter vers Excel");
}

async function handleAIPCA() {
  if (!APP.pcaResult) { toast("Lancez la PCA d'abord", "warn"); return; }
  if (!ChemAI.hasKey()) { toast("Configurez la clé API Gemini dans le panneau Rapport", "warn"); return; }
  setBtnLoading("btn-ai-pca", true, "Analyse IA…");
  const box = document.getElementById("ai-pca-result");
  box.style.display = "block";
  box.innerHTML = '<span class="spinner"></span> Gemini analyse votre PCA…';
  try {
    const text = await ChemAI.interpretPCA(APP.pcaResult, APP.rawData?.headers || [], APP.rawData?.sampleNames || []);
    box.innerHTML = ChemAI.formatResponse(text);
    toast("✅ Interprétation PCA générée", "ok");
  } catch (e) {
    box.innerHTML = `<span style="color:var(--red)">❌ ${e.message}</span>`;
  }
  setBtnLoading("btn-ai-pca", false, "✦ Interpréter la PCA");
}

// ─── PANEL : PLS ─────────────────────────────────────────────────────────────
function setupPLSPanel() {
  document.getElementById("btn-run-pls").addEventListener("click", handleRunPLS);
  document.getElementById("btn-pls-to-excel").addEventListener("click", handlePLSToExcel);
  document.getElementById("btn-ai-pls").addEventListener("click", handleAIPLS);
  document.getElementById("btn-pls-predict").addEventListener("click", handlePLSPredict);

  document.querySelectorAll('[data-chart]').forEach(btn => {
    if (btn.closest("#p-pls")) {
      btn.addEventListener("click", () => {
        btn.closest(".chart-tabs").querySelectorAll(".ctab").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        renderPLSChart(btn.dataset.chart);
      });
    }
  });
}

async function handleRunPLS() {
  setBtnLoading("btn-run-pls", true, "Modélisation PLS…");
  try {
    const X = getXData();
    const yColIdx = parseInt(document.getElementById("pls-y-col").value);
    const nCompSel = document.getElementById("pls-ncomp").value;
    const cvK = document.getElementById("pls-cv").value;
    const splitRatio = parseFloat(document.getElementById("pls-split").value);

    // Extraire Y depuis les données brutes
    const rawY = APP.rawData.data.map(row => row[yColIdx]);
    const yMu = rawY.reduce((s, v) => s + v, 0) / rawY.length;
    const yStd = Math.sqrt(rawY.reduce((s, v) => s + (v - yMu) ** 2, 0) / (rawY.length - 1)) || 1;
    const yStd_data = rawY.map(v => (v - yMu) / yStd);

    // Validation croisée pour choisir nComp optimal
    const maxComp = Math.min(10, Math.min(X.length - 1, X[0].length));
    const k_cv = cvK === "loo" ? X.length : parseInt(cvK);
    const rmsecvArr = ChemMath.PLS.crossValidate(X, yStd_data, maxComp, Math.min(k_cv, X.length));
    const optNComp = nCompSel === "auto"
      ? rmsecvArr.indexOf(Math.min(...rmsecvArr)) + 1
      : parseInt(nCompSel);

    // Entraîner le modèle final
    const splitIdx = Math.floor(X.length * (1 - splitRatio));
    const Xtrain = X.slice(0, splitIdx), ytrain = yStd_data.slice(0, splitIdx);
    const Xtest  = X.slice(splitIdx),   ytest  = yStd_data.slice(splitIdx);

    APP.plsModel = ChemMath.PLS.fit(Xtrain, ytrain, optNComp);
    if (!APP.plsModel) throw new Error("La régression PLS a échoué (matrice singulière ?)");

    // Métriques
    const yPredTrain = ChemMath.PLS.predict(APP.plsModel, Xtrain, yMu, yStd);
    const yPredTest  = ChemMath.PLS.predict(APP.plsModel, Xtest, yMu, yStd);
    const yTrainReal = ytrain.map(v => v * yStd + yMu);
    const yTestReal  = ytest.map(v => v * yStd + yMu);

    const r2 = (y, yHat) => {
      const mu = y.reduce((s, v) => s + v, 0) / y.length;
      const sst = y.reduce((s, v) => s + (v - mu) ** 2, 0);
      const sse = y.reduce((s, v, i) => s + (v - yHat[i]) ** 2, 0);
      return 1 - sse / (sst || 1);
    };
    const rmse = (y, yHat) => Math.sqrt(y.reduce((s, v, i) => s + (v - yHat[i]) ** 2, 0) / y.length);

    APP.plsMetrics = {
      r2_cal:  r2(yTrainReal, yPredTrain),
      r2_cv:   r2(yTestReal, yPredTest),
      rmsec:   rmse(yTrainReal, yPredTrain),
      rmsecv:  rmse(yTestReal, yPredTest),
      nComp:   optNComp,
    };

    // Stocker pour les graphiques
    const yAllPred = ChemMath.PLS.predict(APP.plsModel, X, yMu, yStd);
    APP.plsYReal = rawY;
    APP.plsYPred = yAllPred;

    // Afficher métriques
    document.getElementById("pls-metrics").innerHTML = [
      ["R² calibration",  APP.plsMetrics.r2_cal.toFixed(4)],
      ["R² validation",   APP.plsMetrics.r2_cv.toFixed(4)],
      ["RMSEC",           APP.plsMetrics.rmsec.toFixed(4)],
      ["RMSECV",          APP.plsMetrics.rmsecv.toFixed(4)],
      ["Composantes LV",  optNComp],
    ].map(([l, v]) => `<div class="stat-item"><div class="stat-lbl">${l}</div><div class="stat-val">${v}</div></div>`).join("");

    renderPLSChart("pls-pred");
    document.getElementById("pls-results").style.display = "block";
    toast(`✅ Modèle PLS : R²cv = ${APP.plsMetrics.r2_cv.toFixed(4)}, RMSECV = ${APP.plsMetrics.rmsecv.toFixed(4)}`, "ok");
    log(`PLS ${optNComp} LV — R²cv=${APP.plsMetrics.r2_cv.toFixed(4)}, RMSECV=${APP.plsMetrics.rmsecv.toFixed(4)}`, "ok");
  } catch (e) {
    toast("Erreur PLS : " + e.message, "err");
    log("Erreur PLS : " + e.message, "err");
    console.error(e);
  }
  setBtnLoading("btn-run-pls", false, "∿ Construire le modèle PLS");
}

function renderPLSChart(type) {
  const wrap = document.getElementById("pls-chart-wrap");
  const xHeaders = APP.rawData?.headers || [];
  let svg = "";

  if (type === "pls-pred") {
    svg = ChemCharts.buildPredVsReal(APP.plsYReal, APP.plsYPred, APP.plsMetrics.r2_cv, APP.plsMetrics.rmsecv);
    APP.charts.predVsReal = svg;
  } else if (type === "pls-vip" && APP.plsModel) {
    svg = ChemCharts.buildVIPChart(APP.plsModel.vip, xHeaders);
    APP.charts.vipChart = svg;
  } else if (type === "pls-resid") {
    svg = ChemCharts.buildResidualPlot(APP.plsYPred, APP.plsYReal, APP.rawData?.sampleNames);
    APP.charts.residualPlot = svg;
  }

  wrap.innerHTML = svg || "<p style='color:var(--text-dim);padding:20px;text-align:center'>Graphique non disponible</p>";
}

async function handlePLSToExcel() {
  if (!APP.plsModel) { toast("Lancez le PLS d'abord", "warn"); return; }
  setBtnLoading("btn-pls-to-excel", true, "Export…");
  try {
    await ChemExcel.exportPLSResults(APP.plsModel, APP.plsYReal, APP.plsYPred,
      APP.plsModel.vip, APP.rawData?.headers || [], APP.rawData?.sampleNames || []);
    toast("✅ Résultats PLS exportés dans 'ChemAI_PLS'", "ok");
  } catch (e) { toast("Erreur export : " + e.message, "err"); }
  setBtnLoading("btn-pls-to-excel", false, "⊞ Exporter vers Excel");
}

async function handlePLSPredict() {
  if (!APP.plsModel) { toast("Construisez d'abord un modèle PLS", "warn"); return; }
  setBtnLoading("btn-pls-predict", true, "Prédiction…");
  try {
    const parsed = await ChemExcel.readSelection(true);
    const Xnew = parsed.data;
    const scalerP = APP.scalerParams || ChemMath.Preprocessing.fitScaler(Xnew, "none");
    const Xscaled = ChemMath.Preprocessing.transform(Xnew, scalerP);
    const yPred = ChemMath.PLS.predict(APP.plsModel, Xscaled);

    // Écrire les prédictions dans une nouvelle colonne
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const sel = ctx.workbook.getSelectedRange();
      sel.load("columnCount,rowIndex,columnIndex,rowCount");
      await ctx.sync();
      const predRange = ws.getRangeByIndexes(sel.rowIndex, sel.columnIndex + sel.columnCount, yPred.length + 1, 1);
      predRange.values = [["Ŷ prédit"], ...yPred.map(v => [+v.toFixed(4)])];
      predRange.getCell(0, 0).format.font.bold = true;
      predRange.getCell(0, 0).format.font.color = "#00E676";
      await ctx.sync();
    });
    toast(`✅ ${yPred.length} prédictions écrites`, "ok");
  } catch (e) { toast("Erreur prédiction : " + e.message, "err"); }
  setBtnLoading("btn-pls-predict", false, "↓ Prédire la sélection");
}

async function handleAIPLS() {
  if (!APP.plsModel || !APP.plsMetrics) { toast("Lancez le PLS d'abord", "warn"); return; }
  if (!ChemAI.hasKey()) { toast("Configurez la clé API Gemini", "warn"); return; }
  setBtnLoading("btn-ai-pls", true, "Analyse IA…");
  const box = document.getElementById("ai-pls-result");
  box.style.display = "block";
  box.innerHTML = '<span class="spinner"></span> Gemini analyse votre modèle PLS…';
  try {
    const text = await ChemAI.interpretPLS(APP.plsMetrics, APP.plsModel.vip, APP.rawData?.headers || [], APP.plsMetrics.nComp);
    box.innerHTML = ChemAI.formatResponse(text);
    toast("✅ Interprétation PLS générée", "ok");
  } catch (e) { box.innerHTML = `<span style="color:var(--red)">❌ ${e.message}</span>`; }
  setBtnLoading("btn-ai-pls", false, "✦ Interpréter le modèle PLS");
}

// ─── PANEL : CLUSTERING ──────────────────────────────────────────────────────
function setupClusterPanel() {
  document.getElementById("btn-run-cluster").addEventListener("click", handleRunCluster);
  document.getElementById("btn-cluster-to-excel").addEventListener("click", handleClusterToExcel);
  document.getElementById("btn-ai-cluster").addEventListener("click", handleAICluster);

  document.querySelectorAll('[data-chart]').forEach(btn => {
    if (btn.closest("#p-cluster")) {
      btn.addEventListener("click", () => {
        btn.closest(".chart-tabs").querySelectorAll(".ctab").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        renderClusterChart(btn.dataset.chart);
      });
    }
  });
}

async function handleRunCluster() {
  setBtnLoading("btn-run-cluster", true, "Clustering…");
  try {
    const X = getXData();
    const method = document.getElementById("cluster-method").value;
    const kSel   = document.getElementById("cluster-k").value;
    const linkage = document.getElementById("cluster-linkage").value;

    // Utiliser l'espace PCA si disponible
    const useSpace = document.getElementById("cluster-space").value;
    const Xcluster = (useSpace === "pca" && APP.pcaResult)
      ? APP.pcaResult.scores.map(s => Array.from(s).slice(0, Math.min(5, APP.pcaResult.nComp)))
      : X;

    let k = kSel === "auto" ? null : parseInt(kSel);

    if (method === "kmeans" || method === "both") {
      if (!k) {
        // Elbow automatique
        const elbowData = ChemMath.KMeans.elbow(Xcluster, 8);
        APP.charts.elbowChart = ChemCharts.buildElbowChart(elbowData);
        // Dérivée secondaire pour trouver le coude
        const inertias = elbowData.map(d => d.inertia);
        const diffs2 = inertias.slice(2).map((v, i) => inertias[i] - 2*inertias[i+1] + v);
        k = diffs2.indexOf(Math.max(...diffs2)) + 2;
        toast(`ℹ Elbow automatique : k = ${k}`, "info");
      }
      APP.clusterResult = ChemMath.KMeans.fit(Xcluster, k);
      APP.clusterResult.method = "K-Means";
    }

    if (method === "hierarchical" || method === "both") {
      const hResult = ChemMath.Hierarchical.fit(Xcluster, linkage);
      const kH = k || 3;
      const hLabels = ChemMath.Hierarchical.cutTree(hResult.linkageMatrix, kH, Xcluster.length);
      APP.charts.dendro = ChemCharts.buildDendrogram(hResult.linkageMatrix, APP.rawData?.sampleNames, kH);

      if (method === "hierarchical") {
        APP.clusterResult = { labels: hLabels, k: kH, inertia: null, silhouette: null, centroids: [], method: "Hiérarchique" };
      }
    }

    // Heatmap des données
    APP.charts.heatmap = ChemCharts.buildHeatmap(X, APP.rawData?.sampleNames, APP.rawData?.headers || []);

    // Métriques
    document.getElementById("cluster-metrics").innerHTML = [
      ["Clusters",    APP.clusterResult.k],
      ["Silhouette",  APP.clusterResult.silhouette?.toFixed(3) || "N/A"],
      ["Inertie",     APP.clusterResult.inertia?.toFixed(1) || "N/A"],
      ["Méthode",     APP.clusterResult.method],
    ].map(([l, v]) => `<div class="stat-item"><div class="stat-lbl">${l}</div><div class="stat-val">${v}</div></div>`).join("");

    // Score plot coloré par cluster
    if (APP.pcaResult) {
      APP.charts.clusterScatter = ChemCharts.buildScorePlot(
        APP.pcaResult.scores, APP.rawData?.sampleNames, 0, 1,
        APP.pcaResult.explainedVar, APP.clusterResult.labels
      );
    }

    renderClusterChart("cluster-scatter");
    document.getElementById("cluster-results").style.display = "block";
    toast(`✅ Clustering : ${APP.clusterResult.k} clusters, Silhouette = ${APP.clusterResult.silhouette?.toFixed(3) || 'N/A'}`, "ok");
    log(`Clustering ${APP.clusterResult.method} : k=${APP.clusterResult.k}`, "ok");
  } catch (e) {
    toast("Erreur clustering : " + e.message, "err");
    log("Erreur : " + e.message, "err");
    console.error(e);
  }
  setBtnLoading("btn-run-cluster", false, "⬡ Lancer le clustering");
}

function renderClusterChart(type) {
  const wrap = document.getElementById("cluster-chart-wrap");
  const chartMap = {
    "cluster-scatter": APP.charts.clusterScatter,
    "dendro":          APP.charts.dendro,
    "heatmap":         APP.charts.heatmap,
    "elbow":           APP.charts.elbowChart,
  };
  wrap.innerHTML = chartMap[type] || "<p style='color:var(--text-dim);padding:20px;text-align:center'>Graphique non disponible. Lancez l'analyse correspondante.</p>";
}

async function handleClusterToExcel() {
  if (!APP.clusterResult) { toast("Lancez le clustering d'abord", "warn"); return; }
  setBtnLoading("btn-cluster-to-excel", true, "Export…");
  try {
    await ChemExcel.exportClusterLabels(APP.clusterResult.labels, APP.rawData?.sampleNames || []);
    toast("✅ Labels exportés dans 'ChemAI_Clusters'", "ok");
  } catch (e) { toast("Erreur export : " + e.message, "err"); }
  setBtnLoading("btn-cluster-to-excel", false, "⊞ Exporter labels vers Excel");
}

async function handleAICluster() {
  if (!APP.clusterResult) { toast("Lancez le clustering d'abord", "warn"); return; }
  if (!ChemAI.hasKey()) { toast("Configurez la clé API Gemini", "warn"); return; }
  setBtnLoading("btn-ai-cluster", true, "Analyse IA…");
  const box = document.getElementById("ai-cluster-result");
  box.style.display = "block";
  box.innerHTML = '<span class="spinner"></span> Gemini analyse vos clusters…';
  try {
    const text = await ChemAI.interpretClusters(APP.clusterResult, APP.rawData?.headers || [], APP.clusterResult.method);
    box.innerHTML = ChemAI.formatResponse(text);
    toast("✅ Interprétation clustering générée", "ok");
  } catch (e) { box.innerHTML = `<span style="color:var(--red)">❌ ${e.message}</span>`; }
  setBtnLoading("btn-ai-cluster", false, "✦ Interpréter les clusters");
}

// ─── PANEL : ANOMALIES ───────────────────────────────────────────────────────
function setupAnomalyPanel() {
  document.getElementById("btn-run-anomaly").addEventListener("click", handleRunAnomaly);
  document.getElementById("btn-anomaly-to-excel").addEventListener("click", handleAnomalyToExcel);
  document.getElementById("btn-ai-anomaly").addEventListener("click", handleAIAnomaly);
}

async function handleRunAnomaly() {
  setBtnLoading("btn-run-anomaly", true, "Détection…");
  try {
    const X = getXData();
    const method = document.getElementById("anomaly-method").value;
    const alpha  = parseFloat(document.getElementById("anomaly-alpha").value);

    let scores, threshold, outliers;

    if (method === "mahalanobis") {
      scores = ChemMath.Anomaly.mahalanobis(X);
      threshold = ChemMath.Anomaly.mahalanobisThreshold(X[0].length, alpha);
      outliers = scores.map(s => s > threshold);
    } else if ((method === "hotelling" || method === "q_residual" || method === "combined") && APP.pcaResult) {
      const result = ChemMath.Anomaly.fromPCA(APP.pcaResult, alpha);
      if (method === "hotelling") {
        scores = result.T2;
        threshold = result.T2Limit;
      } else if (method === "q_residual") {
        scores = result.qScores;
        threshold = result.qLimit;
      } else {
        scores = result.T2.map((t, i) => Math.max(t / result.T2Limit, result.qScores[i] / result.qLimit));
        threshold = 1;
      }
      outliers = result.outliers;
    } else {
      scores = ChemMath.Anomaly.mahalanobis(X);
      threshold = ChemMath.Anomaly.mahalanobisThreshold(X[0].length, alpha);
      outliers = scores.map(s => s > threshold);
    }

    APP.anomalyResult = { scores, threshold, outliers, method };

    const nOutliers = outliers.filter(Boolean).length;

    // Résumé
    document.getElementById("anomaly-summary").innerHTML = `
      <div class="anomaly-badge ${nOutliers > 0 ? 'alert' : ''}">
        <strong>${nOutliers}</strong> outlier${nOutliers !== 1 ? 's' : ''} détecté${nOutliers !== 1 ? 's' : ''}
      </div>
      <div class="anomaly-badge">
        Seuil (α=${alpha}) : <strong>${threshold.toFixed(3)}</strong>
      </div>
      <div class="anomaly-badge">
        Score max : <strong>${Math.max(...scores).toFixed(3)}</strong>
      </div>`;

    // Graphique
    const svg = ChemCharts.buildControlChart(scores, threshold, APP.rawData?.sampleNames, `Détection — ${method}`);
    APP.charts.controlChart = svg;
    document.getElementById("anomaly-chart-wrap").innerHTML = svg;

    // Table
    const tbody = document.getElementById("anomaly-tbody");
    const sampleNames = APP.rawData?.sampleNames || [];
    tbody.innerHTML = scores
      .map((s, i) => ({ s, i, isOut: outliers[i] }))
      .sort((a, b) => b.s - a.s)
      .slice(0, 20)
      .map(({ s, i, isOut }) => `<tr style="${isOut ? 'color:var(--red)' : ''}">
        <td>${sampleNames[i] || `S${i+1}`}</td>
        <td style="font-family:var(--font-mono)">${s.toFixed(4)}</td>
        <td style="font-family:var(--font-mono)">${threshold.toFixed(4)}</td>
        <td>${isOut ? '⚠ OUTLIER' : '✓ Normal'}</td>
      </tr>`).join("");

    document.getElementById("anomaly-results").style.display = "block";
    toast(`✅ ${nOutliers} outlier(s) détecté(s) sur ${scores.length} échantillons`, nOutliers > 0 ? "warn" : "ok");
    log(`Anomalies (${method}) : ${nOutliers}/${scores.length} outliers, seuil=${threshold.toFixed(3)}`, "ok");
  } catch (e) {
    toast("Erreur : " + e.message, "err");
    log("Erreur anomalies : " + e.message, "err");
    console.error(e);
  }
  setBtnLoading("btn-run-anomaly", false, "⚠ Détecter les anomalies");
}

async function handleAnomalyToExcel() {
  if (!APP.anomalyResult) { toast("Lancez la détection d'abord", "warn"); return; }
  setBtnLoading("btn-anomaly-to-excel", true, "Marquage…");
  try {
    const outlierIdx = APP.anomalyResult.outliers.map((o, i) => o ? i : -1).filter(i => i >= 0);
    if (outlierIdx.length > 0) await ChemExcel.markOutliers(outlierIdx, "A1");
    toast(`✅ ${outlierIdx.length} outlier(s) marqués en rouge`, "ok");
  } catch (e) { toast("Erreur : " + e.message, "err"); }
  setBtnLoading("btn-anomaly-to-excel", false, "⊞ Marquer dans Excel");
}

async function handleAIAnomaly() {
  if (!APP.anomalyResult) { toast("Lancez la détection d'abord", "warn"); return; }
  if (!ChemAI.hasKey()) { toast("Configurez la clé API Gemini", "warn"); return; }
  setBtnLoading("btn-ai-anomaly", true, "Analyse IA…");
  const box = document.getElementById("ai-anomaly-result");
  box.style.display = "block";
  box.innerHTML = '<span class="spinner"></span> Gemini analyse les anomalies…';
  try {
    const text = await ChemAI.interpretAnomalies(APP.anomalyResult, APP.rawData?.sampleNames || [], APP.anomalyResult.method);
    box.innerHTML = ChemAI.formatResponse(text);
    toast("✅ Interprétation générée", "ok");
  } catch (e) { box.innerHTML = `<span style="color:var(--red)">❌ ${e.message}</span>`; }
  setBtnLoading("btn-ai-anomaly", false, "✦ Analyser les anomalies");
}

// ─── PANEL : RAPPORT ─────────────────────────────────────────────────────────
function setupReportPanel() {
  document.getElementById("btn-save-key").addEventListener("click", () => {
    const key = document.getElementById("ai-api-key").value.trim();
    if (!key) { toast("Saisissez votre clé API", "warn"); return; }
    ChemAI.setApiKey(key);
    toast("✅ Clé API sauvegardée", "ok");
  });

  document.getElementById("btn-report-html").addEventListener("click", handleGenerateReport);
  document.getElementById("btn-report-excel").addEventListener("click", handleExcelReport);
}

async function handleGenerateReport() {
  setBtnLoading("btn-report-html", true, "Génération…");
  try {
    const opts = gatherReportOptions();

    // Interprétation IA globale si clé disponible
    let globalAI = "";
    if (ChemAI.hasKey() && opts.ai) {
      try {
        globalAI = await ChemAI.globalInterpretation({
          nSamples: APP.rawData?.nRows || 0,
          nVars:    APP.rawData?.nCols || 0,
          varNames: APP.rawData?.headers || [],
          pca: APP.pcaResult ? {
            nComp: APP.pcaResult.nComp,
            totalVar: APP.pcaResult.cumulativeVar[APP.pcaResult.nComp-1],
            pc1Vars: APP.rawData?.headers?.slice(0,3).join(', '),
          } : null,
          pls: APP.plsMetrics ? {
            r2cv:    APP.plsMetrics.r2_cv,
            rmsecv:  APP.plsMetrics.rmsecv,
            vipVars: APP.rawData?.headers?.slice(0,3).join(', '),
          } : null,
          clustering: APP.clusterResult ? {
            k: APP.clusterResult.k,
            silhouette: APP.clusterResult.silhouette,
          } : null,
          anomalies: APP.anomalyResult ? {
            nOutliers: APP.anomalyResult.outliers.filter(Boolean).length,
          } : null,
        });
      } catch (e) {
        globalAI = `Interprétation IA indisponible : ${e.message}`;
      }
    }

    const html = generateHTMLReport(opts, globalAI);
    downloadHTML(html, `ChemAI_Rapport_${new Date().toISOString().slice(0,10)}.html`);
    toast("✅ Rapport HTML téléchargé", "ok");
    logReport("Rapport HTML généré", "ok");
  } catch (e) {
    toast("Erreur rapport : " + e.message, "err");
    logReport("Erreur : " + e.message, "err");
    console.error(e);
  }
  setBtnLoading("btn-report-html", false, "🌐 Rapport HTML");
}

function gatherReportOptions() {
  return {
    labo:     document.getElementById("rpt-labo").value,
    auteur:   document.getElementById("rpt-auteur").value,
    ref:      document.getElementById("rpt-ref").value,
    version:  document.getElementById("rpt-version").value,
    data:     document.getElementById("rpt-data").checked,
    preproc:  document.getElementById("rpt-preproc").checked,
    pca:      document.getElementById("rpt-pca").checked,
    pls:      document.getElementById("rpt-pls").checked,
    cluster:  document.getElementById("rpt-cluster").checked,
    anomaly:  document.getElementById("rpt-anomaly").checked,
    ai:       document.getElementById("rpt-ai").checked,
  };
}

function generateHTMLReport(opts, globalAI) {
  const date = new Date().toLocaleDateString("fr-FR");
  const time = new Date().toLocaleTimeString("fr-FR");

  const section = (title, content) => content ? `
    <section class="report-section">
      <h2>${title}</h2>
      ${content}
    </section>` : '';

  const statBox = (items) => `
    <div class="stat-row">
      ${items.map(([l, v]) => `<div class="stat-box"><div class="sl">${l}</div><div class="sv">${v}</div></div>`).join('')}
    </div>`;

  const svgSection = (svgStr, caption) => svgStr ? `
    <div class="chart-container">
      ${svgStr}
      ${caption ? `<p class="caption">${caption}</p>` : ''}
    </div>` : '';

  return `<!DOCTYPE html>
<html lang="fr"><head>
<meta charset="UTF-8">
<title>Rapport ChemAI — ${opts.ref || 'Analyse chimiométrique'}</title>
<style>
  :root{--cyan:#00E5FF;--green:#00E676;--orange:#FF9800;--bg:#080F1A;--bg2:#0D1825;--bg3:#122035;--border:#1A2E47;--text:#D0E4F0;--text-dim:#7B9DB8}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:var(--text);background:var(--bg);padding:30px 40px;max-width:1100px;margin:auto;line-height:1.6}
  h1{font-size:22px;color:var(--cyan);border-bottom:2px solid var(--cyan);padding-bottom:10px;margin-bottom:6px}
  h2{font-size:12px;font-weight:700;color:var(--text-dim);margin:24px 0 10px;border-left:4px solid var(--cyan);padding-left:10px;text-transform:uppercase;letter-spacing:.08em}
  .meta-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;background:var(--bg2);border-radius:8px;padding:16px;margin:16px 0}
  .meta-grid .ml{font-size:9px;color:#3A5570;text-transform:uppercase;letter-spacing:.05em}
  .meta-grid .mv{font-weight:700;font-size:13px;color:var(--cyan);margin-top:2px}
  .report-section{margin-bottom:30px;padding-bottom:20px;border-bottom:1px solid var(--border)}
  table{width:100%;border-collapse:collapse;margin:8px 0;font-size:11px}
  th{background:var(--bg);color:var(--cyan);padding:6px 8px;text-align:left;font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;border-bottom:1px solid var(--border)}
  td{padding:5px 8px;border-bottom:1px solid var(--border);color:var(--text-dim);font-family:'Courier New',monospace}
  tr:nth-child(even) td{background:var(--bg2)}
  .stat-row{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0}
  .stat-box{background:var(--bg3);border-radius:6px;padding:10px 14px;min-width:100px;text-align:center}
  .sl{font-size:9px;color:#3A5570;text-transform:uppercase;letter-spacing:.05em}
  .sv{font-family:'Courier New',monospace;font-size:15px;font-weight:700;color:var(--cyan);margin-top:4px}
  .chart-container{margin:12px 0;border:1px solid var(--border);border-radius:6px;overflow:hidden}
  .caption{font-size:10px;color:var(--text-dim);padding:6px 10px;background:var(--bg2);font-style:italic}
  .ai-content{background:var(--bg3);border-left:3px solid var(--cyan);padding:14px;border-radius:6px;font-size:12px;line-height:1.8;margin:10px 0}
  .footer{margin-top:40px;padding-top:12px;border-top:1px solid var(--border);font-size:10px;color:#3A5570;text-align:center}
  @media print{body{background:#fff;color:#1a1a2e}h1,h2{color:#0D1B2A;border-color:#0D1B2A}.stat-box,.chart-container{border:1px solid #ccc}}
</style>
</head><body>

<h1>Rapport d'Analyse Chimiométrique</h1>
<p style="font-size:11px;color:var(--text-dim);margin-bottom:12px">Analyse multivariée — PCA · PLS · Clustering · Détection d'anomalies</p>

<div class="meta-grid">
  ${[['Laboratoire', opts.labo||'—'],['Auteur', opts.auteur||'—'],['Référence', opts.ref||'—'],['Version', opts.version||'1.0'],
     ['Échantillons', APP.rawData?.nRows||'—'],['Variables', APP.rawData?.nCols||'—'],['Prétraitement', APP.scalingMethod||'aucun'],['Date', date]]
    .map(([l,v])=>`<div><div class="ml">${l}</div><div class="mv">${v}</div></div>`).join('')}
</div>

${opts.data && APP.rawData ? section("1. Description des données", `
  <p>Jeu de données : <strong>${APP.rawData.nRows}</strong> échantillons × <strong>${APP.rawData.nCols}</strong> variables.</p>
  <p>Variables : ${APP.rawData.headers.join(', ')}.</p>
  <p>Méthode de mise à l'échelle : <strong>${APP.scalingMethod || 'aucune'}</strong>. Gestion des NaN : <strong>${APP.missingStrategy}</strong>.</p>
`) : ''}

${opts.pca && APP.pcaResult ? section("2. Analyse en Composantes Principales (PCA)", `
  ${statBox([
    ['Composantes', APP.pcaResult.nComp],
    ['Variance PC1', APP.pcaResult.explainedVar[0]?.toFixed(1)+'%'],
    ['Variance PC2', APP.pcaResult.explainedVar[1]?.toFixed(1)+'%'],
    ['Variance totale', APP.pcaResult.cumulativeVar[APP.pcaResult.nComp-1]?.toFixed(1)+'%'],
  ])}
  <table>
    <tr><th>Composante</th><th>Valeur propre</th><th>Variance (%)</th><th>Variance cumulée (%)</th></tr>
    ${APP.pcaResult.explainedVar.map((v,i)=>`<tr><td>PC${i+1}</td><td>${APP.pcaResult.eigenvalues[i]?.toFixed(4)}</td><td>${v.toFixed(2)}%</td><td>${APP.pcaResult.cumulativeVar[i]?.toFixed(2)}%</td></tr>`).join('')}
  </table>
  ${svgSection(APP.charts.scorePlot, 'Score Plot — Projection des échantillons dans l\'espace des composantes principales')}
  ${svgSection(APP.charts.screePlot, 'Scree Plot — Décroissance de la variance expliquée par composante')}
  ${svgSection(APP.charts.loadingPlot, 'Loading Plot — Contribution des variables aux composantes')}
`) : ''}

${opts.pls && APP.plsMetrics ? section("3. Régression PLS-1", `
  ${statBox([
    ['R² calibration', APP.plsMetrics.r2_cal?.toFixed(4)],
    ['R² validation', APP.plsMetrics.r2_cv?.toFixed(4)],
    ['RMSEC', APP.plsMetrics.rmsec?.toFixed(4)],
    ['RMSECV', APP.plsMetrics.rmsecv?.toFixed(4)],
    ['LV', APP.plsMetrics.nComp],
  ])}
  ${svgSection(APP.charts.predVsReal, 'Graphique Prédit vs Réel — Qualité de la prédiction')}
  ${svgSection(APP.charts.vipChart, 'VIP Scores — Variables les plus informatives (seuil VIP ≥ 1)')}
`) : ''}

${opts.cluster && APP.clusterResult ? section("4. Clustering", `
  ${statBox([
    ['Clusters k', APP.clusterResult.k],
    ['Méthode', APP.clusterResult.method],
    ['Silhouette', APP.clusterResult.silhouette?.toFixed(3)||'N/A'],
    ['Inertie', APP.clusterResult.inertia?.toFixed(1)||'N/A'],
  ])}
  ${svgSection(APP.charts.clusterScatter, 'Scatter plot des clusters dans l\'espace PCA')}
  ${svgSection(APP.charts.dendro, 'Dendrogramme — Structure hiérarchique des groupes')}
  ${svgSection(APP.charts.heatmap, 'Heatmap — Vue d\'ensemble des données')}
`) : ''}

${opts.anomaly && APP.anomalyResult ? section("5. Détection d'anomalies", `
  ${statBox([
    ['Outliers', APP.anomalyResult.outliers.filter(Boolean).length],
    ['Méthode', APP.anomalyResult.method],
    ['Seuil', APP.anomalyResult.threshold?.toFixed(3)],
    ['Score max', Math.max(...APP.anomalyResult.scores).toFixed(3)],
  ])}
  ${svgSection(APP.charts.controlChart, 'Graphique de contrôle — Scores vs seuil statistique')}
`) : ''}

${opts.ai && globalAI ? section("6. Interprétation par Intelligence Artificielle", `
  <div class="ai-content">${ChemAI.formatResponse(globalAI)}</div>
`) : ''}

<div class="footer">
  Rapport généré par <strong>ChemAI Add-in</strong> — ${date} à ${time}<br>
  ${opts.ref ? `Réf. : ${opts.ref} · ` : ''}${opts.auteur ? `Auteur : ${opts.auteur} · ` : ''}Version ${opts.version||'1.0'}
</div>
</body></html>`;
}

async function handleExcelReport() {
  setBtnLoading("btn-report-excel", true, "Export Excel…");
  try {
    if (APP.pcaResult) await ChemExcel.exportPCAResults(APP.pcaResult, APP.rawData?.sampleNames, APP.rawData?.headers);
    if (APP.plsModel)  await ChemExcel.exportPLSResults(APP.plsModel, APP.plsYReal, APP.plsYPred, APP.plsModel.vip, APP.rawData?.headers, APP.rawData?.sampleNames);
    if (APP.clusterResult) await ChemExcel.exportClusterLabels(APP.clusterResult.labels, APP.rawData?.sampleNames);
    toast("✅ Classeur complet exporté", "ok");
  } catch (e) { toast("Erreur : " + e.message, "err"); }
  setBtnLoading("btn-report-excel", false, "⊞ Classeur Excel complet");
}

// ─── Utilitaires ─────────────────────────────────────────────────────────────

function downloadHTML(html, filename) {
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

function switchPanel(id) {
  document.querySelectorAll(".ntab, .panel").forEach(el => el.classList.remove("active"));
  document.querySelector(`[data-panel="${id}"]`)?.classList.add("active");
  document.getElementById(id)?.classList.add("active");
}

function toast(msg, type, dur) {
  const el = document.createElement("div");
  el.className = `toast ${type || "info"}`;
  el.innerHTML = `<span>${{ok:"✅",err:"❌",info:"ℹ",warn:"⚠"}[type]||"ℹ"}</span><span>${msg}</span>`;
  document.getElementById("toast-ct").appendChild(el);
  setTimeout(() => {
    el.style.opacity = "0"; el.style.transform = "translateX(12px)";
    setTimeout(() => el.remove(), 250);
  }, dur || 3500);
}

function log(msg, type) {
  const el = document.getElementById("log-data");
  if (!el) return;
  const e = document.createElement("div");
  e.className = `le ${type||"info"}`;
  e.innerHTML = `<span class="ts">${new Date().toLocaleTimeString("fr-FR")}</span>${msg}`;
  el.appendChild(e);
  el.scrollTop = el.scrollHeight;
}

function logReport(msg, type) {
  const el = document.getElementById("log-report");
  if (!el) return;
  const e = document.createElement("div");
  e.className = `le ${type||"info"}`;
  e.innerHTML = `<span class="ts">${new Date().toLocaleTimeString("fr-FR")}</span>${msg}`;
  el.appendChild(e);
}

function setBtnLoading(id, loading, label) {
  const btn = document.getElementById(id);
  if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading ? `<span class="spinner"></span> ${label}` : label;
}

function setStatus(msg) {
  const el = document.getElementById("hdr-status");
  if (el) el.textContent = msg;
}
