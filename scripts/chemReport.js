/**
 * ================================================================
 * chemReport.js — Générateur de rapport chimiométrique HTML
 *
 * Génère un rapport scientifique complet :
 *  - Section données
 *  - PCA (scores, loadings, scree)
 *  - PLS (métriques, VIP, Prédit vs Réel)
 *  - Clustering (scatter, dendrogramme, heatmap)
 *  - Anomalies (graphique de contrôle)
 *  - Interprétation IA
 *  - Téléchargement HTML
 * ================================================================
 */
"use strict";

const ChemReport = {

  /**
   * Point d'entrée principal
   * @param {object} APP   — état global de l'application
   * @param {object} opts  — options du rapport
   * @param {string} globalAI — texte d'interprétation IA global
   */
  generate(APP, opts, globalAI) {
    const html = this._buildHTML(APP, opts, globalAI);
    const filename = `ChemAI_Rapport_${new Date().toISOString().slice(0, 10)}.html`;
    this._download(html, filename);
    return filename;
  },

  // ─── Construction HTML ────────────────────────────────────────────────────

  _buildHTML(APP, opts, globalAI) {
    const date = new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long", day: "numeric" });
    const time = new Date().toLocaleTimeString("fr-FR");

    return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ChemAI — ${opts.ref || "Rapport chimiométrique"}</title>
${this._styles()}
</head>
<body>

${this._header(opts, date, time)}
${this._toc(opts)}

<div class="content">

${opts.data && APP.rawData ? this._sectionData(APP) : ""}
${opts.preproc            ? this._sectionPreproc(APP) : ""}
${opts.pca && APP.pcaResult ? this._sectionPCA(APP) : ""}
${opts.pls && APP.plsMetrics ? this._sectionPLS(APP) : ""}
${opts.cluster && APP.clusterResult ? this._sectionCluster(APP) : ""}
${opts.anomaly && APP.anomalyResult ? this._sectionAnomaly(APP) : ""}
${opts.ai && globalAI      ? this._sectionAI(globalAI) : ""}
${this._footer(opts, date)}

</div>
</body>
</html>`;
  },

  // ─── En-tête ──────────────────────────────────────────────────────────────

  _header(opts, date, time) {
    const meta = [
      ["Laboratoire", opts.labo || "—"],
      ["Auteur",      opts.auteur || "—"],
      ["Référence",   opts.ref || "—"],
      ["Version",     opts.version || "1.0"],
      ["Date",        date],
      ["Heure",       time],
    ];

    return `
<header class="rpt-header">
  <div class="rpt-logo">⬡ ChemAI</div>
  <div class="rpt-title-block">
    <h1>Rapport d'Analyse Chimiométrique</h1>
    <p class="rpt-subtitle">Analyse multivariée professionnelle — PCA · PLS · Clustering · Anomalies</p>
  </div>
</header>

<div class="meta-grid">
  ${meta.map(([l, v]) => `
    <div class="meta-item">
      <div class="meta-lbl">${l}</div>
      <div class="meta-val">${this._esc(String(v))}</div>
    </div>`).join("")}
</div>
<hr class="rpt-divider">`;
  },

  // ─── Table des matières ───────────────────────────────────────────────────

  _toc(opts) {
    const items = [
      [opts.data,    "§1", "Description des données"],
      [opts.preproc, "§2", "Prétraitement"],
      [opts.pca,     "§3", "Analyse en Composantes Principales (PCA)"],
      [opts.pls,     "§4", "Régression PLS-1"],
      [opts.cluster, "§5", "Clustering"],
      [opts.anomaly, "§6", "Détection d'anomalies"],
      [opts.ai,      "§7", "Interprétation par Intelligence Artificielle"],
    ].filter(([enabled]) => enabled);

    return `
<nav class="toc">
  <div class="toc-title">Table des matières</div>
  <ul>
    ${items.map(([, ref, label]) => `<li><span class="toc-ref">${ref}</span> ${label}</li>`).join("")}
  </ul>
</nav>`;
  },

  // ─── Section : Données ────────────────────────────────────────────────────

  _sectionData(APP) {
    const { rawData } = APP;
    const nanCount   = rawData.data.flat().filter(v => isNaN(v)).length;
    const completePct = ((1 - nanCount / (rawData.nRows * rawData.nCols)) * 100).toFixed(1);

    const statBoxes = [
      ["Échantillons",     rawData.nRows],
      ["Variables",        rawData.nCols],
      ["Valeurs manquantes", nanCount],
      ["Complétude",       completePct + "%"],
    ];

    // Statistiques descriptives en tableau
    const stats = window.ChemMath
      ? ChemMath.Preprocessing.descStats(rawData.data, rawData.headers)
      : [];

    return this._section("1", "Description des données", `

<p>Jeu de données importé : <strong>${rawData.nRows}</strong> échantillons × <strong>${rawData.nCols}</strong> variables.
Variables analysées : ${rawData.headers.slice(0, 10).join(", ")}${rawData.headers.length > 10 ? ` … (+${rawData.headers.length - 10} variables)` : ""}.</p>

${this._statRow(statBoxes)}

${stats.length > 0 ? `
<h3>Statistiques descriptives</h3>
<div class="tscroll">
<table>
  <thead><tr>
    <th>Variable</th><th>n</th><th>Moyenne</th><th>Écart-type</th>
    <th>Min</th><th>Médiane</th><th>Max</th><th>CV (%)</th>
  </tr></thead>
  <tbody>
    ${stats.map(s => `
    <tr>
      <td class="var-name">${this._esc(s.name)}</td>
      <td>${s.n}</td>
      <td>${s.mean.toFixed(4)}</td>
      <td>${s.std.toFixed(4)}</td>
      <td>${s.min.toFixed(4)}</td>
      <td>${s.median.toFixed(4)}</td>
      <td>${s.max.toFixed(4)}</td>
      <td>${(s.std / Math.abs(s.mean || 1) * 100).toFixed(1)}</td>
    </tr>`).join("")}
  </tbody>
</table>
</div>` : ""}
`);
  },

  // ─── Section : Prétraitement ──────────────────────────────────────────────

  _sectionPreproc(APP) {
    const methodLabels = {
      none: "Aucun",
      center: "Centrage (soustraction de la moyenne)",
      standardize: "Autoscaling (z-score, centrage + réduction)",
      normalize: "Normalisation min-max [0, 1]",
      pareto: "Pareto scaling (division par √σ)",
    };
    const method = APP.scalingMethod || "none";

    return this._section("2", "Prétraitement des données", `
<p>Méthode de mise à l'échelle appliquée : <strong>${methodLabels[method] || method}</strong>.</p>
<p>Gestion des valeurs manquantes : <strong>${APP.missingStrategy || "moyenne"}</strong>.</p>
<div class="info-box">
  <strong>Justification :</strong> 
  ${method === "standardize" ? "L'autoscaling (z-score) donne à chaque variable une variance unitaire, ce qui évite que les variables à large amplitude dominent l'analyse. Recommandé lorsque les variables sont en unités différentes." : ""}
  ${method === "pareto" ? "Le Pareto scaling est un compromis entre l'autoscaling et les données brutes. Il réduit l'influence des variables à grande variance sans annuler complètement leur contribution — particulièrement adapté aux données spectrales." : ""}
  ${method === "normalize" ? "La normalisation min-max compresse toutes les variables dans [0, 1]. Utile quand on souhaite préserver les proportions relatives entre échantillons." : ""}
  ${method === "center" ? "Le centrage soustrait simplement la moyenne — les différences relatives entre variables sont préservées." : ""}
  ${method === "none" ? "Les données brutes sont utilisées sans transformation." : ""}
</div>
`);
  },

  // ─── Section : PCA ────────────────────────────────────────────────────────

  _sectionPCA(APP) {
    const { pcaResult, rawData } = APP;
    const xHeaders = rawData?.headers || [];

    // Top loadings PC1
    const topPC1 = xHeaders
      .map((name, j) => ({ name, abs: Math.abs(pcaResult.loadings[0]?.[j] || 0), val: pcaResult.loadings[0]?.[j] || 0 }))
      .sort((a, b) => b.abs - a.abs)
      .slice(0, 5);

    return this._section("3", "Analyse en Composantes Principales (PCA)", `

${this._statRow([
  ["Composantes", pcaResult.nComp],
  ["Variance PC1", pcaResult.explainedVar[0]?.toFixed(1) + "%"],
  ["Variance PC2", pcaResult.explainedVar[1]?.toFixed(1) + "%"],
  ["Variance totale", pcaResult.cumulativeVar[pcaResult.nComp - 1]?.toFixed(1) + "%"],
  ["T²₉₅ seuil", pcaResult.T2Limit95?.toFixed(2)],
])}

<h3>Tableau de la variance expliquée</h3>
<div class="tscroll">
<table>
  <thead><tr>
    <th>Composante</th><th>Valeur propre</th><th>Variance (%)</th><th>Variance cumulée (%)</th><th>Statut</th>
  </tr></thead>
  <tbody>
    ${pcaResult.explainedVar.map((v, i) => `
    <tr>
      <td class="cyan">PC${i + 1}</td>
      <td>${pcaResult.eigenvalues[i]?.toFixed(4)}</td>
      <td>${v.toFixed(2)}%</td>
      <td>${pcaResult.cumulativeVar[i]?.toFixed(2)}%</td>
      <td>${pcaResult.eigenvalues[i] >= 1 ? '<span class="badge-ok">Kaiser ✓</span>' : '<span class="badge-dim">λ < 1</span>'}</td>
    </tr>`).join("")}
  </tbody>
</table>
</div>

${APP.charts?.screePlot ? this._chartBlock(APP.charts.screePlot, "Scree Plot — Décroissance des valeurs propres. Le coude de la courbe indique le nombre optimal de composantes.") : ""}
${APP.charts?.scorePlot ? this._chartBlock(APP.charts.scorePlot, "Score Plot PC1 vs PC2 — Projection des échantillons dans l'espace des deux premières composantes principales. L'ellipse représente la région de confiance à 95%.") : ""}
${APP.charts?.loadingPlot ? this._chartBlock(APP.charts.loadingPlot, "Loading Plot — Contribution des variables aux composantes principales. Les variables proches du cercle unitaire ont une forte contribution.") : ""}

<h3>Contributions des variables (Loadings PC1)</h3>
<p>Variables les plus influentes sur PC1 :</p>
<div class="tscroll">
<table>
  <thead><tr><th>Variable</th><th>Loading PC1</th>${pcaResult.nComp > 1 ? "<th>Loading PC2</th>" : ""}<th>Importance relative</th></tr></thead>
  <tbody>
    ${topPC1.map(({ name, val }, rank) => `
    <tr>
      <td class="var-name">${this._esc(name)}</td>
      <td class="${val > 0 ? "pos" : "neg"}">${val.toFixed(4)}</td>
      ${pcaResult.nComp > 1 ? `<td>${(pcaResult.loadings[1]?.[xHeaders.indexOf(name)] || 0).toFixed(4)}</td>` : ""}
      <td><div class="mini-bar"><div class="mini-fill" style="width:${(topPC1[rank].abs / topPC1[0].abs * 100).toFixed(0)}%"></div></div></td>
    </tr>`).join("")}
  </tbody>
</table>
</div>
`);
  },

  // ─── Section : PLS ────────────────────────────────────────────────────────

  _sectionPLS(APP) {
    const { plsMetrics, plsModel, rawData } = APP;
    const xHeaders = rawData?.headers || [];
    const vip = plsModel?.vip || [];
    const importantVars = xHeaders
      .map((name, j) => ({ name, vip: vip[j] || 0 }))
      .filter(d => d.vip >= 1)
      .sort((a, b) => b.vip - a.vip)
      .slice(0, 10);

    const r2Quality = plsMetrics.r2_cv >= 0.9 ? "Excellent" :
                      plsMetrics.r2_cv >= 0.8 ? "Bon" :
                      plsMetrics.r2_cv >= 0.6 ? "Acceptable" : "Insuffisant";

    return this._section("4", "Régression PLS-1", `

${this._statRow([
  ["R² calibration",    plsMetrics.r2_cal?.toFixed(4)],
  ["R² validation",     plsMetrics.r2_cv?.toFixed(4)],
  ["RMSEC",             plsMetrics.rmsec?.toFixed(4)],
  ["RMSECV",            plsMetrics.rmsecv?.toFixed(4)],
  ["Variables latentes", plsMetrics.nComp],
  ["Qualité",           r2Quality],
])}

<div class="info-box ${plsMetrics.r2_cv >= 0.8 ? 'info-ok' : plsMetrics.r2_cv >= 0.6 ? 'info-warn' : 'info-err'}">
  <strong>Évaluation du modèle :</strong>
  Le modèle PLS explique <strong>${(plsMetrics.r2_cv * 100).toFixed(1)}%</strong> de la variance de Y en validation croisée.
  L'erreur de prédiction estimée (RMSECV) est de <strong>${plsMetrics.rmsecv?.toFixed(4)}</strong> unités.
  ${plsMetrics.r2_cal - plsMetrics.r2_cv > 0.1 ? "<br>⚠ La différence R²cal - R²cv > 10% suggère un possible sur-ajustement. Envisagez de réduire le nombre de composantes." : ""}
</div>

${APP.charts?.predVsReal ? this._chartBlock(APP.charts.predVsReal, "Graphique Prédit vs Réel — Chaque point représente un échantillon. La ligne diagonale représente la prédiction parfaite.") : ""}
${APP.charts?.vipChart   ? this._chartBlock(APP.charts.vipChart, "VIP Scores — Variables avec VIP ≥ 1 (ligne orange) sont les plus informatives pour la prédiction de Y.") : ""}

<h3>Variables importantes (VIP ≥ 1)</h3>
${importantVars.length > 0 ? `
<div class="tscroll">
<table>
  <thead><tr><th>Variable</th><th>VIP Score</th><th>Importance</th></tr></thead>
  <tbody>
    ${importantVars.map(({ name, vip: v }) => `
    <tr>
      <td class="var-name">${this._esc(name)}</td>
      <td class="${v >= 1.5 ? "orange" : "cyan"}">${v.toFixed(3)}</td>
      <td>${v >= 1.5 ? "🔴 Très importante" : "🟡 Importante"}</td>
    </tr>`).join("")}
  </tbody>
</table>
</div>` : "<p>Aucune variable avec VIP ≥ 1. Le modèle repose sur une combinaison diffuse des variables.</p>"}
`);
  },

  // ─── Section : Clustering ────────────────────────────────────────────────

  _sectionCluster(APP) {
    const { clusterResult } = APP;
    const sil = clusterResult.silhouette_score ?? clusterResult.silhouette;
    const silQuality = sil == null ? "N/A" :
                       sil >= 0.7  ? "Structure forte" :
                       sil >= 0.5  ? "Structure raisonnable" :
                       sil >= 0.25 ? "Structure faible" : "Pas de structure";

    return this._section("5", "Clustering", `

${this._statRow([
  ["Clusters k",  clusterResult.k],
  ["Méthode",     clusterResult.method_used || clusterResult.method],
  ["Silhouette",  sil != null ? sil.toFixed(3) : "N/A"],
  ["Qualité",     silQuality],
  ["Inertie",     clusterResult.inertia != null ? clusterResult.inertia.toFixed(1) : "N/A"],
])}

<div class="info-box">
  <strong>Interprétation du coefficient de Silhouette :</strong>
  La valeur ${sil != null ? sil.toFixed(3) : "N/A"} indique une <strong>${silQuality.toLowerCase()}</strong>.
  Un coefficient > 0.5 indique que les groupes sont bien séparés et cohérents.
  Un coefficient entre 0.25 et 0.5 suggère une structure existante mais des chevauchements entre clusters.
</div>

${APP.charts?.clusterScatter ? this._chartBlock(APP.charts.clusterScatter, `Scatter Plot des clusters dans l'espace PCA — ${clusterResult.k} groupes identifiés, colorés distinctement.`) : ""}
${APP.charts?.dendro ? this._chartBlock(APP.charts.dendro, `Dendrogramme Ward — Structure hiérarchique complète. La ligne rouge indique la coupe pour k = ${clusterResult.k} clusters.`) : ""}
${APP.charts?.heatmap ? this._chartBlock(APP.charts.heatmap, "Heatmap — Vue d'ensemble des intensités. Les variables et échantillons similaires tendent à se regrouper.") : ""}
${APP.charts?.elbowChart ? this._chartBlock(APP.charts.elbowChart, "Méthode Elbow — L'inertie décroît à mesure que k augmente. Le coude de la courbe indique le k optimal.") : ""}
`);
  },

  // ─── Section : Anomalies ─────────────────────────────────────────────────

  _sectionAnomaly(APP) {
    const { anomalyResult, rawData } = APP;
    const nOut = anomalyResult.n_outliers ?? anomalyResult.outliers.filter(Boolean).length;
    const outlierNames = (rawData?.sampleNames || [])
      .filter((_, i) => anomalyResult.outliers[i])
      .slice(0, 15);

    return this._section("6", "Détection d'anomalies", `

${this._statRow([
  ["Méthode",      anomalyResult.method],
  ["Seuil",        anomalyResult.threshold?.toFixed(3)],
  ["Outliers",     `${nOut} / ${anomalyResult.scores.length}`],
  ["Taux",         (nOut / anomalyResult.scores.length * 100).toFixed(1) + "%"],
  ["Score max",    Math.max(...anomalyResult.scores).toFixed(3)],
])}

${nOut > 0 ? `
<div class="info-box info-warn">
  <strong>${nOut} échantillon(s) aberrant(s) détecté(s) :</strong>
  ${outlierNames.length > 0 ? outlierNames.map(n => `<span class="badge-err">${this._esc(n)}</span>`).join(" ") : "Voir les indices dans la table ci-dessous"}
  <br>Ces échantillons présentent un score statistique supérieur au seuil (α = ${APP.anomalyAlpha || "5%"}).
  Ils doivent être examinés manuellement pour déterminer s'il s'agit d'erreurs de mesure ou de variations réelles.
</div>` : `
<div class="info-box info-ok">
  <strong>Aucun outlier détecté</strong> au seuil statistique fixé.
  Tous les échantillons se situent dans la région de confiance définie par la méthode ${anomalyResult.method}.
</div>`}

${APP.charts?.controlChart ? this._chartBlock(APP.charts.controlChart, `Graphique de contrôle — Scores individuels des ${anomalyResult.scores.length} échantillons. La ligne pointillée rouge représente le seuil statistique.`) : ""}

<h3>Classement des échantillons par score</h3>
<div class="tscroll">
<table>
  <thead><tr><th>#</th><th>Échantillon</th><th>Score</th><th>Seuil</th><th>Statut</th></tr></thead>
  <tbody>
    ${anomalyResult.scores
      .map((s, i) => ({ s, i, name: rawData?.sampleNames?.[i] || `S${i + 1}`, isOut: anomalyResult.outliers[i] }))
      .sort((a, b) => b.s - a.s)
      .slice(0, 20)
      .map(({ s, i, name, isOut }) => `
    <tr class="${isOut ? "row-err" : ""}">
      <td>${i + 1}</td>
      <td class="var-name">${this._esc(name)}</td>
      <td class="${isOut ? "red" : "cyan"}">${s.toFixed(4)}</td>
      <td>${anomalyResult.threshold?.toFixed(4)}</td>
      <td>${isOut ? '<span class="badge-err">⚠ OUTLIER</span>' : '<span class="badge-ok">✓ Normal</span>'}</td>
    </tr>`).join("")}
  </tbody>
</table>
</div>
`);
  },

  // ─── Section : IA ────────────────────────────────────────────────────────

  _sectionAI(globalAI) {
    return this._section("7", "Interprétation par Intelligence Artificielle", `
<div class="ai-content">
  ${globalAI.replace(/\n/g, "<br>").replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")}
</div>
<p class="ai-note">Interprétation générée par Gemini (Google AI) sur la base des résultats statistiques. 
Cette interprétation est fournie à titre indicatif et doit être validée par un expert chimiomètre.</p>
`);
  },

  // ─── Pied de page ────────────────────────────────────────────────────────

  _footer(opts, date) {
    return `
<footer class="rpt-footer">
  <div class="footer-main">
    Rapport généré par <strong>ChemAI Add-in v2.0</strong> — ${date}
    ${opts.ref ? ` · Réf. : ${this._esc(opts.ref)}` : ""}
    ${opts.auteur ? ` · Auteur : ${this._esc(opts.auteur)}` : ""}
    ${opts.labo ? ` · ${this._esc(opts.labo)}` : ""}
  </div>
  <div class="footer-refs">
    Wold et al. (2001) · Jolliffe (2002) · Jackson & Mudholkar (1979) · Ward (1963)
  </div>
</footer>`;
  },

  // ─── Composants HTML réutilisables ────────────────────────────────────────

  _section(num, title, content) {
    return `
<section class="rpt-section" id="sec-${num}">
  <h2><span class="sec-num">§${num}</span> ${title}</h2>
  ${content}
</section>`;
  },

  _statRow(items) {
    return `<div class="stat-row">${items.map(([l, v]) => `
  <div class="stat-box">
    <div class="sl">${l}</div>
    <div class="sv">${v != null ? v : "—"}</div>
  </div>`).join("")}</div>`;
  },

  _chartBlock(svgString, caption) {
    if (!svgString) return "";
    return `
<div class="chart-container">
  ${svgString}
  ${caption ? `<p class="caption">${caption}</p>` : ""}
</div>`;
  },

  _esc(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  },

  // ─── Styles inline ────────────────────────────────────────────────────────

  _styles() {
    return `<style>
:root{
  --bg:#080F1A; --bg2:#0D1825; --bg3:#122035; --border:#1A2E47;
  --cyan:#00E5FF; --green:#00E676; --orange:#FF9800; --red:#FF5252;
  --purple:#CE93D8; --text:#D0E4F0; --text-dim:#7B9DB8; --faint:#3A5570;
  --mono:'Courier New',monospace; --sans:'Segoe UI',Arial,sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);font-size:12px;color:var(--text);background:var(--bg);padding:28px 36px;max-width:1100px;margin:auto;line-height:1.65}
h1{font-size:20px;color:var(--cyan);margin-bottom:4px}
h2{font-size:13px;font-weight:700;color:var(--text-dim);margin:24px 0 10px;padding:8px 12px;background:var(--bg2);border-left:4px solid var(--cyan);border-radius:4px;text-transform:uppercase;letter-spacing:.06em}
h3{font-size:11px;font-weight:700;color:var(--text-dim);margin:16px 0 8px;text-transform:uppercase;letter-spacing:.05em}
p{margin-bottom:8px;color:var(--text-dim);font-size:11px}
strong{color:var(--text)}

.rpt-header{display:flex;align-items:center;gap:14px;margin-bottom:16px}
.rpt-logo{font-size:28px;color:var(--cyan);filter:drop-shadow(0 0 8px var(--cyan))}
.rpt-subtitle{font-size:10px;color:var(--faint);margin-top:3px}
.rpt-divider{border:none;border-top:1px solid var(--border);margin:16px 0}

.meta-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:14px}
.meta-lbl{font-size:8px;color:var(--faint);text-transform:uppercase;letter-spacing:.06em}
.meta-val{font-size:12px;font-weight:700;color:var(--cyan);margin-top:2px;font-family:var(--mono)}

.toc{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:20px}
.toc-title{font-size:10px;font-weight:700;color:var(--text-dim);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px}
.toc ul{list-style:none;display:grid;grid-template-columns:1fr 1fr;gap:4px}
.toc li{font-size:11px;color:var(--text-dim);padding:3px 0}
.toc-ref{color:var(--cyan);font-family:var(--mono);font-size:10px;margin-right:6px}

.content{margin-top:20px}
.rpt-section{margin-bottom:32px;padding-bottom:24px;border-bottom:1px solid var(--border)}

.sec-num{color:var(--faint);font-size:11px;margin-right:8px}

.stat-row{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}
.stat-box{background:var(--bg3);border:1px solid var(--border);border-radius:6px;padding:10px 14px;min-width:100px;text-align:center}
.sl{font-size:8px;color:var(--faint);text-transform:uppercase;letter-spacing:.05em}
.sv{font-family:var(--mono);font-size:14px;font-weight:700;color:var(--cyan);margin-top:4px}

table{width:100%;border-collapse:collapse;margin:8px 0;font-size:11px}
th{background:var(--bg);color:var(--cyan);padding:6px 10px;text-align:left;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid var(--border)}
td{padding:5px 10px;border-bottom:1px solid var(--border);font-family:var(--mono);font-size:10px;color:var(--text-dim)}
tr:nth-child(even) td{background:var(--bg2)}
tr.row-err td{background:#FF525210}
.tscroll{overflow-x:auto;border-radius:6px;border:1px solid var(--border)}

.var-name{color:var(--text) !important;font-family:var(--sans) !important}
.cyan{color:var(--cyan) !important}
.pos{color:var(--green) !important}
.neg{color:var(--red) !important}
.orange{color:var(--orange) !important}
.red{color:var(--red) !important}

.chart-container{margin:12px 0;border:1px solid var(--border);border-radius:8px;overflow:hidden}
.caption{font-size:10px;color:var(--faint);padding:6px 12px;background:var(--bg2);font-style:italic;line-height:1.5}

.info-box{background:var(--bg3);border-left:3px solid var(--border);border-radius:6px;padding:12px;margin:10px 0;font-size:11px;line-height:1.7}
.info-ok{border-left-color:var(--green)}
.info-warn{border-left-color:var(--orange)}
.info-err{border-left-color:var(--red)}

.badge-ok{background:#00E67620;color:var(--green);padding:2px 6px;border-radius:3px;font-size:9px;font-weight:700}
.badge-err{background:#FF525220;color:var(--red);padding:2px 6px;border-radius:3px;font-size:9px;font-weight:700;margin:0 2px}
.badge-dim{background:var(--bg);color:var(--faint);padding:2px 6px;border-radius:3px;font-size:9px}

.mini-bar{background:var(--bg);border-radius:3px;height:8px;width:100px;overflow:hidden}
.mini-fill{height:100%;background:linear-gradient(90deg,var(--cyan),var(--blue, #40C4FF));border-radius:3px}

.ai-content{background:var(--bg3);border-left:3px solid var(--cyan);border-radius:6px;padding:16px;font-size:11px;line-height:1.85;color:var(--text)}
.ai-note{font-size:9px;color:var(--faint);margin-top:8px;font-style:italic}

.rpt-footer{margin-top:40px;padding-top:12px;border-top:1px solid var(--border);text-align:center}
.footer-main{font-size:10px;color:var(--text-dim)}
.footer-refs{font-size:9px;color:var(--faint);margin-top:4px}

@media print{
  body{background:#fff;color:#1a1a2e;padding:20px}
  h1{color:#0a2540} h2{background:#f0f4f8;color:#0a2540;border-left-color:#0a2540}
  .stat-box{border:1px solid #ccc} .tscroll{border:1px solid #ccc}
  .chart-container{border:1px solid #ccc}
}
</style>`;
  },

  // ─── Téléchargement ──────────────────────────────────────────────────────

  _download(html, filename) {
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  },
};

window.ChemReport = ChemReport;
