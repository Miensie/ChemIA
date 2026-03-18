/**
 * ================================================================
 * chemAI.js — Interprétation IA des résultats chimiométriques
 * Intégration Gemini API pour générer des interprétations scientifiques
 * ================================================================
 */
"use strict";

const ChemAI = {
  _apiKey: null,
  _model: "gemini-2.5-flash-lite",
  _endpoint: "https://generativelanguage.googleapis.com/v1beta/models",

  setApiKey(key) {
    this._apiKey = key;
    try { sessionStorage.setItem("chemAI_key", key); } catch {}
  },

  loadApiKey() {
    try {
      const saved = sessionStorage.getItem("chemAI_key");
      if (saved) {
        this._apiKey = saved;
        const el = document.getElementById("ai-api-key");
        if (el) el.value = "●".repeat(20);
      }
    } catch {}
  },

  hasKey() { return !!this._apiKey; },

  /**
   * Appel générique à l'API Gemini
   */
  async _call(prompt, systemPrompt) {
    if (!this._apiKey) throw new Error("Clé API Gemini non configurée.");

    const url = `${this._endpoint}/${this._model}:generateContent?key=${this._apiKey}`;
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      systemInstruction: {
        parts: [{ text: systemPrompt || this._defaultSystem() }],
      },
      generationConfig: {
        temperature:    0.3,
        maxOutputTokens: 800,
        candidateCount:  1,
      },
    };

    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err?.error?.message || `Erreur API Gemini (${resp.status})`);
    }

    const data = await resp.json();
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || "(Pas de réponse)";
  },

  _defaultSystem() {
    return `Tu es un expert en chimiométrie, chimie analytique et statistiques multivariées.
Tu interprètes des résultats d'analyses statistiques (PCA, PLS, clustering, détection d'outliers)
dans un contexte de laboratoire analytique professionnel.
Réponds en français, de manière structurée, précise et scientifique.
Utilise des termes techniques appropriés. Limite ta réponse à 400 mots maximum.
Structure tes réponses avec des sections claires (sans Markdown lourd).`;
  },

  // ─── Interprétations spécifiques ────────────────────────────────────────────

  /**
   * Interprétation PCA
   */
  async interpretPCA(pcaResult, varNames, sampleNames) {
    const top3Vars_PC1 = this._topLoadings(pcaResult.loadings[0], varNames, 3);
    const top3Vars_PC2 = pcaResult.loadings.length > 1
      ? this._topLoadings(pcaResult.loadings[1], varNames, 3)
      : [];

    const prompt = `
Analyse PCA sur ${sampleNames.length} échantillons et ${varNames.length} variables.

Résultats :
- PC1 : ${pcaResult.explainedVar[0]?.toFixed(1)}% de variance expliquée
  Variables dominantes : ${top3Vars_PC1.map(v => `${v.name} (loading: ${v.val})`).join(', ')}
- PC2 : ${pcaResult.explainedVar[1]?.toFixed(1) ?? 'N/A'}% de variance expliquée
  Variables dominantes : ${top3Vars_PC2.map(v => `${v.name} (loading: ${v.val})`).join(', ')}
- Variance totale expliquée par ${pcaResult.nComp} composantes : ${pcaResult.cumulativeVar[pcaResult.nComp-1]?.toFixed(1)}%
- Valeurs propres : ${pcaResult.eigenvalues.map(e => e.toFixed(3)).join(', ')}
- Noms des variables : ${varNames.join(', ')}

Fournis une interprétation chimiométrique complète :
1. Signification des composantes principales
2. Variables les plus discriminantes
3. Structure des données
4. Recommandations analytiques
    `.trim();

    return this._call(prompt);
  },

  /**
   * Interprétation PLS
   */
  async interpretPLS(metrics, vip, varNames, nComp) {
    const importantVars = vip
      .map((v, i) => ({ v, name: varNames[i] }))
      .filter(d => d.v >= 1.0)
      .sort((a, b) => b.v - a.v)
      .slice(0, 5)
      .map(d => `${d.name} (VIP=${d.v.toFixed(2)})`);

    const prompt = `
Résultats modèle PLS-1 :
- Composantes LV utilisées : ${nComp}
- R² calibration : ${metrics.r2_cal?.toFixed(4) || 'N/A'}
- R² validation croisée : ${metrics.r2_cv?.toFixed(4) || 'N/A'}
- RMSEC (calibration) : ${metrics.rmsec?.toFixed(4) || 'N/A'}
- RMSECV (validation croisée) : ${metrics.rmsecv?.toFixed(4) || 'N/A'}
- Variables avec VIP ≥ 1 : ${importantVars.join(', ') || 'Aucune'}
- Nombre total de variables : ${varNames.length}

Fournis une interprétation du modèle PLS :
1. Qualité du modèle (biais, variance, sur-ajustement ?)
2. Variables spectrales/analytiques les plus informatives
3. Nombre optimal de composantes
4. Recommandations pour améliorer la prédiction
    `.trim();

    return this._call(prompt);
  },

  /**
   * Interprétation clustering
   */
  async interpretClusters(clusterResult, varNames, method) {
    const { k, inertia, silhouette, centroids } = clusterResult;

    const centroidDesc = centroids.slice(0, 4).map((c, ci) => {
      const top2 = varNames
        .map((name, j) => ({ name, v: c[j] }))
        .sort((a, b) => Math.abs(b.v) - Math.abs(a.v))
        .slice(0, 3)
        .map(d => `${d.name}=${d.v.toFixed(2)}`);
      return `Cluster ${ci+1}: [${top2.join(', ')}]`;
    }).join('\n');

    const prompt = `
Clustering ${method || 'K-Means'} sur données chimiométriques :
- Nombre de clusters : ${k}
- Coefficient Silhouette : ${silhouette?.toFixed(3) || 'N/A'} (idéal > 0.5)
- Inertie intra-clusters : ${inertia?.toFixed(1) || 'N/A'}
- Variables : ${varNames.join(', ')}

Profils des centroïdes :
${centroidDesc}

Fournis une interprétation analytique des clusters :
1. Qualité et cohérence de la partition
2. Caractérisation chimique de chaque cluster
3. Variables discriminantes entre groupes
4. Hypothèses sur la nature des groupes
5. Recommandations analytiques
    `.trim();

    return this._call(prompt);
  },

  /**
   * Interprétation anomalies
   */
  async interpretAnomalies(outlierInfo, sampleNames, method) {
    const { outliers, scores, threshold } = outlierInfo;
    const outlierSamples = outliers
      .map((isOut, i) => isOut ? `${sampleNames[i] || 'S'+(i+1)} (score=${scores[i]?.toFixed(2)})` : null)
      .filter(Boolean);

    const prompt = `
Détection d'anomalies (${method}) sur données chimiométriques :
- Méthode : ${method}
- Seuil statistique : ${threshold?.toFixed(3) || 'N/A'}
- Nombre d'outliers détectés : ${outlierSamples.length} / ${outliers.length}
- Échantillons aberrants : ${outlierSamples.slice(0,8).join(', ') || 'Aucun'}
- Score moyen : ${scores ? (scores.reduce((s,v)=>s+v,0)/scores.length).toFixed(3) : 'N/A'}
- Score maximum : ${scores ? Math.max(...scores).toFixed(3) : 'N/A'}

Fournis une interprétation scientifique :
1. Signification statistique des outliers
2. Causes probables (contamination, erreur de mesure, variabilité réelle ?)
3. Impact sur les modèles chimiométriques
4. Recommandations sur le traitement de ces échantillons
    `.trim();

    return this._call(prompt);
  },

  /**
   * Rapport IA global
   */
  async globalInterpretation(context) {
    const prompt = `
Synthèse globale d'une analyse chimiométrique complète :

Données : ${context.nSamples} échantillons, ${context.nVars} variables.
Variables : ${context.varNames?.slice(0,10).join(', ')}${context.nVars > 10 ? '...' : ''}

${context.pca ? `
PCA :
- ${context.pca.nComp} composantes retenues, ${context.pca.totalVar?.toFixed(1)}% de variance expliquée
- Variables PC1 dominantes : ${context.pca.pc1Vars}
` : ''}

${context.pls ? `
PLS :
- R² validation : ${context.pls.r2cv?.toFixed(4)}
- RMSECV : ${context.pls.rmsecv?.toFixed(4)}
- Variables clés (VIP>1) : ${context.pls.vipVars}
` : ''}

${context.clustering ? `
Clustering :
- ${context.clustering.k} groupes identifiés (Silhouette = ${context.clustering.silhouette?.toFixed(3)})
` : ''}

${context.anomalies ? `
Anomalies :
- ${context.anomalies.nOutliers} échantillons aberrants détectés
` : ''}

Rédige une conclusion scientifique générale pour un rapport de laboratoire, incluant :
1. Vue d'ensemble de la qualité des données
2. Structure chimique/analytique révélée
3. Modèles de prédiction et leur applicabilité
4. Points d'attention et limites
5. Recommandations pour la suite
    `.trim();

    return this._call(prompt, `Tu es un expert en chimiométrie rédigeant la conclusion d'un rapport analytique scientifique professionnel. 
Réponds en français, de manière rigoureuse et concise. Style : rapport scientifique.`);
  },

  // ─── Utilitaires ─────────────────────────────────────────────────────────────

  _topLoadings(loadingVec, varNames, n) {
    return varNames
      .map((name, i) => ({ name, val: loadingVec[i]?.toFixed(3) || '0', abs: Math.abs(loadingVec[i] || 0) }))
      .sort((a, b) => b.abs - a.abs)
      .slice(0, n);
  },

  /** Formate le texte IA pour affichage HTML */
  formatResponse(text) {
    return text
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em>$1</em>')
      .replace(/^(\d+\..+)$/gm, '<div style="margin:4px 0;"><strong>$1</strong></div>')
      .replace(/\n\n/g, '<br><br>')
      .replace(/\n/g, '<br>');
  },
};

window.ChemAI = ChemAI;
