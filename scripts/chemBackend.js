/**
 * ================================================================
 * chemBackend.js — Client HTTP vers le backend Python FastAPI
 *
 * L'Add-in peut fonctionner en deux modes :
 *  - Mode autonome (JS pur)   : ChemMath (SVD, PLS, K-Means)
 *  - Mode backend Python      : calculs scikit-learn plus précis
 *
 * Ce fichier gère le mode backend via fetch().
 * ================================================================
 */
"use strict";

const ChemBackend = {
  _baseUrl: "http://localhost:8000",
  _enabled: false,
  _timeout: 30000,  // 30 secondes

  configure(baseUrl) {
    this._baseUrl = baseUrl.replace(/\/$/, "");
  },

  async checkHealth() {
    try {
      const r = await this._fetch("/health", "GET");
      this._enabled = r.status === "ok";
      return r;
    } catch {
      this._enabled = false;
      return null;
    }
  },

  isEnabled() { return this._enabled; },

  // ─── Appels API ─────────────────────────────────────────────────────────────

  async preprocess(X, varNames, sampleNames, method, missingStrategy) {
    return this._fetch("/preprocess/", "POST", {
      X, var_names: varNames, sample_names: sampleNames,
      method, missing_strategy: missingStrategy, nan_threshold: 30,
    });
  },

  async pca(X, varNames, sampleNames, nComponents, scaleData) {
    return this._fetch("/pca/", "POST", {
      X, var_names: varNames, sample_names: sampleNames,
      n_components: nComponents || null,
      algorithm: "svd",
      scale_data: scaleData !== false,
    });
  },

  async pls(X, y, varNames, sampleNames, nComponents, cvFolds, testSize) {
    return this._fetch("/pls/", "POST", {
      X, y, var_names: varNames, sample_names: sampleNames,
      n_components: nComponents || null,
      cv_folds: cvFolds || 5,
      test_size: testSize || 0.2,
    });
  },

  async clustering(X, varNames, method, k, linkage) {
    return this._fetch("/clustering/", "POST", {
      X, var_names: varNames,
      method: method || "kmeans",
      k: k || null,
      linkage: linkage || "ward",
      max_k: 10,
    });
  },

  async anomaly(X, varNames, method, alpha, pcaScores, pcaEigenvalues, pcaQ) {
    return this._fetch("/anomaly/", "POST", {
      X, var_names: varNames,
      method: method || "mahalanobis",
      alpha: alpha || 0.05,
      pca_scores: pcaScores || null,
      pca_eigenvalues: pcaEigenvalues || null,
      pca_Q: pcaQ || null,
    });
  },

  async classify(X, yLabels, varNames, method, cvFolds, testSize) {
    return this._fetch("/classify/", "POST", {
      X, y_labels: yLabels, var_names: varNames,
      method: method || "lda",
      cv_folds: cvFolds || 5,
      test_size: testSize || 0.2,
    });
  },

  async optimizePLS(X, y, varNames, maxComponents, criterion) {
    return this._fetch("/optimize/pls", "POST", {
      X, y, var_names: varNames,
      max_components: maxComponents || 15,
      criterion: criterion || "rmsecv",
    });
  },

  // ─── Transport ──────────────────────────────────────────────────────────────

  async _fetch(path, method, body) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this._timeout);

    try {
      const resp = await fetch(`${this._baseUrl}${path}`, {
        method,
        headers: body ? { "Content-Type": "application/json" } : {},
        body: body ? JSON.stringify(body) : undefined,
        signal: controller.signal,
      });

      clearTimeout(timeout);

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ detail: resp.statusText }));
        throw new Error(err.detail || `Erreur HTTP ${resp.status}`);
      }

      return await resp.json();
    } catch (e) {
      clearTimeout(timeout);
      if (e.name === "AbortError") throw new Error("Timeout : le backend Python ne répond pas");
      throw e;
    }
  },
};

// ── Wrapper unifié : Backend Python si disponible, sinon JS ────────────────

const ChemEngine = {
  /**
   * Lance la PCA via le backend Python si disponible,
   * sinon utilise ChemMath.PCA (JS pur).
   */
  async pca(X, varNames, sampleNames, nComponents) {
    if (ChemBackend.isEnabled()) {
      return ChemBackend.pca(X, varNames, sampleNames, nComponents, true);
    }
    const result = ChemMath.PCA.fit(X, nComponents || "auto");
    return {
      scores:          result.scores.map(r => Array.from(r)),
      loadings:        result.loadings,
      eigenvalues:     result.eigenvalues,
      explained_var:   result.explainedVar,
      cumulative_var:  result.cumulativeVar,
      n_components:    result.nComp,
      T2:              result.T2,
      Q:               result.Q,
      T2_limit_95:     result.T2Limit95,
      Q_limit_95:      Math.max(...result.Q) * 1.2,  // approximation
    };
  },

  async pls(X, y, varNames, sampleNames, nComponents, cvFolds, testSize) {
    if (ChemBackend.isEnabled()) {
      return ChemBackend.pls(X, y, varNames, sampleNames, nComponents, cvFolds, testSize);
    }
    const yMu  = y.reduce((s, v) => s + v, 0) / y.length;
    const yStd = Math.sqrt(y.reduce((s, v) => s + (v - yMu) ** 2, 0) / (y.length - 1));
    const ys   = y.map(v => (v - yMu) / (yStd || 1));
    const nc   = nComponents || 3;
    const model = ChemMath.PLS.fit(X, ys, nc);
    if (!model) throw new Error("Échec PLS (matrice singulière ?)");
    const yPred = ChemMath.PLS.predict(model, X, yMu, yStd);
    const splitIdx = Math.floor(X.length * 0.8);
    const r2 = (yR, yH) => {
      const mu = yR.reduce((s, v) => s + v, 0) / yR.length;
      const sst = yR.reduce((s, v) => s + (v - mu) ** 2, 0);
      const sse = yR.reduce((s, v, i) => s + (v - yH[i]) ** 2, 0);
      return 1 - sse / (sst || 1);
    };
    const rmse = (yR, yH) => Math.sqrt(yR.reduce((s, v, i) => s + (v - yH[i]) ** 2, 0) / yR.length);
    return {
      y_pred_train:    yPred.slice(0, splitIdx),
      y_pred_test:     yPred.slice(splitIdx),
      y_test:          y.slice(splitIdx),
      r2_cal:          r2(y.slice(0, splitIdx), yPred.slice(0, splitIdx)),
      r2_cv:           r2(y.slice(splitIdx), yPred.slice(splitIdx)),
      rmsec:           rmse(y.slice(0, splitIdx), yPred.slice(0, splitIdx)),
      rmsecv:          rmse(y.slice(splitIdx), yPred.slice(splitIdx)),
      n_components:    nc,
      vip_scores:      model.vip,
      coefficients:    model.beta,
      rmsecv_per_comp: [],
      x_scores:        model.T.map(t => t),
      x_loadings:      model.P.map(p => Array.from(p)),
    };
  },

  async clustering(X, varNames, method, k, linkage) {
    if (ChemBackend.isEnabled()) {
      return ChemBackend.clustering(X, varNames, method, k, linkage);
    }
    // JS pur
    const kVal = k || 3;
    const result = ChemMath.KMeans.fit(X, kVal);
    const elbowData = ChemMath.KMeans.elbow(X, 8);
    let linkageMatrix = null;
    if (method === "hierarchical" || method === "both") {
      const hRes = ChemMath.Hierarchical.fit(X, linkage || "ward");
      linkageMatrix = hRes.linkageMatrix;
    }
    return {
      labels: result.labels,
      k: result.k,
      inertia: result.inertia,
      silhouette_score: result.silhouette,
      calinski_score: null,
      centroids: result.centroids,
      linkage_matrix: linkageMatrix,
      elbow_data: elbowData,
      method_used: method,
    };
  },

  async anomaly(X, varNames, method, alpha, pcaResult) {
    if (ChemBackend.isEnabled()) {
      return ChemBackend.anomaly(
        X, varNames, method, alpha,
        pcaResult?.scores?.map(r => Array.from(r)),
        pcaResult?.eigenvalues,
        pcaResult?.Q,
      );
    }
    // JS pur
    let scores, threshold, outliers;
    if (method === "mahalanobis" || !pcaResult) {
      scores    = ChemMath.Anomaly.mahalanobis(X);
      threshold = ChemMath.Anomaly.mahalanobisThreshold(X[0].length, alpha);
      outliers  = scores.map(s => s > threshold);
    } else {
      const res = ChemMath.Anomaly.fromPCA(pcaResult, alpha);
      scores    = method === "q_residual" ? res.qScores : res.T2;
      threshold = method === "q_residual" ? res.qLimit  : res.T2Limit;
      outliers  = res.outliers;
    }
    return {
      scores, threshold, outliers,
      n_outliers: outliers.filter(Boolean).length,
      method,
    };
  },
};

window.ChemBackend = ChemBackend;
window.ChemEngine  = ChemEngine;
