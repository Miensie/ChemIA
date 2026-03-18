/**
 * ================================================================
 * chemMath.js — Moteur de calcul chimiométrique (JS pur)
 * Fonctionne en mode autonome (sans backend Python).
 * Pour les analyses avancées, le backend FastAPI est utilisé.
 *
 * Modules :
 *  - Algèbre linéaire (SVD, multiplication matricielle, etc.)
 *  - Prétraitement (standardisation, normalisation, Pareto)
 *  - PCA (SVD)
 *  - PLS-1 (NIPALS)
 *  - K-Means
 *  - Distance de Mahalanobis
 *  - Statistiques descriptives
 * ================================================================
 */
"use strict";

// ─── Algèbre linéaire de base ────────────────────────────────────────────────

const LA = {
  /** Transposée d'une matrice m×n → n×m */
  T(A) {
    const m = A.length, n = A[0].length;
    const B = Array.from({length: n}, () => new Float64Array(m));
    for (let i = 0; i < m; i++)
      for (let j = 0; j < n; j++)
        B[j][i] = A[i][j];
    return B;
  },

  /** Multiplication A(m×k) × B(k×n) → m×n */
  dot(A, B) {
    const m = A.length, k = B.length, n = B[0].length;
    const C = Array.from({length: m}, () => new Float64Array(n));
    for (let i = 0; i < m; i++)
      for (let l = 0; l < k; l++)
        if (A[i][l] !== 0)
          for (let j = 0; j < n; j++)
            C[i][j] += A[i][l] * B[l][j];
    return C;
  },

  /** Produit matrice × vecteur colonne */
  matvec(A, v) {
    return A.map(row => row.reduce((s, a, j) => s + a * v[j], 0));
  },

  /** Norme euclidienne d'un vecteur */
  norm(v) { return Math.sqrt(v.reduce((s, x) => s + x * x, 0)); },

  /** Normalise un vecteur */
  normalize(v) {
    const n = this.norm(v) || 1;
    return v.map(x => x / n);
  },

  /** Produit scalaire */
  dot1d(a, b) { return a.reduce((s, x, i) => s + x * b[i], 0); },

  /** Soustraction vecteur */
  sub(a, b) { return a.map((x, i) => x - b[i]); },

  /** Copie profonde d'une matrice */
  copy(A) { return A.map(row => Float64Array.from(row)); },

  /** Moyenne colonne par colonne → vecteur de longueur n */
  colMean(X) {
    const m = X.length, n = X[0].length;
    const mu = new Float64Array(n);
    for (let i = 0; i < m; i++)
      for (let j = 0; j < n; j++)
        mu[j] += X[i][j];
    return mu.map(v => v / m);
  },

  /** Écart-type (Bessel) par colonne */
  colStd(X, mu) {
    const m = X.length, n = X[0].length;
    const s = new Float64Array(n);
    for (let i = 0; i < m; i++)
      for (let j = 0; j < n; j++)
        s[j] += (X[i][j] - mu[j]) ** 2;
    return s.map(v => Math.sqrt(v / (m - 1)) || 1);
  },

  /** Retourne une matrice de zéros m×n */
  zeros(m, n) { return Array.from({length: m}, () => new Float64Array(n)); },

  /** Identité n×n */
  eye(n) {
    const I = this.zeros(n, n);
    for (let i = 0; i < n; i++) I[i][i] = 1;
    return I;
  },

  /**
   * SVD partielle par la méthode de déflation + puissance (Lanczos simplifié)
   * Retourne les r premiers vecteurs singuliers : { U, S, Vt }
   * U : m×r, S : r (valeurs singulières), Vt : r×n
   */
  svd(X, r) {
    const m = X.length, n = X[0].length;
    r = Math.min(r || Math.min(m, n), m, n);

    const U = [], S = [], V = [];
    let Xr = this.copy(X);

    for (let k = 0; k < r; k++) {
      // Initialisation aléatoire reproductible
      let v = new Float64Array(n).map(() => Math.random() - 0.5);
      v = this.normalize(v);

      // Itération de la puissance (30 itérations suffisent pour la convergence)
      for (let iter = 0; iter < 60; iter++) {
        const u = this.normalize(this.matvec(Xr, v));
        v = this.normalize(this.matvec(this.T(Xr), u));
      }

      const u = this.normalize(this.matvec(Xr, v));
      const sigma = this.dot1d(u, this.matvec(Xr, v));

      U.push(u);
      S.push(sigma);
      V.push(v);

      // Déflation
      for (let i = 0; i < m; i++)
        for (let j = 0; j < n; j++)
          Xr[i][j] -= sigma * u[i] * v[j];
    }

    return {
      U: U,         // m×r  (tableau de vecteurs)
      S: S,         // r    (valeurs singulières)
      Vt: V,        // r×n  (tableau de vecteurs = lignes de Vt)
    };
  },

  /**
   * Inverse d'une matrice carrée (Gauss-Jordan)
   */
  inv(A) {
    const n = A.length;
    const M = A.map(row => Array.from(row));
    const I = this.eye(n).map(row => Array.from(row));

    for (let col = 0; col < n; col++) {
      let maxRow = col;
      for (let r = col + 1; r < n; r++)
        if (Math.abs(M[r][col]) > Math.abs(M[maxRow][col])) maxRow = r;
      [M[col], M[maxRow]] = [M[maxRow], M[col]];
      [I[col], I[maxRow]] = [I[maxRow], I[col]];
      const piv = M[col][col];
      if (Math.abs(piv) < 1e-14) return null;
      for (let j = 0; j < n; j++) { M[col][j] /= piv; I[col][j] /= piv; }
      for (let r = 0; r < n; r++) {
        if (r === col) continue;
        const f = M[r][col];
        for (let j = 0; j < n; j++) { M[r][j] -= f * M[col][j]; I[r][j] -= f * I[col][j]; }
      }
    }
    return I;
  },
};

// ─── Prétraitement ───────────────────────────────────────────────────────────

const Preprocessing = {
  /**
   * Gestion des valeurs manquantes
   * @param {number[][]} X — matrice m×n (NaN = valeur manquante)
   * @param {string} strategy — "mean" | "median" | "zero" | "drop"
   */
  handleMissing(X, strategy = "mean") {
    const m = X.length, n = X[0].length;

    if (strategy === "drop") {
      return X.filter(row => row.every(v => !isNaN(v) && isFinite(v)));
    }

    const filled = X.map(row => [...row]);

    for (let j = 0; j < n; j++) {
      const valid = X.map(r => r[j]).filter(v => !isNaN(v) && isFinite(v));
      if (!valid.length) continue;

      let fill;
      if (strategy === "mean") {
        fill = valid.reduce((s, v) => s + v, 0) / valid.length;
      } else if (strategy === "median") {
        const sorted = [...valid].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        fill = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
      } else {
        fill = 0;
      }

      for (let i = 0; i < m; i++) {
        if (isNaN(filled[i][j]) || !isFinite(filled[i][j])) filled[i][j] = fill;
      }
    }

    return filled;
  },

  /**
   * Calcule les paramètres de mise à l'échelle (fit)
   */
  fitScaler(X, method) {
    const mu = LA.colMean(X);
    const sd = LA.colStd(X, mu);
    const xMin = new Float64Array(X[0].length);
    const xMax = new Float64Array(X[0].length);

    for (let j = 0; j < X[0].length; j++) {
      xMin[j] = Math.min(...X.map(r => r[j]));
      xMax[j] = Math.max(...X.map(r => r[j]));
    }

    return { mu, sd, xMin, xMax, method };
  },

  /**
   * Applique la mise à l'échelle à X selon les paramètres calculés par fitScaler
   */
  transform(X, params) {
    const { mu, sd, xMin, xMax, method } = params;
    return X.map(row =>
      row.map((v, j) => {
        if (method === "center")      return v - mu[j];
        if (method === "standardize") return (v - mu[j]) / (sd[j] || 1);
        if (method === "normalize")   return (xMax[j] - xMin[j]) > 0 ? (v - xMin[j]) / (xMax[j] - xMin[j]) : 0;
        if (method === "pareto")      return (v - mu[j]) / (Math.sqrt(sd[j]) || 1);
        return v; // none
      })
    );
  },

  /** Statistiques descriptives par colonne */
  descStats(X, colNames) {
    return colNames.map((name, j) => {
      const col = X.map(r => r[j]).filter(v => isFinite(v) && !isNaN(v));
      col.sort((a, b) => a - b);
      const mu = col.reduce((s, v) => s + v, 0) / col.length;
      const sd = Math.sqrt(col.reduce((s, v) => s + (v - mu) ** 2, 0) / (col.length - 1));
      const mid = Math.floor(col.length / 2);
      const median = col.length % 2 ? col[mid] : (col[mid - 1] + col[mid]) / 2;
      return {
        name, n: col.length,
        mean: mu, std: sd, min: col[0], max: col[col.length - 1],
        median, q1: col[Math.floor(col.length * 0.25)], q3: col[Math.floor(col.length * 0.75)],
      };
    });
  },
};

// ─── PCA ────────────────────────────────────────────────────────────────────

const PCA = {
  /**
   * Lance la PCA sur la matrice X (déjà prétraitée)
   * @param {number[][]} X   — m×n
   * @param {number}     nComp — nombre de composantes (ou "auto")
   * @returns PCA result object
   */
  fit(X, nComp = "auto") {
    // ── Garde-fous d'entrée ──
    if (!X || X.length < 2) throw new Error("PCA : au moins 2 échantillons requis.");
    if (!X[0] || X[0].length < 1) throw new Error("PCA : au moins 1 variable requise.");

    // Remplacer les NaN/Inf résiduels par 0
    X = X.map(row => Array.from(row).map(v => (isFinite(v) && !isNaN(v)) ? v : 0));

    const m = X.length, n = X[0].length;
    const maxComp = Math.min(m - 1, n);
    const r = nComp === "auto" ? maxComp : Math.min(nComp, maxComp);

    const { U, S, Vt } = LA.svd(X, r);

    // Valeurs propres = S² / (m-1)
    const eigenvalues = S.map(s => s * s / (m - 1));
    const totalVar = eigenvalues.reduce((s, v) => s + v, 0) ||
                     X.flat().reduce((s, v) => s + v * v, 0) / (m - 1);

    const explainedVar    = eigenvalues.map(e => e / totalVar * 100);
    const cumulativeVar   = explainedVar.reduce((acc, v, i) => {
      acc.push((acc[i - 1] || 0) + v); return acc;
    }, []);

    // Auto : sélectionner le nombre de composantes pour 95% variance
    let nKeep = r;
    if (nComp === "auto") {
      nKeep = cumulativeVar.findIndex(c => c >= 95) + 1;
      if (nKeep <= 0) nKeep = Math.min(r, 5);
    } else {
      nKeep = r;
    }

    // Scores T = X × V
    const V = LA.T(Vt.slice(0, nKeep));   // n × nKeep
    const scores = LA.dot(X, V);          // m × nKeep

    // Loadings = Vt (nKeep × n)
    const loadings = Vt.slice(0, nKeep);

    // Résidus Q = ||X - T P'||²
    const Xrec = LA.dot(scores, Vt.slice(0, nKeep));
    const Q = X.map((row, i) =>
      row.reduce((s, v, j) => s + (v - Xrec[i][j]) ** 2, 0)
    );

    // T² de Hotelling
    const T2 = scores.map(score =>
      score.reduce((s, v, k) => s + v * v / (eigenvalues[k] * (m - 1) || 1), 0)
    );

    return {
      // Convertir Float64Array → tableau JS ordinaire pour compatibilité SVG/JSON
      scores:   scores.map(row => Array.from(row)),     // m × nKeep
      loadings: loadings.map(row => Array.from(row)),   // nKeep × n
      eigenvalues: eigenvalues.slice(0, nKeep),
      explainedVar: explainedVar.slice(0, nKeep),
      cumulativeVar: cumulativeVar.slice(0, nKeep),
      nComp: nKeep,
      Q, T2,
      T2Limit95: nKeep * (m * m - 1) / (m * (m - nKeep)) * 3.0,
    };
  },
};

// ─── PLS-1 NIPALS ───────────────────────────────────────────────────────────

const PLS = {
  /**
   * PLS-1 NIPALS (une seule variable Y)
   * @param {number[][]} X — m×n matrice X (prétraitée)
   * @param {number[]}   y — m vecteur Y
   * @param {number}     nComp
   * @returns PLS model
   */
  fit(X, y, nComp = 3) {
    if (!X || X.length < 3) throw new Error("PLS : au moins 3 échantillons requis.");
    // Nettoyage NaN résiduels
    X = X.map(row => Array.from(row).map(v => (isFinite(v) && !isNaN(v)) ? v : 0));
    y = y.map(v => (isFinite(v) && !isNaN(v)) ? v : 0);

    const m = X.length, n = X[0].length;
    nComp = Math.min(nComp, Math.min(m - 1, n));

    let E = LA.copy(X);
    let f = [...y];

    const W = [], T = [], P = [], Q_pls = [];

    for (let a = 0; a < nComp; a++) {
      // Poids w = X'y / ||X'y||
      let w = new Float64Array(n);
      for (let j = 0; j < n; j++)
        for (let i = 0; i < m; i++)
          w[j] += E[i][j] * f[i];
      w = LA.normalize(w);

      // Scores t = Xw
      const t = E.map(row => LA.dot1d(Array.from(row), Array.from(w)));

      // Loadings p = X't / ||t||²
      const tt = t.reduce((s, v) => s + v * v, 0);
      const p = new Float64Array(n);
      for (let j = 0; j < n; j++)
        for (let i = 0; i < m; i++)
          p[j] += E[i][j] * t[i];
      for (let j = 0; j < n; j++) p[j] /= tt;

      // Coefficient q = y't / ||t||²
      const q = f.reduce((s, v, i) => s + v * t[i], 0) / tt;

      // Déflation
      for (let i = 0; i < m; i++) {
        for (let j = 0; j < n; j++) E[i][j] -= t[i] * p[j];
        f[i] -= t[i] * q;
      }

      W.push(w); T.push(t); P.push(p); Q_pls.push(q);
    }

    // Coefficients β = W*(P'W)^{-1} Q
    // Simplifié pour PLS-1 : β = W R Q (R = W(P'W)^-1)
    const Wmat  = W.map(w => Array.from(w));   // nComp × n
    const Pmat  = P.map(p => Array.from(p));   // nComp × n
    const PtW   = LA.dot(Pmat, LA.T(Wmat));    // nComp × nComp
    const PtWinv = LA.inv(PtW);
    if (!PtWinv) return null;

    const R  = LA.dot(LA.T(Wmat), PtWinv);     // n × nComp
    const beta = new Float64Array(n);
    for (let j = 0; j < n; j++)
      for (let a = 0; a < nComp; a++)
        beta[j] += R[j][a] * Q_pls[a];

    // VIP scores
    const Tmat = LA.T(T); // nComp × m
    const SSY  = Q_pls.map((q, a) => {
      const t = T[a];
      return q * q * t.reduce((s, v) => s + v * v, 0);
    });
    const totalSSY = SSY.reduce((s, v) => s + v, 0);
    const vip = new Float64Array(n);
    for (let j = 0; j < n; j++) {
      let sum = 0;
      for (let a = 0; a < nComp; a++) {
        sum += (W[a][j] ** 2) * SSY[a] * n;
      }
      vip[j] = Math.sqrt(sum / totalSSY);
    }

    return { W, T, P, Q_pls, beta: Array.from(beta), vip: Array.from(vip), nComp };
  },

  /** Prédiction sur X_new avec le modèle */
  predict(model, X_new, yMu, yStd) {
    const { beta } = model;
    return X_new.map(row => {
      const yHat = row.reduce((s, v, j) => s + v * beta[j], 0);
      // Dé-standardiser si Y a été standardisé
      return yMu !== undefined ? yHat * yStd + yMu : yHat;
    });
  },

  /**
   * Validation croisée K-Fold
   * Retourne les RMSECV par nombre de composantes
   */
  crossValidate(X, y, maxComp, k = 5) {
    const m = X.length;
    const indices = Array.from({length: m}, (_, i) => i);
    // Shuffle Fisher-Yates
    for (let i = m - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [indices[i], indices[j]] = [indices[j], indices[i]];
    }

    const foldSize = Math.ceil(m / k);
    const rmsecv = new Float64Array(maxComp).fill(0);

    for (let fold = 0; fold < k; fold++) {
      const testIdx  = indices.slice(fold * foldSize, (fold + 1) * foldSize);
      const trainIdx = indices.filter((_, i) => !testIdx.includes(indices[fold * foldSize + (i % foldSize)]));
      // Simplification : utiliser tous sauf le fold
      const allIdx = indices.filter(i => !testIdx.includes(i));

      const Xtrain = allIdx.map(i => X[i]);
      const ytrain = allIdx.map(i => y[i]);
      const Xtest  = testIdx.map(i => X[i]);
      const ytest  = testIdx.map(i => y[i]);

      for (let nc = 1; nc <= maxComp; nc++) {
        const model = PLS.fit(Xtrain, ytrain, nc);
        if (!model) continue;
        const yPred = PLS.predict(model, Xtest);
        rmsecv[nc - 1] += ytest.reduce((s, v, i) => s + (v - yPred[i]) ** 2, 0);
      }
    }

    return Array.from(rmsecv).map(sse => Math.sqrt(sse / m));
  },
};

// ─── K-Means ─────────────────────────────────────────────────────────────────

const KMeans = {
  /**
   * K-Means (initialisation K-Means++)
   * @param {number[][]} X
   * @param {number}     k
   * @param {number}     maxIter
   */
  fit(X, k, maxIter = 200) {
    const m = X.length, n = X[0].length;

    // K-Means++ initialisation
    const centroids = [X[Math.floor(Math.random() * m)]];
    while (centroids.length < k) {
      const dists = X.map(row => {
        const minD = Math.min(...centroids.map(c =>
          c.reduce((s, v, j) => s + (v - row[j]) ** 2, 0)
        ));
        return minD;
      });
      const total = dists.reduce((s, v) => s + v, 0);
      let r = Math.random() * total;
      let chosen = 0;
      for (let i = 0; i < m; i++) { r -= dists[i]; if (r <= 0) { chosen = i; break; } }
      centroids.push([...X[chosen]]);
    }

    let labels = new Int32Array(m);
    let prevLabels = new Int32Array(m).fill(-1);

    for (let iter = 0; iter < maxIter; iter++) {
      // Assignment
      for (let i = 0; i < m; i++) {
        let best = 0, bestD = Infinity;
        for (let c = 0; c < k; c++) {
          const d = X[i].reduce((s, v, j) => s + (v - centroids[c][j]) ** 2, 0);
          if (d < bestD) { bestD = d; best = c; }
        }
        labels[i] = best;
      }

      if (labels.every((v, i) => v === prevLabels[i])) break;
      prevLabels = Int32Array.from(labels);

      // Update centroids
      const counts = new Int32Array(k);
      const sums   = Array.from({length: k}, () => new Float64Array(n));
      for (let i = 0; i < m; i++) {
        counts[labels[i]]++;
        for (let j = 0; j < n; j++) sums[labels[i]][j] += X[i][j];
      }
      for (let c = 0; c < k; c++) {
        if (counts[c] > 0)
          centroids[c] = Array.from(sums[c]).map(v => v / counts[c]);
      }
    }

    // Inertie (within-cluster sum of squares)
    const inertia = X.reduce((s, row, i) => {
      return s + row.reduce((ss, v, j) => ss + (v - centroids[labels[i]][j]) ** 2, 0);
    }, 0);

    // Silhouette coefficient (approximatif, rapide)
    const silhouette = this._silhouette(X, labels, k);

    return { labels: Array.from(labels), centroids, inertia, silhouette, k };
  },

  /** Méthode elbow : inertie pour k = 2..kMax */
  elbow(X, kMax = 8) {
    const results = [];
    for (let k = 2; k <= Math.min(kMax, X.length - 1); k++) {
      const res = this.fit(X, k);
      results.push({ k, inertia: res.inertia });
    }
    return results;
  },

  _silhouette(X, labels, k) {
    const m = X.length;
    if (m < 4 || k < 2) return 0;
    // Approximation sur 100 points max pour performance
    const sample = m > 100 ? Array.from({length: 100}, (_, i) => Math.floor(i * m / 100)) : Array.from({length: m}, (_, i) => i);

    const scores = sample.map(i => {
      const myCluster = labels[i];
      const sameCluster = sample.filter(j => j !== i && labels[j] === myCluster);
      if (sameCluster.length === 0) return 0;

      const a = sameCluster.reduce((s, j) =>
        s + Math.sqrt(X[i].reduce((ss, v, d) => ss + (v - X[j][d]) ** 2, 0)), 0
      ) / sameCluster.length;

      let b = Infinity;
      for (let c = 0; c < k; c++) {
        if (c === myCluster) continue;
        const otherCluster = sample.filter(j => labels[j] === c);
        if (!otherCluster.length) continue;
        const bd = otherCluster.reduce((s, j) =>
          s + Math.sqrt(X[i].reduce((ss, v, d) => ss + (v - X[j][d]) ** 2, 0)), 0
        ) / otherCluster.length;
        if (bd < b) b = bd;
      }

      const denom = Math.max(a, b);
      return denom > 0 ? (b - a) / denom : 0;
    });

    return scores.reduce((s, v) => s + v, 0) / scores.length;
  },
};

// ─── Clustering hiérarchique (Ward) ──────────────────────────────────────────

const Hierarchical = {
  /**
   * Clustering hiérarchique Ward — retourne la linkage matrix (n-1 × 4)
   * Format: [clusterA, clusterB, distance, size]
   */
  fit(X, linkage = "ward") {
    const m = X.length;
    // Initialisation : chaque point est un cluster
    let clusters = X.map((row, i) => ({ id: i, points: [i], centroid: [...row] }));
    const linkageMatrix = [];
    let nextId = m;

    while (clusters.length > 1) {
      let minDist = Infinity, best = [0, 1];

      for (let i = 0; i < clusters.length - 1; i++) {
        for (let j = i + 1; j < clusters.length; j++) {
          const d = this._dist(clusters[i], clusters[j], X, linkage);
          if (d < minDist) { minDist = d; best = [i, j]; }
        }
      }

      const [i, j] = best;
      const ci = clusters[i], cj = clusters[j];
      const merged = {
        id: nextId++,
        points: [...ci.points, ...cj.points],
        centroid: ci.points.concat(cj.points).reduce((acc, idx) => {
          return acc.map((v, d) => v + X[idx][d] / (ci.points.length + cj.points.length));
        }, new Array(X[0].length).fill(0)),
      };

      linkageMatrix.push([ci.id, cj.id, minDist, merged.points.length]);
      clusters = clusters.filter((_, idx) => idx !== i && idx !== j);
      clusters.push(merged);
    }

    return { linkageMatrix, nSamples: m };
  },

  /** Coupe le dendrogramme pour obtenir k clusters */
  cutTree(linkageMatrix, k, nSamples) {
    const n = nSamples;
    // Initialise : chaque nœud est son propre cluster
    const parent = Array.from({length: 2 * n - 1}, (_, i) => i);

    // Union-Find simplifié
    const find = id => {
      while (parent[id] !== id) id = parent[id];
      return id;
    };

    const nMerges = linkageMatrix.length - (k - 1);
    for (let m = 0; m < nMerges; m++) {
      const [a, b, , ] = linkageMatrix[m];
      const newId = n + m;
      parent[find(a)] = newId;
      parent[find(b)] = newId;
      parent[newId] = newId;
    }

    // Assigner les labels
    const clusterIds = [...new Set(Array.from({length: n}, (_, i) => find(i)))];
    return Array.from({length: n}, (_, i) => clusterIds.indexOf(find(i)));
  },

  _dist(ci, cj, X, method) {
    if (method === "ward") {
      const ni = ci.points.length, nj = cj.points.length;
      return Math.sqrt(
        ci.centroid.reduce((s, v, d) => s + (v - cj.centroid[d]) ** 2, 0) *
        ni * nj / (ni + nj)
      );
    }
    if (method === "complete") {
      let max = -Infinity;
      for (const i of ci.points)
        for (const j of cj.points) {
          const d = X[i].reduce((s, v, d) => s + (v - X[j][d]) ** 2, 0);
          if (d > max) max = d;
        }
      return Math.sqrt(max);
    }
    if (method === "single") {
      let min = Infinity;
      for (const i of ci.points)
        for (const j of cj.points) {
          const d = X[i].reduce((s, v, d) => s + (v - X[j][d]) ** 2, 0);
          if (d < min) min = d;
        }
      return Math.sqrt(min);
    }
    // average
    let sum = 0;
    for (const i of ci.points)
      for (const j of cj.points)
        sum += Math.sqrt(X[i].reduce((s, v, d) => s + (v - X[j][d]) ** 2, 0));
    return sum / (ci.points.length * cj.points.length);
  },
};

// ─── Détection d'anomalies ────────────────────────────────────────────────────

const Anomaly = {
  /**
   * Distance de Mahalanobis pour chaque observation
   * @param {number[][]} X — matrice m×n
   * @returns {number[]} distances
   */
  mahalanobis(X) {
    const m = X.length, n = X[0].length;
    const mu = LA.colMean(X);
    const Xc = X.map(row => row.map((v, j) => v - mu[j]));

    // Matrice de covariance S = X'X / (m-1)
    const S = LA.zeros(n, n);
    for (let i = 0; i < m; i++)
      for (let a = 0; a < n; a++)
        for (let b = 0; b < n; b++)
          S[a][b] += Xc[i][a] * Xc[i][b];
    for (let a = 0; a < n; a++)
      for (let b = 0; b < n; b++)
        S[a][b] /= (m - 1);

    const Sinv = LA.inv(S);
    if (!Sinv) {
      // Cas singulier : retourner la distance euclidienne standardisée
      const sd = LA.colStd(X, mu);
      return Xc.map(row => Math.sqrt(row.reduce((s, v, j) => s + (v / (sd[j] || 1)) ** 2, 0)));
    }

    return Xc.map(row => {
      const Sinv_row = LA.matvec(Sinv, row);
      return Math.sqrt(Math.max(0, LA.dot1d(row, Sinv_row)));
    });
  },

  /**
   * Seuil Chi² (approximation) pour la distance de Mahalanobis
   * @param {number} n   — nombre de variables
   * @param {number} alpha — risque (0.05, 0.01, 0.001)
   */
  mahalanobisThreshold(n, alpha = 0.05) {
    // Approximation du quantile Chi²(n) par la méthode Wilson-Hilferty
    const z = alpha === 0.05 ? 1.645 : alpha === 0.01 ? 2.326 : 3.090;
    const mu_chi = n;
    const sd_chi = Math.sqrt(2 * n);
    return Math.sqrt(Math.max(0, mu_chi + z * sd_chi));
  },

  /**
   * Détection combinée T² + Q résiduel (depuis une PCA)
   */
  fromPCA(pcaResult, alpha = 0.05) {
    const { T2, T2Limit95, Q } = pcaResult;
    const qLimit = Q.reduce((s, v) => s + v, 0) / Q.length +
                   (alpha === 0.05 ? 2 : 3) * Math.sqrt(Q.reduce((s, v) => s + v * v, 0) / Q.length);
    return {
      T2, qScores: Q,
      T2Limit: T2Limit95,
      qLimit,
      outliers: T2.map((t, i) => t > T2Limit95 || Q[i] > qLimit),
    };
  },
};

// ─── Export ───────────────────────────────────────────────────────────────────

window.ChemMath = { LA, Preprocessing, PCA, PLS, KMeans, Hierarchical, Anomaly };