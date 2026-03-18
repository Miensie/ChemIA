/**
 * ================================================================
 * chemExcel.js — Interactions Excel via Office.js
 * Compatible Excel Desktop + Excel Online
 *
 * Règles strictes :
 *  1. Un seul Excel.run par opération
 *  2. getItemOrNullObject() uniquement (jamais getItem() dans try/catch)
 *  3. _write() valide et normalise TOUJOURS les dimensions avant écriture
 *  4. Jamais de tableau vide passé à .values
 * ================================================================
 */
"use strict";

const ChemExcel = {

  // ══════════════════════════════════════════════════════════════
  //  HELPERS INTERNES
  // ══════════════════════════════════════════════════════════════

  /**
   * Obtient ou crée une feuille dans un Excel.run actif.
   * La vide si elle existe déjà.
   */
  async _sheet(ctx, name) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    ws.load("isNullObject");
    await ctx.sync();

    let sheet;
    if (ws.isNullObject) {
      sheet = ctx.workbook.worksheets.add(name);
      await ctx.sync();
    } else {
      sheet = ws;
      const used = ws.getUsedRangeOrNullObject();
      used.load("isNullObject");
      await ctx.sync();
      if (!used.isNullObject) {
        used.clear("All");
        await ctx.sync();
      }
    }
    sheet.activate();
    return sheet;
  },

  /**
   * Nettoyage d'une valeur pour Excel :
   * null/undefined → "" | NaN/Infinity → 0 | string/number → inchangé
   */
  _clean(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && !isFinite(v)) return 0;
    return v;
  },

  /** Arrondit proprement (safe). */
  _n(v, d = 4) {
    if (v === null || v === undefined) return 0;
    const n = Number(v);
    if (!isFinite(n)) return 0;
    return parseFloat(n.toFixed(d));
  },

  /**
   * Écrit un tableau 2D dans Excel avec validation stricte des dimensions.
   * Garantit que TOUTES les lignes ont exactement la même longueur.
   * Ne fait RIEN si rows2D est vide.
   * @returns {number} prochaine ligne disponible
   */
  _write(sheet, rowStart, colStart, rows2D) {
    if (!rows2D || rows2D.length === 0) return rowStart;

    // 1. Convertir chaque ligne en tableau JS ordinaire
    const arrays = rows2D.map(r => Array.from(r));

    // 2. Calculer la largeur maximale réelle
    const nCols = Math.max(...arrays.map(r => r.length), 1);

    // 3. Normaliser : toutes les lignes à exactement nCols colonnes
    const safe = arrays.map(r => {
      const out = new Array(nCols).fill("");
      for (let j = 0; j < nCols; j++) {
        out[j] = j < r.length ? this._clean(r[j]) : "";
      }
      return out;
    });

    const nRows = safe.length;

    // 4. Écrire en une seule opération
    sheet.getRangeByIndexes(rowStart, colStart, nRows, nCols).values = safe;
    return rowStart + nRows;
  },

  /** Écrit une ligne de titre (fond bleu clair). */
  _title(sheet, row, text, nCols) {
    const w = Math.max(nCols, 1);
    const vals = [text, ...new Array(w - 1).fill("")];
    sheet.getRangeByIndexes(row, 0, 1, w).values = [vals];
    const cell = sheet.getRangeByIndexes(row, 0, 1, 1).getCell(0, 0);
    cell.format.font.bold  = true;
    cell.format.font.size  = 11;
    cell.format.font.color = "#FFFFFF";
    sheet.getRangeByIndexes(row, 0, 1, w).format.fill.color = "#336699";
    return row + 1;
  },

  /** Écrit une ligne d'en-tête (fond vert clair). */
  _header(sheet, row, cols) {
    if (!cols || cols.length === 0) return row;
    const vals = cols.map(c => String(c ?? ""));
    sheet.getRangeByIndexes(row, 0, 1, vals.length).values = [vals];
    sheet.getRangeByIndexes(row, 0, 1, vals.length).format.font.bold  = true;
    sheet.getRangeByIndexes(row, 0, 1, vals.length).format.font.color = "#003300";
    sheet.getRangeByIndexes(row, 0, 1, vals.length).format.fill.color = "#C6EFCE";
    return row + 1;
  },

  // ══════════════════════════════════════════════════════════════
  //  LECTURE
  // ══════════════════════════════════════════════════════════════

  async readSelection(hasHeader = true) {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address"]);
      await ctx.sync();

      const raw = range.values;
      if (!raw || raw.length < 2)
        throw new Error("Sélection trop petite (minimum 2 lignes).");

      let headers, dataRows, sampleNames;

      if (hasHeader) {
        headers  = raw[0].map((h, i) => (h !== null && h !== "") ? String(h) : `V${i+1}`);
        dataRows = raw.slice(1);
      } else {
        headers  = raw[0].map((_, i) => `V${i+1}`);
        dataRows = raw;
      }

      const firstColIsText = dataRows.every(
        r => typeof r[0] === "string" && isNaN(parseFloat(r[0]))
      );
      if (firstColIsText) {
        sampleNames = dataRows.map(r => String(r[0]));
        dataRows    = dataRows.map(r => r.slice(1));
        headers     = hasHeader
          ? raw[0].slice(1).map((h, i) => (h !== null && h !== "") ? String(h) : `V${i+1}`)
          : headers.slice(1);
      } else {
        sampleNames = dataRows.map((_, i) => `S${i+1}`);
      }

      const data = dataRows.map(r =>
        r.map(v => { const n = parseFloat(v); return isNaN(n) ? NaN : n; })
      );

      return {
        headers, data, sampleNames,
        address: range.address,
        nRows: data.length,
        nCols: headers.length,
      };
    });
  },

  parseCSV(text, separator = ",", hasHeader = true) {
    const lines = text.trim().split(/\r?\n/);
    const rows  = lines.map(l => l.split(separator).map(v => v.trim().replace(/^"|"$/g, "")));

    let headers, dataRows, sampleNames;
    if (hasHeader) {
      headers  = rows[0].map((h, i) => h || `V${i+1}`);
      dataRows = rows.slice(1);
    } else {
      headers  = rows[0].map((_, i) => `V${i+1}`);
      dataRows = rows;
    }

    const firstColIsText = dataRows.every(
      r => typeof r[0] === "string" && isNaN(parseFloat(r[0]))
    );
    if (firstColIsText) {
      sampleNames = dataRows.map(r => r[0]);
      headers     = headers.slice(1);
      dataRows    = dataRows.map(r => r.slice(1));
    } else {
      sampleNames = dataRows.map((_, i) => `S${i+1}`);
    }

    const data = dataRows.map(r =>
      r.map(v => { const n = parseFloat(v); return isNaN(n) ? NaN : n; })
    );

    return { headers, data, sampleNames, nRows: data.length, nCols: headers.length };
  },

  // ══════════════════════════════════════════════════════════════
  //  EXPORT PCA
  // ══════════════════════════════════════════════════════════════

  async exportPCAResults(pcaResult, sampleNames, varNames) {
    return Excel.run(async (ctx) => {
      const sheet  = await this._sheet(ctx, "ChemAI_PCA");
      const nComp  = Math.max(pcaResult.nComp || 1, 1);
      const pcHdrs = Array.from({length: nComp}, (_, i) => `PC${i+1}`);
      let row = 0;

      // ── Scores ───────────────────────────────────────────────
      const scoreWidth = 1 + nComp;
      row = this._title(sheet, row, "SCORES PCA", scoreWidth);
      row = this._header(sheet, row, ["Échantillon", ...pcHdrs]);

      const scores = pcaResult.scores || [];
      if (scores.length > 0) {
        const scoreData = scores.map((s, i) => {
          // Convertir Float64Array → Array et s'assurer de nComp valeurs exactement
          const vals = Array.from(s).slice(0, nComp).map(v => this._n(v, 6));
          while (vals.length < nComp) vals.push(0);
          return [String(sampleNames?.[i] || `S${i+1}`), ...vals];
        });
        row = this._write(sheet, row, 0, scoreData);
      }
      row += 2;

      // ── Variance ─────────────────────────────────────────────
      const explVar = pcaResult.explainedVar || [];
      if (explVar.length > 0) {
        row = this._title(sheet, row, "VARIANCE EXPLIQUÉE", 4);
        row = this._header(sheet, row, ["Composante", "Valeur propre", "Variance (%)", "Var. cumulée (%)"]);

        const varData = explVar.map((v, i) => [
          `PC${i+1}`,
          this._n(pcaResult.eigenvalues?.[i], 4),
          this._n(v, 2),
          this._n(pcaResult.cumulativeVar?.[i], 2),
        ]);
        row = this._write(sheet, row, 0, varData);
        row += 2;
      }

      // ── Loadings ─────────────────────────────────────────────
      const loadings = pcaResult.loadings || [];
      // Nombre réel de variables dans les loadings
      const nVarsInLoadings = loadings.length > 0 ? Array.from(loadings[0]).length : 0;
      // On utilise uniquement les noms de variables qui correspondent aux loadings
      const vn = Array.isArray(varNames) ? varNames.slice(0, nVarsInLoadings) : [];

      if (vn.length > 0 && loadings.length > 0) {
        const loadWidth = 1 + nComp;
        row = this._title(sheet, row, "LOADINGS", loadWidth);
        row = this._header(sheet, row, ["Variable", ...pcHdrs]);

        const loadData = vn.map((name, j) => {
          const vals = loadings.map(l => this._n(Array.from(l)[j] ?? 0, 6));
          while (vals.length < nComp) vals.push(0);
          return [String(name), ...vals.slice(0, nComp)];
        });
        row = this._write(sheet, row, 0, loadData);
      }

      await ctx.sync();
    });
  },

  // ══════════════════════════════════════════════════════════════
  //  EXPORT PLS
  // ══════════════════════════════════════════════════════════════

  async exportPLSResults(plsModel, yReal, yPred, vip, varNames, sampleNames) {
    return Excel.run(async (ctx) => {
      const sheet = await this._sheet(ctx, "ChemAI_PLS");
      let row = 0;

      // ── Prédictions ──────────────────────────────────────────
      const yr = Array.isArray(yReal) ? yReal : [];
      const yp = Array.isArray(yPred) ? yPred : [];

      if (yr.length > 0) {
        row = this._title(sheet, row, "PRÉDICTIONS PLS", 4);
        row = this._header(sheet, row, ["Échantillon", "Y réel", "Y prédit", "Résidu"]);

        const predData = yr.map((y, i) => {
          const pred = yp[i] ?? 0;
          return [
            String(sampleNames?.[i] || `S${i+1}`),
            this._n(y, 4),
            this._n(pred, 4),
            this._n(pred - y, 4),
          ];
        });
        row = this._write(sheet, row, 0, predData);
        row += 2;
      }

      // ── VIP Scores ───────────────────────────────────────────
      const vipArr = Array.isArray(vip) ? vip : [];
      if (vipArr.length > 0) {
        row = this._title(sheet, row, "VIP SCORES", 3);
        row = this._header(sheet, row, ["Variable", "VIP Score", "Importance"]);

        const vipData = vipArr
          .map((v, i) => ({
            name: String(varNames?.[i] || `V${i+1}`),
            v: this._n(v, 4),
          }))
          .sort((a, b) => b.v - a.v)
          .map(d => [d.name, d.v, d.v >= 1 ? "Important" : d.v >= 0.8 ? "Modéré" : "Faible"]);

        row = this._write(sheet, row, 0, vipData);
      }

      await ctx.sync();
    });
  },

  // ══════════════════════════════════════════════════════════════
  //  EXPORT CLUSTERS
  // ══════════════════════════════════════════════════════════════

  async exportClusterLabels(labels, sampleNames) {
    return Excel.run(async (ctx) => {
      const sheet = await this._sheet(ctx, "ChemAI_Clusters");

      const lbls = Array.isArray(labels) ? labels : [];

      this._header(sheet, 0, ["Échantillon", "Cluster"]);

      if (lbls.length > 0) {
        const clusterData = lbls.map((l, i) => [
          String(sampleNames?.[i] || `S${i+1}`),
          (typeof l === "number" && isFinite(l) ? l : 0) + 1,
        ]);
        this._write(sheet, 1, 0, clusterData);

        // Couleurs pastels par cluster (lisibles fond blanc, compatibles Excel Desktop)
        const colors = ["#C6EFCE","#FFEB9C","#FFC7CE","#BDD7EE","#E2EFDA","#FCE4D6","#DDEBF7","#FFF2CC"];
        const maxK   = Math.max(...lbls.filter(isFinite), 0) + 1;
        for (let k = 0; k < Math.min(maxK, colors.length); k++) {
          lbls.forEach((l, i) => {
            if (l === k) {
              sheet.getRangeByIndexes(i + 1, 1, 1, 1).format.fill.color = colors[k];
            }
          });
        }
      }

      await ctx.sync();
    });
  },

  // ══════════════════════════════════════════════════════════════
  //  MARQUAGE OUTLIERS
  // ══════════════════════════════════════════════════════════════

  async markOutliers(outlierIndices, startAddress) {
    const idxs = Array.isArray(outlierIndices) ? outlierIndices.filter(Number.isFinite) : [];
    if (idxs.length === 0) return;

    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const ref   = sheet.getRange(startAddress || "A1");
      ref.load("rowIndex");
      await ctx.sync();

      const base = ref.rowIndex || 0;

      idxs.slice(0, 100).forEach(idx => {
        const r = sheet.getRangeByIndexes(base + idx + 1, 0, 1, 20);
        r.format.fill.color = "#FFCCCC";
        const b = r.format.borders.getItem("EdgeLeft");
        b.style  = "Continuous";
        b.color  = "#FF0000";
        b.weight = "Thick";
      });

      await ctx.sync();
    });
  },
};

window.ChemExcel = ChemExcel;