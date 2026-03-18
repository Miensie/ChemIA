/**
 * ================================================================
 * chemExcel.js — Interactions Excel via Office.js
 * Compatible Excel Desktop (Windows/macOS) et Excel Online
 *
 * Règles Office.js Desktop :
 *  - Un seul Excel.run par opération (jamais deux Excel.run séquentiels)
 *  - getItemOrNullObject() + isNullObject après sync
 *  - Toutes les écritures batchées avant ctx.sync()
 *  - getUsedRangeOrNullObject() pour feuille vide
 * ================================================================
 */
"use strict";

const ChemExcel = {

  // ─── Utilitaire interne ──────────────────────────────────────────────────────

  /** Dans un Excel.run actif, obtient ou crée une feuille (la vide si elle existe). */
  async _getOrCreateSheet(ctx, name) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    ws.load("isNullObject");
    await ctx.sync();

    if (ws.isNullObject) {
      const s = ctx.workbook.worksheets.add(name);
      s.activate();
      return s;
    }
    // Vider sans crasher si la feuille est vide
    const used = ws.getUsedRangeOrNullObject();
    used.load("isNullObject");
    await ctx.sync();
    if (!used.isNullObject) {
      used.clear("All");
      await ctx.sync();
    }
    ws.activate();
    return ws;
  },

  _round(v, d) {
    if (v === null || v === undefined || !isFinite(v)) return 0;
    return parseFloat(Number(v).toFixed(d));
  },

  _title(sheet, text, row, w) {
    const r = sheet.getRangeByIndexes(row, 0, 1, Math.max(w, 1));
    r.values = [[`=== ${text} ===`]];
    r.getCell(0, 0).format.font.bold  = true;
    r.getCell(0, 0).format.font.color = "#005577";
    r.getCell(0, 0).format.font.size  = 11;
    return row + 1;
  },

  _header(range) {
    range.format.font.bold  = true;
    range.format.font.color = "#003040";
    range.format.fill.color = "#CCEEEE";
  },

  // ─── Lecture plage Excel ────────────────────────────────────────────────────

  async readSelection(hasHeader = true) {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address"]);
      await ctx.sync();

      const raw = range.values;
      if (!raw || raw.length < 2) throw new Error("Sélection trop petite (minimum 2 lignes).");

      let headers, dataRows, sampleNames;

      if (hasHeader) {
        headers  = raw[0].map((h, i) => (h !== null && h !== "") ? String(h) : `V${i+1}`);
        dataRows = raw.slice(1);
      } else {
        headers  = raw[0].map((_, i) => `V${i+1}`);
        dataRows = raw;
      }

      const firstColIsText = dataRows.every(
        row => typeof row[0] === "string" && isNaN(parseFloat(row[0]))
      );
      if (firstColIsText) {
        sampleNames = dataRows.map(row => String(row[0]));
        headers     = headers.slice(1);
        dataRows    = dataRows.map(row => row.slice(1));
        if (hasHeader) headers = raw[0].slice(1).map((h, i) => (h !== null && h !== "") ? String(h) : `V${i+1}`);
      } else {
        sampleNames = dataRows.map((_, i) => `S${i+1}`);
      }

      const data = dataRows.map(row =>
        row.map(v => { const n = parseFloat(v); return isNaN(n) ? NaN : n; })
      );

      return { headers, data, sampleNames, address: range.address, nRows: data.length, nCols: headers.length };
    });
  },

  // ─── Parse CSV (pur JS) ──────────────────────────────────────────────────────

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
      row => typeof row[0] === "string" && isNaN(parseFloat(row[0]))
    );
    if (firstColIsText) {
      sampleNames = dataRows.map(row => row[0]);
      headers     = headers.slice(1);
      dataRows    = dataRows.map(row => row.slice(1));
    } else {
      sampleNames = dataRows.map((_, i) => `S${i+1}`);
    }

    const data = dataRows.map(row =>
      row.map(v => { const n = parseFloat(v); return isNaN(n) ? NaN : n; })
    );

    return { headers, data, sampleNames, nRows: data.length, nCols: headers.length };
  },

  // ─── Export PCA ─────────────────────────────────────────────────────────────

  async exportPCAResults(pcaResult, sampleNames, varNames) {
    return Excel.run(async (ctx) => {
      const sheet     = await this._getOrCreateSheet(ctx, "ChemAI_PCA");
      const nComp     = pcaResult.nComp;
      const compHdrs  = Array.from({length: nComp}, (_, i) => `PC${i+1}`);
      let row = 0;

      // Scores
      row = this._title(sheet, "SCORES PCA", row, nComp + 1);
      sheet.getRangeByIndexes(row, 0, 1, nComp + 1).values = [["Échantillon", ...compHdrs]];
      this._header(sheet.getRangeByIndexes(row, 0, 1, nComp + 1));
      row++;
      const scoreRows = pcaResult.scores.map((s, i) => [
        sampleNames?.[i] || `S${i + 1}`,
        ...Array.from(s).map(v => this._round(v, 6)),
      ]);
      if (scoreRows.length) {
        sheet.getRangeByIndexes(row, 0, scoreRows.length, nComp + 1).values = scoreRows;
        row += scoreRows.length;
      }
      row += 2;

      // Variance
      row = this._title(sheet, "VARIANCE EXPLIQUÉE", row, 4);
      sheet.getRangeByIndexes(row, 0, 1, 4).values = [["Composante","Valeur propre","Variance (%)","Var. cumulée (%)"]];
      this._header(sheet.getRangeByIndexes(row, 0, 1, 4));
      row++;
      const varRows = pcaResult.explainedVar.map((v, i) => [
        `PC${i + 1}`,
        this._round(pcaResult.eigenvalues[i], 4),
        this._round(v, 2),
        this._round(pcaResult.cumulativeVar[i], 2),
      ]);
      sheet.getRangeByIndexes(row, 0, varRows.length, 4).values = varRows;
      row += varRows.length + 2;

      // Loadings
      const vn = varNames || [];
      if (vn.length) {
        row = this._title(sheet, "LOADINGS", row, nComp + 1);
        sheet.getRangeByIndexes(row, 0, 1, nComp + 1).values = [["Variable", ...compHdrs]];
        this._header(sheet.getRangeByIndexes(row, 0, 1, nComp + 1));
        row++;
        const loadRows = vn.map((name, j) => [
          name,
          ...pcaResult.loadings.map(l => this._round(Array.from(l)[j] ?? 0, 6)),
        ]);
        sheet.getRangeByIndexes(row, 0, loadRows.length, nComp + 1).values = loadRows;
      }

      await ctx.sync();  // UN SEUL sync à la fin
    });
  },

  // ─── Export PLS ─────────────────────────────────────────────────────────────

  async exportPLSResults(plsModel, yReal, yPred, vip, varNames, sampleNames) {
    return Excel.run(async (ctx) => {
      const sheet = await this._getOrCreateSheet(ctx, "ChemAI_PLS");
      let row = 0;

      // Prédictions
      row = this._title(sheet, "PRÉDICTIONS PLS", row, 4);
      sheet.getRangeByIndexes(row, 0, 1, 4).values = [["Échantillon","Y réel","Y prédit","Résidu"]];
      this._header(sheet.getRangeByIndexes(row, 0, 1, 4));
      row++;
      const yr = yReal || [], yp = yPred || [];
      const predRows = yr.map((y, i) => [
        sampleNames?.[i] || `S${i + 1}`,
        this._round(y, 4),
        this._round(yp[i] ?? 0, 4),
        this._round((yp[i] ?? 0) - y, 4),
      ]);
      if (predRows.length) {
        sheet.getRangeByIndexes(row, 0, predRows.length, 4).values = predRows;
        row += predRows.length;
      }
      row += 2;

      // VIP
      if (vip && vip.length) {
        row = this._title(sheet, "VIP SCORES", row, 3);
        sheet.getRangeByIndexes(row, 0, 1, 3).values = [["Variable","VIP Score","Importance"]];
        this._header(sheet.getRangeByIndexes(row, 0, 1, 3));
        row++;
        const vipRows = vip
          .map((v, i) => ({ name: varNames?.[i] || `V${i + 1}`, v: this._round(v, 4) }))
          .sort((a, b) => b.v - a.v)
          .map(d => [d.name, d.v, d.v >= 1 ? "Important" : d.v >= 0.8 ? "Modéré" : "Faible"]);
        sheet.getRangeByIndexes(row, 0, vipRows.length, 3).values = vipRows;
      }

      await ctx.sync();
    });
  },

  // ─── Export Clusters ─────────────────────────────────────────────────────────

  async exportClusterLabels(labels, sampleNames) {
    return Excel.run(async (ctx) => {
      const sheet = await this._getOrCreateSheet(ctx, "ChemAI_Clusters");

      sheet.getRangeByIndexes(0, 0, 1, 2).values = [["Échantillon","Cluster"]];
      this._header(sheet.getRangeByIndexes(0, 0, 1, 2));

      if (labels && labels.length) {
        // Écriture en une seule opération (pas de boucle par ligne)
        const dataRows = labels.map((l, i) => [sampleNames?.[i] || `S${i + 1}`, l + 1]);
        sheet.getRangeByIndexes(1, 0, dataRows.length, 2).values = dataRows;

        // Couleur par cluster — couleurs Excel standard, lisibles sur fond blanc
        const clusterColors = ["#C6EFCE","#FFEB9C","#FFC7CE","#BDD7EE","#E2EFDA","#FCE4D6","#DDEBF7","#FFF2CC"];
        const maxK = Math.max(...labels) + 1;
        for (let k = 0; k < Math.min(maxK, clusterColors.length); k++) {
          labels.forEach((l, i) => {
            if (l === k) {
              sheet.getRangeByIndexes(i + 1, 1, 1, 1).format.fill.color = clusterColors[k];
            }
          });
        }
      }

      await ctx.sync();
    });
  },

  // ─── Marquage outliers ───────────────────────────────────────────────────────

  async markOutliers(outlierIndices, startAddress) {
    if (!outlierIndices || !outlierIndices.length) return;

    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const ref   = sheet.getRange(startAddress || "A1");
      ref.load("rowIndex");
      await ctx.sync();

      const base = ref.rowIndex || 0;
      // Batch toutes les opérations de style avant sync
      outlierIndices.slice(0, 100).forEach(idx => {
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