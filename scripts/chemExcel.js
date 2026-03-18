/**
 * ================================================================
 * chemExcel.js — Toutes les interactions Excel via Office.js
 *
 * Fonctions :
 *  - readSelection()      : lit la plage sélectionnée
 *  - parseCSV()           : parse un fichier CSV
 *  - writeResults()       : écrit les résultats dans une feuille
 *  - writeMatrix()        : écrit une matrice avec en-têtes
 *  - insertSVGChart()     : insère un graphique SVG comme image
 *  - markOutliers()       : colorie les cellules aberrantes
 *  - createResultSheet()  : crée ou vide une feuille de résultats
 * ================================================================
 */
"use strict";

const ChemExcel = {

  // ─── Lecture des données ────────────────────────────────────────────────────

  /**
   * Lit la plage sélectionnée dans Excel
   * @param {boolean} hasHeader — la première ligne est-elle un en-tête ?
   * @returns {{ headers: string[], data: number[][], sampleNames: string[], rawValues: any[][] }}
   */
  async readSelection(hasHeader = true) {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync();

      const raw = range.values;
      if (!raw || raw.length < 2) throw new Error("Sélection trop petite (minimum 2 lignes).");

      let headers, dataRows, sampleNames;

      if (hasHeader) {
        headers = raw[0].map((h, i) => h !== null && h !== "" ? String(h) : `V${i+1}`);
        dataRows = raw.slice(1);
      } else {
        headers = raw[0].map((_, i) => `V${i+1}`);
        dataRows = raw;
      }

      // Détecter si la première colonne est un identifiant texte
      const firstColIsText = dataRows.every(row => typeof row[0] === 'string' && isNaN(parseFloat(row[0])));
      if (firstColIsText) {
        sampleNames = dataRows.map(row => String(row[0]));
        headers = headers.slice(1);
        dataRows = dataRows.map(row => row.slice(1));
        if (hasHeader) headers = raw[0].slice(1).map((h, i) => h !== null && h !== "" ? String(h) : `V${i+1}`);
      } else {
        sampleNames = dataRows.map((_, i) => `S${i+1}`);
      }

      // Convertir en nombres
      const data = dataRows.map(row =>
        row.map(v => {
          const n = parseFloat(v);
          return isNaN(n) ? NaN : n;
        })
      );

      return { headers, data, sampleNames, address: range.address, nRows: data.length, nCols: headers.length };
    });
  },

  /**
   * Parse un contenu CSV
   */
  parseCSV(text, separator = ",", hasHeader = true) {
    const lines = text.trim().split(/\r?\n/);
    const rows  = lines.map(l => l.split(separator).map(v => v.trim().replace(/^"|"$/g, '')));

    let headers, dataRows, sampleNames;
    if (hasHeader) {
      headers  = rows[0].map((h, i) => h || `V${i+1}`);
      dataRows = rows.slice(1);
    } else {
      headers  = rows[0].map((_, i) => `V${i+1}`);
      dataRows = rows;
    }

    const firstColIsText = dataRows.every(row => typeof row[0] === 'string' && isNaN(parseFloat(row[0])));
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

  // ─── Écriture des résultats ──────────────────────────────────────────────────

  /**
   * Crée ou vide une feuille avec le nom donné
   */
  async createResultSheet(name) {
    return Excel.run(async (ctx) => {
      let sheet;
      try {
        sheet = ctx.workbook.worksheets.getItem(name);
        sheet.getUsedRange().clear();
      } catch {
        sheet = ctx.workbook.worksheets.add(name);
      }
      sheet.activate();
      await ctx.sync();
      return sheet;
    });
  },

  /**
   * Écrit une matrice dans la feuille active à partir de la cellule (row, col)
   * @param {string}   sheetName
   * @param {string[]} headers
   * @param {any[][]}  data         — lignes de données
   * @param {string[]} rowLabels    — étiquettes de ligne
   * @param {number}   startRow     — 0-indexed
   * @param {number}   startCol
   * @param {string}   [title]
   */
  async writeMatrix(sheetName, headers, data, rowLabels, startRow = 0, startCol = 0, title) {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();

      const ws = sheet.isNullObject
        ? ctx.workbook.worksheets.add(sheetName)
        : sheet;

      let r = startRow;

      // Titre
      if (title) {
        const titleRange = ws.getRangeByIndexes(r, startCol, 1, headers.length + 1);
        titleRange.values = [[title, ...new Array(headers.length).fill("")]];
        titleRange.getCell(0, 0).format.font.bold = true;
        titleRange.getCell(0, 0).format.font.color = "#00E5FF";
        titleRange.getCell(0, 0).format.font.size = 12;
        r++;
      }

      // En-têtes
      const headerValues = rowLabels ? ["", ...headers] : headers;
      const headerRange = ws.getRangeByIndexes(r, startCol, 1, headerValues.length);
      headerRange.values = [headerValues];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#0D1825";
      headerRange.format.font.color = "#00E5FF";
      r++;

      // Données
      const allRows = data.map((row, i) => rowLabels ? [rowLabels[i], ...row] : row);
      if (allRows.length > 0) {
        const dataRange = ws.getRangeByIndexes(r, startCol, allRows.length, allRows[0].length);
        dataRange.values = allRows;
        // Alternance de couleurs
        for (let ri = 0; ri < allRows.length; ri++) {
          const rowRange = ws.getRangeByIndexes(r + ri, startCol, 1, allRows[0].length);
          rowRange.format.fill.color = ri % 2 === 0 ? "#0D1825" : "#122035";
          rowRange.format.font.color = "#B0C8DA";
        }
      }

      await ctx.sync();
      return { endRow: r + allRows.length };
    });
  },

  /**
   * Exporte les scores PCA dans Excel
   */
  async exportPCAResults(pcaResult, sampleNames, varNames) {
    const sheetName = "ChemAI_PCA";
    await this.createResultSheet(sheetName);

    const nComp = pcaResult.nComp;
    const compHeaders = Array.from({length: nComp}, (_, i) => `PC${i+1}`);

    // Scores
    await this.writeMatrix(
      sheetName,
      compHeaders,
      pcaResult.scores.map(row => row.map(v => +v.toFixed(6))),
      sampleNames,
      0, 0,
      "=== SCORES PCA ==="
    );

    // Variance
    const varData = pcaResult.explainedVar.map((v, i) => [
      pcaResult.eigenvalues[i].toFixed(4),
      v.toFixed(2),
      pcaResult.cumulativeVar[i].toFixed(2),
    ]);
    const res = await this.writeMatrix(sheetName, ["Valeur propre", "Variance (%)", "Var. cumulée (%)"],
      varData, compHeaders.map((_, i) => `PC${i+1}`),
      pcaResult.scores.length + 4, 0, "=== VARIANCE EXPLIQUÉE ==="
    );

    // Loadings
    await this.writeMatrix(
      sheetName, varNames,
      pcaResult.loadings.map(row => Array.from(row).map(v => +v.toFixed(6))),
      compHeaders.map((_, i) => `PC${i+1}`),
      res.endRow + 3, 0, "=== LOADINGS ==="
    );
  },

  /**
   * Exporte les résultats PLS
   */
  async exportPLSResults(plsModel, yReal, yPred, vip, varNames, sampleNames) {
    const sheetName = "ChemAI_PLS";
    await this.createResultSheet(sheetName);

    // Prédictions
    const predData = yReal.map((y, i) => [y, +yPred[i].toFixed(4), +(yPred[i]-y).toFixed(4)]);
    const res = await this.writeMatrix(sheetName, ["Y réel", "Y prédit", "Résidu"],
      predData, sampleNames, 0, 0, "=== PRÉDICTIONS PLS ==="
    );

    // VIP
    const vipData = vip.map((v, i) => [+(v).toFixed(4), v >= 1 ? "Important" : v >= 0.8 ? "Modéré" : "Faible"]);
    const sortedVIP = vipData.map((d, i) => ({name: varNames[i], ...d}))
      .sort((a, b) => b[0] - a[0]);
    await this.writeMatrix(sheetName, ["VIP Score", "Importance"],
      sortedVIP.map(d => [d[0], d[1]]),
      sortedVIP.map(d => d.name),
      res.endRow + 3, 0, "=== VIP SCORES ==="
    );
  },

  /**
   * Exporte les labels de cluster et marque les cellules
   */
  async exportClusterLabels(labels, sampleNames) {
    const sheetName = "ChemAI_Clusters";
    await this.createResultSheet(sheetName);

    const clusterColors = ["#00E5FF","#00E676","#FF9800","#CE93D8","#FF5252","#FFEB3B","#40C4FF","#FF80AB"];
    const data = labels.map((l, i) => [sampleNames[i] || `S${i+1}`, l + 1]);

    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      const sheet = ws.isNullObject ? ctx.workbook.worksheets.add(sheetName) : ws;

      // En-têtes
      sheet.getRangeByIndexes(0, 0, 1, 2).values = [["Échantillon", "Cluster"]];
      sheet.getRangeByIndexes(0, 0, 1, 2).format.font.bold = true;

      data.forEach(([name, cluster], i) => {
        const row = sheet.getRangeByIndexes(i + 1, 0, 1, 2);
        row.values = [[name, cluster]];
        const col = clusterColors[(cluster - 1) % clusterColors.length];
        row.getCell(0, 1).format.fill.color = col + "55"; // semi-transparent
      });

      await ctx.sync();
    });
  },

  /**
   * Marque les outliers dans la feuille active avec une bordure rouge
   */
  async markOutliers(outlierIndices, startAddress) {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const refRange = sheet.getRange(startAddress);
      refRange.load("rowIndex");
      await ctx.sync();

      for (const idx of outlierIndices) {
        const row = sheet.getRangeByIndexes(refRange.rowIndex + idx + 1, 0, 1, 20);
        row.format.fill.color = "#FF525220";
        row.format.borders.getItem("EdgeLeft").style = "Continuous";
        row.format.borders.getItem("EdgeLeft").color = "#FF5252";
        row.format.borders.getItem("EdgeLeft").weight = "Thick";
      }
      await ctx.sync();
    });
  },

  /**
   * Insère un SVG comme image dans Excel (via conversion canvas → blob)
   * Note : Office.js ne supporte pas directement SVG.
   * On encode en data URI et on insère via insertPictureFromBase64.
   */
  async insertChartFromSVG(svgString, sheetName, cellAddress) {
    // Conversion SVG → PNG via Canvas (navigateur)
    return new Promise((resolve, reject) => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width  = 800;
        canvas.height = 520;
        const ctx2d = canvas.getContext("2d");

        const img = new Image();
        const blob = new Blob([svgString], { type: "image/svg+xml" });
        const url  = URL.createObjectURL(blob);

        img.onload = async () => {
          ctx2d.fillStyle = "#080F1A";
          ctx2d.fillRect(0, 0, canvas.width, canvas.height);
          ctx2d.drawImage(img, 0, 0, canvas.width, canvas.height);
          URL.revokeObjectURL(url);

          const base64 = canvas.toDataURL("image/png").split(",")[1];

          await Excel.run(async (exCtx) => {
            const ws = exCtx.workbook.worksheets.getItemOrNullObject(sheetName);
            await exCtx.sync();
            const sheet = ws.isNullObject ? exCtx.workbook.worksheets.add(sheetName) : ws;
            const range = sheet.getRange(cellAddress);
            range.load("top,left");
            await exCtx.sync();

            // Insérer l'image
            const shape = sheet.shapes.addImage(base64);
            shape.top  = range.top;
            shape.left = range.left;
            shape.width  = 400;
            shape.height = 260;
            await exCtx.sync();
          });

          resolve();
        };

        img.onerror = reject;
        img.src = url;
      } catch (e) { reject(e); }
    });
  },
};

window.ChemExcel = ChemExcel;
