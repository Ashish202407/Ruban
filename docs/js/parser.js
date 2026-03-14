/**
 * The VC Corner Parser - SheetJS-based Excel parser
 * Reads Setup, Checklist, and sector data via hidden column anchors.
 * Includes client-side formula evaluation for when cached results are missing
 * (e.g. files saved by openpyxl which doesn't evaluate formulas).
 */
const Parser = (() => {
  const YEAR_LABELS = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];

  function parse(data) {
    const wb = XLSX.read(data, { type: "array", cellFormula: true, cellNF: true, sheetStubs: true });
    const setup = parseSetup(wb);
    const activeSections = parseChecklist(wb);
    const sectionData = parseSectorData(wb, setup.businessType, activeSections);
    return { setup, activeSections, sectionData };
  }

  function parseSetup(wb) {
    const ws = wb.Sheets["Setup"];
    if (!ws) return { companyName: "Company", businessType: "AI-SaaS", currency: "INR", fyStart: "April" };

    const val = (cell) => {
      const c = ws[cell];
      return c ? (c.v !== undefined ? String(c.v).trim() : "") : "";
    };

    return {
      companyName: val("B3") || "Company",
      businessType: val("B4") || "AI-SaaS",
      currency: val("B5") || "INR",
      fyStart: val("B6") || "April"
    };
  }

  function parseChecklist(wb) {
    const ws = wb.Sheets["Checklist"];
    if (!ws) return [];

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    const active = [];

    for (let r = 2; r <= range.e.r; r++) {
      const dCell = ws[XLSX.utils.encode_cell({ r, c: 3 })];
      const eCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];

      if (dCell && String(dCell.v).trim().toLowerCase() === "yes" && eCell && eCell.v) {
        active.push(String(eCell.v).trim());
      }
    }
    return active;
  }

  function getSheetName(businessType) {
    const map = { "AI-SaaS": "AI-SaaS", "D2C": "D2C", "Healthcare": "Healthcare" };
    return map[businessType] || businessType;
  }

  /**
   * Read a cell's numeric value - handles cached values, formulas, formatted text
   */
  function getCellNumber(ws, r, c) {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    if (!cell) return 0;

    // Best case: cached numeric value
    if (typeof cell.v === "number") return cell.v;

    // Try parsing the formatted string
    if (cell.w) {
      const cleaned = cell.w.replace(/[,$%₹€£\s]/g, "").trim();
      const num = parseFloat(cleaned);
      if (!isNaN(num)) return num;
    }

    // Try parsing raw value
    if (cell.v !== undefined && cell.v !== null) {
      const num = parseFloat(cell.v);
      if (!isNaN(num)) return num;
    }

    return 0;
  }

  /**
   * Parse sector sheet data, then evaluate any missing formula results
   */
  function parseSectorData(wb, businessType, activeSections) {
    const sheetName = getSheetName(businessType);
    const ws = wb.Sheets[sheetName];
    if (!ws) return {};

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    const result = {};

    // Build section anchors from col H
    const sectionAnchors = {};
    for (let r = 0; r <= range.e.r; r++) {
      const hCell = ws[XLSX.utils.encode_cell({ r, c: 7 })];
      if (hCell && hCell.v && typeof hCell.v === "string") {
        const id = hCell.v.trim();
        if (id && REGISTRY[id]) {
          sectionAnchors[r] = id;
        }
      }
    }

    const anchorRows = Object.keys(sectionAnchors).map(Number).sort((a, b) => a - b);

    // Build a full cell value cache for formula evaluation
    // Map: "B5" -> number
    const cellCache = {};
    for (let r = 0; r <= range.e.r; r++) {
      for (let c = 1; c <= 5; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell && typeof cell.v === "number" && !cell.f) {
          cellCache[addr] = cell.v;
        }
      }
    }

    for (let idx = 0; idx < anchorRows.length; idx++) {
      const startRow = anchorRows[idx];
      const sectionId = sectionAnchors[startRow];

      if (!activeSections.includes(sectionId)) continue;

      const endRow = idx + 1 < anchorRows.length ? anchorRows[idx + 1] - 1 : range.e.r;

      const rows = [];
      for (let r = startRow; r <= endRow; r++) {
        const iCell = ws[XLSX.utils.encode_cell({ r, c: 8 })];
        const rowType = iCell ? String(iCell.v).trim() : "";

        if (rowType === "S") continue;

        const aCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
        const label = aCell ? String(aCell.v || "").trim() : "";
        if (!label) continue;

        const values = [];
        for (let c = 1; c <= 5; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws[addr];

          if (!cell) {
            values.push(0);
            continue;
          }

          // Has a formula - always evaluate (openpyxl doesn't cache values)
          if (cell.f) {
            const evaluated = evaluateFormula(cell.f, cellCache);
            if (evaluated !== null) {
              values.push(evaluated);
              // Cache it for downstream formulas
              cellCache[addr] = evaluated;
              continue;
            }
          }

          // Has a cached numeric value (input cell, no formula)
          if (typeof cell.v === "number") {
            values.push(cell.v);
            continue;
          }

          // Fallback: try parsing
          if (cell.v !== undefined && cell.v !== null) {
            const num = parseFloat(cell.v);
            values.push(isNaN(num) ? 0 : num);
          } else {
            values.push(0);
          }
        }

        rows.push({ label, values, type: rowType });
      }

      result[sectionId] = rows;
    }

    return result;
  }

  /**
   * Simple formula evaluator for common Excel patterns:
   *  - Cell references: B5, C10
   *  - Arithmetic: +, -, *, /
   *  - SUM(range)
   *  - IF(cond, true, false)
   *  - Nested expressions with parentheses
   */
  function evaluateFormula(formula, cellCache) {
    try {
      let expr = formula;

      // Handle SUM(B5:B10) patterns
      expr = expr.replace(/SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/gi, (_, col1, row1, col2, row2) => {
        let sum = 0;
        const c = XLSX.utils.decode_col(col1);
        const r1 = parseInt(row1) - 1; // to 0-indexed
        const r2 = parseInt(row2) - 1;
        for (let r = r1; r <= r2; r++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          sum += cellCache[addr] || 0;
        }
        return String(sum);
      });

      // Handle IF(cond, trueVal, falseVal) - simplistic
      expr = expr.replace(/IF\(([^,]+),([^,]+),([^)]+)\)/gi, (_, cond, trueVal, falseVal) => {
        // Evaluate the condition
        const condResult = evalSimple(cond, cellCache);
        return condResult ? trueVal.trim() : falseVal.trim();
      });

      // Replace cell references with values
      expr = expr.replace(/([A-Z]+)(\d+)/g, (match) => {
        // Convert to 0-indexed address
        const decoded = XLSX.utils.decode_cell(match);
        const addr = XLSX.utils.encode_cell(decoded);
        const val = cellCache[addr];
        return val !== undefined ? String(val) : "0";
      });

      // Evaluate arithmetic
      const result = evalArithmetic(expr);
      return isNaN(result) ? null : result;
    } catch (e) {
      return null;
    }
  }

  /**
   * Evaluate simple condition like "B5<>0" or "B5>0"
   */
  function evalSimple(cond, cellCache) {
    try {
      let expr = cond.trim();
      // Replace cell refs
      expr = expr.replace(/([A-Z]+)(\d+)/g, (match) => {
        const decoded = XLSX.utils.decode_cell(match);
        const addr = XLSX.utils.encode_cell(decoded);
        return String(cellCache[addr] || 0);
      });

      // Handle <> (not equal)
      if (expr.includes("<>")) {
        const parts = expr.split("<>");
        return parseFloat(parts[0]) !== parseFloat(parts[1]);
      }
      if (expr.includes(">=")) {
        const parts = expr.split(">=");
        return parseFloat(parts[0]) >= parseFloat(parts[1]);
      }
      if (expr.includes("<=")) {
        const parts = expr.split("<=");
        return parseFloat(parts[0]) <= parseFloat(parts[1]);
      }
      if (expr.includes(">")) {
        const parts = expr.split(">");
        return parseFloat(parts[0]) > parseFloat(parts[1]);
      }
      if (expr.includes("<")) {
        const parts = expr.split("<");
        return parseFloat(parts[0]) < parseFloat(parts[1]);
      }
      if (expr.includes("=")) {
        const parts = expr.split("=");
        return parseFloat(parts[0]) === parseFloat(parts[1]);
      }
      return parseFloat(expr) !== 0;
    } catch (e) {
      return false;
    }
  }

  /**
   * Evaluate arithmetic expression string (no cell refs - already replaced)
   * Uses Function constructor for safe math evaluation
   */
  function evalArithmetic(expr) {
    try {
      // Clean up: remove leading =, whitespace
      expr = expr.replace(/^=/, "").trim();
      // Only allow numbers, operators, parens, decimal points, minus
      if (/[^0-9+\-*/().eE\s]/.test(expr)) return NaN;
      // Evaluate
      return new Function("return (" + expr + ")")();
    } catch (e) {
      return NaN;
    }
  }

  function getCurrencySymbol(code) {
    const map = { INR: "\u20B9", USD: "$", EUR: "\u20AC", GBP: "\u00A3" };
    return map[code] || code;
  }

  return { parse, getCurrencySymbol, YEAR_LABELS };
})();
