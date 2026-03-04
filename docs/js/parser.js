/**
 * Ruban Parser — SheetJS-based Excel parser
 * Reads Setup, Checklist, and sector data via hidden column anchors
 */
const Parser = (() => {
  const YEAR_LABELS = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];

  /**
   * Parse the uploaded workbook
   * @param {ArrayBuffer} data - file content
   * @returns {Object} parsed data: { setup, activeSections, sectionData }
   */
  function parse(data) {
    const wb = XLSX.read(data, { type: "array" });
    const setup = parseSetup(wb);
    const activeSections = parseChecklist(wb);
    const sectionData = parseSectorData(wb, setup.businessType, activeSections);
    return { setup, activeSections, sectionData };
  }

  /**
   * Read Setup sheet — fixed cells B3-B6
   */
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

  /**
   * Read Checklist sheet — find rows where col D = "Yes", extract section ID from col E
   */
  function parseChecklist(wb) {
    const ws = wb.Sheets["Checklist"];
    if (!ws) return [];

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    const active = [];

    for (let r = 2; r <= range.e.r; r++) {
      const dCell = ws[XLSX.utils.encode_cell({ r, c: 3 })]; // col D (0-indexed: 3)
      const eCell = ws[XLSX.utils.encode_cell({ r, c: 4 })]; // col E (0-indexed: 4)

      if (dCell && String(dCell.v).trim().toLowerCase() === "yes" && eCell && eCell.v) {
        active.push(String(eCell.v).trim());
      }
    }
    return active;
  }

  /**
   * Determine the correct sheet name for a business type
   */
  function getSheetName(businessType) {
    const map = {
      "AI-SaaS": "AI-SaaS",
      "D2C": "D2C",
      "Healthcare": "Healthcare"
    };
    return map[businessType] || businessType;
  }

  /**
   * Parse sector sheet data for active sections using col H anchors and col I row types
   */
  function parseSectorData(wb, businessType, activeSections) {
    const sheetName = getSheetName(businessType);
    const ws = wb.Sheets[sheetName];
    if (!ws) return {};

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    const result = {};

    // Build a map of section anchors: row -> sectionId (from col H)
    const sectionAnchors = {};
    for (let r = 0; r <= range.e.r; r++) {
      const hCell = ws[XLSX.utils.encode_cell({ r, c: 7 })]; // col H
      if (hCell && hCell.v && typeof hCell.v === "string") {
        const id = hCell.v.trim();
        if (id && REGISTRY[id]) {
          sectionAnchors[r] = id;
        }
      }
    }

    // Get sorted anchor rows
    const anchorRows = Object.keys(sectionAnchors).map(Number).sort((a, b) => a - b);

    // For each active section, extract its data rows
    for (let idx = 0; idx < anchorRows.length; idx++) {
      const startRow = anchorRows[idx];
      const sectionId = sectionAnchors[startRow];

      if (!activeSections.includes(sectionId)) continue;

      // Section ends at next anchor or end of data
      const endRow = idx + 1 < anchorRows.length ? anchorRows[idx + 1] - 1 : range.e.r;

      const rows = [];
      for (let r = startRow; r <= endRow; r++) {
        const iCell = ws[XLSX.utils.encode_cell({ r, c: 8 })]; // col I = row type
        const rowType = iCell ? String(iCell.v).trim() : "";

        // Skip section headers and empty rows
        if (rowType === "S") continue;

        const aCell = ws[XLSX.utils.encode_cell({ r, c: 0 })]; // col A = label
        const label = aCell ? String(aCell.v || "").trim() : "";
        if (!label) continue;

        // Read year values from cols B-F
        const values = [];
        for (let c = 1; c <= 5; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell) {
            values.push(typeof cell.v === "number" ? cell.v : parseFloat(cell.v) || 0);
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
   * Get currency symbol from currency code
   */
  function getCurrencySymbol(code) {
    const map = { INR: "\u20B9", USD: "$", EUR: "\u20AC", GBP: "\u00A3" };
    return map[code] || code;
  }

  return { parse, getCurrencySymbol, YEAR_LABELS };
})();
