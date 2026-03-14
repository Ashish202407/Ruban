/**
 * The VC Corner Renderer - Dashboard DOM builder
 * Category grouping, radar scorecard, YoY badges, accent stripes, PDF export
 */
const Renderer = (() => {
  let chartInstances = [];

  function render(parsed) {
    const { setup, activeSections, sectionData } = parsed;
    const currencySymbol = Parser.getCurrencySymbol(setup.currency);

    const dashboard = document.getElementById("dashboard");
    dashboard.innerHTML = "";

    // Header
    dashboard.appendChild(buildHeader(setup));

    // Radar Scorecard (before KPI strip)
    const radarCard = buildRadarScorecard(setup.businessType, activeSections, sectionData, currencySymbol);
    if (radarCard) dashboard.appendChild(radarCard);

    // KPI strip
    const kpis = collectKPIs(activeSections, sectionData, currencySymbol);
    if (kpis.length > 0) {
      dashboard.appendChild(buildKPIStrip(kpis));
    }

    // Chart grid - grouped by category
    const grid = document.createElement("div");
    grid.className = "chart-grid";

    // Sort sections by category order
    const categoryOrder = typeof CATEGORY_ORDER !== "undefined" ? CATEGORY_ORDER : [];
    const sortedSections = [...activeSections].sort((a, b) => {
      const catA = REGISTRY[a]?.category || "zzz";
      const catB = REGISTRY[b]?.category || "zzz";
      const idxA = categoryOrder.indexOf(catA);
      const idxB = categoryOrder.indexOf(catB);
      return (idxA === -1 ? 999 : idxA) - (idxB === -1 ? 999 : idxB);
    });

    let lastCategory = null;
    let cardIndex = 0;

    for (const sectionId of sortedSections) {
      const config = REGISTRY[sectionId];
      if (!config) continue;
      const data = sectionData[sectionId];
      if (!data || data.length === 0) continue;

      // Category group header
      const category = config.category || "Other";
      if (category !== lastCategory) {
        grid.appendChild(buildCategoryHeader(category));
        lastCategory = category;
      }

      const card = buildSectionCard(sectionId, config, data, currencySymbol, cardIndex);
      grid.appendChild(card);
      cardIndex++;
    }

    dashboard.appendChild(grid);
    dashboard.appendChild(buildPrivacyBadge());

    // Wire up PDF export button
    setupExportPDF();
  }

  function buildHeader(setup) {
    const header = document.createElement("div");
    header.className = "dash-header";
    header.innerHTML = `
      <div class="dash-header-content">
        <h1 class="dash-title">${esc(setup.companyName)}</h1>
        <div class="dash-meta">
          <span class="meta-tag">${esc(setup.businessType)}</span>
          <span class="meta-tag">${esc(setup.currency)}</span>
          <span class="meta-tag">FY Start: ${esc(setup.fyStart)}</span>
          <span class="meta-tag">5-Year Projection</span>
        </div>
      </div>
    `;
    return header;
  }

  function collectKPIs(activeSections, sectionData, currencySymbol) {
    const kpis = [];

    for (const sectionId of activeSections) {
      const config = REGISTRY[sectionId];
      if (!config || !config.kpis) continue;
      const data = sectionData[sectionId];
      if (!data) continue;

      for (const kpiDef of config.kpis) {
        if (kpis.length >= 8) break;

        const row = data.find(r => r.label === kpiDef.row);
        if (!row) continue;

        const lastVal = row.values[row.values.length - 1];
        const firstVal = row.values[0];

        let displayVal, subtext = "";
        if (kpiDef.format === "growth") {
          if (firstVal && firstVal !== 0) {
            const growth = ((lastVal - firstVal) / Math.abs(firstVal)) * 100;
            displayVal = (growth >= 0 ? "+" : "") + growth.toFixed(0) + "%";
            subtext = "5Y growth";
          } else {
            continue;
          }
        } else if (kpiDef.format === "percent") {
          displayVal = (lastVal * 100).toFixed(1) + "%";
          subtext = "Year 5";
        } else if (kpiDef.format === "currency") {
          displayVal = fmtCurrency(lastVal, currencySymbol);
          subtext = "Year 5";
        } else if (kpiDef.format === "number") {
          displayVal = fmtNumber(lastVal);
          subtext = "Year 5";
        } else if (kpiDef.format === "ratio") {
          displayVal = lastVal.toFixed(1) + "x";
          subtext = "Year 5";
        } else if (kpiDef.format === "months") {
          displayVal = lastVal.toFixed(0) + " mo";
          subtext = "Year 5";
        } else {
          displayVal = String(lastVal);
        }

        kpis.push({
          label: kpiDef.label,
          value: displayVal,
          subtext,
          trend: lastVal >= firstVal ? "up" : "down"
        });
      }
      if (kpis.length >= 8) break;
    }
    return kpis;
  }

  function buildKPIStrip(kpis) {
    const strip = document.createElement("div");
    strip.className = "kpi-strip";

    for (const kpi of kpis) {
      const card = document.createElement("div");
      card.className = "kpi-card";
      const trendIcon = kpi.trend === "up" ? "&#9650;" : "&#9660;";
      const trendClass = kpi.trend === "up" ? "trend-up" : "trend-down";
      card.innerHTML = `
        <div class="kpi-label">${esc(kpi.label)}</div>
        <div class="kpi-value">${kpi.value} <span class="kpi-trend ${trendClass}">${trendIcon}</span></div>
      `;
      strip.appendChild(card);
    }
    return strip;
  }

  function buildCategoryHeader(category) {
    const el = document.createElement("div");
    el.className = "category-group-header";
    el.textContent = category;
    return el;
  }

  function buildSectionCard(sectionId, config, data, currencySymbol, index) {
    const card = document.createElement("div");
    card.className = "section-card";

    // Accent stripe
    const palette = Charts.PALETTE;
    const accentColor = palette[index % palette.length];
    card.style.setProperty("--section-accent", accentColor);

    // Animation delay
    card.style.animationDelay = (index * 60) + "ms";

    // Title row with YoY badge
    const titleRow = document.createElement("div");
    titleRow.className = "section-title";
    titleRow.textContent = config.title;

    // YoY badge - compute Year4→Year5 delta of primary metric
    const yoyBadge = computeYoYBadge(config, data);
    if (yoyBadge) {
      titleRow.appendChild(yoyBadge);
    }
    card.appendChild(titleRow);

    // Build datasets
    const datasets = [];
    for (const rowLabel of (config.chartRows || [])) {
      const row = data.find(r => r.label === rowLabel);
      if (row) {
        datasets.push({ label: row.label, values: [...row.values] });
      }
    }

    // HTML legend (outside canvas - no overlap with y-axis)
    const isGauge = config.chartType === "gauge";
    if (datasets.length > 0 && config.chartType !== "pie" && !isGauge) {
      const legend = document.createElement("div");
      legend.className = "chart-legend";
      datasets.forEach((ds, i) => {
        const item = document.createElement("span");
        item.className = "chart-legend-item";
        const color = palette[i % palette.length];
        item.innerHTML = `<span class="chart-legend-dot" style="background:${color}"></span>${esc(ds.label)}`;
        legend.appendChild(item);
      });
      card.appendChild(legend);
    }

    // Chart
    const chartWrap = document.createElement("div");
    chartWrap.className = "chart-wrap";
    const canvas = document.createElement("canvas");
    canvas.id = `chart-${sectionId}`;
    chartWrap.appendChild(canvas);
    card.appendChild(chartWrap);

    if (datasets.length > 0) {
      const chart = Charts.create(canvas, config.chartType, Parser.YEAR_LABELS, datasets, currencySymbol, config.yFormat || "currency");
      chartInstances.push(chart);
    }

    // Table toggle
    const tableToggle = document.createElement("button");
    tableToggle.className = "table-toggle";
    tableToggle.innerHTML = `
      <svg viewBox="0 0 16 16" fill="none"><path d="M4 6l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
      View Data
    `;
    tableToggle.addEventListener("click", () => {
      const table = card.querySelector(".data-table");
      const hidden = table.style.display === "none";
      table.style.display = hidden ? "block" : "none";
      tableToggle.classList.toggle("is-open", hidden);
      tableToggle.innerHTML = hidden
        ? `<svg viewBox="0 0 16 16" fill="none"><path d="M4 6l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg> Hide Data`
        : `<svg viewBox="0 0 16 16" fill="none"><path d="M4 6l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg> View Data`;
    });
    card.appendChild(tableToggle);

    const table = buildDataTable(data, currencySymbol);
    table.style.display = "none";
    card.appendChild(table);

    return card;
  }

  function computeYoYBadge(config, data) {
    if (!config.kpis || config.kpis.length === 0) return null;

    // Use the first KPI's row for YoY
    const primaryKpi = config.kpis[0];
    const row = data.find(r => r.label === primaryKpi.row);
    if (!row || row.values.length < 5) return null;

    const y4 = row.values[3];
    const y5 = row.values[4];
    if (y4 === 0 || y4 === null || y4 === undefined) return null;

    const delta = ((y5 - y4) / Math.abs(y4)) * 100;
    if (isNaN(delta) || !isFinite(delta)) return null;

    const badge = document.createElement("span");
    const isPositive = delta >= 0;
    badge.className = `yoy-badge ${isPositive ? "positive" : "negative"}`;
    badge.textContent = (isPositive ? "+" : "") + delta.toFixed(0) + "% YoY";
    return badge;
  }

  // ─── Radar Scorecard ───

  function buildRadarScorecard(businessType, activeSections, sectionData, currencySymbol) {
    const axesDef = getRadarAxes(businessType);
    if (!axesDef) return null;

    const scores = [];
    const labels = [];
    const metricRows = [];

    for (const axis of axesDef) {
      labels.push(axis.label);

      // Try to find the metric in active data
      let score = null;
      let rawValue = null;

      for (const sectionId of activeSections) {
        const data = sectionData[sectionId];
        if (!data) continue;
        const row = data.find(r => r.label === axis.rowLabel);
        if (row) {
          rawValue = row.values[row.values.length - 1];
          // Normalize 0-100 vs benchmark
          score = normalizeScore(rawValue, axis);
          break;
        }
      }

      scores.push(score !== null ? score : 0);
      metricRows.push({
        label: axis.label,
        value: rawValue !== null ? formatRadarValue(rawValue, axis.format) : "-",
        hasData: score !== null
      });
    }

    // If all zeros or no data, don't render
    if (scores.every(s => s === 0)) return null;

    const overallScore = Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);

    // Build card
    const card = document.createElement("div");
    card.className = "radar-scorecard-card";

    // Left side: radar chart
    const chartWrap = document.createElement("div");
    chartWrap.className = "radar-chart-wrap";
    const canvas = document.createElement("canvas");
    canvas.id = "chart-radar-scorecard";
    chartWrap.appendChild(canvas);
    card.appendChild(chartWrap);

    // Right side: score + metrics
    const info = document.createElement("div");
    info.className = "radar-info";

    const scoreDisplay = document.createElement("div");
    scoreDisplay.className = "radar-score-display";
    scoreDisplay.innerHTML = `${overallScore}<small>/100</small>`;
    info.appendChild(scoreDisplay);

    const subtitle = document.createElement("div");
    subtitle.style.cssText = "font-size: 0.78rem; color: var(--text-tertiary); margin-bottom: 0.5rem;";
    subtitle.textContent = "VC Readiness Score";
    info.appendChild(subtitle);

    const metricsList = document.createElement("div");
    metricsList.className = "radar-metrics-list";
    for (const m of metricRows) {
      const row = document.createElement("div");
      row.className = "radar-metric-row";
      row.innerHTML = `
        <span>${esc(m.label)}</span>
        <span class="metric-val" style="${!m.hasData ? 'opacity:0.4' : ''}">${m.value}</span>
      `;
      metricsList.appendChild(row);
    }
    info.appendChild(metricsList);
    card.appendChild(info);

    // Render radar chart
    const benchmarks = axesDef.map(a => a.benchmark);
    const radarChart = Charts.create(canvas, "radar", labels, [
      { label: "Score", values: scores },
      { label: "Benchmark", values: benchmarks }
    ]);
    chartInstances.push(radarChart);

    return card;
  }

  function getRadarAxes(businessType) {
    if (businessType === "AI-SaaS") {
      return [
        { label: "Revenue Growth", rowLabel: "Revenue Growth %", format: "percent", min: 0, max: 3, benchmark: 75 },
        { label: "NDR", rowLabel: "NDR %", format: "percent", min: 0.8, max: 1.5, benchmark: 70 },
        { label: "Gross Margin", rowLabel: "Gross Margin %", format: "percent", min: 0.5, max: 0.9, benchmark: 65 },
        { label: "Rule of 40", rowLabel: "Rule of 40 Score", format: "percent", min: -0.5, max: 1, benchmark: 60 },
        { label: "LTV/CAC", rowLabel: "LTV / CAC", format: "ratio", min: 0, max: 5, benchmark: 55 },
        { label: "Burn Multiple", rowLabel: "Burn Multiple", format: "ratio_inv", min: 0, max: 3, benchmark: 50 },
      ];
    }
    if (businessType === "D2C") {
      return [
        { label: "Revenue Growth", rowLabel: "Total Net Revenue", format: "growth", min: 0, max: 5, benchmark: 70 },
        { label: "Gross Margin", rowLabel: "Gross Margin %", format: "percent", min: 0.3, max: 0.7, benchmark: 60 },
        { label: "LTV/CAC", rowLabel: "LTV / CAC", format: "ratio", min: 0, max: 5, benchmark: 55 },
        { label: "Repeat Rate", rowLabel: "Repeat Purchase Rate", format: "percent", min: 0, max: 0.6, benchmark: 50 },
        { label: "ROAS", rowLabel: "Blended ROAS", format: "ratio", min: 0, max: 6, benchmark: 55 },
        { label: "EBITDA Margin", rowLabel: "EBITDA Margin %", format: "percent", min: -0.3, max: 0.3, benchmark: 45 },
      ];
    }
    if (businessType === "Healthcare") {
      return [
        { label: "Occupancy", rowLabel: "Occupancy Rate %", format: "percent", min: 0.4, max: 0.95, benchmark: 70 },
        { label: "ARPOB Growth", rowLabel: "ARPOB (Avg Rev per Occupied Bed)", format: "growth", min: 0, max: 2, benchmark: 60 },
        { label: "Gross Margin", rowLabel: "Gross Profit", format: "growth", min: 0, max: 3, benchmark: 55 },
        { label: "EBITDA Margin", rowLabel: "EBITDA", format: "growth", min: 0, max: 3, benchmark: 50 },
        { label: "OPD Growth", rowLabel: "Annual OPD Visits", format: "growth", min: 0, max: 2, benchmark: 55 },
        { label: "Collection Rate", rowLabel: "Collection Rate %", format: "percent", min: 0.7, max: 1, benchmark: 65 },
      ];
    }
    return null;
  }

  function normalizeScore(value, axis) {
    if (value === null || value === undefined) return null;
    const range = axis.max - axis.min;
    if (range === 0) return 50;
    let normalized;
    if (axis.format === "ratio_inv") {
      // Lower is better (burn multiple)
      normalized = 1 - ((value - axis.min) / range);
    } else {
      normalized = (value - axis.min) / range;
    }
    return Math.max(0, Math.min(100, Math.round(normalized * 100)));
  }

  function formatRadarValue(value, format) {
    if (format === "percent") return (value * 100).toFixed(1) + "%";
    if (format === "ratio" || format === "ratio_inv") return value.toFixed(1) + "x";
    if (format === "growth") return "+" + ((value) * 100).toFixed(0) + "%";
    return String(value);
  }

  // ─── Export PDF ───

  function setupExportPDF() {
    // Check if button already exists
    const existing = document.querySelector(".export-pdf-btn");
    if (existing) return;

    const navRight = document.querySelector(".nav-right");
    if (!navRight) return;

    const btn = document.createElement("button");
    btn.className = "export-pdf-btn";
    btn.innerHTML = `
      <svg viewBox="0 0 16 16" fill="none">
        <path d="M4 14h8M8 2v9M5 8l3 3 3-3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
      Export PDF
    `;
    btn.addEventListener("click", () => {
      document.body.classList.add("print-mode");
      window.print();
    });

    // Clean up after print
    window.addEventListener("afterprint", () => {
      document.body.classList.remove("print-mode");
    });

    navRight.insertBefore(btn, navRight.firstChild);
  }

  // ─── Data Table ───

  function buildDataTable(rows, currencySymbol) {
    const wrap = document.createElement("div");
    wrap.className = "data-table";

    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    headerRow.innerHTML = `<th></th>${Parser.YEAR_LABELS.map(y => `<th>${y}</th>`).join("")}`;
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (const row of rows) {
      const tr = document.createElement("tr");
      const rowClass = row.type === "T" ? "row-total" : row.type === "P" ? "row-pct" : row.type === "F" ? "row-formula" : "";
      if (rowClass) tr.className = rowClass;

      tr.innerHTML = `<td class="row-label">${esc(row.label)}</td>`;

      for (const val of row.values) {
        const td = document.createElement("td");
        td.className = "row-value";
        if (row.type === "P") {
          td.textContent = (val * 100).toFixed(1) + "%";
        } else {
          td.textContent = fmtCurrency(val, currencySymbol);
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    wrap.appendChild(table);
    return wrap;
  }

  function buildPrivacyBadge() {
    const badge = document.createElement("div");
    badge.className = "privacy-badge";
    badge.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
        <path d="M8 1L2 4v4c0 3.5 2.5 6.4 6 7 3.5-.6 6-3.5 6-7V4L8 1z" stroke="currentColor" stroke-width="1.5" fill="none"/>
        <path d="M5.5 8l2 2 3-3.5" stroke="currentColor" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
      <span>Your data never leaves your device - Dashboard Generator</span>
    `;
    return badge;
  }

  function cleanup() {
    chartInstances.forEach(c => Charts.destroy(c));
    chartInstances = [];
    // Remove export button on cleanup
    const btn = document.querySelector(".export-pdf-btn");
    if (btn) btn.remove();
  }

  // ─── Formatting ───
  function fmtCurrency(val, sym) {
    if (val === 0 || val === null || val === undefined) return sym + "0";
    const abs = Math.abs(val);
    const sign = val < 0 ? "-" : "";
    // Indian notation for ₹, standard K/M/B for others
    if (sym === "₹") {
      if (abs >= 10000000) return sign + sym + (abs / 10000000).toFixed(2) + " Cr";
      if (abs >= 100000) return sign + sym + (abs / 100000).toFixed(2) + " L";
      if (abs >= 1000) return sign + sym + commas((abs / 1000).toFixed(1)) + "K";
      return sign + sym + commas(abs.toFixed(2));
    }
    if (abs >= 1000000000) return sign + sym + (abs / 1000000000).toFixed(2) + "B";
    if (abs >= 1000000) return sign + sym + (abs / 1000000).toFixed(2) + "M";
    if (abs >= 1000) return sign + sym + commas((abs / 1000).toFixed(1)) + "K";
    return sign + sym + commas(abs.toFixed(2));
  }

  function fmtNumber(val) {
    if (val === 0 || val === null) return "0";
    const abs = Math.abs(val);
    const sign = val < 0 ? "-" : "";
    if (abs >= 1000000000) return sign + (abs / 1000000000).toFixed(1) + "B";
    if (abs >= 1000000) return sign + (abs / 1000000).toFixed(1) + "M";
    if (abs >= 1000) return sign + commas((abs / 1000).toFixed(1)) + "K";
    return sign + commas(abs.toFixed(0));
  }

  function commas(n) {
    return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  function esc(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }

  return { render, cleanup };
})();
