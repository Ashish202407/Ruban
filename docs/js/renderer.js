/**
 * Ruban Renderer - Dark dashboard DOM builder
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

    // KPI strip
    const kpis = collectKPIs(activeSections, sectionData, currencySymbol);
    if (kpis.length > 0) {
      dashboard.appendChild(buildKPIStrip(kpis));
    }

    // Chart grid
    const grid = document.createElement("div");
    grid.className = "chart-grid";

    for (const sectionId of activeSections) {
      const config = REGISTRY[sectionId];
      if (!config) continue;
      const data = sectionData[sectionId];
      if (!data || data.length === 0) continue;

      const card = buildSectionCard(sectionId, config, data, currencySymbol);
      grid.appendChild(card);
    }

    dashboard.appendChild(grid);
    dashboard.appendChild(buildPrivacyBadge());
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

  function buildSectionCard(sectionId, config, data, currencySymbol) {
    const card = document.createElement("div");
    card.className = "section-card";

    // Title
    const titleEl = document.createElement("div");
    titleEl.className = "section-title";
    titleEl.textContent = config.title;
    card.appendChild(titleEl);

    // Build datasets
    const datasets = [];
    for (const rowLabel of (config.chartRows || [])) {
      const row = data.find(r => r.label === rowLabel);
      if (row) {
        datasets.push({ label: row.label, values: [...row.values] });
      }
    }

    // HTML legend (outside canvas - no overlap with y-axis)
    if (datasets.length > 0 && config.chartType !== "pie") {
      const legend = document.createElement("div");
      legend.className = "chart-legend";
      datasets.forEach((ds, i) => {
        const item = document.createElement("span");
        item.className = "chart-legend-item";
        const color = Charts.PALETTE[i % Charts.PALETTE.length];
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
      <span>Your data never leaves your device</span>
    `;
    return badge;
  }

  function cleanup() {
    chartInstances.forEach(c => Charts.destroy(c));
    chartInstances = [];
  }

  // ─── Formatting ───
  function fmtCurrency(val, sym) {
    if (val === 0 || val === null || val === undefined) return sym + "0";
    const abs = Math.abs(val);
    const sign = val < 0 ? "-" : "";
    if (abs >= 10000000) return sign + sym + (abs / 10000000).toFixed(2) + " Cr";
    if (abs >= 100000) return sign + sym + (abs / 100000).toFixed(2) + " L";
    if (abs >= 1000) return sign + sym + commas((abs / 1000).toFixed(1)) + "K";
    return sign + sym + commas(abs.toFixed(2));
  }

  function fmtNumber(val) {
    if (val === 0 || val === null) return "0";
    const abs = Math.abs(val);
    const sign = val < 0 ? "-" : "";
    if (abs >= 10000000) return sign + (abs / 10000000).toFixed(1) + " Cr";
    if (abs >= 100000) return sign + (abs / 100000).toFixed(1) + "L";
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
