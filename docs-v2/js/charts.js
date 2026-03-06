/**
 * The VC Corner Charts — Chart.js factory + custom chart types
 */
const Charts = (() => {

  function hexToRgba(hex, a) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `rgba(${r},${g},${b},${a})`;
  }

  function sharpDPR() {
    return Math.max(window.devicePixelRatio || 1, 3);
  }

  if (typeof Chart !== 'undefined') {
    Chart.defaults.devicePixelRatio = sharpDPR();
    Chart.defaults.font.family = "'DM Sans', system-ui, sans-serif";
    Chart.defaults.color = "#68686F";
  }

  /**
   * Main dispatch — routes to the correct chart creator
   */
  function create(canvas, type, labels, datasets, currencySymbol, yFormat) {
    switch (type) {
      case "waterfall":
        return createWaterfall(canvas, labels, datasets, currencySymbol, yFormat);
      case "gauge":
        return createGauge(canvas, datasets);
      case "stacked_area":
        return createStackedArea(canvas, labels, datasets, currencySymbol, yFormat);
      case "combo_valuation":
        return createComboValuation(canvas, labels, datasets, currencySymbol);
      case "donut_kpi":
        return createDonutKPI(canvas, labels, datasets, currencySymbol);
      case "radar":
        return createRadar(canvas, labels, datasets);
      default:
        return createStandard(canvas, type, labels, datasets, currencySymbol, yFormat);
    }
  }

  /**
   * Standard chart (line, bar, stacked_bar, pie/doughnut)
   */
  function createStandard(canvas, type, labels, datasets, currencySymbol, yFormat) {
    const ctx = canvas.getContext("2d");
    const chartType = type === "stacked_bar" ? "bar" : (type === "pie" ? "doughnut" : type);
    const isStacked = type === "stacked_bar";

    const palette = Theme.getPalette();
    const paletteAlpha = palette.map(c => hexToRgba(c, 0.7));
    const paletteBg = palette.map(c => hexToRgba(c, 0.15));
    const ct = Theme.getChartTheme();

    const chartDatasets = datasets.map((ds, i) => {
      const color = palette[i % palette.length];
      const alpha = paletteAlpha[i % paletteAlpha.length];
      const bg = paletteBg[i % paletteBg.length];

      if (chartType === "line") {
        return {
          label: ds.label, data: ds.values,
          borderColor: color, backgroundColor: bg,
          borderWidth: 2.5, tension: 0.35, fill: true,
          pointRadius: 5, pointHoverRadius: 7,
          pointBackgroundColor: color, pointBorderColor: ct.pointBorder, pointBorderWidth: 2.5,
        };
      }
      if (chartType === "doughnut") {
        return {
          label: ds.label, data: ds.values,
          backgroundColor: datasets.map((_, j) => paletteAlpha[j % paletteAlpha.length]),
          borderColor: datasets.map((_, j) => palette[j % palette.length]),
          borderWidth: 1.5, hoverOffset: 6,
        };
      }
      return {
        label: ds.label, data: ds.values,
        backgroundColor: alpha, borderColor: color,
        borderWidth: 1.5, borderRadius: 5, borderSkipped: false,
      };
    });

    const tickFormatter = buildTickFormatter(yFormat || "currency", currencySymbol);
    const tooltipFormatter = buildTooltipFormatter(yFormat || "currency", currencySymbol);

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: sharpDPR(),
      layout: { padding: { top: 8, bottom: 4, left: 4, right: 4 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
          borderColor: ct.tooltipBorder, borderWidth: 1,
          titleFont: { family: "'DM Sans'", size: 13, weight: "600" },
          bodyFont: { family: "'JetBrains Mono'", size: 12 },
          padding: 14, cornerRadius: 8, displayColors: true, boxPadding: 5,
          callbacks: {
            label: function(context) {
              const val = context.parsed.y !== undefined ? context.parsed.y : context.parsed;
              return ` ${context.dataset.label}: ${tooltipFormatter(val)}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          border: { color: ct.borderColor },
          ticks: { color: ct.tickColor, font: { family: "'DM Sans'", size: 12, weight: "500" }, padding: 8 }
        },
        y: {
          grid: { color: ct.gridColor, lineWidth: 0.75, borderDash: [4, 4] },
          border: { display: false },
          ticks: {
            color: ct.tickColor,
            font: { family: "'JetBrains Mono'", size: 11 },
            padding: 10, maxTicksLimit: 6,
            callback: function(value) { return tickFormatter(value); }
          }
        }
      },
      interaction: { mode: "index", intersect: false },
      animation: { duration: 700, easing: "easeOutQuart" }
    };

    if (isStacked) {
      options.scales.x.stacked = true;
      options.scales.y.stacked = true;
    }

    if (chartType === "doughnut") {
      delete options.scales;
      options.cutout = "60%";
      options.plugins.legend = {
        display: true, position: "right",
        labels: {
          color: ct.legendLabel,
          font: { family: "'DM Sans'", size: 11, weight: "500" },
          boxWidth: 10, boxHeight: 10, borderRadius: 3, useBorderRadius: true, padding: 10,
        }
      };
    }

    return new Chart(ctx, { type: chartType, data: { labels, datasets: chartDatasets }, options });
  }

  /**
   * Waterfall chart — stacked bar with transparent bases
   */
  function createWaterfall(canvas, labels, datasets, currencySymbol, yFormat) {
    const ctx = canvas.getContext("2d");
    const ct = Theme.getChartTheme();
    const palette = Theme.getPalette();
    const tickFormatter = buildTickFormatter(yFormat || "currency", currencySymbol);
    const tooltipFormatter = buildTooltipFormatter(yFormat || "currency", currencySymbol);

    // Use last year (Year 5) data for waterfall
    const yearIdx = 4;
    const items = [];
    for (const ds of datasets) {
      items.push({ label: ds.label, value: ds.values[yearIdx] || 0 });
    }

    // Build waterfall: Revenue → costs (negative) → subtotals
    const baseValues = [];
    const barValues = [];
    const bgColors = [];
    const borderColors = [];
    const barLabels = [];

    let runningTotal = 0;
    const green = "#34D399";
    const red = "#F87171";
    const accent = palette[0];

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const label = item.label;
      const value = item.value;
      barLabels.push(label);

      // Subtotals (Gross Profit, EBITDA, PAT) — start from 0
      if (label === "Gross Profit" || label === "EBITDA" || label.startsWith("PAT")) {
        baseValues.push(0);
        barValues.push(value);
        bgColors.push(hexToRgba(green, 0.7));
        borderColors.push(green);
        runningTotal = value;
      }
      // Cost items (negative) — float down from running total
      else if (label === "COGS" || label.includes("Expense")) {
        const absVal = Math.abs(value);
        baseValues.push(runningTotal - absVal);
        barValues.push(absVal);
        bgColors.push(hexToRgba(red, 0.5));
        borderColors.push(red);
        runningTotal -= absVal;
      }
      // Revenue — positive from 0
      else {
        baseValues.push(0);
        barValues.push(value);
        bgColors.push(hexToRgba(accent, 0.7));
        borderColors.push(accent);
        runningTotal = value;
      }
    }

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: sharpDPR(),
      layout: { padding: { top: 24, bottom: 4, left: 4, right: 4 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
          borderColor: ct.tooltipBorder, borderWidth: 1,
          padding: 14, cornerRadius: 8,
          callbacks: {
            label: (context) => {
              if (context.datasetIndex === 0) return null; // hide base
              return ` ${tooltipFormatter(context.parsed.y)}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          border: { color: ct.borderColor },
          ticks: { color: ct.tickColor, font: { family: "'DM Sans'", size: 10, weight: "500" }, maxRotation: 45 }
        },
        y: {
          grid: { color: ct.gridColor, lineWidth: 0.75, borderDash: [4, 4] },
          border: { display: false },
          stacked: true,
          ticks: {
            color: ct.tickColor,
            font: { family: "'JetBrains Mono'", size: 11 },
            padding: 10, maxTicksLimit: 6,
            callback: (value) => tickFormatter(value)
          }
        }
      },
      animation: { duration: 700, easing: "easeOutQuart" }
    };

    // Value labels on top of bars
    const valueLabelsPlugin = {
      id: "waterfallLabels",
      afterDraw(chart) {
        const { ctx: c, chartArea } = chart;
        const meta = chart.getDatasetMeta(1);
        c.save();
        c.font = "600 10px 'JetBrains Mono'";
        c.textAlign = "center";
        c.fillStyle = ct.tickColor;
        meta.data.forEach((bar, i) => {
          const val = barValues[i];
          const formatted = tooltipFormatter(val);
          c.fillText(formatted, bar.x, bar.y - 6);
        });
        c.restore();
      }
    };

    return new Chart(ctx, {
      type: "bar",
      data: {
        labels: barLabels,
        datasets: [
          { label: "Base", data: baseValues, backgroundColor: "transparent", borderWidth: 0, borderSkipped: false },
          { label: "Value", data: barValues, backgroundColor: bgColors, borderColor: borderColors, borderWidth: 1.5, borderRadius: 4, borderSkipped: false }
        ]
      },
      options: { ...options, scales: { ...options.scales, x: { ...options.scales.x, stacked: true } } },
      plugins: [valueLabelsPlugin]
    });
  }

  /**
   * Gauge chart — semicircular arc with needle (Canvas 2D)
   */
  function createGauge(canvas, datasets) {
    const ctx = canvas.getContext("2d");
    const ct = Theme.getChartTheme();
    const dpr = sharpDPR();

    // Get Year 5 values
    const values = datasets.map(ds => {
      const v = ds.values[4] || ds.values[ds.values.length - 1] || 0;
      return { label: ds.label, value: typeof v === "number" ? v : parseFloat(v) || 0 };
    });

    // Render function
    function drawGauge() {
      const rect = canvas.getBoundingClientRect();
      const w = rect.width;
      const h = rect.height;
      canvas.width = w * dpr;
      canvas.height = h * dpr;
      ctx.scale(dpr, dpr);

      ctx.clearRect(0, 0, w, h);

      const gaugeCount = Math.min(values.length, 2);
      const gaugeWidth = w / gaugeCount;

      for (let g = 0; g < gaugeCount; g++) {
        const item = values[g];
        let score = item.value;
        // Normalize: for percentages (Rule of 40), multiply by 100
        if (Math.abs(score) < 2) score = score * 100;
        score = Math.max(-20, Math.min(100, score));

        const cx = gaugeWidth * g + gaugeWidth / 2;
        const cy = h * 0.65;
        const radius = Math.min(gaugeWidth * 0.38, h * 0.45);
        const lineWidth = radius * 0.18;

        // Draw arc segments
        const segments = [
          { end: 20, color: "#F87171" },
          { end: 40, color: "#FBBF24" },
          { end: 60, color: "#34D399" },
          { end: 100, color: "#059669" }
        ];

        let startAngle = Math.PI;
        segments.forEach(seg => {
          const prevEnd = segments.indexOf(seg) === 0 ? -20 : segments[segments.indexOf(seg) - 1].end;
          const segStart = Math.PI + ((prevEnd + 20) / 120) * Math.PI;
          const segEnd = Math.PI + ((seg.end + 20) / 120) * Math.PI;
          ctx.beginPath();
          ctx.arc(cx, cy, radius, segStart, segEnd, false);
          ctx.lineWidth = lineWidth;
          ctx.strokeStyle = hexToRgba(seg.color, 0.3);
          ctx.lineCap = "butt";
          ctx.stroke();
        });

        // Active arc
        const normalizedScore = (score + 20) / 120;
        const activeEnd = Math.PI + normalizedScore * Math.PI;
        const activeColor = score < 20 ? "#F87171" : score < 40 ? "#FBBF24" : score < 60 ? "#34D399" : "#059669";
        ctx.beginPath();
        ctx.arc(cx, cy, radius, Math.PI, activeEnd, false);
        ctx.lineWidth = lineWidth;
        ctx.strokeStyle = activeColor;
        ctx.lineCap = "round";
        ctx.stroke();

        // Needle
        const needleAngle = Math.PI + normalizedScore * Math.PI;
        const needleLen = radius * 0.7;
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.lineTo(cx + needleLen * Math.cos(needleAngle), cy + needleLen * Math.sin(needleAngle));
        ctx.lineWidth = 2;
        ctx.strokeStyle = ct.tickColor;
        ctx.stroke();

        // Center dot
        ctx.beginPath();
        ctx.arc(cx, cy, 4, 0, Math.PI * 2);
        ctx.fillStyle = ct.tickColor;
        ctx.fill();

        // Score text
        ctx.textAlign = "center";
        ctx.font = "700 20px 'JetBrains Mono'";
        ctx.fillStyle = activeColor;
        ctx.fillText(score.toFixed(0), cx, cy + radius * 0.45);

        // Label
        ctx.font = "500 11px 'DM Sans'";
        ctx.fillStyle = ct.tickColor;
        ctx.fillText(item.label, cx, cy + radius * 0.68);
      }
    }

    drawGauge();

    // Return wrapper with destroy() for cleanup compatibility
    return {
      destroy() { ctx.clearRect(0, 0, canvas.width, canvas.height); },
      update() { drawGauge(); }
    };
  }

  /**
   * Stacked Area chart — line chart with fill and stacking
   */
  function createStackedArea(canvas, labels, datasets, currencySymbol, yFormat) {
    const ctx = canvas.getContext("2d");
    const palette = Theme.getPalette();
    const ct = Theme.getChartTheme();
    const tickFormatter = buildTickFormatter(yFormat || "currency", currencySymbol);
    const tooltipFormatter = buildTooltipFormatter(yFormat || "currency", currencySymbol);

    const chartDatasets = datasets.map((ds, i) => {
      const color = palette[i % palette.length];
      return {
        label: ds.label, data: ds.values,
        borderColor: color,
        backgroundColor: hexToRgba(color, 0.25),
        borderWidth: 2,
        tension: 0.4,
        fill: true,
        pointRadius: 3,
        pointBackgroundColor: color,
        pointBorderColor: ct.pointBorder,
        pointBorderWidth: 2,
      };
    });

    return new Chart(ctx, {
      type: "line",
      data: { labels, datasets: chartDatasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        devicePixelRatio: sharpDPR(),
        layout: { padding: { top: 8, bottom: 4, left: 4, right: 4 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
            borderColor: ct.tooltipBorder, borderWidth: 1,
            padding: 14, cornerRadius: 8, displayColors: true, boxPadding: 5,
            callbacks: {
              label: (context) => ` ${context.dataset.label}: ${tooltipFormatter(context.parsed.y)}`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            border: { color: ct.borderColor },
            ticks: { color: ct.tickColor, font: { family: "'DM Sans'", size: 12, weight: "500" }, padding: 8 }
          },
          y: {
            stacked: true,
            grid: { color: ct.gridColor, lineWidth: 0.75, borderDash: [4, 4] },
            border: { display: false },
            ticks: {
              color: ct.tickColor,
              font: { family: "'JetBrains Mono'", size: 11 },
              padding: 10, maxTicksLimit: 6,
              callback: (value) => tickFormatter(value)
            }
          }
        },
        interaction: { mode: "index", intersect: false },
        animation: { duration: 700, easing: "easeOutQuart" }
      }
    });
  }

  /**
   * Combo Valuation — bar + stepped line with dual Y axes
   */
  function createComboValuation(canvas, labels, datasets, currencySymbol) {
    const ctx = canvas.getContext("2d");
    const palette = Theme.getPalette();
    const ct = Theme.getChartTheme();
    const tickFormatter = buildTickFormatter("currency", currencySymbol);
    const tooltipFormatter = buildTooltipFormatter("currency", currencySymbol);

    const barDs = datasets[0] || { label: "Amount Raised", values: [] };
    const lineDs = datasets[1] || { label: "Post-Money Valuation", values: [] };

    return new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: barDs.label, data: barDs.values,
            backgroundColor: hexToRgba(palette[0], 0.7), borderColor: palette[0],
            borderWidth: 1.5, borderRadius: 5, yAxisID: "y"
          },
          {
            label: lineDs.label, data: lineDs.values,
            type: "line", borderColor: palette[1], backgroundColor: "transparent",
            borderWidth: 2.5, stepped: true, pointRadius: 5,
            pointBackgroundColor: palette[1], pointBorderColor: ct.pointBorder, pointBorderWidth: 2.5,
            yAxisID: "y1"
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false, devicePixelRatio: sharpDPR(),
        layout: { padding: { top: 8, bottom: 4, left: 4, right: 4 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
            borderColor: ct.tooltipBorder, borderWidth: 1, padding: 14, cornerRadius: 8,
            callbacks: { label: (ctx2) => ` ${ctx2.dataset.label}: ${tooltipFormatter(ctx2.parsed.y)}` }
          }
        },
        scales: {
          x: { grid: { display: false }, border: { color: ct.borderColor }, ticks: { color: ct.tickColor, font: { family: "'DM Sans'", size: 12 } } },
          y: {
            position: "left", grid: { color: ct.gridColor, lineWidth: 0.75, borderDash: [4, 4] }, border: { display: false },
            ticks: { color: ct.tickColor, font: { family: "'JetBrains Mono'", size: 11 }, callback: (v) => tickFormatter(v) }
          },
          y1: {
            position: "right", grid: { display: false }, border: { display: false },
            ticks: { color: ct.tickColor, font: { family: "'JetBrains Mono'", size: 11 }, callback: (v) => tickFormatter(v) }
          }
        },
        animation: { duration: 700, easing: "easeOutQuart" }
      }
    });
  }

  /**
   * Donut with center KPI text — Year 5 snapshot
   */
  function createDonutKPI(canvas, labels, datasets, currencySymbol) {
    const ctx = canvas.getContext("2d");
    const palette = Theme.getPalette();
    const ct = Theme.getChartTheme();

    // Use Year 5 values
    const year5Values = datasets.map(ds => ds.values[4] || 0);
    const total = year5Values.reduce((a, b) => a + b, 0);

    const centerTextPlugin = {
      id: "centerText",
      afterDraw(chart) {
        const { ctx: c, chartArea } = chart;
        const cx = (chartArea.left + chartArea.right) / 2;
        const cy = (chartArea.top + chartArea.bottom) / 2;
        c.save();
        c.textAlign = "center";
        c.font = "700 18px 'JetBrains Mono'";
        c.fillStyle = ct.tooltipTitle;
        const formatted = buildTooltipFormatter("currency", currencySymbol)(total);
        c.fillText(formatted, cx, cy - 4);
        c.font = "500 10px 'DM Sans'";
        c.fillStyle = ct.tickColor;
        c.fillText("Year 5 Total", cx, cy + 14);
        c.restore();
      }
    };

    return new Chart(ctx, {
      type: "doughnut",
      data: {
        labels: datasets.map(ds => ds.label),
        datasets: [{
          data: year5Values,
          backgroundColor: datasets.map((_, i) => hexToRgba(palette[i % palette.length], 0.7)),
          borderColor: datasets.map((_, i) => palette[i % palette.length]),
          borderWidth: 1.5, hoverOffset: 6
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, devicePixelRatio: sharpDPR(),
        cutout: "65%",
        plugins: {
          legend: {
            display: true, position: "right",
            labels: { color: ct.legendLabel, font: { family: "'DM Sans'", size: 11, weight: "500" }, boxWidth: 10, boxHeight: 10, borderRadius: 3, useBorderRadius: true, padding: 8 }
          },
          tooltip: {
            backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
            borderColor: ct.tooltipBorder, borderWidth: 1, padding: 14, cornerRadius: 8
          }
        },
        animation: { duration: 700, easing: "easeOutQuart" }
      },
      plugins: [centerTextPlugin]
    });
  }

  /**
   * Radar chart — spider chart with benchmark overlay
   */
  function createRadar(canvas, labels, datasets) {
    const ctx = canvas.getContext("2d");
    const palette = Theme.getPalette();
    const ct = Theme.getChartTheme();

    const chartDatasets = [];

    // Actual values dataset
    if (datasets[0]) {
      chartDatasets.push({
        label: "Score",
        data: datasets[0].values,
        backgroundColor: hexToRgba(palette[0], 0.4),
        borderColor: palette[0],
        borderWidth: 2.5,
        pointBackgroundColor: palette[0],
        pointBorderColor: ct.pointBorder,
        pointBorderWidth: 2,
        pointRadius: 5,
      });
    }

    // Benchmark dataset (if provided)
    if (datasets[1]) {
      chartDatasets.push({
        label: "Benchmark",
        data: datasets[1].values,
        backgroundColor: "transparent",
        borderColor: ct.tickColor,
        borderWidth: 1.5,
        borderDash: [6, 4],
        pointRadius: 0,
      });
    }

    return new Chart(ctx, {
      type: "radar",
      data: { labels, datasets: chartDatasets },
      options: {
        responsive: true, maintainAspectRatio: false, devicePixelRatio: sharpDPR(),
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: ct.tooltipBg, titleColor: ct.tooltipTitle, bodyColor: ct.tooltipBody,
            borderColor: ct.tooltipBorder, borderWidth: 1, padding: 14, cornerRadius: 8
          }
        },
        scales: {
          r: {
            grid: { color: ct.gridColor, lineWidth: 0.75 },
            angleLines: { color: ct.gridColor },
            pointLabels: { color: ct.tickColor, font: { family: "'DM Sans'", size: 11, weight: "500" } },
            ticks: { display: false },
            suggestedMin: 0, suggestedMax: 100
          }
        },
        animation: { duration: 700, easing: "easeOutQuart" }
      }
    });
  }

  // ─── Formatters ───

  function buildTickFormatter(yFormat, sym) {
    if (yFormat === "percent") {
      return (value) => (value * 100).toFixed(0) + "%";
    }
    if (yFormat === "percent_whole") {
      return (value) => value.toFixed(0) + "%";
    }
    if (yFormat === "number") {
      return (value) => {
        const abs = Math.abs(value);
        const sign = value < 0 ? "-" : "";
        if (abs >= 1000000) return sign + (abs / 1000000).toFixed(1) + "M";
        if (abs >= 1000) return sign + (abs / 1000).toFixed(1) + "K";
        return sign + abs.toFixed(0);
      };
    }
    if (yFormat === "ratio") {
      return (value) => value.toFixed(1) + "x";
    }
    return (value) => {
      const abs = Math.abs(value);
      const sign = value < 0 ? "-" : "";
      if (abs >= 10000000) return sign + sym + (abs / 10000000).toFixed(1) + "Cr";
      if (abs >= 100000) return sign + sym + (abs / 100000).toFixed(1) + "L";
      if (abs >= 1000) return sign + sym + (abs / 1000).toFixed(1) + "K";
      if (abs === 0) return sym + "0";
      return sign + sym + abs.toFixed(abs < 10 ? 1 : 0);
    };
  }

  function buildTooltipFormatter(yFormat, sym) {
    if (yFormat === "percent") {
      return (value) => (value * 100).toFixed(1) + "%";
    }
    if (yFormat === "percent_whole") {
      return (value) => value.toFixed(1) + "%";
    }
    if (yFormat === "number") {
      return (value) => {
        const abs = Math.abs(value);
        const sign = value < 0 ? "-" : "";
        if (abs >= 1000000) return sign + (abs / 1000000).toFixed(2) + "M";
        if (abs >= 1000) return sign + (abs / 1000).toFixed(1) + "K";
        return sign + abs.toFixed(0);
      };
    }
    if (yFormat === "ratio") {
      return (value) => value.toFixed(2) + "x";
    }
    return (value) => {
      const abs = Math.abs(value);
      const sign = value < 0 ? "-" : "";
      if (abs >= 10000000) return sign + sym + (abs / 10000000).toFixed(2) + " Cr";
      if (abs >= 100000) return sign + sym + (abs / 100000).toFixed(2) + " L";
      if (abs >= 1000) return sign + sym + (abs / 1000).toFixed(1) + "K";
      if (abs === 0) return sym + "0";
      return sign + sym + abs.toFixed(2);
    };
  }

  function destroy(chart) {
    if (chart) {
      if (typeof chart.destroy === "function") chart.destroy();
    }
  }

  return {
    create, destroy, sharpDPR, hexToRgba,
    get PALETTE() { return Theme.getPalette(); },
    get PALETTE_ALPHA() { return Theme.getPalette().map(c => hexToRgba(c, 0.7)); },
  };
})();
