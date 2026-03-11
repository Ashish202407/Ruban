/**
 * Ruban Charts - Chart.js factory, dynamic theme colors, razor-sharp at any zoom
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
   * @param {HTMLCanvasElement} canvas
   * @param {string} type - "line" | "bar" | "stacked_bar" | "pie"
   * @param {string[]} labels
   * @param {Object[]} datasets - [{label, values}]
   * @param {string} currencySymbol
   * @param {string} yFormat - "currency" | "percent" | "percent_whole" | "number" | "ratio"
   */
  function create(canvas, type, labels, datasets, currencySymbol, yFormat) {
    const ctx = canvas.getContext("2d");
    const chartType = type === "stacked_bar" ? "bar" : (type === "pie" ? "doughnut" : type);
    const isStacked = type === "stacked_bar";

    // Dynamic palette and chart theme from Theme module
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

    // Build y-axis tick formatter based on yFormat
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
          grid: { color: ct.gridColor, lineWidth: 0.5 },
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
   * Build a tick formatter for the y-axis
   */
  function buildTickFormatter(yFormat, sym) {
    if (yFormat === "percent") {
      return (value) => {
        return (value * 100).toFixed(0) + "%";
      };
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
    // Default: currency
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

  /**
   * Build a tooltip formatter
   */
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

  function destroy(chart) { if (chart) chart.destroy(); }

  return {
    create, destroy, sharpDPR,
    get PALETTE() { return Theme.getPalette(); },
    get PALETTE_ALPHA() { return Theme.getPalette().map(c => hexToRgba(c, 0.7)); },
  };
})();
