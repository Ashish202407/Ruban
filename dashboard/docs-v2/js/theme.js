/**
 * The VC Corner Theme — Theme/palette state, definitions, controls, re-render
 */
const Theme = (() => {
  // ─── Palette Definitions ───
  const PALETTES = {
    dark: {
      gold:     { label: "Gold",     shades: ["#C8A96E","#5B8DEF","#E06B8A","#6EC8A9","#C490D4","#E8A44C","#52BFD9","#D4716E"] },
      ocean:    { label: "Ocean",    shades: ["#5B8DEF","#E8A44C","#6EC8A9","#C490D4","#E06B8A","#52BFD9","#C8A96E","#8BBF5E"] },
      sage:     { label: "Sage",     shades: ["#6EC8A9","#D4896E","#7B8DEF","#E0C870","#C490D4","#52BFD9","#E06B8A","#8BBF5E"] },
      lavender: { label: "Lavender", shades: ["#A78BDB","#5BC8A9","#E8A44C","#5B8DEF","#E06B8A","#8BBF5E","#52BFD9","#D4896E"] },
      copper:   { label: "Copper",   shades: ["#D4896E","#5B8DEF","#6EC8A9","#C490D4","#E0C870","#52BFD9","#E06B8A","#8BBF5E"] },
    },
    light: {
      slate:  { label: "Slate",  shades: ["#64748B","#2563EB","#D946A8","#0D9488","#7C3AED","#D97706","#0891B2","#BE123C"] },
      indigo: { label: "Indigo", shades: ["#4F46E5","#D946A8","#0D9488","#D97706","#64748B","#BE123C","#0891B2","#15803D"] },
      teal:   { label: "Teal",   shades: ["#0D9488","#D97706","#4F46E5","#BE123C","#64748B","#7C3AED","#0891B2","#D946A8"] },
      rose:   { label: "Rose",   shades: ["#BE123C","#2563EB","#0D9488","#D97706","#7C3AED","#64748B","#D946A8","#0891B2"] },
      amber:  { label: "Amber",  shades: ["#D97706","#2563EB","#0D9488","#BE123C","#7C3AED","#64748B","#D946A8","#0891B2"] },
    }
  };

  const THEME_DEFAULTS = { dark: "gold", light: "slate" };

  // ─── Chart theme colors per base theme ───
  const CHART_THEMES = {
    dark: {
      gridColor: "rgba(42,42,48,0.5)",
      borderColor: "#2A2A30",
      tickColor: "#68686F",
      tooltipBg: "#1C1C20",
      tooltipTitle: "#E8E8EC",
      tooltipBody: "#9898A0",
      tooltipBorder: "#2A2A30",
      pointBorder: "#18181B",
      legendLabel: "#9898A0",
      defaultColor: "#68686F",
    },
    light: {
      gridColor: "rgba(0,0,0,0.06)",
      borderColor: "#E2E2E8",
      tickColor: "#6B7280",
      tooltipBg: "#FFFFFF",
      tooltipTitle: "#1F2937",
      tooltipBody: "#6B7280",
      tooltipBorder: "#E5E7EB",
      pointBorder: "#FFFFFF",
      legendLabel: "#6B7280",
      defaultColor: "#6B7280",
    }
  };

  // ─── State ───
  let currentTheme = "light";
  let currentPaletteId = "slate";
  let parsedData = null;

  // ─── Init ───
  function init() {
    // Restore from localStorage
    const storedTheme = localStorage.getItem("ruban-theme");
    const storedPalette = localStorage.getItem("ruban-palette");

    if (storedTheme === "dark" || storedTheme === "light") {
      currentTheme = storedTheme;
    }

    if (storedPalette && PALETTES[currentTheme][storedPalette]) {
      currentPaletteId = storedPalette;
    } else {
      currentPaletteId = THEME_DEFAULTS[currentTheme];
    }

    applyTheme();
    buildControls();
  }

  // ─── Apply theme to DOM ───
  function applyTheme() {
    document.documentElement.setAttribute("data-theme", currentTheme);
    localStorage.setItem("ruban-theme", currentTheme);
    localStorage.setItem("ruban-palette", currentPaletteId);

    if (typeof Chart !== "undefined") {
      Chart.defaults.color = CHART_THEMES[currentTheme].defaultColor;
    }
  }

  // ─── Build nav controls ───
  function buildControls() {
    const container = document.getElementById("theme-controls");
    if (!container) return;
    container.innerHTML = "";

    // Theme toggle pill
    const toggle = document.createElement("div");
    toggle.className = "theme-toggle";

    const sunBtn = document.createElement("button");
    sunBtn.className = "theme-toggle-btn" + (currentTheme === "light" ? " active" : "");
    sunBtn.setAttribute("aria-label", "Light theme");
    sunBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>`;
    sunBtn.addEventListener("click", () => switchTheme("light"));

    const moonBtn = document.createElement("button");
    moonBtn.className = "theme-toggle-btn" + (currentTheme === "dark" ? " active" : "");
    moonBtn.setAttribute("aria-label", "Dark theme");
    moonBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`;
    moonBtn.addEventListener("click", () => switchTheme("dark"));

    toggle.appendChild(sunBtn);
    toggle.appendChild(moonBtn);
    container.appendChild(toggle);

    // Palette picker (custom dropdown)
    const picker = document.createElement("div");
    picker.className = "palette-picker";

    const pickerBtn = document.createElement("button");
    pickerBtn.className = "palette-picker-btn";
    pickerBtn.setAttribute("aria-label", "Choose color palette");
    updatePickerBtn(pickerBtn);

    const dropdown = document.createElement("div");
    dropdown.className = "palette-dropdown";
    buildDropdownOptions(dropdown);

    pickerBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = dropdown.classList.contains("open");
      dropdown.classList.toggle("open", !isOpen);
      pickerBtn.classList.toggle("open", !isOpen);
    });

    document.addEventListener("click", () => {
      dropdown.classList.remove("open");
      pickerBtn.classList.remove("open");
    });

    picker.appendChild(pickerBtn);
    picker.appendChild(dropdown);
    container.appendChild(picker);
  }

  function updatePickerBtn(btn) {
    const palette = PALETTES[currentTheme][currentPaletteId];
    btn.innerHTML = `
      <span class="palette-picker-swatches">
        ${palette.shades.slice(0, 4).map(c => `<span class="palette-swatch-dot" style="background:${c}"></span>`).join("")}
      </span>
      <span class="palette-picker-label">${palette.label}</span>
      <svg width="10" height="10" viewBox="0 0 16 16" fill="none"><path d="M4 6l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
    `;
  }

  function buildDropdownOptions(dropdown) {
    dropdown.innerHTML = "";
    const palettes = PALETTES[currentTheme];
    for (const [id, pal] of Object.entries(palettes)) {
      const option = document.createElement("button");
      option.className = "palette-option" + (id === currentPaletteId ? " active" : "");
      option.innerHTML = `
        <span class="palette-option-swatches">
          ${pal.shades.slice(0, 5).map(c => `<span class="palette-swatch-dot" style="background:${c}"></span>`).join("")}
        </span>
        <span class="palette-option-label">${pal.label}</span>
      `;
      option.addEventListener("click", (e) => {
        e.stopPropagation();
        selectPalette(id);
      });
      dropdown.appendChild(option);
    }
  }

  // ─── Theme switching ───
  function switchTheme(theme) {
    if (theme === currentTheme) return;
    currentTheme = theme;

    // If current palette doesn't exist in new theme, fall back to default
    if (!PALETTES[currentTheme][currentPaletteId]) {
      currentPaletteId = THEME_DEFAULTS[currentTheme];
    }

    applyTheme();
    buildControls();
    rerender();
  }

  function selectPalette(id) {
    if (id === currentPaletteId) return;
    currentPaletteId = id;
    localStorage.setItem("ruban-palette", currentPaletteId);
    buildControls();
    rerender();
  }

  // ─── Re-render dashboard ───
  function rerender() {
    if (!parsedData) return;

    const dashboard = document.getElementById("dashboard");
    if (dashboard) {
      dashboard.classList.add("no-animate");
    }

    Renderer.cleanup();
    Renderer.render(parsedData);

    // Remove no-animate after a frame so future navigations still animate
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (dashboard) dashboard.classList.remove("no-animate");
      });
    });
  }

  // ─── Public API ───
  function getPalette() {
    return PALETTES[currentTheme][currentPaletteId].shades;
  }

  function getChartTheme() {
    return CHART_THEMES[currentTheme];
  }

  function stashParsedData(data) {
    parsedData = data;
  }

  return {
    init,
    getPalette,
    getChartTheme,
    stashParsedData,
    rerender,
    get currentTheme() { return currentTheme; },
    get currentPaletteId() { return currentPaletteId; },
  };
})();
