/**
 * The VC Corner Theme - Theme/palette state, definitions, controls, re-render
 */
const Theme = (() => {
  // ─── Palette Definitions ───
  const PALETTES = {
    dark: {
      gold:     { label: "Gold",     shades: ["#78350F","#92400E","#B45309","#D97706","#F59E0B","#FBBF24","#FCD34D","#FDE68A"] },
      ocean:    { label: "Ocean",    shades: ["#1E3A5F","#1E40AF","#1D4ED8","#2563EB","#3B82F6","#60A5FA","#93C5FD","#BFDBFE"] },
      sage:     { label: "Sage",     shades: ["#064E3B","#065F46","#047857","#059669","#10B981","#34D399","#6EE7B7","#A7F3D0"] },
      lavender: { label: "Lavender", shades: ["#4C1D95","#5B21B6","#6D28D9","#7C3AED","#8B5CF6","#A78BFA","#C4B5FD","#DDD6FE"] },
      copper:   { label: "Copper",   shades: ["#7C2D12","#9A3412","#C2410C","#EA580C","#F97316","#FB923C","#FDBA74","#FED7AA"] },
    },
    light: {
      slate:  { label: "Slate",  shades: ["#0F172A","#1E293B","#334155","#475569","#64748B","#94A3B8","#CBD5E1","#E2E8F0"] },
      indigo: { label: "Indigo", shades: ["#312E81","#3730A3","#4338CA","#4F46E5","#6366F1","#818CF8","#A5B4FC","#C7D2FE"] },
      teal:   { label: "Teal",   shades: ["#134E4A","#115E59","#0F766E","#0D9488","#14B8A6","#2DD4BF","#5EEAD4","#99F6E4"] },
      rose:   { label: "Rose",   shades: ["#881337","#9F1239","#BE123C","#E11D48","#F43F5E","#FB7185","#FDA4AF","#FECDD3"] },
      amber:  { label: "Amber",  shades: ["#78350F","#92400E","#B45309","#D97706","#F59E0B","#FBBF24","#FCD34D","#FDE68A"] },
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
