/**
 * Ruban App — Orchestrator
 * File upload → parse → render
 */
(function () {
  Theme.init();

  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const uploadScreen = document.getElementById("upload-screen");
  const dashboardScreen = document.getElementById("dashboard-screen");
  const backBtn = document.getElementById("back-btn");
  const errorMsg = document.getElementById("error-msg");

  // ─── Drag & drop ───
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  // ─── Click to upload ───
  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
  });

  // ─── Back button ───
  backBtn.addEventListener("click", () => {
    Theme.stashParsedData(null);
    Renderer.cleanup();
    dashboardScreen.style.display = "none";
    uploadScreen.style.display = "flex";
    fileInput.value = "";
    errorMsg.style.display = "none";
  });

  // ─── File handler ───
  function handleFile(file) {
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      showError("Please upload an Excel file (.xlsx)");
      return;
    }

    errorMsg.style.display = "none";

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const parsed = Parser.parse(data);

        if (parsed.activeSections.length === 0) {
          showError('No sections marked "Yes" in the Checklist sheet. Check your template.');
          return;
        }

        const hasData = Object.values(parsed.sectionData).some(rows => rows.length > 0);
        if (!hasData) {
          showError("No data found. Fill in the template before uploading.");
          return;
        }

        // Transition
        uploadScreen.style.display = "none";
        dashboardScreen.style.display = "block";

        // Scroll to top
        window.scrollTo(0, 0);

        // Stash parsed data for theme re-renders
        Theme.stashParsedData(parsed);

        // Render
        Renderer.cleanup();
        Renderer.render(parsed);
      } catch (err) {
        console.error("Parse error:", err);
        showError("Error parsing file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function showError(msg) {
    errorMsg.textContent = msg;
    errorMsg.style.display = "block";
  }
})();
