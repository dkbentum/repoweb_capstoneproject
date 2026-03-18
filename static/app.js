const form = document.getElementById("generator-form");
const filesInput = document.getElementById("subject-files");
const fileList = document.getElementById("file-list");
const statusEl = document.getElementById("status");
const generateBtn = document.getElementById("generate-btn");
const resultsPanel = document.getElementById("results");
const compiledLink = document.getElementById("compiled-link");
const zipLink = document.getElementById("zip-link");
const summaryEl = document.getElementById("summary");
const studentLinksEl = document.getElementById("student-links");
const warningsEl = document.getElementById("warnings");
const downloadTemplateBtn = document.getElementById("download-template-btn");

const templateCards = Array.from(document.querySelectorAll(".template-card"));
const templateRadios = Array.from(document.querySelectorAll('input[name="template_id"]'));

const primaryColorInput = document.getElementById("primary-color-input");
const primaryColorHex = document.getElementById("primary-color-hex");
const secondaryColorInput = document.getElementById("secondary-color-input");
const secondaryColorHex = document.getElementById("secondary-color-hex");
const surfaceColorInput = document.getElementById("surface-color-input");
const surfaceColorHex = document.getElementById("surface-color-hex");
const randomPaletteBtn = document.getElementById("random-palette-btn");

const logoInput = document.getElementById("school-logo");
const logoStyleSelect = document.getElementById("logo-style");
const logoOpacityRow = document.getElementById("logo-opacity-row");
const logoOpacityInput = document.getElementById("logo-opacity");
const logoOpacityValue = document.getElementById("logo-opacity-value");
const logoPreviewWrap = document.getElementById("logo-preview-wrap");
const logoPreview = document.getElementById("logo-preview");
const logoMeta = document.getElementById("logo-meta");

function setStatus(message, type = "") {
  statusEl.textContent = message;
  statusEl.classList.remove("ok", "error");
  if (type) {
    statusEl.classList.add(type);
  }
}

function normalizeHexColor(value, fallback = "#1638B7") {
  let hex = String(value || "").trim().toUpperCase();
  if (!hex.startsWith("#")) {
    hex = `#${hex}`;
  }
  if (/^#[0-9A-F]{6}$/.test(hex)) {
    return hex;
  }
  return fallback;
}

function ordinal(value) {
  const n = Number(value || 0);
  if (!n) {
    return "-";
  }
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) {
    return `${n}th`;
  }
  const mod10 = n % 10;
  if (mod10 === 1) return `${n}st`;
  if (mod10 === 2) return `${n}nd`;
  if (mod10 === 3) return `${n}rd`;
  return `${n}th`;
}

function refreshTemplateCards() {
  templateCards.forEach((card) => {
    const input = card.querySelector('input[type="radio"]');
    card.classList.toggle("selected", Boolean(input && input.checked));
  });
}

function bindColorPair(colorInput, hexInput, fallback) {
  const syncFromColor = () => {
    const hex = normalizeHexColor(colorInput.value, fallback);
    colorInput.value = hex;
    hexInput.value = hex;
  };

  const syncFromHex = () => {
    const hex = normalizeHexColor(hexInput.value, colorInput.value || fallback);
    colorInput.value = hex;
    hexInput.value = hex;
  };

  colorInput.addEventListener("input", syncFromColor);
  hexInput.addEventListener("blur", syncFromHex);
  syncFromColor();
}

function hslToHex(h, s, l) {
  const sat = s / 100;
  const light = l / 100;
  const c = (1 - Math.abs(2 * light - 1)) * sat;
  const x = c * (1 - Math.abs((h / 60) % 2 - 1));
  const m = light - c / 2;

  let r = 0;
  let g = 0;
  let b = 0;

  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];

  const toHex = (channel) => Math.round((channel + m) * 255).toString(16).padStart(2, "0").toUpperCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function applyPalette(primary, secondary, surface) {
  const primaryHex = normalizeHexColor(primary, "#1638B7");
  const secondaryHex = normalizeHexColor(secondary, "#0F2F92");
  const surfaceHex = normalizeHexColor(surface, "#E8EEFF");

  primaryColorInput.value = primaryHex;
  primaryColorHex.value = primaryHex;
  secondaryColorInput.value = secondaryHex;
  secondaryColorHex.value = secondaryHex;
  surfaceColorInput.value = surfaceHex;
  surfaceColorHex.value = surfaceHex;

  document.documentElement.style.setProperty("--brand", primaryHex);
}

function randomizePalette() {
  const hue = Math.floor(Math.random() * 360);
  const primary = hslToHex(hue, 70, 42);
  const secondary = hslToHex((hue + 24) % 360, 64, 34);
  const surface = hslToHex((hue + 8) % 360, 58, 93);
  applyPalette(primary, secondary, surface);
}

function refreshLogoStyle() {
  const isWatermark = logoStyleSelect.value === "watermark";
  logoOpacityRow.classList.toggle("hidden", !isWatermark);
}

function refreshLogoOpacity() {
  logoOpacityValue.textContent = `${logoOpacityInput.value}%`;
}

function refreshLogoPreview() {
  const [file] = logoInput.files || [];
  if (!file) {
    logoPreviewWrap.classList.add("hidden");
    logoPreview.removeAttribute("src");
    logoMeta.textContent = "Logo preview";
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    logoPreview.src = String(reader.result);
    logoPreviewWrap.classList.remove("hidden");
    logoMeta.textContent = `${file.name} (${Math.max(1, Math.round(file.size / 1024))} KB)`;
  };
  reader.readAsDataURL(file);
}

function refreshFileList() {
  const files = Array.from(filesInput.files || []);
  fileList.innerHTML = "";

  if (!files.length) {
    const li = document.createElement("li");
    li.textContent = "No files selected yet.";
    fileList.appendChild(li);
    return;
  }

  files.forEach((file) => {
    const li = document.createElement("li");
    const kb = Math.max(1, Math.round(file.size / 1024));
    li.textContent = `${file.name} (${kb} KB)`;
    fileList.appendChild(li);
  });
}

function renderResults(data) {
  resultsPanel.classList.remove("hidden");
  const logoSummary = data.has_logo ? `Logo: ${data.logo_style_label}` : "Logo: none";
  summaryEl.textContent = `Generated ${data.student_count} student report card(s) across ${data.subject_count} subject(s). Weighting: ClassScore ${data.class_weight}% / ExamScore ${data.exam_weight}%. Template: ${data.template_label}. Colors: ${data.primary_color}, ${data.secondary_color}, ${data.surface_color}. ${logoSummary}. Job ID: ${data.job_id}`;

  compiledLink.href = data.compiled_pdf_url;
  zipLink.href = data.zip_url;

  studentLinksEl.innerHTML = "";
  data.students.forEach((student) => {
    const li = document.createElement("li");

    const meta = document.createElement("span");
    const idText = student.student_id ? ` (${student.student_id})` : "";
    const subjectText = Object.entries(student.subject_positions || {})
      .map(([subject, pos]) => `${subject}: ${ordinal(pos)}`)
      .join(" | ");
    meta.textContent = `${student.student_name}${idText} - Overall ${student.overall_position_label} - Avg ${student.overall_average}${subjectText ? ` - ${subjectText}` : ""}`;

    const anchor = document.createElement("a");
    anchor.href = student.pdf_url;
    anchor.target = "_blank";
    anchor.rel = "noopener";
    anchor.textContent = "Download PDF";

    li.appendChild(meta);
    li.appendChild(anchor);
    studentLinksEl.appendChild(li);
  });

  if (Array.isArray(data.warnings) && data.warnings.length > 0) {
    warningsEl.innerHTML = `<strong>Warnings:</strong><br>${data.warnings.map((w) => `- ${w}`).join("<br>")}`;
  } else {
    warningsEl.textContent = "";
  }
}

async function handleSubmit(event) {
  event.preventDefault();

  const files = Array.from(filesInput.files || []);
  if (files.length === 0) {
    setStatus("Select at least one subject file.", "error");
    return;
  }

  const classWeight = Number(form.elements.class_weight.value);
  const examWeight = Number(form.elements.exam_weight.value);

  if (Number.isNaN(classWeight) || Number.isNaN(examWeight)) {
    setStatus("Enter numeric percentages for ClassScore and ExamScore.", "error");
    return;
  }

  if (Math.abs(classWeight + examWeight - 100) > 0.001) {
    setStatus("ClassScore% and ExamScore% must add up to 100.", "error");
    return;
  }

  applyPalette(primaryColorHex.value, secondaryColorHex.value, surfaceColorHex.value);

  const formData = new FormData();
  const rawFormData = new FormData(form);
  for (const [key, value] of rawFormData.entries()) {
    if (key !== "subject_files") {
      formData.append(key, value);
    }
  }
  files.forEach((file) => formData.append("subject_files", file));

  generateBtn.disabled = true;
  setStatus("Processing files, computing positions, and generating styled report cards...", "");

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: formData,
    });

    const payload = await response.json();

    if (!response.ok) {
      const message = payload.error || "Could not generate report cards.";
      setStatus(message, "error");
      return;
    }

    renderResults(payload);
    setStatus("Report cards generated successfully.", "ok");
    resultsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    setStatus(`Request failed: ${error.message}`, "error");
  } finally {
    generateBtn.disabled = false;
  }
}

filesInput.addEventListener("change", refreshFileList);
form.addEventListener("submit", handleSubmit);
downloadTemplateBtn.addEventListener("click", () => {
  window.location.href = "/api/template";
});

templateRadios.forEach((radio) => {
  radio.addEventListener("change", refreshTemplateCards);
});

bindColorPair(primaryColorInput, primaryColorHex, "#1638B7");
bindColorPair(secondaryColorInput, secondaryColorHex, "#0F2F92");
bindColorPair(surfaceColorInput, surfaceColorHex, "#E8EEFF");

randomPaletteBtn.addEventListener("click", randomizePalette);

logoInput.addEventListener("change", refreshLogoPreview);
logoStyleSelect.addEventListener("change", refreshLogoStyle);
logoOpacityInput.addEventListener("input", refreshLogoOpacity);

refreshTemplateCards();
refreshLogoStyle();
refreshLogoOpacity();
refreshLogoPreview();
refreshFileList();
applyPalette(primaryColorHex.value, secondaryColorHex.value, surfaceColorHex.value);