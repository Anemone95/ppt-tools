const state = window.PPT_TOOLS_STATE || {};
const LINE_NUMBER_COLOR = "8996A8";
const DEFAULT_HIGHLIGHT_COLOR = "FFF1A8";

const languageSelect = document.querySelector("#languageSelect");
const codeInput = document.querySelector("#codeInput");
const codeOutput = document.querySelector("#codeOutput");
const lineNumbers = document.querySelector("#lineNumbers");
const lineHighlights = document.querySelector("#lineHighlights");
const lineCount = document.querySelector("#lineCount");
const previewLanguage = document.querySelector("#previewLanguage");
const previewShell = document.querySelector("#previewShell");
const lineNumberToggle = document.querySelector("#lineNumberToggle");
const fontSizeSelect = document.querySelector("#fontSizeSelect");
const numberGapSelect = document.querySelector("#numberGapSelect");
const formatButton = document.querySelector("#formatButton");
const applyButton = document.querySelector("#applyButton");
const cancelButton = document.querySelector("#cancelButton");
const toast = document.querySelector("#toast");

let detectedLanguage = "javascript";
let toastTimer;
const highlights = new Map(); // line number -> RRGGBB

const formatterConfigs = {
  javascript: { parser: "babel", plugins: ["babel", "estree"] },
  typescript: { parser: "typescript", plugins: ["typescript", "estree"] },
  json: { parser: "json", plugins: ["babel", "estree"] },
  xml: { parser: "html", plugins: ["html"] },
  css: { parser: "css", plugins: ["postcss"] },
  markdown: { parser: "markdown", plugins: ["markdown"] },
  java: { parser: "java", plugins: ["java"] },
};

const languageNames = {
  auto: "AUTO",
  javascript: "JAVASCRIPT",
  typescript: "TYPESCRIPT",
  python: "PYTHON",
  java: "JAVA",
  cpp: "C / C++",
  csharp: "C#",
  go: "GO",
  rust: "RUST",
  json: "JSON",
  bash: "SHELL",
  shell: "SHELL",
  sql: "SQL",
  xml: "HTML / XML",
  css: "CSS",
  markdown: "MARKDOWN",
  plaintext: "PLAIN TEXT",
};

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function highlightCode(source, language) {
  try {
    if (language === "auto") {
      const result = window.hljs.highlightAuto(source);
      detectedLanguage = result.language || "plaintext";
      return result.value;
    }

    detectedLanguage = language;
    return window.hljs.highlight(source, { language }).value;
  } catch {
    detectedLanguage = "plaintext";
    return escapeHtml(source);
  }
}

function renderLineNumbers(totalLines) {
  lineNumbers.replaceChildren(...Array.from({ length: totalLines }, (_, index) => {
    const line = index + 1;
    const row = document.createElement("span");
    row.className = "line-number";
    row.dataset.line = line;
    row.textContent = line;
    row.title = "Click to highlight this line";

    const color = highlights.get(line);
    if (color) {
      const swatch = document.createElement("input");
      swatch.type = "color";
      swatch.className = "highlight-swatch";
      swatch.value = `#${color.toLowerCase()}`;
      swatch.title = "Change highlight color";
      row.classList.add("is-highlighted");
      row.style.setProperty("--highlight", `#${color}`);
      row.prepend(swatch);
    }
    return row;
  }));
}

function renderHighlightBands() {
  lineHighlights.replaceChildren(...[...highlights].map(([line, color]) => {
    const band = document.createElement("span");
    band.className = "line-highlight";
    band.style.setProperty("--line-index", line - 1);
    band.style.setProperty("--highlight", `#${color}`);
    return band;
  }));
}

function render() {
  const source = codeInput.value;
  const sourceLines = source.split("\n");
  const totalLines = sourceLines.length;
  const maxLineLength = sourceLines.reduce(
    (max, line) => Math.max(max, line.replaceAll("\t", "    ").length), 0);

  highlights.forEach((_, line) => {
    if (line > totalLines) highlights.delete(line);
  });
  codeOutput.innerHTML = highlightCode(source, languageSelect.value);
  lineHighlights.style.setProperty("--max-line-length", maxLineLength);
  renderLineNumbers(totalLines);
  renderHighlightBands();
  lineCount.textContent = `${totalLines} ${totalLines === 1 ? "line" : "lines"}`;
  previewLanguage.textContent = languageNames[detectedLanguage] || detectedLanguage.toUpperCase();
}

function updateOutputOptions() {
  const fontSize = Number(fontSizeSelect.value);
  previewShell.classList.toggle("hide-line-numbers", !lineNumberToggle.checked);
  previewShell.style.setProperty("--preview-code-size", `${Math.max(11, fontSize * 0.72)}px`);
  previewShell.style.setProperty("--line-number-gap", `${numberGapSelect.value}ch`);
  numberGapSelect.disabled = !lineNumberToggle.checked;
}

async function formatCode() {
  const language = languageSelect.value === "auto" ? detectedLanguage : languageSelect.value;
  const config = formatterConfigs[language];

  if (!config) {
    showToast("Formatting supports JavaScript, TypeScript, Java, JSON, HTML, CSS, and Markdown", true);
    return;
  }
  if (!codeInput.value.trim()) {
    showToast("Enter some code before formatting", true);
    return;
  }

  try {
    const formatted = await window.prettier.format(codeInput.value, {
      parser: config.parser,
      plugins: config.plugins.map((name) => window.prettierPlugins[name]),
      printWidth: 80,
      tabWidth: 2,
      useTabs: false,
    });
    codeInput.value = formatted.replace(/\n$/, "");
    render();
    showToast("Code formatted");
  } catch {
    showToast("Could not format this code — check its syntax", true);
  }
}

function cssColorToHex(color) {
  return (color.match(/\d+/g) || [])
    .slice(0, 3)
    .map((part) => Number(part).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function collectStyledLines() {
  const lines = [[]];

  function walk(node, inheritedStyle) {
    if (node.nodeType === Node.TEXT_NODE) {
      node.textContent.replaceAll("\t", "    ").split("\n").forEach((text, index, parts) => {
        if (text) {
          lines.at(-1).push({
            text,
            color: cssColorToHex(inheritedStyle.color),
            bold: Number(inheritedStyle.fontWeight) >= 600,
            italic: inheritedStyle.fontStyle === "italic",
          });
        }
        if (index < parts.length - 1) lines.push([]);
      });
      return;
    }

    const style = node === codeOutput ? inheritedStyle : getComputedStyle(node);
    node.childNodes.forEach((child) => walk(child, style));
  }

  walk(codeOutput, getComputedStyle(codeOutput));
  return lines;
}

// Builds the slide text (paragraphs joined by \r) and 1-based character runs for PowerPoint.
function buildSlideCode() {
  const lines = collectStyledLines();
  const numberWidth = String(lines.length).length;
  const numberGap = Number(numberGapSelect.value);
  const defaultColor = cssColorToHex(getComputedStyle(codeOutput).color);
  const prefixLength = lineNumberToggle.checked ? numberWidth + numberGap : 0;
  const maxLength = prefixLength + lines.reduce((max, segments) =>
    Math.max(max, segments.reduce((sum, segment) => sum + segment.text.length, 0)), 0);
  const runs = [];
  const highlightRuns = [];
  let text = "";

  function append(value, color, bold = false, italic = false) {
    const start = text.length + 1;
    const last = runs.at(-1);
    text += value;
    if (color === defaultColor && !bold && !italic) return;
    if (last && last.start + last.length === start && last.color === color &&
        last.bold === bold && last.italic === italic) {
      last.length += value.length;
    } else {
      runs.push({ start, length: value.length, color, bold, italic });
    }
  }

  lines.forEach((segments, index) => {
    if (index > 0) text += "\r";
    const lineStart = text.length + 1;
    if (lineNumberToggle.checked) {
      append(`${String(index + 1).padStart(numberWidth, " ")}${" ".repeat(numberGap)}`, LINE_NUMBER_COLOR);
    }
    segments.forEach((segment) => append(segment.text, segment.color, segment.bold, segment.italic));

    const color = highlights.get(index + 1);
    if (color) {
      // Highlighted lines are padded to the longest line so every band has the same width.
      // The blank braille cell (U+2800) ends the line, so PowerPoint highlights the padding too.
      append(`${" ".repeat(maxLength - (text.length + 1 - lineStart))}\u2800`, defaultColor);
      highlightRuns.push({ start: lineStart, length: text.length + 1 - lineStart, color });
    }
  });

  return { text, runs, defaultColor, highlights: highlightRuns };
}

function serializeHighlights() {
  return [...highlights]
    .sort((a, b) => a[0] - b[0])
    .map(([line, color]) => `${line}:${color}`)
    .join(",");
}

function postToPowerPoint(message) {
  window.webkit.messageHandlers.pptTools.postMessage(JSON.stringify(message));
}

function applyToSlide() {
  if (!codeInput.value.trim()) {
    showToast("Enter some code first", true);
    return;
  }

  codeInput.value = codeInput.value.replace(/\s+$/, "");
  render();
  postToPowerPoint({
    action: "apply",
    language: languageSelect.value,
    fontSize: Number(fontSizeSelect.value),
    lineNumbers: lineNumberToggle.checked,
    numberGap: Number(numberGapSelect.value),
    source: codeInput.value,
    lineHighlights: serializeHighlights(),
    ...buildSlideCode(),
  });
}

function cancel() {
  postToPowerPoint({ action: "cancel" });
}

function showToast(message, isError = false) {
  clearTimeout(toastTimer);
  toast.querySelector("span").textContent = message;
  toast.classList.toggle("is-error", isError);
  toast.classList.add("is-visible");
  toastTimer = window.setTimeout(() => toast.classList.remove("is-visible"), 2600);
}

function selectIfPresent(select, value) {
  if ([...select.options].some((option) => option.value === value)) select.value = value;
}

function loadState() {
  if (state.mode === "edit") {
    const label = "Update code on slide";
    document.querySelector("#modeNote").textContent = "Editing the selected code block";
    document.querySelector("#applyLabel").textContent = label;
    document.querySelector("#hintAction").textContent = label;
  }
  selectIfPresent(languageSelect, state.language);
  selectIfPresent(fontSizeSelect, state.fontSize);
  selectIfPresent(numberGapSelect, state.numberGap);
  if (state.lineNumbers !== undefined) lineNumberToggle.checked = state.lineNumbers === "1";
  codeInput.value = state.source || "";
  (state.lineHighlights || "").split(",").forEach((entry) => {
    const match = entry.match(/^(\d+):([0-9A-F]{6})$/i);
    if (match) highlights.set(Number(match[1]), match[2].toUpperCase());
  });
}

function toggleHighlight(event) {
  const row = event.target.closest(".line-number");
  if (!row || event.target.classList.contains("highlight-swatch")) return;

  const line = Number(row.dataset.line);
  if (highlights.has(line)) {
    highlights.delete(line);
  } else {
    highlights.set(line, DEFAULT_HIGHLIGHT_COLOR);
  }
  renderLineNumbers(codeInput.value.split("\n").length);
  renderHighlightBands();
}

function changeHighlightColor(event) {
  if (!event.target.classList.contains("highlight-swatch")) return;

  const row = event.target.closest(".line-number");
  highlights.set(Number(row.dataset.line), event.target.value.slice(1).toUpperCase());
  row.style.setProperty("--highlight", event.target.value);
  renderHighlightBands();
}

codeInput.addEventListener("input", render);
languageSelect.addEventListener("change", render);
lineNumberToggle.addEventListener("change", updateOutputOptions);
numberGapSelect.addEventListener("change", updateOutputOptions);
fontSizeSelect.addEventListener("change", updateOutputOptions);
lineNumbers.addEventListener("click", toggleHighlight);
lineNumbers.addEventListener("input", changeHighlightColor);
formatButton.addEventListener("click", formatCode);
applyButton.addEventListener("click", applyToSlide);
cancelButton.addEventListener("click", cancel);
document.addEventListener("keydown", (event) => {
  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
    event.preventDefault();
    applyToSlide();
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancel();
  }
});

loadState();
render();
updateOutputOptions();
codeInput.setSelectionRange(0, 0);
codeInput.focus();
