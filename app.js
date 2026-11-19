const state = {
  files: [],
  datasets: [],
  previewRows: [],
  previewColumnCount: 0,
  previewRange: { start: -1, end: -1 },
  previewFrame: null,
  previewMode: "data",
};

const PREVIEW_ROW_HEIGHT = 29;
const PREVIEW_OVERSCAN = 8;
const byId = (id) => document.getElementById(id);
const refs = {
  fileInput: byId("file-input"),
  fileList: byId("file-list"),
  layoutInputs: Array.from(document.querySelectorAll('input[name="layout"]')),
  outputName: byId("output-name"),
  statisticsDuration: byId("statistics-duration"),
  includeMetadata: byId("chi-include-metadata"),
  status: byId("status"),
  previewWrap: document.querySelector(".preview-wrap"),
  previewHead: document.querySelector("#preview-table thead"),
  previewBody: document.querySelector("#preview-table tbody"),
  buildPreview: byId("build-preview"),
  exportButton: byId("export-button"),
  removeSelected: byId("remove-selected"),
  moveUp: byId("move-up"),
  moveDown: byId("move-down"),
  clearFiles: byId("clear-files"),
  previewDataTab: byId("preview-data-tab"),
  previewStatisticsTab: byId("preview-statistics-tab"),
};

function checkedValue(name) {
  return document.querySelector(`input[name="${name}"]:checked`)?.value || "";
}

function setStatus(message, isError = false) {
  refs.status.textContent = message;
  refs.status.classList.toggle("error", isError);
}

function refreshFileList() {
  refs.fileList.innerHTML = "";
  state.files.forEach((file, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `${index + 1}. ${file.name}`;
    refs.fileList.append(option);
  });
}

function readFileText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error(`读取失败: ${file.name}`));
    reader.readAsText(file);
  });
}

function splitLines(text) {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function splitCsvLine(line) {
  const cells = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      if (quoted && line[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (character === "," && !quoted) {
      cells.push(cell.trim());
      cell = "";
    } else {
      cell += character;
    }
  }
  cells.push(cell.trim());
  return cells;
}

function isNumeric(value) {
  return value !== "" && Number.isFinite(Number(value));
}

function asExcelValue(value) {
  const trimmed = String(value ?? "").trim();
  if (trimmed && Number.isFinite(Number(trimmed))) return Number(trimmed);
  return trimmed;
}

function isDataHeader(cells) {
  if (cells.length < 2) return false;
  const hasTime = cells.some((cell) => /^time\b|^time\s*\//i.test(cell));
  const hasPotential = cells.some((cell) => /^potential\b/i.test(cell));
  const hasCurrent = cells.some((cell) => /^current\b|^i\d+\s*\/\s*a$/i.test(cell));
  return hasTime || (hasPotential && hasCurrent);
}

function parseMetadata(lines) {
  const metadata = new Map();
  let noteIndex = 1;
  lines.forEach((line, index) => {
    const value = line.trim();
    if (!value) return;
    if (index === 0) {
      metadata.set("日期时间", value);
      return;
    }
    const match = value.match(/^\s*([^:=]+?)\s*[:=]\s*(.*)$/);
    if (match) {
      const key = match[1].trim();
      let label = key;
      let suffix = 2;
      while (metadata.has(label)) label = `${key} ${suffix++}`;
      metadata.set(label, match[2].trim());
    } else {
      metadata.set(`说明 ${noteIndex++}`, value);
    }
  });
  return metadata;
}

function parseChi(text, fileName) {
  const metadataLines = [];
  let header = null;
  let dataRows = [];

  for (const rawLine of splitLines(text)) {
    const line = rawLine.trim();
    if (!line) continue;
    const cells = splitCsvLine(line);
    if (!header) {
      if (isDataHeader(cells)) {
        header = cells;
      } else {
        metadataLines.push(line);
      }
      continue;
    }

    if (cells.join("\u0000") === header.join("\u0000")) continue;
    const numericCount = cells.filter(isNumeric).length;
    if (cells.length >= 2 && numericCount >= 2) {
      dataRows.push(cells.map(asExcelValue));
    }
  }

  if (!header) {
    throw new Error(`未找到可识别的数据表头：${fileName}。支持 Time 列，或 Potential/V 配合 Current/A、i1/A 等电流列。`);
  }
  if (!dataRows.length) {
    throw new Error(`表头后没有可转换的数值数据：${fileName}`);
  }

  return {
    title: fileName,
    header,
    dataRows,
    metadata: parseMetadata(metadataLines),
  };
}

function collectOptions() {
  return {
    layout: checkedValue("layout") || "horizontal",
    outputName: refs.outputName.value.trim() || "chi-converted-data",
    statisticsDuration: Math.max(0.1, Number(refs.statisticsDuration.value) || 30),
    includeMetadata: refs.includeMetadata.checked,
  };
}

async function buildDatasets() {
  if (!state.files.length) throw new Error("请先添加至少一个 CHI TXT 文件。");
  const datasets = [];
  for (const file of state.files) {
    datasets.push(parseChi(await readFileText(file), file.name));
  }
  state.datasets = datasets;
  return datasets;
}

function buildHorizontalRows(datasets) {
  const rows = [];
  const widths = datasets.map((dataset) => dataset.header.length);
  const totalColumns = widths.reduce((sum, width) => sum + width, 0);
  rows.push(datasets.flatMap((dataset) => [dataset.title, ...Array(dataset.header.length - 1).fill("")]));
  rows.push(datasets.flatMap((dataset) => dataset.header));
  const maxDataRows = Math.max(...datasets.map((dataset) => dataset.dataRows.length));
  for (let rowIndex = 0; rowIndex < maxDataRows; rowIndex += 1) {
    const row = [];
    datasets.forEach((dataset) => {
      const values = dataset.dataRows[rowIndex] || [];
      row.push(...Array.from({ length: dataset.header.length }, (_, column) => values[column] ?? ""));
    });
    rows.push(Array.from({ length: totalColumns }, (_, column) => row[column] ?? ""));
  }
  return rows;
}

function buildVerticalRows(datasets) {
  const columns = [];
  datasets.forEach((dataset) => dataset.header.forEach((header) => {
    if (!columns.includes(header)) columns.push(header);
  }));
  const rows = [["文件名", ...columns]];
  datasets.forEach((dataset) => {
    const indexes = new Map(dataset.header.map((header, index) => [header, index]));
    dataset.dataRows.forEach((dataRow) => {
      rows.push([dataset.title, ...columns.map((header) => {
        const index = indexes.get(header);
        return index === undefined ? "" : dataRow[index] ?? "";
      })]);
    });
  });
  return rows;
}

function buildDataRows(datasets, options) {
  return options.layout === "vertical"
    ? buildVerticalRows(datasets)
    : buildHorizontalRows(datasets);
}

function buildMetadataRows(datasets) {
  const keys = [];
  datasets.forEach((dataset) => dataset.metadata.forEach((_, key) => {
    if (!keys.includes(key)) keys.push(key);
  }));
  return [
    ["参数", ...datasets.map((dataset) => dataset.title)],
    ...keys.map((key) => [key, ...datasets.map((dataset) => dataset.metadata.get(key) ?? "")]),
  ];
}

function findTimeColumnIndex(header) {
  return header.findIndex((cell) => /^time(?:\b|\s*\/)/i.test(String(cell).trim()));
}

function roundedNumber(value, digits = 10) {
  if (!Number.isFinite(value)) return "";
  return Number(value.toFixed(digits));
}

function formatTime(value) {
  return Number.isFinite(value) ? roundedNumber(value, 6) : "";
}

function uniqueMeasurementLabels(dataset, timeIndex) {
  const counts = new Map();
  return dataset.header.map((header, index) => {
    if (index === timeIndex) return "";
    const base = String(header || `列 ${index + 1}`).trim() || `列 ${index + 1}`;
    const count = (counts.get(base) || 0) + 1;
    counts.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
}

function calculateDatasetStatistics(dataset, duration) {
  const timeIndex = findTimeColumnIndex(dataset.header);
  if (timeIndex < 0) {
    return {
      title: dataset.title,
      rangeLabel: "无 Time 列，无法按时长统计",
      actualDuration: "",
      pointCount: 0,
      averages: new Map(),
    };
  }

  const timedRows = dataset.dataRows
    .map((row) => ({ row, time: Number(row[timeIndex]) }))
    .filter(({ time }) => Number.isFinite(time));
  if (!timedRows.length) {
    return {
      title: dataset.title,
      rangeLabel: "Time 列没有数值",
      actualDuration: "",
      pointCount: 0,
      averages: new Map(),
    };
  }

  const endTime = Math.max(...timedRows.map(({ time }) => time));
  const cutoff = endTime - duration;
  const selected = timedRows.filter(({ time }) => time >= cutoff && time <= endTime);
  const startTime = Math.min(...selected.map(({ time }) => time));
  const labels = uniqueMeasurementLabels(dataset, timeIndex);
  const averages = new Map();

  dataset.header.forEach((_, columnIndex) => {
    if (columnIndex === timeIndex) return;
    const values = selected
      .map(({ row }) => Number(row[columnIndex]))
      .filter((value) => Number.isFinite(value));
    if (!values.length) return;
    averages.set(labels[columnIndex], values.reduce((sum, value) => sum + value, 0) / values.length);
  });

  return {
    title: dataset.title,
    rangeLabel: `${formatTime(startTime)}–${formatTime(endTime)} s`,
    actualDuration: roundedNumber(endTime - startTime, 6),
    pointCount: selected.length,
    averages,
  };
}

function buildStatisticsRows(datasets, duration) {
  const statistics = datasets.map((dataset) => calculateDatasetStatistics(dataset, duration));
  const labels = [];
  statistics.forEach((item) => item.averages.forEach((_, label) => {
    if (!labels.includes(label)) labels.push(label);
  }));
  return [
    ["文件名", "统计范围", "实际时长/s", "数据点数", ...labels.map((label) => `${label} 平均值`)],
    ...statistics.map((item) => [
      item.title,
      item.rangeLabel,
      item.actualDuration,
      item.pointCount,
      ...labels.map((label) => roundedNumber(item.averages.get(label), 12)),
    ]),
  ];
}

function sheetFromRows(rows, columnWidth = 16) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  worksheet["!cols"] = Array.from({ length: width }, () => ({ wch: columnWidth }));
  return worksheet;
}

function buildWorkbook(datasets, options) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel 导出组件尚未加载，请检查网络后刷新页面。");
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheetFromRows(buildDataRows(datasets, options)), "Origin Data");
  XLSX.utils.book_append_sheet(
    workbook,
    sheetFromRows(buildStatisticsRows(datasets, options.statisticsDuration), 20),
    "Statistics"
  );
  if (options.includeMetadata) {
    XLSX.utils.book_append_sheet(workbook, sheetFromRows(buildMetadataRows(datasets), 24), "Header Info");
  }
  return workbook;
}

function exportXlsx(datasets, options) {
  XLSX.writeFile(buildWorkbook(datasets, options), `${options.outputName}.xlsx`, { compression: true });
}

function appendSpacerRow(height) {
  if (height <= 0) return;
  const tr = document.createElement("tr");
  tr.className = "preview-spacer";
  const td = document.createElement("td");
  td.colSpan = Math.max(1, state.previewColumnCount);
  td.style.height = `${height}px`;
  tr.append(td);
  refs.previewBody.append(tr);
}

function renderVirtualPreview(force = false) {
  if (!state.previewRows.length) {
    refs.previewBody.innerHTML = "";
    return;
  }
  const scrollTop = refs.previewWrap.scrollTop;
  const viewportHeight = refs.previewWrap.clientHeight || 420;
  const start = Math.max(0, Math.floor(scrollTop / PREVIEW_ROW_HEIGHT) - PREVIEW_OVERSCAN);
  const end = Math.min(state.previewRows.length, Math.ceil((scrollTop + viewportHeight) / PREVIEW_ROW_HEIGHT) + PREVIEW_OVERSCAN);
  if (!force && start === state.previewRange.start && end === state.previewRange.end) return;
  state.previewRange = { start, end };
  refs.previewBody.innerHTML = "";
  appendSpacerRow(start * PREVIEW_ROW_HEIGHT);
  for (let rowIndex = start; rowIndex < end; rowIndex += 1) {
    const tr = document.createElement("tr");
    tr.className = "preview-data-row";
    for (let column = 0; column < state.previewColumnCount; column += 1) {
      const td = document.createElement("td");
      td.textContent = state.previewRows[rowIndex][column] ?? "";
      tr.append(td);
    }
    refs.previewBody.append(tr);
  }
  appendSpacerRow((state.previewRows.length - end) * PREVIEW_ROW_HEIGHT);
}

function renderPreview(datasets, options, mode = state.previewMode) {
  state.previewMode = mode;
  state.previewRows = mode === "statistics"
    ? buildStatisticsRows(datasets, options.statisticsDuration)
    : buildDataRows(datasets, options);
  state.previewColumnCount = state.previewRows.reduce((max, row) => Math.max(max, row.length), 0);
  state.previewRange = { start: -1, end: -1 };
  refs.previewHead.innerHTML = "";
  refs.previewBody.innerHTML = "";
  const head = document.createElement("tr");
  for (let column = 0; column < state.previewColumnCount; column += 1) {
    const th = document.createElement("th");
    th.textContent = `列 ${column + 1}`;
    head.append(th);
  }
  refs.previewHead.append(head);
  refs.previewWrap.scrollTop = 0;
  refs.previewDataTab.classList.toggle("active", mode === "data");
  refs.previewStatisticsTab.classList.toggle("active", mode === "statistics");
  renderVirtualPreview(true);
}

async function handlePreview() {
  try {
    const datasets = await buildDatasets();
    renderPreview(datasets, collectOptions());
    refs.previewDataTab.disabled = false;
    refs.previewStatisticsTab.disabled = false;
    const timedCount = datasets.filter(({ header }) => findTimeColumnIndex(header) >= 0).length;
    setStatus(`已读取 ${datasets.length} 个文件。${timedCount ? `其中 ${timedCount} 个文件可按 Time 列统计。` : "这批文件没有 Time 列，已读取原始数据。"}`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function handleExport() {
  try {
    const datasets = await buildDatasets();
    exportXlsx(datasets, collectOptions());
    setStatus(`转换完成：${collectOptions().outputName}.xlsx`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

function selectedIndex() {
  return refs.fileList.selectedIndex;
}

function moveSelected(offset) {
  const index = selectedIndex();
  const target = index + offset;
  if (index < 0 || target < 0 || target >= state.files.length) return;
  [state.files[index], state.files[target]] = [state.files[target], state.files[index]];
  refreshFileList();
  refs.fileList.selectedIndex = target;
}

refs.fileInput.addEventListener("change", (event) => {
  const files = Array.from(event.target.files || []);
  state.files.push(...files);
  refreshFileList();
  setStatus(`已添加 ${files.length} 个 TXT 文件。`);
  refs.fileInput.value = "";
});

refs.removeSelected.addEventListener("click", () => {
  const selected = new Set(Array.from(refs.fileList.selectedOptions, (option) => Number(option.value)));
  state.files = state.files.filter((_, index) => !selected.has(index));
  refreshFileList();
});
refs.moveUp.addEventListener("click", () => moveSelected(-1));
refs.moveDown.addEventListener("click", () => moveSelected(1));
refs.clearFiles.addEventListener("click", () => {
  state.files = [];
  state.datasets = [];
  state.previewRows = [];
  state.previewColumnCount = 0;
  state.previewRange = { start: -1, end: -1 };
  state.previewMode = "data";
  refreshFileList();
  refs.previewHead.innerHTML = "";
  refs.previewBody.innerHTML = "";
  refs.previewDataTab.disabled = true;
  refs.previewStatisticsTab.disabled = true;
  refs.previewDataTab.classList.add("active");
  refs.previewStatisticsTab.classList.remove("active");
  setStatus("文件列表已清空。");
});
refs.previewWrap.addEventListener("scroll", () => {
  if (state.previewFrame !== null) return;
  state.previewFrame = requestAnimationFrame(() => {
    state.previewFrame = null;
    renderVirtualPreview();
  });
});
refs.buildPreview.addEventListener("click", handlePreview);
refs.exportButton.addEventListener("click", handleExport);
refs.previewDataTab.addEventListener("click", () => {
  if (state.datasets.length) renderPreview(state.datasets, collectOptions(), "data");
});
refs.previewStatisticsTab.addEventListener("click", () => {
  if (state.datasets.length) renderPreview(state.datasets, collectOptions(), "statistics");
});
refs.layoutInputs.forEach((input) => input.addEventListener("change", () => {
  if (state.datasets.length) renderPreview(state.datasets, collectOptions());
}));
refs.statisticsDuration.addEventListener("change", () => {
  if (state.datasets.length) renderPreview(state.datasets, collectOptions());
});
