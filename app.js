const state = {
  files: [],
  datasets: [],
};

const byId = (id) => document.getElementById(id);

const refs = {
  fileInput: byId("file-input"),
  fileList: byId("file-list"),
  profile: byId("profile"),
  layout: byId("layout"),
  exportFormat: byId("export-format"),
  outputName: byId("output-name"),
  delimiter: byId("delimiter"),
  genericLastLines: byId("generic-last-lines"),
  rawLastLines: byId("raw-last-lines"),
  chiExtractMode: byId("chi-extract-mode"),
  chiLastPoints: byId("chi-last-points"),
  chiLastSeconds: byId("chi-last-seconds"),
  timeMin: byId("time-min"),
  timeMax: byId("time-max"),
  genericOptions: byId("generic-options"),
  rawOptions: byId("raw-options"),
  chiOptions: byId("chi-options"),
  chiLastPointsField: byId("chi-last-points-field"),
  chiLastSecondsField: byId("chi-last-seconds-field"),
  chiTimeMinField: byId("chi-time-min-field"),
  chiTimeMaxField: byId("chi-time-max-field"),
  status: byId("status"),
  previewHead: document.querySelector("#preview-table thead"),
  previewBody: document.querySelector("#preview-table tbody"),
  buildPreview: byId("build-preview"),
  exportButton: byId("export-button"),
  removeSelected: byId("remove-selected"),
  moveUp: byId("move-up"),
  moveDown: byId("move-down"),
  clearFiles: byId("clear-files"),
};

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

function asNumber(value) {
  if (typeof value !== "string") {
    return value;
  }
  const stripped = value.trim();
  if (!stripped) {
    return value;
  }
  if (/^-?\d+(,\d{3})*$/.test(stripped)) {
    return Number(stripped.replace(/,/g, ""));
  }
  const numeric = Number(stripped);
  return Number.isFinite(numeric) ? numeric : value;
}

function parseNumericValue(value) {
  const normalized = asNumber(value);
  return typeof normalized === "number" && Number.isFinite(normalized) ? normalized : NaN;
}

function transformCell(value) {
  return asNumber(value);
}

function toDisplayName(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }
  const parts = text.split(/[\\/]/);
  return parts[parts.length - 1] || text;
}

function filterLastItems(items, count) {
  return count > 0 ? items.slice(-count) : items;
}

function formatNumber(value, digits = 6) {
  return Number.isFinite(value) ? Number(value.toFixed(digits)) : "";
}

function normalizeHeaderText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function describeTimeRange(mode, rows, timeIndex, options) {
  if (!rows.length) {
    return "空区间";
  }
  const firstTime = Number(rows[0][timeIndex]);
  const lastTime = Number(rows[rows.length - 1][timeIndex]);

  if (mode === "last_seconds") {
    return `最后 ${options.chiLastSeconds}s (${formatNumber(firstTime, 3)}-${formatNumber(lastTime, 3)}s)`;
  }
  if (mode === "time_range") {
    return `${formatNumber(firstTime, 3)}-${formatNumber(lastTime, 3)}s`;
  }
  if (mode === "last_points") {
    return `最后 ${rows.length} 行 (${formatNumber(firstTime, 3)}-${formatNumber(lastTime, 3)}s)`;
  }
  return `全部数据 (${formatNumber(firstTime, 3)}-${formatNumber(lastTime, 3)}s)`;
}

function detectTimeColumnIndex(header) {
  const index = header.findIndex((cell) => /time/i.test(String(cell)));
  return index >= 0 ? index : 0;
}

function detectCurrentColumnIndex(header, timeIndex, dataRows = []) {
  const preferredIndex = header.findIndex((cell, index) => {
    const text = normalizeHeaderText(cell);
    if (index === timeIndex) {
      return false;
    }
    return (
      text === "i" ||
      text === "i/a" ||
      text.startsWith("i/") ||
      text.includes("/i") ||
      text.includes("current") ||
      text.includes("<i>") ||
      text.includes("current/")
    );
  });
  if (preferredIndex >= 0) {
    return preferredIndex;
  }

  // Fall back to the first non-time column that contains numeric data.
  for (let columnIndex = 0; columnIndex < header.length; columnIndex += 1) {
    if (columnIndex === timeIndex) {
      continue;
    }
    const hasNumericValues = dataRows.some((row) => Number.isFinite(parseNumericValue(row[columnIndex])));
    if (hasNumericValues) {
      return columnIndex;
    }
  }

  return timeIndex === 0 ? 1 : 0;
}

function parseGeneric(text, fileName, options) {
  let lines = splitLines(text);
  lines = filterLastItems(lines, options.genericLastLines);
  const rows = lines
    .filter((line) => line.trim())
    .map((line) =>
      options.delimiter
        ? line.split(options.delimiter).map((cell) => cell.trim())
        : [line]
    );
  return { kind: "plain", title: fileName, rows };
}

function parseRaw(text, fileName, options) {
  let lines = splitLines(text);
  lines = filterLastItems(lines, options.rawLastLines);
  const rows = lines
    .filter((line) => line.trim())
    .map((line) => [line]);
  return { kind: "plain", title: fileName, rows };
}

function parseChi(text, fallbackName, options) {
  let title = toDisplayName(fallbackName);
  const rows = [];
  let headerFound = false;

  for (const rawLine of splitLines(text)) {
    const line = rawLine.trim();
    if (line.startsWith("File:")) {
      title = toDisplayName(line.slice(5).trim() || fallbackName);
    }
    if (!headerFound && line.startsWith(options.chiHeaderPrefix)) {
      headerFound = true;
      rows.push(line.split(",").map((cell) => cell.trim()));
      continue;
    }
    if (headerFound && line) {
      rows.push(line.split(",").map((cell) => cell.trim()));
    }
  }

  if (!rows.length) {
    throw new Error(`未找到 CHI 表头: ${fallbackName}`);
  }

  const header = rows[0];
  const timeIndex = detectTimeColumnIndex(header);

  const dataRows = rows
    .slice(1)
    .filter((row) => Number.isFinite(Number(row[timeIndex])));

  if (!dataRows.length) {
    throw new Error(`没有可识别的 CHI 数据行: ${fallbackName}`);
  }

  const currentIndex = detectCurrentColumnIndex(header, timeIndex, dataRows);

  let filtered = dataRows;
  if (options.chiExtractMode === "last_points") {
    filtered = filterLastItems(dataRows, options.chiLastPoints);
  } else if (options.chiExtractMode === "last_seconds") {
    const lastTime = Number(dataRows[dataRows.length - 1][timeIndex]);
    const cutoff = lastTime - options.chiLastSeconds;
    filtered = dataRows.filter((row) => Number(row[timeIndex]) >= cutoff);
  } else if (options.chiExtractMode === "time_range") {
    filtered = dataRows.filter((row) => {
      const time = Number(row[timeIndex]);
      if (options.timeMin !== null && time < options.timeMin) {
        return false;
      }
      if (options.timeMax !== null && time > options.timeMax) {
        return false;
      }
      return true;
    });
  }

  if (!filtered.length) {
    throw new Error(`选定时间范围没有数据: ${fallbackName}`);
  }

  const numericValues = filtered
    .map((row) => parseNumericValue(row[currentIndex]))
    .filter((value) => Number.isFinite(value));
  if (!numericValues.length) {
    throw new Error(`无法识别电流列数值: ${fallbackName}`);
  }
  const averageValue = numericValues.length
    ? numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length
    : NaN;
  const selectionLabel = describeTimeRange(
    options.chiExtractMode,
    filtered,
    timeIndex,
    options
  );

  return {
    kind: "chi",
    title,
    header,
    timeIndex,
    currentIndex,
    dataRows: filtered,
    rows: [header, ...filtered],
    selectionLabel,
    sourceLabel: `${title} | ${selectionLabel}`,
    summaryLabel: header[currentIndex] || "电流列",
    averageValue,
    summaryRow: [title, selectionLabel, header[currentIndex] || "电流列", averageValue],
  };
}

async function buildDatasets() {
  const options = collectOptions();
  if (!state.files.length) {
    throw new Error("请先添加至少一个 TXT 文件。");
  }

  const datasets = [];
  for (const file of state.files) {
    const text = await readFileText(file);
    if (options.profile === "chi") {
      datasets.push(parseChi(text, file.name, options));
    } else if (options.profile === "raw") {
      datasets.push(parseRaw(text, file.name, options));
    } else {
      datasets.push(parseGeneric(text, file.name, options));
    }
  }
  state.datasets = datasets;
  return { datasets, options };
}

function collectOptions() {
  return {
    profile: refs.profile.value,
    layout: refs.layout.value,
    exportFormat: refs.exportFormat.value,
    outputName: refs.outputName.value.trim() || "txt-excel-export",
    includeTitle: true,
    skipEmpty: true,
    titleGap: 0,
    colGap: 1,
    rowGap: 0,
    delimiter: refs.delimiter.value,
    genericLastLines: Math.max(0, Number(refs.genericLastLines.value) || 0),
    rawLastLines: Math.max(0, Number(refs.rawLastLines.value) || 0),
    chiHeaderPrefix: "Time/s",
    chiExtractMode: refs.chiExtractMode.value,
    chiLastPoints: Math.max(0, Number(refs.chiLastPoints.value) || 0),
    chiLastSeconds: Math.max(0, Number(refs.chiLastSeconds.value) || 0),
    timeMin: refs.timeMin.value === "" ? null : Number(refs.timeMin.value),
    timeMax: refs.timeMax.value === "" ? null : Number(refs.timeMax.value),
  };
}

function datasetWidth(dataset) {
  return dataset.rows.reduce((max, row) => Math.max(max, row.length), 1);
}

function pushGapRows(output, count) {
  for (let i = 0; i < count; i += 1) {
    output.push([]);
  }
}

function rowsForGenericLayout(datasets, options) {
  if (options.layout === "vertical") {
    const output = [];
    datasets.forEach((dataset) => {
      if (options.includeTitle) {
        output.push([dataset.title]);
        pushGapRows(output, options.titleGap);
      }
      dataset.rows.forEach((row) => {
        output.push(row.map((cell) => transformCell(cell)));
      });
      pushGapRows(output, options.rowGap);
    });
    return output;
  }

  if (options.layout === "horizontal") {
    const grid = [];
    let currentCol = 0;
    datasets.forEach((dataset) => {
      let dataStartRow = 0;
      if (options.includeTitle) {
        grid[0] = grid[0] || [];
        grid[0][currentCol] = dataset.title;
        dataStartRow = options.titleGap + 1;
      }
      dataset.rows.forEach((row, rowOffset) => {
        const targetRow = dataStartRow + rowOffset;
        grid[targetRow] = grid[targetRow] || [];
        row.forEach((cell, colOffset) => {
          grid[targetRow][currentCol + colOffset] = transformCell(cell);
        });
      });
      currentCol += datasetWidth(dataset) + options.colGap;
    });
    return grid.map((row = []) => row.map((cell) => (cell === undefined ? "" : cell)));
  }

  throw new Error("CSV 只支持横向拼接或纵向合并。");
}

function buildChiHorizontalRows(datasets, options) {
  const grid = [];
  let currentCol = 0;

  datasets.forEach((dataset) => {
    let currentRow = 0;
    if (options.includeTitle) {
      grid[currentRow] = grid[currentRow] || [];
      grid[currentRow][currentCol] = dataset.title;
      currentRow += 1;
      pushGridGapRows(grid, currentRow, options.titleGap);
      currentRow += options.titleGap;
    }

    grid[currentRow] = grid[currentRow] || [];
    grid[currentRow][currentCol] = dataset.selectionLabel;
    currentRow += 1;

    dataset.rows.forEach((row, rowOffset) => {
      const targetRow = currentRow + rowOffset;
      grid[targetRow] = grid[targetRow] || [];
      row.forEach((cell, colOffset) => {
        grid[targetRow][currentCol + colOffset] = transformCell(cell);
      });
    });
    currentRow += dataset.rows.length;

    const summaryRows = [
      ["统计项", "值"],
      ["文件名", dataset.title],
      ["时间区间", dataset.selectionLabel],
      ["统计列", dataset.summaryLabel],
      ["平均值", dataset.averageValue],
    ];

    summaryRows.forEach((row, rowOffset) => {
      const targetRow = currentRow + rowOffset;
      grid[targetRow] = grid[targetRow] || [];
      row.forEach((cell, colOffset) => {
        grid[targetRow][currentCol + colOffset] = transformCell(cell);
      });
    });

    currentCol += Math.max(dataset.header.length, 2) + options.colGap;
  });

  return grid.map((row = []) => row.map((cell) => (cell === undefined ? "" : cell)));
}

function pushGridGapRows(grid, startRow, count) {
  for (let i = 0; i < count; i += 1) {
    grid[startRow + i] = grid[startRow + i] || [];
  }
}

function buildChiVerticalRows(datasets) {
  const first = datasets[0];
  const output = [["来源", ...first.header]];

  datasets.forEach((dataset) => {
    dataset.dataRows.forEach((row) => {
      output.push([dataset.sourceLabel, ...row.map((cell) => transformCell(cell))]);
    });
  });

  output.push([]);
  output.push(["文件名", "时间区间", "统计列", "平均值"]);
  datasets.forEach((dataset) => {
    output.push([
      transformCell(dataset.title),
      transformCell(dataset.selectionLabel),
      transformCell(dataset.summaryLabel),
      transformCell(dataset.averageValue),
    ]);
  });

  return output;
}

function rowsForLayout(datasets, options) {
  if (options.profile === "chi") {
    if (options.layout === "horizontal") {
      return buildChiHorizontalRows(datasets, options);
    }
    if (options.layout === "vertical") {
      return buildChiVerticalRows(datasets);
    }
    throw new Error("CSV 只支持横向拼接或纵向合并。");
  }

  return rowsForGenericLayout(datasets, options);
}

function buildWorkbook(datasets, options) {
  const workbook = XLSX.utils.book_new();
  if (options.layout === "both") {
    const horizontalRows = rowsForLayout(datasets, { ...options, layout: "horizontal" });
    const verticalRows = rowsForLayout(datasets, { ...options, layout: "vertical" });
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(horizontalRows), "Horizontal Merge");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(verticalRows), "Vertical Merge");
    return workbook;
  }

  const rows = rowsForLayout(datasets, options);
  const sheetName = options.layout === "horizontal" ? "Horizontal Merge" : "Vertical Merge";
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return workbook;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function exportCsv(datasets, options) {
  const rows = rowsForLayout(datasets, options);
  const csv = rows
    .map((row) =>
      row
        .map((cell) => {
          const value = cell === undefined || cell === null ? "" : String(cell);
          return /[",\n]/.test(value) ? `"${value.replace(/"/g, "\"\"")}"` : value;
        })
        .join(",")
    )
    .join("\r\n");
  const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" });
  downloadBlob(blob, `${options.outputName}.csv`);
}

function exportXlsx(datasets, options) {
  if (typeof XLSX === "undefined") {
    throw new Error("XLSX 库未加载完成，请检查网络或稍后重试。");
  }
  const workbook = buildWorkbook(datasets, options);
  XLSX.writeFile(workbook, `${options.outputName}.xlsx`, { compression: true });
}

function renderPreview(datasets, options) {
  refs.previewHead.innerHTML = "";
  refs.previewBody.innerHTML = "";

  const previewRows =
    options.layout === "both"
      ? rowsForLayout(datasets, { ...options, layout: "horizontal" })
      : rowsForLayout(datasets, options);
  const slicedRows = previewRows.slice(0, 12);
  const maxColumns = slicedRows.reduce((max, row) => Math.max(max, row.length), 0);

  const headRow = document.createElement("tr");
  for (let col = 0; col < maxColumns; col += 1) {
    const th = document.createElement("th");
    th.textContent = `列 ${col + 1}`;
    headRow.append(th);
  }
  refs.previewHead.append(headRow);

  slicedRows.forEach((row) => {
    const tr = document.createElement("tr");
    for (let col = 0; col < maxColumns; col += 1) {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.append(td);
    }
    refs.previewBody.append(tr);
  });
}

function updateModeVisibility() {
  const profile = refs.profile.value;
  refs.genericOptions.classList.toggle("hidden", profile !== "generic");
  refs.rawOptions.classList.toggle("hidden", profile !== "raw");
  refs.chiOptions.classList.toggle("hidden", profile !== "chi");

  const chiMode = refs.chiExtractMode.value;
  refs.chiLastPointsField.classList.toggle("hidden", chiMode !== "last_points");
  refs.chiLastSecondsField.classList.toggle("hidden", chiMode !== "last_seconds");
  refs.chiTimeMinField.classList.toggle("hidden", chiMode !== "time_range");
  refs.chiTimeMaxField.classList.toggle("hidden", chiMode !== "time_range");

}

function selectedIndex() {
  return refs.fileList.selectedIndex;
}

function moveSelected(offset) {
  const index = selectedIndex();
  if (index < 0) {
    return;
  }
  const target = index + offset;
  if (target < 0 || target >= state.files.length) {
    return;
  }
  [state.files[index], state.files[target]] = [state.files[target], state.files[index]];
  refreshFileList();
  refs.fileList.selectedIndex = target;
}

async function handlePreview() {
  try {
    const { datasets, options } = await buildDatasets();
    renderPreview(datasets, options);
    setStatus(`已载入 ${datasets.length} 个文件，预览显示前 12 行。`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function handleExport() {
  try {
    const { datasets, options } = await buildDatasets();
    if (options.exportFormat === "csv") {
      if (options.layout === "both") {
        throw new Error("CSV 不支持同时导出两张表，请改成横向或纵向。");
      }
      exportCsv(datasets, options);
    } else {
      exportXlsx(datasets, options);
    }
    setStatus(`导出完成: ${options.outputName}.${options.exportFormat}`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

refs.fileInput.addEventListener("change", (event) => {
  const files = Array.from(event.target.files || []);
  state.files.push(...files);
  refreshFileList();
  setStatus(`已添加 ${files.length} 个文件。`);
  refs.fileInput.value = "";
});

refs.removeSelected.addEventListener("click", () => {
  const selected = Array.from(refs.fileList.selectedOptions).map((option) => Number(option.value));
  state.files = state.files.filter((_, index) => !selected.includes(index));
  refreshFileList();
});

refs.moveUp.addEventListener("click", () => moveSelected(-1));
refs.moveDown.addEventListener("click", () => moveSelected(1));
refs.clearFiles.addEventListener("click", () => {
  state.files = [];
  refreshFileList();
  refs.previewHead.innerHTML = "";
  refs.previewBody.innerHTML = "";
  setStatus("文件列表已清空。");
});

refs.profile.addEventListener("change", updateModeVisibility);
refs.layout.addEventListener("change", updateModeVisibility);
refs.chiExtractMode.addEventListener("change", updateModeVisibility);
refs.exportFormat.addEventListener("change", () => {
  if (refs.exportFormat.value === "csv" && refs.layout.value === "both") {
    refs.layout.value = "horizontal";
  }
  updateModeVisibility();
});
refs.buildPreview.addEventListener("click", handlePreview);
refs.exportButton.addEventListener("click", handleExport);

updateModeVisibility();
