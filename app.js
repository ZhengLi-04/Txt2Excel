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
  includeTitle: byId("include-title"),
  skipEmpty: byId("skip-empty"),
  coerceNumbers: byId("coerce-numbers"),
  titleGap: byId("title-gap"),
  colGap: byId("col-gap"),
  rowGap: byId("row-gap"),
  delimiter: byId("delimiter"),
  genericLastLines: byId("generic-last-lines"),
  rawLastLines: byId("raw-last-lines"),
  chiHeaderPrefix: byId("chi-header-prefix"),
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
  refs.status.style.background = isError ? "#f4d2d2" : "";
  refs.status.style.color = isError ? "#7a1d1d" : "";
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

function transformCell(value, shouldCoerceNumbers) {
  return shouldCoerceNumbers ? asNumber(value) : value;
}

function filterLastItems(items, count) {
  return count > 0 ? items.slice(-count) : items;
}

function parseGeneric(text, fileName, options) {
  let lines = splitLines(text);
  lines = filterLastItems(lines, options.genericLastLines);
  const rows = lines
    .filter((line) => !(options.skipEmpty && !line.trim()))
    .map((line) =>
      options.delimiter
        ? line.split(options.delimiter).map((cell) => cell.trim())
        : [line]
    );
  return { title: fileName, rows };
}

function parseRaw(text, fileName, options) {
  let lines = splitLines(text);
  lines = filterLastItems(lines, options.rawLastLines);
  const rows = lines
    .filter((line) => !(options.skipEmpty && !line.trim()))
    .map((line) => [line]);
  return { title: fileName, rows };
}

function parseChi(text, fallbackName, options) {
  let title = fallbackName;
  const rows = [];
  let headerFound = false;

  for (const rawLine of splitLines(text)) {
    const line = rawLine.trim();
    if (line.startsWith("File:")) {
      title = line.slice(5).trim() || fallbackName;
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
  const dataRows = rows.slice(1);
  let filtered = dataRows;

  if (options.chiExtractMode === "last_points") {
    filtered = filterLastItems(dataRows, options.chiLastPoints);
  } else if (options.chiExtractMode === "last_seconds") {
    const lastTime = [...dataRows].reverse().find((row) => Number.isFinite(Number(row[0])));
    if (!lastTime) {
      throw new Error(`时间列无法识别: ${fallbackName}`);
    }
    const cutoff = Number(lastTime[0]) - options.chiLastSeconds;
    filtered = dataRows.filter((row) => Number(row[0]) >= cutoff);
  } else if (options.chiExtractMode === "time_range") {
    filtered = dataRows.filter((row) => {
      const time = Number(row[0]);
      if (!Number.isFinite(time)) {
        return false;
      }
      if (options.timeMin !== null && time < options.timeMin) {
        return false;
      }
      if (options.timeMax !== null && time > options.timeMax) {
        return false;
      }
      return true;
    });
  }

  return { title, rows: [header, ...filtered] };
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
    includeTitle: refs.includeTitle.checked,
    skipEmpty: refs.skipEmpty.checked,
    coerceNumbers: refs.coerceNumbers.checked,
    titleGap: Math.max(0, Number(refs.titleGap.value) || 0),
    colGap: Math.max(0, Number(refs.colGap.value) || 0),
    rowGap: Math.max(0, Number(refs.rowGap.value) || 0),
    delimiter: refs.delimiter.value,
    genericLastLines: Math.max(0, Number(refs.genericLastLines.value) || 0),
    rawLastLines: Math.max(0, Number(refs.rawLastLines.value) || 0),
    chiHeaderPrefix: refs.chiHeaderPrefix.value.trim() || "Time/s",
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

function rowsForLayout(datasets, options) {
  if (options.layout === "vertical") {
    const output = [];
    datasets.forEach((dataset) => {
      if (options.includeTitle) {
        output.push([dataset.title]);
        for (let i = 0; i < options.titleGap; i += 1) {
          output.push([]);
        }
      }
      dataset.rows.forEach((row) => {
        output.push(row.map((cell) => transformCell(cell, options.coerceNumbers)));
      });
      for (let i = 0; i < options.rowGap; i += 1) {
        output.push([]);
      }
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
          grid[targetRow][currentCol + colOffset] = transformCell(cell, options.coerceNumbers);
        });
      });
      currentCol += datasetWidth(dataset) + options.colGap;
    });
    return grid.map((row = []) => row.map((cell) => (cell === undefined ? "" : cell)));
  }

  throw new Error("CSV 只支持横向拼接或纵向合并。");
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
  const slicedRows = previewRows.slice(0, 8);
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

  if (profile === "chi") {
    refs.layout.value = refs.layout.value === "both" ? "both" : refs.layout.value;
  }

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
    setStatus(`已载入 ${datasets.length} 个文件，预览显示前 8 行。`);
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
refs.chiExtractMode.addEventListener("change", updateModeVisibility);
refs.exportFormat.addEventListener("change", () => {
  if (refs.exportFormat.value === "csv" && refs.layout.value === "both") {
    refs.layout.value = "horizontal";
  }
});
refs.buildPreview.addEventListener("click", handlePreview);
refs.exportButton.addEventListener("click", handleExport);

updateModeVisibility();
