const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const XlsxPopulate = require("xlsx-populate");

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, "index.html"));
}
app.whenReady().then(createWindow);

function columnLetter(col) {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function columnNumber(letter) {
  return letter
    .split("")
    .reduce((r, c) => r * 26 + (c.charCodeAt(0) - 64), 0);
}

ipcMain.handle("select-file", async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });
  return canceled ? null : filePaths[0];
});

ipcMain.handle("get-sheets", async (_, filePath) => {
  const wb = await XlsxPopulate.fromFileAsync(filePath);
  return wb.sheets().map((s) => s.name());
});

ipcMain.handle("get-columns", async (_, filePath, sheetName) => {
  const wb = await XlsxPopulate.fromFileAsync(filePath);
  const sheet = wb.sheet(sheetName);
  const used = sheet.usedRange();
  if (!used) return [];
  const lastCol = used.endCell().columnNumber();
  const cols = [];
  for (let i = 1; i <= lastCol; i++) {
    const letter = columnLetter(i);
    const header = String(sheet.cell(1, i).value() || letter).trim();
    cols.push({ letter, header });
  }
  return cols;
});

ipcMain.handle(
  "process-files",
  async (_, { fileF, sheetsF, fileD, sheetD, colD, saveModeF, saveModeD }) => {
    const wbF = await XlsxPopulate.fromFileAsync(fileF);
    const wbD = await XlsxPopulate.fromFileAsync(fileD);
    const wsD = wbD.sheet(sheetD);

    const setF = new Set();
    for (const { sheet, col } of sheetsF) {
      const ws = wbF.sheet(sheet);
      const used = ws.usedRange();
      if (!used) continue;
      const lastRow = used.endCell().rowNumber();
      const vals =
        lastRow >= 2
          ? ws
              .range(2, columnNumber(col), lastRow, columnNumber(col))
              .value()
              .flat()
          : [];
      vals
        .filter((v) => v != null)
        .map((v) => String(v).trim().toLowerCase())
        .forEach((v) => setF.add(v));
    }

    // Стилизация D
    const usedD = wsD.usedRange();
    const lastD = usedD ? usedD.endCell().rowNumber() : 1;
    for (let r = 2; r <= lastD; r++) {
      const cell = wsD.cell(r, columnNumber(colD));
      const val = String(cell.value() || "").trim().toLowerCase();
      cell.style({ fontColor: setF.has(val) ? "FFFF0000" : "FF00AA00" });
    }

    // Стилизация F
    for (const { sheet, col } of sheetsF) {
      const ws = wbF.sheet(sheet);
      const used = ws.usedRange();
      if (!used) continue;
      const lastRow = used.endCell().rowNumber();
      for (let r = 2; r <= lastRow; r++) {
        const cell = ws.cell(r, columnNumber(col));
        const val = String(cell.value() || "").trim().toLowerCase();
        if (setF.has(val)) cell.style({ fontColor: "FFFF0000" });
      }
    }

    const outF =
      saveModeF === "overwrite"
        ? fileF
        : fileF.replace(/\.xlsx$/, "_colored.xlsx");
    const outD =
      saveModeD === "overwrite"
        ? fileD
        : fileD.replace(/\.xlsx$/, "_colored.xlsx");
    await Promise.all([wbF.toFileAsync(outF), wbD.toFileAsync(outD)]);
    return { outF, outD };
  }
);
