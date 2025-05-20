const { ipcRenderer } = require("electron");

const selectFile = () =>
  ipcRenderer.invoke("select-file").then((f) => f || null);
const getSheets = (f) => ipcRenderer.invoke("get-sheets", f);
const getColumns = (f, s) => ipcRenderer.invoke("get-columns", f, s);

// Загрузка колонок для блока
async function loadColsForBlock(filePath, sheetName, idx) {
  const cols = await getColumns(filePath, sheetName);
  const sel = document.querySelector(`.col-select[data-idx="${idx}"]`);
  sel.innerHTML = cols
    .map((c) => `<option value="${c.letter}">${c.letter} — ${c.header}</option>`)
    .join("");
}

window.addEventListener("DOMContentLoaded", () => {
  let currentF = null;
  let currentD = null;
  let allSheets = [];
  let availableSheets = [];
  const btnF = document.getElementById("btnF");
  const pathF = document.getElementById("pathF");
  const addBtn = document.getElementById("addSheetF");
  const container = document.getElementById("sheetsF-container");

  btnF.onclick = async () => {
    const f = await selectFile();
    if (!f) return;
    currentF = f;
    pathF.textContent = f.split(/[/\\]/).pop();
    allSheets = await getSheets(f);
    availableSheets = [...allSheets];
    container.innerHTML = "";
    updateAddBtn();
  };

  function updateAddBtn() {
    addBtn.disabled = availableSheets.length === 0;
  }

  addBtn.onclick = () => {
    if (!currentF) return alert("Сначала выберите файл.");
    if (availableSheets.length === 0) return;
    addSheetBlock();
  };

  function addSheetBlock() {
    const idx = container.children.length;
    const div = document.createElement("div");
    div.className = "sheet-row";
    div.dataset.selectedSheet = "";

    // Селект листов
    const sheetSel = document.createElement("select");
    sheetSel.className = "form-select sheet-select me-2";
    sheetSel.dataset.idx = idx;
    availableSheets.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      sheetSel.appendChild(opt);
    });

    // Селект колонок
    const colSel = document.createElement("select");
    colSel.className = "form-select col-select me-2";
    colSel.dataset.idx = idx;

    // Кнопка удаления
    const removeBtn = document.createElement("button");
    removeBtn.className = "btn btn-sm btn-danger";
    removeBtn.innerHTML = '<i class="bi bi-trash"></i>';

    div.append(sheetSel, colSel, removeBtn);
    container.appendChild(div);

    // Инициализация: первый элемент
    const initial = sheetSel.value;
    div.dataset.selectedSheet = initial;
    availableSheets = availableSheets.filter((s) => s !== initial);
    updateAddBtn();
    loadColsForBlock(currentF, initial, idx);

    sheetSel.onchange = async () => {
      const prev = div.dataset.selectedSheet;
      if (prev) availableSheets.push(prev);
      const newS = sheetSel.value;
      availableSheets = availableSheets.filter((s) => s !== newS);
      div.dataset.selectedSheet = newS;
      updateAddBtn();
      loadColsForBlock(currentF, newS, idx);
    };

    removeBtn.onclick = () => {
      const sel = div.dataset.selectedSheet;
      if (sel) availableSheets.push(sel);
      updateAddBtn();
      div.remove();
    };
  }

  // Шаг D и запуск
  const btnD = document.getElementById("btnD");
  const pathD = document.getElementById("pathD");
  const sheetD = document.getElementById("sheetD");
  const colD = document.getElementById("colD");

  btnD.onclick = async () => {
    const f = await selectFile();
    if (!f) return;
    currentD = f;
    pathD.textContent = f.split(/[/\\]/).pop();
    sheetD.innerHTML = (await getSheets(f)).map((n) => `<option>${n}</option>`).join("");
    loadColsForBlock(f, sheetD.value, "D");
  };
  sheetD.onchange = () => loadColsForBlock(currentD, sheetD.value, "D");

  document.getElementById("run").onclick = async () => {
    const sheetsF = Array.from(document.querySelectorAll(".sheet-row")).map((div) => ({
      sheet: div.querySelector(".sheet-select").value,
      col: div.querySelector(".col-select").value,
    }));
    const cfg = {
      fileF: currentF,
      sheetsF,
      fileD: currentD,
      sheetD: sheetD.value,
      colD: colD.value,
      saveModeF: document.querySelector('input[name="saveModeF"]:checked').value,
      saveModeD: document.querySelector('input[name="saveModeD"]:checked').value,
    };
    const res = await ipcRenderer.invoke("process-files", cfg);
    alert(`F: ${res.outF}\nD: ${res.outD}`);
  };
});
