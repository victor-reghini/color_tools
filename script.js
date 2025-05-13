let jsonData = [];
let allColumns = [];

const fileInput = document.getElementById('fileInput');
const columnAdd = document.getElementById('columnAdd');
const columnList = document.getElementById('columnList');
const exportBtn = document.getElementById('exportBtn');

fileInput.addEventListener('change', handleFile);
exportBtn.addEventListener('click', exportToXLSX);

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];

    jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    if (jsonData.length === 0) return;

    allColumns = Object.keys(jsonData[0]);
    populateColumnSelector();
    clearColumnList();
    updateTableAndJson();
  };

  reader.readAsArrayBuffer(file);
}

function populateColumnSelector() {
  columnAdd.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.textContent = '--- Escolha uma coluna ---';
  placeholder.disabled = true;
  placeholder.selected = true;
  columnAdd.appendChild(placeholder);

  allColumns.forEach(col => {
    const option = document.createElement('option');
    option.value = col;
    option.textContent = col;
    columnAdd.appendChild(option);
  });

  columnAdd.onchange = () => {
    const selectedCol = columnAdd.value;
    if (!isColumnInList(selectedCol)) {
      addColumnToList(selectedCol);
      updateTableAndJson();
    }
    columnAdd.selectedIndex = 0;
  };
}

function clearColumnList() {
  columnList.innerHTML = '';
}

function addColumnToList(colName) {
  const li = document.createElement('li');
  li.dataset.column = colName;

  const span = document.createElement('span');
  span.textContent = colName;

  const removeBtn = document.createElement('button');
  removeBtn.textContent = '❌';
  removeBtn.onclick = () => {
    columnList.removeChild(li);
    updateTableAndJson();
  };

  li.appendChild(span);
  li.appendChild(removeBtn);
  columnList.appendChild(li);
}

function isColumnInList(colName) {
  return Array.from(columnList.children).some(li => li.dataset.column === colName);
}

function getSelectedColumns() {
  return Array.from(columnList.children).map(li => li.dataset.column);
}

function updateTableAndJson() {
  const selectedCols = getSelectedColumns();
  if (!selectedCols.length) {
    document.getElementById('previewTable').innerHTML = '<tr><td>Nenhuma coluna selecionada.</td></tr>';
    document.getElementById('output').textContent = '[]';
    return;
  }

  const filteredData = jsonData.map(row => {
    let filteredRow = {};
    selectedCols.forEach(col => filteredRow[col] = row[col]);
    return filteredRow;
  });

  renderTable(filteredData, selectedCols);
}

function renderTable(data, columns) {
  const table = document.getElementById('previewTable');
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');
    columns.forEach(col => {
      const td = document.createElement('td');
      td.textContent = row[col] ?? '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

function exportToXLSX() {
  const selectedCols = getSelectedColumns();
  if (!selectedCols.length) {
    alert("Nenhuma coluna selecionada.");
    return;
  }

  const filteredData = jsonData.map(row => {
    let filteredRow = {};
    selectedCols.forEach(col => filteredRow[col] = row[col]);
    return filteredRow;
  });

  const worksheet = XLSX.utils.json_to_sheet(filteredData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Visualização");
  XLSX.writeFile(workbook, "visualizacao.xlsx");
}

// Enable drag-and-drop reordering
Sortable.create(columnList, {
  animation: 150,
  onEnd: updateTableAndJson
});
