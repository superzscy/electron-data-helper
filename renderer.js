const ExcelJS = require('exceljs');
const { dialog } = require('@electron/remote');
const fs = require('fs');

document.getElementById('selectFileBtn').addEventListener('click', async () => {
  const result = await dialog.showOpenDialog({
    filters: [{ name: 'Excel 文件', extensions: ['xlsx'] }],
    properties: ['openFile']
  });

  if (result.canceled || result.filePaths.length === 0) {
    return;
  }

  const filePath = result.filePaths[0];

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  const output = document.getElementById('output');
  output.innerHTML = '';

  const table = document.createElement('table');
  table.border = 1;
  table.cellPadding = 5;
  table.style.borderCollapse = 'collapse';

  worksheet.eachRow((row, rowNumber) => {
    const tr = document.createElement('tr');
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const td = document.createElement(rowNumber === 1 ? 'th' : 'td');
      td.innerText = cell.value != null ? cell.value.toString() : '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  output.appendChild(table);
});
