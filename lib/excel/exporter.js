const XLSX = require('xlsx');

function exportWorkbookFromRows(rows, fileName) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, '결과');
  XLSX.writeFile(workbook, fileName);
}

module.exports = {
  exportWorkbookFromRows,
};
