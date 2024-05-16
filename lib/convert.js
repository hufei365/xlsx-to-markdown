const xlsx = require('xlsx');

/**
 * 将指定工作表和列转换为 Markdown 表格
 * @param {string} filePath - XLSX 文件路径
 * @param {string} sheetName - 工作表名称
 * @param {Array<string>} columns - 要转换的列
 * @returns {string} - 转换后的 Markdown 表格
 */
function convertToMarkdown(filePath, sheetName, columns) {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`Sheet ${sheetName} not found in the workbook.`);
  }

  const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  const header = json[0];
  const data = json.slice(1);

  const columnIndices = columns.map(col => header.indexOf(col));
  if (columnIndices.includes(-1)) {
    throw new Error(`One or more columns not found in the sheet.`);
  }

  const markdownRows = [];

  // 表格 header
  const markdownHeader = columnIndices.map(i => header[i]).join(' | ');
  const markdownSeparator = columnIndices.map(() => '---').join(' | ');
  markdownRows.push(`| ${markdownHeader} |`);
  markdownRows.push(`| ${markdownSeparator} |`);

  // 表格 body
  data.forEach(row => {
    const markdownRow = columnIndices.map(i => row[i] !== undefined ? row[i] : '').join(' | ');
    markdownRows.push(`| ${markdownRow} |`);
  });

  return markdownRows.join('\n');
}

module.exports = convertToMarkdown;
