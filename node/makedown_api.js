/**
 * 将表格数据转换为 Markdown 表格格式
 * @param {Array} rowData 表格数据
 * @returns {string} 转换后的 Markdown 表格
 */
const convertToMarkdownTable = (rowData) => {
  let markdownTable = '|';

  rowData.forEach(row => {
    row.values.forEach(cell => {
      if (cell.cell_value.link) {
        markdownTable += `[${cell.cell_value.link.text}](${cell.cell_value.link.url}) |`;
      } else {
        markdownTable += `${cell.cell_value.text} |`;
      }
    });

    markdownTable += '\n|';
  });

  return markdownTable.trim();
};

module.exports = { convertToMarkdownTable };