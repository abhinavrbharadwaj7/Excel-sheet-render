import ExcelJS from 'exceljs';

export const processCellValue = (cell) => {
  if (!cell) return '';
  
  // Handle formula results
  if (cell.type === ExcelJS.ValueType.Formula) {
    const result = cell.result;
    if (result instanceof Date) {
      return result.toLocaleDateString();
    }
    return result?.toString() || '';
  }
  
  // Handle direct cell values
  if (cell.value === null || cell.value === undefined) return '';
  
  // Handle dates
  if (cell.value instanceof Date) {
    return cell.value.toLocaleDateString();
  }
  
  // Handle rich text and other objects
  if (typeof cell.value === 'object') {
    if (cell.value.richText) {
      return cell.value.richText.map(t => t.text).join('');
    }
    if (cell.value.text) {
      return cell.value.text;
    }
    // Convert any other object to string safely
    return Object.prototype.toString.call(cell.value);
  }
  
  // Handle all other types
  return String(cell.value).trim();
};

export const processExcelFile = async (file) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  return workbook.worksheets.map(worksheet => {
    // Get max columns including all rows
    let maxCol = 0;
    worksheet.eachRow(row => {
      maxCol = Math.max(maxCol, row.actualCellCount);
    });

    // Process rows
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      const rowData = {
        id: rowNumber,
        __rowNum: rowNumber
      };

      // Process each cell in the row
      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        const cell = row.getCell(colNumber);
        const value = processCellValue(cell);
        if (value !== '') {
          rowData[`col${colNumber}`] = { value };
        }
      }

      rows.push(rowData);
    });

    return {
      name: worksheet.name,
      headers: Array.from({ length: maxCol }, (_, i) => ({
        id: `col${i + 1}`
      })),
      rows: rows.slice(1) // Skip header row
    };
  });
};
