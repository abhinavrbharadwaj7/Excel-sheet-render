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
  try {
    // Convert file to ArrayBuffer using a Promise wrapper
    const buffer = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });

    // Create Uint8Array from buffer to ensure data integrity
    const data = new Uint8Array(buffer);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data.buffer);

    return workbook.worksheets
      .filter(worksheet => {
        let hasData = false;
        worksheet.eachRow(row => {
          if (row.values && row.values.some(value => value !== null && value !== undefined)) {
            hasData = true;
          }
        });
        return hasData;
      })
      .map(worksheet => {
        // Find the actual used range by scanning all rows and columns
        let maxCol = 0;
        let maxRow = 0;
        worksheet.eachRow((row, rowNumber) => {
          maxRow = Math.max(maxRow, rowNumber);
          row.eachCell((cell, colNumber) => {
            maxCol = Math.max(maxCol, colNumber);
          });
        });

        // Process rows
        const rows = [];
        for (let rowNumber = 2; rowNumber <= maxRow; rowNumber++) { // Start from 2 to skip header
          const row = worksheet.getRow(rowNumber);
          const rowData = {
            id: rowNumber,
            __rowNum: rowNumber
          };

          let hasContent = false;
          for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
            const cell = row.getCell(colNumber);
            const value = processCellValue(cell);
            if (value !== '') {
              rowData[`col${colNumber}`] = { value };
              hasContent = true;
            }
          }

          if (hasContent) {
            rows.push(rowData);
          }
        }

        return {
          name: worksheet.name,
          headers: Array.from({ length: maxCol }, (_, i) => ({
            id: `col${i + 1}`
          })),
          rows
        };
      });
  } catch (error) {
    console.error('Excel processing error:', error);
    throw new Error('Failed to process Excel file. Please ensure the file is not corrupted.');
  }
};
