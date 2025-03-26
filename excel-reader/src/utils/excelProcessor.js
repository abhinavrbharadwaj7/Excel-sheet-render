const ExcelJS = require('exceljs'); // Changed to require syntax

export const processCellValue = (cell) => {
  if (!cell) return '';
  
  try {
    if (cell.type === ExcelJS.ValueType.Formula) {
      const result = cell.result;
      return result ? String(result) : '';  // Ensure string conversion
    }
    
    if (cell.value === null || cell.value === undefined) return '';
    
    if (cell.value instanceof Date) {
      return cell.value.toLocaleDateString();
    }
    
    if (typeof cell.value === 'object') {
      if (cell.value.richText) {
        return cell.value.richText.map(t => t.text).join('');
      }
      if (cell.value.text) {
        return cell.value.text;
      }
      if (cell.value.error) {
        return '#ERROR';  // Handle error values
      }
      return '';
    }
    
    return String(cell.value).trim();  // Ensure string conversion
  } catch (error) {
    console.warn('Error processing cell:', error);
    return '';
  }
};

export const processExcelFile = async (file) => {
  try {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    return workbook.worksheets.map(worksheet => {
      const mergedCells = new Map();
      
      // Process merged cells using _merges object
      if (worksheet.mergeCells && worksheet._merges) {
        Object.values(worksheet._merges).forEach(merge => {
          const { top, left, bottom, right } = merge;
          const mainCell = worksheet.getCell(top, left);
          const value = processCellValue(mainCell);
          
          for (let row = top; row <= bottom; row++) {
            for (let col = left; col <= right; col++) {
              mergedCells.set(`${row}-${col}`, {
                value,
                isMain: row === top && col === left,
                span: {
                  rowSpan: bottom - top + 1,
                  colSpan: right - left + 1
                }
              });
            }
          }
        });
      }

      // Get max columns
      let maxCol = 0;
      worksheet.eachRow(row => {
        row.eachCell((cell, col) => {
          maxCol = Math.max(maxCol, col);
        });
      });

      // Process rows with merged cell handling
      const rows = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header

        const rowData = {
          id: rowNumber,
          __rowNum: rowNumber
        };

        for (let col = 1; col <= maxCol; col++) {
          const cellKey = `${rowNumber}-${col}`;
          const mergedInfo = mergedCells.get(cellKey);
          
          if (mergedInfo) {
            if (mergedInfo.isMain) {
              rowData[`col${col}`] = {
                value: mergedInfo.value,
                ...mergedInfo.span
              };
            }
          } else {
            const value = processCellValue(row.getCell(col));
            if (value) {
              rowData[`col${col}`] = { value };
            }
          }
        }

        if (Object.keys(rowData).length > 2) { // More than just id and __rowNum
          rows.push(rowData);
        }
      });

      return {
        name: worksheet.name,
        headers: Array.from({ length: maxCol }, (_, i) => ({
          id: `col${i + 1}`,
          label: String.fromCharCode(65 + i)
        })),
        rows
      };
    });
  } catch (error) {
    console.error('Excel processing error:', error);
    throw new Error('Failed to process Excel file. Please try again.');
  }
};
