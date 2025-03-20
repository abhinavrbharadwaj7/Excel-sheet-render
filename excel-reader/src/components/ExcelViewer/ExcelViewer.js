// components/ExcelViewer/ExcelViewer.js
import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { DataGrid, GridToolbar } from '@mui/x-data-grid';
import { Tabs, Tab, Typography, CircularProgress, IconButton } from '@mui/material';
import Brightness4Icon from '@mui/icons-material/Brightness4';
import Brightness7Icon from '@mui/icons-material/Brightness7';
import './ExcelViewer.css';

const ExcelViewer = ({ darkMode, onToggleTheme }) => {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isDragActive, setIsDragActive] = useState(false);

  // File handling functions
  const handleDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
    setIsDragActive(false);
  }, []);

  const handleFileInput = (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
  };

  // Cell processing utilities
  const safeToString = (value) => {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if (value instanceof Date) return value.toLocaleDateString();
      if (value.richText) return value.richText.map(t => t.text).join('');
      if (value.hyperlink) return value.hyperlink.text || value.hyperlink;
      if (value.error) return `#${value.error}`;
      return Object.values(value).join(' ');
    }
    return String(value).trim();
  };

  const processCellValue = (cell) => {
    try {
      if (!cell) return '';
      if (cell.type === ExcelJS.ValueType.Formula) return safeToString(cell.result);
      if (cell.type === ExcelJS.ValueType.Error) return `#${cell.value?.error || 'ERROR'}`;
      if (cell.richText) return cell.richText.map(t => t.text).join('');
      if (cell.hyperlink) return cell.hyperlink.text || cell.hyperlink;
      return safeToString(cell.value);
    } catch (error) {
      console.warn('Cell processing error:', error);
      return '';
    }
  };

  const getExcelColumnName = (num) => {
    let columnName = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      columnName = String.fromCharCode(65 + remainder) + columnName;
      num = Math.floor((num - 1) / 26);
    }
    return columnName || 'A';
  };

  const processMergedCells = (worksheet) => {
    const mergedMap = new Map();
    try {
      worksheet.mergedCells?.forEach(merge => {
        const { top, left, bottom, right } = merge;
        const mainCell = worksheet.getCell(top, left);
        const value = processCellValue(mainCell);
        
        for (let row = top; row <= bottom; row++) {
          for (let col = left; col <= right; col++) {
            mergedMap.set(`${row}-${col}`, {
              value,
              isMain: row === top && col === left,
              colSpan: right - left + 1,
              rowSpan: bottom - top + 1
            });
          }
        }
      });
    } catch (error) {
      console.warn('Merged cell processing error:', error);
    }
    return mergedMap;
  };

  const processFile = async (file) => {
    setLoading(true);
    setError(null);
    
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const buffer = e.target.result;
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(buffer);

          const sheets = wb.worksheets.map(worksheet => {
            let maxCol = 0;
            worksheet.eachRow((row) => {
              maxCol = Math.max(maxCol, row._cells.length);
            });

            // Generate Excel-style column headers (A, B, C...)
            const headers = Array.from({ length: maxCol }, (_, i) => ({
              field: getExcelColumnName(i + 1),
              headerName: getExcelColumnName(i + 1)
            }));

            // Process rows with Excel-style row numbers
            const rows = [];
            worksheet.eachRow((row, rowNumber) => {
              if (rowNumber === 1) return;

              const rowData = { 
                id: `row-${rowNumber}`,
                rowNumber
              };

              let hasData = false;
              headers.forEach((header, colIndex) => {
                const cell = row.getCell(colIndex + 1);
                const value = processCellValue(cell);
                if (value !== '') {
                  rowData[header.field] = value;
                  hasData = true;
                }
              });

              if (hasData) {
                rows.push(rowData);
              }
            });

            return {
              name: worksheet.name,
              headers,
              rows
            };
          });

          setWorkbook({
            fileName: file.name,
            sheets: sheets.filter(sheet => sheet.rows.length > 0)
          });
        } catch (err) {
          setError(err.message || 'File processing failed');
        }
        setLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      setError(err.message || 'File read error');
      setLoading(false);
    }
  };

  const generateColumns = (headers) => {
    return [
      {
        field: 'rowNumber',
        headerName: '#',
        width: 50,
        sortable: false,
        filterable: false,
        headerClassName: 'bold-header row-number-header',
        renderCell: (params) => (
          <div className="row-number-cell">
            {params.value}
          </div>
        )
      },
      ...headers.map(header => ({
        field: header.field,
        headerName: header.headerName,
        minWidth: 150,
        flex: 1,
        headerClassName: 'bold-header',
        renderCell: (params) => (
          <div 
            className={`cell-content ${typeof params.value === 'number' ? 'number-cell' : ''}`}
            title={`${header.field}${params.row.rowNumber}: ${String(params.value || '')}`}
          >
            {params.value || <span className="empty-cell">-</span>}
            <span className="cell-address">{header.field}{params.row.rowNumber}</span>
          </div>
        )
      }))
    ];
  };

  return (
    <div className={`excel-viewer-container ${darkMode ? 'dark' : ''}`}>
      <div className="header-container">
        <div className="header-left">
          <Typography 
            variant="h5" 
            onClick={() => window.location.reload()}
            sx={{ fontWeight: 700, cursor: 'pointer' }}
          >
            Excel Viewer
          </Typography>
          <IconButton onClick={onToggleTheme}>
            {darkMode ? <Brightness7Icon /> : <Brightness4Icon />}
          </IconButton>
        </div>
        {workbook && (
          <Typography variant="subtitle1">
            {workbook.fileName} • {workbook.sheets.length} sheets
          </Typography>
        )}
      </div>

      {!workbook && (
        <div className={`drop-zone ${isDragActive ? 'active' : ''}`}
          onDragOver={(e) => e.preventDefault() || setIsDragActive(true)}
          onDragLeave={() => setIsDragActive(false)}
          onDrop={handleDrop}>
          <input type="file" accept=".xlsx,.xls" onChange={handleFileInput} hidden id="file-input" />
          <label htmlFor="file-input">
            <Typography variant="h6">
              {isDragActive ? 'Drop Excel file here' : 'Drag & Drop or Click to Upload'}
            </Typography>
            <Typography variant="body2" color="textSecondary">
              Supports .xlsx and .xls files
            </Typography>
          </label>
        </div>
      )}

      {loading && (
        <div className="loading-overlay">
          <CircularProgress size={60} thickness={4} />
          <Typography variant="h6">Processing file...</Typography>
        </div>
      )}

      {error && (
        <div className="error-message">
          <Typography color="error" variant="h6">{error}</Typography>
        </div>
      )}

      {workbook?.sheets?.[activeSheet] && (
        <div className="data-grid-container">
          <Tabs value={activeSheet} onChange={(_, v) => setActiveSheet(v)} variant="scrollable">
            {workbook.sheets.map((sheet, i) => (
              <Tab key={i} label={sheet.name} sx={{ textTransform: 'none' }} />
            ))}
          </Tabs>

          <DataGrid
            rows={workbook.sheets[activeSheet].rows}
            columns={generateColumns(
              workbook.sheets[activeSheet].headers,
              workbook.sheets[activeSheet].mergedMap
            )}
            autoHeight
            pageSize={10}
            components={{ Toolbar: GridToolbar }}
            disableSelectionOnClick
            sx={{
              '& .MuiDataGrid-cell': {
                display: 'flex',
                alignItems: 'center',
                position: 'relative',
                padding: '0 !important'
              },
              '& .MuiDataGrid-virtualScroller': {
                overflowX: 'auto',
                overflowY: 'scroll'
              }
            }}
          />
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;