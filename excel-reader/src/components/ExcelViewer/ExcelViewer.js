// components/ExcelViewer/ExcelViewer.js
import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { DataGrid, GridToolbar } from '@mui/x-data-grid';
import { Tabs, Tab, Box, Typography, CircularProgress, IconButton } from '@mui/material';
import Brightness4Icon from '@mui/icons-material/Brightness4';
import Brightness7Icon from '@mui/icons-material/Brightness7';
import './ExcelViewer.css';

const ExcelViewer = ({ darkMode, onToggleTheme }) => {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isDragActive, setIsDragActive] = useState(false);

  const processCellValue = (cell) => {
    try {
      if (!cell) return '';
      
      // Get actual cell value
      const value = cell.type === ExcelJS.ValueType.Formula ? cell.result : cell.value;
      
      if (value === null || value === undefined) return '';
      
      // Handle different types
      if (value instanceof Date) return value.toLocaleDateString();
      if (typeof value === 'object' && value.richText) {
        return value.richText.map(rt => rt.text || '').join('');
      }
      if (typeof value === 'number') return value;
      
      return String(value).trim();
    } catch (err) {
      console.warn('Error processing cell:', err);
      return '';
    }
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
            // Get max column count by checking each row
            let maxCol = 0;
            worksheet.eachRow((row) => {
              maxCol = Math.max(maxCol, row._cells.length);
            });

            // Process headers (first row)
            const headers = [];
            const firstRow = worksheet.getRow(1);
            for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
              const cell = firstRow.getCell(colNumber);
              const headerValue = processCellValue(cell) || `Column ${colNumber}`;
              headers.push(headerValue);
            }

            // Process data rows
            const rows = [];
            worksheet.eachRow((row, rowNumber) => {
              if (rowNumber === 1) return; // Skip header row

              const rowData = { id: `row-${rowNumber}` };
              let hasData = false;

              // Process each column in the row
              headers.forEach((header, index) => {
                const cell = row.getCell(index + 1);
                const value = processCellValue(cell);
                rowData[header] = value;
                if (value !== '') hasData = true;
              });

              if (hasData) {
                rows.push(rowData);
              }
            });

            return {
              name: worksheet.name,
              tables: [{
                headers: headers,
                rows: rows
              }]
            };
          }).filter(sheet => sheet.tables[0].headers.length > 0 && sheet.tables[0].rows.length > 0);

          if (!sheets.length) {
            throw new Error('No valid data found in the Excel file');
          }

          setWorkbook({
            fileName: file.name,
            sheets
          });
        } catch (err) {
          console.error('Processing error:', err);
          setError(err.message);
        }
        setLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error('File error:', err);
      setError(err.message);
      setLoading(false);
    }
  };

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

  const generateColumns = (headers) => {
    return headers.map(header => ({
      field: header,
      headerName: header,
      minWidth: 180,
      flex: 1,
      renderCell: (params) => (
        <div 
          className={`cell-content ${typeof params.value === 'number' ? 'number-cell' : ''}`}
          title={String(params.value)}
        >
          {params.value || <span className="empty-cell">-</span>}
        </div>
      )
    }));
  };

  return (
    <div className={`excel-viewer-container ${darkMode ? 'dark' : ''}`}>
      <div className="header-container">
        <div className="header-left">
          <Typography 
            variant="h5" 
            onClick={() => window.location.reload()}
            sx={{ 
              cursor: 'pointer',
              '&:hover': { opacity: 0.8 },
              fontWeight: 700
            }}
          >
            Excel Viewer
          </Typography>
          <IconButton onClick={onToggleTheme} color="inherit">
            {darkMode ? <Brightness7Icon /> : <Brightness4Icon />}
          </IconButton>
        </div>
        {workbook && (
          <Typography variant="subtitle1" sx={{ opacity: 0.8 }}>
            {workbook.fileName} • {workbook.sheets.length} sheets
          </Typography>
        )}
      </div>

      {!workbook && (
        <div 
          className={`drop-zone ${isDragActive ? 'active' : ''}`}
          onDragOver={(e) => e.preventDefault() || setIsDragActive(true)}
          onDragLeave={() => setIsDragActive(false)}
          onDrop={handleDrop}
        >
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileInput}
            hidden
            id="file-input"
          />
          <label htmlFor="file-input">
            <Typography variant="h6">
              {isDragActive ? 'Drop Excel file here' : 'Drag & Drop or Click to Upload'}
            </Typography>
            <Typography variant="body2" color="textSecondary" mt={1}>
              Supports .xlsx and .xls files
            </Typography>
          </label>
        </div>
      )}

      {loading && (
        <div className="loading-overlay">
          <CircularProgress size={60} thickness={4} />
          <Typography variant="h6" mt={2}>
            Loading {workbook?.fileName || 'file'}...
          </Typography>
        </div>
      )}

      {error && (
        <div className="error-message">
          <Typography color="error" variant="h6">
            Error: {error}
          </Typography>
        </div>
      )}

      {workbook && workbook.sheets.length > 0 && (
        <div className="data-grid-container">
          <Tabs
            value={activeSheet}
            onChange={(e, newValue) => setActiveSheet(newValue)}
            variant="scrollable"
          >
            {workbook.sheets.map((sheet, index) => (
              <Tab 
                key={index} 
                label={sheet.name} 
                sx={{ textTransform: 'none' }}
              />
            ))}
          </Tabs>

          {workbook.sheets[activeSheet]?.tables?.map((table, tableIndex) => (
            <div key={tableIndex} className="table-container">
              <DataGrid
                rows={table.rows}
                columns={generateColumns(table.headers)}
                autoHeight
                pageSize={10}
                rowsPerPageOptions={[10, 25, 50]}
                components={{ Toolbar: GridToolbar }}
                disableSelectionOnClick
                sx={{
                  mt: 2,
                  '& .MuiDataGrid-columnHeaders': {
                    backgroundColor: darkMode ? '#2d2d2d' : '#f8f9fa',
                  }
                }}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;