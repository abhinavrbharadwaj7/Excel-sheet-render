// components/ExcelViewer/ExcelViewer.js
import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { DataGrid, GridToolbar } from '@mui/x-data-grid';
import { Tabs, Tab, Typography, CircularProgress, IconButton } from '@mui/material';
import Brightness4Icon from '@mui/icons-material/Brightness4';
import Brightness7Icon from '@mui/icons-material/Brightness7';
import './ExcelViewer.css';
import { processExcelFile } from '../../utils/excelProcessor';

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

  const processFile = async (file) => {
    setLoading(true);
    setError(null);
    
    try {
      const sheets = await processExcelFile(file);
      setWorkbook({
        fileName: file.name,
        sheets: sheets.filter(s => s.rows.length > 0)
      });
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateColumns = (headers) => {
    return [
      {
        field: '__rowNum',
        headerName: '',
        width: 50,
        sortable: false,
        filterable: false,
        renderCell: (params) => (
          <div className="row-number-cell">
            {params.row.__rowNum}
          </div>
        )
      },
      ...headers.map((header, index) => ({
        field: `col${index + 1}`,
        headerName: String.fromCharCode(65 + index), // Use A, B, C... for headers
        flex: 1,
        minWidth: 150,
        sortable: false, // Disable sorting
        filterable: false, // Disable filtering
        renderCell: (params) => {
          const cell = params.value || {};
          const style = cell.style || {};
          
          const cellClasses = [
            'cell-content',
            style.bold ? 'bold' : '',
            style.align ? `align-${style.align}` : '',
            params.row.isHeader ? 'header-cell' : '',
            params.row.sectionIndex > 0 && 
            params.api.getRowIndex(params.row.id) === 0 ? 'section-start' : ''
          ].filter(Boolean).join(' ');

          return (
            <div 
              className={cellClasses}
              style={{
                ...(style.background && {
                  backgroundColor: `#${style.background.substring(2)}`
                })
              }}
            >
              {cell.value || ''}
            </div>
          );
        }
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
          <Tabs 
            value={activeSheet} 
            onChange={(_, v) => setActiveSheet(v)} 
            variant="scrollable"
            sx={{ borderBottom: 1, borderColor: 'divider' }}
          >
            {workbook.sheets.map((sheet, i) => (
              <Tab key={i} label={sheet.name} sx={{ textTransform: 'none' }} />
            ))}
          </Tabs>

          <div className="table-container">
            <DataGrid
              rows={workbook.sheets[activeSheet].rows}
              columns={generateColumns(workbook.sheets[activeSheet].headers)}
              getRowId={(row) => row.id}
              autoHeight
              hideFooter
              disableColumnMenu
              disableSelectionOnClick
              components={{
                Toolbar: null, // Remove the toolbar completely
              }}
              sx={{
                '& .MuiDataGrid-cell': {
                  padding: '0 8px',
                  maxHeight: '200px',
                  overflow: 'hidden'
                }
              }}
            />
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;