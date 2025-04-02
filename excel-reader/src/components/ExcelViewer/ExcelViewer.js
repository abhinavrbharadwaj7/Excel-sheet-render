// components/ExcelViewer/ExcelViewer.js
import React, { useState, useCallback, useEffect } from 'react';
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
  const [columnWidths, setColumnWidths] = useState({});
  const [resizingColumn, setResizingColumn] = useState(null);

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
      // Validate file type
      const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
      ];
      
      if (!validTypes.includes(file.type)) {
        throw new Error('Please upload a valid Excel file (.xlsx or .xls)');
      }

      const sheets = await processExcelFile(file);
      if (sheets.length === 0) {
        throw new Error('No data found in the Excel file');
      }
      
      setWorkbook({
        fileName: file.name,
        sheets
      });
    } catch (err) {
      setError(err.message);
      console.error('File processing error:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleResizeStart = (field, e) => {
    setResizingColumn(field);
    e.preventDefault();
  };

  const handleResize = useCallback((e) => {
    if (!resizingColumn) return;
    
    setColumnWidths(prev => ({
      ...prev,
      [resizingColumn]: Math.max(150, e.clientX - e.target.getBoundingClientRect().left)
    }));
  }, [resizingColumn]);

  const handleResizeEnd = () => {
    setResizingColumn(null);
  };

  useEffect(() => {
    if (resizingColumn) {
      window.addEventListener('mousemove', handleResize);
      window.addEventListener('mouseup', handleResizeEnd);
      
      return () => {
        window.removeEventListener('mousemove', handleResize);
        window.removeEventListener('mouseup', handleResizeEnd);
      };
    }
  }, [resizingColumn, handleResize]);

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
            {String(params.row.__rowNum)}
          </div>
        )
      },
      ...headers.map((_, index) => ({
        field: `col${index + 1}`,
        headerName: String.fromCharCode(65 + index),
        flex: 1,
        minWidth: 120,
        sortable: false,
        filterable: false,
        renderCell: (params) => {
          const cell = params.value;
          if (!cell) return '';
          
          // Check if value is numeric
          const isNumber = !isNaN(cell.value) && cell.value !== '';
          
          return (
            <div 
              className={`cell-content ${cell.rowSpan || cell.colSpan ? 'merged-cell' : ''} ${isNumber ? 'number-cell' : ''}`}
              style={{
                gridRow: `span ${cell.rowSpan || 1}`,
                gridColumn: `span ${cell.colSpan || 1}`
              }}
            >
              {String(cell.value || '')}
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
              getRowHeight={() => 'auto'}
              getEstimatedRowHeight={() => 100}
              components={{
                Toolbar: null,
              }}
              sx={{
                '& .MuiDataGrid-main': {
                  overflow: 'visible'
                },
                '& .MuiDataGrid-cell': {
                  padding: 0,
                  border: '1px solid #e0e0e0'
                },
                '& .MuiDataGrid-columnHeader': {
                  padding: '12px',
                  backgroundColor: '#f8f9fa',
                  borderBottom: '2px solid #e0e0e0',
                  fontWeight: 600
                },
                '& .MuiDataGrid-virtualScroller': {
                  overflow: 'visible'
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