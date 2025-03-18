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

  const handleTitleClick = () => {
    window.location.reload();
  };

  const processCellValue = (cell) => {
    try {
      if (cell === null || cell === undefined) return '';
      if (typeof cell === 'object') {
        if (cell.text) return cell.text.trim();
        if (cell.result) return cell.result.toString();
        if (cell.hyperlink) return cell.hyperlink;
        if (cell.formula) return `=${cell.formula}`;
        return '';
      }
      if (cell instanceof Date) return cell.toLocaleDateString();
      if (typeof cell === 'number') return cell;
      return String(cell).trim();
    } catch (error) {
      console.warn('Error processing cell:', error);
      return '';
    }
  };

  const processFile = async (file) => {
    setLoading(true);
    setError(null);
    
    try {
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      
      reader.onload = async () => {
        const buffer = reader.result;
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);

        const sheets = [];
        wb.eachSheet((worksheet) => {
          const tables = [];
          let currentTable = null;
          let headers = [];
          let inTable = false;

          worksheet.eachRow({ includeEmpty: true }, (row) => {
            const rowValues = row.values.slice(1).map(processCellValue);
            const isEmptyRow = rowValues.every(cell => !cell);

            if (!inTable && !isEmptyRow) {
              inTable = true;
              headers = rowValues;
              currentTable = {
                headers: headers.filter(h => h),
                rows: []
              };
            } else if (inTable) {
              if (isEmptyRow) {
                if (currentTable.rows.length > 0) {
                  tables.push(currentTable);
                }
                inTable = false;
                currentTable = null;
              } else {
                const rowData = headers.reduce((acc, header, index) => {
                  acc[header] = rowValues[index] || '';
                  return acc;
                }, {});
                currentTable.rows.push({
                  ...rowData,
                  id: `${row.number}-${Math.random().toString(36).slice(2, 9)}`
                });
              }
            }
          });

          if (currentTable?.rows.length > 0) {
            tables.push(currentTable);
          }

          sheets.push({
            name: worksheet.name,
            tables: tables.filter(t => t.headers.length > 0 && t.rows.length > 0)
          });
        });

        setWorkbook({
          fileName: file.name,
          sheets: sheets.filter(sheet => sheet.tables.length > 0),
          created: new Date()
        });
        setLoading(false);
      };
    } catch (err) {
      setError(`Error processing file: ${err.message}`);
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
            onClick={handleTitleClick}
            sx={{ 
              cursor: 'pointer',
              '&:hover': { opacity: 0.8 }
            }}
          >
            Excel Viewer
          </Typography>
          <IconButton onClick={onToggleTheme} color="inherit">
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
              Supports all Excel files
            </Typography>
          </label>
        </div>
      )}

      {loading && (
        <div className="loading-container">
          <CircularProgress />
          <Typography variant="body1" ml={2}>
            Processing file...
          </Typography>
        </div>
      )}

      {error && (
        <Typography color="error" textAlign="center" my={4}>
          {error}
        </Typography>
      )}

      {workbook && (
        <div>
          <Tabs
            value={activeSheet}
            onChange={(e, newValue) => setActiveSheet(newValue)}
            variant="scrollable"
            sx={{ mb: 2 }}
          >
            {workbook.sheets.map((sheet, index) => (
              <Tab key={index} label={`${sheet.name} (${sheet.tables.length})`} />
            ))}
          </Tabs>

          {workbook.sheets[activeSheet]?.tables?.map((table, tableIndex) => (
            <Box key={tableIndex} mt={4}>
              <Typography variant="h6" mb={2}>
                Table {tableIndex + 1} ({table.rows.length} rows)
              </Typography>
              <DataGrid
                rows={table.rows}
                columns={generateColumns(table.headers)}
                autoHeight
                pageSize={25}
                rowsPerPageOptions={[25, 50, 100]}
                components={{ Toolbar: GridToolbar }}
                density="comfortable"
                disableSelectionOnClick
                sx={{
                  '& .cell-content': {
                    width: '100%',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis'
                  }
                }}
              />
            </Box>
          ))}
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;