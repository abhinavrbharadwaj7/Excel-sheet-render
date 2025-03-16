import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { DataGrid, GridToolbar } from '@mui/x-data-grid';
import { Box, Typography, CircularProgress, IconButton } from '@mui/material';
import Brightness4Icon from '@mui/icons-material/Brightness4';
import Brightness7Icon from '@mui/icons-material/Brightness7';
import './ExcelViewer.css';

const ExcelViewer = ({ darkMode, onToggleTheme }) => {
  const [workbook, setWorkbook] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isDragActive, setIsDragActive] = useState(false);

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
          let tableData = [];
          let headers = [];
          let inMaterialTable = false;

          worksheet.eachRow({ includeEmpty: true }, (row) => {
            const rowValues = row.values.slice(1).map(cell => {
              if (cell === null || cell === undefined) return '';
              if (typeof cell === 'object' && cell.text) return cell.text.trim();
              if (typeof cell === 'number') return cell.toString();
              if (cell instanceof Date) return cell.toISOString();
              return typeof cell === 'string' ? cell.trim() : '';
            });

            if (rowValues[0] === 'Component' && rowValues[1] === 'Substance') {
              inMaterialTable = true;
              headers = rowValues.filter(h => h);
              return;
            }

            if (inMaterialTable) {
              if (rowValues.every(cell => !cell)) {
                inMaterialTable = false;
                return;
              }

              const cleanRow = headers.reduce((acc, header, index) => {
                let value = rowValues[index] || '-';
                
                if (typeof value === 'string' && value.includes('e-')) {
                  value = Number(value).toFixed(8);
                }

                if (!isNaN(value) && value !== '-') {
                  value = Number(value);
                  if (value.toString().includes('.')) {
                    value = Number(value.toFixed(6));
                  }
                }

                acc[header] = value;
                return acc;
              }, {});

              if (Object.values(cleanRow).some(v => v !== '-')) {
                tableData.push({ ...cleanRow, id: tableData.length });
              }
            }
          });

          sheets.push({
            name: worksheet.name,
            table: {
              headers: headers,
              rows: tableData
            }
          });
        });

        setWorkbook({
          fileName: file.name,
          sheets: sheets,
          created: new Date(),
          modified: new Date()
        });
        setLoading(false);
      };
    } catch (err) {
      setError('Error processing file: ' + err.message);
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

  const columns = (headers) => {
    return headers.map(header => ({
      field: header,
      headerName: header,
      width: 200,
      headerClassName: 'data-grid-header',
      cellClassName: 'data-grid-cell',
      valueFormatter: ({ value }) => {
        if (typeof value === 'number') {
          return value.toLocaleString(undefined, {
            maximumFractionDigits: 6,
            useGrouping: false
          });
        }
        return value;
      },
      renderCell: ({ value }) => (
        <div className={typeof value === 'number' ? 'number-cell' : ''}>
          {value}
        </div>
      )
    }));
  };

  return (
    <div className={`excel-viewer-container ${darkMode ? 'dark' : ''}`}>
      <div className="header-container">
        <div className="header-left">
          <Typography variant="h5">
            Material Declaration Viewer
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
              Supports .xlsx and .xls files
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

      {workbook && workbook.sheets[0]?.table?.headers && (
        <Box height="75vh" mt={4}>
          <DataGrid
            rows={workbook.sheets[0].table.rows}
            columns={columns(workbook.sheets[0].table.headers)}
            pageSize={25}
            rowsPerPageOptions={[25, 50, 100]}
            components={{ Toolbar: GridToolbar }}
            density="compact"
            disableSelectionOnClick
            sx={{
              '& .number-cell': {
                fontFamily: 'Roboto Mono, monospace',
                justifyContent: 'flex-end',
                paddingRight: '16px !important'
              },
              '& .MuiDataGrid-virtualScroller': {
                overflowX: 'auto'
              }
            }}
          />
        </Box>
      )}
    </div>
  );
};

export default ExcelViewer;