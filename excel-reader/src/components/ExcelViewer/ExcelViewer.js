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

  const handleTitleClick = () => {
    window.location.reload();
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
          let currentTable = { headers: [], rows: [] };
          let hasHeaders = false;

          worksheet.eachRow({ includeEmpty: true }, (row) => {
            const rowValues = row.values.slice(1).map(cell => {
              if (cell === null || cell === undefined) return '';
              
              // Handle different cell types
              if (typeof cell === 'object') {
                // Rich text
                if (cell.richText) {
                  return cell.richText.map(part => part.text).join('').trim();
                }
                // Hyperlink
                if (cell.hyperlink && cell.text) {
                  return cell.text.trim();
                }
                // Formula
                if (cell.formula) {
                  return cell.result !== undefined ? String(cell.result).trim() : '';
                }
                // Date
                if (cell instanceof Date) {
                  return cell.toLocaleDateString();
                }
                // SharedString or other object types
                if (cell.text) {
                  return cell.text.trim();
                }
                // If still an object, try to get a meaningful string representation
                return '';
              }
              
              // Handle numbers and other primitive types
              if (typeof cell === 'number') {
                return Number.isFinite(cell) ? cell : '';
              }
              
              return String(cell).trim();
            });

            if (!hasHeaders && rowValues.some(cell => cell)) {
              currentTable.headers = rowValues.map(h => h || 'Unnamed Column');
              hasHeaders = true;
              return;
            }

            if (hasHeaders && rowValues.some(cell => cell)) {
              const rowData = currentTable.headers.reduce((acc, header, index) => {
                const value = rowValues[index];
                
                // Convert to number if possible, otherwise keep as string
                acc[header] = !isNaN(value) && value !== '' ? Number(value) : value;
                return acc;
              }, {});
              
              currentTable.rows.push({
                ...rowData,
                id: `${row.number}-${Math.random().toString(36).slice(2, 9)}`
              });
            }
          });

          if (currentTable.headers.length > 0 && currentTable.rows.length > 0) {
            tables.push(currentTable);
          }

          sheets.push({
            name: worksheet.name,
            tables: tables
          });
        });

        setWorkbook({
          fileName: file.name,
          sheets: sheets.filter(sheet => sheet.tables.length > 0),
          created: new Date(),
          modified: new Date()
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
    return headers?.map(header => ({
      field: header,
      headerName: header || 'Unnamed Column',
      minWidth: 180,
      maxWidth: 400,
      flex: 1,
      headerClassName: 'data-grid-header',
      cellClassName: 'data-grid-cell',
      valueFormatter: (params) => {
        const value = params?.value;
        if (value === null || value === undefined || value === '') {
          return '';
        }
        if (typeof value === 'number') {
          // Format numbers with appropriate precision
          return Number.isInteger(value) ? 
            value.toLocaleString() : 
            value.toLocaleString(undefined, {
              minimumFractionDigits: 2,
              maximumFractionDigits: 6
            });
        }
        return String(value);
      },
      renderCell: (params) => {
        const value = params?.value;
        if (value === null || value === undefined || value === '') {
          return <span className="empty-cell">-</span>;
        }
        
        return (
          <div 
            className={typeof value === 'number' ? 'number-cell' : ''}
            style={{ 
              width: '100%',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap' 
            }}
            title={String(value)}
          >
            {value}
          </div>
        );
      }
    })) || [];
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

      {workbook?.sheets[0]?.tables?.map((table, tableIndex) => {
        const safeHeaders = table.headers?.filter(Boolean) || [];
        const safeRows = table.rows?.filter(r => 
          Object.values(r).some(v => v !== '' && v !== null)
        ) || [];

        return (
          <Box key={tableIndex} mt={4} height="70vh">
            <Typography variant="h6" mb={2} fontWeight="600">
              {workbook.sheets[0].name} - Table {tableIndex + 1} ({safeRows.length} rows)
            </Typography>
            
            {safeHeaders.length > 0 && safeRows.length > 0 ? (
              <DataGrid
                rows={safeRows}
                columns={generateColumns(safeHeaders)}
                pageSize={25}
                rowsPerPageOptions={[25, 50, 100]}
                components={{ Toolbar: GridToolbar }}
                density="comfortable"
                disableSelectionOnClick
                sx={{
                  '& .number-cell': {
                    fontFamily: 'Roboto Mono, monospace',
                    justifyContent: 'flex-end',
                    paddingRight: '16px !important'
                  },
                  '& .empty-cell': {
                    color: darkMode ? '#888' : '#666',
                    fontStyle: 'italic'
                  },
                  '& .MuiDataGrid-cellContent': {
                    fontSize: '0.9rem'
                  },
                  '& .MuiDataGrid-columnHeaderTitle': {
                    fontSize: '0.95rem',
                    fontWeight: 600
                  }
                }}
              />
            ) : (
              <Typography color="textSecondary">
                No displayable data found in this table
              </Typography>
            )}
          </Box>
        )
      })}
    </div>
  );
};

export default ExcelViewer;