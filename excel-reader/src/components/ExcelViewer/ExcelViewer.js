import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { DataGrid } from '@mui/x-data-grid';
import { Tabs, Tab, Box, Typography, CircularProgress } from '@mui/material';
import './ExcelViewer.css';

const ExcelViewer = () => {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheet, setActiveSheet] = useState(0);
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
        wb.eachSheet((worksheet, sheetId) => {
          const rows = [];
          worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            rows.push(row.values.slice(1));
          });
          sheets.push({
            name: worksheet.name,
            rows: rows,
            rowCount: worksheet.rowCount,
            columnCount: worksheet.columnCount
          });
        });
        
        setWorkbook({
          fileName: file.name,
          sheets: sheets,
          created: wb.created,
          modified: wb.modified
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

  const generateColumns = (rows) => {
    if (!rows || rows.length === 0) return [];
    return rows[0].map((_, index) => ({
      field: `col${index}`,
      headerName: String.fromCharCode(65 + index),
      width: 120,
      editable: false
    }));
  };

  const generateRows = (rows) => {
    return rows.map((row, index) => ({
      id: index,
      ...row.reduce((acc, val, idx) => {
        acc[`col${idx}`] = val?.toString() || '';
        return acc;
      }, {})
    }));
  };

  return (
    <div className="excel-viewer-container">
      <div className="header-container">
        <Typography variant="h5">Excel Viewer</Typography>
        {workbook && (
          <Typography variant="subtitle1">
            {workbook.fileName} • {workbook.sheets.length} sheets
          </Typography>
        )}
      </div>

      {!workbook && (
        <div 
          className={`drop-zone ${isDragActive ? 'active' : ''}`}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragActive(true);
          }}
          onDragLeave={() => setIsDragActive(false)}
          onDrop={handleDrop}
        >
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileInput}
            style={{ display: 'none' }}
            id="file-input"
          />
          <label htmlFor="file-input">
            <Typography variant="h6">
              {isDragActive ? 'Drop Excel file here' : 'Drag & Drop or Click to Upload'}
            </Typography>
            <Typography variant="body2" color="textSecondary" style={{ marginTop: 10 }}>
              Supports .xlsx and .xls files
            </Typography>
          </label>
        </div>
      )}

      {loading && (
        <div className="loading-container">
          <CircularProgress />
          <Typography variant="body1" style={{ marginLeft: 10 }}>
            Processing file...
          </Typography>
        </div>
      )}

      {error && (
        <Typography color="error" align="center" my={4}>
          {error}
        </Typography>
      )}

      {workbook && (
        <div>
          <Tabs
            value={activeSheet}
            onChange={(e, newValue) => setActiveSheet(newValue)}
            variant="scrollable"
            scrollButtons="auto"
            className="sheet-tabs"
          >
            {workbook.sheets.map((sheet, index) => (
              <Tab 
                key={index}
                label={`${sheet.name} (${sheet.rowCount}x${sheet.columnCount})`}
              />
            ))}
          </Tabs>

          <Box mt={2} height={600}>
            <DataGrid
              rows={generateRows(workbook.sheets[activeSheet].rows)}
              columns={generateColumns(workbook.sheets[activeSheet].rows)}
              pageSize={50}
              rowsPerPageOptions={[50, 100, 200]}
              checkboxSelection={false}
              disableSelectionOnClick
              loading={loading}
            />
          </Box>

          <div className="file-info">
            <Typography variant="body2">
              File Info: {workbook.fileName} • Created: {workbook.created?.toLocaleDateString()} • 
              Modified: {workbook.modified?.toLocaleDateString()}
            </Typography>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;