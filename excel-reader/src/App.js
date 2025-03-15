import React, { useState, useMemo } from 'react';
import { StyledEngineProvider } from '@mui/material/styles';
import ExcelViewer from './components/ExcelViewer/ExcelViewer';
import { CssBaseline, ThemeProvider, createTheme } from '@mui/material';
import './DataGrid.css';
import './App.css';

function App() {
  const [mode, setMode] = useState('light');

  const theme = useMemo(() => createTheme({
    palette: {
      mode,
    },
    components: {
      MuiDataGrid: {
        styleOverrides: {
          root: {
            border: 'none',
          },
          columnHeaders: {
            backgroundColor: mode === 'dark' ? '#1e1e1e' : '#f3f7fa',
          },
        },
      },
    },
  }), [mode]);

  return (
    <StyledEngineProvider injectFirst>
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <div className="App">
          <ExcelViewer darkMode={mode === 'dark'} onToggleTheme={() => setMode(mode === 'light' ? 'dark' : 'light')} />
        </div>
      </ThemeProvider>
    </StyledEngineProvider>
  );
}

export default App;