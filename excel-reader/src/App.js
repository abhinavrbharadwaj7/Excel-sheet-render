import React, { useState, useMemo } from 'react';
import { StyledEngineProvider } from '@mui/material/styles';
import ExcelViewer from './components/ExcelViewer/ExcelViewer';
import { CssBaseline, ThemeProvider, createTheme } from '@mui/material';
import './App.css';

function App() {
  const [mode, setMode] = useState('light');

  const theme = useMemo(() => createTheme({
    palette: {
      mode,
      primary: { main: '#6366f1' },
      secondary: { main: '#a855f7' },
      background: {
        default: mode === 'dark' ? '#121212' : '#f3f4f6',
      },
    },
    components: {
      MuiDataGrid: {
        styleOverrides: {
          root: {
            border: 'none',
            '& .MuiDataGrid-columnHeaders': {
              backgroundColor: mode === 'dark' ? '#2d2d2d' : '#f8f9fa',
            },
            '& .MuiDataGrid-cell': {
              borderBottomColor: mode === 'dark' ? 'rgba(255,255,255,0.1)' : '#f0f0f0'
            }
          }
        }
      }
    }
  }), [mode]);

  return (
    <StyledEngineProvider injectFirst>
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <ExcelViewer 
          darkMode={mode === 'dark'} 
          onToggleTheme={() => setMode(prev => prev === 'light' ? 'dark' : 'light')}
        />
      </ThemeProvider>
    </StyledEngineProvider>
  );
}

export default App;