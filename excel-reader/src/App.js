// App.js
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
        paper: mode === 'dark' ? '#1e1e1e' : '#ffffff',
      },
    },
    components: {
      MuiDataGrid: {
        styleOverrides: {
          root: ({ theme }) => ({
            border: 'none',
            backgroundColor: theme.palette.background.paper,
            '& .MuiDataGrid-columnHeaders': {
              backgroundColor: theme.palette.mode === 'dark' ? 
                '#2d2d2d' : 
                '#f8f9fa',
              color: theme.palette.text.primary,
            },
            '& .MuiDataGrid-cell': {
              borderBottomColor: theme.palette.mode === 'dark' ? 
                'rgba(255,255,255,0.1)' : 
                '#f0f0f0'
            }
          })
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