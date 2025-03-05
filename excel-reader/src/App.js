import React from 'react';
import { StyledEngineProvider } from '@mui/material/styles';
import ExcelViewer from './components/ExcelViewer/ExcelViewer';
import { CssBaseline, ThemeProvider, createTheme } from '@mui/material';
import './DataGrid.css';
import './App.css';

const theme = createTheme({
  components: {
    MuiDataGrid: {
      styleOverrides: {
        root: {
          border: 'none',
        },
        columnHeaders: {
          backgroundColor: '#f3f7fa',
        },
      },
    },
  },
});

function App() {
  return (
    <StyledEngineProvider injectFirst>
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <div className="App">
          <ExcelViewer />
        </div>
      </ThemeProvider>
    </StyledEngineProvider>
  );
}

export default App;