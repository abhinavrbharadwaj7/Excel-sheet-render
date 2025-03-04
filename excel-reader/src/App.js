import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import './App.css';

function App() {
  const [data, setData] = useState([]);
  const [sheets, setSheets] = useState({});
  const [currentSheet, setCurrentSheet] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const fileInput = useRef(null);

  const handleFileUpload = (e) => {
    const reader = new FileReader();
    reader.readAsBinaryString(e.target.files[0]);
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      
      // Handle multiple sheets
      const sheetsData = {};
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        sheetsData[sheetName] = XLSX.utils.sheet_to_json(sheet);
      });
      
      setSheets(sheetsData);
      setCurrentSheet(workbook.SheetNames[0]);
      setData(sheetsData[workbook.SheetNames[0]]);
      setIsModalOpen(true);
    };
  };

  const handleUploadClick = () => {
    fileInput.current.click();
  };

  const handleSheetChange = (sheetName) => {
    setCurrentSheet(sheetName);
    setData(sheets[sheetName]);
  };

  return (
    <div className="App">
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        style={{ display: 'none' }}
        ref={fileInput}
      />
      
      <button className="upload-btn" onClick={handleUploadClick}>
        Upload Excel File
      </button>

      {isModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              <h2>Excel Data</h2>
              <button onClick={() => setIsModalOpen(false)}>Close</button>
            </div>

            {Object.keys(sheets).length > 1 && (
              <div className="sheet-tabs">
                {Object.keys(sheets).map(sheetName => (
                  <button
                    key={sheetName}
                    className={`sheet-tab ${currentSheet === sheetName ? 'active' : ''}`}
                    onClick={() => handleSheetChange(sheetName)}
                  >
                    {sheetName}
                  </button>
                ))}
              </div>
            )}

            <div className="table-container">
              {data.length > 0 && (
                <table className="table">
                  <thead>
                    <tr>
                      {Object.keys(data[0]).map((key) => (
                        <th key={key}>{key}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {data.map((row, index) => (
                      <tr key={index}>
                        {Object.values(row).map((value, index) => (
                          <td key={index}>{value}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;