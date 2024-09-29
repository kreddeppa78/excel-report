import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const ExcelProcessor1 = () => {
  const [data, setData] = useState([]);
  const priorityActivityCodes = ['ACT123', 'ACT456']; // Replace with your priority activity codes

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      processExcelData(worksheet);
    };
    reader.readAsBinaryString(file);
  };

  const processExcelData = (data) => {
    // Merge Employee and Activity Code into one column
    const processedData = data.map((row) => ({
      Time: row.Time,
      Employee_Activity: `${row.Employee} - ${row['Activity Code']}`,
      Comments: row.Comments,
     // ActivityCode: row['Activity Code'],
    }));

    // Sort the data based on priority activity codes
    processedData.sort((a, b) => {
      const aPriority = priorityActivityCodes.includes(a.ActivityCode) ? -1 : 1;
      const bPriority = priorityActivityCodes.includes(b.ActivityCode) ? -1 : 1;
      return aPriority - bPriority;
    });

    setData(processedData);
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Apply header styles
    const headerStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "4F81BD" } },
      alignment: { horizontal: "center", vertical: "center" }
    };

    // Apply header styles to the first row
    for (let cell in worksheet) {
      if (cell[0] === '!') continue; // Skip metadata
      const cellRef = XLSX.utils.decode_cell(cell);
      if (cellRef.r === 0) {
        worksheet[cell].s = headerStyle; // Apply header style
      }
    }

    // Highlight priority rows
    const priorityStyle = {
      fill: { fgColor: { rgb: "FFC7CE" } }
    };

    for (let i = 1; i < data.length + 1; i++) {
      const activityCode = data[i - 1].ActivityCode;
      if (priorityActivityCodes.includes(activityCode)) {
        for (let j = 0; j < Object.keys(data[0]).length; j++) {
          const cellRef = XLSX.utils.encode_cell({ r: i, c: j });
          worksheet[cellRef].s = priorityStyle; // Apply priority style
        }
      }
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Processed Data');
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, 'processed_data.xlsx');
  };

  return (
    <div style={{ padding: '20px', textAlign: 'center', fontFamily: 'Arial, sans-serif' }}>
         <h1 style={{ color: '#4F81BD' }}>WEB APP FOR REPORTS GENERATION</h1>
  <h2 style={{ color: '#4F81BD' }}> Please Upload your Excel file</h2>
  <input
    type="file"
    accept=".xlsx, .xls"
    onChange={handleFileUpload}
    style={{
      margin: '20px 0',
      padding: '10px',
      border: '2px solid #4F81BD',
      borderRadius: '5px',
      fontSize: '16px',
    }}
  />
  <br></br>
  {data.length > 0 && (
    <button
      onClick={exportToExcel}
      style={{
        backgroundColor: '#4F81BD',
        color: 'white',
        padding: '10px 20px',
        fontSize: '16px',
        border: 'none',
        borderRadius: '5px',
        cursor: 'pointer',
      }}
    >
       
      Download Processed Excel
    </button>
  )}
</div>

  );
};

export default ExcelProcessor1;
