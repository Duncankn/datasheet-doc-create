import { read, utils } from 'xlsx';
import { useEffect, useState } from 'react';

const DataImport = ({ onDataImported }) => {
    const [headerRow, setHeaderRow] = useState({});
    const [jsonRows, setJsonRows] = useState([]);
    const [rowData, setRowData] = useState({});
    const [selectedColumn, setSelectedColumn] = useState('');
    const [selectedData, setSelectedData] = useState('');
    const [rowNumber, setRowNumber] = useState('');
    const [selectedModelNum, setModelNum] = useState('');
    const [selectedColourTemp, setColourTemp] = useState('');
    const [selectedColour, setColour] = useState('');
    const [productCodeOptions, setProductCodeOptions] = useState([]);
    const [selectedOption, setSelectedOption] = useState('');
    const [logisticInfo, setLogisticInfo] = useState({});

    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (file) {
          const fileReader = new FileReader();
      
          fileReader.onload = (evt) => {
            const data = new Uint8Array(evt.target.result);
            const workbook = read(data, { type: 'array' });
      
            // Access the first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
    
            // Convert the worksheet to an array of objects
            const jsonData = utils.sheet_to_json(worksheet, { header: 1 });
            const mJsonData = jsonData.slice(4);
            const usefulJsonData = mJsonData.slice(0,1).concat(mJsonData.slice(3));
            const header = usefulJsonData[0];
            const firstRow = usefulJsonData[1]
            console.log(jsonData);
    
            setHeaderRow(header);
            setJsonRows(usefulJsonData);
            const singleRowObj = {};

            firstRow.forEach((value, cellIndex) => {
                singleRowObj[header[cellIndex]] = value;
            });
            setRowData(singleRowObj);

            console.log(header);
            console.log(singleRowObj);
          };
    
          fileReader.readAsArrayBuffer(file);
        }
    };

    const handleMmcodeMasterUpload = (e) => {
        const file = e.target.files[0];
        if (file) {
            const fileReader = new FileReader();
        
            fileReader.onload = (evt) => {
              const data = new Uint8Array(evt.target.result);
              const workbook = read(data, { type: 'array' });
        
              // Access the first sheet
              const sheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[sheetName];

              // Convert the worksheet to an array of objects
              const jsonData = utils.sheet_to_json(worksheet, { header: 1 });
              const mJsonData = jsonData.slice(7);

              // Find rows in the second file based on the selected model no.
              const filteredRows = mJsonData.filter((row) => row[1] === selectedModelNum);

              // Create a dropdown list for selection
              const options = filteredRows.map((row) => ({
                label: row[3]+"/"+row[7], // Assuming column 8 (index 7) has the product code
                value: row[8]
              }));

              console.log(selectedModelNum);
              console.log(options);

              // Update your state or render the dropdown list
              setProductCodeOptions(options);
            }

            fileReader.readAsArrayBuffer(file);
        }
    }

    const handleLogisticFileUpload = (e) => {
      const file = e.target.files[0];
      if (file) {
          const fileReader = new FileReader();
      
          fileReader.onload = (evt) => {
            const data = new Uint8Array(evt.target.result);
            const workbook = read(data, { type: 'array' });
      
            // Access the first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert the worksheet to an array of objects
            const jsonData = utils.sheet_to_json(worksheet, { header: 1 });

            // Find rows in the second file based on the selected model no.
            const filteredRows = jsonData.filter((row) => row[3] === selectedModelNum);
            const jsonLogisticData = filteredRows.map((row) => {
              return {
                package: row[0],
                series: row[1],
                brand: row[2],
                modelNum: row[3],
                YKmodelNum: row[4],
                power: row[5],
                unit: row[6],
                length: row[7],
                width: row[9],
                height: row[11],
                weight: row[12]
              }
            });
            setLogisticInfo(jsonLogisticData[0]);

            console.log(jsonLogisticData[0]);

          }

          fileReader.readAsArrayBuffer(file);
      }
  }

    const handleColumnChange = (e) => {
        setSelectedColumn(e.target.value);
    };

    const handleDataChange = (e) => {
        setSelectedData(e.target.value);
        setRowNumber(jsonRows.slice(1).indexOf(jsonRows.slice(1).find((row) => row[selectedColumn] === selectedData)) + 1)
    };

    useEffect(() => {
        const rowDataObj = {};

        if (selectedColumn) {
          setRowNumber(jsonRows.slice(1).indexOf(jsonRows.slice(1).find((row) => row[selectedColumn] === selectedData)) + 1);
          setJsonRows(jsonRows);
          
          //setIsDocumentReady(true);
        }
        if (rowNumber) {
            jsonRows[rowNumber].forEach((value, index) => {
                rowDataObj[jsonRows[0][index]] = value;
              });
            setRowData(rowDataObj);
        }
        //console.log(rowData);
        
    }, [selectedColumn, selectedData, jsonRows, rowNumber, rowData, onDataImported]);

    useEffect(() => {
        onDataImported(rowData, selectedOption, logisticInfo);
        setModelNum(rowData["Customer Model No.                (NEW ErP)"]);
        setColourTemp(rowData["Correlated Colour Temperature (CCT) in K"]);
        setColour(rowData["Fitting Color"]);
        console.log(selectedModelNum + selectedColourTemp + selectedColour);
    }, [rowData])

    return (
      <div className="flex flex-grow overflow-y-auto items-center justify-center bg-gray-100">
        <div className="flex flex-col bg-white shadow-lg rounded-lg p-8 w-full max-w-md">
          <h1 className="text-2xl font-bold mb-4">Luminaire Datasheet Generator</h1>
          <div className="mb-4">
            <label htmlFor="file-input" className="block text-gray-700 font-bold mb-2">
              Select a general data file:
            </label>
            <div className="relative">
              <input
                type="file"
                accept=".xlsx, .csv"
                onChange={handleFileUpload}
                id="file-input"
                className="text-sm text-grey-500
                  file:mr-5 file:py-2 file:px-6
                  file:rounded-full file:border-0
                  file:text-sm file:font-medium
                  file:bg-blue-50 file:text-blue-700
                  hover:file:cursor-pointer hover:file:bg-amber-50
                  hover:file:text-amber-700"
              />
            </div>
          </div>
          <div className="mb-4">
            <label htmlFor="column-select" className="block text-gray-700 font-bold mb-2">
              Select a column from general data:
            </label>
            <div className="relative">
              <select
                value={selectedColumn}
                onChange={handleColumnChange}
                id="column-select"
                className="w-full bg-white border border-gray-300 rounded-md py-2 pl-3 pr-10 text-sm leading-5 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 appearance-none"
              >
                <option value="">Select a column from general data</option>
                {Object.entries(headerRow).map(([column, headerCell]) => (
                  <option key={column} value={column}>
                    {headerCell}
                  </option>
                ))}
              </select>
            </div>
          </div>
          {selectedColumn && (
            <div className="mb-4">
              <label htmlFor="data-select" className="block text-gray-700 font-bold mb-2">
                Select a data:
              </label>
              <div className="relative">
                <select
                  value={selectedData}
                  onChange={handleDataChange}
                  id="data-select"
                  className="w-full bg-white border border-gray-300 rounded-md py-2 pl-3 pr-10 text-sm leading-5 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 appearance-none"
                >
                  <option value="">Select a data</option>
                  {jsonRows.slice(1).map((row) => (
                    <option key={row[selectedColumn]} value={row[selectedColumn]}>
                      {row[selectedColumn]}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          )}

          <p className="mb-4 text-gray-700">Row number: {rowNumber}</p>

          <div className="mb-4">
            <label htmlFor="file-input" className="block text-gray-700 font-bold mb-2">
              Select a range master file:
            </label>
            <div className="relative">
              <input
                type="file"
                accept=".xlsx, .csv"
                onChange={handleMmcodeMasterUpload}
                id="file-input"
                className="text-sm text-grey-500
                  file:mr-5 file:py-2 file:px-6
                  file:rounded-full file:border-0
                  file:text-sm file:font-medium
                  file:bg-blue-50 file:text-blue-700
                  hover:file:cursor-pointer hover:file:bg-amber-50
                  hover:file:text-amber-700"
              />
            </div>
            <select value={selectedOption} onChange={(e) => setSelectedOption(e.target.value)}>
              {productCodeOptions.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
            <p>Selected option: {selectedOption}</p>
          </div>
          <div className="mb-4">
            <label htmlFor="file-input" className="block text-gray-700 font-bold mb-2">
              Select the logistic information file:
            </label>
            <div className="relative">
              <input
                type="file"
                accept=".xlsx, .csv"
                onChange={handleLogisticFileUpload}
                id="file-input"
                className="text-sm text-grey-500
                  file:mr-5 file:py-2 file:px-6
                  file:rounded-full file:border-0
                  file:text-sm file:font-medium
                  file:bg-blue-50 file:text-blue-700
                  hover:file:cursor-pointer hover:file:bg-amber-50
                  hover:file:text-amber-700"
              />
            </div>
          </div>
        </div>
      </div>
    );
}

export default DataImport;