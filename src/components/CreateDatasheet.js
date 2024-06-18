import React, { useState } from 'react';
import ImportData from "./importData";
import DatasheetTemplate from './DatasheetTemplate';
import megamanLogo from "../images/logo_megaman.jpg"


const CreateDatasheet = () => {
    const [rowData, setRowData] = useState([]);
    const [mmCode, setMmCode] = useState("");
    const [logisticInfo, setLogisticInfo] = useState({});
    
    const handleDataImported = (data, code, logistic) => {
        setRowData(data);
        setMmCode(code);
        setLogisticInfo(logistic);
        //console.log(logisticInfo);
    };

    return (
        <div className="h-screen bg-gray-100">
            <div className="absolute top-0 left-0 p-4">
                <img src={megamanLogo} alt="Logo" className="h-16" />
            </div>
            <div className="flex flex-col h-screen pt-20">
                <div className="overflow-y-auto flex-grow m-2">
                    <ImportData onDataImported={handleDataImported} />
                </div>
                {rowData && (
                <div className="m-2 flex-grow overflow-y-auto">
                    <DatasheetTemplate data={rowData} code={mmCode} logistic={logisticInfo} />
                </div>
                )}
            </div>
        </div>
    );
}

export default CreateDatasheet;