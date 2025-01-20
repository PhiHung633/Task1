import axios from "axios";
import XLSX from "xlsx"

const processExcel = async () => {
    try {
        const url = "https://go.microsoft.com/fwlink/?LinkID=521962";

        const response = await axios.get(url, {responseType: 'arraybuffer'});
        const workbook = XLSX.read(response.data, {type: 'buffer'});

        const sheetName = workbook.SheetNames[0];
        const workSheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(workSheet);

        const normilizedData = jsonData.map(row => {
            const normilizedRow = {};
            for(const key in row){
                const trimmedKey = key.trim();
                normilizedRow[trimmedKey]=row[key];
            }
            return normilizedRow;
        })
        const filteredData = normilizedData.filter(row => {
            const salesValue =row['Sales'];
            if(!salesValue) return false;
            const numericValue = parseFloat(salesValue.toString().replace(/[$,]/g, ''));
            return numericValue > 50000;
        });
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Filtered Data");

        const outputFile = "filtered_sales.xlsx";
        XLSX.writeFile(newWorkbook, outputFile)
        console.log(`File save success: ${outputFile}`);
    } catch (error) {
        console.error("Error: ", error.message)
    }
};

processExcel();