const XLSX = require('xlsx-populate');

// Get the value of a specific cell of Excel
const getSingleCellValue = async (path = '', sheetNumber = Number, col = '', row = '') => {

    const data = await XLSX.fromFileAsync(path);

    const res = await data.sheet(sheetNumber).cell(`${col}${row}`).value();

    console.log(res);

    return res;

}

// Set a value in a specific cell of Excel
const setSingleCellValue = async (path = '', sheetNumber = Number, col = '', row = '', value = '') => {

    const data = await XLSX.fromFileAsync(path);

    const res = await data.sheet(sheetNumber).cell(`${col}${row}`).value(value);
    
    data.toFileAsync(path);

    console.log('value writed');    

}

// Get multiple cell values from Excel
const getMultipleCellValue = async (path = '', sheetNumber = Number, startCell = '', endCell = '') => {

    const data = await XLSX.fromFileAsync(path);

    const res = await data.sheet(sheetNumber).range(`${startCell}:${endCell}`).value();

    console.log(res);
    
    return res;

}

// Set multiple cell values
const SetMultipleCellValue = async (path = '', sheetNumber = Number, startCell = '', endCell = '', copyCell = '', values = []) => {

    const value = [values];

    const data = await XLSX.fromFileAsync(path);

    if (startCell && endCell && copyCell == null) {
        const res = await data.sheet(sheetNumber).cell(`${startCell}:${endCell}`).value(value);
    }

    if (copyCell) {
        const res = await data.sheet(sheetNumber).cell(copyCell).value(value);
    }

    data.toFileAsync(path);

    console.log('values writed');

}

// Copy multiple cell values
const CopyMultipleCellValue = async (path = '', sheetNumber = Number, startCell = '', endCell = '', copyCell = '') => {

    const values = await getMultipleCellValue(path, 0, startCell, endCell);

    const data = await XLSX.fromFileAsync(path);

    const res = await data.sheet(sheetNumber).cell(copyCell).value(values);

    data.toFileAsync(path);

    console.log('values writed');

}

////////////////////////////////////////////////////////

/**
 * 
 * WriteToExcel(A1,A20)
WriteToExcel(A1, F1)
WriteToExcel(A1, V) te escribe desde A1 hasta infinito a lo vertical
WriteToExcel (A1, H)
Te escribe desde A1 hasta lo infinito en lo horizontal
 * 
 */

const WriteToExcel = async (path = '', sheetNumber = Number , startCell = '', endCell = '', value = [], horizontal = true) => {
    
    try {
        
        // Formatting the VALUE array
        const values = [value];

        // Get the data from ALL Excel
        const data = await XLSX.fromFileAsync(path);

        if (startCell && endCell) {
            // Set the data in specific RANGE position
            await data.sheet(sheetNumber).range(`${startCell}:${endCell}`).value(values);
        }
        else if (startCell && endCell == null) {
            // Set the data in specific position to infinite (horizontal default)
            if (horizontal == true) {
                await data.sheet(sheetNumber).cell(startCell).value(values);
            }
            else {
                const verticalArray = value.map(item => [item]);
                await data.sheet(sheetNumber).cell(startCell).value(verticalArray);
            }
        }
        else {
            console.error('An error has happened: ', error);
            return {
                res: "bad"
            }
        }

        data.toFileAsync(path);

        console.log('Values writed');
        
        return {
            res: 'ok'
        }

    } catch (error) {
        
        console.error('An error has happened: ', error);

        return {
            res: error
        }

    }

}

////////////////////////////////////////////////////////

module.exports = { getMultipleCellValue, CopyMultipleCellValue, SetMultipleCellValue, getSingleCellValue, setSingleCellValue, WriteToExcel }