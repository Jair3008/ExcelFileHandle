const xlsx = require('xlsx-populate');

const writeToExcel = async (path = '', sheetNumber = 0 , value = [], config = {
    vertical: false, // Default horizontal
    range: '', // If you need write in multiples cells
    cell: '' // If you need write in single cell
}) => {

    // If config values aren't string type
    if (config.cell) {
        if (typeof (config.cell) !== 'string') {
            console.error('cagaste hubo un error 2');
            return {
                res: 'error'
            }
        }
    }

    if (config.range) {
        if (typeof (config.range) !== 'string') {
            console.error('cagaste hubo un error 2');
            return {
                res: 'error'
            }
        }
    }

    // If the value parameter no exist or is null
    if (!value || value.length <= 0) {
        console.error('cagaste hubo un error 2');
        return {
            res: 'error'
        }
    }

    // If the user has selected two configs not supported
    if (config.cell && config.range) {
        if (config.range.length > 0 && config.cell.length > 0) {
            console.error('cagaste hubo un error 3');
            return {
                res: 'error'
            }
        }
    }

    ///////////////////////////////////////////

    const excel = await xlsx.fromFileAsync(path);

    if (!config.vertical){
        if (config.cell) {
            // If the user wants inset only an value
            if (config.cell.length > 0 && Array.isArray(value) == false) {
                console.log('only an value in specific cell');
                await excel.sheet(sheetNumber).cell(config.cell.toUpperCase()).value(value);
            }
            if (config.cell.length > 0 && Array.isArray(value) == true) {
                console.log('multiples values in a specific cell');
                await excel.sheet(sheetNumber).cell(config.cell.toUpperCase()).value([value]);
            }
        }
    }

    if (config.vertical){
        if (config.cell) {
            // If the user wants inset only an value
            if (config.cell.length > 0 && Array.isArray(value) == false) {
                console.log('only an value in specific cell');
                await excel.sheet(sheetNumber).cell(config.cell.toUpperCase()).value(value);
            }
            if (config.cell.length > 0 && Array.isArray(value) == true) {
                console.log('multiples values in a specific cell (Vertical)');
                const array = config.cell.split('');
                var row = array.map((e) => {
                    if (!isNaN(e)){
                        return e;
                    }
                })
                row = parseInt(row.join(''), 10);
                var col = array.map((el) => {
                    if (isNaN(el)){
                        return el;
                    }
                    else{
                        return '';
                    }
                })
                col = col.join('');
                // console.log(row, col);
                for (let i = 0; i < value.length; i++) {
                    await excel.sheet(sheetNumber).row(row++).cell(col).value(value[i]);
                }                
            }
        }
    }
    // Si el usuario digita un rango de filas en vertical
    if (config.vertical){
        if (config.range) {
            // If the user wants inset only an value
            if (config.range.length > 0 && Array.isArray(value) == false) {
                console.log('only an value in specific cell');
                await excel.sheet(sheetNumber).cell(config.cell.toUpperCase()).value(value);
            }
            if (config.range.length > 0 && Array.isArray(value) == true) {
                console.log('multiples values in a specific range (Vertical)');
                const array = config.range.split(':');
                let startRow, endRow, col;
                // get the column char
                col = array[0].split('');
                let newCol = col.map((e) => {
                    if (isNaN(e)) {
                        return e;
                    }
                    else {
                        return '';
                    }
                })
                newCol = newCol.join('')
                startRow = array[0].split('').map((e) => {
                    if (!isNaN(e)) {
                        return e;
                    }
                    else {
                        return '';
                    }
                })
                startRow = parseInt(startRow.join(''), 10);
                endRow = array[1].split('').map((e) => {
                    if (!isNaN(e)) {
                        return e;
                    }
                    else {
                        return '';
                    }
                })
                endRow = parseInt(endRow.join(''), 10);
                // console.log(newCol, startRow, endRow);
                let j = 0;
                for (let i = 0; i < endRow; i++) {
                    
                    if (j >= value.length) {
                        j = 0;
                        console.log(j);
                        
                        await excel.sheet(sheetNumber).row(startRow).cell(newCol).value(value[j]);
                        startRow++;
                        j++;
                    }
                    else {
                        console.log(j);
                        await excel.sheet(sheetNumber).row(startRow).cell(newCol).value(value[j]);
                        startRow++;
                        j++;
                    }
                }
            }
        }
    }

    if (!config.vertical){
        if (config.range){
            // If the user wants inset in array an value
            if (config.range.length > 0 && Array.isArray(value) == true) {
                console.log('multiples value in specific range'); 
                console.log(value);
                await excel.sheet(sheetNumber).range(config.range.toUpperCase()).value([value]);
            }
        }
    }

    excel.toFileAsync(path);

}

// writeToExcel('./test.xlsx', 0, ["hola3", "queso3", 'mango', 'pepino'], {range: 'B1:B10', vertical: true});

module.exports = { writeToExcel }