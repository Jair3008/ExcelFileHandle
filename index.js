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

writeToExcel('./test.xlsx', 0, ["hola3", "queso3", 'barret', 'monsta'], {cell: "B1", vertical: true});