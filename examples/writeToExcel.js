const isaiX = require('../index.js');

const values = ["1", "2", "3"];

isaiX.WriteToExcel('../test.xlsx', 0, 'F2', 'F10', values);