const isaiX = require('../index.js');

const values = ["5", "2", "3"];

isaiX.WriteToExcel('../test.xlsx', 0, 'A1', 'A3', values, false);