let XLSX = require('xlsx');
let util = require('util');
let fs = require('fs');

let workbook = XLSX.readFile('CZIL.xls');
let sheetName = workbook['SheetNames'][0];
let sheet = workbook['Sheets'][sheetName];

let gliders = [];
for (var key in sheet) {
    if (sheet.hasOwnProperty(key) && key.toLowerCase()[0] === 'a') {
        let row = sheet[key].v
            .split(',')
            .filter(_ => !_.includes('odstranit'))
            .map(_ => _.trim())
            .map(_ => _.replace(/  +/g, ' '))
            .map(_ => _.replace(/"/g, '""'));

        gliders.push(...row.map(x => ({
            value: x,
            key: key.slice(1)
        })));
    }
}

// remove headers from file
gliders = gliders.slice(2);

let fileStream = fs.createWriteStream('list.csv');
fileStream.once('open', function (fd) {
    fileStream.write(`Glider,key\n`);
    for (let i = 0; i < gliders.length; i++) {
        fileStream.write(`"${gliders[i].value}",${gliders[i].key}\n`);
    }
    fileStream.end();
});