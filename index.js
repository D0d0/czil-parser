var XLSX = require('xlsx');
var util = require('util');
var fs = require('fs');

var workbook = XLSX.readFile('CZIL.xls');

var sheetName = workbook['SheetNames'][0];

//console.log('vInfo', util.inspect(workbook['Sheets'][sheet], false, null, true /* enable colors */));


var gliders = [];
var sheet = workbook['Sheets'][sheetName];
for (var key in sheet) {
    if (sheet.hasOwnProperty(key) && key.toLowerCase()[0] === 'a') {
        // console.log(util.inspect(sheet[key].v, false, null, true /* enable colors */));
        // sheet[key].v.split(',').map(x => x.trim())

        let row = sheet[key].v
            .split(',')
            .filter(x => !x.includes('odstranit'))
            .map(x => x.trim())
            .map(x => x.replace(/  +/g, ' '));
        gliders.push(...row);
    }
}


// remove headers from file
gliders = gliders.slice(2);

console.log(util.inspect(gliders, false, null, true /* enable colors */));


var stream = fs.createWriteStream("my_file.txt");
stream.once('open', function(fd) {
    for (var i = 0; i < gliders.length; i++){
        stream.write(`${gliders[i]}\n`);
    }
    stream.end();
});