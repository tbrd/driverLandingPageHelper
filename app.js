var xlsx = require("xlsx");
var fs = require('fs');

var workbook = xlsx.readFile("input.xlsx");
var worksheet = workbook.Sheets["without_cars"];

if (!worksheet) {
    throw "Worksheet not found: 'without_cars'";
}
var region;

/**
 * Columns
 * A:       Region
 * B:       City
 * C:       City
 * D:       Type (car/motorbike)
 * E:       Hourly Pay
 * F:       Drop Rate
 * G:       Text to display
 **/

var rows = [];
var city;

for (cellLocation in worksheet) {
    /* all keys that do not begin with "!" correspond to cell addresses */
    if (cellLocation[0] === "!") continue;

    var column = cellLocation.substr(0, 1);
    var row = parseInt(cellLocation.substr(1), 10);

    // ignore first two rows
    if (row <= 2) continue;
    
    var value = worksheet[cellLocation].v;
    rows[row] = rows[row] || {};
    rows[row][column] = value;
}

// fix cities
rows.forEach(function (row, i) {
    if (!row['C']) {
        row['C'] = rows[i - 1]['C'];
    }
});

var cityData = {};

rows.map(row =>
    ({
        city: row.C.toLowerCase(),
        type: row.D.toLowerCase() === 'motorbike' ? 'scooter' : row.D.toLowerCase(),
        text: row.G
    })
).sort((a, b) =>
    a.city < b.city ? -1 :
    a.city > b.city ? 1 :
    a.type > b.type ? -1 :
    a.type < b.type ? 1 :
    0
).forEach(row => {
    cityData[row.city] = cityData[row.city] || {};
    cityData[row.city][row.type] = row.text;
});

var json = JSON
    // pretty print JSON
    .stringify(cityData, null, 2)
    // remove quotes from keys
    .replace(/\"([^(\")"]+)\":/g,"$1:")
    // replace double quotes with single quotes
    .replace(/\"/g, "'");

var stream = fs.createWriteStream('output.json', { flags : 'w' });
stream.write(json);


// csv.generate({seed: 1, columns: 2, length: 20}, function(err, data){
//   csv.parse(data, function(err, data){
//     csv.transform(data, function(data){
//       return data.map(function(value){return value.toUpperCase()});
//     }, function(err, data){
//       csv.stringify(data, function(err, data){
//         process.stdout.write(data);
//       });
//     });
//   });
// });


// xlsxj({
//     input: "HiringLandingPagePayRates.xlsx", 
//     output: "driverLandingPageData.json",
//     sheet: "without_cars"
// }, function(err, result) {
//     if(err) {
//         console.error(err);
//     } else {
//         console.log(result);
//     }
// });