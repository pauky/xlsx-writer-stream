/**
 * Created by air on 16/11/21.
 */
var XLSXWriterStream = require('../index');

var writer = new XLSXWriterStream({
    file: 'example.xlsx'
    , onFinish: function (file) {
        console.log('finish');
    }
});

// Optional: Adjust column widths
writer.defineColumns([
    { width: 20 },
    { width: 10 }
]);

// Optional: Set cell map title
writer.setCellMap(['name', 'value']);

// Add some simple rows
writer.addRow(['pauky', 'ykk']);
writer.addRow(['glowry', 'yrw', 'test']);

// Add multiple rows
writer.addRows([['1', '2'],['a', 'b']]);

// Finalize the spreadsheet. If you don't do this, the readstream will not end.
writer.finalize();