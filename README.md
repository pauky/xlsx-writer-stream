# xlsx-writer-stream

  A xlsx writer by stream in Node.js

## Installation

```
$ npm install xlsx-writer-stream
```

## Example

```js
var XLSXWriterStream = require('xlsx-wirter-stream');

var writer = new XLSXWriter({
    file: 'example.xlsx'
});

// Optional: Adjust column widths
writer.defineColumns([
    { width: 20 }, // width is in 'characters'
    { width: 10 }
]);

// Optional: Set cell map title
writer.setCellMap(['name', 'value']);

// Add some simple rows
writer.addRow(['pauky', 'ykk']);
writer.addRow(['glowry', 'yrw', 'test']);

// Add multiple row
writer.addRows([['1', '2'],['a', 'b']]);

// Finalize the spreadsheet. If you don't do this, the readstream will not end.
writer.finalize();
```

# License

  MIT
