/**
 * XlsxWriterStream
 */
'use strict';
var fs = require("fs");
var blobs = require('./blobs');
var Archiver = require('archiver');
var XlsxWriterStream = function (options) {
    this.init(options);
};

/**
 * 初始化
 * @param options
 */
XlsxWriterStream.prototype.init = function (options) {
    var _self = this;
    var PassThrough = require('stream').PassThrough;
    this.sheetStream = new PassThrough();

    var defaults = {
        file: '',
        defaultWidth: 15,
        zip: {
            forceUTC: true
        },
        columns: []
    };
    this.options = this._extend(defaults, options);
    this._resetSheet();
    this.defineColumns(this.options.columns);
    this.zip = Archiver('zip', this.options.zip);
    this.zip.catchEarlyExitAttached = true;
    this.zip.append(this.sheetStream, {
        name: 'xl/worksheets/sheet1.xml'
    });

    if (this.options.file) {
        this.fileStream = fs.createWriteStream(this.options.file);
        this.zip.pipe(this.fileStream);
        this.fileStream.on('finish', function () {
            _self.options.onFinish(_self.options.file);
        });
    }

};

/**
 *
 * @returns {*|exports|XlsxWriterStream.zip}
 */
XlsxWriterStream.prototype.getStream = function () {
    return this.zip;
};

/**
 * reset set
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._resetSheet = function () {
    this.sheetData = '';
    this.strings = [];
    this.stringMap = {};
    this.stringIndex = 0;
    this.stringData = null;
    this.currentRow = 0;
    this.cellMap = [];
    this.cellLabelMap = {};
    this.columns = [];
    this.relData = '';
    this.relationships = [];
    this.haveHeader = false;
    this.finalized = false;
    return this._write(blobs.sheetHeader);
};

/**
 * set cell map title info
 * @param headers
 */
XlsxWriterStream.prototype.setCellMap = function (cellMap) {
    this.cellMap = cellMap || [];
};

/**
 * define columns width
 * @param columns
 * @returns {*}
 */
XlsxWriterStream.prototype.defineColumns = function (columns) {
    if (this.haveHeader) {
        throw new Error("Columns cannot be added after rows! Unfortunately Excel will crash\nif column definitions come after sheet data. Please move your `defineColumns()`\ncall before any `addRow()` calls, or define options.columns in the XlsxWriterStream\nconstructor.");
    }
    this.options.columns = columns;
    return this._write(this._generateColumnDefinition());
};


/**
 *
 * @param dest
 * @param src
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._extend = function (dest, src) {
    var key, val;
    for (key in src) {
        val = src[key];
        dest[key] = val;
    }
    return dest;
};

/**
 *
 * @returns {string}
 * @private
 */
XlsxWriterStream.prototype._generateColumnDefinition = function () {
    if (!this.options.columns || !this.options.columns.length) {
        return '';
    }
    var _self = this;
    var columnDefinition = '';
    columnDefinition += blobs.startColumns;
    this.options.columns.forEach(function (val, i) {
        columnDefinition += blobs.column(val.width || _self.options.defaultWidth, i + 1);
    });
    columnDefinition += blobs.endColumns;
    return columnDefinition;
};

/**
 * add simple row
 * @param row
 * @returns {*}
 */
XlsxWriterStream.prototype.addRow = function (row) {
    var _self = this;
    if (!this.haveHeader) {
        this._write(blobs.sheetDataHeader);
        if (this.cellMap.length) {
            this._startRow();
            this.cellMap.forEach(function (val, i) {
                _self._addCell(val, i + 1);
            });
            this._endRow();
        }
        this.haveHeader = true;
    }
    this._startRow();
    row.forEach(function (val, i) {
        _self._addCell(val || "", i + 1);
    });
    return this._endRow();
};

/**
 * 写入多行数据
 * add multiple rows
 * @param rows
 */
XlsxWriterStream.prototype.addRows = function (rows) {
    var _self = this;
    rows.forEach(function (val, i) {
        _self.addRow(val);
    });
};

/**
 * write the start of row
 * @returns {number}
 * @private
 */
XlsxWriterStream.prototype._startRow = function () {
    this.rowBuffer = blobs.startRow(this.currentRow);
    return this.currentRow += 1;
};

/**
 * write the end of row
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._endRow = function () {
    return this._write(this.rowBuffer + blobs.endRow);
};

/**
 * write data to stream
 * @param data
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._write = function (data) {
    return this.sheetStream.write(data);
};

/**
 * add cell
 * @param value
 * @param col
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._addCell = function (value, col) {
    var cell, date, index, row;
    if (value == null) {
        value = '';
    }
    row = this.currentRow;
    cell = this._getCellIdentifier(row, col);
    if (Object.prototype.toString.call(value) === '[object Object]') {
        if (!value.value || !value.hyperlink) {
            throw new Error("A hyperlink cell must have both 'value' and 'hyperlink' keys.");
        }
        this._addCell(value.value, col);
        this._createRelationship(cell, value.hyperlink);
        return;
    }
    if (typeof value === 'number') {
        return this.rowBuffer += blobs.numberCell(value, cell);
    } else if (value instanceof Date) {
        date = this._dateToOADate(value);
        return this.rowBuffer += blobs.dateCell(date, cell);
    } else {
        index = this._lookupString(value);
        return this.rowBuffer += blobs.cell(index, cell);
    }
};

/**
 *
 * @param row
 * @param col
 * @returns {string}
 * @private
 */
XlsxWriterStream.prototype._getCellIdentifier = function (row, col) {
    var a, colIndex, input;
    colIndex = '';
    if (this.cellLabelMap[col]) {
        colIndex = this.cellLabelMap[col];
    } else {
        if (col === 0) {
            row = 1;
            col = 1;
        }
        input = (+col - 1).toString(26);
        while (input.length) {
            a = input.charCodeAt(input.length - 1);
            colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
            input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
        }
        this.cellLabelMap[col] = colIndex;
    }
    return colIndex + row;
};

/**
 *
 * @param cell
 * @param target
 * @returns {Number}
 * @private
 */
XlsxWriterStream.prototype._createRelationship = function (cell, target) {
    return this.relationships.push({
        cell: cell,
        target: target
    });
};

/**
 *
 * @param value
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._lookupString = function (value) {
    if (!this.stringMap[value]) {
        this.stringMap[value] = this.stringIndex;
        this.strings.push(value);
        this.stringIndex += 1;
    }
    return this.stringMap[value];
};

/**
 *
 * @returns {string}
 * @private
 */
XlsxWriterStream.prototype._generateStrings = function () {
    var string, stringTable, _i, _len, _ref;
    stringTable = '';
    _ref = this.strings;
    for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        string = _ref[_i];
        stringTable += blobs.string(this.escapeXml(string));
    }
    return this.stringsData = blobs.stringsHeader(this.strings.length) + stringTable + blobs.stringsFooter;
};

/**
 * filter xml format
 * @param str
 * @returns {XML|string}
 */
XlsxWriterStream.prototype.escapeXml = function (str) {
    if (str == null) {
        str = '';
    }
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
};

/**
 *
 * @returns {*}
 * @private
 */
XlsxWriterStream.prototype._generateRelationships = function () {
    return this.relsData = blobs.externalWorksheetRels(this.relationships);
};

/**
 * end the writer
 * @returns {*}
 */
XlsxWriterStream.prototype.finalize = function () {
    if (this.finalized) {
        throw new Error("This XLSX was already finalized.");
    }
    this.finalized = true;
    if (this.haveHeader) {
        this._write(blobs.sheetDataFooter);
    }
    this._write(blobs.worksheetRels(this.relationships));
    this._generateStrings();
    this._generateRelationships();
    this.sheetStream.end(blobs.sheetFooter);
    return this._finalizeZip();
};

/**
 * zip the data of xml to xlsx
 * @returns {*|void|this}
 * @private
 */
XlsxWriterStream.prototype._finalizeZip = function () {
    return this.zip
        .append(blobs.contentTypes, {
            name: '[Content_Types].xml'
        })
        .append(blobs.rels, {
            name: '_rels/.rels'
        })
        .append(blobs.workbook, {
            name: 'xl/workbook.xml'
        })
        .append(blobs.styles, {
            name: 'xl/styles.xml'
        })
        .append(blobs.workbookRels, {
            name: 'xl/_rels/workbook.xml.rels'
        })
        .append(this.relsData, {
            name: 'xl/worksheets/_rels/sheet1.xml.rels'
        })
        .append(this.stringsData, {
            name: 'xl/sharedStrings.xml'
        })
        .finalize();
};

module.exports = XlsxWriterStream;



