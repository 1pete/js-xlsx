var encodeCol = require('./encodeCol');
var encodeRow = require('./encodeRow');

module.exports = function encodeCell(cell) { return encodeCol(cell.c) + encodeRow(cell.r); };
