var decodeCol = require('./decodeCol');
var decodeRow = require('./decodeRow');
var splitCell = require('./splitCell');

module.exports = function decodeCell(cstr) {
  var splt = splitCell(cstr);
  return { c: decodeCol(splt[0]), r: decodeRow(splt[1]) };
};
