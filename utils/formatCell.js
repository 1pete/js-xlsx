var safeFormatCell = require('./safeFormatCell');

module.exports = function formatCell(cell, v) {
  if (cell == null || cell.t == null) return '';
  if (cell.w != null) return cell.w;
  if (v == null) return safeFormatCell(cell, cell.v);
  return safeFormatCell(cell, v);
};
