var SSF = require('../modules/SSF');

module.exports = function safeFormatCell(cell, v) {
  if (cell.z != null) {
    try {
      return (cell.w = SSF.format(cell.z, v));
    } catch (e) { /* do nothing */ }
  }
  if (!cell.XF) return v;
  try { return (cell.w = SSF.format(cell.XF.ifmt || 0, v)); } catch (e) { return '' + v; }
};
