function unfixRow(cstr) { return cstr.replace(/\$(\d+)$/, '$1'); }
module.exports = function decodeRow(rowstr) { return parseInt(unfixRow(rowstr), 10) - 1; }
