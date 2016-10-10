module.exports = function splitCell(cstr) {
  return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, '$1,$2').split(',');
};
