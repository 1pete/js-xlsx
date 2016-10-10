function unfixCol(cstr) { return cstr.replace(/^\$([A-Z])/, '$1'); }
module.exports = function decodeCol(colstr) {
  var c = unfixCol(colstr);
  var d = 0;
  var i = 0;
  for (; i !== c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
  return d - 1;
};
