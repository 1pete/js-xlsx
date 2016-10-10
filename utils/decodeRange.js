var decodeCell = require('./decodeCell');

module.exports = function decodeRange(range) {
  var x = range.split(':').map(decodeCell);
  return { s: x[0], e: x[x.length - 1] };
};
