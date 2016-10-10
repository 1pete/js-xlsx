var OFFCRYPTO = {};

OFFCRYPTO.rc4 = function (key, data) {
  var S = new Array(256);
  var c = 0;
  var i = 0;
  var j = 0;
  var t = 0;
  for (i = 0; i !== 256; ++i) S[i] = i;
  for (i = 0; i !== 256; ++i) {
    j = (j + S[i] + (key[i % key.length]).charCodeAt(0)) & 255;
    t = S[i]; S[i] = S[j]; S[j] = t;
  }
  i = j = 0;
  var out = Buffer(data.length);
  for (c = 0; c !== data.length; ++c) {
    i = (i + 1) & 255;
    j = (j + S[i]) % 256;
    t = S[i]; S[i] = S[j]; S[j] = t;
    out[c] = (data[c] ^ S[(S[i] + S[j]) & 255]);
  }
  return out;
};

OFFCRYPTO.md5 = function () { throw 'unimplemented'; };

module.exports = OFFCRYPTO;
