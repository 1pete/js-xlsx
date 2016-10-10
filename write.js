/* xlsx.js (C) 2013-2015 SheetJS -- http://sheetjs.com */

var jszip = require('jszip');

var SSF = require('./modules/SSF');
var StyleBuilder = require('./modules/StyleBuilder');

var encodeCell = require('./utils/encodeCell');
var encodeCol = require('./utils/encodeCol');
var encodeRange = require('./utils/encodeRange');
var encodeRow = require('./utils/encodeRow');
var safeDecodeRange = require('./utils/safeDecodeRange');

var style_builder;

var _getchar = function _gc1(x) { return String.fromCharCode(x); };

var has_buf = (typeof Buffer !== 'undefined');

function new_raw_buf(len) {
  return new (has_buf ? Buffer : Array)(len);
}

var chr0 = /\u0000/g;

function isval(x) { return x !== undefined && x !== null; }

function keys(o) { return Object.keys(o); }

function evert_key(obj, key) {
  var o = [];
  var K = keys(obj);
  for (var i = 0; i !== K.length; ++i) o[obj[K[i]][key]] = K[i];
  return o;
}

function evert(obj) {
  var o = [];
  var K = keys(obj);
  for (var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
  return o;
}

function evert_num(obj) {
  var o = [];
  var K = keys(obj);
  for (var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i], 10);
  return o;
}

function evert_arr(obj) {
  var o = [];
  var K = keys(obj);
  for (var i = 0; i !== K.length; ++i) {
    if (o[obj[K[i]]] == null) o[obj[K[i]]] = [];
    o[obj[K[i]]].push(K[i]);
  }
  return o;
}

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}

var encodings = {
  '&quot;': '"',
  '&apos;': "'",
  '&gt;': '>',
  '&lt;': '<',
  '&amp;': '&',
};
var rencoding = evert(encodings);

var decregex = /[&<>'"]/g;
var charegex = /[\u0000-\u0008\u000b-\u001f]/g;
function escapexml(text) {
  var s = text + '';
  return s.replace(decregex, function (y) { return rencoding[y]; })
    .replace(charegex, function (ss) { return '_x' + ('000' + ss.charCodeAt(0).toString(16)).substr(-4) + '_'; });
}

function parsexmlbool(value) {
  switch (value) {
    case '1': case 'true': case 'TRUE': return true;
    default: return false;
  }
}

var utf8read = function utf8reada(orig) {
  var out = '';
  var i = 0;
  var c = 0;
  var d = 0;
  var e = 0;
  var f = 0;
  var w = 0;
  while (i < orig.length) {
    c = orig.charCodeAt(i++);
    if (c < 128) { out += String.fromCharCode(c); continue; }
    d = orig.charCodeAt(i++);
    if (c > 191 && c < 224) { out += String.fromCharCode(((c & 31) << 6) | (d & 63)); continue; }
    e = orig.charCodeAt(i++);
    if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
    f = orig.charCodeAt(i++);
    w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63)) - 65536;
    out += String.fromCharCode(0xD800 + ((w >>> 10) & 1023));
    out += String.fromCharCode(0xDC00 + (w & 1023));
  }
  return out;
};

if (has_buf) {
  var utf8readb = function utf8readb(data) {
    var out = new Buffer(2 * data.length);
    var w;
    var i;
    var j = 1;
    var k = 0;
    var ww = 0;
    var c;
    for (i = 0; i < data.length; i += j) {
      j = 1;
      if ((c = data.charCodeAt(i)) < 128) w = c;
      else if (c < 224) {
        w = (c & 31) * 64 + (data.charCodeAt(i + 1) & 63);
        j = 2;
      } else if (c < 240) {
        w = (c & 15) * 4096 + (data.charCodeAt(i + 1) & 63) * 64 + (data.charCodeAt(i + 2) & 63);
        j = 3;
      } else {
        j = 4;
        w = (c & 7) * 262144 + (data.charCodeAt(i + 1) & 63) * 4096 + (data.charCodeAt(i + 2) & 63) * 64 + (data.charCodeAt(i + 3) & 63);
        w -= 65536; ww = 0xD800 + ((w >>> 10) & 1023); w = 0xDC00 + (w & 1023);
      }
      if (ww !== 0) { out[k++] = ww & 255; out[k++] = ww >>> 8; ww = 0; }
      out[k++] = w % 256; out[k++] = w >>> 8;
    }
    out.length = k;
    return out.toString('ucs2');
  };
  var corpus = 'foo bar baz\u00e2\u0098\u0083\u00f0\u009f\u008d\u00a3';
  if (utf8read(corpus) == utf8readb(corpus)) utf8read = utf8readb;
  var utf8readc = function utf8readc(data) { return Buffer(data, 'binary').toString('utf8'); };
  if (utf8read(corpus) == utf8readc(corpus)) utf8read = utf8readc;
}

var wtregex = /(^\s|\s$|\n)/;
function writetag(f, g) { return '<' + f + (g.match(wtregex) ? ' xml:space="preserve"' : '') + '>' + g + '</' + f + '>'; }

function wxt_helper(h) { return keys(h).map(function (k) { return ' ' + k + '="' + h[k] + '"'; }).join(''); }
function writextag(f, g, h) { return '<' + f + (isval(h) ? wxt_helper(h) : '') + (isval(g) ? (g.match(wtregex) ? ' xml:space="preserve"' : '') + '>' + g + '</' + f : '/') + '>'; }

function write_w3cdtf(d, t) {
  try {
    return d.toISOString().replace(/\.\d*/, '');
  } catch (e) {
    if (t) throw e;
  }
  return null;
}

function write_vt(s) {
  switch (typeof s) {
    case 'string': return writextag('vt:lpwstr', s);
    case 'number': return writextag((s | 0) == s ? 'vt:i4' : 'vt:r8', String(s));
    case 'boolean': return writextag('vt:bool', s ? 'true' : 'false');
  }
  if (s instanceof Date) return writextag('vt:filetime', write_w3cdtf(s));
  throw new Error('Unable to serialize ' + s);
}

var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
var XMLNS = {
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  mx: 'http://schemas.microsoft.com/office/mac/excel/2008/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  sjs: 'http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties',
  vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  xsd: 'http://www.w3.org/2001/XMLSchema',
};

XMLNS.main = [
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
  'http://schemas.microsoft.com/office/excel/2006/main',
  'http://schemas.microsoft.com/office/excel/2006/2',
];

function readIEEE754(buf, idx, isLE, nl, ml) {
  if (isLE === undefined) isLE = true;
  if (!nl) nl = 8;
  if (!ml && nl === 8) ml = 52;
  var e;
  var m;
  var el = nl * 8 - ml - 1;
  var eMax = (1 << el) - 1;
  var eBias = eMax >> 1;
  var bits = -7;
  var d = isLE ? -1 : 1;
  var i = isLE ? (nl - 1) : 0;
  var s = buf[idx + i];

  i += d;
  e = s & ((1 << (-bits)) - 1); s >>>= (-bits); bits += el;
  for (; bits > 0; e = e * 256 + buf[idx + i], i += d, bits -= 8);
  m = e & ((1 << (-bits)) - 1); e >>>= (-bits); bits += ml;
  for (; bits > 0; m = m * 256 + buf[idx + i], i += d, bits -= 8);
  if (e === eMax) return m ? NaN : ((s ? -1 : 1) * Infinity);
  else if (e === 0) e = 1 - eBias;
  else {
    m += Math.pow(2, ml);
    e -= eBias;
  }
  return (s ? -1 : 1) * m * Math.pow(2, e - ml);
}

var __readUInt8 = function (b, idx) { return b[idx]; };
var __readUInt16LE = function (b, idx) { return b[idx + 1] * (1 << 8) + b[idx]; };
var __readInt16LE = function (b, idx) { var u = b[idx + 1] * (1 << 8) + b[idx]; return (u < 0x8000) ? u : (0xffff - u + 1) * -1; };
var __readUInt32LE = function (b, idx) { return b[idx + 3] * (1 << 24) + (b[idx + 2] << 16) + (b[idx + 1] << 8) + b[idx]; };
var __readInt32LE = function (b, idx) { return (b[idx + 3] << 24) | (b[idx + 2] << 16) | (b[idx + 1] << 8) | b[idx]; };

var __toBuffer;
var ___toBuffer;
__toBuffer = ___toBuffer = function toBuffer_(bufs) { var x = []; for (var i = 0; i < bufs[0].length; ++i) { x.push.apply(x, bufs[0][i]); } return x; };
var __utf16le;
var ___utf16le;
__utf16le = ___utf16le = function utf16le_(b, s, e) { var ss = []; for (var i = s; i < e; i += 2) ss.push(String.fromCharCode(__readUInt16LE(b, i))); return ss.join(''); };
var __hexlify;
var ___hexlify;
__hexlify = ___hexlify = function hexlify_(b, s, l) { return b.slice(s, (s + l)).map(function (x) { return (x < 16 ? '0' : '') + x.toString(16); }).join(''); };
var __utf8;
__utf8 = function (b, s, e) { var ss = []; for (var i = s; i < e; i++) ss.push(String.fromCharCode(__readUInt8(b, i))); return ss.join(''); };
var __lpstr;
var ___lpstr;
__lpstr = ___lpstr = function lpstr_(b, i) { var len = __readUInt32LE(b, i); return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : ''; };
var __lpwstr;
var ___lpwstr;
__lpwstr = ___lpwstr = function lpwstr_(b, i) { var len = 2 * __readUInt32LE(b, i); return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : ''; };
var __double;
var ___double;
__double = ___double = function (b, idx) { return readIEEE754(b, idx); };

var is_buf = function is_buf_a(a) { return Array.isArray(a); };
if (has_buf) {
  __utf16le = function utf16le_b(b, s, e) { if (!Buffer.isBuffer(b)) return ___utf16le(b, s, e); return b.toString('utf16le', s, e); };
  __hexlify = function (b, s, l) { return Buffer.isBuffer(b) ? b.toString('hex', s, s + l) : ___hexlify(b, s, l); };
  __lpstr = function lpstr_b(b, i) { if (!Buffer.isBuffer(b)) return ___lpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8', i + 4, i + 4 + len - 1) : ''; };
  __lpwstr = function lpwstr_b(b, i) { if (!Buffer.isBuffer(b)) return ___lpwstr(b, i); var len = 2 * b.readUInt32LE(i); return b.toString('utf16le', i + 4, i + 4 + len - 1); };
  __utf8 = function utf8_b(s, e) { return this.toString('utf8', s, e); };
  __toBuffer = function (bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs); };
  __double = function double_(b, i) { if (Buffer.isBuffer(b)) return b.readDoubleLE(i); return ___double(b, i); };
  is_buf = function is_buf_b(a) { return Buffer.isBuffer(a) || Array.isArray(a); };
}

function ReadShift(size, t) {
  var o = '';
  var oI;
  var oR;
  var oo = [];
  var w;
  var vv;
  var i;
  var loc;
  switch (t) {
    case 'dbcs':
      loc = this.l;
      if (has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l + 2 * size).toString('utf16le');
      else for (i = 0; i !== size; ++i) { o += String.fromCharCode(__readUInt16LE(this, loc)); loc += 2; }
      size *= 2;
      break;

    case 'utf8': o = __utf8(this, this.l, this.l + size); break;
    case 'utf16le': size *= 2; o = __utf16le(this, this.l, this.l + size); break;
    case 'lpstr': o = __lpstr(this, this.l); size = 5 + o.length; break;
    case 'lpwstr': o = __lpwstr(this, this.l); size = 5 + o.length; if (o[o.length - 1] === '\u0000') size += 2; break;

    case 'cstr': size = 0; o = '';
      while ((w = __readUInt8(this, this.l + size++)) !== 0) oo.push(_getchar(w));
      o = oo.join(''); break;
    case 'wstr': size = 0; o = '';
      while ((w = __readUInt16LE(this, this.l + size)) !== 0) { oo.push(_getchar(w)); size += 2; }
      size += 2; o = oo.join(''); break;

    case 'dbcs-cont': o = ''; loc = this.l;
      for (i = 0; i !== size; ++i) {
        if (this.lens && this.lens.indexOf(loc) !== -1) {
          w = __readUInt8(this, loc);
          this.l = loc + 1;
          vv = ReadShift.call(this, size - i, w ? 'dbcs-cont' : 'sbcs-cont');
          return oo.join('') + vv;
        }
        oo.push(_getchar(__readUInt16LE(this, loc)));
        loc += 2;
      } o = oo.join(''); size *= 2; break;

    case 'sbcs-cont': o = ''; loc = this.l;
      for (i = 0; i !== size; ++i) {
        if (this.lens && this.lens.indexOf(loc) !== -1) {
          w = __readUInt8(this, loc);
          this.l = loc + 1;
          vv = ReadShift.call(this, size - i, w ? 'dbcs-cont' : 'sbcs-cont');
          return oo.join('') + vv;
        }
        oo.push(_getchar(__readUInt8(this, loc)));
        loc += 1;
      } o = oo.join(''); break;

    default:
      switch (size) {
        case 1:
          oI = __readUInt8(this, this.l);
          this.l++;
          return oI;
        case 2:
          oI = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l);
          this.l += 2;
          return oI;
        case 4:
          if (t === 'i' || (this[this.l + 3] & 0x80) === 0) {
            oI = __readInt32LE(this, this.l);
            this.l += 4;
            return oI;
          }
          oR = __readUInt32LE(this, this.l);
          this.l += 4;
          return oR;
        case 8: if (t === 'f') { oR = __double(this, this.l); this.l += 8; return oR; }
        case 16: o = __hexlify(this, this.l, size); break;
      }
  }
  this.l += size;
  return o;
}

function WriteShift(t, val, f) {
  var size;
  var i;
  if (f === 'dbcs') {
    for (i = 0; i !== val.length; ++i) this.writeUInt16LE(val.charCodeAt(i), this.l + 2 * i);
    size = 2 * val.length;
  } else {
    switch (t) {
      case 1: size = 1; this[this.l] = val & 255; break;
      case 3: size = 3; this[this.l + 2] = val & 255; val >>>= 8; this[this.l + 1] = val & 255; val >>>= 8; this[this.l] = val & 255; break;
      case 4: size = 4; this.writeUInt32LE(val, this.l); break;
      case 8: size = 8; if (f === 'f') { this.writeDoubleLE(val, this.l); break; }
      case 16: break;
      case -4: size = 4; this.writeInt32LE(val, this.l); break;
    }
  }
  this.l += size; return this;
}

function CheckField(hexstr, fld) {
  var m = __hexlify(this, this.l, hexstr.length >> 1);
  if (m !== hexstr) throw fld + 'Expected ' + hexstr + ' saw ' + m;
  this.l += hexstr.length >> 1;
}

function prep_blob(blob, pos) {
  blob.l = pos;
  blob.read_shift = ReadShift;
  blob.chk = CheckField;
  blob.write_shift = WriteShift;
}

function parsenoop(blob, length) { blob.l += length; }

function new_buf(sz) {
  var o = new_raw_buf(sz);
  prep_blob(o, 0);
  return o;
}

function buf_array() {
  var bufs = [];
  var blksz = 2048;
  var newblk = function ba_newblk(sz) {
    var o = new_buf(sz);
    prep_blob(o, 0);
    return o;
  };

  var curbuf = newblk(blksz);

  var endbuf = function ba_endbuf() {
    curbuf.length = curbuf.l;
    if (curbuf.length > 0) bufs.push(curbuf);
    curbuf = null;
  };

  var next = function ba_next(sz) {
    if (sz < curbuf.length - curbuf.l) return curbuf;
    endbuf();
    return (curbuf = newblk(Math.max(sz + 1, blksz)));
  };

  var end = function ba_end() {
    endbuf();
    return __toBuffer([bufs]);
  };

  var push = function ba_push(buf) { endbuf(); curbuf = buf; next(blksz); };

  return { next: next, push: push, end: end, _bufs: bufs };
}

function write_record(ba, type, payload, length) {
  var t = evert_RE[type];
  var l;
  if (!length) length = XLSBRecordEnum[t].p || (payload || []).length || 0;
  l = 1 + (t >= 0x80 ? 1 : 0) + 1 + length;
  if (length >= 0x80) ++l; if (length >= 0x4000) ++l; if (length >= 0x200000) ++l;
  var o = ba.next(l);
  if (t <= 0x7F) o.write_shift(1, t);
  else {
    o.write_shift(1, (t & 0x7F) + 0x80);
    o.write_shift(1, (t >> 7));
  }
  for (var i = 0; i !== 4; ++i) {
    if (length >= 0x80) {
      o.write_shift(1, (length & 0x7F) + 0x80);
      length >>= 7;
    } else {
      o.write_shift(1, length);
      break;
    }
  }
  if (length > 0 && is_buf(payload)) ba.push(payload);
}

function parse_StrRun(data) {
  return { ich: data.read_shift(2), ifnt: data.read_shift(2) };
}

/* [MS-XLSB] 2.1.7.121 */
function parse_RichStr(data, length) {
  var start = data.l;
  var flags = data.read_shift(1);
  var str = parse_XLWideString(data);
  var rgsStrRun = [];
  var z = { t: str, h: str };
  if ((flags & 1) !== 0) { /* fRichStr */
    var dwSizeStrRun = data.read_shift(4);
    for (var i = 0; i !== dwSizeStrRun; ++i) rgsStrRun.push(parse_StrRun(data));
    z.r = rgsStrRun;
  } else z.r = '<t>' + escapexml(str) + '</t>';
  if ((flags & 2) !== 0) { /* fExtStr */
  }
  data.l = start + length;
  return z;
}
function write_RichStr(str, o) {
  if (o == null) o = new_buf(5 + 2 * str.t.length);
  o.write_shift(1, 0);
  write_XLWideString(str.t, o);
  return o;
}

/* [MS-XLSB] 2.5.9 */
function parse_XLSBCell(data) {
  var col = data.read_shift(4);
  var iStyleRef = data.read_shift(2);
  iStyleRef += data.read_shift(1) << 16;
  data.read_shift(1);
  return { c: col, iStyleRef: iStyleRef };
}
function write_XLSBCell(cell, o) {
  if (o == null) o = new_buf(8);
  o.write_shift(-4, cell.c);
  o.write_shift(3, cell.iStyleRef === undefined ? cell.iStyleRef : cell.s);
  o.write_shift(1, 0); /* fPhShow */
  return o;
}

function parse_XLSBCodeName(data, length) { return parse_XLWideString(data, length); }

function parse_XLNullableWideString(data) {
  var cchCharacters = data.read_shift(4);
  return cchCharacters === 0 || cchCharacters === 0xFFFFFFFF ? '' : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLNullableWideString(data, o) {
  if (!o) o = new_buf(127);
  o.write_shift(4, data.length > 0 ? data.length : 0xFFFFFFFF);
  if (data.length > 0) o.write_shift(0, data, 'dbcs');
  return o;
}

/* [MS-XLSB] 2.5.168 */
function parse_XLWideString(data) {
  var cchCharacters = data.read_shift(4);
  return cchCharacters === 0 ? '' : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLWideString(data, o) {
  if (o == null) o = new_buf(4 + 2 * data.length);
  o.write_shift(4, data.length);
  if (data.length > 0) o.write_shift(0, data, 'dbcs');
  return o;
}

/* [MS-XLSB] 2.5.114 */
var parse_RelID = parse_XLNullableWideString;
var write_RelID = write_XLNullableWideString;


/* [MS-XLSB] 2.5.122 */
/* [MS-XLS] 2.5.217 */
function parse_RkNumber(data) {
  var b = data.slice(data.l, data.l + 4);
  var fX100 = b[0] & 1;
  var fInt = b[0] & 2;
  data.l += 4;
  b[0] &= 0xFC; // b[0] &= ~3;
  var RK = fInt === 0 ? __double([0, 0, 0, 0, b[0], b[1], b[2], b[3]], 0) : __readInt32LE(b, 0) >> 2;
  return fX100 ? RK / 100 : RK;
}

/* [MS-XLSB] 2.5.153 */
function parse_UncheckedRfX(data) {
  var cell = { s: {}, e: {} };
  cell.s.r = data.read_shift(4);
  cell.e.r = data.read_shift(4);
  cell.s.c = data.read_shift(4);
  cell.e.c = data.read_shift(4);
  return cell;
}

function write_UncheckedRfX(r, o) {
  if (!o) o = new_buf(16);
  o.write_shift(4, r.s.r);
  o.write_shift(4, r.e.r);
  o.write_shift(4, r.s.c);
  o.write_shift(4, r.e.c);
  return o;
}

function parse_Xnum(data) { return data.read_shift(8, 'f'); }
function write_Xnum(data, o) { return (o || new_buf(8)).write_shift(8, 'f', data); }

/* [MS-XLSB] 2.5.198.2 */
var BErr = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0F: '#VALUE!',
  0x17: '#REF!',
  0x1D: '#NAME?',
  0x24: '#NUM!',
  0x2A: '#N/A',
  0x2B: '#GETTING_DATA',
  0xFF: '#WTF?',
};

function parse_BrtColor(data) {
  var out = {};
  var d = data.read_shift(1);
  out.fValidRGB = d & 1;
  out.xColorType = d >>> 1;
  out.index = data.read_shift(1);
  out.nTintAndShade = data.read_shift(2, 'i');
  out.bRed = data.read_shift(1);
  out.bGreen = data.read_shift(1);
  out.bBlue = data.read_shift(1);
  out.bAlpha = data.read_shift(1);
}

/* [MS-XLSB] 2.5.52 */
function parse_FontFlags(data) {
  var d = data.read_shift(1);
  data.l++;
  var out = {
    fItalic: d & 0x2,
    fStrikeout: d & 0x8,
    fOutline: d & 0x10,
    fShadow: d & 0x20,
    fCondense: d & 0x40,
    fExtend: d & 0x80,
  };
  return out;
}

var VT_I4 = 0x0003;
var VT_VARIANT = 0x000C;

var VT_STRING = 0x0050;
var VT_USTR = 0x0051;
var VT_CUSTOM = [VT_STRING, VT_USTR];

var ct2type = {
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml': 'workbooks',
  'application/vnd.ms-excel.binIndexWs': 'TODO',
  'application/vnd.ms-excel.chartsheet': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml': 'TODO',
  'application/vnd.ms-excel.dialogsheet': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml': 'TODO',
  'application/vnd.ms-excel.macrosheet': 'TODO',
  'application/vnd.ms-excel.macrosheet+xml': 'TODO',
  'application/vnd.ms-excel.intlmacrosheet': 'TODO',
  'application/vnd.ms-excel.binIndexMs': 'TODO',
  'application/vnd.openxmlformats-package.core-properties+xml': 'coreprops',
  'application/vnd.openxmlformats-officedocument.custom-properties+xml': 'custprops',
  'application/vnd.openxmlformats-officedocument.extended-properties+xml': 'extprops',
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml': 'TODO',
  'application/vnd.ms-excel.comments': 'comments',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml': 'comments',

  'application/vnd.ms-excel.pivotTable': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml': 'TODO',

  'application/vnd.ms-excel.calcChain': 'calcchains',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml': 'calcchains',

  'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings': 'TODO',

  'application/vnd.ms-office.activeX': 'TODO',
  'application/vnd.ms-office.activeX+xml': 'TODO',

  'application/vnd.ms-excel.attachedToolbars': 'TODO',

  'application/vnd.ms-excel.connections': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml': 'TODO',

  'application/vnd.ms-excel.externalLink': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml': 'TODO',

  'application/vnd.ms-excel.sheetMetadata': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml': 'TODO',

  'application/vnd.ms-excel.pivotCacheDefinition': 'TODO',
  'application/vnd.ms-excel.pivotCacheRecords': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml': 'TODO',

  'application/vnd.ms-excel.queryTable': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml': 'TODO',

  'application/vnd.ms-excel.userNames': 'TODO',
  'application/vnd.ms-excel.revisionHeaders': 'TODO',
  'application/vnd.ms-excel.revisionLog': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml': 'TODO',

  'application/vnd.ms-excel.tableSingleCells': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml': 'TODO',
  'application/vnd.ms-excel.slicer': 'TODO',
  'application/vnd.ms-excel.slicerCache': 'TODO',
  'application/vnd.ms-excel.slicer+xml': 'TODO',
  'application/vnd.ms-excel.slicerCache+xml': 'TODO',
  'application/vnd.ms-excel.wsSortMap': 'TODO',
  'application/vnd.ms-excel.table': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.theme+xml': 'themes',
  'application/vnd.ms-excel.Timeline+xml': 'TODO', /* verify */
  'application/vnd.ms-excel.TimelineCache+xml': 'TODO', /* verify */

  'application/vnd.ms-office.vbaProject': 'vba',
  'application/vnd.ms-office.vbaProjectSignature': 'vba',

  'application/vnd.ms-office.volatileDependencies': 'TODO',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml': 'TODO',

  'application/vnd.ms-excel.controlproperties+xml': 'TODO',

  'application/vnd.openxmlformats-officedocument.model+data': 'TODO',

  'application/vnd.ms-excel.Survey+xml': 'TODO',

  'application/vnd.openxmlformats-officedocument.drawing+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml': 'TODO',
  'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml': 'TODO',

  'application/vnd.openxmlformats-officedocument.vmlDrawing': 'TODO',

  'application/vnd.openxmlformats-package.relationships+xml': 'rels',
  'application/vnd.openxmlformats-officedocument.oleObject': 'TODO',

  sheet: 'js',
};

var CT_LIST = (function () {
  var o = {
    workbooks: {
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      xlsb: 'application/vnd.ms-excel.sheet.binary.macroEnabled.main',
      xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
    },
    strs: { /* Shared Strings */
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
      xlsb: 'application/vnd.ms-excel.sharedStrings',
    },
    sheets: {
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      xlsb: 'application/vnd.ms-excel.worksheet',
    },
    styles: {/* Styles */
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      xlsb: 'application/vnd.ms-excel.styles',
    },
  };
  keys(o).forEach(function (k) { if (!o[k].xlsm) o[k].xlsm = o[k].xlsx; });
  keys(o).forEach(function (k) { keys(o[k]).forEach(function (v) { ct2type[o[k][v]] = k; }); });
  return o;
}());

var type2ct = evert_arr(ct2type);

XMLNS.CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

var CTYPE_XML_ROOT = writextag('Types', null, {
  xmlns: XMLNS.CT,
  'xmlns:xsd': XMLNS.xsd,
  'xmlns:xsi': XMLNS.xsi,
});

var CTYPE_DEFAULTS = [
  ['xml', 'application/xml'],
  ['bin', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main'],
  ['rels', type2ct.rels[0]],
].map(function (x) {
  return writextag('Default', null, { Extension: x[0], ContentType: x[1] });
});

function write_ct(ct, opts) {
  var o = [];
  var v;
  o[o.length] = (XML_HEADER);
  o[o.length] = (CTYPE_XML_ROOT);
  o = o.concat(CTYPE_DEFAULTS);
  var f1 = function (w) {
    if (ct[w] && ct[w].length > 0) {
      v = ct[w][0];
      o[o.length] = (writextag('Override', null, {
        PartName: (v[0] === '/' ? '' : '/') + v,
        ContentType: CT_LIST[w][opts.bookType || 'xlsx'],
      }));
    }
  };
  var f2 = function (w) {
    ct[w].forEach(function (v) {
      o[o.length] = (writextag('Override', null, {
        PartName: (v[0] === '/' ? '' : '/') + v,
        ContentType: CT_LIST[w][opts.bookType || 'xlsx'],
      }));
    });
  };
  var f3 = function (t) {
    (ct[t] || []).forEach(function (v) {
      o[o.length] = (writextag('Override', null, {
        PartName: (v[0] === '/' ? '' : '/') + v,
        ContentType: type2ct[t][0],
      }));
    });
  };
  f1('workbooks');
  f2('sheets');
  f3('themes');
  ['strs', 'styles'].forEach(f1);
  ['coreprops', 'extprops', 'custprops'].forEach(f3);
  if (o.length > 2) { o[o.length] = ('</Types>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}
/* 9.3.2 OPC Relationships Markup */
var RELS = {
  WB: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  SHEET: 'http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
};

XMLNS.RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

var RELS_ROOT = writextag('Relationships', null, {
  xmlns: XMLNS.RELS,
});

/* TODO */
function write_rels(rels) {
  var o = [];
  o[o.length] = (XML_HEADER);
  o[o.length] = (RELS_ROOT);
  keys(rels['!id']).forEach(function (rid) {
    var rel = rels['!id'][rid];
    o[o.length] = (writextag('Relationship', null, rel));
  });
  if (o.length > 2) { o[o.length] = ('</Relationships>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}

var CORE_PROPS = [
  ['cp:category', 'Category'],
  ['cp:contentStatus', 'ContentStatus'],
  ['cp:keywords', 'Keywords'],
  ['cp:lastModifiedBy', 'LastAuthor'],
  ['cp:lastPrinted', 'LastPrinted'],
  ['cp:revision', 'RevNumber'],
  ['cp:version', 'Version'],
  ['dc:creator', 'Author'],
  ['dc:description', 'Comments'],
  ['dc:identifier', 'Identifier'],
  ['dc:language', 'Language'],
  ['dc:subject', 'Subject'],
  ['dc:title', 'Title'],
  ['dcterms:created', 'CreatedDate', 'date'],
  ['dcterms:modified', 'ModifiedDate', 'date'],
];

XMLNS.CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
RELS.CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

var CORE_PROPS_XML_ROOT = writextag('cp:coreProperties', null, {
  'xmlns:cp': XMLNS.CORE_PROPS,
  'xmlns:dc': XMLNS.dc,
  'xmlns:dcterms': XMLNS.dcterms,
  'xmlns:dcmitype': XMLNS.dcmitype,
  'xmlns:xsi': XMLNS.xsi,
});

function cp_doit(f, g, h, o, p) {
  if (p[f] != null || g == null || g === '') return;
  p[f] = g;
  o[o.length] = (h ? writextag(f, g, h) : writetag(f, g));
}

function write_core_props(cp, opts) {
  var o = [XML_HEADER, CORE_PROPS_XML_ROOT];
  var p = {};
  if (!cp) return o.join('');


  if (cp.CreatedDate != null) cp_doit('dcterms:created', typeof cp.CreatedDate === 'string' ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts.WTF), { 'xsi:type': 'dcterms:W3CDTF' }, o, p);
  if (cp.ModifiedDate != null) cp_doit('dcterms:modified', typeof cp.ModifiedDate === 'string' ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts.WTF), { 'xsi:type': 'dcterms:W3CDTF' }, o, p);

  for (var i = 0; i !== CORE_PROPS.length; ++i) { var f = CORE_PROPS[i]; cp_doit(f[0], cp[f[1]], null, o, p); }
  if (o.length > 2) { o[o.length] = ('</cp:coreProperties>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}
/* 15.2.12.3 Extended File Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var EXT_PROPS = [
  ['Application', 'Application', 'string'],
  ['AppVersion', 'AppVersion', 'string'],
  ['Company', 'Company', 'string'],
  ['DocSecurity', 'DocSecurity', 'string'],
  ['Manager', 'Manager', 'string'],
  ['HyperlinksChanged', 'HyperlinksChanged', 'bool'],
  ['SharedDoc', 'SharedDoc', 'bool'],
  ['LinksUpToDate', 'LinksUpToDate', 'bool'],
  ['ScaleCrop', 'ScaleCrop', 'bool'],
  ['HeadingPairs', 'HeadingPairs', 'raw'],
  ['TitlesOfParts', 'TitlesOfParts', 'raw'],
];

XMLNS.EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
RELS.EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

var EXT_PROPS_XML_ROOT = writextag('Properties', null, {
  xmlns: XMLNS.EXT_PROPS,
  'xmlns:vt': XMLNS.vt,
});

function write_ext_props(cp) {
  var o = [];
  var W = writextag;
  if (!cp) cp = {};
  cp.Application = 'SheetJS';
  o[o.length] = (XML_HEADER);
  o[o.length] = (EXT_PROPS_XML_ROOT);

  EXT_PROPS.forEach(function (f) {
    if (cp[f[1]] === undefined) return;
    var v;
    switch (f[2]) {
      case 'string': v = cp[f[1]]; break;
      case 'bool': v = cp[f[1]] ? 'true' : 'false'; break;
    }
    if (v !== undefined) o[o.length] = (W(f[0], v));
  });

  o[o.length] = (W('HeadingPairs', W('vt:vector', W('vt:variant', '<vt:lpstr>Worksheets</vt:lpstr>') + W('vt:variant', W('vt:i4', String(cp.Worksheets))), { size: 2, baseType: 'variant' })));
  o[o.length] = (W('TitlesOfParts', W('vt:vector', cp.SheetNames.map(function (s) { return '<vt:lpstr>' + s + '</vt:lpstr>'; }).join(''), { size: cp.Worksheets, baseType: 'lpstr' })));
  if (o.length > 2) { o[o.length] = ('</Properties>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}
/* 15.2.12.2 Custom File Properties Part */
XMLNS.CUST_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
RELS.CUST_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';

var CUST_PROPS_XML_ROOT = writextag('Properties', null, {
  xmlns: XMLNS.CUST_PROPS,
  'xmlns:vt': XMLNS.vt,
});

function write_cust_props(cp) {
  var o = [XML_HEADER, CUST_PROPS_XML_ROOT];
  if (!cp) return o.join('');
  var pid = 1;
  keys(cp).forEach(function custprop(k) {
    ++pid;
    o[o.length] = (writextag('property', write_vt(cp[k]), {
      fmtid: '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
      pid: pid,
      name: k,
    }));
  });
  if (o.length > 2) { o[o.length] = '</Properties>'; o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}

function parse_FILETIME(blob) {
  var dwLowDateTime = blob.read_shift(4);
  var dwHighDateTime = blob.read_shift(4);
  return new Date(((dwHighDateTime / 1e7 * Math.pow(2, 32) + dwLowDateTime / 1e7) - 11644473600) * 1000).toISOString().replace(/\.000/, '');
}

function parse_lpstr(blob, type, pad) {
  var str = blob.read_shift(0, 'lpstr');
  if (pad) blob.l += (4 - ((str.length + 1) & 3)) & 3;
  return str;
}

/* [MS-OSHARED] 2.3.3.1.6 Lpwstr */
function parse_lpwstr(blob, type, pad) {
  var str = blob.read_shift(0, 'lpwstr');
  if (pad) blob.l += (4 - ((str.length + 1) & 3)) & 3;
  return str;
}


/* [MS-OSHARED] 2.3.3.1.11 VtString */
/* [MS-OSHARED] 2.3.3.1.12 VtUnalignedString */
function parse_VtStringBase(blob, stringType, pad) {
  if (stringType === 0x1F /* VT_LPWSTR*/) return parse_lpwstr(blob);
  return parse_lpstr(blob, stringType, pad);
}

function parse_VtString(blob, t, pad) { return parse_VtStringBase(blob, t, pad === false ? 0 : 4); }
function parse_VtUnalignedString(blob, t) { if (!t) throw new Error('dafuq?'); return parse_VtStringBase(blob, t, 0); }

/* [MS-OSHARED] 2.3.3.1.9 VtVecUnalignedLpstrValue */
function parse_VtVecUnalignedLpstrValue(blob) {
  var length = blob.read_shift(4);
  var ret = [];
  for (var i = 0; i !== length; ++i) ret[i] = blob.read_shift(0, 'lpstr');
  return ret;
}

function parse_VtVecUnalignedLpstr(blob) {
  return parse_VtVecUnalignedLpstrValue(blob);
}

function parse_VtHeadingPair(blob) {
  var headingString = parse_TypedPropertyValue(blob, VT_USTR);
  var headerParts = parse_TypedPropertyValue(blob, VT_I4);
  return [headingString, headerParts];
}

/* [MS-OSHARED] 2.3.3.1.14 VtVecHeadingPairValue */
function parse_VtVecHeadingPairValue(blob) {
  var cElements = blob.read_shift(4);
  var out = [];
  for (var i = 0; i != cElements / 2; ++i) out.push(parse_VtHeadingPair(blob));
  return out;
}

function parse_VtVecHeadingPair(blob) {
  return parse_VtVecHeadingPairValue(blob);
}


/* [MS-OLEPS] 2.9 BLOB */
function parse_BLOB(blob) {
  var size = blob.read_shift(4);
  var bytes = blob.slice(blob.l, blob.l + size);
  if (size & 3 > 0) blob.l += (4 - (size & 3)) & 3;
  return bytes;
}

function parse_ClipboardData(blob) {
  var o = {};
  o.Size = blob.read_shift(4);
  blob.l += o.Size;
  return o;
}

function parse_TypedPropertyValue(blob, type, _opts) {
  var t = blob.read_shift(2);
  var ret;
  var opts = _opts || {};
  blob.l += 2;

  if (type !== VT_VARIANT) {
    if (t !== type && VT_CUSTOM.indexOf(type) === -1) {
      throw new Error('Expected type ' + type + ' saw ' + t);
    }
  }
  switch (type === VT_VARIANT ? t : type) {
    case 0x02 /* VT_I2*/: ret = blob.read_shift(2, 'i'); if (!opts.raw) blob.l += 2; return ret;
    case 0x03 /* VT_I4*/: ret = blob.read_shift(4, 'i'); return ret;
    case 0x0B /* VT_BOOL*/: return blob.read_shift(4) !== 0x0;
    case 0x13 /* VT_UI4*/: ret = blob.read_shift(4); return ret;
    case 0x1E /* VT_LPSTR*/: return parse_lpstr(blob, t, 4).replace(chr0, '');
    case 0x1F /* VT_LPWSTR*/: return parse_lpwstr(blob);
    case 0x40 /* VT_FILETIME*/: return parse_FILETIME(blob);
    case 0x41 /* VT_BLOB*/: return parse_BLOB(blob);
    case 0x47 /* VT_CF*/: return parse_ClipboardData(blob);
    case 0x50 /* VT_STRING*/: return parse_VtString(blob, t, !opts.raw && 4).replace(chr0, '');
    case 0x51 /* VT_USTR*/: return parse_VtUnalignedString(blob, t, 4).replace(chr0, '');
    case 0x100C /* VT_VECTOR|VT_VARIANT*/: return parse_VtVecHeadingPair(blob);
    case 0x101E /* VT_LPSTR*/: return parse_VtVecUnalignedLpstr(blob);
    default: throw new Error('TypedPropertyValue unrecognized type ' + type + ' ' + t);
  }
}

RELS.SST = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
var straywsregex = /^\s|\s$|[\t\n\r]/;
function write_sst_xml(sst, opts) {
  if (!opts.bookSST) return '';
  var o = [XML_HEADER];
  o[o.length] = (writextag('sst', null, {
    xmlns: XMLNS.main[0],
    count: sst.Count,
    uniqueCount: sst.Unique,
  }));
  for (var i = 0; i !== sst.length; ++i) {
    if (sst[i] == null) continue;
    var s = sst[i];
    var sitag = '<si>';
    if (s.r) sitag += s.r;
    else {
      sitag += '<t';
      if (s.t.match(straywsregex)) sitag += ' xml:space="preserve"';
      sitag += '>' + escapexml(s.t) + '</t>';
    }
    sitag += '</si>';
    o[o.length] = (sitag);
  }
  if (o.length > 2) { o[o.length] = ('</sst>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}

function parse_BrtBeginSst(data) {
  return [data.read_shift(4), data.read_shift(4)];
}

function write_BrtBeginSst(sst, o) {
  if (!o) o = new_buf(8);
  o.write_shift(4, sst.Count);
  o.write_shift(4, sst.Unique);
  return o;
}

var write_BrtSSTItem = write_RichStr;

function write_sst_bin(sst) {
  var ba = buf_array();
  write_record(ba, 'BrtBeginSst', write_BrtBeginSst(sst));
  for (var i = 0; i < sst.length; ++i) write_record(ba, 'BrtSSTItem', write_BrtSSTItem(sst[i]));
  write_record(ba, 'BrtEndSst');
  return ba.end();
}

var DEF_MDW = 7;
var MDW = DEF_MDW;
function px2char(px) { return (((px - 5) / MDW * 100 + 0.5) | 0) / 100; }
function char2width(chr) { return (((chr * MDW + 5) / MDW * 256) | 0) / 256; }

function write_numFmts(NF) {
  var o = ['<numFmts>'];
  [
    [5, 8],
    [23, 26],
    [41, 44],
    [63, 66],
    [164, 392],
  ].forEach(function (r) {
    for (var i = r[0]; i <= r[1]; ++i) if (NF[i] !== undefined) o[o.length] = (writextag('numFmt', null, { numFmtId: i, formatCode: escapexml(NF[i]) }));
  });
  if (o.length === 1) return '';
  o[o.length] = ('</numFmts>');
  o[0] = writextag('numFmts', null, { count: o.length - 2 }).replace('/>', '>');
  return o.join('');
}

function write_cellXfs(cellXfs) {
  var o = [];
  o[o.length] = (writextag('cellXfs', null));
  cellXfs.forEach(function (c) {
    o[o.length] = (writextag('xf', null, c));
  });
  o[o.length] = ('</cellXfs>');
  if (o.length === 2) return '';
  o[0] = writextag('cellXfs', null, { count: o.length - 2 }).replace('/>', '>');
  return o.join('');
}

var STYLES_XML_ROOT = writextag('styleSheet', null, {
  xmlns: XMLNS.main[0],
  'xmlns:vt': XMLNS.vt,
});

RELS.STY = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

function write_sty_xml(wb, opts) {
  if (style_builder) {
    return style_builder.toXml();
  }

  var o = [XML_HEADER, STYLES_XML_ROOT];
  var w;
  if ((w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
  o[o.length] = ('<fonts count="1"><font><sz val="8"/><color theme="1"/><name val="Tahoma"/><family val="2"/><scheme val="minor"/></font></fonts>');
  o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
  o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
  o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
  if ((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
  o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
  o[o.length] = ('<dxfs count="0"/>');
  o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

  if (o.length > 2) {
    o[o.length] = ('</styleSheet>');
    o[1] = o[1].replace('/>', '>');
  }
  return o.join('');
}

function parse_BrtFmt(data, length) {
  var ifmt = data.read_shift(2);
  var stFmtCode = parse_XLWideString(data, length - 2);
  return [ifmt, stFmtCode];
}

function parse_BrtFont(data, length) {
  var out = { flags: {} };
  out.dyHeight = data.read_shift(2);
  out.grbit = parse_FontFlags(data, 2);
  out.bls = data.read_shift(2);
  out.sss = data.read_shift(2);
  out.uls = data.read_shift(1);
  out.bFamily = data.read_shift(1);
  out.bCharSet = data.read_shift(1);
  data.l++;
  out.brtColor = parse_BrtColor(data, 8);
  out.bFontScheme = data.read_shift(1);
  out.name = parse_XLWideString(data, length - 21);

  out.flags.Bold = out.bls === 0x02BC;
  out.flags.Italic = out.grbit.fItalic;
  out.flags.Strikeout = out.grbit.fStrikeout;
  out.flags.Outline = out.grbit.fOutline;
  out.flags.Shadow = out.grbit.fShadow;
  out.flags.Condense = out.grbit.fCondense;
  out.flags.Extend = out.grbit.fExtend;
  out.flags.Sub = out.sss & 0x2;
  out.flags.Sup = out.sss & 0x1;
  return out;
}

/* [MS-XLSB] 2.4.816 BrtXF */
function parse_BrtXF(data, length) {
  var ixfeParent = data.read_shift(2);
  var ifmt = data.read_shift(2);
  parsenoop(data, length - 4);
  return { ixfe: ixfeParent, ifmt: ifmt };
}

function write_sty_bin() {
  var ba = buf_array();
  write_record(ba, 'BrtBeginStyleSheet');
  write_record(ba, 'BrtEndStyleSheet');
  return ba.end();
}
RELS.THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

function write_theme(opts) {
  if (opts.themeXml) { return opts.themeXml; }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:spDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style></a:spDef><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/></a:theme>';
}

function parse_BrtCalcChainItem$(data) {
  var out = {};
  out.i = data.read_shift(4);
  var cell = {};
  cell.r = data.read_shift(4);
  cell.c = data.read_shift(4);
  out.r = encodeCell(cell);
  var flags = data.read_shift(1);
  if (flags & 0x2) out.l = '1';
  if (flags & 0x8) out.a = '1';
  return out;
}

function parse_BrtBeginComment(data) {
  var out = {};
  out.iauthor = data.read_shift(4);
  var rfx = parse_UncheckedRfX(data, 16);
  out.rfx = rfx.s;
  out.ref = encodeCell(rfx.s);
  data.l += 16;
  return out;
}

var parse_BrtCommentAuthor = parse_XLWideString;
var parse_BrtCommentText = parse_RichStr;

function parse_XLSBCellParsedFormula(data, length) {
  data.read_shift(4);
  return parsenoop(data, length - 4);
}

RELS.WS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

function get_sst_id(sst, str) {
  var len = sst.length;
  for (var i = 0; i < len; ++i) if (sst[i].t === str) { sst.Count ++; return i; }
  sst[len] = { t: str };
  sst.Count ++;
  sst.Unique ++;
  return len;
}

function get_cell_style(styles, cell, opts) {
  if (style_builder) {
    if (/^\d+$/.exec(cell.s)) { return cell.s; }  // if its already an integer index, let it be
    if (cell.s && (cell.s == +cell.s)) { return cell.s; }  // if its already an integer index, let it be
    var s = cell.s || {};
    if (cell.z) s.numFmt = cell.z;
    return style_builder.addStyle(s);
  }

  var z = opts.revssf[cell.z != null ? cell.z : 'General'];
  var len = styles.length;
  for (var i = 0; i !== len; ++i) if (styles[i].numFmtId === z) return i;
  styles[len] = {
    numFmtId: z,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    applyNumberFormat: 1,
  };
  return len;
}

function write_ws_xml_merges(merges) {
  if (merges.length === 0) return '';
  var o = '<mergeCells count="' + merges.length + '">';
  for (var i = 0; i !== merges.length; ++i) o += '<mergeCell ref="' + encodeRange(merges[i]) + '"/>';
  return o + '</mergeCells>';
}

function write_ws_xml_cols(ws, cols) {
  var o = ['<cols>'];
  var col;
  var width;
  for (var i = 0; i !== cols.length; ++i) {
    if (!(col = cols[i])) continue;
    var p = { min: i + 1, max: i + 1 };
    width = -1;
    if (col.wpx) width = px2char(col.wpx);
    else if (col.wch) width = col.wch;
    if (width > -1) { p.width = char2width(width); p.customWidth = 1; }
    o[o.length] = (writextag('col', null, p));
  }
  o[o.length] = '</cols>';
  return o.join('');
}

function write_ws_xml_cell(cell, ref, ws, opts) {
  if (cell.v === undefined && cell.s === undefined) return '';
  var vv = '';
  var oldt = cell.t;
  var oldv = cell.v;
  switch (cell.t) {
    case 'b': vv = cell.v ? '1' : '0'; break;
    case 'n': vv = '' + cell.v; break;
    case 'e': vv = BErr[cell.v]; break;
    case 'd':
      cell.t = 'n';
      vv = '' + (cell.v = datenum(cell.v));
      if (cell.z === undefined) {
        if (cell.s && cell.s.numFmt) cell.z = cell.s.numFmt;
        else cell.z = SSF._table[14];
      }
      break;
    default: vv = cell.v; break;
  }
  var v = writetag(cell.v.toString().charAt(0) === '=' ? 'f' : 'v', escapexml(vv));
  var o = { r: ref };
  var os = get_cell_style(opts.cellXfs, cell, opts);
  if (os !== 0) o.s = os;
  switch (cell.t) {
    case 'n': break;
    case 'd': o.t = 'd'; break;
    case 'b': o.t = 'b'; break;
    case 'e': o.t = 'e'; break;
    default:
      if (opts.bookSST) {
        v = writetag('v', '' + get_sst_id(opts.Strings, cell.v));
        o.t = 's'; break;
      }
      o.t = 'str'; break;
  }
  if (cell.t !== oldt) {
    cell.t = oldt;
    cell.v = oldv;
  }
  return writextag('c', v, o);
}

function write_ws_xml_data(ws, opts, idx, wb) {
  var o = [];
  var r = [];
  var range = safeDecodeRange(ws['!ref']);
  var cell;
  var ref;
  var rr = '';
  var cols = [];
  var R;
  var C;
  for (C = range.s.c; C <= range.e.c; ++C) cols[C] = encodeCol(C);
  for (R = range.s.r; R <= range.e.r; ++R) {
    r = [];
    rr = encodeRow(R);
    for (C = range.s.c; C <= range.e.c; ++C) {
      ref = cols[C] + rr;
      if (ws[ref] === undefined) continue;
      if ((cell = write_ws_xml_cell(ws[ref], ref, ws, opts, idx, wb)) != null) r.push(cell);
    }
    if (r.length > 0) o[o.length] = (writextag('row', r.join(''), { r: rr }));
  }
  return o.join('');
}

var WS_XML_ROOT = writextag('worksheet', null, {
  xmlns: XMLNS.main[0],
  'xmlns:r': XMLNS.r,
});

function write_ws_xml(idx, opts, wb) {
  var o = [XML_HEADER, WS_XML_ROOT];
  var s = wb.SheetNames[idx];
  var sidx = 0;
  var rdata = '';
  var ws = wb.Sheets[s];
  if (ws === undefined) ws = {};
  var ref = ws['!ref']; if (ref === undefined) ref = 'A1';
  o[o.length] = (writextag('dimension', null, { ref: ref }));

  if (ws['!cols'] !== undefined && ws['!cols'].length > 0) o[o.length] = (write_ws_xml_cols(ws, ws['!cols']));
  o[sidx = o.length] = '<sheetData/>';
  if (ws['!ref'] !== undefined) {
    rdata = write_ws_xml_data(ws, opts, idx, wb);
    if (rdata.length > 0) o[o.length] = (rdata);
  }
  if (o.length > sidx + 1) { o[o.length] = ('</sheetData>'); o[sidx] = o[sidx].replace('/>', '>'); }

  if (ws['!merges'] !== undefined && ws['!merges'].length > 0) o[o.length] = (write_ws_xml_merges(ws['!merges']));

  if (o.length > 2) { o[o.length] = ('</worksheet>'); o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}

function parse_BrtRowHdr(data, length) {
  var z = [];
  z.r = data.read_shift(4);
  data.l += length - 4;
  return z;
}

var parse_BrtWsDim = parse_UncheckedRfX;
var write_BrtWsDim = write_UncheckedRfX;

function parse_BrtWsProp(data, length) {
  var z = {};
  data.l += 19;
  z.name = parse_XLSBCodeName(data, length - 19);
  return z;
}

function parse_BrtCellBlank(data) {
  var cell = parse_XLSBCell(data);
  return [cell];
}
function write_BrtCellBlank(cell, val, o) {
  if (o == null) o = new_buf(8);
  return write_XLSBCell(val, o);
}


function parse_BrtCellBool(data) {
  var cell = parse_XLSBCell(data);
  var fBool = data.read_shift(1);
  return [cell, fBool, 'b'];
}

function parse_BrtCellError(data) {
  var cell = parse_XLSBCell(data);
  var fBool = data.read_shift(1);
  return [cell, fBool, 'e'];
}

function parse_BrtCellIsst(data) {
  var cell = parse_XLSBCell(data);
  var isst = data.read_shift(4);
  return [cell, isst, 's'];
}

function parse_BrtCellReal(data) {
  var cell = parse_XLSBCell(data);
  var value = parse_Xnum(data);
  return [cell, value, 'n'];
}

function parse_BrtCellRk(data) {
  var cell = parse_XLSBCell(data);
  var value = parse_RkNumber(data);
  return [cell, value, 'n'];
}

function parse_BrtCellSt(data) {
  var cell = parse_XLSBCell(data);
  var value = parse_XLWideString(data);
  return [cell, value, 'str'];
}

function parse_BrtFmlaBool(data, length, opts) {
  var cell = parse_XLSBCell(data);
  var value = data.read_shift(1);
  var o = [cell, value, 'b'];
  if (opts.cellFormula) {
    parse_XLSBCellParsedFormula(data, length - 9);
    o[3] = '';
  } else {
    data.l += length - 9;
  }
  return o;
}

function parse_BrtFmlaError(data, length, opts) {
  var cell = parse_XLSBCell(data);
  var value = data.read_shift(1);
  var o = [cell, value, 'e'];
  if (opts.cellFormula) {
    parse_XLSBCellParsedFormula(data, length - 9);
    o[3] = '';
  } else {
    data.l += length - 9;
  }
  return o;
}

function parse_BrtFmlaNum(data, length, opts) {
  var cell = parse_XLSBCell(data);
  var value = parse_Xnum(data);
  var o = [cell, value, 'n'];
  if (opts.cellFormula) {
    parse_XLSBCellParsedFormula(data, length - 16);
    o[3] = '';
  } else {
    data.l += length - 16;
  }
  return o;
}

function parse_BrtFmlaString(data, length, opts) {
  var start = data.l;
  var cell = parse_XLSBCell(data);
  var value = parse_XLWideString(data);
  var o = [cell, value, 'str'];
  if (opts.cellFormula) {
    parse_XLSBCellParsedFormula(data, start + length - data.l);
  } else {
    data.l = start + length;
  }
  return o;
}

var parse_BrtMergeCell = parse_UncheckedRfX;

function parse_BrtHLink(data, length) {
  var end = data.l + length;
  var rfx = parse_UncheckedRfX(data, 16);
  var relId = parse_XLNullableWideString(data);
  var loc = parse_XLWideString(data);
  var tooltip = parse_XLWideString(data);
  var display = parse_XLWideString(data);
  data.l = end;
  return { rfx: rfx, relId: relId, loc: loc, tooltip: tooltip, display: display };
}

function write_ws_bin_cell(ba, cell, R, C, opts) {
  if (cell.v == null) return;
  switch (cell.t) {
    case 'b': vv = cell.v ? '1' : '0'; break;
    case 'n': case 'e': vv = '' + cell.v; break;
    default: vv = cell.v; break;
  }
  var o = { r: R, c: C };
  o.s = get_cell_style(opts.cellXfs, cell, opts);
  switch (cell.t) {
    case 's': case 'str':
      if (opts.bookSST) {
        vv = get_sst_id(opts.Strings, cell.v);
        o.t = 's'; break;
      }
      o.t = 'str'; break;
    case 'n': break;
    case 'b': o.t = 'b'; break;
    case 'e': o.t = 'e'; break;
  }
  write_record(ba, 'BrtCellBlank', write_BrtCellBlank(cell, o));
}

function write_CELLTABLE(ba, ws, idx, opts) {
  var range = safeDecodeRange(ws['!ref'] || 'A1');
  var ref;
  var rr = '';
  var cols = [];
  write_record(ba, 'BrtBeginSheetData');
  for (var R = range.s.r; R <= range.e.r; ++R) {
    rr = encodeRow(R);
    for (var C = range.s.c; C <= range.e.c; ++C) {
      if (R === range.s.r) cols[C] = encodeCol(C);
      ref = cols[C] + rr;
      if (!ws[ref]) continue;
      write_ws_bin_cell(ba, ws[ref], R, C, opts);
    }
  }
  write_record(ba, 'BrtEndSheetData');
}

function write_ws_bin(idx, opts, wb) {
  var ba = buf_array();
  var s = wb.SheetNames[idx];
  var ws = wb.Sheets[s] || {};
  var r = safeDecodeRange(ws['!ref'] || 'A1');
  write_record(ba, 'BrtBeginSheet');
  write_record(ba, 'BrtWsDim', write_BrtWsDim(r));
  write_CELLTABLE(ba, ws, idx, opts, wb);
  write_record(ba, 'BrtEndSheet');
  return ba.end();
}

var WB_XML_ROOT = writextag('workbook', null, {
  xmlns: XMLNS.main[0],
  'xmlns:r': XMLNS.r,
});

function safe1904(wb) {
  try { return parsexmlbool(wb.Workbook.WBProps.date1904) ? 'true' : 'false'; } catch (e) { return 'false'; }
}

function write_wb_xml(wb) {
  var o = [XML_HEADER];
  o[o.length] = WB_XML_ROOT;
  o[o.length] = (writextag('workbookPr', null, { date1904: safe1904(wb) }));
  o[o.length] = '<sheets>';
  for (var i = 0; i !== wb.SheetNames.length; ++i) {
    o[o.length] = writextag('sheet', null, {
      name: wb.SheetNames[i].substr(0, 31),
      sheetId: '' + (i + 1),
      'r:id': 'rId' + (i + 1),
    });
  }
  o[o.length] = '</sheets>';
  if (o.length > 2) { o[o.length] = '</workbook>'; o[1] = o[1].replace('/>', '>'); }
  return o.join('');
}
/* [MS-XLSB] 2.4.301 BrtBundleSh */
function parse_BrtBundleSh(data, length) {
  var z = {};
  z.hsState = data.read_shift(4); // ST_SheetState
  z.iTabID = data.read_shift(4);
  z.strRelID = parse_RelID(data, length - 8);
  z.name = parse_XLWideString(data);
  return z;
}
function write_BrtBundleSh(data, o) {
  if (!o) o = new_buf(127);
  o.write_shift(4, data.hsState);
  o.write_shift(4, data.iTabID);
  write_RelID(data.strRelID, o);
  write_XLWideString(data.name.substr(0, 31), o);
  return o;
}

function parse_BrtWbProp(data, length) {
  data.read_shift(4);
  var dwThemeVersion = data.read_shift(4);
  var strName = (length > 8) ? parse_XLWideString(data) : '';
  return [dwThemeVersion, strName];
}

function write_BrtWbProp(data, o) {
  if (!o) o = new_buf(8);
  o.write_shift(4, 0);
  o.write_shift(4, 0);
  return o;
}

function parse_BrtFRTArchID$(data, length) {
  var o = {};
  data.read_shift(4);
  o.ArchID = data.read_shift(4);
  data.l += length - 8;
  return o;
}

function write_BUNDLESHS(ba, wb) {
  write_record(ba, 'BrtBeginBundleShs');
  for (var idx = 0; idx !== wb.SheetNames.length; ++idx) {
    var d = { hsState: 0, iTabID: idx + 1, strRelID: 'rId' + (idx + 1), name: wb.SheetNames[idx] };
    write_record(ba, 'BrtBundleSh', write_BrtBundleSh(d));
  }
  write_record(ba, 'BrtEndBundleShs');
}

function write_BrtFileVersion(data, o) {
  if (!o) o = new_buf(127);
  for (var i = 0; i !== 4; ++i) o.write_shift(4, 0);
  write_XLWideString('SheetJS', o);
  // write_XLWideString(XLSX.version, o);
  // write_XLWideString(XLSX.version, o);
  write_XLWideString('7262', o);
  o.length = o.l;
  return o;
}

function write_BOOKVIEWS(ba) {
  write_record(ba, 'BrtBeginBookViews');
  write_record(ba, 'BrtEndBookViews');
}

function write_BrtCalcProp(data, o) {
  if (!o) o = new_buf(26);
  o.write_shift(4, 0); /* force recalc */
  o.write_shift(4, 1);
  o.write_shift(4, 0);
  write_Xnum(0, o);
  o.write_shift(-4, 1023);
  o.write_shift(1, 0x33);
  o.write_shift(1, 0x00);
  return o;
}

function write_BrtFileRecover(data, o) {
  if (!o) o = new_buf(1);
  o.write_shift(1, 0);
  return o;
}

function write_wb_bin(wb, opts) {
  var ba = buf_array();
  write_record(ba, 'BrtBeginBook');
  write_record(ba, 'BrtFileVersion', write_BrtFileVersion());
  write_record(ba, 'BrtWbProp', write_BrtWbProp());
  write_BOOKVIEWS(ba, wb, opts);
  write_BUNDLESHS(ba, wb, opts);
  write_record(ba, 'BrtCalcProp', write_BrtCalcProp());
  write_record(ba, 'BrtFileRecover', write_BrtFileRecover());
  write_record(ba, 'BrtEndBook');

  return ba.end();
}

function write_wb(wb, name, opts) {
  return (name.substr(-4) === '.bin' ? write_wb_bin : write_wb_xml)(wb, opts);
}

function write_ws(data, name, opts, wb) {
  return (name.substr(-4) === '.bin' ? write_ws_bin : write_ws_xml)(data, opts, wb);
}

function write_sty(data, name, opts) {
  return (name.substr(-4) === '.bin' ? write_sty_bin : write_sty_xml)(data, opts);
}

function write_sst(data, name, opts) {
  return (name.substr(-4) === '.bin' ? write_sst_bin : write_sst_xml)(data, opts);
}

var XLSBRecordEnum = {
  0x0000: { n: 'BrtRowHdr', f: parse_BrtRowHdr },
  0x0001: { n: 'BrtCellBlank', f: parse_BrtCellBlank },
  0x0002: { n: 'BrtCellRk', f: parse_BrtCellRk },
  0x0003: { n: 'BrtCellError', f: parse_BrtCellError },
  0x0004: { n: 'BrtCellBool', f: parse_BrtCellBool },
  0x0005: { n: 'BrtCellReal', f: parse_BrtCellReal },
  0x0006: { n: 'BrtCellSt', f: parse_BrtCellSt },
  0x0007: { n: 'BrtCellIsst', f: parse_BrtCellIsst },
  0x0008: { n: 'BrtFmlaString', f: parse_BrtFmlaString },
  0x0009: { n: 'BrtFmlaNum', f: parse_BrtFmlaNum },
  0x000A: { n: 'BrtFmlaBool', f: parse_BrtFmlaBool },
  0x000B: { n: 'BrtFmlaError', f: parse_BrtFmlaError },
  0x0010: { n: 'BrtFRTArchID$', f: parse_BrtFRTArchID$ },
  0x0013: { n: 'BrtSSTItem', f: parse_RichStr },
  0x0014: { n: 'BrtPCDIMissing', f: parsenoop },
  0x0015: { n: 'BrtPCDINumber', f: parsenoop },
  0x0016: { n: 'BrtPCDIBoolean', f: parsenoop },
  0x0017: { n: 'BrtPCDIError', f: parsenoop },
  0x0018: { n: 'BrtPCDIString', f: parsenoop },
  0x0019: { n: 'BrtPCDIDatetime', f: parsenoop },
  0x001A: { n: 'BrtPCDIIndex', f: parsenoop },
  0x001B: { n: 'BrtPCDIAMissing', f: parsenoop },
  0x001C: { n: 'BrtPCDIANumber', f: parsenoop },
  0x001D: { n: 'BrtPCDIABoolean', f: parsenoop },
  0x001E: { n: 'BrtPCDIAError', f: parsenoop },
  0x001F: { n: 'BrtPCDIAString', f: parsenoop },
  0x0020: { n: 'BrtPCDIADatetime', f: parsenoop },
  0x0021: { n: 'BrtPCRRecord', f: parsenoop },
  0x0022: { n: 'BrtPCRRecordDt', f: parsenoop },
  0x0023: { n: 'BrtFRTBegin', f: parsenoop },
  0x0024: { n: 'BrtFRTEnd', f: parsenoop },
  0x0025: { n: 'BrtACBegin', f: parsenoop },
  0x0026: { n: 'BrtACEnd', f: parsenoop },
  0x0027: { n: 'BrtName', f: parsenoop },
  0x0028: { n: 'BrtIndexRowBlock', f: parsenoop },
  0x002A: { n: 'BrtIndexBlock', f: parsenoop },
  0x002B: { n: 'BrtFont', f: parse_BrtFont },
  0x002C: { n: 'BrtFmt', f: parse_BrtFmt },
  0x002D: { n: 'BrtFill', f: parsenoop },
  0x002E: { n: 'BrtBorder', f: parsenoop },
  0x002F: { n: 'BrtXF', f: parse_BrtXF },
  0x0030: { n: 'BrtStyle', f: parsenoop },
  0x0031: { n: 'BrtCellMeta', f: parsenoop },
  0x0032: { n: 'BrtValueMeta', f: parsenoop },
  0x0033: { n: 'BrtMdb', f: parsenoop },
  0x0034: { n: 'BrtBeginFmd', f: parsenoop },
  0x0035: { n: 'BrtEndFmd', f: parsenoop },
  0x0036: { n: 'BrtBeginMdx', f: parsenoop },
  0x0037: { n: 'BrtEndMdx', f: parsenoop },
  0x0038: { n: 'BrtBeginMdxTuple', f: parsenoop },
  0x0039: { n: 'BrtEndMdxTuple', f: parsenoop },
  0x003A: { n: 'BrtMdxMbrIstr', f: parsenoop },
  0x003B: { n: 'BrtStr', f: parsenoop },
  0x003C: { n: 'BrtColInfo', f: parsenoop },
  0x003E: { n: 'BrtCellRString', f: parsenoop },
  0x003F: { n: 'BrtCalcChainItem$', f: parse_BrtCalcChainItem$ },
  0x0040: { n: 'BrtDVal', f: parsenoop },
  0x0041: { n: 'BrtSxvcellNum', f: parsenoop },
  0x0042: { n: 'BrtSxvcellStr', f: parsenoop },
  0x0043: { n: 'BrtSxvcellBool', f: parsenoop },
  0x0044: { n: 'BrtSxvcellErr', f: parsenoop },
  0x0045: { n: 'BrtSxvcellDate', f: parsenoop },
  0x0046: { n: 'BrtSxvcellNil', f: parsenoop },
  0x0080: { n: 'BrtFileVersion', f: parsenoop },
  0x0081: { n: 'BrtBeginSheet', f: parsenoop },
  0x0082: { n: 'BrtEndSheet', f: parsenoop },
  0x0083: { n: 'BrtBeginBook', f: parsenoop, p: 0 },
  0x0084: { n: 'BrtEndBook', f: parsenoop },
  0x0085: { n: 'BrtBeginWsViews', f: parsenoop },
  0x0086: { n: 'BrtEndWsViews', f: parsenoop },
  0x0087: { n: 'BrtBeginBookViews', f: parsenoop },
  0x0088: { n: 'BrtEndBookViews', f: parsenoop },
  0x0089: { n: 'BrtBeginWsView', f: parsenoop },
  0x008A: { n: 'BrtEndWsView', f: parsenoop },
  0x008B: { n: 'BrtBeginCsViews', f: parsenoop },
  0x008C: { n: 'BrtEndCsViews', f: parsenoop },
  0x008D: { n: 'BrtBeginCsView', f: parsenoop },
  0x008E: { n: 'BrtEndCsView', f: parsenoop },
  0x008F: { n: 'BrtBeginBundleShs', f: parsenoop },
  0x0090: { n: 'BrtEndBundleShs', f: parsenoop },
  0x0091: { n: 'BrtBeginSheetData', f: parsenoop },
  0x0092: { n: 'BrtEndSheetData', f: parsenoop },
  0x0093: { n: 'BrtWsProp', f: parse_BrtWsProp },
  0x0094: { n: 'BrtWsDim', f: parse_BrtWsDim, p: 16 },
  0x0097: { n: 'BrtPane', f: parsenoop },
  0x0098: { n: 'BrtSel', f: parsenoop },
  0x0099: { n: 'BrtWbProp', f: parse_BrtWbProp },
  0x009A: { n: 'BrtWbFactoid', f: parsenoop },
  0x009B: { n: 'BrtFileRecover', f: parsenoop },
  0x009C: { n: 'BrtBundleSh', f: parse_BrtBundleSh },
  0x009D: { n: 'BrtCalcProp', f: parsenoop },
  0x009E: { n: 'BrtBookView', f: parsenoop },
  0x009F: { n: 'BrtBeginSst', f: parse_BrtBeginSst },
  0x00A0: { n: 'BrtEndSst', f: parsenoop },
  0x00A1: { n: 'BrtBeginAFilter', f: parsenoop },
  0x00A2: { n: 'BrtEndAFilter', f: parsenoop },
  0x00A3: { n: 'BrtBeginFilterColumn', f: parsenoop },
  0x00A4: { n: 'BrtEndFilterColumn', f: parsenoop },
  0x00A5: { n: 'BrtBeginFilters', f: parsenoop },
  0x00A6: { n: 'BrtEndFilters', f: parsenoop },
  0x00A7: { n: 'BrtFilter', f: parsenoop },
  0x00A8: { n: 'BrtColorFilter', f: parsenoop },
  0x00A9: { n: 'BrtIconFilter', f: parsenoop },
  0x00AA: { n: 'BrtTop10Filter', f: parsenoop },
  0x00AB: { n: 'BrtDynamicFilter', f: parsenoop },
  0x00AC: { n: 'BrtBeginCustomFilters', f: parsenoop },
  0x00AD: { n: 'BrtEndCustomFilters', f: parsenoop },
  0x00AE: { n: 'BrtCustomFilter', f: parsenoop },
  0x00AF: { n: 'BrtAFilterDateGroupItem', f: parsenoop },
  0x00B0: { n: 'BrtMergeCell', f: parse_BrtMergeCell },
  0x00B1: { n: 'BrtBeginMergeCells', f: parsenoop },
  0x00B2: { n: 'BrtEndMergeCells', f: parsenoop },
  0x00B3: { n: 'BrtBeginPivotCacheDef', f: parsenoop },
  0x00B4: { n: 'BrtEndPivotCacheDef', f: parsenoop },
  0x00B5: { n: 'BrtBeginPCDFields', f: parsenoop },
  0x00B6: { n: 'BrtEndPCDFields', f: parsenoop },
  0x00B7: { n: 'BrtBeginPCDField', f: parsenoop },
  0x00B8: { n: 'BrtEndPCDField', f: parsenoop },
  0x00B9: { n: 'BrtBeginPCDSource', f: parsenoop },
  0x00BA: { n: 'BrtEndPCDSource', f: parsenoop },
  0x00BB: { n: 'BrtBeginPCDSRange', f: parsenoop },
  0x00BC: { n: 'BrtEndPCDSRange', f: parsenoop },
  0x00BD: { n: 'BrtBeginPCDFAtbl', f: parsenoop },
  0x00BE: { n: 'BrtEndPCDFAtbl', f: parsenoop },
  0x00BF: { n: 'BrtBeginPCDIRun', f: parsenoop },
  0x00C0: { n: 'BrtEndPCDIRun', f: parsenoop },
  0x00C1: { n: 'BrtBeginPivotCacheRecords', f: parsenoop },
  0x00C2: { n: 'BrtEndPivotCacheRecords', f: parsenoop },
  0x00C3: { n: 'BrtBeginPCDHierarchies', f: parsenoop },
  0x00C4: { n: 'BrtEndPCDHierarchies', f: parsenoop },
  0x00C5: { n: 'BrtBeginPCDHierarchy', f: parsenoop },
  0x00C6: { n: 'BrtEndPCDHierarchy', f: parsenoop },
  0x00C7: { n: 'BrtBeginPCDHFieldsUsage', f: parsenoop },
  0x00C8: { n: 'BrtEndPCDHFieldsUsage', f: parsenoop },
  0x00C9: { n: 'BrtBeginExtConnection', f: parsenoop },
  0x00CA: { n: 'BrtEndExtConnection', f: parsenoop },
  0x00CB: { n: 'BrtBeginECDbProps', f: parsenoop },
  0x00CC: { n: 'BrtEndECDbProps', f: parsenoop },
  0x00CD: { n: 'BrtBeginECOlapProps', f: parsenoop },
  0x00CE: { n: 'BrtEndECOlapProps', f: parsenoop },
  0x00CF: { n: 'BrtBeginPCDSConsol', f: parsenoop },
  0x00D0: { n: 'BrtEndPCDSConsol', f: parsenoop },
  0x00D1: { n: 'BrtBeginPCDSCPages', f: parsenoop },
  0x00D2: { n: 'BrtEndPCDSCPages', f: parsenoop },
  0x00D3: { n: 'BrtBeginPCDSCPage', f: parsenoop },
  0x00D4: { n: 'BrtEndPCDSCPage', f: parsenoop },
  0x00D5: { n: 'BrtBeginPCDSCPItem', f: parsenoop },
  0x00D6: { n: 'BrtEndPCDSCPItem', f: parsenoop },
  0x00D7: { n: 'BrtBeginPCDSCSets', f: parsenoop },
  0x00D8: { n: 'BrtEndPCDSCSets', f: parsenoop },
  0x00D9: { n: 'BrtBeginPCDSCSet', f: parsenoop },
  0x00DA: { n: 'BrtEndPCDSCSet', f: parsenoop },
  0x00DB: { n: 'BrtBeginPCDFGroup', f: parsenoop },
  0x00DC: { n: 'BrtEndPCDFGroup', f: parsenoop },
  0x00DD: { n: 'BrtBeginPCDFGItems', f: parsenoop },
  0x00DE: { n: 'BrtEndPCDFGItems', f: parsenoop },
  0x00DF: { n: 'BrtBeginPCDFGRange', f: parsenoop },
  0x00E0: { n: 'BrtEndPCDFGRange', f: parsenoop },
  0x00E1: { n: 'BrtBeginPCDFGDiscrete', f: parsenoop },
  0x00E2: { n: 'BrtEndPCDFGDiscrete', f: parsenoop },
  0x00E3: { n: 'BrtBeginPCDSDTupleCache', f: parsenoop },
  0x00E4: { n: 'BrtEndPCDSDTupleCache', f: parsenoop },
  0x00E5: { n: 'BrtBeginPCDSDTCEntries', f: parsenoop },
  0x00E6: { n: 'BrtEndPCDSDTCEntries', f: parsenoop },
  0x00E7: { n: 'BrtBeginPCDSDTCEMembers', f: parsenoop },
  0x00E8: { n: 'BrtEndPCDSDTCEMembers', f: parsenoop },
  0x00E9: { n: 'BrtBeginPCDSDTCEMember', f: parsenoop },
  0x00EA: { n: 'BrtEndPCDSDTCEMember', f: parsenoop },
  0x00EB: { n: 'BrtBeginPCDSDTCQueries', f: parsenoop },
  0x00EC: { n: 'BrtEndPCDSDTCQueries', f: parsenoop },
  0x00ED: { n: 'BrtBeginPCDSDTCQuery', f: parsenoop },
  0x00EE: { n: 'BrtEndPCDSDTCQuery', f: parsenoop },
  0x00EF: { n: 'BrtBeginPCDSDTCSets', f: parsenoop },
  0x00F0: { n: 'BrtEndPCDSDTCSets', f: parsenoop },
  0x00F1: { n: 'BrtBeginPCDSDTCSet', f: parsenoop },
  0x00F2: { n: 'BrtEndPCDSDTCSet', f: parsenoop },
  0x00F3: { n: 'BrtBeginPCDCalcItems', f: parsenoop },
  0x00F4: { n: 'BrtEndPCDCalcItems', f: parsenoop },
  0x00F5: { n: 'BrtBeginPCDCalcItem', f: parsenoop },
  0x00F6: { n: 'BrtEndPCDCalcItem', f: parsenoop },
  0x00F7: { n: 'BrtBeginPRule', f: parsenoop },
  0x00F8: { n: 'BrtEndPRule', f: parsenoop },
  0x00F9: { n: 'BrtBeginPRFilters', f: parsenoop },
  0x00FA: { n: 'BrtEndPRFilters', f: parsenoop },
  0x00FB: { n: 'BrtBeginPRFilter', f: parsenoop },
  0x00FC: { n: 'BrtEndPRFilter', f: parsenoop },
  0x00FD: { n: 'BrtBeginPNames', f: parsenoop },
  0x00FE: { n: 'BrtEndPNames', f: parsenoop },
  0x00FF: { n: 'BrtBeginPName', f: parsenoop },
  0x0100: { n: 'BrtEndPName', f: parsenoop },
  0x0101: { n: 'BrtBeginPNPairs', f: parsenoop },
  0x0102: { n: 'BrtEndPNPairs', f: parsenoop },
  0x0103: { n: 'BrtBeginPNPair', f: parsenoop },
  0x0104: { n: 'BrtEndPNPair', f: parsenoop },
  0x0105: { n: 'BrtBeginECWebProps', f: parsenoop },
  0x0106: { n: 'BrtEndECWebProps', f: parsenoop },
  0x0107: { n: 'BrtBeginEcWpTables', f: parsenoop },
  0x0108: { n: 'BrtEndECWPTables', f: parsenoop },
  0x0109: { n: 'BrtBeginECParams', f: parsenoop },
  0x010A: { n: 'BrtEndECParams', f: parsenoop },
  0x010B: { n: 'BrtBeginECParam', f: parsenoop },
  0x010C: { n: 'BrtEndECParam', f: parsenoop },
  0x010D: { n: 'BrtBeginPCDKPIs', f: parsenoop },
  0x010E: { n: 'BrtEndPCDKPIs', f: parsenoop },
  0x010F: { n: 'BrtBeginPCDKPI', f: parsenoop },
  0x0110: { n: 'BrtEndPCDKPI', f: parsenoop },
  0x0111: { n: 'BrtBeginDims', f: parsenoop },
  0x0112: { n: 'BrtEndDims', f: parsenoop },
  0x0113: { n: 'BrtBeginDim', f: parsenoop },
  0x0114: { n: 'BrtEndDim', f: parsenoop },
  0x0115: { n: 'BrtIndexPartEnd', f: parsenoop },
  0x0116: { n: 'BrtBeginStyleSheet', f: parsenoop },
  0x0117: { n: 'BrtEndStyleSheet', f: parsenoop },
  0x0118: { n: 'BrtBeginSXView', f: parsenoop },
  0x0119: { n: 'BrtEndSXVI', f: parsenoop },
  0x011A: { n: 'BrtBeginSXVI', f: parsenoop },
  0x011B: { n: 'BrtBeginSXVIs', f: parsenoop },
  0x011C: { n: 'BrtEndSXVIs', f: parsenoop },
  0x011D: { n: 'BrtBeginSXVD', f: parsenoop },
  0x011E: { n: 'BrtEndSXVD', f: parsenoop },
  0x011F: { n: 'BrtBeginSXVDs', f: parsenoop },
  0x0120: { n: 'BrtEndSXVDs', f: parsenoop },
  0x0121: { n: 'BrtBeginSXPI', f: parsenoop },
  0x0122: { n: 'BrtEndSXPI', f: parsenoop },
  0x0123: { n: 'BrtBeginSXPIs', f: parsenoop },
  0x0124: { n: 'BrtEndSXPIs', f: parsenoop },
  0x0125: { n: 'BrtBeginSXDI', f: parsenoop },
  0x0126: { n: 'BrtEndSXDI', f: parsenoop },
  0x0127: { n: 'BrtBeginSXDIs', f: parsenoop },
  0x0128: { n: 'BrtEndSXDIs', f: parsenoop },
  0x0129: { n: 'BrtBeginSXLI', f: parsenoop },
  0x012A: { n: 'BrtEndSXLI', f: parsenoop },
  0x012B: { n: 'BrtBeginSXLIRws', f: parsenoop },
  0x012C: { n: 'BrtEndSXLIRws', f: parsenoop },
  0x012D: { n: 'BrtBeginSXLICols', f: parsenoop },
  0x012E: { n: 'BrtEndSXLICols', f: parsenoop },
  0x012F: { n: 'BrtBeginSXFormat', f: parsenoop },
  0x0130: { n: 'BrtEndSXFormat', f: parsenoop },
  0x0131: { n: 'BrtBeginSXFormats', f: parsenoop },
  0x0132: { n: 'BrtEndSxFormats', f: parsenoop },
  0x0133: { n: 'BrtBeginSxSelect', f: parsenoop },
  0x0134: { n: 'BrtEndSxSelect', f: parsenoop },
  0x0135: { n: 'BrtBeginISXVDRws', f: parsenoop },
  0x0136: { n: 'BrtEndISXVDRws', f: parsenoop },
  0x0137: { n: 'BrtBeginISXVDCols', f: parsenoop },
  0x0138: { n: 'BrtEndISXVDCols', f: parsenoop },
  0x0139: { n: 'BrtEndSXLocation', f: parsenoop },
  0x013A: { n: 'BrtBeginSXLocation', f: parsenoop },
  0x013B: { n: 'BrtEndSXView', f: parsenoop },
  0x013C: { n: 'BrtBeginSXTHs', f: parsenoop },
  0x013D: { n: 'BrtEndSXTHs', f: parsenoop },
  0x013E: { n: 'BrtBeginSXTH', f: parsenoop },
  0x013F: { n: 'BrtEndSXTH', f: parsenoop },
  0x0140: { n: 'BrtBeginISXTHRws', f: parsenoop },
  0x0141: { n: 'BrtEndISXTHRws', f: parsenoop },
  0x0142: { n: 'BrtBeginISXTHCols', f: parsenoop },
  0x0143: { n: 'BrtEndISXTHCols', f: parsenoop },
  0x0144: { n: 'BrtBeginSXTDMPS', f: parsenoop },
  0x0145: { n: 'BrtEndSXTDMPs', f: parsenoop },
  0x0146: { n: 'BrtBeginSXTDMP', f: parsenoop },
  0x0147: { n: 'BrtEndSXTDMP', f: parsenoop },
  0x0148: { n: 'BrtBeginSXTHItems', f: parsenoop },
  0x0149: { n: 'BrtEndSXTHItems', f: parsenoop },
  0x014A: { n: 'BrtBeginSXTHItem', f: parsenoop },
  0x014B: { n: 'BrtEndSXTHItem', f: parsenoop },
  0x014C: { n: 'BrtBeginMetadata', f: parsenoop },
  0x014D: { n: 'BrtEndMetadata', f: parsenoop },
  0x014E: { n: 'BrtBeginEsmdtinfo', f: parsenoop },
  0x014F: { n: 'BrtMdtinfo', f: parsenoop },
  0x0150: { n: 'BrtEndEsmdtinfo', f: parsenoop },
  0x0151: { n: 'BrtBeginEsmdb', f: parsenoop },
  0x0152: { n: 'BrtEndEsmdb', f: parsenoop },
  0x0153: { n: 'BrtBeginEsfmd', f: parsenoop },
  0x0154: { n: 'BrtEndEsfmd', f: parsenoop },
  0x0155: { n: 'BrtBeginSingleCells', f: parsenoop },
  0x0156: { n: 'BrtEndSingleCells', f: parsenoop },
  0x0157: { n: 'BrtBeginList', f: parsenoop },
  0x0158: { n: 'BrtEndList', f: parsenoop },
  0x0159: { n: 'BrtBeginListCols', f: parsenoop },
  0x015A: { n: 'BrtEndListCols', f: parsenoop },
  0x015B: { n: 'BrtBeginListCol', f: parsenoop },
  0x015C: { n: 'BrtEndListCol', f: parsenoop },
  0x015D: { n: 'BrtBeginListXmlCPr', f: parsenoop },
  0x015E: { n: 'BrtEndListXmlCPr', f: parsenoop },
  0x015F: { n: 'BrtListCCFmla', f: parsenoop },
  0x0160: { n: 'BrtListTrFmla', f: parsenoop },
  0x0161: { n: 'BrtBeginExternals', f: parsenoop },
  0x0162: { n: 'BrtEndExternals', f: parsenoop },
  0x0163: { n: 'BrtSupBookSrc', f: parsenoop },
  0x0165: { n: 'BrtSupSelf', f: parsenoop },
  0x0166: { n: 'BrtSupSame', f: parsenoop },
  0x0167: { n: 'BrtSupTabs', f: parsenoop },
  0x0168: { n: 'BrtBeginSupBook', f: parsenoop },
  0x0169: { n: 'BrtPlaceholderName', f: parsenoop },
  0x016A: { n: 'BrtExternSheet', f: parsenoop },
  0x016B: { n: 'BrtExternTableStart', f: parsenoop },
  0x016C: { n: 'BrtExternTableEnd', f: parsenoop },
  0x016E: { n: 'BrtExternRowHdr', f: parsenoop },
  0x016F: { n: 'BrtExternCellBlank', f: parsenoop },
  0x0170: { n: 'BrtExternCellReal', f: parsenoop },
  0x0171: { n: 'BrtExternCellBool', f: parsenoop },
  0x0172: { n: 'BrtExternCellError', f: parsenoop },
  0x0173: { n: 'BrtExternCellString', f: parsenoop },
  0x0174: { n: 'BrtBeginEsmdx', f: parsenoop },
  0x0175: { n: 'BrtEndEsmdx', f: parsenoop },
  0x0176: { n: 'BrtBeginMdxSet', f: parsenoop },
  0x0177: { n: 'BrtEndMdxSet', f: parsenoop },
  0x0178: { n: 'BrtBeginMdxMbrProp', f: parsenoop },
  0x0179: { n: 'BrtEndMdxMbrProp', f: parsenoop },
  0x017A: { n: 'BrtBeginMdxKPI', f: parsenoop },
  0x017B: { n: 'BrtEndMdxKPI', f: parsenoop },
  0x017C: { n: 'BrtBeginEsstr', f: parsenoop },
  0x017D: { n: 'BrtEndEsstr', f: parsenoop },
  0x017E: { n: 'BrtBeginPRFItem', f: parsenoop },
  0x017F: { n: 'BrtEndPRFItem', f: parsenoop },
  0x0180: { n: 'BrtBeginPivotCacheIDs', f: parsenoop },
  0x0181: { n: 'BrtEndPivotCacheIDs', f: parsenoop },
  0x0182: { n: 'BrtBeginPivotCacheID', f: parsenoop },
  0x0183: { n: 'BrtEndPivotCacheID', f: parsenoop },
  0x0184: { n: 'BrtBeginISXVIs', f: parsenoop },
  0x0185: { n: 'BrtEndISXVIs', f: parsenoop },
  0x0186: { n: 'BrtBeginColInfos', f: parsenoop },
  0x0187: { n: 'BrtEndColInfos', f: parsenoop },
  0x0188: { n: 'BrtBeginRwBrk', f: parsenoop },
  0x0189: { n: 'BrtEndRwBrk', f: parsenoop },
  0x018A: { n: 'BrtBeginColBrk', f: parsenoop },
  0x018B: { n: 'BrtEndColBrk', f: parsenoop },
  0x018C: { n: 'BrtBrk', f: parsenoop },
  0x018D: { n: 'BrtUserBookView', f: parsenoop },
  0x018E: { n: 'BrtInfo', f: parsenoop },
  0x018F: { n: 'BrtCUsr', f: parsenoop },
  0x0190: { n: 'BrtUsr', f: parsenoop },
  0x0191: { n: 'BrtBeginUsers', f: parsenoop },
  0x0193: { n: 'BrtEOF', f: parsenoop },
  0x0194: { n: 'BrtUCR', f: parsenoop },
  0x0195: { n: 'BrtRRInsDel', f: parsenoop },
  0x0196: { n: 'BrtRREndInsDel', f: parsenoop },
  0x0197: { n: 'BrtRRMove', f: parsenoop },
  0x0198: { n: 'BrtRREndMove', f: parsenoop },
  0x0199: { n: 'BrtRRChgCell', f: parsenoop },
  0x019A: { n: 'BrtRREndChgCell', f: parsenoop },
  0x019B: { n: 'BrtRRHeader', f: parsenoop },
  0x019C: { n: 'BrtRRUserView', f: parsenoop },
  0x019D: { n: 'BrtRRRenSheet', f: parsenoop },
  0x019E: { n: 'BrtRRInsertSh', f: parsenoop },
  0x019F: { n: 'BrtRRDefName', f: parsenoop },
  0x01A0: { n: 'BrtRRNote', f: parsenoop },
  0x01A1: { n: 'BrtRRConflict', f: parsenoop },
  0x01A2: { n: 'BrtRRTQSIF', f: parsenoop },
  0x01A3: { n: 'BrtRRFormat', f: parsenoop },
  0x01A4: { n: 'BrtRREndFormat', f: parsenoop },
  0x01A5: { n: 'BrtRRAutoFmt', f: parsenoop },
  0x01A6: { n: 'BrtBeginUserShViews', f: parsenoop },
  0x01A7: { n: 'BrtBeginUserShView', f: parsenoop },
  0x01A8: { n: 'BrtEndUserShView', f: parsenoop },
  0x01A9: { n: 'BrtEndUserShViews', f: parsenoop },
  0x01AA: { n: 'BrtArrFmla', f: parsenoop },
  0x01AB: { n: 'BrtShrFmla', f: parsenoop },
  0x01AC: { n: 'BrtTable', f: parsenoop },
  0x01AD: { n: 'BrtBeginExtConnections', f: parsenoop },
  0x01AE: { n: 'BrtEndExtConnections', f: parsenoop },
  0x01AF: { n: 'BrtBeginPCDCalcMems', f: parsenoop },
  0x01B0: { n: 'BrtEndPCDCalcMems', f: parsenoop },
  0x01B1: { n: 'BrtBeginPCDCalcMem', f: parsenoop },
  0x01B2: { n: 'BrtEndPCDCalcMem', f: parsenoop },
  0x01B3: { n: 'BrtBeginPCDHGLevels', f: parsenoop },
  0x01B4: { n: 'BrtEndPCDHGLevels', f: parsenoop },
  0x01B5: { n: 'BrtBeginPCDHGLevel', f: parsenoop },
  0x01B6: { n: 'BrtEndPCDHGLevel', f: parsenoop },
  0x01B7: { n: 'BrtBeginPCDHGLGroups', f: parsenoop },
  0x01B8: { n: 'BrtEndPCDHGLGroups', f: parsenoop },
  0x01B9: { n: 'BrtBeginPCDHGLGroup', f: parsenoop },
  0x01BA: { n: 'BrtEndPCDHGLGroup', f: parsenoop },
  0x01BB: { n: 'BrtBeginPCDHGLGMembers', f: parsenoop },
  0x01BC: { n: 'BrtEndPCDHGLGMembers', f: parsenoop },
  0x01BD: { n: 'BrtBeginPCDHGLGMember', f: parsenoop },
  0x01BE: { n: 'BrtEndPCDHGLGMember', f: parsenoop },
  0x01BF: { n: 'BrtBeginQSI', f: parsenoop },
  0x01C0: { n: 'BrtEndQSI', f: parsenoop },
  0x01C1: { n: 'BrtBeginQSIR', f: parsenoop },
  0x01C2: { n: 'BrtEndQSIR', f: parsenoop },
  0x01C3: { n: 'BrtBeginDeletedNames', f: parsenoop },
  0x01C4: { n: 'BrtEndDeletedNames', f: parsenoop },
  0x01C5: { n: 'BrtBeginDeletedName', f: parsenoop },
  0x01C6: { n: 'BrtEndDeletedName', f: parsenoop },
  0x01C7: { n: 'BrtBeginQSIFs', f: parsenoop },
  0x01C8: { n: 'BrtEndQSIFs', f: parsenoop },
  0x01C9: { n: 'BrtBeginQSIF', f: parsenoop },
  0x01CA: { n: 'BrtEndQSIF', f: parsenoop },
  0x01CB: { n: 'BrtBeginAutoSortScope', f: parsenoop },
  0x01CC: { n: 'BrtEndAutoSortScope', f: parsenoop },
  0x01CD: { n: 'BrtBeginConditionalFormatting', f: parsenoop },
  0x01CE: { n: 'BrtEndConditionalFormatting', f: parsenoop },
  0x01CF: { n: 'BrtBeginCFRule', f: parsenoop },
  0x01D0: { n: 'BrtEndCFRule', f: parsenoop },
  0x01D1: { n: 'BrtBeginIconSet', f: parsenoop },
  0x01D2: { n: 'BrtEndIconSet', f: parsenoop },
  0x01D3: { n: 'BrtBeginDatabar', f: parsenoop },
  0x01D4: { n: 'BrtEndDatabar', f: parsenoop },
  0x01D5: { n: 'BrtBeginColorScale', f: parsenoop },
  0x01D6: { n: 'BrtEndColorScale', f: parsenoop },
  0x01D7: { n: 'BrtCFVO', f: parsenoop },
  0x01D8: { n: 'BrtExternValueMeta', f: parsenoop },
  0x01D9: { n: 'BrtBeginColorPalette', f: parsenoop },
  0x01DA: { n: 'BrtEndColorPalette', f: parsenoop },
  0x01DB: { n: 'BrtIndexedColor', f: parsenoop },
  0x01DC: { n: 'BrtMargins', f: parsenoop },
  0x01DD: { n: 'BrtPrintOptions', f: parsenoop },
  0x01DE: { n: 'BrtPageSetup', f: parsenoop },
  0x01DF: { n: 'BrtBeginHeaderFooter', f: parsenoop },
  0x01E0: { n: 'BrtEndHeaderFooter', f: parsenoop },
  0x01E1: { n: 'BrtBeginSXCrtFormat', f: parsenoop },
  0x01E2: { n: 'BrtEndSXCrtFormat', f: parsenoop },
  0x01E3: { n: 'BrtBeginSXCrtFormats', f: parsenoop },
  0x01E4: { n: 'BrtEndSXCrtFormats', f: parsenoop },
  0x01E5: { n: 'BrtWsFmtInfo', f: parsenoop },
  0x01E6: { n: 'BrtBeginMgs', f: parsenoop },
  0x01E7: { n: 'BrtEndMGs', f: parsenoop },
  0x01E8: { n: 'BrtBeginMGMaps', f: parsenoop },
  0x01E9: { n: 'BrtEndMGMaps', f: parsenoop },
  0x01EA: { n: 'BrtBeginMG', f: parsenoop },
  0x01EB: { n: 'BrtEndMG', f: parsenoop },
  0x01EC: { n: 'BrtBeginMap', f: parsenoop },
  0x01ED: { n: 'BrtEndMap', f: parsenoop },
  0x01EE: { n: 'BrtHLink', f: parse_BrtHLink },
  0x01EF: { n: 'BrtBeginDCon', f: parsenoop },
  0x01F0: { n: 'BrtEndDCon', f: parsenoop },
  0x01F1: { n: 'BrtBeginDRefs', f: parsenoop },
  0x01F2: { n: 'BrtEndDRefs', f: parsenoop },
  0x01F3: { n: 'BrtDRef', f: parsenoop },
  0x01F4: { n: 'BrtBeginScenMan', f: parsenoop },
  0x01F5: { n: 'BrtEndScenMan', f: parsenoop },
  0x01F6: { n: 'BrtBeginSct', f: parsenoop },
  0x01F7: { n: 'BrtEndSct', f: parsenoop },
  0x01F8: { n: 'BrtSlc', f: parsenoop },
  0x01F9: { n: 'BrtBeginDXFs', f: parsenoop },
  0x01FA: { n: 'BrtEndDXFs', f: parsenoop },
  0x01FB: { n: 'BrtDXF', f: parsenoop },
  0x01FC: { n: 'BrtBeginTableStyles', f: parsenoop },
  0x01FD: { n: 'BrtEndTableStyles', f: parsenoop },
  0x01FE: { n: 'BrtBeginTableStyle', f: parsenoop },
  0x01FF: { n: 'BrtEndTableStyle', f: parsenoop },
  0x0200: { n: 'BrtTableStyleElement', f: parsenoop },
  0x0201: { n: 'BrtTableStyleClient', f: parsenoop },
  0x0202: { n: 'BrtBeginVolDeps', f: parsenoop },
  0x0203: { n: 'BrtEndVolDeps', f: parsenoop },
  0x0204: { n: 'BrtBeginVolType', f: parsenoop },
  0x0205: { n: 'BrtEndVolType', f: parsenoop },
  0x0206: { n: 'BrtBeginVolMain', f: parsenoop },
  0x0207: { n: 'BrtEndVolMain', f: parsenoop },
  0x0208: { n: 'BrtBeginVolTopic', f: parsenoop },
  0x0209: { n: 'BrtEndVolTopic', f: parsenoop },
  0x020A: { n: 'BrtVolSubtopic', f: parsenoop },
  0x020B: { n: 'BrtVolRef', f: parsenoop },
  0x020C: { n: 'BrtVolNum', f: parsenoop },
  0x020D: { n: 'BrtVolErr', f: parsenoop },
  0x020E: { n: 'BrtVolStr', f: parsenoop },
  0x020F: { n: 'BrtVolBool', f: parsenoop },
  0x0210: { n: 'BrtBeginCalcChain$', f: parsenoop },
  0x0211: { n: 'BrtEndCalcChain$', f: parsenoop },
  0x0212: { n: 'BrtBeginSortState', f: parsenoop },
  0x0213: { n: 'BrtEndSortState', f: parsenoop },
  0x0214: { n: 'BrtBeginSortCond', f: parsenoop },
  0x0215: { n: 'BrtEndSortCond', f: parsenoop },
  0x0216: { n: 'BrtBookProtection', f: parsenoop },
  0x0217: { n: 'BrtSheetProtection', f: parsenoop },
  0x0218: { n: 'BrtRangeProtection', f: parsenoop },
  0x0219: { n: 'BrtPhoneticInfo', f: parsenoop },
  0x021A: { n: 'BrtBeginECTxtWiz', f: parsenoop },
  0x021B: { n: 'BrtEndECTxtWiz', f: parsenoop },
  0x021C: { n: 'BrtBeginECTWFldInfoLst', f: parsenoop },
  0x021D: { n: 'BrtEndECTWFldInfoLst', f: parsenoop },
  0x021E: { n: 'BrtBeginECTwFldInfo', f: parsenoop },
  0x0224: { n: 'BrtFileSharing', f: parsenoop },
  0x0225: { n: 'BrtOleSize', f: parsenoop },
  0x0226: { n: 'BrtDrawing', f: parsenoop },
  0x0227: { n: 'BrtLegacyDrawing', f: parsenoop },
  0x0228: { n: 'BrtLegacyDrawingHF', f: parsenoop },
  0x0229: { n: 'BrtWebOpt', f: parsenoop },
  0x022A: { n: 'BrtBeginWebPubItems', f: parsenoop },
  0x022B: { n: 'BrtEndWebPubItems', f: parsenoop },
  0x022C: { n: 'BrtBeginWebPubItem', f: parsenoop },
  0x022D: { n: 'BrtEndWebPubItem', f: parsenoop },
  0x022E: { n: 'BrtBeginSXCondFmt', f: parsenoop },
  0x022F: { n: 'BrtEndSXCondFmt', f: parsenoop },
  0x0230: { n: 'BrtBeginSXCondFmts', f: parsenoop },
  0x0231: { n: 'BrtEndSXCondFmts', f: parsenoop },
  0x0232: { n: 'BrtBkHim', f: parsenoop },
  0x0234: { n: 'BrtColor', f: parsenoop },
  0x0235: { n: 'BrtBeginIndexedColors', f: parsenoop },
  0x0236: { n: 'BrtEndIndexedColors', f: parsenoop },
  0x0239: { n: 'BrtBeginMRUColors', f: parsenoop },
  0x023A: { n: 'BrtEndMRUColors', f: parsenoop },
  0x023C: { n: 'BrtMRUColor', f: parsenoop },
  0x023D: { n: 'BrtBeginDVals', f: parsenoop },
  0x023E: { n: 'BrtEndDVals', f: parsenoop },
  0x0241: { n: 'BrtSupNameStart', f: parsenoop },
  0x0242: { n: 'BrtSupNameValueStart', f: parsenoop },
  0x0243: { n: 'BrtSupNameValueEnd', f: parsenoop },
  0x0244: { n: 'BrtSupNameNum', f: parsenoop },
  0x0245: { n: 'BrtSupNameErr', f: parsenoop },
  0x0246: { n: 'BrtSupNameSt', f: parsenoop },
  0x0247: { n: 'BrtSupNameNil', f: parsenoop },
  0x0248: { n: 'BrtSupNameBool', f: parsenoop },
  0x0249: { n: 'BrtSupNameFmla', f: parsenoop },
  0x024A: { n: 'BrtSupNameBits', f: parsenoop },
  0x024B: { n: 'BrtSupNameEnd', f: parsenoop },
  0x024C: { n: 'BrtEndSupBook', f: parsenoop },
  0x024D: { n: 'BrtCellSmartTagProperty', f: parsenoop },
  0x024E: { n: 'BrtBeginCellSmartTag', f: parsenoop },
  0x024F: { n: 'BrtEndCellSmartTag', f: parsenoop },
  0x0250: { n: 'BrtBeginCellSmartTags', f: parsenoop },
  0x0251: { n: 'BrtEndCellSmartTags', f: parsenoop },
  0x0252: { n: 'BrtBeginSmartTags', f: parsenoop },
  0x0253: { n: 'BrtEndSmartTags', f: parsenoop },
  0x0254: { n: 'BrtSmartTagType', f: parsenoop },
  0x0255: { n: 'BrtBeginSmartTagTypes', f: parsenoop },
  0x0256: { n: 'BrtEndSmartTagTypes', f: parsenoop },
  0x0257: { n: 'BrtBeginSXFilters', f: parsenoop },
  0x0258: { n: 'BrtEndSXFilters', f: parsenoop },
  0x0259: { n: 'BrtBeginSXFILTER', f: parsenoop },
  0x025A: { n: 'BrtEndSXFilter', f: parsenoop },
  0x025B: { n: 'BrtBeginFills', f: parsenoop },
  0x025C: { n: 'BrtEndFills', f: parsenoop },
  0x025D: { n: 'BrtBeginCellWatches', f: parsenoop },
  0x025E: { n: 'BrtEndCellWatches', f: parsenoop },
  0x025F: { n: 'BrtCellWatch', f: parsenoop },
  0x0260: { n: 'BrtBeginCRErrs', f: parsenoop },
  0x0261: { n: 'BrtEndCRErrs', f: parsenoop },
  0x0262: { n: 'BrtCrashRecErr', f: parsenoop },
  0x0263: { n: 'BrtBeginFonts', f: parsenoop },
  0x0264: { n: 'BrtEndFonts', f: parsenoop },
  0x0265: { n: 'BrtBeginBorders', f: parsenoop },
  0x0266: { n: 'BrtEndBorders', f: parsenoop },
  0x0267: { n: 'BrtBeginFmts', f: parsenoop },
  0x0268: { n: 'BrtEndFmts', f: parsenoop },
  0x0269: { n: 'BrtBeginCellXFs', f: parsenoop },
  0x026A: { n: 'BrtEndCellXFs', f: parsenoop },
  0x026B: { n: 'BrtBeginStyles', f: parsenoop },
  0x026C: { n: 'BrtEndStyles', f: parsenoop },
  0x0271: { n: 'BrtBigName', f: parsenoop },
  0x0272: { n: 'BrtBeginCellStyleXFs', f: parsenoop },
  0x0273: { n: 'BrtEndCellStyleXFs', f: parsenoop },
  0x0274: { n: 'BrtBeginComments', f: parsenoop },
  0x0275: { n: 'BrtEndComments', f: parsenoop },
  0x0276: { n: 'BrtBeginCommentAuthors', f: parsenoop },
  0x0277: { n: 'BrtEndCommentAuthors', f: parsenoop },
  0x0278: { n: 'BrtCommentAuthor', f: parse_BrtCommentAuthor },
  0x0279: { n: 'BrtBeginCommentList', f: parsenoop },
  0x027A: { n: 'BrtEndCommentList', f: parsenoop },
  0x027B: { n: 'BrtBeginComment', f: parse_BrtBeginComment },
  0x027C: { n: 'BrtEndComment', f: parsenoop },
  0x027D: { n: 'BrtCommentText', f: parse_BrtCommentText },
  0x027E: { n: 'BrtBeginOleObjects', f: parsenoop },
  0x027F: { n: 'BrtOleObject', f: parsenoop },
  0x0280: { n: 'BrtEndOleObjects', f: parsenoop },
  0x0281: { n: 'BrtBeginSxrules', f: parsenoop },
  0x0282: { n: 'BrtEndSxRules', f: parsenoop },
  0x0283: { n: 'BrtBeginActiveXControls', f: parsenoop },
  0x0284: { n: 'BrtActiveX', f: parsenoop },
  0x0285: { n: 'BrtEndActiveXControls', f: parsenoop },
  0x0286: { n: 'BrtBeginPCDSDTCEMembersSortBy', f: parsenoop },
  0x0288: { n: 'BrtBeginCellIgnoreECs', f: parsenoop },
  0x0289: { n: 'BrtCellIgnoreEC', f: parsenoop },
  0x028A: { n: 'BrtEndCellIgnoreECs', f: parsenoop },
  0x028B: { n: 'BrtCsProp', f: parsenoop },
  0x028C: { n: 'BrtCsPageSetup', f: parsenoop },
  0x028D: { n: 'BrtBeginUserCsViews', f: parsenoop },
  0x028E: { n: 'BrtEndUserCsViews', f: parsenoop },
  0x028F: { n: 'BrtBeginUserCsView', f: parsenoop },
  0x0290: { n: 'BrtEndUserCsView', f: parsenoop },
  0x0291: { n: 'BrtBeginPcdSFCIEntries', f: parsenoop },
  0x0292: { n: 'BrtEndPCDSFCIEntries', f: parsenoop },
  0x0293: { n: 'BrtPCDSFCIEntry', f: parsenoop },
  0x0294: { n: 'BrtBeginListParts', f: parsenoop },
  0x0295: { n: 'BrtListPart', f: parsenoop },
  0x0296: { n: 'BrtEndListParts', f: parsenoop },
  0x0297: { n: 'BrtSheetCalcProp', f: parsenoop },
  0x0298: { n: 'BrtBeginFnGroup', f: parsenoop },
  0x0299: { n: 'BrtFnGroup', f: parsenoop },
  0x029A: { n: 'BrtEndFnGroup', f: parsenoop },
  0x029B: { n: 'BrtSupAddin', f: parsenoop },
  0x029C: { n: 'BrtSXTDMPOrder', f: parsenoop },
  0x029D: { n: 'BrtCsProtection', f: parsenoop },
  0x029F: { n: 'BrtBeginWsSortMap', f: parsenoop },
  0x02A0: { n: 'BrtEndWsSortMap', f: parsenoop },
  0x02A1: { n: 'BrtBeginRRSort', f: parsenoop },
  0x02A2: { n: 'BrtEndRRSort', f: parsenoop },
  0x02A3: { n: 'BrtRRSortItem', f: parsenoop },
  0x02A4: { n: 'BrtFileSharingIso', f: parsenoop },
  0x02A5: { n: 'BrtBookProtectionIso', f: parsenoop },
  0x02A6: { n: 'BrtSheetProtectionIso', f: parsenoop },
  0x02A7: { n: 'BrtCsProtectionIso', f: parsenoop },
  0x02A8: { n: 'BrtRangeProtectionIso', f: parsenoop },
  0x0400: { n: 'BrtRwDescent', f: parsenoop },
  0x0401: { n: 'BrtKnownFonts', f: parsenoop },
  0x0402: { n: 'BrtBeginSXTupleSet', f: parsenoop },
  0x0403: { n: 'BrtEndSXTupleSet', f: parsenoop },
  0x0404: { n: 'BrtBeginSXTupleSetHeader', f: parsenoop },
  0x0405: { n: 'BrtEndSXTupleSetHeader', f: parsenoop },
  0x0406: { n: 'BrtSXTupleSetHeaderItem', f: parsenoop },
  0x0407: { n: 'BrtBeginSXTupleSetData', f: parsenoop },
  0x0408: { n: 'BrtEndSXTupleSetData', f: parsenoop },
  0x0409: { n: 'BrtBeginSXTupleSetRow', f: parsenoop },
  0x040A: { n: 'BrtEndSXTupleSetRow', f: parsenoop },
  0x040B: { n: 'BrtSXTupleSetRowItem', f: parsenoop },
  0x040C: { n: 'BrtNameExt', f: parsenoop },
  0x040D: { n: 'BrtPCDH14', f: parsenoop },
  0x040E: { n: 'BrtBeginPCDCalcMem14', f: parsenoop },
  0x040F: { n: 'BrtEndPCDCalcMem14', f: parsenoop },
  0x0410: { n: 'BrtSXTH14', f: parsenoop },
  0x0411: { n: 'BrtBeginSparklineGroup', f: parsenoop },
  0x0412: { n: 'BrtEndSparklineGroup', f: parsenoop },
  0x0413: { n: 'BrtSparkline', f: parsenoop },
  0x0414: { n: 'BrtSXDI14', f: parsenoop },
  0x0415: { n: 'BrtWsFmtInfoEx14', f: parsenoop },
  0x0416: { n: 'BrtBeginConditionalFormatting14', f: parsenoop },
  0x0417: { n: 'BrtEndConditionalFormatting14', f: parsenoop },
  0x0418: { n: 'BrtBeginCFRule14', f: parsenoop },
  0x0419: { n: 'BrtEndCFRule14', f: parsenoop },
  0x041A: { n: 'BrtCFVO14', f: parsenoop },
  0x041B: { n: 'BrtBeginDatabar14', f: parsenoop },
  0x041C: { n: 'BrtBeginIconSet14', f: parsenoop },
  0x041D: { n: 'BrtDVal14', f: parsenoop },
  0x041E: { n: 'BrtBeginDVals14', f: parsenoop },
  0x041F: { n: 'BrtColor14', f: parsenoop },
  0x0420: { n: 'BrtBeginSparklines', f: parsenoop },
  0x0421: { n: 'BrtEndSparklines', f: parsenoop },
  0x0422: { n: 'BrtBeginSparklineGroups', f: parsenoop },
  0x0423: { n: 'BrtEndSparklineGroups', f: parsenoop },
  0x0425: { n: 'BrtSXVD14', f: parsenoop },
  0x0426: { n: 'BrtBeginSxview14', f: parsenoop },
  0x0427: { n: 'BrtEndSxview14', f: parsenoop },
  0x042A: { n: 'BrtBeginPCD14', f: parsenoop },
  0x042B: { n: 'BrtEndPCD14', f: parsenoop },
  0x042C: { n: 'BrtBeginExtConn14', f: parsenoop },
  0x042D: { n: 'BrtEndExtConn14', f: parsenoop },
  0x042E: { n: 'BrtBeginSlicerCacheIDs', f: parsenoop },
  0x042F: { n: 'BrtEndSlicerCacheIDs', f: parsenoop },
  0x0430: { n: 'BrtBeginSlicerCacheID', f: parsenoop },
  0x0431: { n: 'BrtEndSlicerCacheID', f: parsenoop },
  0x0433: { n: 'BrtBeginSlicerCache', f: parsenoop },
  0x0434: { n: 'BrtEndSlicerCache', f: parsenoop },
  0x0435: { n: 'BrtBeginSlicerCacheDef', f: parsenoop },
  0x0436: { n: 'BrtEndSlicerCacheDef', f: parsenoop },
  0x0437: { n: 'BrtBeginSlicersEx', f: parsenoop },
  0x0438: { n: 'BrtEndSlicersEx', f: parsenoop },
  0x0439: { n: 'BrtBeginSlicerEx', f: parsenoop },
  0x043A: { n: 'BrtEndSlicerEx', f: parsenoop },
  0x043B: { n: 'BrtBeginSlicer', f: parsenoop },
  0x043C: { n: 'BrtEndSlicer', f: parsenoop },
  0x043D: { n: 'BrtSlicerCachePivotTables', f: parsenoop },
  0x043E: { n: 'BrtBeginSlicerCacheOlapImpl', f: parsenoop },
  0x043F: { n: 'BrtEndSlicerCacheOlapImpl', f: parsenoop },
  0x0440: { n: 'BrtBeginSlicerCacheLevelsData', f: parsenoop },
  0x0441: { n: 'BrtEndSlicerCacheLevelsData', f: parsenoop },
  0x0442: { n: 'BrtBeginSlicerCacheLevelData', f: parsenoop },
  0x0443: { n: 'BrtEndSlicerCacheLevelData', f: parsenoop },
  0x0444: { n: 'BrtBeginSlicerCacheSiRanges', f: parsenoop },
  0x0445: { n: 'BrtEndSlicerCacheSiRanges', f: parsenoop },
  0x0446: { n: 'BrtBeginSlicerCacheSiRange', f: parsenoop },
  0x0447: { n: 'BrtEndSlicerCacheSiRange', f: parsenoop },
  0x0448: { n: 'BrtSlicerCacheOlapItem', f: parsenoop },
  0x0449: { n: 'BrtBeginSlicerCacheSelections', f: parsenoop },
  0x044A: { n: 'BrtSlicerCacheSelection', f: parsenoop },
  0x044B: { n: 'BrtEndSlicerCacheSelections', f: parsenoop },
  0x044C: { n: 'BrtBeginSlicerCacheNative', f: parsenoop },
  0x044D: { n: 'BrtEndSlicerCacheNative', f: parsenoop },
  0x044E: { n: 'BrtSlicerCacheNativeItem', f: parsenoop },
  0x044F: { n: 'BrtRangeProtection14', f: parsenoop },
  0x0450: { n: 'BrtRangeProtectionIso14', f: parsenoop },
  0x0451: { n: 'BrtCellIgnoreEC14', f: parsenoop },
  0x0457: { n: 'BrtList14', f: parsenoop },
  0x0458: { n: 'BrtCFIcon', f: parsenoop },
  0x0459: { n: 'BrtBeginSlicerCachesPivotCacheIDs', f: parsenoop },
  0x045A: { n: 'BrtEndSlicerCachesPivotCacheIDs', f: parsenoop },
  0x045B: { n: 'BrtBeginSlicers', f: parsenoop },
  0x045C: { n: 'BrtEndSlicers', f: parsenoop },
  0x045D: { n: 'BrtWbProp14', f: parsenoop },
  0x045E: { n: 'BrtBeginSXEdit', f: parsenoop },
  0x045F: { n: 'BrtEndSXEdit', f: parsenoop },
  0x0460: { n: 'BrtBeginSXEdits', f: parsenoop },
  0x0461: { n: 'BrtEndSXEdits', f: parsenoop },
  0x0462: { n: 'BrtBeginSXChange', f: parsenoop },
  0x0463: { n: 'BrtEndSXChange', f: parsenoop },
  0x0464: { n: 'BrtBeginSXChanges', f: parsenoop },
  0x0465: { n: 'BrtEndSXChanges', f: parsenoop },
  0x0466: { n: 'BrtSXTupleItems', f: parsenoop },
  0x0468: { n: 'BrtBeginSlicerStyle', f: parsenoop },
  0x0469: { n: 'BrtEndSlicerStyle', f: parsenoop },
  0x046A: { n: 'BrtSlicerStyleElement', f: parsenoop },
  0x046B: { n: 'BrtBeginStyleSheetExt14', f: parsenoop },
  0x046C: { n: 'BrtEndStyleSheetExt14', f: parsenoop },
  0x046D: { n: 'BrtBeginSlicerCachesPivotCacheID', f: parsenoop },
  0x046E: { n: 'BrtEndSlicerCachesPivotCacheID', f: parsenoop },
  0x046F: { n: 'BrtBeginConditionalFormattings', f: parsenoop },
  0x0470: { n: 'BrtEndConditionalFormattings', f: parsenoop },
  0x0471: { n: 'BrtBeginPCDCalcMemExt', f: parsenoop },
  0x0472: { n: 'BrtEndPCDCalcMemExt', f: parsenoop },
  0x0473: { n: 'BrtBeginPCDCalcMemsExt', f: parsenoop },
  0x0474: { n: 'BrtEndPCDCalcMemsExt', f: parsenoop },
  0x0475: { n: 'BrtPCDField14', f: parsenoop },
  0x0476: { n: 'BrtBeginSlicerStyles', f: parsenoop },
  0x0477: { n: 'BrtEndSlicerStyles', f: parsenoop },
  0x0478: { n: 'BrtBeginSlicerStyleElements', f: parsenoop },
  0x0479: { n: 'BrtEndSlicerStyleElements', f: parsenoop },
  0x047A: { n: 'BrtCFRuleExt', f: parsenoop },
  0x047B: { n: 'BrtBeginSXCondFmt14', f: parsenoop },
  0x047C: { n: 'BrtEndSXCondFmt14', f: parsenoop },
  0x047D: { n: 'BrtBeginSXCondFmts14', f: parsenoop },
  0x047E: { n: 'BrtEndSXCondFmts14', f: parsenoop },
  0x0480: { n: 'BrtBeginSortCond14', f: parsenoop },
  0x0481: { n: 'BrtEndSortCond14', f: parsenoop },
  0x0482: { n: 'BrtEndDVals14', f: parsenoop },
  0x0483: { n: 'BrtEndIconSet14', f: parsenoop },
  0x0484: { n: 'BrtEndDatabar14', f: parsenoop },
  0x0485: { n: 'BrtBeginColorScale14', f: parsenoop },
  0x0486: { n: 'BrtEndColorScale14', f: parsenoop },
  0x0487: { n: 'BrtBeginSxrules14', f: parsenoop },
  0x0488: { n: 'BrtEndSxrules14', f: parsenoop },
  0x0489: { n: 'BrtBeginPRule14', f: parsenoop },
  0x048A: { n: 'BrtEndPRule14', f: parsenoop },
  0x048B: { n: 'BrtBeginPRFilters14', f: parsenoop },
  0x048C: { n: 'BrtEndPRFilters14', f: parsenoop },
  0x048D: { n: 'BrtBeginPRFilter14', f: parsenoop },
  0x048E: { n: 'BrtEndPRFilter14', f: parsenoop },
  0x048F: { n: 'BrtBeginPRFItem14', f: parsenoop },
  0x0490: { n: 'BrtEndPRFItem14', f: parsenoop },
  0x0491: { n: 'BrtBeginCellIgnoreECs14', f: parsenoop },
  0x0492: { n: 'BrtEndCellIgnoreECs14', f: parsenoop },
  0x0493: { n: 'BrtDxf14', f: parsenoop },
  0x0494: { n: 'BrtBeginDxF14s', f: parsenoop },
  0x0495: { n: 'BrtEndDxf14s', f: parsenoop },
  0x0499: { n: 'BrtFilter14', f: parsenoop },
  0x049A: { n: 'BrtBeginCustomFilters14', f: parsenoop },
  0x049C: { n: 'BrtCustomFilter14', f: parsenoop },
  0x049D: { n: 'BrtIconFilter14', f: parsenoop },
  0x049E: { n: 'BrtPivotCacheConnectionName', f: parsenoop },
  0x0800: { n: 'BrtBeginDecoupledPivotCacheIDs', f: parsenoop },
  0x0801: { n: 'BrtEndDecoupledPivotCacheIDs', f: parsenoop },
  0x0802: { n: 'BrtDecoupledPivotCacheID', f: parsenoop },
  0x0803: { n: 'BrtBeginPivotTableRefs', f: parsenoop },
  0x0804: { n: 'BrtEndPivotTableRefs', f: parsenoop },
  0x0805: { n: 'BrtPivotTableRef', f: parsenoop },
  0x0806: { n: 'BrtSlicerCacheBookPivotTables', f: parsenoop },
  0x0807: { n: 'BrtBeginSxvcells', f: parsenoop },
  0x0808: { n: 'BrtEndSxvcells', f: parsenoop },
  0x0809: { n: 'BrtBeginSxRow', f: parsenoop },
  0x080A: { n: 'BrtEndSxRow', f: parsenoop },
  0x080C: { n: 'BrtPcdCalcMem15', f: parsenoop },
  0x0813: { n: 'BrtQsi15', f: parsenoop },
  0x0814: { n: 'BrtBeginWebExtensions', f: parsenoop },
  0x0815: { n: 'BrtEndWebExtensions', f: parsenoop },
  0x0816: { n: 'BrtWebExtension', f: parsenoop },
  0x0817: { n: 'BrtAbsPath15', f: parsenoop },
  0x0818: { n: 'BrtBeginPivotTableUISettings', f: parsenoop },
  0x0819: { n: 'BrtEndPivotTableUISettings', f: parsenoop },
  0x081B: { n: 'BrtTableSlicerCacheIDs', f: parsenoop },
  0x081C: { n: 'BrtTableSlicerCacheID', f: parsenoop },
  0x081D: { n: 'BrtBeginTableSlicerCache', f: parsenoop },
  0x081E: { n: 'BrtEndTableSlicerCache', f: parsenoop },
  0x081F: { n: 'BrtSxFilter15', f: parsenoop },
  0x0820: { n: 'BrtBeginTimelineCachePivotCacheIDs', f: parsenoop },
  0x0821: { n: 'BrtEndTimelineCachePivotCacheIDs', f: parsenoop },
  0x0822: { n: 'BrtTimelineCachePivotCacheID', f: parsenoop },
  0x0823: { n: 'BrtBeginTimelineCacheIDs', f: parsenoop },
  0x0824: { n: 'BrtEndTimelineCacheIDs', f: parsenoop },
  0x0825: { n: 'BrtBeginTimelineCacheID', f: parsenoop },
  0x0826: { n: 'BrtEndTimelineCacheID', f: parsenoop },
  0x0827: { n: 'BrtBeginTimelinesEx', f: parsenoop },
  0x0828: { n: 'BrtEndTimelinesEx', f: parsenoop },
  0x0829: { n: 'BrtBeginTimelineEx', f: parsenoop },
  0x082A: { n: 'BrtEndTimelineEx', f: parsenoop },
  0x082B: { n: 'BrtWorkBookPr15', f: parsenoop },
  0x082C: { n: 'BrtPCDH15', f: parsenoop },
  0x082D: { n: 'BrtBeginTimelineStyle', f: parsenoop },
  0x082E: { n: 'BrtEndTimelineStyle', f: parsenoop },
  0x082F: { n: 'BrtTimelineStyleElement', f: parsenoop },
  0x0830: { n: 'BrtBeginTimelineStylesheetExt15', f: parsenoop },
  0x0831: { n: 'BrtEndTimelineStylesheetExt15', f: parsenoop },
  0x0832: { n: 'BrtBeginTimelineStyles', f: parsenoop },
  0x0833: { n: 'BrtEndTimelineStyles', f: parsenoop },
  0x0834: { n: 'BrtBeginTimelineStyleElements', f: parsenoop },
  0x0835: { n: 'BrtEndTimelineStyleElements', f: parsenoop },
  0x0836: { n: 'BrtDxf15', f: parsenoop },
  0x0837: { n: 'BrtBeginDxfs15', f: parsenoop },
  0x0838: { n: 'brtEndDxfs15', f: parsenoop },
  0x0839: { n: 'BrtSlicerCacheHideItemsWithNoData', f: parsenoop },
  0x083A: { n: 'BrtBeginItemUniqueNames', f: parsenoop },
  0x083B: { n: 'BrtEndItemUniqueNames', f: parsenoop },
  0x083C: { n: 'BrtItemUniqueName', f: parsenoop },
  0x083D: { n: 'BrtBeginExtConn15', f: parsenoop },
  0x083E: { n: 'BrtEndExtConn15', f: parsenoop },
  0x083F: { n: 'BrtBeginOledbPr15', f: parsenoop },
  0x0840: { n: 'BrtEndOledbPr15', f: parsenoop },
  0x0841: { n: 'BrtBeginDataFeedPr15', f: parsenoop },
  0x0842: { n: 'BrtEndDataFeedPr15', f: parsenoop },
  0x0843: { n: 'BrtTextPr15', f: parsenoop },
  0x0844: { n: 'BrtRangePr15', f: parsenoop },
  0x0845: { n: 'BrtDbCommand15', f: parsenoop },
  0x0846: { n: 'BrtBeginDbTables15', f: parsenoop },
  0x0847: { n: 'BrtEndDbTables15', f: parsenoop },
  0x0848: { n: 'BrtDbTable15', f: parsenoop },
  0x0849: { n: 'BrtBeginDataModel', f: parsenoop },
  0x084A: { n: 'BrtEndDataModel', f: parsenoop },
  0x084B: { n: 'BrtBeginModelTables', f: parsenoop },
  0x084C: { n: 'BrtEndModelTables', f: parsenoop },
  0x084D: { n: 'BrtModelTable', f: parsenoop },
  0x084E: { n: 'BrtBeginModelRelationships', f: parsenoop },
  0x084F: { n: 'BrtEndModelRelationships', f: parsenoop },
  0x0850: { n: 'BrtModelRelationship', f: parsenoop },
  0x0851: { n: 'BrtBeginECTxtWiz15', f: parsenoop },
  0x0852: { n: 'BrtEndECTxtWiz15', f: parsenoop },
  0x0853: { n: 'BrtBeginECTWFldInfoLst15', f: parsenoop },
  0x0854: { n: 'BrtEndECTWFldInfoLst15', f: parsenoop },
  0x0855: { n: 'BrtBeginECTWFldInfo15', f: parsenoop },
  0x0856: { n: 'BrtFieldListActiveItem', f: parsenoop },
  0x0857: { n: 'BrtPivotCacheIdVersion', f: parsenoop },
  0x0858: { n: 'BrtSXDI15', f: parsenoop },
  0xFFFF: { n: '', f: parsenoop },
};

var evert_RE = evert_key(XLSBRecordEnum, 'n');

function fix_opts_func(defaults) {
  return function fix_opts(opts) {
    for (var i = 0; i !== defaults.length; ++i) {
      var d = defaults[i];
      if (opts[d[0]] === undefined) opts[d[0]] = d[1];
      if (d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
    }
  };
}

var fix_write_opts = fix_opts_func([
  ['bookSST', false], /* Generate Shared String Table */
  ['bookType', 'xlsx'], /* Type of workbook (xlsx/m/b) */
  ['WTF', false], /* WTF mode (throws errors) */
]);

function add_rels(rels, rId, f, type, relobj) {
  if (!relobj) relobj = {};
  if (!rels['!id']) rels['!id'] = {};
  relobj.Id = 'rId' + rId;
  relobj.Type = type;
  relobj.Target = f;
  if (rels['!id'][relobj.Id]) throw new Error('Cannot rewrite rId ' + rId);
  rels['!id'][relobj.Id] = relobj;
  rels[('/' + relobj.Target).replace('//', '/')] = relobj;
}

function write_zip(wb, opts) {
  if (wb && !wb.SSF) {
    wb.SSF = SSF.get_table();
  }
  if (wb && wb.SSF) {
    SSF.load_table(wb.SSF);
    opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
  }
  opts.rels = {}; opts.wbrels = {};
  opts.Strings = []; opts.Strings.Count = 0; opts.Strings.Unique = 0;
  var wbext = opts.bookType === 'xlsb' ? 'bin' : 'xml';
  var ct = {
    workbooks: [],
    sheets: [],
    calcchains: [],
    themes: [],
    styles: [],
    coreprops: [],
    extprops: [],
    custprops: [],
    strs: [],
    comments: [],
    vba: [],
    TODO: [],
    rels: [],
    xmlns: '',
  };
  fix_write_opts(opts = opts || {});
  var zip = new jszip();
  var f = '';
  var rId = 0;

  opts.cellXfs = [];
  get_cell_style(opts.cellXfs, {}, { revssf: { General: 0 } });

  f = 'docProps/core.xml';
  zip.file(f, write_core_props(wb.Props, opts));
  ct.coreprops.push(f);
  add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

  f = 'docProps/app.xml';
  if (!wb.Props) wb.Props = {};
  wb.Props.SheetNames = wb.SheetNames;
  wb.Props.Worksheets = wb.SheetNames.length;
  zip.file(f, write_ext_props(wb.Props, opts));
  ct.extprops.push(f);
  add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

  if (wb.Custprops !== wb.Props && keys(wb.Custprops || {}).length > 0) {
    f = 'docProps/custom.xml';
    zip.file(f, write_cust_props(wb.Custprops, opts));
    ct.custprops.push(f);
    add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
  }

  f = 'xl/workbook.' + wbext;
  zip.file(f, write_wb(wb, f, opts));
  ct.workbooks.push(f);
  add_rels(opts.rels, 1, f, RELS.WB);

  for (rId = 1; rId <= wb.SheetNames.length; ++rId) {
    f = 'xl/worksheets/sheet' + rId + '.' + wbext;
    zip.file(f, write_ws(rId - 1, f, opts, wb));
    ct.sheets.push(f);
    add_rels(opts.wbrels, rId, 'worksheets/sheet' + rId + '.' + wbext, RELS.WS);
  }

  if (opts.Strings != null && opts.Strings.length > 0) {
    f = 'xl/sharedStrings.' + wbext;
    zip.file(f, write_sst(opts.Strings, f, opts));
    ct.strs.push(f);
    add_rels(opts.wbrels, ++rId, 'sharedStrings.' + wbext, RELS.SST);
  }

  f = 'xl/theme/theme1.xml';
  zip.file(f, write_theme(opts));
  ct.themes.push(f);
  add_rels(opts.wbrels, ++rId, 'theme/theme1.xml', RELS.THEME);

  f = 'xl/styles.' + wbext;
  zip.file(f, write_sty(wb, f, opts));
  ct.styles.push(f);
  add_rels(opts.wbrels, ++rId, 'styles.' + wbext, RELS.STY);

  zip.file('[Content_Types].xml', write_ct(ct, opts));
  zip.file('_rels/.rels', write_rels(opts.rels));
  zip.file('xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));
  return zip;
}

function write_zip_type(wb, opts) {
  var o = opts || {};
  style_builder = new StyleBuilder(opts);

  var z = write_zip(wb, o);
  switch (o.type) {
    case 'base64': return z.generate({ type: 'base64' });
    case 'binary': return z.generate({ type: 'string' });
    case 'buffer': return z.generate({ type: 'nodebuffer' });
    default: throw new Error('Unrecognized type ' + o.type);
  }
}

function writeSync(wb, opts) {
  var o = opts || {};
  return write_zip_type(wb, o);
}

module.exports = writeSync;
