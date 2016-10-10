module.exports = function encodeRange(cs, ce) {
  if (ce == null || typeof ce === 'number') return encodeRange(cs.s, cs.e);
  if (typeof cs !== 'string') cs = encodeCell(cs); if (typeof ce !== 'string') ce = encodeCell(ce);
  return cs === ce ? cs : cs + ':' + ce;
};
