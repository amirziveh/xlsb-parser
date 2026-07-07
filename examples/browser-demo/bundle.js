// node_modules/fflate/esm/browser.js
var u8 = Uint8Array;
var u16 = Uint16Array;
var i32 = Int32Array;
var fleb = new u8([
  0,
  0,
  0,
  0,
  0,
  0,
  0,
  0,
  1,
  1,
  1,
  1,
  2,
  2,
  2,
  2,
  3,
  3,
  3,
  3,
  4,
  4,
  4,
  4,
  5,
  5,
  5,
  5,
  0,
  /* unused */
  0,
  0,
  /* impossible */
  0
]);
var fdeb = new u8([
  0,
  0,
  0,
  0,
  1,
  1,
  2,
  2,
  3,
  3,
  4,
  4,
  5,
  5,
  6,
  6,
  7,
  7,
  8,
  8,
  9,
  9,
  10,
  10,
  11,
  11,
  12,
  12,
  13,
  13,
  /* unused */
  0,
  0
]);
var clim = new u8([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]);
var freb = function(eb, start) {
  var b = new u16(31);
  for (var i = 0; i < 31; ++i) {
    b[i] = start += 1 << eb[i - 1];
  }
  var r = new i32(b[30]);
  for (var i = 1; i < 30; ++i) {
    for (var j = b[i]; j < b[i + 1]; ++j) {
      r[j] = j - b[i] << 5 | i;
    }
  }
  return { b, r };
};
var _a = freb(fleb, 2);
var fl = _a.b;
var revfl = _a.r;
fl[28] = 258, revfl[258] = 28;
var _b = freb(fdeb, 0);
var fd = _b.b;
var revfd = _b.r;
var rev = new u16(32768);
for (i = 0; i < 32768; ++i) {
  x = (i & 43690) >> 1 | (i & 21845) << 1;
  x = (x & 52428) >> 2 | (x & 13107) << 2;
  x = (x & 61680) >> 4 | (x & 3855) << 4;
  rev[i] = ((x & 65280) >> 8 | (x & 255) << 8) >> 1;
}
var x;
var i;
var hMap = (function(cd, mb, r) {
  var s = cd.length;
  var i = 0;
  var l = new u16(mb);
  for (; i < s; ++i) {
    if (cd[i])
      ++l[cd[i] - 1];
  }
  var le = new u16(mb);
  for (i = 1; i < mb; ++i) {
    le[i] = le[i - 1] + l[i - 1] << 1;
  }
  var co;
  if (r) {
    co = new u16(1 << mb);
    var rvb = 15 - mb;
    for (i = 0; i < s; ++i) {
      if (cd[i]) {
        var sv = i << 4 | cd[i];
        var r_1 = mb - cd[i];
        var v = le[cd[i] - 1]++ << r_1;
        for (var m = v | (1 << r_1) - 1; v <= m; ++v) {
          co[rev[v] >> rvb] = sv;
        }
      }
    }
  } else {
    co = new u16(s);
    for (i = 0; i < s; ++i) {
      if (cd[i]) {
        co[i] = rev[le[cd[i] - 1]++] >> 15 - cd[i];
      }
    }
  }
  return co;
});
var flt = new u8(288);
for (i = 0; i < 144; ++i)
  flt[i] = 8;
var i;
for (i = 144; i < 256; ++i)
  flt[i] = 9;
var i;
for (i = 256; i < 280; ++i)
  flt[i] = 7;
var i;
for (i = 280; i < 288; ++i)
  flt[i] = 8;
var i;
var fdt = new u8(32);
for (i = 0; i < 32; ++i)
  fdt[i] = 5;
var i;
var flrm = /* @__PURE__ */ hMap(flt, 9, 1);
var fdrm = /* @__PURE__ */ hMap(fdt, 5, 1);
var max = function(a) {
  var m = a[0];
  for (var i = 1; i < a.length; ++i) {
    if (a[i] > m)
      m = a[i];
  }
  return m;
};
var bits = function(d, p, m) {
  var o = p / 8 | 0;
  return (d[o] | d[o + 1] << 8) >> (p & 7) & m;
};
var bits16 = function(d, p) {
  var o = p / 8 | 0;
  return (d[o] | d[o + 1] << 8 | d[o + 2] << 16) >> (p & 7);
};
var shft = function(p) {
  return (p + 7) / 8 | 0;
};
var slc = function(v, s, e) {
  if (s == null || s < 0)
    s = 0;
  if (e == null || e > v.length)
    e = v.length;
  return new u8(v.subarray(s, e));
};
var ec = [
  "unexpected EOF",
  "invalid block type",
  "invalid length/literal",
  "invalid distance",
  "stream finished",
  "no stream handler",
  ,
  // determined by compression function
  "no callback",
  "invalid UTF-8 data",
  "extra field too long",
  "date not in range 1980-2099",
  "filename too long",
  "stream finishing",
  "invalid zip data"
  // determined by unknown compression method
];
var err = function(ind, msg, nt) {
  var e = new Error(msg || ec[ind]);
  e.code = ind;
  if (Error.captureStackTrace)
    Error.captureStackTrace(e, err);
  if (!nt)
    throw e;
  return e;
};
var inflt = function(dat, st, buf, dict) {
  var sl = dat.length, dl = dict ? dict.length : 0;
  if (!sl || st.f && !st.l)
    return buf || new u8(0);
  var noBuf = !buf;
  var resize = noBuf || st.i != 2;
  var noSt = st.i;
  if (noBuf)
    buf = new u8(sl * 3);
  var cbuf = function(l2) {
    var bl = buf.length;
    if (l2 > bl) {
      var nbuf = new u8(Math.max(bl * 2, l2));
      nbuf.set(buf);
      buf = nbuf;
    }
  };
  var final = st.f || 0, pos = st.p || 0, bt = st.b || 0, lm = st.l, dm = st.d, lbt = st.m, dbt = st.n;
  var tbts = sl * 8;
  do {
    if (!lm) {
      final = bits(dat, pos, 1);
      var type = bits(dat, pos + 1, 3);
      pos += 3;
      if (!type) {
        var s = shft(pos) + 4, l = dat[s - 4] | dat[s - 3] << 8, t = s + l;
        if (t > sl) {
          if (noSt)
            err(0);
          break;
        }
        if (resize)
          cbuf(bt + l);
        buf.set(dat.subarray(s, t), bt);
        st.b = bt += l, st.p = pos = t * 8, st.f = final;
        continue;
      } else if (type == 1)
        lm = flrm, dm = fdrm, lbt = 9, dbt = 5;
      else if (type == 2) {
        var hLit = bits(dat, pos, 31) + 257, hcLen = bits(dat, pos + 10, 15) + 4;
        var tl = hLit + bits(dat, pos + 5, 31) + 1;
        pos += 14;
        var ldt = new u8(tl);
        var clt = new u8(19);
        for (var i = 0; i < hcLen; ++i) {
          clt[clim[i]] = bits(dat, pos + i * 3, 7);
        }
        pos += hcLen * 3;
        var clb = max(clt), clbmsk = (1 << clb) - 1;
        var clm = hMap(clt, clb, 1);
        for (var i = 0; i < tl; ) {
          var r = clm[bits(dat, pos, clbmsk)];
          pos += r & 15;
          var s = r >> 4;
          if (s < 16) {
            ldt[i++] = s;
          } else {
            var c = 0, n = 0;
            if (s == 16)
              n = 3 + bits(dat, pos, 3), pos += 2, c = ldt[i - 1];
            else if (s == 17)
              n = 3 + bits(dat, pos, 7), pos += 3;
            else if (s == 18)
              n = 11 + bits(dat, pos, 127), pos += 7;
            while (n--)
              ldt[i++] = c;
          }
        }
        var lt = ldt.subarray(0, hLit), dt = ldt.subarray(hLit);
        lbt = max(lt);
        dbt = max(dt);
        lm = hMap(lt, lbt, 1);
        dm = hMap(dt, dbt, 1);
      } else
        err(1);
      if (pos > tbts) {
        if (noSt)
          err(0);
        break;
      }
    }
    if (resize)
      cbuf(bt + 131072);
    var lms = (1 << lbt) - 1, dms = (1 << dbt) - 1;
    var lpos = pos;
    for (; ; lpos = pos) {
      var c = lm[bits16(dat, pos) & lms], sym = c >> 4;
      pos += c & 15;
      if (pos > tbts) {
        if (noSt)
          err(0);
        break;
      }
      if (!c)
        err(2);
      if (sym < 256)
        buf[bt++] = sym;
      else if (sym == 256) {
        lpos = pos, lm = null;
        break;
      } else {
        var add = sym - 254;
        if (sym > 264) {
          var i = sym - 257, b = fleb[i];
          add = bits(dat, pos, (1 << b) - 1) + fl[i];
          pos += b;
        }
        var d = dm[bits16(dat, pos) & dms], dsym = d >> 4;
        if (!d)
          err(3);
        pos += d & 15;
        var dt = fd[dsym];
        if (dsym > 3) {
          var b = fdeb[dsym];
          dt += bits16(dat, pos) & (1 << b) - 1, pos += b;
        }
        if (pos > tbts) {
          if (noSt)
            err(0);
          break;
        }
        if (resize)
          cbuf(bt + 131072);
        var end = bt + add;
        if (bt < dt) {
          var shift = dl - dt, dend = Math.min(dt, end);
          if (shift + bt < 0)
            err(3);
          for (; bt < dend; ++bt)
            buf[bt] = dict[shift + bt];
        }
        for (; bt < end; ++bt)
          buf[bt] = buf[bt - dt];
      }
    }
    st.l = lm, st.p = lpos, st.b = bt, st.f = final;
    if (lm)
      final = 1, st.m = lbt, st.d = dm, st.n = dbt;
  } while (!final);
  return bt != buf.length && noBuf ? slc(buf, 0, bt) : buf.subarray(0, bt);
};
var et = /* @__PURE__ */ new u8(0);
var b2 = function(d, b) {
  return d[b] | d[b + 1] << 8;
};
var b4 = function(d, b) {
  return (d[b] | d[b + 1] << 8 | d[b + 2] << 16 | d[b + 3] << 24) >>> 0;
};
var b8 = function(d, b) {
  return b4(d, b) + b4(d, b + 4) * 4294967296;
};
function inflateSync(data, opts) {
  return inflt(data, { i: 2 }, opts && opts.out, opts && opts.dictionary);
}
var td = typeof TextDecoder != "undefined" && /* @__PURE__ */ new TextDecoder();
var tds = 0;
try {
  td.decode(et, { stream: true });
  tds = 1;
} catch (e) {
}
var dutf8 = function(d) {
  for (var r = "", i = 0; ; ) {
    var c = d[i++];
    var eb = (c > 127) + (c > 223) + (c > 239);
    if (i + eb > d.length)
      return { s: r, r: slc(d, i - 1) };
    if (!eb)
      r += String.fromCharCode(c);
    else if (eb == 3) {
      c = ((c & 15) << 18 | (d[i++] & 63) << 12 | (d[i++] & 63) << 6 | d[i++] & 63) - 65536, r += String.fromCharCode(55296 | c >> 10, 56320 | c & 1023);
    } else if (eb & 1)
      r += String.fromCharCode((c & 31) << 6 | d[i++] & 63);
    else
      r += String.fromCharCode((c & 15) << 12 | (d[i++] & 63) << 6 | d[i++] & 63);
  }
};
function strFromU8(dat, latin1) {
  if (latin1) {
    var r = "";
    for (var i = 0; i < dat.length; i += 16384)
      r += String.fromCharCode.apply(null, dat.subarray(i, i + 16384));
    return r;
  } else if (td) {
    return td.decode(dat);
  } else {
    var _a2 = dutf8(dat), s = _a2.s, r = _a2.r;
    if (r.length)
      err(8);
    return s;
  }
}
var slzh = function(d, b) {
  return b + 30 + b2(d, b + 26) + b2(d, b + 28);
};
var zh = function(d, b, z) {
  var fnl = b2(d, b + 28), efl = b2(d, b + 30), fn = strFromU8(d.subarray(b + 46, b + 46 + fnl), !(b2(d, b + 8) & 2048)), es = b + 46 + fnl;
  var _a2 = z64hs(d, es, efl, z, b4(d, b + 20), b4(d, b + 24), b4(d, b + 42)), sc = _a2[0], su = _a2[1], off = _a2[2];
  return [b2(d, b + 10), sc, su, fn, es + efl + b2(d, b + 32), off];
};
var z64hs = function(d, b, l, z, sc, su, off) {
  var nsc = sc == 4294967295, nsu = su == 4294967295, noff = off == 4294967295, e = b + l;
  var nf = nsc + nsu + noff;
  if (z && nf) {
    for (; b + 4 < e; b += 4 + b2(d, b + 2)) {
      if (b2(d, b) == 1) {
        return [
          nsc ? b8(d, b + 4 + 8 * nsu) : sc,
          nsu ? b8(d, b + 4) : su,
          noff ? b8(d, b + 4 + 8 * (nsu + nsc)) : off,
          1
        ];
      }
    }
    if (z < 2)
      err(13);
  }
  return [sc, su, off, 0];
};
function unzipSync(data, opts) {
  var files = {};
  var e = data.length - 22;
  for (; b4(data, e) != 101010256; --e) {
    if (!e || data.length - e > 65558)
      err(13);
  }
  ;
  var c = b2(data, e + 8);
  if (!c)
    return {};
  var o = b4(data, e + 16);
  var z = b4(data, e - 20) == 117853008;
  if (z) {
    var ze = b4(data, e - 12);
    z = b4(data, ze) == 101075792;
    if (z) {
      c = b4(data, ze + 32);
      o = b4(data, ze + 48);
    }
  }
  var fltr = opts && opts.filter;
  for (var i = 0; i < c; ++i) {
    var _a2 = zh(data, o, z), c_2 = _a2[0], sc = _a2[1], su = _a2[2], fn = _a2[3], no = _a2[4], off = _a2[5], b = slzh(data, off);
    o = no;
    if (!fltr || fltr({
      name: fn,
      size: sc,
      originalSize: su,
      compression: c_2
    })) {
      if (!c_2)
        files[fn] = slc(data, b, b + sc);
      else if (c_2 == 8)
        files[fn] = inflateSync(data.subarray(b, b + sc), { out: new u8(su) });
      else
        err(14, "unknown compression type " + c_2);
    }
  }
  return files;
}

// src/types.ts
var XlsbSizeError = class extends Error {
  name = "XlsbSizeError";
  limit;
  actual;
  constructor(message, limit, actual) {
    super(message);
    this.limit = limit;
    this.actual = actual;
  }
};

// src/record-stream.ts
var dec16 = new TextDecoder("utf-16le");
var dec8 = new TextDecoder("utf-8");
var BRT_ROW_HEADER = 0;
var BRT_CELL_BLANK = 1;
var BRT_CELL_RK = 2;
var BRT_CELL_ERROR = 3;
var BRT_CELL_BOOL = 4;
var BRT_CELL_REAL = 5;
var BRT_CELL_ST = 6;
var BRT_CELL_ISST = 7;
var BRT_FMLA_STRING = 8;
var BRT_FMLA_NUM = 9;
var BRT_FMLA_BOOL = 10;
var BRT_FMLA_ERROR = 11;
var BRT_SHORT_BLANK = 12;
var BRT_SHORT_RK = 13;
var BRT_SHORT_ERROR = 14;
var BRT_SHORT_BOOL = 15;
var BRT_SHORT_REAL = 16;
var BRT_SHORT_ST = 17;
var BRT_SHORT_ISST = 18;
var BRT_SST_ITEM = 19;
var BRT_BUNDLE_SH = 156;
var BRT_BUNDLE_SH_NEW = 3585;
var ERRORS = {
  0: "#NULL!",
  7: "#DIV/0!",
  15: "#VALUE!",
  23: "#REF!",
  29: "#NAME?",
  36: "#NUM!",
  42: "#N/A",
  43: "#GETTING_DATA"
};
function* records(data) {
  let off = 0;
  while (off < data.length) {
    const recStart = off;
    if (off >= data.length) break;
    let t = data[off++];
    if ((t & 128) !== 0) {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record type byte at offset ${recStart} announces a second byte but only ${data.length} bytes total remain`
        );
      }
      t = (t & 127) << 7 | data[off++];
    }
    let s = 0, sh = 0, b;
    do {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size varint overruns the buffer`
        );
      }
      b = data[off++];
      s |= (b & 127) << sh;
      sh += 7;
    } while (b & 128);
    if (off + s > data.length) {
      throw new Error(
        `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size ${s} but only ${data.length - off} bytes remain`
      );
    }
    yield { type: t, data: data.subarray(off, off + s) };
    off += s;
  }
}
function readU16(d, off) {
  return d[off] | d[off + 1] << 8;
}
function readU32(d, off) {
  return (d[off] | d[off + 1] << 8 | d[off + 2] << 16 | d[off + 3] << 24) >>> 0;
}
function readF64(d, off) {
  return new DataView(d.buffer, d.byteOffset + off, 8).getFloat64(0, true);
}
function readWideString(d, off) {
  const len = readU32(d, off);
  return dec16.decode(d.subarray(off + 4, off + 4 + len * 2));
}
function readRichString(d, off) {
  return readWideString(d, off + 1);
}
var RK_SCRATCH_BUF = new ArrayBuffer(8);
var RK_SCRATCH_DV = new DataView(RK_SCRATCH_BUF);
var RK_SCRATCH_U32 = new Uint32Array(RK_SCRATCH_BUF);
function decodeRk(rk) {
  const fx100 = rk & 1;
  const fInt = rk >> 1 & 1;
  const num = rk >>> 2;
  let val;
  if (fInt) {
    val = num << 2 >> 2;
  } else {
    RK_SCRATCH_U32[0] = 0;
    RK_SCRATCH_U32[1] = num << 2;
    val = RK_SCRATCH_DV.getFloat64(0, true);
  }
  if (fx100) val /= 100;
  return val;
}
function hex(d, max2 = 48) {
  return Array.from(d.subarray(0, Math.min(max2, d.length))).map((b) => b.toString(16).padStart(2, "0")).join(" ");
}

// src/workbook.ts
function parseWorkbook(data) {
  const names = [];
  for (const r of records(data)) {
    if (r.type === BRT_BUNDLE_SH_NEW && r.data.length >= 12) {
      const rIdLen = readU32(r.data, 8);
      const nameOff = 12 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    } else if (r.type === BRT_BUNDLE_SH && r.data.length >= 9) {
      const rIdLen = readU32(r.data, 5);
      const nameOff = 9 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    }
  }
  return names;
}

// src/shared-strings.ts
function parseSharedStrings(data) {
  const list = [];
  for (const r of records(data)) {
    if (r.type === BRT_SST_ITEM && r.data.length >= 5) list.push(readRichString(r.data, 0));
  }
  return list;
}

// src/styles.ts
var BRT_FMT = 729;
var BRT_CELL_XF = 505;
var BUILTIN_DATE_FMT_IDS = /* @__PURE__ */ new Set([
  14,
  15,
  16,
  17,
  18,
  19,
  20,
  21,
  22,
  27,
  28,
  29,
  30,
  31,
  32,
  33,
  34,
  35,
  36,
  45,
  46,
  47,
  50,
  51,
  52,
  53,
  54,
  55,
  56,
  57,
  58,
  78,
  79,
  80,
  81
]);
var SERIAL_EPOCH_MS = Date.UTC(1899, 11, 30);
function parseStyles(data) {
  const cellXfs = [];
  const numFmts = /* @__PURE__ */ new Map();
  for (const r of records(data)) {
    if (r.type === BRT_FMT && r.data.length >= 4) {
      const numFmtId = readU16(r.data, 0);
      try {
        const len = readU32At(r.data, 2);
        if (len > 0 && len < 1e3 && 6 + len * 2 <= r.data.length) {
          const s = new TextDecoder("utf-16le").decode(r.data.subarray(6, 6 + len * 2));
          numFmts.set(numFmtId, s);
        }
      } catch {
      }
    } else if (r.type === BRT_CELL_XF && r.data.length >= 2) {
      cellXfs.push(readU16(r.data, 0));
    }
  }
  return { cellXfs, numFmts };
}
function readU32At(d, off) {
  return readU32(d, off);
}
function isDateFormatId(numFmtId) {
  return BUILTIN_DATE_FMT_IDS.has(numFmtId);
}
function isDateNumFmtString(fmt) {
  let escaped = false;
  let quote = false;
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (c === "\\") {
      escaped = true;
      continue;
    }
    if (c === '"') {
      quote = !quote;
      continue;
    }
    if (quote) continue;
    if (/[dmyYhHsS]/.test(c)) return true;
    if ((c === "A" || c === "a") && fmt.slice(i, i + 2).match(/am|AM/)) return true;
    if ((c === "P" || c === "p") && fmt.slice(i, i + 2).match(/pm|PM/)) return true;
  }
  return false;
}
function numFmtIdToSerialConvert(serial) {
  const ms = SERIAL_EPOCH_MS + serial * 864e5;
  return new Date(ms).toISOString();
}

// src/sheet.ts
function parseSheet(data, ss, opts = {}) {
  const rows = [];
  const styles = opts.styles ?? null;
  const maxRows = opts.maxRows;
  let curRow = null;
  let prevCol = -1;
  for (const r of records(data)) {
    const d = r.data;
    if (r.type === BRT_ROW_HEADER && d.length >= 4) {
      if (maxRows !== void 0 && rows.length >= maxRows) {
        throw new XlsbSizeError(`Sheet exceeded maxRows=${maxRows} limit`, maxRows, rows.length);
      }
      curRow = { row: readU32(d, 0), cols: {} };
      rows.push(curRow);
      prevCol = -1;
      continue;
    }
    if (!curRow) continue;
    if (r.type >= BRT_CELL_BLANK && r.type <= BRT_FMLA_ERROR) {
      const col = readU32(d, 0);
      const ixf = d.length >= 6 ? d[4] | d[5] << 8 | (d[5] & 128 ? 4294901760 : 0) : void 0;
      const cell = readCell(r.type, d, 8, ss);
      if (cell) {
        cell.ixf = ixf;
        applyDateMeta(cell, ixf, styles);
        curRow.cols[col] = cell;
      }
      prevCol = col;
    } else if (r.type >= BRT_SHORT_BLANK && r.type <= BRT_SHORT_ISST) {
      const col = prevCol + 1;
      const ixf = d.length >= 4 ? readU16(d, 2) : void 0;
      const cell = readShortCell(r.type, d, 4, ss);
      if (cell) {
        cell.ixf = ixf;
        applyDateMeta(cell, ixf, styles);
        curRow.cols[col] = cell;
      }
      prevCol = col;
    }
  }
  return rows;
}
function applyDateMeta(cell, ixf, styles) {
  if (!styles || ixf === void 0 || ixf < 0) return;
  const idx = ixf & 65535;
  if (idx >= styles.cellXfs.length) return;
  const numFmtId = styles.cellXfs[idx];
  cell.numFmtId = numFmtId;
  const customFmt = styles.numFmts.get(numFmtId);
  const isDate = customFmt !== void 0 && isDateNumFmtString(customFmt) || isDateFormatId(numFmtId);
  if (!isDate) return;
  cell.isDate = true;
  if (cell.t === "n" && typeof cell.v === "number") {
    cell.dateValue = numFmtIdToSerialConvert(cell.v);
  }
}
function readCell(type, d, off, ss) {
  switch (type) {
    case BRT_CELL_BLANK:
      return { t: "blank" };
    case BRT_CELL_RK:
      if (off + 4 > d.length) return null;
      return { t: "n", v: decodeRk(readU32(d, off)) };
    case BRT_CELL_REAL:
      if (off + 8 > d.length) return null;
      return { t: "n", v: readF64(d, off) };
    case BRT_CELL_ISST:
      if (off + 4 > d.length) return null;
      return { t: "s", v: ss[readU32(d, off)] ?? `[SST#${readU32(d, off)}]` };
    case BRT_CELL_BOOL:
      if (off + 1 > d.length) return null;
      return { t: "b", v: d[off] !== 0 };
    case BRT_CELL_ERROR:
      if (off + 1 > d.length) return null;
      return { t: "e", err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    case BRT_CELL_ST:
      if (off + 5 > d.length) return null;
      return { t: "s", v: readRichString(d, off) };
    case BRT_FMLA_NUM:
      if (off + 8 > d.length) return null;
      return { t: "n", v: readF64(d, off) };
    case BRT_FMLA_STRING:
      if (off + 4 > d.length) return null;
      return { t: "s", v: readWideString(d, off) };
    case BRT_FMLA_BOOL:
      if (off + 1 > d.length) return null;
      return { t: "b", v: d[off] !== 0 };
    case BRT_FMLA_ERROR:
      if (off + 1 > d.length) return null;
      return { t: "e", err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    default:
      return null;
  }
}
function readShortCell(type, d, off, ss) {
  switch (type) {
    case BRT_SHORT_BLANK:
      return { t: "blank" };
    case BRT_SHORT_RK:
      if (off + 4 > d.length) return null;
      return { t: "n", v: decodeRk(readU32(d, off)) };
    case BRT_SHORT_ERROR:
      if (off + 1 > d.length) return null;
      return { t: "e", err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    case BRT_SHORT_BOOL:
      if (off + 1 > d.length) return null;
      return { t: "b", v: d[off] !== 0 };
    case BRT_SHORT_REAL:
      if (off + 8 > d.length) return null;
      return { t: "n", v: readF64(d, off) };
    case BRT_SHORT_ST:
      if (off + 5 > d.length) return null;
      return { t: "s", v: readRichString(d, off) };
    case BRT_SHORT_ISST:
      if (off + 4 > d.length) return null;
      return { t: "s", v: ss[readU32(d, off)] ?? `[SST#${readU32(d, off)}]` };
    default:
      return null;
  }
}

// src/dump.ts
function dumpBinary(path, data, maxRec = 200) {
  const recordsArr = [];
  const typeSummary = {};
  let total = 0;
  for (const r of records(data)) {
    const key = "0x" + r.type.toString(16).toUpperCase().padStart(4, "0");
    typeSummary[key] = (typeSummary[key] || 0) + 1;
    total++;
    if (recordsArr.length >= maxRec) continue;
    const strings = [];
    for (let off = 0; off + 4 < r.data.length; ) {
      const len = readU32(r.data, off);
      if (len > 0 && len < 200 && off + 4 + len * 2 <= r.data.length) {
        try {
          const s = dec16.decode(r.data.subarray(off + 4, off + 4 + len * 2)).replace(/\0/g, "");
          if (s.length >= 2) strings.push(s);
        } catch {
        }
        off += 4 + len * 2;
      } else {
        off++;
      }
    }
    recordsArr.push({
      type: key,
      typeNum: r.type,
      size: r.data.length,
      hex: hex(r.data, 32),
      strings
    });
  }
  if (total > maxRec) {
    recordsArr.push({
      type: "...",
      typeNum: -1,
      size: 0,
      hex: `[${total - maxRec} more records omitted]`,
      strings: []
    });
  }
  return { path, size: data.length, recCount: total, records: recordsArr, typeSummary };
}

// src/pivot-cache.ts
function formatYMD(y, m, d) {
  return String(y).padStart(4, "0") + "-" + String(m).padStart(2, "0") + "-" + String(d).padStart(2, "0");
}
function parsePivotCache(name, def, recs) {
  const fieldNames = [];
  const sharedItems = [];
  let curItems = [];
  let fallbackItems = [];
  let hierItems = [];
  let has1F81 = false;
  function pushField() {
    if (fieldNames.length === 0) return;
    const src = has1F81 ? curItems : curItems.length > 0 ? curItems : fallbackItems.length > 0 ? fallbackItems : hierItems.length > 0 ? hierItems : curItems;
    sharedItems.push(src);
  }
  for (const r of records(def)) {
    if (r.type === 7041) {
      pushField();
      curItems = [];
      fallbackItems = [];
      hierItems = [];
      has1F81 = false;
      const d = r.data;
      if (d.length >= 24) {
        const nameLen = readU32(d, 20);
        if (nameLen > 0 && nameLen < 100 && 24 + nameLen * 2 <= d.length) {
          fieldNames.push(dec16.decode(d.subarray(24, 24 + nameLen * 2)));
        }
      }
    } else if (r.type === 31) {
      const d = r.data;
      if (d.length >= 4) {
        const slen = readU32(d, 0);
        if (slen > 0 && slen < 500) {
          const endOff = 4 + slen * 2;
          if (endOff <= d.length) {
            fallbackItems.push(dec16.decode(d.subarray(4, endOff)));
          }
        }
      }
    } else if (r.type === 24 && r.data.length > 4) {
      const len = readU32(r.data, 0);
      if (len > 0 && len < 200 && 4 + len * 2 <= r.data.length) {
        hierItems.push(dec16.decode(r.data.subarray(4, 4 + len * 2)));
      }
    } else if (r.type === 8065 && r.data.length >= 6 && (r.data[0] === 32 || r.data[0] === 2)) {
      has1F81 = true;
      const d = r.data;
      if (d[0] === 32) {
        const count = readU32(d, 2);
        for (let i = 0, off = 6; i < count && off + 8 <= d.length; i++, off += 8) {
          curItems.push(
            formatYMD(
              d[off] | d[off + 1] << 8,
              d[off + 2] | d[off + 3] << 8,
              d[off + 4] | d[off + 5] << 8
            )
          );
        }
      } else if (d[0] === 2) {
        const count = readU32(d, 2);
        for (let i = 0, off = 6; i < count && off + 4 <= d.length; i++) {
          const slen = readU32(d, off);
          if (slen > 0 && slen < 500 && off + 4 + slen * 2 <= d.length) {
            curItems.push(dec16.decode(d.subarray(off + 4, off + 4 + slen * 2)));
            off += 4 + slen * 2;
          } else break;
        }
      }
    } else if (r.type === 32 && r.data.length >= 8) {
      const d = r.data;
      curItems.push(formatYMD(d[0] | d[1] << 8, d[2] | d[3] << 8, d[4] | d[5] << 8));
    }
  }
  pushField();
  const recBodies = [];
  for (const r of records(recs)) {
    if (r.type === 8449) break;
    if (r.type !== 33) continue;
    recBodies.push(r.data);
    if (recBodies.length >= 50) break;
  }
  const fieldCount = fieldNames.length;
  const fields = new Array(fieldCount).fill(null);
  const recOffsets = new Int32Array(recBodies.length);
  for (let fi = 0; fi < fieldCount; fi++) {
    for (let ri = 0; ri < recBodies.length; ri++) {
      const body = recBodies[ri];
      const off = recOffsets[ri];
      if (off < 0 || off + 4 > body.length) continue;
      const len = readU32(body, off);
      if (len >= 3 && len < 200 && off + 4 + len * 2 <= body.length) {
        try {
          const s = dec16.decode(body.subarray(off + 4, off + 4 + len * 2));
          let valid = true;
          let alpha = 0;
          for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 && c !== 10 && c !== 13) {
              valid = false;
              break;
            }
            if (c >= 65 && c <= 90 || c >= 97 && c <= 122) alpha++;
          }
          if (valid && alpha > 0) {
            fields[fi] = "str";
            break;
          }
        } catch {
        }
      }
    }
    if (fields[fi] === "str") {
      for (let ri = 0; ri < recBodies.length; ri++) {
        const body = recBodies[ri];
        const off = recOffsets[ri];
        if (off >= 0) {
          const len = readU32(body, off);
          recOffsets[ri] = off + 4 + len * 2;
        }
      }
      continue;
    }
    let isF64 = false;
    for (let ri = 0; ri < recBodies.length; ri++) {
      const body = recBodies[ri];
      const off = recOffsets[ri];
      if (off < 0 || off + 8 > body.length) continue;
      const f = readF64(body, off);
      const lo = readU32(body, off);
      const hi = readU32(body, off + 4);
      if (isFinite(f) && !isNaN(f) && (lo !== 0 || hi !== 0) && Math.abs(f) > 1e-10 && Math.abs(f) < 1e20 && !(hi === 0 && lo < 1e5)) {
        isF64 = true;
        break;
      }
    }
    fields[fi] = isF64 ? "f64" : "u32";
    const sz = isF64 ? 8 : 4;
    for (let ri = 0; ri < recBodies.length; ri++) {
      if (recOffsets[ri] >= 0) recOffsets[ri] += sz;
    }
  }
  for (let i = 1; i < fields.length - 1; i++) {
    if (fields[i] === "u32" && fields[i - 1] === "f64" && fields[i + 1] === "f64")
      fields[i] = "f64";
  }
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 1; i < fields.length - 1; i++) {
      if (fields[i] === "u32" && fields[i - 1] === "f64" && fields[i + 1] === "f64") {
        fields[i] = "f64";
        changed = true;
      }
    }
  }
  let lastStr = -1;
  for (let i = 0; i < fields.length; i++) {
    if (fields[i] === "str") lastStr = i;
  }
  for (const body of recBodies) {
    let consumed = 0;
    for (let i = 0; i < fields.length; i++) {
      if (fields[i] === "str") {
        if (consumed + 4 > body.length) {
          consumed = -1;
          break;
        }
        const len = readU32(body, consumed);
        consumed += len >= 3 && consumed + 4 + len * 2 <= body.length ? 4 + len * 2 : 4;
      } else {
        consumed += fields[i] === "f64" ? 8 : 4;
      }
    }
    if (consumed < 0) continue;
    const deficit = body.length - consumed;
    if (deficit > 0 && deficit % 4 === 0) {
      const mis = deficit / 4;
      let converted = 0;
      for (let i = fields.length - 1; i > lastStr && converted < mis; i--) {
        if (fields[i] === "u32") {
          fields[i] = "f64";
          converted++;
        }
      }
    }
  }
  if (recBodies.length > 0) {
    let best = fields.length;
    for (let check = fields.length; check >= 1; check--) {
      let allFit = true;
      for (const body of recBodies) {
        let off = 0;
        for (let i = 0; i < check && off <= body.length; i++) {
          if (fields[i] === "str") {
            if (off + 4 > body.length) {
              allFit = false;
              break;
            }
            const len = readU32(body, off);
            off += len >= 3 && off + 4 + len * 2 <= body.length ? 4 + len * 2 : 4;
          } else {
            off += fields[i] === "f64" ? 8 : 4;
          }
        }
        if (off > body.length) {
          allFit = false;
          break;
        }
      }
      if (allFit) {
        best = check;
        break;
      }
    }
    if (best < fields.length) fields.length = best;
  }
  const rows = [];
  for (const r of records(recs)) {
    if (r.type === 8449) break;
    if (r.type !== 33) continue;
    const d = r.data;
    const values = [];
    let off = 0;
    let fi = 0;
    while (off < d.length && fi < fields.length) {
      const ft = fields[fi];
      if (ft === "str") {
        const len = readU32(d, off);
        if (len >= 3 && off + 4 + len * 2 <= d.length) {
          values.push(dec16.decode(d.subarray(off + 4, off + 4 + len * 2)));
          off += 4 + len * 2;
        } else {
          values.push(null);
          off += 4;
        }
      } else if (ft === "f64") {
        if (off + 8 <= d.length) {
          const f = readF64(d, off);
          values.push(
            f === 0 ? 0 : Math.abs(f) >= 1 ? parseFloat(f.toFixed(4)) : parseFloat(f.toFixed(8))
          );
          off += 8;
        } else {
          values.push(null);
          off += 4;
        }
      } else {
        if (off + 4 <= d.length) {
          const v = readU32(d, off);
          const si = fi < sharedItems.length ? sharedItems[fi] : null;
          if (si && si.length > 0 && v < si.length && si[v] !== null) {
            values.push(si[v]);
          } else {
            values.push(v);
          }
        } else values.push(null);
        off += 4;
      }
      fi++;
    }
    rows.push(values);
  }
  if (rows.length > 0) {
    let lastPopulated = 0;
    for (let ri = 0; ri < Math.min(5, rows.length); ri++) {
      for (let fi = rows[ri].length - 1; fi >= 0; fi--) {
        if (rows[ri][fi] !== void 0 && fi > lastPopulated) lastPopulated = fi;
      }
    }
    if (lastPopulated + 1 < fieldNames.length) fieldNames.length = lastPopulated + 1;
  }
  return {
    name,
    fieldNames,
    rows,
    rowCount: recBodies.length > 50 ? recBodies.length : rows.length
  };
}

// src/handle.ts
function tick() {
  return new Promise((r) => setTimeout(r, 0));
}
function normalizeOptions(arg) {
  if (arg === void 0) return {};
  if (typeof arg === "function") return { onProgress: arg };
  return arg;
}
async function openXlsb(data, options) {
  const opts = normalizeOptions(options);
  const onProgress = opts.onProgress;
  const maxZipBytes = opts.maxZipBytes;
  onProgress?.("Decompressing ZIP...", 0);
  await new Promise((r) => setTimeout(r, 50));
  const u82 = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
  let zip;
  try {
    zip = unzipSync(u82);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error("ZIP decompression failed: " + msg, { cause: e });
  }
  if (maxZipBytes !== void 0) {
    let total = 0;
    for (const k of Object.keys(zip)) total += zip[k].length;
    if (total > maxZipBytes) {
      throw new XlsbSizeError(
        `Decompressed ZIP size ${total} bytes exceeds maxZipBytes limit ${maxZipBytes}`,
        maxZipBytes,
        total
      );
    }
  }
  const wb = zip["xl/workbook.bin"];
  if (!wb) throw new Error("xl/workbook.bin not found");
  onProgress?.("Parsing workbook...", 5);
  const sheetNames = parseWorkbook(wb);
  await tick();
  let sharedStrings = [];
  if (zip["xl/sharedStrings.bin"]) {
    onProgress?.("Parsing shared strings...", 10);
    sharedStrings = parseSharedStrings(zip["xl/sharedStrings.bin"]);
    await tick();
  }
  let styles = null;
  if (zip["xl/styles.bin"]) {
    onProgress?.("Parsing styles...", 12);
    try {
      styles = parseStyles(zip["xl/styles.bin"]);
    } catch {
      styles = null;
    }
    await tick();
  }
  const sheetBytes = sheetNames.map((_, i) => {
    return zip[`xl/worksheets/sheet${i + 1}.bin`] ?? null;
  });
  const handle = {
    sheetNames,
    sharedStrings,
    styles,
    async *iterSheetRows(sheetIndex, iterOpts = {}) {
      const bytes = sheetBytes[sheetIndex];
      if (!bytes) return;
      const maxRows = iterOpts.maxRows;
      const onProgressIter = iterOpts.onProgress;
      const ss = this.sharedStrings;
      const styles2 = this.styles;
      let curRow = null;
      let prevCol = -1;
      let yielded = 0;
      let lastProgressPct = -1;
      for (const r of records(bytes)) {
        const d = r.data;
        if (r.type === BRT_ROW_HEADER && d.length >= 4) {
          if (curRow) {
            yield curRow;
            yielded++;
            if (onProgressIter && yielded % 1e3 === 0) {
              const pct = Math.min(99, Math.floor(yielded / 1e5 * 100));
              if (pct !== lastProgressPct) {
                onProgressIter(`Row ${yielded}`, pct);
                lastProgressPct = pct;
              }
              await tick();
            }
            if (maxRows !== void 0 && yielded >= maxRows) return;
          }
          curRow = { row: readU32(d, 0), cols: {} };
          prevCol = -1;
          continue;
        }
        if (!curRow) continue;
        if (r.type >= BRT_CELL_BLANK && r.type <= BRT_FMLA_ERROR) {
          const col = readU32(d, 0);
          const ixf = d.length >= 6 ? d[4] | d[5] << 8 | (d[5] & 128 ? 4294901760 : 0) : void 0;
          const cell = readCell(r.type, d, 8, ss);
          if (cell) {
            cell.ixf = ixf;
            applyDateMeta(cell, ixf, styles2);
            curRow.cols[col] = cell;
          }
          prevCol = col;
        } else if (r.type >= BRT_SHORT_BLANK && r.type <= BRT_SHORT_ISST) {
          const col = prevCol + 1;
          const ixf = d.length >= 4 ? readU16(d, 2) : void 0;
          const cell = readShortCell(r.type, d, 4, ss);
          if (cell) {
            cell.ixf = ixf;
            applyDateMeta(cell, ixf, styles2);
            curRow.cols[col] = cell;
          }
          prevCol = col;
        }
      }
      if (curRow) {
        yield curRow;
        yielded++;
      }
      if (onProgressIter) onProgressIter(`Done (${yielded} rows)`, 100);
    },
    async collectSheet(sheetIndex, iterOpts = {}) {
      const name = sheetNames[sheetIndex] ?? `Sheet${sheetIndex + 1}`;
      const rows = [];
      for await (const row of this.iterSheetRows(sheetIndex, iterOpts)) rows.push(row);
      const totalCells = rows.reduce((a, r) => a + Object.keys(r.cols).length, 0);
      return { name, rows, totalCells };
    }
  };
  return handle;
}

// src/index.ts
function tick2() {
  return new Promise((r) => setTimeout(r, 0));
}
function normalizeOptions2(arg) {
  if (arg === void 0) return {};
  if (typeof arg === "function") return { onProgress: arg };
  return arg;
}
async function parseXlsb(data, options) {
  const opts = normalizeOptions2(options);
  const onProgress = opts.onProgress;
  const dumpBinaries = opts.dumpBinaries === true;
  const readXml = opts.readXml === true;
  const parsePivotCaches = opts.parsePivotCaches === true;
  const maxRowsPerSheet = opts.maxRowsPerSheet;
  const maxZipBytes = opts.maxZipBytes;
  onProgress?.("Decompressing ZIP...", 0);
  await new Promise((r) => setTimeout(r, 50));
  const u82 = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
  let zip;
  try {
    zip = unzipSync(u82);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error("ZIP decompression failed: " + msg, { cause: e });
  }
  if (maxZipBytes !== void 0) {
    let total = 0;
    for (const k of Object.keys(zip)) total += zip[k].length;
    if (total > maxZipBytes) {
      throw new XlsbSizeError(
        `Decompressed ZIP size ${total} bytes exceeds maxZipBytes limit ${maxZipBytes}`,
        maxZipBytes,
        total
      );
    }
  }
  const out = {
    sheets: [],
    sharedStrings: [],
    xmlFiles: {},
    binaryDumps: [],
    pivotCaches: [],
    summary: { fileCount: 0, totalRecords: 0 }
  };
  const wb = zip["xl/workbook.bin"];
  if (!wb) throw new Error("xl/workbook.bin not found");
  onProgress?.("Parsing workbook...", 5);
  const sheetNames = parseWorkbook(wb);
  await tick2();
  if (zip["xl/sharedStrings.bin"]) {
    onProgress?.("Parsing shared strings...", 10);
    out.sharedStrings = parseSharedStrings(zip["xl/sharedStrings.bin"]);
    await tick2();
  }
  let styles = null;
  if (zip["xl/styles.bin"]) {
    onProgress?.("Parsing styles...", 12);
    try {
      styles = parseStyles(zip["xl/styles.bin"]);
    } catch {
      styles = null;
    }
    await tick2();
  }
  for (let i = 0; i < sheetNames.length; i++) {
    const key = `xl/worksheets/sheet${i + 1}.bin`;
    const sd = zip[key];
    if (sd) {
      onProgress?.(`Sheet "${sheetNames[i]}"...`, 15 + Math.round(i / sheetNames.length * 20));
      const rows = parseSheet(sd, out.sharedStrings, { maxRows: maxRowsPerSheet, styles });
      const totalCells = rows.reduce((a, r) => a + Object.keys(r.cols).length, 0);
      out.sheets.push({ name: sheetNames[i], rows, totalCells });
      await tick2();
    }
  }
  if (parsePivotCaches) {
    const pcd1 = zip["xl/pivotCache/pivotCacheDefinition1.bin"];
    const pcd2 = zip["xl/pivotCache/pivotCacheDefinition2.bin"];
    const pcr1 = zip["xl/pivotCache/pivotCacheRecords1.bin"];
    const pcr2 = zip["xl/pivotCache/pivotCacheRecords2.bin"];
    if (pcd1 && pcr1) {
      onProgress?.("Pivot cache 1...", 33);
      out.pivotCaches.push(parsePivotCache("PivotCache1", pcd1, pcr1));
      await tick2();
    }
    if (pcd2 && pcr2) {
      onProgress?.("Pivot cache 2...", 34);
      out.pivotCaches.push(parsePivotCache("PivotCache2", pcd2, pcr2));
      await tick2();
    }
  }
  if (dumpBinaries) {
    const binPaths = Object.keys(zip).filter((k) => k.endsWith(".bin")).sort();
    const total = binPaths.length;
    let doneBins = 0;
    for (const path of binPaths) {
      doneBins++;
      const pct = 35 + Math.round(doneBins / total * 55);
      onProgress?.(`${path.split("/").pop()}...`, pct);
      const dump = dumpBinary(path, zip[path]);
      out.binaryDumps.push(dump);
      out.summary.fileCount++;
      out.summary.totalRecords += dump.recCount;
      await tick2();
    }
  } else {
    const binPaths = Object.keys(zip).filter((k) => k.endsWith(".bin"));
    out.summary.fileCount += binPaths.length;
  }
  await tick2();
  if (readXml) {
    const xmlPaths = Object.keys(zip).filter((k) => k.endsWith(".xml") || k.endsWith(".rels"));
    for (let i = 0; i < xmlPaths.length; i++) {
      const path = xmlPaths[i];
      if (i % 5 === 0) {
        onProgress?.(`XML files...`, 92 + Math.round(i / xmlPaths.length * 6));
        await tick2();
      }
      try {
        out.xmlFiles[path] = dec8.decode(zip[path]);
      } catch {
      }
    }
  }
  onProgress?.("Done", 100);
  await tick2();
  return out;
}
export {
  XlsbSizeError,
  isDateFormatId,
  isDateNumFmtString,
  openXlsb,
  parseStyles,
  parseXlsb
};
