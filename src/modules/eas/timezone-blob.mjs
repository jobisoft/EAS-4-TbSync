/**
 * 172-byte packed VTIMEZONE blob used by EAS Calendar (codepage 4) up to
 * AS 14.x. The base64 form is what goes on the wire as
 * `<Calendar:Timezone>`.
 *
 * Layout:
 *   @000  utcOffset       LONG (i32 LE) - bias from local to UTC, minutes
 *   @004  standardName    32 × WCHAR (UTF-16 LE)
 *   @068  standardDate    SYSTEMTIME (8 × u16 LE)
 *   @084  standardBias    LONG (i32 LE)
 *   @088  daylightName    32 × WCHAR (UTF-16 LE)
 *   @152  daylightDate    SYSTEMTIME
 *   @168  daylightBias    LONG (i32 LE) - typically -60
 *
 * SYSTEMTIME for DST switch dates encodes "Nth weekday of month":
 *   wYear=0, wMonth=1..12, wDayOfWeek=0..6 (Sun..Sat), wDay=1..5 (5 = last)
 *
 * Port of `EAS-4-TbSync/content/includes/tools.js` TimeZoneDataStructure.
 */

const SIZE = 172;

export class TimeZoneBlob {
  constructor() {
    this.buf = new DataView(new ArrayBuffer(SIZE));
  }

  set easTimeZone64(b64) {
    for (let i = 0; i < SIZE; i++) this.buf.setUint8(i, 0);
    if (!b64) return;
    const content = atob(b64);
    const len = Math.min(content.length, SIZE);
    for (let i = 0; i < len; i++) this.buf.setUint8(i, content.charCodeAt(i));
  }

  get easTimeZone64() {
    let s = "";
    for (let i = 0; i < SIZE; i++)
      s += String.fromCharCode(this.buf.getUint8(i));
    return btoa(s);
  }

  _getStr(off) {
    let s = "";
    for (let i = 0; i < 32; i++) {
      const cc = this.buf.getUint16(off + i * 2, true);
      if (cc === 0) break;
      s += String.fromCharCode(cc);
    }
    return s;
  }

  _setStr(off, str) {
    for (let i = 0; i < 32; i++) this.buf.setUint16(off + i * 2, 0, true);
    const limit = Math.min(str.length, 32);
    for (let i = 0; i < limit; i++) {
      this.buf.setUint16(off + i * 2, str.charCodeAt(i), true);
    }
  }

  _getSystemtime(off) {
    const buf = this.buf;
    return {
      get wYear() {
        return buf.getUint16(off + 0, true);
      },
      get wMonth() {
        return buf.getUint16(off + 2, true);
      },
      get wDayOfWeek() {
        return buf.getUint16(off + 4, true);
      },
      get wDay() {
        return buf.getUint16(off + 6, true);
      },
      get wHour() {
        return buf.getUint16(off + 8, true);
      },
      get wMinute() {
        return buf.getUint16(off + 10, true);
      },
      get wSecond() {
        return buf.getUint16(off + 12, true);
      },
      get wMilliseconds() {
        return buf.getUint16(off + 14, true);
      },
      set wYear(v) {
        buf.setUint16(off + 0, v, true);
      },
      set wMonth(v) {
        buf.setUint16(off + 2, v, true);
      },
      set wDayOfWeek(v) {
        buf.setUint16(off + 4, v, true);
      },
      set wDay(v) {
        buf.setUint16(off + 6, v, true);
      },
      set wHour(v) {
        buf.setUint16(off + 8, v, true);
      },
      set wMinute(v) {
        buf.setUint16(off + 10, v, true);
      },
      set wSecond(v) {
        buf.setUint16(off + 12, v, true);
      },
      set wMilliseconds(v) {
        buf.setUint16(off + 14, v, true);
      },
    };
  }

  get standardDate() {
    return this._getSystemtime(68);
  }
  get daylightDate() {
    return this._getSystemtime(152);
  }

  get utcOffset() {
    return this.buf.getInt32(0, true);
  }
  set utcOffset(v) {
    this.buf.setInt32(0, v, true);
  }

  get standardBias() {
    return this.buf.getInt32(84, true);
  }
  set standardBias(v) {
    this.buf.setInt32(84, v, true);
  }
  get daylightBias() {
    return this.buf.getInt32(168, true);
  }
  set daylightBias(v) {
    this.buf.setInt32(168, v, true);
  }

  get standardName() {
    return this._getStr(4);
  }
  set standardName(v) {
    this._setStr(4, v);
  }
  get daylightName() {
    return this._getStr(88);
  }
  set daylightName(v) {
    this._setStr(88, v);
  }
}

/** True when the blob is all zero (server didn't send useful TZ data). */
export function isAllZero(b64) {
  if (!b64) return true;
  try {
    const s = atob(b64);
    for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) !== 0) return false;
    return true;
  } catch {
    return true;
  }
}
