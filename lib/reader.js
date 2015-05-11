
// Overview:
// https://msdn.microsoft.com/EN-US/library/office/gg278316.aspx

// Todo Items
// ----------
//
// Shared string runs https://msdn.microsoft.com/en-us/library/office/gg278314.aspx

'use strict';

var util = require('util');
var debug = require('debug')('mk-co-xlsx');
var moment = require('moment');
var ZipFile = require('./zipfile');

var XMLParser = require('mk-co-xml').Parser;

function ExcelReader(options) {
  options = options || {};
  this.tz = options.tz === 'utc' ? 'utc' : 'local';
  this.skipEmptyRows = (options.skipEmptyRows != null) ? options.skipEmptyRows : true;

  this.name    = null;
  this.zipfile = null;
  this.source  = null;

  this.sheets = [];
  // The names of the sheets.

  this.sharedStrings = [];

  this.tz = null;
  // Determines the time zone that dates will be converted to.  Valid values are
  // 'utc' and 'local'.

  this.dateSystem = null;
  // Determines how serial dates are calculated.  Can be 1900 or 1904.

  this.dateFlags = null;
  // An array of Booleans corresponding to each defined style.  (That is,
  // element 0 in this array is the flag for style "0" in the spreadsheet.)  If
  // the flag is true the style represents a date.
  //
  // http://blogs.msdn.com/b/brian_jones/archive/2007/05/29/simple-spreadsheetml-file-part-3-formatting.aspx

}

module.exports = ExcelReader;

ExcelReader.prototype = Object.create(null, {

  fromFilename: {
    value: function*(filename, options) {
      options = options || {};

      this.tz = options.tz || 'utc';

      this.name = filename;
      this.zipfile = yield ZipFile.openFile(filename);
      this.sharedStrings = yield this._readSharedStrings();
      this.dateFlags = yield this._readDateFlags();

      this.sheets = yield this._readSheets();
      this.dateSystem = yield this._readDateSystem();
    }
  },

  _parseXML: {
    value: function* _parseXML(filename, callback) {
      // Internal utility to read an entire XML file, pass each event through
      // `callback`, and close the stream.  This is used for the simpler XML
      // tasks, not the worksheet.
      //
      // The callback can return false to abort processing.

      var parser = new XMLParser();
      var stream = yield this.zipfile.createReadStream(filename);
      try {
        parser.readStream(filename, stream);

        while (true) {
          var x = yield parser.read();
          if (!x)
            break;

          var result = callback(x);
          if (result === false)
            break;
        }
      } finally {
        try {
          parser.close();
        }
        catch (err) {
        }
      }
    }
  },

  _readDateSystem: {
    value: function* _readDateSystem() {
      // Reads date system from the workbookPr element.  The default date system
      // is 1900, but the workbookPr element can have date1904="1".

      var system = 1900; // the default

      yield this._parseXML('xl/workbook.xml', function (e) {
        if (e.tag === 'workbookPr') {
          if (e.attrs.date1904 === '1')
            system = 1904;
          return false;
        }
      });

      return system;
    }
  },

  _readDateFlags: {
    value: function* _readDateFlags() {
      // Returns an array of Booleans that determines whether each style should
      // be converted to date.

      var inCellXfs = false;
      // We want the <xf> elements in the <cellXfs> but there are others.

      var styleFormats = [];
      var mapFmtToFlag = {};

      yield this._parseXML('xl/styles.xml', function(e) {

        if (e.type === 'start' && e.tag === 'numFmt') {
          // <numFmt numFmtId="164" formatCode="\ mm\/dd\/yyyy"/>
          mapFmtToFlag[e.attrs.numFmtId] = isFormatCodeDate(e.attrs.formatCode);
          return;
        }

        if (e.tag === 'cellXfs') {
          if (e.type == 'start') {
            inCellXfs = true;
            return;
          } else {
            return false; // stop the XML stream - we found what we need
          }
        }

        if (inCellXfs && e.tag === 'xf' && e.type === 'start') {
          styleFormats.push(e.attrs.numFmtId);
        }
      });

      var flags = styleFormats.map(function(numFmtId) {
        var f = mapFmtToFlag[numFmtId];
        if (f === undefined) {
          var n = parseInt(numFmtId);
          f = (n >= 14 && n <= 17);
        }
        return f;
      });

      return flags;
    }
  },

  _readSheets: {
    value: function* _readSheets() {
      // Returns an array of the sheet names.

      var sheets = [];

      yield this._parseXML('xl/workbook.xml', function(e) {
        if (e.type === 'start' && e.tag === 'sheet') {
          sheets.push(e.attrs.name);
        }
      });

      return sheets;
    }
  },

  _readSharedStrings: {
    value: function*() {
      // Returns the shared string table as an array.  Cells with text have
      // values that are simply indexes into this table.

      var filename = 'xl/sharedStrings.xml';

      if (!this.zipfile.findFile(filename))
        return;

      var strings = [];
      var inT = false;

      yield this._parseXML(filename, function(e) {
        if (inT && e.type === 'text') {
          strings.push(e.text);
          inT = false;
        } else {
          inT = (e.type === 'start' && e.tag === 't');
        }
      });

      return strings;
    }
  },

  getSheet: {
    value: function* getSheet(sheetIndex) {
      // Returns a Sheet object for the given sheet.  The index is the
      // zero-based index of the sheet, so the first sheet is 0.

      if (sheetIndex < 0 || sheetIndex >= this.sheets.length)
        throw new Error('Invalid sheet index ' + sheetIndex);

      var filename = 'xl/worksheets/sheet' + (sheetIndex + 1) + '.xml';
      var stream = yield this.zipfile.createReadStream(filename);

      return new Sheet(this, filename, stream);
    }
  }
});

function Sheet(workbook, filename, stream) {
  this.workbook = workbook;
  this.parser = new XMLParser();
  this.parser.readStream(filename, stream);

  this.rows = [];
  // Buffered up rows that will be returned by read().

  this.events = [];
  // Buffered up events that will create rows when we get more events.  This
  // occurs when we find a <row> but the </row> hasn't been read yet.
}

Sheet.prototype = Object.create(null, {

  closed: {
    get: function() {
      return this.parser == null && this.rows.length === 0;
    }
  },

  read: {
    value: function* read() {
      while (true) {
        if (this.rows.length)
          return this.rows.shift();

        if (!this.parser)
          return null;

        yield* this.refill();
      }
    }
  },

  readMany: {
    value: function* readMany(count) {
      // Returns an array of rows.
      //
      // If `count` is null or undefined (not provided) all buffered rows are
      // returned, which could be zero (an empty array).
      //
      // Otherwise it will attempt to return `count` records, reading more from
      // the stream if necessary.  If the stream has less than `count`, however
      // many could be read are returned.  This may return an empty array.
      //
      // Returns null when all rows have been read.

      var rows;

      console.log('>>>> readMany: count=%s rows=%s parser=%s', count, this.rows.length, (this.parser != null));

      if (!count) {
        if (!this.rows.length && this.parser) {
          yield* this.refill();

          console.log('>>>> refill: rows=%s parser=%s', this.rows.length, (this.parser != null));
        }
        rows = this.rows;
        this.rows = [];
        return rows;
      }

      while (this.rows.length < count && this.parser) {
        yield* this.refill();
      }

      count = Math.min(count, this.rows.length); // in case of end-of-stream

      rows      = this.rows.slice(0, count);
      this.rows = this.rows.slice(count);

      return rows;
    }
  },

  putBack: {
    value: function putBack(row) {
      this.rows.unshift(row);
    }
  },

  putBackMany: {
    value: function putBackMany(rows) {
      this.rows = rows.concat(this.rows);
    }
  },

  refill: {
    value: function*() {
      var before = this.events.length;

      var newEvents = yield* this.parser.readMany();
      if (newEvents == null) {
        this.parser.close();
        this.parser = null;
        return;
      }

      this.events = this.events.concat(newEvents);

      var after = this.events.length;

      this.parse();
    }
  },

  parse: {
    value: function() {
      // Parses `this.events` into as many rows as possible and discards
      // consumed events.

      var e = this.events;
      var l = e.length;

      var i = 0;

      while (i < l) {
        // Scan ahead for the next <row>.

        if (e[i].tag !== 'row' || e[i].type !== 'start') {
          i += 1;
          continue;
        }

        // Now find </row>.
        for (var end = i + 1; end < l; end++) {
          if (e[end].type === 'end' && e[end].tag === 'row') {
            break;
          }
        }

        if (end === l) {
          // The row is incomplete.  Discard everything we've consumed so far
          // and exit.  The calling code will refill the buffer and call us
          // again to pick up where we left off.
          this.events = this.events.slice(i);
          return;
        }

        var row = this.parseRow(e, i+1, end-1);
        if (row)
          this.rows.push(row);

        i = end + 1;
      }

      // We've consumed all events.
      this.events = [];
    }
  },

  parseRow: {
    value: function parseRow(events, start, end) {

      // Note: We're guaranteed there is an event *after* end which is </row>.
      // This means we can peek ahead without checking for falling off.

      var cols = [];

      var i = start;
      while (i <= end) {

        var e = events[i];
        if (e.tag !== 'c' || e.type !== 'start') {
          i += 1;
          continue;
        }

        var before = i;

        // Scan for the end of the column.  If we see the value tag, grab it.
        // Note that

        var value;
        for (++i; i <= end && events[i].tag !== 'c'; i++) {
          if (events[i].text)
            value = events[i].text;
        }
        if (i > end)
          throw new Error('Missing </c>');

        i++; // skip </c>

        // It is possible to not have a value (<v>) in which case we ignore,
        // leaving undefined in `cols`.

        if (value !== undefined) {
          var ref   = e.attrs.r;
          var index = this._indexFromRef(ref);
          var type  = e.attrs.t;

          if (type === 's') {
            // The value is an index into the shared string table.
            value = this.workbook.sharedStrings[parseInt(value)];
          } else if (type === 'd') {
            value = parseISODate(this.workbook.tz, value);
          } else if (e.attrs.s) {
            // It has a style which will be an index into the global flags
            // array.  See if we can infer the type from the style.
            var isDate = this.workbook.dateFlags[parseInt(e.attrs.s)];
            if (isDate) {
              value = parseSerialDate(this.workbook.dateSystem, this.workbook.tz, value);
            } else {
              value = parseFloat(value);
            }
          } else {
            value = parseFloat(value);
          }

          cols[index] = value;
        }
      }

      if (cols.length || !this.workbook.skipEmptyRows) {
        return cols;
      }
    }
  },

  _indexFromRef: {
    // Given a cell reference like "A1", returns the zero-based column index:
    // A1 -> 0 (A is the first)
    // Z1 -> 25

    value: function _indexFromRef(ref) {
      var m = /^([A-Z]+)\d+$/.exec(ref);
      var chars = m[1];
      var index = 0;
      for (var i = 0; i < chars.length; i++) {
        index *= 26;
        index += chars[i].charCodeAt(0) - 64;
      }
      index -= 1;
      return index;
    }
  },

  close: {
    value: function() {
      if (this.parser) {
        try { this.parser.close(); } catch (err) { }
        this.parser = null;
      }
    }
  }
});

// Dates in SpreadsheetML
//
// http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/02/16/dates-in-spreadsheetml.aspx
// http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/03/08/dates-in-strict-spreadsheetml-files.aspx

function parseISODate(tz, value) {
  if (value[2] === ':') {
    // Just a time.  For now return as a string, but we should probably set to
    // 0001-01-01 or 1900 or whatever Excel users would expect.
    return value;
  }

  var m;

  if (tz === 'utc')
    m = moment.utc(value, moment.ISO_8601);
  else
    m = moment(value, moment.ISO_8601);

  return m.toDate();
}

function parseSerialDate(dateSystem, tz, value) {

  value = parseFloat(value);

  var date    = Math.floor(value);
  var seconds = Math.round(86400 * (value - date));

  // There are two date systems, one starting in 1900 and the other in 1904.
  // Normalize to 1900.

  if (dateSystem === 1904)
    date += 1462;

  // Lotus 123 had a bug that considered 1900 a leap year so it thought there
  // was a 1900-02-29.  Excel (purposely) copied the bug for backwards
  // compatibility.

  if (date >= 60) {
    date -= 1;
  }

  // Convert to unix timestamp which is easy to use with moment.

  var unix = ((date - 1) * 86400) + seconds - 2208988800;

  return moment.unix(unix).toDate();
}

function isFormatCodeDate(code) {
  // Determine if the Excel format code is a date format.
  //
  // In general we just need to see 'm', 'd', and 'y' somewhere in the code.
  // However, there are some things to consider:
  //
  // * A format code can have 4 parts separated by ';'.  We're just going to use
  //   the first (positive) format.  If it is a date, that's probably all that
  //   exists.
  // * Literal text can be inserted using either a backslash or quotes.  (I'm
  //   not sure if a blackslash can be used in quotes to escape them.)
  // * An 'm' can actually be minutes if used after an 'h' or before an 's'.  If
  //   we see a 'y' and 'd', we'll assume at least one of the 'm' characters is
  //   a month.


  var FLAGS = {
    'd' : 0x01,
    'm' : 0x02,
    'y' : 0x04
  };
  var ALL = 0x07;

  var esc = false;
  var quo = false;

  var seen = 0x00;

  for (var i = 0; i < code.length && seen !== ALL; i++) {
    if (esc) {
      esc = false;
      continue;
    }

    var ch = code[i];

    if (quo) {
      if (ch === '"')
        quo = false;
      continue;
    }

    if (ch === ';')
      break;

    if (ch === '\\')
      esc = true;
    else if (ch === '"')
      quo = true;
    else {
      seen |= (FLAGS[ch] || 0);
    }
  }

  return seen == ALL;
}
