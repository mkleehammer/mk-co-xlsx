
// Overview:
// https://msdn.microsoft.com/EN-US/library/office/gg278316.aspx

// TODO: Shared string runs https://msdn.microsoft.com/en-us/library/office/gg278314.aspx

// TODO: Scan number formats for custom date formats (look for Y, etc.).

'use strict';

var debug = require('debug')('mk-co-xlsx');
var moment = require('moment');
var ZipFile = require('./zipfile');
var util = require('util');

var XMLParser = require('mk-co-xml').Parser;

function ExcelReader(options) {
  options = options || {};
  this.tz = options.tz === 'utc' ? 'utc' : 'local';

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

  this.styles = null;
  // An array of style information used for determining if numbers are dates or
  // not.  Styles in Excel are identified by the zero-based index.  Values look
  // like: { numFmtId: 1, type: 'n' } { numFmtId: 14, type: 'd' }
  //
  // Excel hardcodes a few number formats.  The type field tells us if we think
  // it is a number or date.

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
      this.styles = yield this._readCellStyles();
      this.sheets = yield this._readSheets();

      this.dateSystem = yield this._readDateSystem();
    }
  },

  _parseXML: {
    value: function* _parseXML(filename, callback) {
      // Internal utility to read XML and close the stream.  The callback will
      // be called for each element.  Normally the entire XML file will be
      // processed, but you can return false from the callback to abort.

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
      // Reads date system from the workbookPr element.

      var system = 1900; // the default

      function callback(x) {
        if (x.tag === 'workbookPr') {
          if (x.attributes.date1904 === '1')
            system = 1904;
          return false;
        }
      }

      yield this._parseXML('xl/workbook.xml', callback);

      return system;
    }
  },

  _readCellStyles: {
    value: function* _readCellStyles() {
      var styles = [];

      var found = false;

      function callback(x) {
        if (x.tag === 'cellXfs') {

          if (x.type == 'elementStart') {
            found = true;
            return;
          }

          return false;
        }

        if (found && x.tag === 'xf' && x.type === 'elementStart') {
          var n = parseInt(x.attributes.numFmtId);
          styles.push({
            numFmtId: n,
            type: (n >= 14 && n <= 17) ? 'd' : 'n'
          });
        }
      }

      yield this._parseXML('xl/styles.xml', callback);

      return styles;
    }
  },

  _readSheets: {
    value: function* _readSheets() {

      var sheets = [];

      function callback(x) {
        if (x.type === 'elementStart' && x.tag === 'sheet') {
          sheets.push(x.attributes.name);
        }
      }

      yield this._parseXML('xl/workbook.xml', callback);

      return sheets;
    }
  },

  _readSharedStrings: {
    value: function*() {
      var filename = 'xl/sharedStrings.xml';

      if (!this.zipfile.findFile(filename))
        return;

      var a = [];
      var inT = false;

      function callback(x) {
        if (inT && x.type === 'text') {
          a.push(x.text);
          inT = false;
        } else {
          inT = (x.type === 'elementStart' && x.tag === 't');
        }
      }

      yield this._parseXML(filename, callback);

      return a;
    }
  },

  getSheet: {
    value: function* getSheet(sheetIndex) {
      // Returns a row stream for the given sheet.
      if (sheetIndex < 0 || sheetIndex >= this.sheets.length)
        throw new Error('Invalid sheet index ' + sheetIndex);

      var filename = 'xl/worksheets/sheet' + (sheetIndex + 1) + '.xml';
      var stream = yield this.zipfile.createReadStream(filename);

      var options = { tz: this.tz };

      return new Sheet(this, filename, stream, options);
    }
  }
});

function Sheet(workbook, filename, stream, options) {
  this.workbook = workbook;
  this.parser = new XMLParser();
  this.parser.readStream(filename, stream);

  this.tz = options.tz;
}

Sheet.prototype = Object.create(null, {

  read: {
    value: function* read() {
      if (!this.parser)
        return null;

      var x = yield* this._skipTo('row');
      if (!x)
        return null;

      var row = {
        index: parseInt(x.attributes.r),
        cols: []
      };

      while (true) {
        x = yield this.parser.read();
        if (!x) {
          this.close();
          return null;
        }

        if (x.type === 'elementEnd' && x.tag === 'row')
          return row;

        if (x.type === 'elementStart' && x.tag === 'c') {
          yield this._readCell(row.cols, x);
        }
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

  // Dates in SpreadsheetML:
  // http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/02/16/dates-in-spreadsheetml.aspx
  // http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2012/03/08/dates-in-strict-spreadsheetml-files.aspx

  _readCell: {
    value: function* _readCell(cols, x) {
      var ref = x.attributes.r;
      var type  = x.attributes.t;
      var style = x.attributes.s ? parseInt(x.attributes.s) : null;

      var v = yield this._skipTo('v', 'c');
      if (!v || v.tag !== 'v')
        throw new Error('Unable to find value for cell ' + ref);

      v = yield this.parser.read();
      if (!v || v.type !== 'text')
        throw new Error('Expected value text for cell ' + ref + ' not ' + JSON.stringify(v));

      var value;

      if (type === 's') {
        // string
        value = this.workbook.sharedStrings[parseInt(v.text)];
      } else if (type === 'd') {
        value = parseISODate(this.tz, v.text);
      } else if (style) {
        var s = this.workbook.styles[style];
        if (s && s.type === 'd')
          value = parseSerialDate(this.workbook.dateSystem, this.tz, v.text);
        else
          value = parseFloat(v.text);
      } else {
        value = parseFloat(v.text);
      }

      var index = this._indexFromRef(ref);
      cols[index] = value;
    }
  },

  close: {
    value: function() {
      if (this.parser) {
        try { this.parser.close(); } catch (err) { }
        this.parser = null;
      }
    }
  },

  _skipTo: {
    value: function* _skipTo() {
      // _skipTo(tag1, tag2, ...) -> element or null
      //
      // Accepts a list of tags to search for.  Reads and discards elements
      // until it finds one of the tags is found, in which case it returns the
      // element, or until end-of-file is reached in which case the parser is
      // closed and null is returned.

      var tags = arguments;

      while (true) {
        var x = yield this.parser.read();
        if (!x) {
          this.close();
          return null;
        }

        if (x.tag) {
          for (var i = 0, c = tags.length; i < c; i++)
            if (x.tag === tags[i])
              return x;
        }

        // debug('_skipTo: looking for "%s" - skipping %j', tags, x);
      }
    }
  },

});

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
