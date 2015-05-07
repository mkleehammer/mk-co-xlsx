
var assert = require("assert");

var path = require('path');
var co = require('co');

var ExcelReader = require('../lib/reader.js');


describe('basic', function() {
  it('should parse a mix', function(done) {

    // The file small.xlsx has two sheets.  The second sheet has a mix of text,
    // dates, numbers, and currency.

    var expectedRows = [
      ['Text', 'Date', 'Number', 'Styled Num'],
      ['Hello',  new Date(Date.UTC(2001, 0, 1)), 134, 100],
      ['Sailor', new Date(Date.UTC(2003, 0, 2)), 3.14, 200],
      null
    ];

    co(function*() {
      var reader = new ExcelReader();

      var fqn = path.join(__dirname, 'small.xlsx');
      yield reader.fromFilename(fqn);

      var sheet = yield reader.getSheet(1); // (2nd sheet)

      var got = null;

      for (var i = 0; i < expectedRows.length; i++) {
        var expected = expectedRows[i];
        got = yield sheet.read();

        if (got === null && expected === null)
          break;

        assert.deepEqual(got.cols, expected);
      }

      assert(got == null);
    })
      .then(function() {
        done();
      }, function(err) {
        console.log(err.stack);
        done(err);
      });
  });
});
