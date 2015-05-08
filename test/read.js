#!/usr/bin/env node --harmony

// A command line utility that simply reads a file.

'use strict';

var path = require('path');
var co = require('co');
var ExcelReader = require('../lib/reader.js');

var debug = require('debug')('mk-co-xlsx');

function* main() {
  var program = require('commander');
  program
    .version('1.0.0')
    .usage('[options] file')
    .option('-v --verbose', 'print more', increaseVerbosity, 0)
    .option('--max [max]', 'stop after max records', parseInt, 0)
    .parse(process.argv);

  if (program.args.length !== 1) {
    program.help();
    console.log('\Pass one and only one filename');
    process.exit(1);
  }

  var reader = new ExcelReader({ skipEmptyRows: false });
  var fqn = path.join(__dirname, program.args[0]);
  yield reader.fromFilename(fqn);

  var sheet = yield reader.getSheet(0);

  var rows = 0;

  var maxCount = program.max || 999999;

  while (true) {
    var row = yield sheet.read();
    if (!row)
      break;

    rows += 1;

    if (rows >= maxCount) {
      console.log('--max reached');
      break;
    }

    if (program.verbose > 0 && program.verbose <= 2) {
      if ((rows % 100) === 0) {
        if (program.verbose === 1)
          console.log('row %s', rows);
        else
          console.log('row %s %j', rows, row);
      }
    } else if (program.verbose > 2) {
      console.log(row);
    }
  }

  console.log('rows:', rows);
}

function increaseVerbosity(v, total) {
  return total + 1;
}

co(main)
  .then(function () { },
        function (err) {
          console.error(err.stack);
        });
