
// A node zip-file reader that provides a readable stream for each entry.
//
// I needed a co-friendly way to read zip files.

'use strict';

var fs = require('fs');
var cofs = require('co-fs');
var zlib = require('zlib');
var debug = require('debug')('mk-co-xlsx:zipfile');

module.exports = ZipFile;

var EOCD_SIGNATURE = 0x06054b50;
var EOCD_LEN = 22; // Does not include comment
var EOCD_MAX_COMMENT = 65535;
var CD_SIGNATURE = 0x02014b50;
var CD_LEN = 46;
var LFH_SIGNATURE = 0x04034b50;
var LFH_LEN = 30;
var COMPRESS_STORE = 0;
var COMPRESS_DEFLATE = 8;

function ZipFile(filename, stat) {
  this.filename = filename;
  this.stat = stat;
  this.files = [];
}

ZipFile.openFile = function* openFile(filename) {
  var fd = yield cofs.open(filename, 'r');

  try {
    var stat = yield cofs.fstat(fd);

    if (stat.size < EOCD_LEN) {
      throw new Error(filename + ' is not a zip file');
    }

    var zipfile = new ZipFile(filename, stat);

    yield zipfile._readEOCD(fd);

    return zipfile;

  } finally {
    if (fd) {
      try { yield fd.close(fd); } catch (error) { }
    }
  }
};

ZipFile.prototype._readEOCD = function* _readEOCD(fd) {

  var maxlen = EOCD_LEN + EOCD_MAX_COMMENT;
  
  var buffer = new Buffer(maxlen);
  yield cofs.read(fd, buffer, 0, maxlen, this.stat.size - maxlen);

  // var EOCD_SIGNATURE = 0x06054b50;

  var i = maxlen - 1 - EOCD_LEN + 4;
  while (i >= 0) {
    if (buffer[i] === 0x50 && buffer.readUInt32LE(i) === EOCD_SIGNATURE)
      break;
    i--;
  }
  if (i < 0)
    throw new Error('Did not find EOCD');

  buffer = buffer.slice(i, i+EOCD_LEN);
  
  if (buffer.readUInt16LE(4) !== 0)
    throw new Error('Multi-disk zip files are not supported');

  var count  = buffer.readUInt16LE(8);
  var offset = buffer.readUInt32LE(16);
  var length = buffer.readUInt32LE(12);

  buffer = new Buffer(CD_LEN);
  var lfh = new Buffer(LFH_LEN);
  var stringBuffer;

  for (i = 0; i < count; i++) {
    yield cofs.read(fd, buffer, 0, CD_LEN, offset);

    if (buffer.readUInt32LE(0) !== CD_SIGNATURE)
      throw new Error('Did not find CD ' + i);
    
    var filenameLength = buffer.readUInt16LE(28);
    var extraLength    = buffer.readUInt16LE(30);
    var commentLength  = buffer.readUInt16LE(32);

    if (!stringBuffer || stringBuffer.length < filenameLength)
      stringBuffer = new Buffer(filenameLength);
    yield cofs.read(fd, stringBuffer, 0, filenameLength, offset + CD_LEN);
    var filename = stringBuffer.toString('ascii', 0, filenameLength);
    
    // Note: The extra field length in the Central Directory Entry is not the
    // same as that in the Local File Header.  The filename should be the same,
    // but since we already have the LFH we might as well use its lengths to be
    // sure.

    var lfhOffset = buffer.readUInt32LE(42);

    yield cofs.read(fd, lfh, 0, LFH_LEN, lfhOffset);
    console.assert(lfh.readUInt32LE(0) === LFH_SIGNATURE);

    var dataOffset = lfhOffset + LFH_LEN + 
        lfh.readUInt16LE(26) + // filename
        lfh.readUInt16LE(28);  // comment


    var file = {
      filename: filename,

      compressionMethod: buffer.readUInt16LE(10),
      compressedSize: buffer.readUInt32LE(20),
      uncompressedSize: buffer.readUInt32LE(24),

      dataOffset: dataOffset
    };

    this.files.push(file);

    offset += CD_LEN + filenameLength + extraLength + commentLength;
  }
};


ZipFile.prototype.findFile = function findFile(filename) {
  // Returns the directory entry for the given filename if it is in the zip
  // file.  Otherwise returns null.

  for (let i = 0; i < this.files.length; i++)
    if (this.files[i].filename === filename)
      return this.files[i];
  return null;
};

ZipFile.prototype._findFile = function _findFile(filename) {
  for (let i = 0; i < this.files.length; i++) {
    if (this.files[i].filename === filename)
      return this.files[i];
  }
  throw new Error('There is no file named "' + filename + '" in this zip file');
};


ZipFile.prototype.createReadStream = function* createReadStream(filename) {
  // Returns a zipfile entry object for the given filename.

  var file = this._findFile(filename);
  
  var stream = fs.createReadStream(this.filename, {
    start: file.dataOffset, 
    end: file.dataOffset + file.compressedSize - 1
  });

  switch (file.compressionMethod) {
  case COMPRESS_STORE:
    break;
  case COMPRESS_DEFLATE:
    stream = stream.pipe(zlib.createInflateRaw());
    break;
  default:
    throw new Error('Unsupported compression method. file=' + file.filename + ' method=' + file.compressionMethod);
  }
  
  return stream;
};
