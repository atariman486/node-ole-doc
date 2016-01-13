chai = require('chai');
expect = chai.expect;

path = require('path');
Buffer = require('buffer').Buffer;

oleDoc = require('../lib/ole-doc').OleCompoundDoc;

describe('A missing Word document' , function() {

  it('should notify an error on open', function(done) {
    var filename = path.resolve(__dirname, "data/missing.doc");
    document = new oleDoc(filename);
    document.on('err', function(err) {
      expect(err).to.exist;
      expect(err).to.match(/missing\.doc/);
      done();
    });
    document.on('ready', function() {
      done("ready should not be called for a missing file");
    });
    document.read();
  });

});

describe('A Word document', function() {

  it('can be opened correctly', function(done) {
    var filename = path.resolve(__dirname, "data/test1.doc");
    doc = new oleDoc(filename);
    doc.on('err', function(err) {
      done(err);
    });
    doc.on('ready', function() {
      done();
    });
    doc.read();
  });

  it('generates a valid Word stream', function(done) {
    var filename = path.resolve(__dirname, "data/test1.doc");
    doc = new oleDoc(filename);
    doc.on('err', function(err) {
      done(err);
    });
    doc.on('ready', function() {
      chunks = [];
      var stream = doc.stream('WordDocument');
      stream.on('data', function(chunk) { chunks.push(chunk); });
      stream.on('error', function(error) { done(error); });
      stream.on('end', function() {
        var buffer = Buffer.concat(chunks);
        var magicNumber = buffer.readUInt16LE(0);
        expect(magicNumber).to.equal(0xa5ec);
        done();
      });
    });
    doc.read();
  });

});
