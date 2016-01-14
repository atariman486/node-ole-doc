# ole-doc

Read streams from an OLE Compound Document, e.g. StructuredStorage.

## Example Usage

### Normal
```
OleDoc = require('../lib/ole-doc').OleCompoundDoc;
var doc = new OleDoc( 'filename.ext' );
doc.on('err', function(err) {
  // do something with err
});
doc.on('ready', function() {
  var stream = doc.storage('StorageName').stream('StreamName');
  // do something with stream
});
doc.read();
```

### With Custom Header

```
var headerSize = 24;
doc.read(headerSize, function(buffer) {
  // do something with buffer
  // return false to stop reading (if you don't like the header for some reason)
  // return true to continue reading like normal
});
```
