// util
var util = require('./lib/util')

// method
var Frozen = require('./lib/mod/frozen')
var Entry = require('./lib/mod/entry')
var Valid = require('./lib/mod/valid')

// main
var XMLSX = function(init) {
  var initData = util.readXMLbuffer(init)
  this._formats = initData.formats
  this._sheetData = initData.sheetData
}

XMLSX.prototype._getSheet = function() {
  return {
    formats: this._formats,
    sheetData: this._sheetData,
  }
}

// ------- step func
// frozen row or col
XMLSX.prototype.frozen = function(range) {
  var f = new Frozen(range)
  f.setFrozen(this._formats)
  return this
}

// write data [[row],[row],[row]]
XMLSX.prototype.entry = function(data) {
  var i = new Entry(data)
  i.setSheet({
    formats: this._formats,
    sheetData: this._sheetData,
  })
  return this
}

// valid data [{A1: [1, 2, 3]}]
XMLSX.prototype.valid = function(validArray) {
  var v = new Valid(validArray)
  v.setSheet(this._formats)
  return this
}

// ------- final func
// get xlsx buffer
XMLSX.prototype.done = function(callback) {
  util.closeSet(
    {
      formats: this._formats,
      sheetData: this._sheetData,
    },
    callback
  )
}

// get xlsx object
XMLSX.prototype.output = function() {
  var format = {
    frozen: Frozen.analysis(this._formats),
    valid: Valid.analysis(this._formats),
  }
  var sheetData = Entry.analysis(this._getSheet())
  return {
    format: format,
    sheetData: sheetData,
    colStyleData: util.getColStyle(sheetData),
  }
}

module.exports = XMLSX
