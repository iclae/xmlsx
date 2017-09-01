// source
var sheet = require('./lib/temp/sheet')
var data = require('./lib/temp/data')

// util
var util = require('./lib/util')

// method
var Frozen = require('./lib/mod/frozen')
var Entry = require('./lib/mod/entry')
var Valid = require('./lib/mod/valid')

// main
var XMLSX = function() {
  this._formats = sheet.getSheet()
  this._sheetData = data.getData()
}

// final
XMLSX.prototype.done = function(callback) {
  util.closeSet(
    {
      formats: this._formats,
      sheetData: this._sheetData,
    },
    callback
  )
}

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

module.exports = XMLSX
