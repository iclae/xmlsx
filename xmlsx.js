var sheet = require('./temp/sheet')
var data = require('./temp/data')
var util = require('./util')

// func
var Frozen = require('./mod/frozen')
var Write = require('./mod/write')
var Valid = require('./mod/valid')

// main
var XMLSX = function() {
  this._formats = sheet.getSheet()
  this._sheetData = data.getData()
}

// final
XMLSX.prototype.getXlsx = function(callback) {
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
XMLSX.prototype.write = function(data) {
  var i = new Write(data)
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
