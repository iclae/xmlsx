var util = require('./lib/util')
var Store = require('./lib/store')

// mod
var Frozen = require('./lib/mod/frozen')
var Entry = require('./lib/mod/entry')
var Valid = require('./lib/mod/valid')

// main
function XMLSX(buffer) {
  this._store = new Store()
  this._store.cache(buffer)
}

// ------- step func
// frozen row TODO: frozen col
XMLSX.prototype.frozen = function(range) {
  var f = new Frozen(range)
  f.setFrozen(this._store)
  return this
}

// write data [[row],[row],[row]]
XMLSX.prototype.entry = function(data) {
  var e = new Entry(data)
  e.setSheet(this._store)
  return this
}

// valid data [{A1: [1, 2, 3]}, {'B1:B100': [a, b, c]}]
XMLSX.prototype.valid = function(validArray) {
  var v = new Valid(validArray)
  v.setSheet(this._store)
  return this
}

// ------- final func
// get xlsx buffer
XMLSX.prototype.done = function(callback) {
  this._store.build(callback)
}

// get xlsx object
XMLSX.prototype.output = function() {
  return this._store.clear()
}

module.exports = XMLSX
