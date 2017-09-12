var Store = require('./lib/store')
var util = require('./lib/util')
// mods
var frozen = require('./lib/mod/frozen')
var valid = require('./lib/mod/valid')
var hide = require('./lib/mod/hide')

// main
function Xmlsx(buffer) {
  this._store = new Store()
  this._store.cache(buffer)
}

// Objec func
Xmlsx.getCell = util.getCellByXY
Xmlsx.formatDate = util.formatDate

// ------- step func
// write data [[row],[row],[row]]
Xmlsx.prototype.entry = function(data) {
  this._store.load(data)
  return this
}

// frozen row TODO: frozen col
Xmlsx.prototype.frozen = function(range) {
  frozen.setFrozen(this._store, range)
  return this
}

// valid data [{A1: [1, 2, 3]}, {'B1:B100': [a, b, c]}]
Xmlsx.prototype.valid = function(validArray) {
  valid.setValid(this._store, validArray)
  return this
}

// hide cell 'A1'
Xmlsx.prototype.hide = function(cell) {
  hide.setHide(this._store, cell)
  return this
}

// ------- final func
// get xlsx buffer
Xmlsx.prototype.done = function(callback) {
  this._store.compile()
  this._store.build(callback)
}

// get xlsx object
Xmlsx.prototype.output = function() {
  if (this._store.opens.frozen) frozen.parseFrozen(this._store)
  if (this._store.opens.valid) valid.parseValid(this._store)
  this._store.parse()
  return this._store.clear()
}

module.exports = Xmlsx
