var Store = require('./lib/store')

// mods
var frozen = require('./lib/mod/frozen')
var valid = require('./lib/mod/valid')

// main
function XMLSX(buffer) {
  this._store = new Store()
  this._store.cache(buffer)
}

// ------- step func
// write data [[row],[row],[row]]
XMLSX.prototype.entry = function(data) {
  this._store.load(data)
  return this
}

// frozen row TODO: frozen col
XMLSX.prototype.frozen = function(range) {
  frozen.setFrozen(this._store, range)
  return this
}

// valid data [{A1: [1, 2, 3]}, {'B1:B100': [a, b, c]}]
XMLSX.prototype.valid = function(validArray) {
  valid.setValid(this._store, validArray)
  return this
}

// ------- final func
// get xlsx buffer
XMLSX.prototype.done = function(callback) {
  this._store.build(callback)
}

// get xlsx object
XMLSX.prototype.output = function() {
  if (this._store.opens.frozen) frozen.parseFrozen(this._store)
  if (this._store.opens.valid) valid.parseValid(this._store)
  this._store.parse()
  return this._store.clear()
}

module.exports = XMLSX
