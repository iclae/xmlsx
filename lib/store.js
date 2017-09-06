var util = require('./util')

var tempData = require('./temp/data')
var tempFormats = require('./temp/formats')

var parseData = require('./mod/entry').parseData
// var parseFrozen = require('./mod/frozen').parseFrozen
// var parseValid = require('./mod/valid').parseValid

function Store() {
  this.sheet = {
    formats: tempFormats.getFormats(),
    sheetData: tempData.getStringData(),
  }
}

Store.prototype = {
  cache: cache,
  build: build,
  clear: clear,
}

// save tmp data
function cache(buffer) {
  if (buffer) {
    this.sheet = util.readXLSXBuffer(buffer)
  }
}

// build xlsx buffer
function build(callback) {
  var FORMAT = util.path.join(__dirname, './src/xl/worksheets/sheet1.xml')
  var DATA = util.path.join(__dirname, './src/xl/sharedStrings.xml')
  var formatStr = util.converToXML(this.sheet.formats)
  var sheetData = util.converToXML(this.sheet.sheetData)
  var escaped = /&amp;quot;/g
  var buildQueue = []
  formatStr = formatStr.replace(escaped, '&quot;')
  buildQueue.push({ path: FORMAT, data: formatStr })
  buildQueue.push({ path: DATA, data: sheetData })

  util.execBuild(buildQueue, callback)
}

// output parse sheet
function clear() {
  var data = {
    sheetData: parseData(this),
    // TODO: it can work but need a better way to use, so stay comment
    // formats: { valid: parseValid(this), frozen: parseFrozen(this) },
  }
  return data
}

module.exports = Store
