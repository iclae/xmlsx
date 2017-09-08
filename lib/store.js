var util = require('./util')
var tempData = require('./temp.json')

function Store() {
  util.assign(this, tempData)
  this.opens = {
    frozen: false,
    valid: false,
  }
  this.sheetInfo = {
    dataCount: 0,
    rowCount: 0,
  }
  this.sheet = {
    entry: [],
  }
}

Store.prototype = {
  cache: cache,
  load: load,
  build: build,
  parse: parse,
  clear: clear,
}

// save tmp data
function cache(buffer) {
  if (buffer) {
    util.assign(this, util.readXLSXBuffer(buffer))
  }
}

// load cell datas
function load(datas) {
  var self = this
  var es = datas || []
  if (!es.length) return
  util.each(es, function(vals) {
    var rowCursor = 65
    var rowIndex = self.sheetInfo.rowCount
    var target = self.worksheet.sheetData[0].row[rowIndex]
    self.sheetInfo.rowCount += 1
    if (!target) {
      target = {
        $: { r: String(self.sheetInfo.rowCount) },
        c: [],
      }
      util.each(vals, function(val) {
        self.sst.si.push({ t: val })
        target.c.push({
          $: {
            r: String.fromCharCode(rowCursor++) + self.sheetInfo.rowCount,
            t: 's',
          },
          v: String(self.sheetInfo.dataCount),
        })
        self.sheetInfo.dataCount += 1
      })
    } else {
      util.each(vals, function(val, i) {
        self.sst.si.push({ t: val })
        if (target.c[i]) {
          target.c[i].v = String(self.sheetInfo.dataCount)
        } else {
          target.c[i] = {
            $: {
              r: String.fromCharCode(rowCursor++) + self.sheetInfo.rowCount,
              t: 's',
            },
            v: String(self.sheetInfo.dataCount),
          }
        }
        self.sheetInfo.dataCount += 1
      })
    }
    self.worksheet.sheetData[0].row[rowIndex] = target
    self.sst.$.count = String(self.sheetInfo.dataCount)
    self.sst.$.uniqueCount = String(self.sheetInfo.dataCount)
  })
}

// build xlsx buffer
function build(callback) {
  var WORKSHEET = util.path.join(__dirname, './src/xl/worksheets/sheet1.xml')
  var SST = util.path.join(__dirname, './src/xl/sharedStrings.xml')

  var worksheet = util.converToXML({ worksheet: this.worksheet })
  var sst = util.converToXML({ sst: this.sst })
  var escaped = /&amp;quot;/g
  var buildQueue = []
  worksheet = worksheet.replace(escaped, '&quot;')
  buildQueue.push({ path: WORKSHEET, data: worksheet })
  buildQueue.push({ path: SST, data: sst })

  util.execBuild(buildQueue, callback)
}

// parse xlsx obj to datas
function parse() {
  var self = this
  var ts = this.sst.si || []
  var rowFormat = self.worksheet.sheetData[0].row
  var newTs = []
  var newData = []
  var last = 0
  if (!ts.length) return
  newTs = util.map(ts, 't[0]')

  util.each(rowFormat, function(row) {
    var rowData = []
    var r = row.c
    util.each(r, function(vs) {
      var ri = vs.$.r[0].charCodeAt() - 65
      var t = vs.$.t
      if (t && t === 's') rowData[ri] = newTs[vs.v[0]]
      else {
        if (vs.v && vs.v.length) rowData[ri] = String(parseFloat(vs.v[0]))
        else rowData[ri] = ''
      }
      self.sheetInfo.dataCount += 1
    })
    rowData = util.map(rowData, function(row) {
      return row ? row : ''
    })
    self.sheetInfo.rowCount += 1
    newData.push(rowData)
  })

  // filter void array
  util.each(newData, function(d, index) {
    if (d.length) last = index
  })
  self.sheet.entry = newData.slice(0, last + 1)
}

// clear output xlsx obj
function clear() {
  return this.sheet
}

module.exports = Store
