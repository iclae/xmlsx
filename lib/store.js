var path = require('path')
var util = require('./util')

var async = require('async')
var each = require('lodash').forEach
var map = require('lodash').map
var assign = require('lodash').assign
var flatten = require('lodash').flatten
var times = require('lodash').times
var find = require('lodash').find
var maxBy = require('lodash').maxBy

var tempData = require('./temp.json')
var styleMapping = require('./style.json')

function Store() {
  assign(this, tempData)
  this.opens = {
    frozen: false,
    valid: false,
  }
  // to output
  this.sheet = {
    entry: [],
  }
  // to compile
  this.preEntry = {
    data: [],
    style: {},
  }
  // cells border
  this.edge = {
    x: 0,
    y: 0,
  }
}

Store.prototype = {
  cache: cache,
  bound: bound,
  load: load,
  style: style,
  compile: compile,
  build: build,
  parse: parse,
  clear: clear,
}

// save tmp data
function cache(buffer) {
  if (buffer) {
    assign(this, util.readXLSXBuffer(buffer))
  }
}

// set max edge
function bound(e) {
  this.edge = {
    x: maxBy([e.x, this.edge.x]),
    y: maxBy([e.y, this.edge.y]),
  }
}

// load cell datas
function load(datas) {
  var self = this
  var es = datas || []
  if (!es.length) return
  self.preEntry.data = flatten(es)
  self.sheet.entry = es
  self.bound(util.getEdgeByData(es))
}

// set styles
function style(s) {
  var cell = s.cell.toUpperCase()
  this.preEntry.style[cell] = s.style
  this.bound(util.getEdgeByCell(cell, false))
}

// pre build
function compile() {
  var self = this
  var rows = []
  var count = 0

  async.parallel({
    worksheet: function() {
      times(self.edge.y, function(y) {
        var row = {
          $: { r: String(y + 1) },
          c: [],
        }
        times(self.edge.x, function(x) {
          var v = self.sheet.entry[y] ? self.sheet.entry[y][x] : null
          var cell = util.getCellByXY(x, y)
          var cellStyle = self.preEntry.style[cell] || 'none'
          var col = {
            $: {
              r: util.getCellByXY(x, y),
              t: 's',
              s: styleMapping[cellStyle],
            },
          }
          if (v || v === '') col.v = String(count++)
          row.c.push(col)
        })
        rows.push(row)
      })
      self.worksheet.sheetData[0].row = rows
    },
    sst: function() {
      var sis = map(self.preEntry.data, function(v) {
        return { t: String(v) }
      })
      self.sst.si = sis
      self.sst.$.count = String(sis.length)
      self.sst.$.uniqueCount = String(sis.length)
    },
  })
}

// build xlsx buffer
function build(callback) {
  var WORKSHEET = path.join(__dirname, './src/xl/worksheets/sheet1.xml')
  var SST = path.join(__dirname, './src/xl/sharedStrings.xml')
  var STYLESHEET = path.join(__dirname, './src/xl/styles.xml')

  var worksheet = util.converToXML({ worksheet: this.worksheet })
  var sst = util.converToXML({ sst: this.sst })
  var stylesheet = util.converToXML({ styleSheet: this.styleSheet })
  var escaped = /&amp;quot;/g
  var buildQueue = []
  worksheet = worksheet.replace(escaped, '&quot;')
  buildQueue.push({ path: WORKSHEET, data: worksheet })
  buildQueue.push({ path: SST, data: sst })
  buildQueue.push({ path: STYLESHEET, data: stylesheet })

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
  newTs = map(ts, 't[0]')

  each(rowFormat, function(row) {
    var rowData = []
    var r = row.c
    each(r, function(vs) {
      var ri = vs.$.r[0].charCodeAt() - 65
      var t = vs.$.t
      var val
      if (!vs.v) return
      val = vs.v[0]
      if (t && t === 's') rowData[ri] = newTs[val]
      else {
        if (vs.v.length) rowData[ri] = String(parseFloat(val))
        else rowData[ri] = ''
      }
    })
    rowData = map(rowData, function(row) {
      return row ? row : ''
    })
    newData.push(rowData)
  })

  // filter void array
  each(newData, function(d, index) {
    if (d.length) last = index
  })
  newData = newData.slice(0, last + 1)
  self.edge = util.getEdgeByData(newData)
  self.sheet.entry = newData
}

// clear output xlsx obj
function clear() {
  return this.sheet
}

module.exports = Store
