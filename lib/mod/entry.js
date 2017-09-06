var each = require('../util').each

function Entry(data) {
  this.data = data || []
}
function getFormatSource() {
  return {
    sheetData: [{ row: [] }],
  }
}

Entry.prototype.setSheet = function setSheet(store) {
  var es = this.data
  var formatSource
  var count = 0
  if (es.length) {
    formatSource = getFormatSource()
    each(es, function(row, rowIndex) {
      var rs = {
        $: { r: String(rowIndex + 1) },
        c: [],
      }
      var A = 65
      each(row, function(cell) {
        store.sheet.sheetData.sst.si.push({ t: cell })

        rs.c.push({
          $: {
            r: String.fromCharCode(A++) + (rowIndex + 1),
            s: '1',
            t: 's',
          },
          v: String(count++),
        })
      })
      formatSource.sheetData[0].row.push(rs)
    })
    store.sheet.formats.worksheet.sheetData = formatSource.sheetData
    store.sheet.sheetData.sst.$.count = String(count)
    store.sheet.sheetData.sst.$.uniqueCount = String(count)
  }
}

Entry.parseData = function parseData(store) {
  var ts = store.sheet.sheetData.sst.si
  var rowFormat = store.sheet.formats.worksheet.sheetData[0].row
  var newTs = []
  var newData = []
  var last = 0
  // flatten data
  each(ts, function(t) {
    newTs.push(t.t[0])
  })
  // format data
  each(rowFormat, function(row) {
    var rowData = []
    var r = row.c
    each(r, function(vs) {
      var ri = vs.$.r[0].charCodeAt() - 65
      var t = vs.$.t
      if (t && t === 's') rowData[ri] = newTs[vs.v[0]]
      else {
        if (vs.v && vs.v.length) rowData[ri] = String(parseFloat(vs.v[0]))
        else rowData[ri] = ''
      }
    })
    newData.push(rowData)
  })

  // filter void array
  each(newData, function(d, index) {
    if (d.length) last = index
  })
  newData = newData.slice(0, last + 1)
  return newData
}

module.exports = Entry
