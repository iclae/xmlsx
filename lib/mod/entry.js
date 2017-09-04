var each = require('../util').each

function Entry(data) {
  this.data = data || []
}

Entry.prototype.setSheet = function setSheet(sheet) {
  var formatSource = getFormatSource()
  var count = 0
  if (this.data && sheet) {
    each(this.data || [], function(row, rowIndex) {
      var rs = {
        $: { r: String(rowIndex + 1) },
        c: [],
      }
      var A = 65
      each(row, function(cell) {
        sheet.sheetData.sst.si.push({ t: cell })

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
    sheet.formats.worksheet.sheetData = formatSource.sheetData
    sheet.sheetData.sst.$.count = String(count)
    sheet.sheetData.sst.$.uniqueCount = String(count)
  }
}

Entry.analysis = function analysis(sheet) {
  var ts = sheet.sheetData.sst.si
  var rowFormat = sheet.formats.worksheet.sheetData[0].row
  var newTs = []
  var formatTs = []
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
      if (vs.v) rowData[ri] = newTs[vs.v[0]]
    })
    each(rowData, function(rd, index) {
      if (!rd) rowData[index] = ''
    })
    formatTs.push(rowData)
  })
  // filter void array
  each(formatTs, function(d, index) {
    if (d.length) last = index
  })
  formatTs = formatTs.slice(0, last + 1)
  return formatTs
}

function getFormatSource() {
  return {
    sheetData: [{ row: [] }],
  }
}

module.exports = Entry
