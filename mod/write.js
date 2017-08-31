var each = require('../util').each

function Write(sheetData) {
  this.data = sheetData || []
}

Write.prototype.setSheet = function setSheet(sheet) {
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

function getFormatSource() {
  return {
    sheetData: [{ row: [] }],
  }
}

module.exports = Write
