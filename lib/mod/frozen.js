function setFrozen(store, range) {
  var r = ''
  store.opens.frozen = true
  if (range) {
    r = String(range)
    store.worksheet.sheetViews[0].sheetView[0].pane[0].$.ySplit = String(r) + '.0'
    store.worksheet.sheetViews[0].sheetView[0].pane[0].$.topLeftCell = 'A' + (Number(r) + 1)
    store.worksheet.sheetViews[0].sheetView[0].selection[0].$.activeCell = 'B' + (Number(r) + 2)
    store.worksheet.sheetViews[0].sheetView[0].selection[0].$.sqref = 'B' + (Number(r) + 2)
  }
}

function parseFrozen(store) {
  var sheetViews = store.worksheet.sheetViews
  var forzenNumber = parseInt(sheetViews[0].sheetView[0].pane[0].$.ySplit, 10)
  store.sheet.frozen = String(forzenNumber)
}

module.exports = {
  setFrozen: setFrozen,
  parseFrozen: parseFrozen,
}
