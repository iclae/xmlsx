function Frozen(range) {
  this.range = range || 0
}
function getSource() {
  return {
    sheetViews: [
      {
        sheetView: [
          {
            $: { workbookViewId: '0' },
            pane: [
              {
                $: {
                  ySplit: '1.0',
                  topLeftCell: 'A1',
                  activePane: 'bottomLeft',
                  state: 'frozen',
                },
              },
            ],
            selection: {
              $: {
                activeCell: 'B2',
                sqref: 'B2',
                pane: 'bottomLeft',
              },
            },
          },
        ],
      },
    ],
  }
}

Frozen.prototype.setFrozen = function setFrozen(store) {
  var source
  if (this.range) {
    source = getSource()
    source.sheetViews[0].sheetView[0].pane[0].$.ySplit = String(this.range) + '.0'
    source.sheetViews[0].sheetView[0].pane[0].$.topLeftCell = 'A' + (Number(this.range) + 1)
    source.sheetViews[0].sheetView[0].selection.$.activeCell = 'B' + (Number(this.range) + 2)
    source.sheetViews[0].sheetView[0].selection.$.sqref = 'B' + (Number(this.range) + 2)
    store.sheet.formats.worksheet.sheetViews = source.sheetViews
  }
}

Frozen.parseFrozen = function parseFrozen(store) {
  var sheetViews = store.sheet.formats.worksheet.sheetViews
  var forzenNumber = 0
  if (sheetViews) {
    forzenNumber = parseInt(store.sheet.formats.worksheet.sheetViews[0].sheetView[0].pane[0].$.ySplit, 10)
  }
  return forzenNumber
}

module.exports = Frozen
