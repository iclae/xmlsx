function Frozen(range) {
  this.range = range
}

Frozen.prototype.setFrozen = function setFrozen(sheet) {
  var source = getSource()
  if (this.range && sheet) {
    source.sheetViews[0].sheetView[0].pane[0].$.ySplit = String(this.range) + '.0'
    source.sheetViews[0].sheetView[0].pane[0].$.topLeftCell = 'A' + (Number(this.range) + 1)
    source.sheetViews[0].sheetView[0].selection.$.activeCell = 'B' + (Number(this.range) + 2)
    source.sheetViews[0].sheetView[0].selection.$.sqref = 'B' + (Number(this.range) + 2)
    sheet.worksheet.sheetViews = source.sheetViews
  }
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

module.exports = Frozen
