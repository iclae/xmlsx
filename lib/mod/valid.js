var each = require('../util').each

function Valid(validArr) {
  this.data = validArr
}
function getSource() {
  return {
    dataValidations: [{ dataValidation: [] }],
  }
}

Valid.prototype.setSheet = function setSheet(store) {
  var vs = this.data
  var source
  if (vs.length) {
    source = getSource()
    each(vs, function(v) {
      each(v, function(val, key) {
        source.dataValidations[0].dataValidation.push({
          $: {
            type: 'list',
            allowBlank: '1',
            showErrorMessage: '1',
            sqref: key.toUpperCase(),
          },
          formula1: '&quot;' + val.join() + '&quot;',
        })
      })
    })
    store.sheet.formats.worksheet.dataValidations = source.dataValidations
  }
}

Valid.parseValid = function parseValid(store) {
  var dvs = store.sheet.formats.worksheet.dataValidations
  var validArr = []
  each(dvs, function(dv) {
    each(dv.dataValidation, function(v) {
      var validObj = {}
      var valiStr = v.formula1[0]
      validObj[v.$.sqref] = valiStr.slice(1, valiStr.length - 1).split(',')
      validArr.push(validObj)
    })
  })
  return validArr
}

module.exports = Valid
