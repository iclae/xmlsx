var each = require('../util').each

function Valid(validArr) {
  this.data = validArr
}

Valid.prototype.setSheet = function setSheet(sheet) {
  var source = getSource()
  if (this.data.length && sheet) {
    each(this.data, function(v) {
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
    sheet.worksheet.dataValidations = source.dataValidations
  }
}

Valid.analysis = function analysis(format) {
  var dvs = format.worksheet.dataValidations
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

function getSource() {
  return {
    dataValidations: [{ dataValidation: [] }],
  }
}

module.exports = Valid
