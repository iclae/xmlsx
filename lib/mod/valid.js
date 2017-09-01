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

function getSource() {
  return {
    dataValidations: [{ dataValidation: [] }],
  }
}

module.exports = Valid
