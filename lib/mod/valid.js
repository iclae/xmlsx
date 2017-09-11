var each = require('lodash').forEach

function setValid(store, validArr) {
  var vs = validArr || []
  store.opens.valid = true
  if (vs.length) {
    each(vs, function(v) {
      each(v, function(val, key) {
        store.worksheet.dataValidations[0].dataValidation.push({
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
  }
}

function parseValid(store) {
  var dvs = store.worksheet.dataValidations
  store.sheet.valid = []
  each(dvs, function(dv) {
    each(dv.dataValidation, function(v) {
      var validObj = {}
      var valiStr = v.formula1[0]
      validObj[v.$.sqref] = valiStr.slice(1, valiStr.length - 1).split(',')
      store.sheet.valid.push(validObj)
    })
  })
}

module.exports = {
  setValid: setValid,
  parseValid: parseValid,
}
