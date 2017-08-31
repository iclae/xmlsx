var xml2js = require('xml2js')
var Zip = require('adm-zip')
var async = require('async')
var each = require('lodash.foreach')
var fs = require('fs')
var path = require('path')

var xmlBuilder = new xml2js.Builder()
var zip = new Zip()

function closeSet(sets, cb) {
  var formats = sets.formats
  var data = sets.sheetData
  var formatStr = buildObjToXML(formats)
  var sheetData = buildObjToXML(data)
  var escaped = /&amp;quot;/g
  formatStr = formatStr.replace(escaped, '&quot;')
  async.parallel(
    [
      function(next) {
        fs.writeFile('node_modules/xmlsx/xmls/xl/worksheets/sheet1.xml', formatStr, next)
      },
      function(next) {
        fs.writeFile('node_modules/xmlsx/xmls/xl/sharedStrings.xml', sheetData, next)
      },
    ],
    function(err) {
      if (err) throw new Error(err)
      buildXlsx(cb)
    }
  )
}

/**
 * @param {function} cb callback cb(buffer)
 */
function buildXlsx(cb) {
  zip.addLocalFolder('node_modules/xmlsx/xmls')
  zip.toBuffer(
    function(buf) {
      cb(null, buf)
    },
    function(err) {
      cb(err)
    }
  )
}

/**
 * @param {object} obj xmlObject
 * @return {string} xmlString
 */
function buildObjToXML(obj) {
  return xmlBuilder.buildObject(obj)
}

module.exports = {
  closeSet: closeSet,
  each: each,
}
