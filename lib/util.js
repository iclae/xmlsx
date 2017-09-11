var fs = require('fs')
var path = require('path')

var Zip = require('adm-zip')
var xml2js = require('xml2js')

var async = require('async')
var map = require('lodash').map

/**
 * @param {Buffer} buf 
 * @return {object} data
 */
function readXLSXBuffer(buf) {
  var z = new Zip(buf)
  var data = {}

  async.parallel(
    {
      worksheet: function(next) {
        var sheetData = z.getEntry('xl/worksheets/sheet1.xml').getData()
        xml2js.parseString(sheetData, next)
      },
      sst: function(next) {
        var data = z.getEntry('xl/sharedStrings.xml').getData()
        xml2js.parseString(data, next)
      },
      // styleSheet: function(next) {
      //   var data = z.getEntry('xl/styles.xml').getData()
      //   xml2js.parseString(data, next)
      // },
    },
    function(err, result) {
      if (err) throw new Error(err)
      data.worksheet = result.worksheet.worksheet
      data.sst = result.sst.sst
      // data.styleSheet = result.styleSheet.styleSheet
    }
  )

  return data
}

/**
 * @param {object} obj xmlObject
 * @return {string} xmlString
 */
function converToXML(obj) {
  var xmlBuilder = new xml2js.Builder()
  return xmlBuilder.buildObject(obj)
}

/**
 * @param {array} queue 
 * @param {function} callback 
 */
function execBuild(queue, callback) {
  // create build func
  var packQueues = map(queue, function(q) {
    return function(next) {
      fs.writeFile(q.path, q.data, next)
    }
  })
  // exec queue
  async.parallel(packQueues, function(err) {
    if (err) throw new Error(err)
    buildXlsx(callback)
  })

  function buildXlsx(cb) {
    var SOURCE = path.join(__dirname, './src')
    var zip = new Zip()
    zip.addLocalFolder(SOURCE)
    zip.toBuffer(
      function(buf) {
        cb(null, buf)
      },
      function(err) {
        cb(err)
      }
    )
  }
}

module.exports = {
  execBuild: execBuild,
  converToXML: converToXML,
  readXLSXBuffer: readXLSXBuffer,
}
