var fs = require('fs')
var path = require('path')

var Zip = require('adm-zip')
var xml2js = require('xml2js')

var async = require('async')
var each = require('lodash').forEach
var map = require('lodash').map

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
}

/**
 * @param {function} cb callback cb(buffer)
 */
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

/**
 * @param {Buffer} buf 
 * @return {object} data
 */
function readXLSXBuffer(buf) {
  var z = new Zip(buf)
  var es = z.getEntries()
  var data = {}

  each(es, function(e) {
    if (e.entryName === 'xl/worksheets/sheet1.xml') {
      xml2js.parseString(e.getData(), function(err, d) {
        data.formats = d
      })
    } else if (e.entryName === 'xl/sharedStrings.xml') {
      xml2js.parseString(e.getData(), function(err, d) {
        data.sheetData = d
      })
    }
  })

  return data
}

// TODO: i need think a better way to use
// function getColStyle(sheetData) {
//   var uzd = unzip(sheetData)
//   var newSheetData = []
//   each(uzd, function(col) {
//     var newCol = []
//     if (col[0] && col.length > 1) {
//       each(col, function(item) {
//         if (item) newCol.push(item)
//       })
//       newSheetData.push(newCol)
//     }
//   })
//   return newSheetData
// }

module.exports = {
  each: each,
  path: path,

  execBuild: execBuild,
  converToXML: converToXML,
  readXLSXBuffer: readXLSXBuffer,
}
