var fs = require('fs')
var path = require('path')
var xml2js = require('xml2js')
var Zip = require('adm-zip')
var async = require('async')
var each = require('lodash.foreach')
var unzip = require('lodash.unzip')

var initFormat = require('./temp/format')
var initData = require('./temp/data')

var xmlBuilder = new xml2js.Builder()
var zip = new Zip()

var FORMAT = path.join(__dirname, './xmls/xl/worksheets/sheet1.xml')
var DATA = path.join(__dirname, './xmls/xl/sharedStrings.xml')
var XMLS = path.join(__dirname, './xmls')

/**
 * @param {object} sets { formats, sheetData }
 * @param {function} cb cb(err, buffer)
 */
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
        fs.writeFile(FORMAT, formatStr, next)
      },
      function(next) {
        fs.writeFile(DATA, sheetData, next)
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
  zip.addLocalFolder(XMLS)
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

/**
 * @param {Buffer} buf
 * @return 
 */
function readXMLbuffer(buf) {
  var rzip
  var entrys
  var data = {
    formats: {},
    sheetData: [],
  }
  if (buf) {
    rzip = new Zip(buf)
    entrys = rzip.getEntries()
    each(entrys, function(e) {
      if (e.entryName === 'xl/worksheets/sheet1.xml') {
        xml2js.parseString(e.getData(), function(err, res) {
          data.formats = res
        })
      } else if (e.entryName === 'xl/sharedStrings.xml') {
        xml2js.parseString(e.getData(), function(err, res) {
          data.sheetData = res
        })
      }
    })
  } else {
    data.formats = initFormat.getFormat()
    data.sheetData = initData.getData()
  }

  return data
}

/**
 * @param {Object} sheetData 
 */
function getColStyle(sheetData) {
  var uzd = unzip(sheetData)
  var newSheetData = []
  each(uzd, function(col) {
    var newCol = []
    if (col[0] && col.length > 1) {
      each(col, function(item) {
        if (item) newCol.push(item)
      })
      newSheetData.push(newCol)
    }
  })
  return newSheetData
}

module.exports = {
  each: each,
  getColStyle: getColStyle,
  closeSet: closeSet,
  readXMLbuffer: readXMLbuffer,
}
