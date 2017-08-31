# xmlsx
xmlsx = xml + xlsx

## example

    var xmlsx = new XMLSX()

    xmlsx
     .frozen('2')
     .write([
       ['我', '秦始皇', '打钱'],
       ['aa', 'bb']]
      )
     .valid([
       { A1: [2, 4, 6] },
       { B5: ['f1', 'f4', 'qb'] }
      ])
     .getXlsx(function(err, buf) {
       fs.writeFile('test.xlsx', buf)
      })