# xmlsx
xmlsx = xml + xlsx

## example

    var xmlsx = new XMLSX()

    xmlsx
     .frozen('2') // 冻结某行以上
     .write([ // 输入表格数据
       ['我', '秦始皇', '打钱'],
       ['aa', 'bb']]
      )
     .valid([ // 数据验证
       { A1: [2, 4, 6] },
       { B5: ['f1', 'f4', 'qb'] }
      ])
     .getXlsx(function(err, buf) {
       fs.writeFile('test.xlsx', buf)
      })