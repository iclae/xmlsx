# xmlsx
xmlsx = xml + xlsx

## Method

* frozen  frozen a line 冻结某行以上
* entry input entry 输入表格数据
* valid set valid 数据验证
* output get xlsx json 解析xlsx并输出内容
* done get xlsx buffer 获取设定后的xlsx

## Example
### Make Xlsx
    var xmlsx = new XMLSX()

    xmlsx
      .entry([
        ['1', 2, '三'], 
        ['foo', 'bar'],
        [],
        ['', 'space', ],
        ['', '', '@_@']
      ])
      .frozen('5'//or 5)
      .valid([
        { A1: [male, famale] }, 
        { B5: ['JS', 'Node', '东北话'] }, 
        { 'C2:C1000': ['甜豆腐脑', '咸豆腐脑'] }
      ])
      .done(function(err, buf) {
        fs.writeFile('test.xlsx', buf)
      })

### Read Xlsx
    fs.readFile('test.xlsx', function(err, buf) {
      var xmlsx = new XMLSX(buf)
      xmlsx.output()
      //-> { entry: [...] }

      xmlsx.frozen().output()
      //-> { ......, frozen: '5' }

      xmlsx.valid().frozen().output()
      //-> { ......, valid: [{ A1: [ male, famale ]}] }
    })