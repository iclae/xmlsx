# xmlsx
xmlsx = xml + xlsx

## Method

* frozen 冻结某行以上
* entry 输入表格数据
* valid 数据验证
* output 解析xlsx并输出内容
* done 获取设定后的xlsx

## Example
### Make Xlsx
    var xmlsx = new XMLSX()

    xmlsx
      .entry([
        ['我', '秦始皇', '打钱'], 
        ['foo', 'bar']
      ])
      .valid([
        { A1: [male, famale] }, 
        { B5: ['JS', 'Node', '东北话'] }, 
        { 'C2:C1000': ['甜豆腐脑', '咸豆腐脑'] }
      ])
      .done(function(err, buf) {
        fs.writeFile('test.xlsx', buf)
      })

### Read Xlsx
    fs.readFile('test.xlsx', function(err, data) {
      var xmlsx = new XMLSX(data)
      console.log(xmlsx.output())
    })