# xmlsx
xmlsx = xml + xlsx

## Method

* frozen: frozen a line (冻结某行以上)

      @type step
      @param [Number | String] - rowNumber

* entry: input entry (输入表格数据)

      @type step
      @param [Array[Array]] - inputData

* valid: set valid (数据验证)

      @type step
      @param [Array[Object]] - validData

* hide: hide one cell (隐藏某个cell)

      @type step
      @param [String] - hideCell

* output: get xlsx json (解析xlsx并输出内容)

      @type final
      @return @see entry param

* done: get xlsx buffer (获取设定后的xlsx)

      @type final
      @return [Buffer] xlsxBuffer

* getCell: get a xlsx axis (转换xy坐标为xlsx的位置)

      @type tool
      @param [Number, Number] x,y start on 0
      @return [String] cell eg:'A1'

* formatDate: format Date cell (将日期格式的cell值转换为Date)

      @type tool
      @param [Number|String] dateCellVal
      @return [Date] date


## Example
### Make Xlsx
    var xmlsx = new Xmlsx()

    xmlsx
      .entry([
        ['1', 2, '三'], 
        ['foo', 'bar'],
        [],
        ['', 'space', ],
        ['', '', '@_@']
      ])
      .frozen('5')
      .hide('A2')
      .valid([
        { A1: ['male', 'famale'] }, 
        { B5: ['JS', 'Node', '东北话'] }, 
        { 'C2:C1000': ['甜豆腐脑', '咸豆腐脑'] }
      ])
      .done(function(err, buf) {
        fs.writeFile('test.xlsx', buf)
      })

### Read Xlsx
    fs.readFile('test.xlsx', function(err, buf) {
      var xmlsx = new Xmlsx(buf)
      xmlsx.output()
      //-> { entry: [[...], [...], ...] }

      xmlsx.frozen().output()
      //-> { ..., frozen: '5' }

      xmlsx.valid().frozen().output()
      //-> { ..., valid: [{ A1: [ male, famale ]}] }
    })

### Tools
      Xmlsx.getCell(0, 0)   //-> 'A1'
      Xmlsx.getCell(1, 2)   //-> 'B3'

      Xmlsx.formatDate(42735)  //-> Sun Jan 01 2017 00:00:00 GMT+0800 (CST) (2017/1/1)
