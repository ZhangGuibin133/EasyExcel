# EasyExcel
忽略Excel格式,只是把Excel文件当作数据源,在POI的基础上,提供一个简单的输入输出实现
row:用一个List<String>实现
Sheet:用一个List<row>实现
WorkBook:用一个表名对应sheet的Map实现
输出: loadSheetS,按内容输出对应的WorkBook的Map
      loadSheet,输出表1的sheet
      loadSheetsWithFieldNames,按指定表头输出相应的记录
输入: saveSheets,将Map内容输出为工作表,如果验证为数字,保存为数字类型,其余按文本保存
      saveSheets,将List<row>按名字保存,默认表名为sheet1
