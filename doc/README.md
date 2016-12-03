# 变量说明
	metaDataMap：Map类型，key-value数据集
	rptResultList：List类型，报表的数据集，集合中保存的是DTO对象，有序。
	表格批注：文本，添加批注之后，需要把批注中的内容清空再填写批注内容。
==================================================================================
报表模板元素定义的格式为 #{type.[options]} ,
type为元素类型，options部分每种元素不尽相同。在1.2.1版本之后，options会慢慢被批注方式替代。
结合resource中的模板阅读本说明，效果更佳。

# head表头:
#{header.property}，property为metaDataMap中的key，程序会通过property从map中获取value。head元素在写入时的坐标和模板中定义的相同，不会有位移。
可以设置合并单元格，需要设置单元格格式。
批注：
增加：margin right 1，margin left 1，margin center 0，在交叉报表时控制表头表尾的位置。
注意：margin对非center的操作可能会有问题，请测试后使用。

# footer表尾:
#{footer.property}，property为metaDataMap中的key，程序会通过property从map中获取value。footer元素在写入时会根据写入的数据区的行数进行位移，如：模板中一个footer元素定义的坐标为B10，报表数据写入了10行，那么在报表中的footer坐标会位移到B19。
可以设置合并单元格，需要设置单元格格式。
	
#  const常量元素:
常量元素分两种，一种是static，另一种是param。v纵向位移，h横向位移，f不位移。
#{const.v.static.value}，static类型的元素只根据写入的数据行数位移，表格中的值是value部分。
#{const.v.param.property}，param类型的元素和static位移的方式一样，只是值是从map中取。
#{const.f.param.property} 固定位置的常量元素
可以设置合并单元格，需要设置单元格格式。
	
# sum求合:
和const一样，这里也预留了一个占位符v。求合一般用在数据区域的下方。
#{sum.v.name}，元素会根据写入数据行数位移，值就是value。
#{sum.v.name[expressions]},
可以设置合并单元格，需要设置单元格式。
	
# normal正常数据元素:
normal元素是直接按行写入list中的dto属性，一行对应一个dto对象，一个元素对应一个dto属性。
#{normal.property}property是dto的属性名。程序会根据属性名获取dto中的属性值，写入到元素单元格里，当写完数据时程序记录写入数据行数。
#{normal.sbtcode.[merge 2]}，如果当前单元格的内容和右边一个单元格的内容一样，则合并这两个单元格。
批注定义：
“merge v”，纵向合并，自动合并纵向相同的单元格。
group，分组标示，一张报表只能有一个分组标记，配合“subtotal”使用。
特殊的定义：
normal还提供了序号功能，定义方法是#{normal.SN}，当然，SN一定不可以是dto中的属性。
#{normal.SN}，自增序列，从1开始。
# subtotal 小计
和normal搭配使用，切只能与normal配合使用，使用时必须有且只有一个normal标示为group。
小计逻辑：即使只有一条数据也会进行小计。
注意：程序不会对数据进行排序，所以需要在传参是把数据排好序。
有一下类型：
#{subtotal.static.小计}，#{subtotal.static}静态常量，
#{subtotal.calc}计算小计标记，会按group计算小计。
	
# crosstab 交叉数据元素:
crosstab.ctrl.[param] 交叉数据参数，用来控制交叉报表生成的行为，不存在单独的构建步骤。注释：已经去掉此功能，用批注的方式代替。
crosstab.ctrl.row_not_merger,行不合并。
#{crosstab.column.title1}列表头
#{crosstab.row.name}行表头
#{crosstab.data.endamt}数据，暂时只支持一个数据单元
交叉报表可以使用批注增加参数，控制报表展示，如：
表头：order [asc|desc]  控制排序