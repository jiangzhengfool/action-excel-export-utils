# Excel 导出工具

![测试结果1](https://gitcode.net/qiutongcunyi/excel-export-utils/-/blob/master/static/%E6%B5%8B%E8%AF%95%E7%BB%93%E6%9E%9C1.png "测试结果1")

![测试结果2](https://gitcode.net/qiutongcunyi/excel-export-utils/-/blob/master/static/%E6%B5%8B%E8%AF%95%E7%BB%93%E6%9E%9C2.png "测试结果2")

## 简介
和其他Excel导出工具不同，这个工具不能凭空生成一个Excel文件，而是必须指定一个Excel模板文件。工具将会数据填充到Excel模板中，并最终导出为新的Excel文件。

## 主要场景

对于一些需要导出复杂样式的Excel报表使用，可以提前设计好复杂样式的Excel模板，比如字体、样式、图片、报表等。再将数据插入到指定的位置，从而完成最终的Excel报表。

## 主要功能

- 支持一个Excel表格有多个Sheet模板。
- 支持一个Sheet目标页有多个Table 数据区域。
- 支持根据一个Sheet模板生成多个Sheet结果页。
- 支持替换占位符变量，变量样式为`${sign}`，支持一个单元格类多个占位符变量。
- 支持设置Table区域参考指定的行的样式。
- 支持设置每个单元格的值的设置处理和单元格的自定义样式的处理。

## 优势

- 可以实现一些非常复杂的Excel表格的导出。
- 通过提供了默认的处理，对于一些简单的表格导出，也提供的较方便的使用。
- 通过参考单元格样式的方法，来解决大数据量表格的单元格样式超过64000个的限制问题。对于绝大多数表格，真实的独立的样式不会很多，基本都是几种样式互相引用。
异常的文本为`java.lang.IllegalStateException: The maximum number of Cell Styles was exceeded. You can define up to 64000 style in a .xlsx Workbook`
