## QSheet

* 用于展示数据、可编辑输入、基于canvas与javascript开发的表格控件
* 支持excel文件、xml、json、打印、单元格格式、多表等等功能，更多请参考api
* IE9+



## 目录

* apidoc  api文档index.html

* src

  * css 
    * table.css 控件样式文件
  * js
    * apidoc.json api文档json文件
    * defaultOption.js 表格数据格式默认属性参考
    * popup.js 下拉框插件
    * table.js 表格控件核心文件
    * table.min.js min版本
    * tableApi.js api文档js文件
    * tableInfo.js demo数据
  * demo
    * example1.html  示例1
    * example2.html  示例2

  * vendor

    * FileSaver.js 文件导出需要的第三方文件

    * jszip.min.js 压缩文件导入需要的第三方文件


## 如何使用apidoc

* ```
  npm install apidoc -g
  ```

* 打开apidoc下的index.html



## 如何使用该控件

* 更详细 参考example1.html  example2.html  api文档



```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>example1</title>
    <link rel="stylesheet" href="../css/table.css"> <!--引入css-->
</head>
<body>
    <div id="container"></div>  <!--容器-->


    <script src="../js/table.js" ></script>    <!--引入核心的接口文件-->
    <script type="text/javascript">
    var WB = new Workbook('container');     //容器 data(可选)
        
    // do something...
        
    WB.startPaint(true)	//绘制表格

    </script>
</body>
</html>
```

