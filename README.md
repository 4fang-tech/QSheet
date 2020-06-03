## QSheet

- Used to display data, editable input, form controls based on canvas and javascript
- Support excel file, xml, json, print, cell format, multi-table, etc., please refer to api for more
- IE9+



## Directory

- apidoc:&nbsp;*api document index.html*

- src

  - css
    - table.css:&nbsp;*control style file*

  - js
    - apidoc.json:&nbsp;*api document json file*
    - defaultOption.js:&nbsp;*table data format default attribute reference*
    - popup.js:&nbsp;*drop-down box plugin*
    - table.js:&nbsp;*core file of table control*
    - table.min.js:&nbsp;*min version*
    - tableApi.js:&nbsp;*api document js file*
    - tableInfo.js:&nbsp;*demo data*

  - demo
    - example1.html:&nbsp;*example 1* 
    - example2.html:&nbsp;*example 2*

  - vendor
    - FileSaver.js:&nbsp;*third-party files required for file export*
    - jszip.min.js:&nbsp;*compressed file import third-party files required*

## How to use apidoc

- ```
  npm install apidoc -g
  ```

- Open index.html under apidoc



## How to use the control

- For more details, refer to example1.html example2.html api documentation



```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>example1</title>
    <link rel="stylesheet" href="../css/table.css"> <!--Introduce css-->
</head>
<body>
    <div id="container"></div> <!--container-->


    <script src="../js/table.js" ></script> <!--Introduce the core interface file-->
    <script type="text/javascript">
    var WB = new Workbook('container'); //container data (optional)
        
    // do something...
        
    WB.startPaint(true) //Draw a table

    </script>
</body>
</html>
```

