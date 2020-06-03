## QSheet

- Used to display data, editable input, form controls based on canvas and javascript
- Support excel file, xml, json, print, cell format, multi-table, etc., please refer to api for more
- IE9+



## Directory

- apidoc: *api document index.html*

- src

  - css
    - table.css: *control style file*

  - js
    - apidoc.json: *api document json file*
    - defaultOption.js: *table data format default attribute reference*
    - popup.js: *drop-down box plugin*
    - table.js: *core file of table control*
    - table.min.js: *min version*
    - tableApi.js: *api document js file*
    - tableInfo.js: *demo data*

  - demo
    - example1.html: *example 1* 
    - example2.html: *example 2*

  - vendor
    - FileSaver.js: *third-party files required for file export*
    - jszip.min.js: *compressed file import third-party files required*

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

