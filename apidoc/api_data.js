define({ "api": [
  {
    "type": "null",
    "url": "/null",
    "title": "activeCol",
    "name": "activeCol",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the current column of the current table</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.activeCol,activeSheet.activeCol</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "activeRow",
    "name": "activeRow",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the current row of the current table</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.activeRow,activeRow.activeRow</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "allowMoveRange",
    "name": "allowMoveRange",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether the current table allows dragging cells(Defaults:true)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.allowMoveRange = value(boolean),activeSheet.allowMoveRang = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "alwaysShowButton",
    "name": "alwaysShowButton",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to always display the button of the current table cell(Defaults:true)</p> <ul> <li>In case of false value, btn will be displayed only when editing is entered</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>Set btn reference setCellType</li> <li>WB.activeSheet.alwaysShowButton = value(boolean),activeSheet.alwaysShowButton = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "alwaysShowInArea",
    "name": "alwaysShowInArea",
    "group": "Attribute_activeSheet",
    "description": "<p>If the current table cell btn is always displayed, is it displayed only after clicking in the cell(Defaults:true)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.alwaysShowInArea = value(boolean),activeSheet.alwaysShowInArea = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "canOpenLayer",
    "name": "canOpenLayer",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to click the current click cell can trigger the corresponding event(Defaults:true)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.canOpenLayer = value(boolean),activeSheet.canOpenLayer = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "cellPadding.bottom",
    "name": "cellPadding_bottom",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the bottom padding to get the cell content of the current table(Defaults:2)</p> <ul> <li>You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.cellPadding.bottom = value(Number),activeSheet.cellPadding.bottom = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "cellPadding.left",
    "name": "cellPadding_left",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the left margin to get the cell contents of the current table(Defaults:2)</p> <ul> <li>You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.cellPadding.left = value(Number),activeSheet.cellPadding.left = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "cellPadding.right",
    "name": "cellPadding_right",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the right padding to get the cell content of the current table(Defaults:2)</p> <ul> <li>You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.cellPadding.right = value(Number),activeSheet.cellPadding.right = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "cellPadding.top",
    "name": "cellPadding_top",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the top padding to get the cell content of the current table(Defaults:2)</p> <ul> <li>You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.cellPadding.top = value(Number),activeSheet.cellPadding.top = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "colHeaderData.defaultDataNode.style.fillColor",
    "name": "colHeaderData_defaultDataNode_style_fillColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The fill color of the current column header(Defaults:#008844)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.colHeaderData.defaultDataNode.style.fillColor = value(color)，activeSheet.colHeaderData.defaultDataNode.style.fillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "colHeaderData.defaultDataNode.style.fontColor",
    "name": "colHeaderData_defaultDataNode_style_fontColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The font color of the column header of the current table(Defaults:#fff)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.colHeaderData.defaultDataNode.style.fontColor = value(color)，activeSheet.colHeaderData.defaultDataNode.style.fontColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "colHeaderData.defaultH",
    "name": "colHeaderData_defaultH",
    "group": "Attribute_activeSheet",
    "description": "<p>The height of the column header of the current table(Defaults:30)</p> <ul> <li>Can be set using WB.hdrHeight(height, index) method</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.colHeaderData.defaultH = value(Number)，activeSheet.colHeaderData.defaultH = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "colHeaderData.showColHeading",
    "name": "colHeaderData_showColHeading",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to display the column header of the current table(Defaults:false)</p> <ul> <li>Can be set using WB.showColHeading(boolean, index) method</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.colHeaderData.showColHeading = value(boolean)，activeSheet.colHeaderData.showColHeading = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "defaults.colWidth",
    "name": "defaults_colWidth",
    "group": "Attribute_activeSheet",
    "description": "<p>The default column width of the current table(Defaults:60)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.defaults.colWidth = value(number)，activeSheet.defaults.colWidth = value(number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "defaults.rowHeight",
    "name": "defaults_rowHeight",
    "group": "Attribute_activeSheet",
    "description": "<p>The default row height of the current table(Defaults:30)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.defaults.rowHeight = value(number) ，activeSheet.defaults.rowHeight = value(number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "endCol",
    "name": "endCol",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the end column of the visible area of the current table</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.endCol,activeSheet.endCol</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "endRow",
    "name": "endRow",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the end row of the visible area of the current table</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.endRow,activeSheet.endRow</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "gridLinesColor",
    "name": "gridLinesColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The grid line color of the current table(Defaults:ccc)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.gridLinesColor = value(color),activeSheet.gridLinesColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "headeNotLineInFix",
    "name": "headeNotLineInFix",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to open the head mode of the current table(No grid lines above the frozen row)(Defaults:false)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.headeNotLineInFix = value(boolean),activeSheet.headeNotLineInFix = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "isSelectionHideBorder",
    "name": "isSelectionHideBorder",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to hide the check box of the selected cell of the current table(Defaults:false)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.isSelectionHideBorder = value(boolean),activeSheet.isSelectionHideBorder = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "notColSelectionNum",
    "name": "notColSelectionNum",
    "group": "Attribute_activeSheet",
    "description": "<p>Which column of the current table has no check box(Defaults:-1 All have check boxes)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.notColSelectionNum = value(Int),activeSheet.notColSelectionNum = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "notRowSelectionNum",
    "name": "notRowSelectionNum",
    "group": "Attribute_activeSheet",
    "description": "<p>Which row of the current table has no check box(Defaults:-1 All have check boxes)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.notRowSelectionNum = value(Int),activeSheet.notRowSelectionNum = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.footpagestyle",
    "name": "printSetting_footpagestyle",
    "group": "Attribute_activeSheet",
    "description": "<p>Print page number format of current table  (&amp;p/&amp;P) (Defaults:第 &amp;p 页，共 &amp;P 页)</p> <ul> <li>&amp;p &amp;P fixed(These two symbols will be replaced with the corresponding page and the total number of pages)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.footpagestyle = value(string),activeSheet.printSetting.footpagestyle = value(string)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.isshowfooterpageinfo",
    "name": "printSetting_isshowfooterpageinfo",
    "group": "Attribute_activeSheet",
    "description": "<p>Page number display information</p> <ul> <li>0:Do not show(Defaults)</li> <li>1:Footer centered</li> <li>2:Footer left</li> <li>3:Footer right</li> <li>4:Header centered</li> <li>5:Header left</li> <li>6:Header right</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.isshowfooterpageinfo = value(Int),activeSheet.printSetting.isshowfooterpageinfo = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.marginBottom",
    "name": "printSetting_marginBottom",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the bottom margin when printing the current table</p> <ul> <li>The unit is mm(Defaults 5mm)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.marginBottom = value(Number),activeSheet.printSetting.marginBottom = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.marginLeft",
    "name": "printSetting_marginLeft",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the left margin when printing the current table</p> <ul> <li>The unit is mm(Defaults 5mm)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.marginLeft = value(Number),activeSheet.printSetting.marginLeft = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.marginRight",
    "name": "printSetting_marginRight",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the right margin when printing the current table</p> <ul> <li>The unit is mm(Defaults 5mm)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.marginRight = value(Number),activeSheet.printSetting.marginRight = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.marginTop",
    "name": "printSetting_marginTop",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the top margin when printing the current table</p> <ul> <li>The unit is mm(Defaults 5mm)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.marginTop = value(Number),activeSheet.printSetting.marginTop = value(Number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.orientation",
    "name": "printSetting_orientation",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the paper orientation when the current table is printed</p> <ul> <li>1:Vertical(Defaults)</li> <li>2:Horizontal</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.orientation = value(int),activeSheet.printSetting.orientation = value(int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.paper",
    "name": "printSetting_paper",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the paper type for the current table printing(A3-A7)</p> <ul> <li>A4(Defaults)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.paper = value(String),activeSheet.printSetting.paper = value(String)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printArea.c1",
    "name": "printSetting_printArea_c1",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the column of the starting cell of the printing area of the current table page</p> <ul> <li>''Not set(Defaults)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printArea.c1 = value(Int),activeSheet.printSetting.printArea.c1 = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printArea.c2",
    "name": "printSetting_printArea_c2",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the column of the end cell of the printing area of the current table page</p> <ul> <li>''Not set(Defaults)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printArea.c2 = value(Int),activeSheet.printSetting.printArea.c2 = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printArea.r1",
    "name": "printSetting_printArea_r1",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the row of the starting cell of the printing area of the current table page</p> <ul> <li>''Not set(Defaults)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printArea.r1 = value(Int),activeSheet.printSetting.printArea.r1 = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printArea.r2",
    "name": "printSetting_printArea_r2",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the row of the end cell of the printing area of the current table page</p> <ul> <li>''Not set(Defaults)</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printArea.r2 = value(Int),activeSheet.printSetting.printArea.r2 = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printDirection",
    "name": "printSetting_printDirection",
    "group": "Attribute_activeSheet",
    "description": "<p>Set the printing direction of the current table</p> <ul> <li>1:N-shaped(Defaults)</li> <li>2:Z-shaped</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printDirection = value(int),activeSheet.printSetting.printDirection = value(int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printGridLine",
    "name": "printSetting_printGridLine",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to print the grid lines of the current table (Defaults:false)</p> <ul> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printGridLine = value(boolean),activeSheet.printSetting.printGridLine = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printSetting.printHeadings",
    "name": "printSetting_printHeadings",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to print the row and column labels of the current table(Defaults:false)</p> <ul> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.printSetting.printHeadings = value(boolean),activeSheet.printSetting.printHeadings = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rangeCol1",
    "name": "rangeCol1",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the start column of the selected range of the current table(==activeCol)</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rangeCol1,activeRow.rangeCol1</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rangeCol2",
    "name": "rangeCol2",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the end column of the selected range of the current table</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rangeCol2,activeSheet.rangeCol2</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rangeRow1",
    "name": "rangeRow1",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the start row of the selected range of the current table(==activeRow)</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rangeRow1,activeSheet.rangeRow1</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rangeRow2",
    "name": "rangeRow2",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the end row of the selected range of the current table</p> <ul> <li>(r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rangeRow2,activeSheet.rangeCol2</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rowHeaderData.defaultDataNode.style.fillColor",
    "name": "rowHeaderData_defaultDataNode_style_fillColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The fill color of the header of the current table(Defaults:#008844)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rowHeaderData.defaultDataNode.style.fillColor = value(color),activeSheet.rowHeaderData.defaultDataNode.style.fillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rowHeaderData.defaultDataNode.style.fontColor",
    "name": "rowHeaderData_defaultDataNode_style_fontColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The font color of the header of the current table(Defaults:#fff)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rowHeaderData.defaultDataNode.style.fontColor = value(color),activeSheet.rowHeaderData.defaultDataNode.style.fontColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rowHeaderData.defaultW",
    "name": "rowHeaderData_defaultW",
    "group": "Attribute_activeSheet",
    "description": "<p>The width of the current table header(Defaults:30)</p> <ul> <li>Can be set using WB.hdrWidth(width, index) method</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rowHeaderData.defaultW = value(number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rowHeaderData.showRowHeading",
    "name": "rowHeaderData_showRowHeading",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to display the header of the current table(Defaults:false)</p> <ul> <li>Can be set using WB.showRowHeading(boolean, index) method</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.rowHeaderData.showRowHeading = value(boolean),activeSheet.rowHeaderData.showRowHeading = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "selHdrTopLeft",
    "name": "selHdrTopLeft",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether the current table has selected all in the upper left corner(Defaults:false)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.selHdrTopLeft = value(boolean),activeSheet.selHdrTopLeft = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "selectMode",
    "name": "selectMode",
    "group": "Attribute_activeSheet",
    "description": "<p>The selected mode of the current table cell</p> <ul> <li>0:Freely select range(Defaults)</li> <li>1:row mode,click the cell to select the row</li> <li>2:Column mode,click the cell to select the column</li> <li>3:Cell mode can not be dragged</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.selectMode = value(0|1|2|3),activeSheet.selectMode = value(0|1|2|3)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "selectOption.selectBorderColor",
    "name": "selectOption_selectBorderColor",
    "group": "Attribute_activeSheet",
    "description": "<p>Checkbox color of current table(Defaults:green)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.selectOption.selectBorderColor  = value(color),activeSheet.selectOption.selectBorderColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "selectOption.selectFillColor",
    "name": "selectOption_selectFillColor",
    "group": "Attribute_activeSheet",
    "description": "<p>The fill color of the check box of the current table(Defaults:rgba(0,0,245,0.1))</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.selectOption.selectFillColor  = value(color),activeSheet.selectOption.selectFillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "sheetName",
    "name": "sheetName",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the worksheet name of the current table</p> <ul> <li>To set the worksheet name, please use the WB.sheetName(Index, Name) method. WB is a new example</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.sheetName,activeSheet.sheetName</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showFixedLine",
    "name": "showFixedLine",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to display the frozen line of the current table(Defaults:false)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.showFixedLine = value(boolean) ,activeSheet.showFixedLine = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showGridLines",
    "name": "showGridLines",
    "group": "Attribute_activeSheet",
    "description": "<p>Whether to display the grid lines of the current table(Defaults:true)</p> <ul> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.showGridLines = value(boolean),activeSheet.showGridLines = valueboolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "startCol",
    "name": "startCol",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the start column of the visible area of the current table</p> <ul> <li>Count from the next column of the frozen column(If there are two columns frozen in front, it will be listed as the third column)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.startCol,activeSheet.startCol</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "startRow",
    "name": "startRow",
    "group": "Attribute_activeSheet",
    "description": "<p>Get the start row of the visible area of the current table</p> <ul> <li>Count from the next line of the frozen line(If there are two lines frozen in front, it will be listed as the third line)</li> <li>Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting</li> <li>WB.activeSheet.startRow,activeSheet.startRow</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_activeSheet"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "bottomBtn",
    "name": "bottomBtn",
    "group": "Attribute_prototype",
    "description": "<p>Show btn button group when showtabs display mode is set to 2</p> <ul> <li>WB.bottomBtn=['yes','cancel']</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "btnBorderColor",
    "name": "btnBorderColor",
    "group": "Attribute_prototype",
    "description": "<p>The border color of the cell button (celltype ==1 celltype==2)(Defaults: #C5C1AA)</p> <ul> <li>WB.btnBorderColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "btnBorderColorH",
    "name": "btnBorderColorH",
    "group": "Attribute_prototype",
    "description": "<p>The border color of the cell button (hover) (celltype ==1 celltype==2) (Defaults: rgba(30,144,255,.4))</p> <ul> <li>WB.btnBorderColorH = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "btnFillColor",
    "name": "btnFillColor",
    "group": "Attribute_prototype",
    "description": "<p>The fill color of the cell button (celltype ==1 celltype==2) (Defaults: #F0F0F0)</p> <ul> <li>WB.btnFillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "btnFillColorH",
    "name": "btnFillColorH",
    "group": "Attribute_prototype",
    "description": "<p>The fill color of the cell button (hover) (celltype==1 celltype==2) (Defaults: rgba(240,248,255))</p> <ul> <li>WB.btnFillColorH = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "checkBoxBorderColor",
    "name": "checkBoxBorderColor",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the border color of the check box button(Defaults:#C5C1AA)</p> <ul> <li>WB.checkBoxBorderColor=&quot;color&quot;</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "checkBoxBorderColorH",
    "name": "checkBoxBorderColorH",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the border color of the check box button (hover)(Defaults:rgba(30,144,255,.4))</p> <ul> <li>WB.checkBoxBorderColorH=&quot;color&quot;</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "checkBoxFillColor",
    "name": "checkBoxFillColor",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the fill color of the check box button(Defaults:#F0F0F0)</p> <ul> <li>WB.checkBoxFillColor=&quot;color&quot;</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "checkBoxFillColorH",
    "name": "checkBoxFillColorH",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the fill color of the checkbox button (hover)(Defaults:rgba(240,248,255))</p> <ul> <li>WB.checkBoxFillColorH=&quot;color&quot;</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "defaultCellStyle",
    "name": "defaultCellStyle",
    "group": "Attribute_prototype",
    "description": "<p>Get or set the default property value of the cell. The value set is an object. The default value is used without changing a certain property. The default values are as follows</p> <ul> <li>&quot;fontColor&quot;:&quot;#000&quot;</li> <li>&quot;hAlign&quot;:1</li> <li>&quot;cellType&quot;:{&quot;name&quot;:0}</li> <li>&quot;lock&quot;:false</li> <li>&quot;vAlign&quot;:2</li> <li>&quot;formatter&quot;:&quot;&quot;</li> <li>&quot;font&quot;:this.workbook.defaultFontSize+&quot;px/&quot;+this.workbook.defaultLineHeight+&quot;px &quot;+ this.workbook.defaultFontName</li> <li>&quot;wordWrap&quot;:false</li> <li>WB.defaultCellStyle={&quot;fontColor&quot;:&quot;red&quot;}  Set the default font color of the cell to red</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "dpi",
    "name": "dpi",
    "group": "Attribute_prototype",
    "description": "<p>Set or get dpi (Defaults 96)</p> <ul> <li>WB.dpi = 值(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "editSelEnd",
    "name": "editSelEnd",
    "group": "Attribute_prototype",
    "description": "<p>Set the end position of the selected text when triggering input(Defaults:0)</p> <ul> <li>WB.editSelEnd=3   End of the third word</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "editSelStart",
    "name": "editSelStart",
    "group": "Attribute_prototype",
    "description": "<p>Set the start position of the selected text when triggering input(Defaults:0)</p> <ul> <li>WB.editSelStart=1   Start with the second word</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "enterToRight",
    "name": "enterToRight",
    "group": "Attribute_prototype",
    "description": "<p>Whether the cell moves to the right after pressing the Enter key (Defaults: false), when the property is changed to true, the carriage return moves the cell to the last column and jumps to the next row</p> <ul> <li>WB.enterToRight = value(bool)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "isEdit",
    "name": "isEdit",
    "group": "Attribute_prototype",
    "description": "<p>Set whether table initialization enters edit state(Defaults:false)</p> <ul> <li>WB.isEdit=true  table initialization into edit</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "isFocus",
    "name": "isFocus",
    "group": "Attribute_prototype",
    "description": "<p>Can judge whether it is in editing according to isFocus or isEdit(Defaults:false)</p>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "newSheetCol",
    "name": "newSheetCol",
    "group": "Attribute_prototype",
    "description": "<p>Set the number of columns when creating a new table(Defaults: 8)</p> <ul> <li>WB.newSheetCol = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "newSheetRow",
    "name": "newSheetRow",
    "group": "Attribute_prototype",
    "description": "<p>Set the number of rows when creating a new table(Defaults: 8)</p> <ul> <li>WB.newSheetRow = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "scrollSize",
    "name": "scrollSize",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the scroll bar size(Defaults:7)</p> <ul> <li>WB.scrollSize=8  The scroll bar size is 8px</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabArrowHeight",
    "name": "tabArrowHeight",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the arrow height of the tabs switch table(Defaults:8)</p> <ul> <li>WB.tabArrowHeight=8  Tabs bar switch table arrow height is 8px</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabArrowStartX",
    "name": "tabArrowStartX",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the position of the tabs column switch table arrow (the first arrow) from the left side of the tabs column(Defaults:15)</p> <ul> <li>WB.tabArrowStartX=15  The position of the arrow of the tabs bar switching table is 15px from the left side of the tabs bar</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabArrowWidth",
    "name": "tabArrowWidth",
    "group": "Attribute_prototype",
    "description": "<p>Set or get the arrow width of the tabs switch table(Defaults:5)</p> <ul> <li>WB.tabArrowWidth=5  Tabs bar switch table arrow width is 5px</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsHeight",
    "name": "tabsHeight",
    "group": "Attribute_prototype",
    "description": "<p>Set the height of the tabs bar</p> <ul> <li>WB.tabsHeight=30  The height of the tabs bar is 30px</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Attribute_prototype"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "adaption",
    "name": "adaption",
    "group": "Attribute_workbook",
    "description": "<p>Whether to enable adaptive(Defaults:false)</p> <ul> <li>WB.workbook.adaption = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "allowDelete",
    "name": "allowDelete",
    "group": "Attribute_workbook",
    "description": "<p>Whether to allow the delete key to delete the selected text(Defaults:true)</p> <ul> <li>WB.workbook.allowDelete = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "allowResize",
    "name": "allowResize",
    "group": "Attribute_workbook",
    "description": "<p>Whether to adjust the column width and height by dragging(Defaults:true)</p> <ul> <li>WB.workbook.allowResize = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "allowTabs",
    "name": "allowTabs",
    "group": "Attribute_workbook",
    "description": "<p>Whether to use tab key to switch cells(Defaults:true)</p> <ul> <li>WB.workbook.allowTabs = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "behaviorMode",
    "name": "behaviorMode",
    "group": "Attribute_workbook",
    "description": "<p>Workbook mode</p> <ul> <li>1:sheet mode (Double click to edit)(Defaults)</li> <li>2:grid mode (Click cell to enter edit box but no check box)</li> <li>3:sheet click mode (Click to enter edit and check box)</li> <li>WB.workbook.behaviorMode = value</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "devScreenHeight",
    "name": "devScreenHeight",
    "group": "Attribute_workbook",
    "description": "<p>Vertical computer resolution at the time of development</p> <ul> <li>Whether to enable adaptation based on the adaption property, refer to the WB.adaption(bool, devW, devH) method</li> <li>WB.workbook.devScreenHeight = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "devScreenWidth",
    "name": "devScreenWidth",
    "group": "Attribute_workbook",
    "description": "<p>Horizontal computer resolution at the time of development</p> <ul> <li>Whether to enable adaptation based on the adaption property, refer to the WB.adaption(bool, devW, devH) method</li> <li>WB.workbook.devScreenWidth = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "height",
    "name": "height",
    "group": "Attribute_workbook",
    "description": "<p>Workbook height(Defaults:Viewport height minus system scroll bar size)</p> <ul> <li>WB.workbook.height = value(number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "marginCopies",
    "name": "marginCopies",
    "group": "Attribute_workbook",
    "description": "<p>Print the data of two tables on the same sheet of paper The interval between the two data (tables)</p> <ul> <li>This attribute only takes effect when print is equal to 2 and printOnSamePaper is equal to true</li> <li>The unit is mm</li> <li>WB.workbook.marginCopies = value(int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "print",
    "name": "print",
    "group": "Attribute_workbook",
    "description": "<p>Print mode</p> <ul> <li>1:Print current worksheet(Defaults)</li> <li>2:Print workbook</li> <li>3:Print the currently selected area of the current table</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>WB.workbook.print = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printOnSamePaper",
    "name": "printOnSamePaper",
    "group": "Attribute_workbook",
    "description": "<p>Whether multiple tables are printed on the same paper(Defaults:false)</p> <ul> <li>Only effective when print is 2</li> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>WB.workbook.printOnSamePaper = value(Boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "printer",
    "name": "printer",
    "group": "Attribute_workbook",
    "description": "<p>printer</p> <ul> <li>Method WB.setPrint(obj,Index) can set print settings</li> <li>WB.workbook.printer = value(string)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "rowArrowColor",
    "name": "rowArrowColor",
    "group": "Attribute_workbook",
    "description": "<p>The fill color of the arrow of the current line(Defaults:#000)</p> <ul> <li>WB.workbook.rowArrowColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "scrollMode",
    "name": "scrollMode",
    "group": "Attribute_workbook",
    "description": "<p>Scroll bar mode</p> <ul> <li>1:The scroll bar starts from the top or left</li> <li>2:The scroll bar is below or to the right of the freeze line(Defaults)</li> <li>WB.workbook.scrollMode = value</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showContextMenu",
    "name": "showContextMenu",
    "group": "Attribute_workbook",
    "description": "<p>Whether to display the right-click menu bar of the workbook(Defaults:false)</p> <ul> <li>WB.workbook.showContextMenu = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showHScrollBar",
    "name": "showHScrollBar",
    "group": "Attribute_workbook",
    "description": "<p>How to display the horizontal scroll bar of the workbook(Defaults:1)</p> <ul> <li>0:Do not show</li> <li>1:Displayed when the workbook is wide(Defaults)</li> <li>WB.workbook.showHScrollBar = value</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showRowArrow",
    "name": "showRowArrow",
    "group": "Attribute_workbook",
    "description": "<p>Whether to display the leftmost arrow when the current row is selected(Defaults:false)</p> <ul> <li>WB.workbook.showRowArrow = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showTabs",
    "name": "showTabs",
    "group": "Attribute_workbook",
    "description": "<p>Whether to show the tabs column</p> <ul> <li>0:Do not show(Defaults)</li> <li>1:show</li> <li>2:Button group</li> <li>WB.workbook.showTabs = value(Int)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showVScrollBar",
    "name": "showVScrollBar",
    "group": "Attribute_workbook",
    "description": "<p>How to display the vertical scroll bar of the workbook</p> <ul> <li>0:Do not show</li> <li>1:Displayed when the workbook is high(Defaults)</li> <li>WB.workbook.showVScrollBar = value</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showWorkBookBorder",
    "name": "showWorkBookBorder",
    "group": "Attribute_workbook",
    "description": "<p>Whether to display the border of the workbook(Defaults:true)</p> <ul> <li>WB.workbook.showWorkBookBorder = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.borderColor",
    "name": "tabsOptions_borderColor",
    "group": "Attribute_workbook",
    "description": "<p>The border color of each table in the tab bar(Defaults:#DCDCDC)</p> <ul> <li>WB.workbook.tabsOptions.borderColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.fillColor",
    "name": "tabsOptions_fillColor",
    "group": "Attribute_workbook",
    "description": "<p>Tabs bar background fill color(Defaults:#FDF5E6)</p> <ul> <li>WB.workbook.tabsOptions.fillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.font",
    "name": "tabsOptions_font",
    "group": "Attribute_workbook",
    "description": "<p>Tab bar font(Defaults:14px/20px 宋体)</p> <ul> <li>WB.workbook.tabsOptions.font = value(Font size/line height font type)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.fontColor",
    "name": "tabsOptions_fontColor",
    "group": "Attribute_workbook",
    "description": "<p>Tab bar font color(Defaults:#444)</p> <ul> <li>WB.workbook.tabsOptions.fontColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.fontSelColor",
    "name": "tabsOptions_fontSelColor",
    "group": "Attribute_workbook",
    "description": "<p>Font color after table selection(Defaults:#CD853F)</p> <ul> <li>WB.workbook.tabsOptions.fontSelColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.selFillColor",
    "name": "tabsOptions_selFillColor",
    "group": "Attribute_workbook",
    "description": "<p>Tab fill color when the current table is selected(Defaults:#fff)</p> <ul> <li>WB.workbook.tabsOptions.selFillColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "tabsOptions.showAdd",
    "name": "tabsOptions_showAdd",
    "group": "Attribute_workbook",
    "description": "<p>Whether to display the tabs column to add a table + sign(Defaults:true)</p> <ul> <li>WB.workbook.tabsOptions.showAdd = value(boolean)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "width",
    "name": "width",
    "group": "Attribute_workbook",
    "description": "<p>Workbook width(Defaults:Viewport width minus system scroll bar size)</p> <ul> <li>WB.workbook.width = value(number)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "workBookBorderColor",
    "name": "workBookBorderColor",
    "group": "Attribute_workbook",
    "description": "<p>Set the border color of the workbook(Defaults:#ccc)</p> <ul> <li>WB.workbook.workBookBorderColor = value(color)</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/tableApi.js",
    "groupTitle": "Attribute_workbook"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onafteraddrow",
    "name": "onafteraddrow",
    "group": "Event",
    "description": "<p>Enter the event after adding a new line</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onafteraddrow=function(r,c){\n    //do something...\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "oncellbuttonclick",
    "name": "oncellbuttonclick",
    "group": "Event",
    "description": "<p>Click the cell button to trigger the event</p> <ul> <li>The function returns the table index and the current row and current column</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.oncellbuttonclick=function(idx,r,c){\n    alert(\"oncellbuttonclick\"+idx+\",\"+r+\",\"+c)\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onclickcellopenlayer",
    "name": "onclickcellopenlayer",
    "group": "Event",
    "description": "<p>Click on the cell to trigger the event</p> <ul> <li>The function returns the table index and the current row and current column</li> </ul>",
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onclickcellopenlayer=function(idx,r,c){\n    alert(\"onclickcellopenlayer\"+idx+\",\"+r+\",\"+c)\n}",
        "type": "javascript"
      }
    ],
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onenteraddrow",
    "name": "onenteraddrow",
    "group": "Event",
    "description": "<p>Press Enter on the last line to add a line</p> <ul> <li>When the onenteraddrow function returns true, add a row</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onenteraddrow=function(r,c,v){\n    return true\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onequalsign",
    "name": "onequalsign",
    "group": "Event",
    "description": "<p>Input box input = number trigger event</p> <ul> <li>The function returns the current row and current column</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onequalsign=function(r,c){\n    // do something...\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onformularesult",
    "name": "onformularesult",
    "group": "Event",
    "description": "<p>The event that returns the result of the formula</p> <ul> <li>Function return (formula r c)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onformularesult = function(f,r,c){\n    //do something...\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onopenmorelist",
    "name": "onopenmorelist",
    "group": "Event",
    "description": "<p>Refer to which item is triggered by onopenpopup, the function needs to be placed in onopenpopup success and called</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onopenpopup",
    "name": "onopenpopup",
    "group": "Event",
    "description": "<p>Clicking on the cell with celltype==4 will open the drop-down window</p> <ul> <li>Need to import popup.js file</li> <li>Returns 10 values(w,h,x,y,list,r,c,child,fontSize,noSetIndex)</li> <li>noSetIndex(Special item, you can judge that the corresponding event occurs on the item)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onopenpopup = function(w,h,x,y,list,r,c,child,fontSize,noSetIndex){\n    var data =[];\n    if(list){\n        for(var i =0;i<list.length;i++){\n            data.push({name:list[i]})\n        }\n    };\n    WB.popup._init({\n        shade: false,\n        area:[w+'px',list.length*h+\"px\"],\n        offset:[x+'px', y+'px'],\n        backgroundColor: '#dccccc',\n        rowHeight:34+\"px\",\n        data:data,\n        fontSize:fontSize,\n        success: function(index,value,listIndex) {\n            if(listIndex==noSetIndex){\n                WB.onopenmorelist = function(){\n                    // alert(\"something\")\n                };\n                if(typeof(WB.onopenmorelist)=='function'){\n                    WB.onopenmorelist()\n                }\n            }else{\n                WB.setValue(r,c,value)\n                child.innerHTML = value\n                WB.startPaint(true)\n            }\n            this.close(index); \n        },\n    })\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onrowmodedblclick",
    "name": "onrowmodedblclick",
    "group": "Event",
    "description": "<p>Trigger event in double-click mode selection</p> <ul> <li>The function returns the table index and the current row and current column</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onrowmodedblclick=function(idx,r,c){\n    //do something...\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onshoweditor",
    "name": "onshoweditor",
    "group": "Event",
    "description": "<p>Select text according to the set editSelStart (start position of the cursor) and editSelEnd (end position of the cursor)</p> <ul> <li>If there is a callback but these two parameters are not set, all text is selected by default</li> <li>If there is no callback, the cursor is at the start position by default</li> <li>child is a text edit box</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onshoweditor=function(child){\n   var text = child.innerHTML;\n   WB.editSelStart = text.length;\n   WB.editSelEnd = text.length;\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "onvaluechange",
    "name": "onvaluechange",
    "group": "Event",
    "description": "<p>The function returns r, c, new value, old value</p> <ul> <li>Call this function to return if it is -1 do not perform setValue and setCellFormula operations If it is 1, then call the onrelevantchange event!</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onvaluechange=function(idx,r,c){\n    //do something...\n};",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "ownerdraw",
    "name": "ownerdraw",
    "group": "Event",
    "description": "<p>Draw the contents of the cell yourself</p> <ul> <li>The function returns this r c index</li> <li>If the self-drawn cell does not exist, please call WB.createData(r,c,index).Cell text, value has no value or no attributes</li> <li>Please don't call startPaint function inside the function</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.createData(0,2,0)\nWB.ownerdraw=function(content,i,j,ind){\n    if(i==0&&j==2&&ind==0){\n        content.setLock(i,j,i,j,true)\n        var selectInfo = content.getSelectInfo(i,j,i,j,0)\n        var x = selectInfo.selectX;\n        var y = selectInfo.selectY;\n        var w = selectInfo.w;\n        var h = selectInfo.h;\n        content.ctx.arc(x+w/2,y+h/2,h/2-4,0,Math.PI*2,true)\n        content.ctx.moveTo(x+w/2-2,y+h/2-3)\n        content.ctx.arc(x+w/2-5,y+h/2-3,3,0,Math.PI*2,true)\n        content.ctx.moveTo(x+w/2+8,y+h/2-3)\n        content.ctx.arc(x+w/2+5,y+h/2-3,3,0,Math.PI*2,true)\n        content.ctx.moveTo(x+w/2+7,y+h/2+2)\n        content.ctx.arc(x+w/2,y+h/2+2,7,0,Math.PI,false)\n        content.ctx.fillStyle = \"#6495ED\";\n        content.ctx.strokeStyle = \"white\"\n        content.ctx.fill()\n        content.ctx.stroke()\n    }\n    return false;\n};",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Event"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "adaption",
    "name": "adaption",
    "group": "Function",
    "description": "<p>Adaptive function</p> <ul> <li>WB.adaption(bool,devW,devH)  You can also turn on adaptation by setting the adaption attribute</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "bool",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "devW",
            "description": "<p>The horizontal resolution of the computer screen during development</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "devH",
            "description": "<p>The vertical resolution of the computer screen during development</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.adaption(true,1280,720)   //Turn on adaptive",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "addChildCanvas",
    "name": "addChildCanvas",
    "group": "Function",
    "description": "<p>Add a child canvas element to insert into the page</p> <ul> <li>The function is mainly used to pop up another layer in the existing canvas table, which is another workbook</li> <li>WB.addChildCanvas(options)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "options",
            "description": "<p>Sub-canvas related parameters</p> <ul> <li>id(Child container id) This parameter needs to be passed to specify the container of the child canvas</li> <li>parentId(Parent container id)</li> <li>width(Defaults:250)</li> <li>height(Defaults:300)</li> <li>x(Defaults:0)</li> <li>y(Defaults:0)</li> <li>shade(Defaults:false)</li> </ul>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.addChildCanvas({\n    width:250,  //The width of the child canvas and the width of the container\n    height:300, //The height of the child canvas and the height of the container\n    x:280,  //Relative to the x coordinate of the parent container\n    y:90,   //Relative to the parent container's y coordinate\n    parentId:\"container\",   //Parent container default document.body\n    id:\"container2\",    //Container id of child canvan\n    shade:true, //Whether to show the shade\n});\n\nvar WB2 = new Workbook(\"container2\")\n\nWB2.setTableSize(250,300)\nWB2.maxCol(2)\nWB2.maxRow(3)\nWB2.setColWidth (125,0,1)\nWB2.showColHeading(true) \nWB2.setHeaderName(-1,0,'project')\nWB2.setHeaderName(-1,1,'Numerical value')\nWB2.activeSheet.colHeaderData.defaultDataNode.style.fillColor = \"#A9A9A9\"\nWB2.workbook.tabsOptions.fontColor = '#1E90FF'\nWB2.workbook.showTabs = 2\nWB2.bottomBtn=['yes','cancel']\nWB2.tabsHeight = 30;\nWB2.workbook.behaviorMode = 3\nWB.setActiveCanvas('container');\nWB2.setValue(0,1,\"Hidden bomb\",0)\nWB2.setLock(0,1,0,1,true)\nWB2.setValue(0,0,\"data\",0)\nWB2.startPaint()\nWB.onclickcellopenlayer = function(ind,r,c){\n    if(r==0&&c==1){\n        WB.showChildCanvas(\"container2\")   //show\n        WB.setActiveCanvas('container2');\n    }\n}\nWB2.onclickcellopenlayer = function(ind,r,c){\n    if(r==0&&c==1){\n        WB2.hiddenChildCanvas(\"container2\")   //hide\n        WB.setActiveCanvas('container');\n    }\n}",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "addSheet",
    "name": "addSheet",
    "group": "Function",
    "description": "<p>Add a new table</p> <ul> <li>WB.addSheet(data)   Create a new table and pass in the data</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": true,
            "field": "data",
            "description": "<p>The data of the table is optional</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.cellTypeContent(n,list,i)",
    "title": "cellTypeContent",
    "name": "cellTypeContent",
    "group": "Function",
    "description": "<p>Set the cell type to return an object format</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "n",
            "description": "<ul> <li>0：No button</li> <li>1：The button type is three dots</li> <li>2：Button type is subscript arrow</li> <li>3：checkBox(The value of this cell is not drawn and cannot be edited)</li> <li>4: select(Drop-down box)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Array",
            "optional": true,
            "field": "list",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "i",
            "description": "<p>Special item (reference onopenpopup)</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.cellTypeContent(4,[\"a\",\"b\",\"c\"])",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.clearAllBorder(R1,C1,R2,C2,Index)",
    "title": "clearAllBorder",
    "name": "clearAllBorder",
    "group": "Function",
    "description": "<p>Clear all the borders of the cells in the range (without parameters, the current range is cleared by default)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.clearAllBorder(2,2,5,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.clearAllStyle(R1,C1,R2,C2,Index)",
    "title": "clearAllStyle",
    "name": "clearAllStyle",
    "group": "Function",
    "description": "<p>Restore the default style of the cells in the range (no parameters default to the currently selected range of cells)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "clearFormatBrushStyle",
    "name": "clearFormatBrushStyle",
    "group": "Function",
    "description": "<p>Clear format brush style</p> <ul> <li>WB.clearFormatBrushStyle()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.deleteText(R1,C1,R2,C2,Index)",
    "title": "clearText",
    "name": "clearText",
    "group": "Function",
    "description": "<p>Clear the text of a range of cells(if there is text and value; text and value is set to empty)</p> <ul> <li>Cells that are locked/uneditable/checkbox cannot be deleted (no parameter defaults to clear the current range)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.clearText (1,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.col(column,index)",
    "title": "col",
    "name": "col",
    "group": "Function",
    "description": "<p>Set or return the active column of the current worksheet</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "column",
            "description": "<p>column(column&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.col()   //Get current active column\nWB.col(2)  //Set the current active column to the third column",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.createData(r,c,Index)",
    "title": "createData",
    "name": "createData",
    "group": "Function",
    "description": "<p>Initialize an empty cell (a cell with no value, no style, etc. will not be recorded, to be initialized within the range of existing rows and columns)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r",
            "description": "<p>row(r&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c",
            "description": "<p>col(c&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.deleteCell(R1,C1,R2,C2,Index)",
    "title": "deleteCell",
    "name": "deleteCell",
    "group": "Function",
    "description": "<p>Delete cell (including text and style, no parameter, delete current range by default)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.deleteCell (1,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.deleteFillColor(R1,C1,R2,C2,Index)",
    "title": "deleteFillColor",
    "name": "deleteFillColor",
    "group": "Function",
    "description": "<p>Delete the fill color of the cells in the range (no parameter clears the fill color in the range by default)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.deleteFillColor(0,0,2,2)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.deleteHeaderNames(R,C,Index)",
    "title": "deleteHeaderNames",
    "name": "deleteHeaderNames",
    "group": "Function",
    "description": "<p>Delete the name of the column header and row header</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>When deleting the row header, C passes -1 (remove both the column header and the row header name R&gt;=0 C&gt;=0) R&gt;=0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>R pass -1 (C&gt;=0) when deleting column header</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.deleteHeaderNames(1,-1)\nWB.deleteHeaderNames(-1,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "deleteSheets",
    "name": "deleteSheets",
    "group": "Function",
    "description": "<p>Delete worksheet</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheet",
            "defaultValue": "Current",
            "description": "<p>index] Starting index</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheets",
            "defaultValue": "1",
            "description": "<p>Number, how many tables are deleted</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.deleteSheets(1,2)  //Delete the second and third tables (at least one table remains)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.deleteText(R1,C1,R2,C2,Index)",
    "title": "deleteText",
    "name": "deleteText",
    "group": "Function",
    "description": "<p>Delete the contents of a range of cells (if there are text and value attributes; these two attributes will be deleted, which is different from clearText)</p> <ul> <li>Cells that are locked/uneditable/checkbox cannot be deleted; (the default range is deleted without parameters)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.deleteText (1,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "drawPrint",
    "name": "drawPrint",
    "group": "Function",
    "description": "<p>print</p> <ul> <li>Receive a callback and return an array (base64)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.drawPrint(function(arr){\n    console.log(arr)\n})",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.editDelete(nShiftType,R,C,N,Index)",
    "title": "editDelete",
    "name": "editDelete",
    "group": "Function",
    "description": "<p>Delete the current or specified row or column</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "nShiftType",
            "description": "<ul> <li>F1ShiftRows:(Delete row)</li> <li>F1ShiftCols(Delete column)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>Which row to delete, R&gt;=0, when the column is deleted, R passes -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>Which column to delete, C&gt;=0, when deleting rows, C passes -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "N",
            "defaultValue": "1",
            "description": "<p>N&gt;=1Number of rows or columns</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.editDelete  (\"F1ShiftRows\",0,-1)   //Delete the first line\nWB.editDelete  (\"F1ShiftCols\",-1,1)   //Delete the second column",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.editInsert(InsertType,R,C,N,Index)",
    "title": "editInsert",
    "name": "editInsert",
    "group": "Function",
    "description": "<p>Insert a new row or column in the current row or specified row</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "InsertType",
            "description": "<ul> <li>F1ShiftRows:(Insert a line, move down the line)</li> <li>F1ShiftCols(Insert a column, Move the column to the right)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>Insert in front of which row, R&gt;=0, R passes -1 when inserting column</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>Insert in front of which column, C&gt;=0, C insert -1 when inserting rows</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "N",
            "defaultValue": "1",
            "description": "<p>N&gt;=1Number of rows or columns</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.editInsert (\"F1ShiftRows\",0,-1)  //Insert a line before the first line\nWB.editInsert (\"F1ShiftCols\",-1,2)  //Insert a column before the third column",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.exportFileAsJson()",
    "title": "exportFileAsJson",
    "name": "exportFileAsJson",
    "group": "Function",
    "description": "<p>Save the current workbook as a json file and download</p>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.exportFileAsXml()",
    "title": "exportFileAsXml",
    "name": "exportFileAsXml",
    "group": "Function",
    "description": "<p>Save the current workbook as an xml file and download it (each table generates an xml)</p>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fixedFirstCol(startC,Index)",
    "title": "fixedFirstCol",
    "name": "fixedFirstCol",
    "group": "Function",
    "description": "<p>Freeze the first column of the visible area of the active table</p> <ul> <li>Call this function and the existing frozen rows will be removed</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startC",
            "defaultValue": "0",
            "description": "<p>Set which column is the starting column of the visual view</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fixedFirstCol()",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fixedFirstRow(startR,Index)",
    "title": "fixedFirstRow",
    "name": "fixedFirstRow",
    "group": "Function",
    "description": "<p>Freeze the first row of the visible area of the active table</p> <ul> <li>Call this function and the existing frozen column will be removed</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startR",
            "defaultValue": "0",
            "description": "<p>Set which line is the start row of the visual view</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fixedFirstRow()",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fixedRowAndCol(R,C,startR,startC,Index）",
    "title": "fixedRowAndCol",
    "name": "fixedRowAndCol",
    "group": "Function",
    "description": "<p>Freeze split pane</p> <ul> <li>fixedRows (How many rows are frozen and which rows start to freeze)</li> <li>fixedRow (What is the top row of the view after freezing)</li> <li>fixedCols (How many columns are frozen and which one starts to freeze)</li> <li>fixedRow (What is the leftmost column of the view after freezing;</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>In which line to start freezing (the freezing line is below this line) (R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>In which column to start freezing (C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startR",
            "defaultValue": "0",
            "description": "<p>Start row(startR&gt;=0,startR&lt;R)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startC",
            "defaultValue": "0",
            "description": "<p>Start column(startC&gt;=0,startC&lt;C)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fixedRowAndCol(1,1)   //Freeze the split cell at the position of cell (1,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fontColor(R1,C1,R2,C2,Color,Index)",
    "title": "fontColor",
    "name": "fontColor",
    "group": "Function",
    "description": "<p>Set the font color of the cell text in the range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "color",
            "defaultValue": "#000",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fontColor(0,0,2,2,'pink')   //Set the font to pink",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fontStrikeout(R1,C1,R2,C2,bool,Index)",
    "title": "fontStrikeout",
    "name": "fontStrikeout",
    "group": "Function",
    "description": "<p>Strikethrough of the cells in the set range (no parameter defaults to the strikethrough of the currently selected cell)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fontStrikeout(0,0,2,2,true)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.fontUnderline(R1,C1,R2,C2,bool,Index)",
    "title": "fontUnderline",
    "name": "fontUnderline",
    "group": "Function",
    "description": "<p>Underline of cells in the setting range (no parameter defaults to underlining of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.fontUnderline(0,0,2,2)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getActiveSheet",
    "name": "getActiveSheet",
    "group": "Function",
    "description": "<p>Set or get the current table</p> <ul> <li>WB.getActiveSheet ()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getAlignment(R1,C1,R2,C2,Index)",
    "title": "getAlignment",
    "name": "getAlignment",
    "group": "Function",
    "description": "<p>Get the alignment of the cells in the range (no parameter defaults to get the alignment of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getAlignment(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getBorder(R1,C1,R2,C2,Index)",
    "title": "getBorder",
    "name": "getBorder",
    "group": "Function",
    "description": "<p>Get the border of the cell in the range (without parameters, the border of the currently selected range is obtained by default)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getBorder(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getCanEdit(R1,C1,R2,C2,Index)",
    "title": "getCanEdit",
    "name": "getCanEdit",
    "group": "Function",
    "description": "<p>Get whether the cells in the range can be edited (no parameter defaults to the uneditable state of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getCanEdit(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getCellFormat(R1,C1,R2,C2,Index)",
    "title": "getCellFormat",
    "name": "getCellFormat",
    "group": "Function",
    "description": "<p>Get the format of the cell (without parameters, the format of the currently selected cell is obtained by default)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getCellFormat(1,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getCellFormula(R1,C1,R2,C2,Index)",
    "title": "getCellFormula",
    "name": "getCellFormula",
    "group": "Function",
    "description": "<p>Get the formula of the cell (the default is to get the formula of the currently selected range without parameters)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getCellFormula(1,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getCellRC",
    "name": "getCellRC",
    "group": "Function",
    "description": "<p>Get the range of rows and columns of the current table</p> <ul> <li>WB.getCellRC()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getCellType(R1,C1,R2,C2,Index)",
    "title": "getCellType",
    "name": "getCellType",
    "group": "Function",
    "description": "<p>Get the celltype of a range of cells (no parameter defaults to the celltype of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getCellType(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getEndSheet",
    "name": "getEndSheet",
    "group": "Function",
    "description": "<p>Get end table and index in tab bar of current workbook</p> <ul> <li>WB.getEndSheet ()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFillColor(R1,C1,R2,C2,Index)",
    "title": "getFillColor",
    "name": "getFillColor",
    "group": "Function",
    "description": "<p>Get the fill color of the cells in the range (no parameter defaults to the fill color of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getFillColor(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFont(R1,C1,R2,C2,Index)",
    "title": "getFont",
    "name": "getFont",
    "group": "Function",
    "description": "<p>Get the font of the selected cell (no parameter defaults to get the font of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getFont(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFontColor(R1,C1,R2,C2,Index)",
    "title": "getFontColor",
    "name": "getFontColor",
    "group": "Function",
    "description": "<p>Get the font color of the cells in the range (no parameter defaults to get the font color of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getFontColor(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFontList(r,c,Index)",
    "title": "getFontList",
    "name": "getFontList",
    "group": "Function",
    "description": "<p>Get every item of the font</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r",
            "description": "<p>row(r&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c",
            "description": "<p>column(c&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFontStrikeout(R1,C1,R2,C2,Index)",
    "title": "getFontStrikeout",
    "name": "getFontStrikeout",
    "group": "Function",
    "description": "<p>Whether to get the strikethrough of the cells in the range (no parameter defaults to get the strikethrough status of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getFontStrikeout(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getFontUnderline(R1,C1,R2,C2,Index)",
    "title": "getFontUnderline",
    "name": "getFontUnderline",
    "group": "Function",
    "description": "<p>Get whether the cells in the range are underlined (no parameter defaults to the underline state of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getFontUnderline(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getLock(R1,C1,R2,C2,Index)",
    "title": "getLock",
    "name": "getLock",
    "group": "Function",
    "description": "<p>Get whether the cell is locked (no parameter defaults to get the locked state of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getLock(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getSelectInfo(R1,C1,R2,C2,Index)",
    "title": "getSelectInfo",
    "name": "getSelectInfo",
    "group": "Function",
    "description": "<p>Get the coordinate width and height information of the selected cell</p> <ul> <li>Without parameters, the coordinate width and height of the currently selected range are obtained by default</li> <li>The width and height returned is the width and height seen</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getStartAndEnd(index)",
    "title": "getStartAndEnd",
    "name": "getStartAndEnd",
    "group": "Function",
    "description": "<p>Get the start row、start column、end row and end column of the view</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getStartAndEnd()",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getStartSheet",
    "name": "getStartSheet",
    "group": "Function",
    "description": "<p>Get the starting table and index in the tab bar of the current workbook</p> <ul> <li>WB.getStartSheet ()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getText(R1,C1,R2,C2,Index)",
    "title": "getText",
    "name": "getText",
    "group": "Function",
    "description": "<p>Get the value of the text attribute of the active cell, if there is no text attribute, get the value of the value attribute(no parameter defaults to the value of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getText(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getValue(R1,C1,R2,C2,Index)",
    "title": "getValue",
    "name": "getValue",
    "group": "Function",
    "description": "<p>Get the value of the cell (no parameter defaults to get the value of the currently selected cell)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getValue(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.getWordWrap(R1,C1,R2,C2,Index)",
    "title": "getWordWrap",
    "name": "getWordWrap",
    "group": "Function",
    "description": "<p>Get the status of whether the cells in the range are automatically wrapped (the default is to get the status of the word wrapping in the currently selected range without parameters)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.getWordWrap(0,0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.hdrHeight(height,index)",
    "title": "hdrHeight",
    "name": "hdrHeight",
    "group": "Function",
    "description": "<p>Set the height of the column header</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "height",
            "description": "<p>height&lt;=0  Will hide the column header and the height is 0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.hdrHeight(40)  //Set the height of the column header to 40px",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.hdrWidth(width,index)",
    "title": "hdrWidth",
    "name": "hdrWidth",
    "group": "Function",
    "description": "<p>Set the width of the row header</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "width",
            "description": "<p>width&lt;=0  Will hide the row header and the width is 0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.hdrWidth(40)  //Set the width of the row header to 40px",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "hiddenChildCanvas",
    "name": "hiddenChildCanvas",
    "group": "Function",
    "description": "<p>Hide all elements of child canvas</p> <ul> <li>WB2.hiddenChildCanvas(&quot;container2&quot;)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": "<p>Id of the child canvas container</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "importFile",
    "name": "importFile",
    "group": "Function",
    "description": "<p>File import (the file only supports .json, .xml, .xlsx, excel.zip)</p> <ul> <li>WB.exportFileAsXml()  f</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "f",
            "description": "<p>file</p>"
          },
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": ""
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "insertSheets",
    "name": "insertSheets",
    "group": "Function",
    "description": "<p>Insert n tables in front of the currently indexed table</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheet",
            "defaultValue": "Current",
            "description": "<p>index] Starting index</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheets",
            "defaultValue": "1",
            "description": "<p>Number, how many tables to insert  nSheets&gt;=1.</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.insertSheets(1,1)   //Insert a table before the second table\nWB.insertSheets()   //Insert a table before the current table",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "maxCol",
    "name": "maxCol",
    "group": "Function",
    "description": "<p>Get or set the number of columns in the current worksheet</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "num",
            "defaultValue": "Current",
            "description": "<p>number of columns]</p> <ul> <li>num&gt;Current number of columns.If the incoming num is greater than the current number of columns, increase the number of columns</li> <li>0&lt;=num&amp;&amp;num&lt;Current number of columns.Pass in num less than the current and greater than or equal to 0, then delete the number of columns (corresponding to merge, cell data clearing) corresponding to merge reduction (num&lt;0,num=0)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.maxCol(10)        //The total number of columns is 10\nWB.maxCol()        //Get the total number of columns",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "maxRow",
    "name": "maxRow",
    "group": "Function",
    "description": "<p>Get or set the number of rows in the current worksheet</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "num",
            "defaultValue": "Current",
            "description": "<p>table rows]</p> <ul> <li>num&gt;Current table rows.If the incoming num is greater than the current number of rows, the number of rows is increased</li> <li>0&lt;=num&amp;&amp;num&lt;Current table rows.The incoming num is less than the current and greater than or equal to 0, then delete the number of rows (corresponding to the merge, the cell data is cleared) corresponding to the merge reduction (num&lt;0,num=0)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.maxRow(9)       //The total number of rows is 9\nWB.maxRow()        //Get the total number of rows",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.mergeCells(R1,C1,R2,C2,Index)",
    "title": "mergeCells",
    "name": "mergeCells",
    "group": "Function",
    "description": "<p>Merge cells (no parameter defaults to the merge of the currently selected range)</p> <ul> <li>The merge will change the existing merged data, and only the content of the first cell will be retained after the merge</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.mergeCells(2,1,3,2)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "newWorkbook",
    "name": "newWorkbook",
    "group": "Function",
    "description": "<p>Create a new workbook and clear the current workbook data</p> <ul> <li>WB.newWorkbook()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": true,
            "field": "template",
            "description": "<p>Can be passed into the template</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "numSheets",
    "name": "numSheets",
    "group": "Function",
    "description": "<p>Set the number of sheets in the workbook</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "num",
            "description": "<ul> <li>num&gt;Total number of tables(Insert table after last table)</li> <li>num&lt;Total number of tables&amp;&amp;num&gt;=1(Delete table from behind)</li> <li>num&lt;=0(Will not delete at least one table)</li> </ul>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.numSheets(3)   //The total number of tables is set to 3",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.removeFixed(Index)",
    "title": "removeFixed",
    "name": "removeFixed",
    "group": "Function",
    "description": "<p>Unfreeze pane</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.removeFixed ()  // Remove all freezes",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.removeMergeCells(R1,C1,R2,C2,Index)",
    "title": "removeMergeCells",
    "name": "removeMergeCells",
    "group": "Function",
    "description": "<p>Cancel merged cells (no parameters default to cancel the merge of the current range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.removeMergeCells(2,1,3,2)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.removeRowColor(Index)",
    "title": "removeRowColor",
    "name": "removeRowColor",
    "group": "Function",
    "description": "<p>Cancel interlace discoloration</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.removeSplitcolHeader(Col,index)",
    "title": "removeSplitcolHeader",
    "name": "removeSplitcolHeader",
    "group": "Function",
    "description": "<p>Delete split column header</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "Col",
            "description": "<p>col(Col&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.removeSplitcolHeader(4)   //Delete the split column header of the fifth column",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "resumeEvent",
    "name": "resumeEvent",
    "group": "Function",
    "description": "<p>The resumeEvent function sets the value of stopEventCount. The value is decremented by 1 each time it is called until it is 0. stopEventCount is 0 and the on event is not executed.</p> <ul> <li>WB.resumeEvent()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "description": "<p>If true stopEventCount is directly assigned a value of 0</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.row(row,index)",
    "title": "row",
    "name": "row",
    "group": "Function",
    "description": "<p>Set or return the active row of the current worksheet</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "row",
            "description": "<p>row(row&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.row()   //Get current active line\nWB.row(2)  //Set the current active column to the third row",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setActiveCanvas",
    "name": "setActiveCanvas",
    "group": "Function",
    "description": "<p>Set the active canvas when there are multiple canvases on the same page (the default is the last new canvas, click on the corresponding canvas activecanvas will automatically switch)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setActiveCanvas(\"container2\")",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setActiveEditCell(R,C)",
    "title": "setActiveEditCell",
    "name": "setActiveEditCell",
    "group": "Function",
    "description": "<p>Open the specified edit cell</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>row(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>column(C&gt;=0)</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setActiveEditCell(1,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setAllBorder(R1,C1,R2,C2,NAll,CrALL,Index)",
    "title": "setAllBorder",
    "name": "setAllBorder",
    "group": "Function",
    "description": "<p>Set the full border of the cells in the range (no parameter defaults to the full border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NAll",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrALL",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setAllBorder(2,2,5,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setBottomBorder(R1,C1,R2,C2,NBottom,CrBottom,Index)",
    "title": "setBottomBorder",
    "name": "setBottomBorder",
    "group": "Function",
    "description": "<p>Set the bottom border of the cells in the range (no parameters default to the bottom border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NBottom",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrBottom",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setBottomBorder(1,1,3,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setCanEdit(R1,C1,R2,C2,bool,Index)",
    "title": "setCanEdit",
    "name": "setCanEdit",
    "group": "Function",
    "description": "<p>Set whether the cell can be edited (no parameter default setting currently selected range can not be edited)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "defaultValue": "false",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setCanEdit(1,1,1,1,false)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setCellFormat(R1,C1,R2,C2,type,Index)",
    "title": "setCellFormat",
    "name": "setCellFormat",
    "group": "Function",
    "description": "<p>Format cells in a range</p> <ul> <li>date：yyyy年m月d日 ,yyyy年m月 ,m月d日,yyyy-m-d ,yyyy-m,m-d ,yyyy-m-d:uppercase ,yyyy-m:uppercase ,星期n ,yyyy/m/d ,yyyy/m,m/d</li> <li>Scientific notation：E  (0E:Keep a decimal,00E:Keep two decimal places,The number of 0 is the number of reserved decimals)</li> <li>Numerical format：NR(A+thousands  (Remain as many decimals as there are 0s in front of N, use thousands separators after +thousands after N, and do not use thousands separators if there are none) R(A reference currency</li> <li>currency：MR(A+$   (Remain as many decimals as there are 0s in front of M, and set the color to red for negative numbers with R, yes (there are parentheses (only when A)), and A represents a positive number display, R(A There is no limit to the three character positions, but it must be placed after M. If there is a + currency symbol behind M, the currency symbol is added in front of it. If not, the currency symbol is not used Yuan display is almost ¥), US dollar, Euro, British pound, French franc, Korean won)</li> <li>percentage：P+%  (Remain as many decimals as there are 0s in front of P)</li> <li>Amount of capital：capitalMoney(Numbers are converted to upper-case amounts, and decimals are reserved to the point)</li> <li>text：Text format (return as is, including formulas)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "type",
            "description": "<p>Valid cell format value</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setCellFormat(0,0,2,c2,\"yyyy年m月d日\")   //The format of the cells in the set range is the date (\"yyyy年m月d日\")",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setCellFormula(R1,C1,Formula,Index)",
    "title": "setCellFormula",
    "name": "setCellFormula",
    "group": "Function",
    "description": "<p>Set cell formula</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>row(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>column(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Formula",
            "description": "<p>(Must start with = and at least one character after)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setCellFormula(0,0,'=A1+B1')   //Set the cell formula to =A1+B1",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)",
    "title": "setCellPadding",
    "name": "setCellPadding",
    "group": "Function",
    "description": "<p>Set padding of table cells</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "topSize",
            "description": "<p>padding-top</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "rightSize",
            "description": "<p>padding-right值</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "bottomSize",
            "description": "<p>padding-bottom</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "leftSize",
            "description": "<p>padding-left</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setCellPadding(0,0,0,5)    //Set the right padding of the cell to 5px",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setCellType(R1,C1,R2,C2,obj,Index)",
    "title": "setCellType",
    "name": "setCellType",
    "group": "Function",
    "description": "<p>Set the cellType of the cells in the range For obj parameter format, please refer to cellTypeContent method(var obj = WB.cellTypeContent(4,[&quot;apple&quot;,&quot;banana&quot;]))</p> <ul> <li>WB.setCellType(0,0,0,0,obj)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "obj",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setColHidden(C1,C2,BHidden,Index)",
    "title": "setColHidden",
    "name": "setColHidden",
    "group": "Function",
    "description": "<p>Set the column to hide, if the mouse drags the column width is less than 0 will also hide</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "BHidden",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setColHidden(1,2,true)  //Hide the second and third columns",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setColWidth(width,c1,c2,index)",
    "title": "setColWidth",
    "name": "setColWidth",
    "group": "Function",
    "description": "<p>Set the width of the columns in the range (hidden columns cannot be set)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Numer",
            "optional": false,
            "field": "width",
            "description": "<p>Value (width&lt;=0, hidden)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>Start column(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "c2",
            "description": "<p>End column(c1&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setColWidth(80,0,0)  //Set the width of the first column to 80",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setDefaultFontName",
    "name": "setDefaultFontName",
    "group": "Function",
    "description": "<p>Set the default font of the workbook</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "string",
            "optional": false,
            "field": "fontname",
            "description": "<p>Font name</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setDefaultFontName('楷体')",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setDefaultFontSize",
    "name": "setDefaultFontSize",
    "group": "Function",
    "description": "<p>Set the default font size of the workbook</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "fontsize",
            "description": "<p>font size</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setDefaultFontSize(16)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setDefaultLineHeight",
    "name": "setDefaultLineHeight",
    "group": "Function",
    "description": "<p>Set the default font line height of the workbook</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "lineheight",
            "description": "<p>Font line height</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setDefaultLineHeight(20)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setDocument",
    "name": "setDocument",
    "group": "Function",
    "description": "<p>Set the document that triggered the click event</p> <ul> <li>WB.setDocument(D)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "obj",
            "optional": false,
            "field": "d",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setDocument(window.top.document)   //Set to top",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFillColor(color,R1,C1,R2,C2,Index)",
    "title": "setFillColor",
    "name": "setFillColor",
    "group": "Function",
    "description": "<p>Set the fill color of the cells in the range (no parameter defaults to the fill color of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "color",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFillColor('pink',0,0,2,2)   //Set the fill color of the cells in the range to pink",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFontBold(R1,C1,R2,C2,BBold,Index)",
    "title": "setFontBold",
    "name": "setFontBold",
    "group": "Function",
    "description": "<p>Whether the font of the cells in the setting range is bold (no parameter defaults to setting the bold of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "BBold",
            "defaultValue": "true",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFontBold(1,1,1,1,true)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFontItalic(R1,C1,R2,C2,BItalic,Index)",
    "title": "setFontItalic",
    "name": "setFontItalic",
    "group": "Function",
    "description": "<p>Whether the font of the cells in the set range is italic (no parameter defaults to set the italic of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "BItalic",
            "defaultValue": "true",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFontItalic(1,1,1,1,true)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFontLineH(R1,C1,R2,C2,NLineHeight,Index)",
    "title": "setFontLineH",
    "name": "setFontLineH",
    "group": "Function",
    "description": "<p>Set the font line height of the cells in the range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "NLineHeight",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFontLineH(1,1,1,1,20)  //Set the font line height of the cells in the range to 20px",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFontName(R1,C1,R2,C2,PName,Index)",
    "title": "setFontName",
    "name": "setFontName",
    "group": "Function",
    "description": "<p>Set the font type of the cells in the range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "PName",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFontName(0,0,2,2,'宋体')",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFontSize(R1,C1,R2,C2,NSize,Index)",
    "title": "setFontSize",
    "name": "setFontSize",
    "group": "Function",
    "description": "<p>Set the font size of the cells in the range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "NSize",
            "description": "<p>font size</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setFontSize(1,1,1,1,20)  //Set the cell font size in the range to 20px",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setFormatBrushStyle(R,C,Index)",
    "title": "setFormatBrushStyle",
    "name": "setFormatBrushStyle",
    "group": "Function",
    "description": "<p>Format brush style</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setHAlignment(R1,C1,R2,C2,HAlign,Index)",
    "title": "setHAlignment",
    "name": "setHAlignment",
    "group": "Function",
    "description": "<p>Set the horizontal alignment of cell text in the range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "hAlign",
            "defaultValue": "1",
            "description": "<ul> <li>1: Regular (aligned according to cell data type)</li> <li>2：left</li> <li>3：center</li> <li>4：Right</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setHAlignment(0,0,0,0,2) //Set the cell text in the range to be left aligned",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setHeaderName(R,C,headerName,Index)",
    "title": "setHeaderName",
    "name": "setHeaderName",
    "group": "Function",
    "description": "<p>Change the names of column headers and row headers</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>Change the row header, change the column header when R passes -1(r&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>Change the column header, change the row header when C passes -1 (c&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "setHeaderName",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setHeaderName(1,-1,'name')\nWB.setHeaderName(-1,1,'name')",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setINnerBorder(R1,C1,R2,C2,NInner,CrInner,Index)",
    "title": "setINnerBorder",
    "name": "setINnerBorder",
    "group": "Function",
    "description": "<p>Set the inner border of the cell in the range (no parameter defaults to the inner border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NInner",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrInner",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setINnerBorder(1,1,5,5)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setLeftBorder(R1,C1,R2,C2,NLeft,CrLeft,Index)",
    "title": "setLeftBorder",
    "name": "setLeftBorder",
    "group": "Function",
    "description": "<p>Set the left border of the cells in the range (no parameter defaults to the left border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nLeft",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "crLeft",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setLeftBorder(1,1,3,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setLock(R1,C1,R2,C2,bool,Index)",
    "title": "setLock",
    "name": "setLock",
    "group": "Function",
    "description": "<p>Set whether the cells in the range are locked (no parameter defaults to the locking of the cells in the currently selected range)</p> <ul> <li>The locked cell cannot be edited, no check box is selected</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "defaultValue": "true",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setOutBorder(R1,C1,R2,C2,NOut,CrOut,Index)",
    "title": "setOutBorder",
    "name": "setOutBorder",
    "group": "Function",
    "description": "<p>Set the outer border of the cells in the range (no parameter defaults to the outer border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NOut",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrOut",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setOutBorder(1,1,5,5)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setPrint",
    "name": "setPrint",
    "group": "Function",
    "description": "<p>Set to print, the default value is used for parameters that are not passed</p> <ul> <li>marginTop</li> <li>marginBottom</li> <li>marginLeft</li> <li>marginRight</li> <li>paper (Paper type)</li> <li>printHeadings (Whether to print row and column labels)</li> <li>printGridLine (Whether to print grid lines)</li> <li>orientation (Paper orientation)</li> <li>printDirection (Print direction:)(1:N shape 2：Z shape)</li> <li>marginCopies (The interval between two data (tables) on the same paper)</li> <li>printer (printer)</li> <li>isshowfooterpageinfo (Display page number information parameters &quot;0&quot;, &quot;1&quot;, &quot;2&quot;, &quot;3&quot;, &quot;4&quot;, &quot;5&quot;, &quot;6&quot; correspond to: &quot;not displayed&quot;, &quot;footer centered&quot;, &quot;footer left&quot; &quot;Header right&quot;, &quot;Header centered&quot;, &quot;Header left&quot;, &quot;Header right&quot;)</li> <li>footpagestyle (Page number format)</li> <li>startR (Start row of print area)</li> <li>endR (end row of print area)</li> <li>startC (Start column of print area)</li> <li>endC (end column of print area)</li> <li>print (1 Print the current worksheet (default) 2 Print the workbook 3 Print the currently selected area of the current table)</li> <li>printOnSamePaper (Only effective when print is 2, whether multiple sheets are placed on the same sheet, default false)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "obj",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setRightBorder(R1,C1,R2,C2,NRight,CrRight,Index)",
    "title": "setRightBorder",
    "name": "setRightBorder",
    "group": "Function",
    "description": "<p>Set the right border of the cells in the range (no parameter defaults to the right border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NRight",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrRight",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setRightBorder(1,1,3,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setRowColor(color,oddEven,start,Index)",
    "title": "setRowColor",
    "name": "setRowColor",
    "group": "Function",
    "description": "<p>Set row color（Interlace color change, equivalent to setting the fill color of the cell)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "color",
            "defaultValue": "rgb(255,255,224)",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "oddEven",
            "defaultValue": "0",
            "description": "<ul> <li>0 : Even rows</li> <li>1 ：Odd rows</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "start",
            "defaultValue": "0",
            "description": "<p>From which row to start;</p> <ul> <li>0 ：Start at the top of the table</li> <li>1 ：Start from the following part of the frozen row</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setRowColor('rgb(255,255,224)',0)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setRowHeight(height,R1,R2,Index)",
    "title": "setRowHeight",
    "name": "setRowHeight",
    "group": "Function",
    "description": "<p>Set the height of the lines within the range (hidden lines cannot be set)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Numer",
            "optional": false,
            "field": "height",
            "description": "<p>Value (height&lt;=0, hidden)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setRowHeight(80,0,1)  //The first and second lines set the height of the line to 80",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setRowHidden(R1,R2,BHidden,Index)",
    "title": "setRowHidden",
    "name": "setRowHidden",
    "group": "Function",
    "description": "<p>Set the line to hide, if the mouse drags the line height is less than 0 will also hide</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "BHidden",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setRowHidden(1,2,true)  //Hide the second and third lines",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setScrollPosition(x,y,Index)",
    "title": "setScrollPosition",
    "name": "setScrollPosition",
    "group": "Function",
    "description": "<p>Set the position of the scroll bar</p> <ul> <li>In case of freezing, this function should be placed after the freezing function</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "x",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "y",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setTableSize(width,height)",
    "title": "setTableSize",
    "name": "setTableSize",
    "group": "Function",
    "description": "<p>Set the width and height of the canvas and container</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "width",
            "description": "<p>width(width&gt;0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "height",
            "description": "<p>height(height&gt;0)</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setTableSize(500,300)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setText(R,C,Text,Index)",
    "title": "setText",
    "name": "setText",
    "group": "Function",
    "description": "<p>Set the text value of the cell</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>row(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>col(C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Text",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setTopBorder(R1,C1,R2,C2,NTop,CrTop,Index)",
    "title": "setTopBorder",
    "name": "setTopBorder",
    "group": "Function",
    "description": "<p>Set the top border of the cells in the range (no parameter defaults to the top border of the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "NTop",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line</li> <li>2: Medium Line</li> <li>5: Thick Line</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "CrTop",
            "defaultValue": "#000",
            "description": "<p>Border color</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setTopBorder(1,1,3,3)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setVAlignment(R1,C1,R2,C2,VAlign,Index)",
    "title": "setVAlignment",
    "name": "setVAlignment",
    "group": "Function",
    "description": "<p>Set the vertical alignment of cell text in a range</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "vAlign",
            "defaultValue": "2",
            "description": "<ul> <li>1：top</li> <li>2：middle</li> <li>3：bottom</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setVAlignment(0,0,0,0,1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.setValue(R,C,Value,Index)",
    "title": "setValue",
    "name": "setValue",
    "group": "Function",
    "description": "<p>Set the cell's value (value text)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>row(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>column(C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Value",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.setValue(0,0,'test',1)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.sheetIndex(Index)",
    "title": "sheetIndex",
    "name": "sheetIndex",
    "group": "Function",
    "description": "<p>Set or get the current table index</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index&gt;=0&amp;&amp;&lt;Total number of tables</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.sheetIndex(1)   //Switch to the second table in the tab bar\nWB.sheetIndex()    //Get the current table index",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.sheetName()",
    "title": "sheetName",
    "name": "sheetName",
    "group": "Function",
    "description": "<p>Set or get the current worksheet name</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index&gt;=0&amp;&amp;&lt;Total number of tables</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "Name",
            "defaultValue": "Current",
            "description": "<p>table name] Table Name</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.sheetName(1,\"table two\")   //Change the name of the second table to'table two'\nWB.sheetName()   //Get the current table name",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "showChildCanvas",
    "name": "showChildCanvas",
    "group": "Function",
    "description": "<p>Show all elements of child canvas</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": ""
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.showChildCanvas(\"container2\")",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.showColHeading(boolean,index)",
    "title": "showColHeading",
    "name": "showColHeading",
    "group": "Function",
    "description": "<p>Whether to display column headers</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "boolean",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.showColHeading(false)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.showFixedLine(boolean,index)",
    "title": "showFixedLine",
    "name": "showFixedLine",
    "group": "Function",
    "description": "<p>Whether to display the frozen line of the current table</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "boolean",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.showFixedLine(true)   //show",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.showRowHeading(boolean,index)",
    "title": "showRowHeading",
    "name": "showRowHeading",
    "group": "Function",
    "description": "<p>Whether to display row headers</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "boolean",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.showRowHeading(false)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.splitcolHeader(Col,Name,Count,Height,index)",
    "title": "splitcolHeader",
    "name": "splitcolHeader",
    "group": "Function",
    "description": "<p>Split column header, only support splitting two rows</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "Col",
            "description": "<p>col(Col&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Name",
            "description": ""
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Count",
            "defaultValue": "2",
            "description": "<p>Number of columns split</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "Height",
            "defaultValue": "Half",
            "description": "<p>the height of the column head]  Height&lt;The height of the column head</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.splitcolHeader(4,\"name\") //Set the name of the split column header of the fifth column to \"name\", split the two columns by default, and the height is half the height of the column header",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "startPaint",
    "name": "startPaint",
    "group": "Function",
    "description": "<p>The startPaint function sets the value of stopPaintedCount.Each time the value is decremented by 1 until it is 0.redraw when stopPaintedCount is 0.</p> <ul> <li>WB.startPaint()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "description": "<p>If true stopPaintedCount is directly assigned a value of 0</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "stopEvent",
    "name": "stopEvent",
    "group": "Function",
    "description": "<p>The stopEvent function sets the value of stopEventCount, which is increased by 1 each time it is called.</p> <ul> <li>WB.stopEvent()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "stopPaint",
    "name": "stopPaint",
    "group": "Function",
    "description": "<p>The stopPaint function sets the value of stopPaintCount, which is increased by 1 every time it is called.</p> <ul> <li>WB.stopPaint()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.useFormatBrushStyle(R1,C1,R2,C2,Index)",
    "title": "useFormatBrushStyle",
    "name": "useFormatBrushStyle",
    "group": "Function",
    "description": "<p>Apply format brush styles within the range (no parameters default to the currently selected range of cells)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "WB.wordWrap(R1,C1,R2,C2,bool,Index)",
    "title": "wordWrap",
    "name": "wordWrap",
    "group": "Function",
    "description": "<p>Set whether the cells in the range are automatically wrapped (no parameters default to automatically wrap in the currently selected range)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>Start row(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>Start column(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>End row(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>End column(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>Whether to wrap</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "Current_table_index",
            "description": "<p>table index</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.wordWrap (0,0,1,1,true)",
        "type": "javascript"
      }
    ],
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  }
] });
