define({ "api": [
  {
    "type": "null",
    "url": "/null",
    "title": "activeCol",
    "name": "activeCol",
    "group": "Attribute_activeSheet",
    "description": "<ul> <li>获取当前表的当前列  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</li> <li>通过WB.sheetIndex(Index)设置当前表  则这样调用WB.activeSheet.activeCol</li> <li>或者var activeSheet = WB.getActiveSheet(Index)获取当前表  则这样调用activeSheet.activeCol</li> </ul>",
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
    "description": "<p>获取当前表的当前行  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.activeRow,activeRow.activeRow</li> </ul>",
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
    "description": "<p>当前表是否允许拖选单元格(视角上的选择范围)(默认:true)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.allowMoveRange = 值(boolean),activeSheet.allowMoveRang = 值(boolean)</li> </ul>",
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
    "description": "<p>是否始终显示当前表单元格的button(默认:true)  false值的情况下只有进入编辑的情况才会显示btn</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>设置btn参考setCellType</li> <li>WB.activeSheet.alwaysShowButton = 值(boolean),activeSheet.alwaysShowButton = 值(boolean)</li> </ul>",
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
    "description": "<p>始终显示当前表单元格btn的情况下  是否只有在点击在该单元格再显示(默认:true)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.alwaysShowInArea = 值(boolean),activeSheet.alwaysShowInArea = 值(boolean)</li> </ul>",
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
    "description": "<p>是否点当前表击单元格能触发相应事件(默认:true)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.canOpenLayer = 值(boolean),activeSheet.canOpenLayer = 值(boolean)</li> </ul>",
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
    "description": "<p>当前表的单元格内容的下内边距(默认:2)  亦可可通过调用方法WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.cellPadding.bottom = 值(number),activeSheet.cellPadding.bottom = 值(number)</li> </ul>",
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
    "description": "<p>当前表的单元格内容的左内边距(默认:2)  亦可通过调用方法WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.cellPadding.left = 值(number),activeSheet.cellPadding.left = 值(number)</li> </ul>",
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
    "description": "<p>当前表的单元格内容的右内边距(默认:2)  亦可可通过调用方法WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.cellPadding.right = 值(number),activeSheet.cellPadding.right = 值(number)</li> </ul>",
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
    "description": "<p>当前表的单元格内容的上内边距(默认:2)  亦可可通过调用方法WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.cellPadding.top = 值(number),activeSheet.cellPadding.top = 值(number)</li> </ul>",
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
    "description": "<p>当前表列头填充颜色(默认:#008844)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.colHeaderData.defaultDataNode.style.fillColor = 值(有效颜色值)，activeSheet.colHeaderData.defaultDataNode.style.fillColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>当前表列头字体颜色(默认:#fff)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.colHeaderData.defaultDataNode.style.fontColor = 值(有效颜色值)，activeSheet.colHeaderData.defaultDataNode.style.fontColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>当前表列头的高度(默认:30)  可用WB.hdrHeight(height,index)函数设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.colHeaderData.defaultH，activeSheet.colHeaderData.defaultH</li> </ul>",
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
    "description": "<p>是否显示当前表列头(默认:false)  可用WB.showColHeading(boolean,index)函数设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.colHeaderData.showColHeading，activeSheet.colHeaderData.showColHeading</li> </ul>",
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
    "description": "<p>当前表默认列宽(默认:60)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.defaults.colWidth = 值(number)，activeSheet.defaults.colWidth = 值(number)</li> </ul>",
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
    "description": "<p>当前表默认行高(默认:30)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.defaults.rowHeight = 值(number) ，activeSheet.defaults.rowHeight = 值(number)</li> </ul>",
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
    "description": "<p>获取当前表可视区域结束列</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.endCol,activeSheet.endCol</li> </ul>",
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
    "description": "<p>获取当前表可视区域结束行</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.endRow,activeSheet.endRow</li> </ul>",
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
    "description": "<p>当前表的网格线颜色(默认:ccce)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.gridLinesColor = 值(有效颜色值),activeSheet.gridLinesColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>是否开启当前表头模式（冻结行上方没有网格线）(默认:false)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.headeNotLineInFix = 值(boolean),activeSheet.headeNotLineInFix = 值(boolean)</li> </ul>",
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
    "description": "<p>是否隐藏选中当前表单元格的选中框(默认:false)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.isSelectionHideBorder = 值(boolean),activeSheet.isSelectionHideBorder = 值(boolean)</li> </ul>",
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
    "description": "<p>当前表哪一列没有选中框(默认:-1全部有选中框)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.notColSelectionNum = 值(num),activeSheet.notColSelectionNum = 值(num)</li> </ul>",
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
    "description": "<p>当前表哪一行没有选中框(默认:-1全部有选中框)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.notRowSelectionNum = 值(num),activeSheet.notRowSelectionNum = 值(行NUM)</li> </ul>",
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
    "description": "<p>当前表打印页码格式  第 &amp;p 页，共 &amp;P 页  (&amp;p)固定(&amp;p &amp;P注意区分大小写 这两个符号会替换成当前第几页 一共几页得数字) 方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.footpagestyle = 值(string),activeSheet.printSetting.footpagestyle = 值(string)</li> </ul>",
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
    "description": "<p>当前表  方法WB.setPrint(obj,Index) 显示页码信息参数&quot;0&quot;(默认), &quot;1&quot;, &quot;2&quot;, &quot;3&quot;, &quot;4&quot;, &quot;5&quot;, &quot;6&quot;对应:&quot;不显示&quot;, &quot;页脚居中&quot;, &quot;页脚居左&quot;, &quot;页脚居右&quot;, &quot;页眉居中&quot;, &quot;页眉居左&quot;, &quot;页眉居右</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.isshowfooterpageinfo = 值(int),activeSheet.printSetting.isshowfooterpageinfo = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印下边距;单位为毫米默认5mm   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.marginBottom = 值(int),activeSheet.printSetting.marginBottom = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印左边距;单位为毫米默认5mm   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.marginLeft = 值(int),activeSheet.printSetting.marginLeft = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印右边距;单位为毫米默认5mm   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.marginRight = 值(int),activeSheet.printSetting.marginRight = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印上边距;单位为毫米默认5mm  方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.marginTop = 值(int),activeSheet.printSetting.marginTop = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表纸张方向 默认（1） 纵向  2横向 方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.orientation = 值(int),activeSheet.printSetting.orientation = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印的纸张类型默认 A4 (A3-7)   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.paper = 值(String),activeSheet.printSetting.paper = 值(String)</li> </ul>",
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
    "description": "<p>设置当前表页打印区域开始单元格的c1   默认''空不设置  方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printArea.c1 = 值(int),activeSheet.printSetting.printArea.c1 = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表页打印区域结束单元格的c1   默认''空不设置  方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printArea.c2 = 值(int),activeSheet.printSetting.printArea.c2 = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表页打印区域开始单元格的r1   默认''空不设置  方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printArea.r1 = 值(int),activeSheet.printSetting.printArea.r1 = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表页打印区域结束单元格的r2   默认''空不设置  方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printArea.r2 = 值(int),activeSheet.printSetting.printArea.r2 = 值(int)</li> </ul>",
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
    "description": "<p>设置当前表打印方向 默认（1） 先列后行呈N字形  2先行后列呈Z字形   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printDirection = 值(int),activeSheet.printSetting.printDirection = 值(int)</li> </ul>",
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
    "description": "<p>是否打印当前表网格线 默认（false）   方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printGridLine = 值(bool),activeSheet.printSetting.printGridLine = 值(bool)</li> </ul>",
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
    "description": "<p>是否打印当前表行号列标 默认（false） 方法WB.setPrint(obj,Index)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.printSetting.printHeadings = 值(bool),activeSheet.printSetting.printHeadings = 值(bool)</li> </ul>",
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
    "description": "<p>获取当前表的选中范围开始列(==activeCol)  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rangeCol1,activeRow.rangeCol1</li> </ul>",
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
    "description": "<p>获取当前表的选中范围的结束列  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rangeCol2,activeSheet.rangeCol2</li> </ul>",
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
    "description": "<p>获取当前表的选中范围开始行(==activeRow)  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rangeRow1,activeSheet.rangeRow1</li> </ul>",
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
    "description": "<p>获取当前表的选中范围结束行  可通过WB.getCellRC(Index)方法获取r1,c1,r2,c2四个参数 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rangeRow2,activeSheet.rangeCol2</li> </ul>",
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
    "description": "<p>当前表行头填充颜色(默认:#008844)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rowHeaderData.defaultDataNode.style.fillColor = 值(有效颜色值),activeSheet.rowHeaderData.defaultDataNode.style.fillColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>当前表行头字体颜色(默认:#fff)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rowHeaderData.defaultDataNode.style.fontColor = 值(有效颜色值),activeSheet.rowHeaderData.defaultDataNode.style.fontColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>当前表行头的宽度(默认:30)  可用WB.hdrWidth(width,index)函数设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rowHeaderData.defaultW = 值(number)</li> </ul>",
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
    "description": "<p>是否显示当前表行头(默认:false) 可用WB.showRowHeading(boolean,index)函数设置</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.rowHeaderData.showRowHeading = 值(boolean),activeSheet.rowHeaderData.showRowHeading = 值(boolean)</li> </ul>",
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
    "description": "<p>当前表是否选中了左上角的全选(默认:false)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.selHdrTopLeft = 值(boolean),activeSheet.selHdrTopLeft = 值(boolean)</li> </ul>",
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
    "description": "<p>当前表单元格选中模式(0range模式，自由选择，1:row行模式，点击单元格选中行，2:col 列模式，点击单元格选中列，3:cell单元格模式 不可拖选)(默认:0)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.selectMode = 值(0|1|2|3),activeSheet.selectMode = 值(0|1|2|3)</li> </ul>",
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
    "description": "<p>当前表选中框颜色(默认:green)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.selectOption.selectBorderColor  = 值(有效颜色值),activeSheet.selectOption.selectBorderColor  = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>当前表选中框填充颜色(默认:rgba(0,0,245,0.1))</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.selectOption.selectFillColor  = 值(有效颜色值),activeSheet.selectOption.selectFillColor  = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>获取当前表的工作表名  设置工作表名称请调用WB.sheetName(Index，Name)方法 其中WB是new出来的实例</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.sheetName,activeSheet.sheetName</li> </ul>",
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
    "description": "<p>是否显示当前表冻结线(默认:false)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.showFixedLine = 值(boolean) ,activeSheet.showFixedLine = 值(boolean)</li> </ul>",
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
    "description": "<p>是否显示当前表网格线(默认:true)</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.showGridLines = 值(boolean),activeSheet.showGridLines = 值(boolean)</li> </ul>",
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
    "description": "<p>获取当前表可视区域开始列(从冻结列的最后一列的下一列开始算) 若前面有冻结了两列  则开始列为第三列</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.startCol,activeSheet.startCol</li> </ul>",
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
    "description": "<p>获取当前表可视区域开始行(从冻结行的最后一行的下一行开始算) 若前面有冻结了两行  则开始列为第三行</p> <ul> <li>当前表的获取跟设置请看(sheetIndex,getActiveSheet)</li> <li>WB.activeSheet.startRow,activeSheet.startRow</li> </ul>",
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
    "description": "<p>在showtabs显示方式设置为2的情况下 显示btn按钮组</p> <ul> <li>WB.bottomBtn=['确定','取消']</li> </ul>",
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
    "description": "<p>单元格btn边框(celltype ==1  celltype==2)(默认:#C5C1AA)</p> <ul> <li>WB.btnBorderColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>单元格btn边框（hover）(celltype ==1  celltype==2)(默认:rgba(30,144,255,.4))</p> <ul> <li>WB.btnBorderColorH = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>单元格btn填充颜色(celltype ==1  celltype==2)(默认:#F0F0F0)</p> <ul> <li>WB.btnFillColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>单元格填充颜色(hover)(celltype ==1  celltype==2)(默认:rgba(240,248,255))</p> <ul> <li>WB.btnFillColorH = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>设置获取复选框btn边框颜色</p> <ul> <li>WB.checkBoxBorderColor=&quot;颜色值&quot;</li> </ul>",
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
    "description": "<p>设置获取复选框边框（hover）颜色</p> <ul> <li>WB.checkBoxBorderColorH=&quot;颜色值&quot;</li> </ul>",
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
    "description": "<p>设置获取复选框填充颜色</p> <ul> <li>WB.checkBoxFillColor=&quot;颜色值&quot;</li> </ul>",
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
    "description": "<p>设置获取复选框填充颜色(hover)</p> <ul> <li>WB.checkBoxFillColorH=&quot;颜色值&quot;</li> </ul>",
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
    "description": "<p>获取或者设置单元格的默认属性值,设置的值为一个对象 没有改变某个属性则使用默认的值 默认值有如下</p> <ul> <li>&quot;fontColor&quot;:&quot;#000&quot;</li> <li>&quot;hAlign&quot;:1</li> <li>&quot;cellType&quot;:{&quot;name&quot;:0}</li> <li>&quot;lock&quot;:false</li> <li>&quot;vAlign&quot;:2</li> <li>&quot;formatter&quot;:&quot;&quot;</li> <li>&quot;font&quot;:this.workbook.defaultFontSize+&quot;px/&quot;+this.workbook.defaultLineHeight+&quot;px &quot;+ this.workbook.defaultFontName</li> <li>&quot;wordWrap&quot;:false</li> <li>WB.defaultCellStyle={&quot;fontColor&quot;:&quot;red&quot;}  设置单元格默认属性的字体颜色为红色;</li> </ul>",
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
    "description": "<p>设置获取dpi(默认96)</p> <ul> <li>WB.dpi = 值(Int)</li> </ul>",
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
    "description": "<p>设置触发输入的时候选中文本的结束位置</p> <ul> <li>WB.editSelEnd=3   第三个字结束</li> </ul>",
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
    "description": "<p>设置触发输入的时候选中文本的开始位置</p> <ul> <li>WB.editSelStart=1   从第二个字开始</li> </ul>",
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
    "description": "<p>按下回车键之后单元格是否向右移动(默认:false),改属性为true时,回车右移单元格至最后一列跳到下一行</p> <ul> <li>WB.enterToRight = 值(bool)</li> </ul>",
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
    "description": "<p>设置table初始化是否进入编辑状态（默认false）</p> <ul> <li>WB.isEdit=true  table初始化进入编辑</li> </ul>",
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
    "description": "<p>可根据isFocus或者isEdit判断是否在编辑中（默认false）</p>",
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
    "description": "<p>设置点击新建一个表的时候的列数</p> <ul> <li>WB.newSheetCol = 值(Int)</li> </ul>",
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
    "description": "<p>设置点击新建一个表的时候的行数</p> <ul> <li>WB.newSheetRow = 值(Int)</li> </ul>",
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
    "description": "<p>设置获取滚动条大小</p> <ul> <li>WB.scrollSize=8   滚动大小为8px</li> </ul>",
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
    "description": "<p>设置获取tabs栏切换表箭头高度</p> <ul> <li>WB.tabArrowHeight=8  tabs栏切换表箭头高度为8px</li> </ul>",
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
    "description": "<p>设置获取tabs栏切换表箭头(第一个箭头)距离tabs栏左边（x水平）的位置</p> <ul> <li>WB.tabArrowStartX=15  tabs栏切换表箭头距离tabs栏左边（x水平）的位置为15px</li> </ul>",
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
    "description": "<p>设置获取tabs栏切换表箭头宽度</p> <ul> <li>WB.tabArrowWidth=5  tabs栏切换表箭头宽度为5px</li> </ul>",
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
    "description": "<p>设置tabs栏的高度</p> <ul> <li>WB.tabsHeight=30   tabs栏的高度为30px</li> </ul>",
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
    "description": "<p>是否开启自适应(默认:false)   开启自适应请在new实例之后调用WB.adaption(bool,devW,devH)</p> <ul> <li>WB.workbook.adaption = 值(boolean)</li> </ul>",
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
    "description": "<p>是否允许删除键删除选中的文本(默认:true)</p> <ul> <li>WB.workbook.allowDelete = 值(boolean)</li> </ul>",
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
    "description": "<p>否允许通过拖动调整列宽行高(默认:true)</p> <ul> <li>WB.workbook.allowResize = 值(boolean)</li> </ul>",
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
    "description": "<p>是否允许使用tab键切换单元格(默认:true)</p> <ul> <li>WB.workbook.allowTabs = 值(boolean)</li> </ul>",
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
    "description": "<p>工作簿模式(默认:1)</p> <ul> <li>WB.workbook.behaviorMode = 值( 1:sheet模式   2:grid模式（点击单元格就有输入框、没有选中框） 3:sheet单击可编辑模式、有选中框)</li> </ul>",
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
    "description": "<p>开发时候的电脑分辨率的高 结合adaption属性是否开启自适应  开启自适应请在new实例之后调用WB.adaption(bool,devW,devH)</p> <ul> <li>WB.workbook.devScreenHeight = 值(boolean)</li> </ul>",
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
    "description": "<p>开发时候的电脑分辨率的宽 结合adaption属性是否开启自适应  开启自适应请在new实例之后调用WB.adaption(bool,devW,devH)</p> <ul> <li>WB.workbook.devScreenWidth = 值(boolean)</li> </ul>",
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
    "description": "<p>工作簿宽度(默认:视口高度-系统滚动条size)</p> <ul> <li>WB.workbook.height = 值(number)</li> </ul>",
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
    "description": "<p>在同一张纸上方打印两个表的数据 两条数据(表)之间的间隔单位mm</p> <ul> <li>WB.workbook.marginCopies = 值(int)</li> </ul>",
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
    "description": "<p>打印   1打印当前工作表(默认) 2打印工作簿 3打印当前表当前选中区域 方法WB.setPrint(obj,Index)</p> <ul> <li>WB.workbook.print = 值(Int)</li> </ul>",
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
    "description": "<p>printOnSamePaper 仅当print为2时生效 在纸张可以放下其他表的情况下 多张表是否放在同一张纸上 默认false 方法WB.setPrint(obj,Index)</p> <ul> <li>WB.workbook.printOnSamePaper = 值(Boolean)</li> </ul>",
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
    "description": "<p>打印机   方法WB.setPrint(obj,Index)</p> <ul> <li>WB.workbook.printer = 值(string)</li> </ul>",
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
    "description": "<p>当前行箭头的填充色(默认:#000)</p> <ul> <li>WB.workbook.rowArrowColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>滚动条模式(默认:2)</p> <ul> <li>WB.workbook.scrollMode = 值(1滚动条从顶部或者最左部开始  2滚动条再冻结线下方或者右方)</li> </ul>",
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
    "description": "<p>是否显示右击菜单栏(默认:false)</p> <ul> <li>WB.workbook.showContextMenu = 值(boolean)</li> </ul>",
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
    "description": "<p>工作簿水平滚动条的显示方式(默认:1)</p> <ul> <li>WB.workbook.showHScrollBar = 值(0不显示 1工作簿较宽的时候显示)</li> </ul>",
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
    "description": "<p>是否显示当前行的箭头(默认:false)</p> <ul> <li>WB.workbook.showRowArrow = 值(boolean)</li> </ul>",
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
    "description": "<p>是否显示tabs栏(0 不显示 1显示 2按钮组 默认:0)</p> <ul> <li>WB.workbook.showTabs = 值(Int)</li> </ul>",
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
    "description": "<p>工作簿垂直滚动条的显示方式(默认:1)</p> <ul> <li>WB.workbook.showVScrollBar = 值(0不显示 1工作簿较宽的时候显示)</li> </ul>",
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
    "description": "<p>是否显示工作簿的边框(默认:true)</p> <ul> <li>WB.workbook.showWorkBookBorder = 值(boolean)</li> </ul>",
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
    "description": "<p>tab栏每个表的边框颜色(默认:#DCDCDC)</p> <ul> <li>WB.workbook.tabsOptions.borderColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>tabs栏的背景填充色(默认:#FDF5E6)</p> <ul> <li>WB.workbook.tabsOptions.fillColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>tab栏的字体(默认:14px/20px 宋体)</p> <ul> <li>WB.workbook.tabsOptions.font = 值(字体大小/行高 字体类型)</li> </ul>",
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
    "description": "<p>tab栏的字体颜色(默认:#444)</p> <ul> <li>WB.workbook.tabsOptions.fontColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>tab栏选中之后的字体颜色(默认:#CD853F)</p> <ul> <li>WB.workbook.tabsOptions.fontSelColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>tab选中当前表的填充颜色(默认:#fff)</p> <ul> <li>WB.workbook.tabsOptions.selFillColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>是否显示tabs栏添加一个表的的+号(默认:true)</p> <ul> <li>WB.workbook.tabsOptions.showAdd = false</li> </ul>",
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
    "description": "<p>工作簿宽度(默认:视口宽度-系统滚动条size)</p> <ul> <li>WB.workbook.width = 值(number)</li> </ul>",
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
    "description": "<p>设置工作簿边框颜色(默认:#ccc)</p> <ul> <li>WB.workbook.workBookBorderColor = 值(有效颜色值)</li> </ul>",
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
    "description": "<p>回车新增一行后的事件 f(r,c)</p> <ul> <li>WB.onafteraddrow=function(r,c){//};</li> </ul>",
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
    "title": "oncellbuttonclick",
    "name": "oncellbuttonclick",
    "group": "Event",
    "description": "<p>点击cell button 触发事件，具体函数看需求实现 函数返回(表索引 r c)</p> <ul> <li>WB.oncellbuttonclick=function(idx,r,c){alert(&quot;oncellbuttonclick&quot;+idx+&quot;,&quot;+r+&quot;,&quot;+c)};</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>点击cellBtn触发相应事件.</p>"
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
    "title": "onclickcellopenlayer",
    "name": "onclickcellopenlayer",
    "group": "Event",
    "description": "<p>点击单元格是否触发事件，具体函数看需求实现  函数返回(表索引 r c)</p> <ul> <li>WB.onclickcellopenlayer=function(idx,r,c){alert(&quot;onclickcellopenlayer&quot;+idx+&quot;,&quot;+r+&quot;,&quot;+c)};</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>点击单元格触发相应事件.</p>"
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
    "description": "<p>在最后一行根据函数的返回值(true)判断是否增加一行可以指定在最后一行的那个单元格返回true f(r,c,v)</p> <ul> <li>WB.onenteraddrow=function(r,c,v){return true};</li> </ul>",
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
    "title": "onequalsign",
    "name": "onequalsign",
    "group": "Event",
    "description": "<p>输入框输入 =（等号） 事件 函数返回(r c)</p> <ul> <li>WB.onequalsign=function(r,c){//)};</li> </ul>",
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
    "title": "onformularesult",
    "name": "onformularesult",
    "group": "Event",
    "description": "<p>返回公式结果事件 函数返回(公式 r c)   f(formula,R1,C1);</p> <ul> <li>WB.onformularesult = function(f,r,c){}</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>传入的回调函数</p>"
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
    "title": "onopenmorelist",
    "name": "onopenmorelist",
    "group": "Event",
    "description": "<p>参考onopenpopup 点击哪一项触发该事件,函数需要自己放在onopenpopup success里调用</p>",
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
    "description": "<p>引入popup.js,cellType.name==4的单元格会弹出弹窗,该操作会new出来一个popup实例,具体函数看需求实现,返回10个值w（宽）,h(高),x（x坐标）,y（y坐标）,list（列表数据） r(行) c(列) child(对应的输入框),fontSize(字体大小),noSetIndex(特殊的项，可判断点击在该项上发生对应事件)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>点击触发cellType.name==4相应事件.</p>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.onopenpopup = function(w,h,x,y,list,r,c,child,fontSize,noSetIndex){\n    var data =[];\n    if(list){\n        for(var i =0;i<list.length;i++){\n            data.push({name:list[i]})\n        }\n    };\n    WB.popup._init({\n        shade: false,\n        area:[w+'px',list.length*h+\"px\"],\n        offset:[x+'px', y+'px'],\n        backgroundColor: '#dccccc',\n        rowHeight:34+\"px\",\n        data:data,\n        fontSize:fontSize,\n        success: function(index,value,listIndex) {\n            if(listIndex==noSetIndex){\n                WB.onopenmorelist = function(){\n                    // alert(\"打开更多弹窗\")\n                };\n                if(typeof(WB.onopenmorelist)=='function'){\n                    WB.onopenmorelist()\n                }\n            }else{\n                WB.setValue(r,c,value)\n                child.innerHTML = value\n                WB.startPaint(true)\n            }\n            this.close(index); \n        },\n    })\n}",
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
    "description": "<p>行选模式双击下触发事件，具体函数看需求实现  函数返回(表索引 r c)</p> <ul> <li>WB.onrowmodedblclick=function(idx,r,c){};</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>行选模式双击下触发事件.</p>"
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
    "title": "onshoweditor",
    "name": "onshoweditor",
    "group": "Event",
    "description": "<p>根据用户设置的editSelStart(光标开始位置)和editSelEnd(光标结束位置选定文本),如果有传callback但没有设置这两个参数，默认选中全部文本；没有callback光标默认再开始位置：child是编辑文本的box（div）</p> <ul> <li>WB.onshoweditor=function(child){var text = child.innerHTML;WB.editSelStart = text.length;WB.editSelEnd = text.length;};设置光标位置再最后</li> </ul>",
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
    "title": "onvaluechange",
    "name": "onvaluechange",
    "group": "Event",
    "description": "<p>该函数返回r,c,新值，旧值，具体函数看需求实现,调用函数返回如果是-1不执行setValue，setCellFormula操作,如果是1则调用onrelevantchange事件！</p> <ul> <li>WB.onvaluechange=function(idx,r,c){};</li> </ul>",
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
    "title": "ownerdraw",
    "name": "ownerdraw",
    "group": "Event",
    "description": "<p>给用户自己画单元格的内容 ownerdraw函数返回当前对象的this r c index 若要自绘的单元格不存在请先调用WB.createData(r,c,index) 创建该单元格并且单元格text,value无值或者无这俩属性。注意：请不要在函数内调用startPaint函数</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>用户自己画单元格的内容回调函数.</p>"
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
    "description": "<p>调用自适应函数  WB.adaption(bool,devW,devH)  亦可以通过设置adaption属性来开启自适应</p> <ul> <li>WB.adaption(true,1280,720)   开启自适应</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "bool",
            "description": "<p>布尔值</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "devW",
            "description": "<p>开发时候电脑屏幕分辨率的宽</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "devH",
            "description": "<p>开发时候电脑屏幕分辨率的高</p>"
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
    "title": "addChildCanvas",
    "name": "addChildCanvas",
    "group": "Function",
    "description": "<p>新增一个子canvas元素插入页面中(函数主要用于在现有的canvas表中弹出另一个层,该层是另一个workbook)</p> <ul> <li>WB.addChildCanvas(options)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "options",
            "description": "<p>子canvas相关参数</p> <ul> <li>id(子容器id) 该参数需传 用于指定子canvas的容器 如container2(需new出来 var WB2 = new Workbook(&quot;container2&quot;))</li> <li>parentId(父容器id不传插入到body中)</li> <li>width(默认250)</li> <li>height(默认300)</li> <li>x(坐标x 默认0)</li> <li>y(坐标y 默认0)</li> <li>shade(蒙层 默认false)</li> </ul>"
          }
        ]
      }
    },
    "examples": [
      {
        "title": "demo:",
        "content": "WB.addChildCanvas({\n    width:250,  //子canvas宽以及容器宽\n    height:300, //子canvas高以及容器高\n    x:280,  //相对于父容器的x坐标\n    y:90,   //相对于父容器的y坐标\n    parentId:\"container\",   //父容器 默认document.body\n    id:\"container2\",    //子canvan容器id\n    shade:true, //是否显示蒙层\n});\n\nvar WB2 = new Workbook(\"container2\")\n\nWB2.setTableSize(250,300)\nWB2.maxCol(2)\nWB2.maxRow(3)\nWB2.setColWidth (125,0,1)\nWB2.showColHeading(true) \nWB2.setHeaderName(-1,0,'项目')\nWB2.setHeaderName(-1,1,'数值')\nWB2.activeSheet.colHeaderData.defaultDataNode.style.fillColor = \"#A9A9A9\"\nWB2.workbook.tabsOptions.fontColor = '#1E90FF'\nWB2.workbook.showTabs = 2\nWB2.bottomBtn=['确定','取消']\nWB2.tabsHeight = 30;\nWB2.workbook.behaviorMode = 3\nWB.setActiveCanvas('container');\nWB2.setValue(0,1,\"隐藏弹层\",0)\nWB2.setLock(0,1,0,1,true)\nWB2.setValue(0,0,\"数据\",0)\nWB2.startPaint()\nWB.onclickcellopenlayer = function(ind,r,c){\n    if(r==0&&c==1){\n        WB.showChildCanvas(\"container2\")   //显示\n        WB.setActiveCanvas('container2');\n    }\n}\nWB2.onclickcellopenlayer = function(ind,r,c){\n    if(r==0&&c==1){\n        WB2.hiddenChildCanvas(\"container2\")   //隐藏\n        WB.setActiveCanvas('container');\n    }\n}",
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
    "description": "<p>新增一个表: WB.addSheet(data)</p> <ul> <li>WB.addSheet(data)   新建一个表并且传入数据</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": true,
            "field": "data",
            "description": "<p>表的数据可选传入</p>"
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
    "title": "cellTypeContent",
    "name": "cellTypeContent",
    "group": "Function",
    "description": "<p>设置单元格类型 返回的是一个对象的形式 WB.cellTypeContent(n,list,i)</p> <ul> <li>WB.cellTypeContent(4,[&quot;a&quot;,&quot;b&quot;,&quot;c&quot;])</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "n",
            "description": "<ul> <li>0：无button</li> <li>1：button类型为三个小点；</li> <li>2：button类型为下标箭头；</li> <li>3：checkBox(为checkBox的情况，该单元格的值不画，且不可编辑)</li> <li>4: select下拉框</li> </ul>"
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
            "description": "<p>哪一项触发onopenmorelist事件</p>"
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
    "title": "clearAllBorder",
    "name": "clearAllBorder",
    "group": "Function",
    "description": "<p>清除范围边框(没参数默认清除当前范围) WB.clearAllBorder(R1,C1,R2,C2,Index)</p> <ul> <li>WB.clearAllBorder(2,2,5,3)  清除范围边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "clearAllStyle",
    "name": "clearAllStyle",
    "group": "Function",
    "description": "<p>恢复范围单元格默认样式样式(没参数默认设置当前选中范围的单元格) WB.clearAllStyle(R1,C1,R2,C2,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "description": "<p>清除格式刷样式 WB.clearFormatBrushStyle()</p>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "clearText",
    "name": "clearText",
    "group": "Function",
    "description": "<p>清空范围单元格文本(若有text跟value;text value设置为空)WB.deleteText(R1,C1,R2,C2,Index) 锁定/不可编辑/checkbox的单元格不能删除;(没参数默认清空当前范围)</p> <ul> <li>WB.clearText (1,0)  删除（1，0）单元格的文本</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "col",
    "name": "col",
    "group": "Function",
    "description": "<p>设置或者返回当前工作表的活动列 WB.col(column,index)</p> <ul> <li>WB.col()   获取当前活动列</li> <li>WB.col(2) \t设置当前活动列为第三列</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "column",
            "description": "<p>列(column&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "createData",
    "name": "createData",
    "group": "Function",
    "description": "<p>初始化一个空的单元格（一个单元格没值没有样式等是不会记录的，要在现有的行列范围内初始化）: WB.createData(r,c,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r",
            "description": "<p>行(r&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c",
            "description": "<p>列(c&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "deleteCell",
    "name": "deleteCell",
    "group": "Function",
    "description": "<p>删除单元格记录的数据(包括文本 样式 没参数默认删除当前范围) : WB.deleteCell(R1,C1,R2,C2,Index)</p> <ul> <li>WB.deleteCell (1,0)  删除（1，0）单元格</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "deleteFillColor",
    "name": "deleteFillColor",
    "group": "Function",
    "description": "<p>删除范围内单元格填充颜色(没参数默认清除范围内的填充色) WB.deleteFillColor(R1,C1,R2,C2,Index)</p> <ul> <li>WB.deleteFillColor(0,0,2,2)      除范围内单元格的填充色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "deleteHeaderNames",
    "name": "deleteHeaderNames",
    "group": "Function",
    "description": "<p>删除列行头name   WB.deleteHeaderNames(R,C,Index)</p> <ul> <li>WB.deleteHeaderNames(1,-1)  删除第二行行头name</li> <li>WB.deleteHeaderNames(-1,1)  删除第二列列头name</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>删除行头C传-1(同时删除列头行头name R&gt;=0 C&gt;=0)，R&gt;=0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>删除列头R传-1，C&gt;=0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "deleteSheets",
    "name": "deleteSheets",
    "group": "Function",
    "description": "<p>删除工作表: WB.deleteSheets（nSheet,nSheets）</p> <ul> <li>WB.deleteSheets(1,2)  删除第二第三个表（至少留有一个表）</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheet",
            "defaultValue": "当前索引",
            "description": "<p>开始的索引.</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheets",
            "defaultValue": "1",
            "description": "<p>数目,删除多少个表.</p>"
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
    "title": "deleteText",
    "name": "deleteText",
    "group": "Function",
    "description": "<p>删除范围单元格内容(若有text跟value属性;会删除该两个属性与clearText有点点不同) WB.deleteText(R1,C1,R2,C2,Index) 锁定/不可编辑/checkbox的单元格不能删除;(没参数默认删除当前范围)</p> <ul> <li>WB.deleteText (1,0)  删除（1,0）单元格的文本</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "drawPrint",
    "name": "drawPrint",
    "group": "Function",
    "description": "<p>打印 接收一个callback 返回一个数组,base64格式的页面图片</p>",
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
            "description": "<p>表索引</p>"
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
    "title": "editDelete",
    "name": "editDelete",
    "group": "Function",
    "description": "<p>删除当前或者指定的行列: WB.editDelete(nShiftType,R,C,N,Index)</p> <ul> <li>WB.editDelete  (&quot;F1ShiftRows&quot;,0,-1)   删除第一行</li> <li>WB.editDelete  (&quot;F1ShiftCols&quot;,-1,1)   删除第二列</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "nShiftType",
            "description": "<ul> <li>F1ShiftRows:(删除行)R</li> <li>F1ShiftCols(删除列)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>删除哪一行,R&gt;=0,删除列的时候R传 -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>删除哪一列,C&gt;=0,删除行的时候C传 -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "N",
            "defaultValue": "1",
            "description": "<p>N&gt;=1行数或者列数</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "editInsert",
    "name": "editInsert",
    "group": "Function",
    "description": "<p>在当前行列或者指定行列插入新的一行或者一列: WB.editInsert(InsertType,R,C,N,Index)</p> <ul> <li>WB.editInsert (&quot;F1ShiftRows&quot;,0,-1)  在第一行前面插入一行</li> <li>WB.editInsert (&quot;F1ShiftCols&quot;,-1,2)   在第三列前面插入一列</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "InsertType",
            "description": "<ul> <li>F1ShiftRows:(插入一行、所属行下移)</li> <li>F1ShiftCols(插入一列：所属列右移)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>在那一行前面插入,R&gt;=0,插入列的时候R传 -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>在那一列前面插入,C&gt;=0,插入行的时候C传 -1</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "N",
            "defaultValue": "1",
            "description": "<p>N&gt;=1行数或者列数</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "exportFileAsJson",
    "name": "exportFileAsJson",
    "group": "Function",
    "description": "<p>保存当前workbook为json文件并下载</p> <ul> <li>WB.exportFileAsJson()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "exportFileAsXml",
    "name": "exportFileAsXml",
    "group": "Function",
    "description": "<p>保存当前workbook为xml文件并下载(每个表生成一个xml)</p> <ul> <li>WB.exportFileAsXml()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "fixedFirstCol",
    "name": "fixedFirstCol",
    "group": "Function",
    "description": "<p>冻结活动表的可视区域的首列 WB.fixedFirstCol(startC,Index)</p> <ul> <li>WB.fixedFirstCol()  冻结活动表的可视区域的首列</li> <li>注：调用该函数，现有的行冻结会移除</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startC",
            "defaultValue": "0",
            "description": "<p>设定可视视图的开始列是哪一列</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "fixedFirstRow",
    "name": "fixedFirstRow",
    "group": "Function",
    "description": "<p>冻结活动表的可视区域的首行 WB.fixedFirstRow(startR,Index)</p> <ul> <li>WB.fixedFirstRow()  冻结活动表的可视区域的首行</li> <li>注：调用该函数，现有的列冻结会移除</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startR",
            "defaultValue": "0",
            "description": "<p>设定可视视图的开始行是哪一行</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "fixedRowAndCol",
    "name": "fixedRowAndCol",
    "group": "Function",
    "description": "<p>冻结拆分窗格 WB.fixedRowAndCol（R,C,startR,startC,Index）</p> <ul> <li>WB.fixedRowAndCol(1,1)   在单元格（1,1）处冻结拆分单元格 呈十字</li> <li>注：可视区域第一行为startRow   第一列为startCol</li> <li>fixedRows(冻结了多少行,哪一行开始冻结);fixedRow(起始行;fixedCols(冻结了多少列,哪一列开始冻结);fixedRow(起始列）;</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>在哪一行开始冻结（冻结线在该行下面）(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>在哪一列开始冻结(C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startR",
            "defaultValue": "0",
            "description": "<p>开始行(startR&gt;=0,注意当R&gt;0时startR&lt;R)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "startC",
            "defaultValue": "0",
            "description": "<p>开始列(startC&gt;=0,注意当C&gt;0时startC&lt;C)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "fontColor",
    "name": "fontColor",
    "group": "Function",
    "description": "<p>设置范围内单元格文本字体颜色 WB.fontColor(R1,C1,R2,C2,Color,Index)</p> <ul> <li>WB.fontColor(0,0,2,2,'pink')   设置范围内字体为粉色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R2",
            "description": "<p>结束行(r2&gt;=0,r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C2",
            "description": "<p>结束列(c2&gt;=0,c2&gt;=c1)</p>"
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
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "fontStrikeout",
    "name": "fontStrikeout",
    "group": "Function",
    "description": "<p>区域是否设置删除线(没参数默认设置当前选中范围的删除线) WB.fontStrikeout(R1,C1,R2,C2,bool,Index)</p> <ul> <li>WB.fontStrikeout(0,0,2,2,true)     设置区域内文本有删划线</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>是否设置删除线</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "fontUnderline",
    "name": "fontUnderline",
    "group": "Function",
    "description": "<p>区域是否设置下划线(没参数默认设置当前选中范围的下划线) WB.fontUnderline(R1,C1,R2,C2,bool,Index)</p> <ul> <li>WB.fontUnderline(0,0,2,2)     设置区域内文本有下划线</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>是否设置下划线</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getActiveSheet",
    "name": "getActiveSheet",
    "group": "Function",
    "description": "<p>设定或者获取当前表: WB.getActiveSheet(Index)</p> <ul> <li>WB.getActiveSheet () 获取当前表</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>索引（0开始）.</p>"
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
    "title": "getAlignment",
    "name": "getAlignment",
    "group": "Function",
    "description": "<p>获取选中单元格的对齐方式(没参数默认获取当前选中范围的对齐方式) WB.getAlignment(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getAlignment(0,0)   获取（0，0）单元格的对齐方式</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getBorder",
    "name": "getBorder",
    "group": "Function",
    "description": "<p>获取选中单元格边框样式(没参数默认获取当前选中范围的边框) WB.getBorder(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getBorder(0,0)  获取（0，0）单元格的边框样式</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getCanEdit",
    "name": "getCanEdit",
    "group": "Function",
    "description": "<p>获取单元格的是否可编辑(没参数默认获取当前选中范围的不可编辑状态) WB.getCanEdit(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getCanEdit(0,0)   获取（0，0）单元格的可编辑状态</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getCellFormat",
    "name": "getCellFormat",
    "group": "Function",
    "description": "<p>获取单元格格式(没参数默认获取当前选中范围的单元格格式) WB.getCellFormat(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getCellFormat(1,1)    获取单元格（1,1）格式</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getCellFormula",
    "name": "getCellFormula",
    "group": "Function",
    "description": "<p>获取单元格公式(没参数默认获取当前选中范围的公式) WB.getCellFormula(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getCellFormula(1,1)    获取指定单元格（1,1）格式</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getCellRC",
    "name": "getCellRC",
    "group": "Function",
    "description": "<p>获取当前表的r c 范围</p> <ul> <li>WB.getCellRC()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getCellType",
    "name": "getCellType",
    "group": "Function",
    "description": "<p>获取单元格的celltype(没参数默认获取当前选中范围的celltype) WB.getCellType(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getCellType(0,0)   获取（0，0）单元格的celltype</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getEndSheet",
    "name": "getEndSheet",
    "group": "Function",
    "description": "<p>返回结束表索引及名称 ： WB.getEndSheet()</p> <ul> <li>WB.getEndSheet ()  结束表索引及名称</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getFillColor",
    "name": "getFillColor",
    "group": "Function",
    "description": "<p>获取单元格的填充色(没参数默认获取当前选中范围的填充色) WB.getFillColor(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getFillColor(0,0)   获取（0，0）单元格的填充色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getFont",
    "name": "getFont",
    "group": "Function",
    "description": "<p>获取选中单元格的字体样式(没参数默认获取当前选中范围的字体) WB.getFont(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getFont(0,0)  获取（0，0）单元格的字体样式</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getFontColor",
    "name": "getFontColor",
    "group": "Function",
    "description": "<p>获取单元格的字体颜色(没参数默认获取当前选中范围的字体颜色) WB.getFontColor(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getFontColor(0,0)   获取（0，0）单元格的字体颜色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getFontList",
    "name": "getFontList",
    "group": "Function",
    "description": "<p>获取字体的每一项 WB.getFontList(r,c,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r",
            "description": "<p>行(r&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c",
            "description": "<p>列(c&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getFontStrikeout",
    "name": "getFontStrikeout",
    "group": "Function",
    "description": "<p>获取单元格的是否有删除线(没参数默认获取当前选中范围的删除线状态) WB.getFontStrikeout(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getFontStrikeout(0,0)   获取（0，0）单元格的删除线状态</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getFontUnderline",
    "name": "getFontUnderline",
    "group": "Function",
    "description": "<p>获取单元格的是否有下划线(没参数默认获取当前选中范围的下划线状态) WB.getFontUnderline(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getFontUnderline(0,0)   获取（0，0）单元格的下划线状态</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getLock",
    "name": "getLock",
    "group": "Function",
    "description": "<p>获取单元格的是否锁定(没参数默认获取当前选中范围的锁定状态) WB.getLock(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getLock(0,0)   获取（0，0）单元格的锁定状态</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getSelectInfo",
    "name": "getSelectInfo",
    "group": "Function",
    "description": "<p>获取选中单元格信息(没参数默认获取当前选中范围信息): WB.getSelectInfo(R1,C1,R2,C2,Index) 调用后返回单元格坐标宽高</p> <ul> <li>WB.getSelectInfo () 获取选中单元格信息(返回的w h是所看到的宽高)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getStartAndEnd",
    "name": "getStartAndEnd",
    "group": "Function",
    "description": "<p>获取视图的开始行、开始列、结束行、结束列: WB.getStartAndEnd(index)</p> <ul> <li>WB.getStartAndEnd() 开始结束的行列信息</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getStartSheet",
    "name": "getStartSheet",
    "group": "Function",
    "description": "<p>返回开始表表索引及名称 ： WB.getStartSheet()</p> <ul> <li>tabs栏的表格有限的width下显示多少个表 start-&gt;end</li> <li>WB.getStartSheet ()  开始表索引及名称</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "getText",
    "name": "getText",
    "group": "Function",
    "description": "<p>获取激活单元格文本,没有text则获取value(没参数默认获取当前选中范围的值) WB.getText(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getText(0,0)   获取（0，0）单元格的文本</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getValue",
    "name": "getValue",
    "group": "Function",
    "description": "<p>获取单元格value值(没参数默认获取当前选中范围单元格的value) WB.getValue(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getValue(0,0)   获取（0，0）单元格的value</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "getWordWrap",
    "name": "getWordWrap",
    "group": "Function",
    "description": "<p>获取单元格的是否自动换行(没参数默认获取当前选中范围的自动换行状态) WB.getWordWrap(R1,C1,R2,C2,Index)</p> <ul> <li>WB.getWordWrap(0,0)   获取（0，0）单元格的自动换行状态</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "hdrHeight",
    "name": "hdrHeight",
    "group": "Function",
    "description": "<p>设置列标题栏高度 WB.hdrHeight(height,index)</p> <ul> <li>WB.hdrHeight(40)  设置列头高度为40px</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "height",
            "description": "<p>如果height&lt;=0  会隐藏列头并且高度为0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "hdrWidth",
    "name": "hdrWidth",
    "group": "Function",
    "description": "<p>设置行标题栏宽度 WB.hdrWidth(width,index)</p> <ul> <li>WB.hdrWidth(40)  设置行头宽度为40px</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "width",
            "description": "<p>width&lt;=0  会隐藏行头并且宽度为0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "hiddenChildCanvas",
    "name": "hiddenChildCanvas",
    "group": "Function",
    "description": "<p>隐藏子canvas全部元素</p> <ul> <li>WB2.hiddenChildCanvas(&quot;container2&quot;)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": "<p>子canvas容器id</p>"
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
    "description": "<p>文件导入(文件仅支持.json、.xml、.xlsx、excel.zip)</p> <ul> <li>WB.exportFileAsXml()  f</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "f",
            "description": "<p>文件</p>"
          },
          {
            "group": "Parameter",
            "type": "Function",
            "optional": false,
            "field": "callback",
            "description": "<p>导入之后的回到函数</p>"
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
    "description": "<p>当前索引的表的前面插入n个表: WB.insertSheets（nSheet,nSheets）</p> <ul> <li>WB.insertSheets(1,1)    在第二个表前面插入一个表</li> <li>WB.insertSheets()    当前表前面插入一个表</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheet",
            "defaultValue": "当前索引",
            "description": "<p>开始的索引.</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nSheets",
            "defaultValue": "1",
            "description": "<p>数目,插入多少个表 nSheets&gt;=1.</p>"
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
    "title": "maxCol",
    "name": "maxCol",
    "group": "Function",
    "description": "<p>获取或设定当前工作表的列数:  WB.maxCol(num,Index)</p> <ul> <li>WB.maxCol(10)         总列数为10列</li> <li>WB.maxCol()        获取总列数</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "num",
            "defaultValue": "当前表列数",
            "description": "<ul> <li>num&gt;len（当前表的总列数），传入num大于当前列数  则增加列数.</li> <li>0&lt;=num&amp;&amp;num&lt; len，传入num小于当前并且大于等于0  则删除列数（对应合并，单元格数据清除）对应合并减少(num&lt;0,num=0)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "maxRow",
    "name": "maxRow",
    "group": "Function",
    "description": "<p>获取或设定当前工作表的列数:  WB.maxRow(num,Index)</p> <ul> <li>WB.maxRow(9)         总行数为9行</li> <li>WB.maxRow()        获取总行数</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "num",
            "defaultValue": "当前表行数",
            "description": "<ul> <li>num&gt;len（当前表的总行数），传入num大于当前行数  则增加行数.</li> <li>0&lt;=num&amp;&amp;num&lt; len，传入num小于当前并且大于等于0  则删除行数（对应合并，单元格数据清除）对应合并减少(num&lt;0,num=0)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "mergeCells",
    "name": "mergeCells",
    "group": "Function",
    "description": "<p>合并单元格(没参数默认设置当前选中范围的合并) WB.mergeCells(R1,C1,R2,C2,Index)</p> <ul> <li>WB.mergeCells(2,1,3,2)            单元格(2,1)(2,2)(3,1)(3,2)被合并</li> <li>合并会改变现有的合并数据，合并之后也只会保留头一个单元格的内容</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "newWorkbook",
    "name": "newWorkbook",
    "group": "Function",
    "description": "<p>新建一个workbook 当前workbook数据清空</p> <ul> <li>WB.newWorkbook()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": true,
            "field": "template",
            "description": "<p>可传入模板</p>"
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
    "description": "<p>设置工作簿中工作表数: WB.numSheets(num)</p> <ul> <li>WB.numSheets(3)   表的总数设置为3个</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "num",
            "description": "<ul> <li>num&gt;全部表.length(在最后表之后插入表，如length=2  num=5  则添加三个表)；</li> <li>num&lt;全部表.length&amp;&amp;num&gt;=1(从后面删除表，如length=5  num=2  则删除最后三个表)</li> <li>num&lt;=0(不会删除   至少留有一张表)</li> </ul>"
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
    "title": "removeFixed",
    "name": "removeFixed",
    "group": "Function",
    "description": "<p>取消冻结窗格 WB.removeFixed(Index)</p> <ul> <li>WB.removeFixed ()   移除所有冻结</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "removeMergeCells",
    "name": "removeMergeCells",
    "group": "Function",
    "description": "<p>取消合并单元格(没参数默认取消当前范围的合并) WB.removeMergeCells(R1,C1,R2,C2,Index)</p> <ul> <li>WB.removeMergeCells(2,1,3,2)           取消合并</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "removeRowColor",
    "name": "removeRowColor",
    "group": "Function",
    "description": "<p>取消隔行变色  WB.removeRowColor(Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "removeSplitcolHeader",
    "name": "removeSplitcolHeader",
    "group": "Function",
    "description": "<p>删除拆分的列头  WB.removeSplitcolHeader(Col,index)</p> <ul> <li>WB.removeSplitcolHeader(4)    删除第五列的拆分列头</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "Col",
            "description": "<p>列(Col&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "resumeEvent",
    "name": "resumeEvent",
    "group": "Function",
    "description": "<p>resumeEvent函数设置stopEventCount的值，每调用一次该值减1，直到为0，stopEventCount为0时 on事件才会执行</p> <ul> <li>WB.resumeEvent()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "description": "<p>布尔值：true stopEventCount直接赋值为0</p>"
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
    "title": "row",
    "name": "row",
    "group": "Function",
    "description": "<p>设置或者返回当前工作表的活动行 WB.row(row,index)</p> <ul> <li>WB.row()   获取当前活动行</li> <li>WB.row(2) \t设置当前活动列为第三行</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "row",
            "description": "<p>行(row&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setActiveCanvas",
    "name": "setActiveCanvas",
    "group": "Function",
    "description": "<p>同页面多canvas的情况下设置活动的单元格（默认是new出来的最后一个canvas，点击对应的canvas activecanvas会自动切换）</p> <ul> <li>WB.setActiveCanvas(&quot;container2&quot;)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": "<p>容器id</p>"
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
    "title": "setActiveEditCell",
    "name": "setActiveEditCell",
    "group": "Function",
    "description": "<p>打开指定编辑单元格 WB.setActiveEditCell(R,C)</p> <ul> <li>WB.setActiveEditCell(1,1)     激活(1 1)单元格</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>行(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>列(C&gt;=0)</p>"
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
    "title": "setAllBorder",
    "name": "setAllBorder",
    "group": "Function",
    "description": "<p>设置范围内全边框(没参数默认设置当前选中范围的全边框)  WB.setAllBorder(R1,C1,R2,C2,NAll,CrALL,Index)</p> <ul> <li>WB.setAllBorder(2,2,5,3)  设置范围内全边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setBottomBorder",
    "name": "setBottomBorder",
    "group": "Function",
    "description": "<p>设置范围内下边框(没参数默认设置当前选中范围的下边框)  WB.setBottomBorder(R1,C1,R2,C2,NBottom,CrBottom,Index)</p> <ul> <li>WB.setBottomBorder(1,1,3,3)  设置范围内下边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>下边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setCanEdit",
    "name": "setCanEdit",
    "group": "Function",
    "description": "<p>设置单元格是否可以编辑(没参数默认设置当前选中范围不能编辑) WB.setCanEdit(R1,C1,R2,C2,bool,Index)</p> <ul> <li>WB.setCanEdit(1,1,1,1,false)    单元格1，1不能编辑</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "defaultValue": "false",
            "description": "<p>默认单元格不可以编辑</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setCellFormat",
    "name": "setCellFormat",
    "group": "Function",
    "description": "<p>设置范围内单元格格式 WB.setCellFormat(R1,C1,R2,C2,type,Index)</p> <ul> <li>WB.setCellFormat(0,0,2,c2,&quot;yyyy年m月d日&quot;)    设定范围内单元格的格式为日期(&quot;yyyy年m月d日&quot;格式)</li> <li>日期：yyyy年m月d日 ,yyyy年m月 ,m月d日,yyyy-m-d ,yyyy-m,m-d ,yyyy-m-d:uppercase ,yyyy-m:uppercase ,星期n ,yyyy/m/d ,yyyy/m,m/d;</li> <li>科学计数法：E  (0E代表保留一位小数,00E代表两位小数,以此类推,E则不保留小数);</li> <li>数值：NR(A+thousands  (N的前面有多少个0就保留多少位小数，N后面有+thousands则使用千位分隔符,无则不使用千位分隔符) R(A同理货币</li> <li>货币：MR(A+$   (M的前面有多少个0就保留多少位小数,对于负数有R则设置颜色位红色，有(代表加上括号（仅有A的时候），有A代表转成正数形式显示,R(A三个字符位置不限，但要放在M之后,M后面有+货币符号则在前面添加货币符号显示,无则不使用货币符号 ¥|$|€|￡|₣|₩ （人民币（日元显示差不多也是¥）、美元、欧元、英镑、法郎、韩元）</li> <li>百分比：P+%  (P的前面有多少个0就保留多少位小数)</li> <li>金额大写：capitalMoney(数字转大写金额,小数保留到角分)</li> <li>文本：text 文本格式(原样返回,包括公式)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "type",
            "description": "<p>有效的单元格格式值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setCellFormula",
    "name": "setCellFormula",
    "group": "Function",
    "description": "<p>设置单元格公式 WB.setCellFormula(R1,C1,Formula,Index)</p> <ul> <li>WB.setCellFormula(0,0,'=A1+B1')   设置单元格公式为=A1+B1</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Formula",
            "description": "<p>(必须以=开头且后面最少有一个字符)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setCellPadding",
    "name": "setCellPadding",
    "group": "Function",
    "description": "<p>设置表单元格的padding  WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index)</p> <ul> <li>WB.setCellPadding(0,0,0,5)    设置单元格右边框为5px</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "topSize",
            "description": "<p>padding-top值的大小</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "rightSize",
            "description": "<p>padding-right值的大小</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "bottomSize",
            "description": "<p>padding-bottom值的大小</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "leftSize",
            "description": "<p>padding-left值的大小</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setCellType",
    "name": "setCellType",
    "group": "Function",
    "description": "<p>设置范围内单元格cellType 传入的obj请根据var obj = WB.cellTypeContent(4,[&quot;apple&quot;,&quot;banana&quot;])传入 WB.setCellType(R1,C1,R2,C2,obj,Index)或者直接传一个对象</p> <ul> <li>WB.setCellType(0,0,0,0,obj)    设置第一个单元格的cellType为select,list为[&quot;apple&quot;,&quot;banana&quot;]</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setColHidden",
    "name": "setColHidden",
    "group": "Function",
    "description": "<p>设置列隐藏，如果鼠标拖拽列宽为&lt;=0也会隐藏 WB.setColHidden(C1,C2,BHidden,Index)</p> <ul> <li>WB.setColHidden(1,2,true)  隐藏第二第三列</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "BHidden",
            "description": "<p>是否隐藏</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setColWidth",
    "name": "setColWidth",
    "group": "Function",
    "description": "<p>输入设置范围行高度(隐藏的列) : WB.setColWidth(width,c1,c2,index)</p> <ul> <li>WB.setColWidth   (80,0,0)  第一列设置列的宽度为80</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Numer",
            "optional": false,
            "field": "width",
            "description": "<p>输入一个数值(width&lt;=0,隐藏)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "c2",
            "description": "<p>结束列(c1&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setDefaultFontName",
    "name": "setDefaultFontName",
    "group": "Function",
    "description": "<p>设置工作簿的默认字体</p> <ul> <li>WB.setDefaultFontName('楷体')</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "string",
            "optional": false,
            "field": "fontname",
            "description": "<p>字体名称</p>"
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
    "title": "setDefaultFontSize",
    "name": "setDefaultFontSize",
    "group": "Function",
    "description": "<p>设置工作簿的默认字体大小</p> <ul> <li>WB.setDefaultFontSize(16)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "fontsize",
            "description": "<p>字体大小</p>"
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
    "title": "setDefaultLineHeight",
    "name": "setDefaultLineHeight",
    "group": "Function",
    "description": "<p>设置工作簿的默认字体行高</p> <ul> <li>WB.setDefaultLineHeight(20)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "int",
            "optional": false,
            "field": "lineheight",
            "description": "<p>字体行高</p>"
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
    "title": "setDocument",
    "name": "setDocument",
    "group": "Function",
    "description": "<p>设置触发点击事件的document  WB.setDocument(D)</p> <ul> <li>WB.setDocument(window.top.document)   设置为顶层</li> </ul>",
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
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "setFillColor",
    "name": "setFillColor",
    "group": "Function",
    "description": "<p>设置范围内单元格填充颜色(没rc参数默认设置当前选中范围的填充色) WB.setFillColor(color,R1,C1,R2,C2,Index)</p> <ul> <li>WB.setFillColor('pink',0,0,2,2)   设置范围内单元格的填充色为粉色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "color",
            "description": "<p>有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFontBold",
    "name": "setFontBold",
    "group": "Function",
    "description": "<p>设置范围内字体是否粗体(没参数默认设置当前选中范围的粗体)  WB.setFontBold(R1,C1,R2,C2,BBold,Index)</p> <ul> <li>WB.setFontBold(1,1,1,1,true)  设置范围内字体为粗体</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "BBold",
            "defaultValue": "true",
            "description": "<p>字是否粗体</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFontItalic",
    "name": "setFontItalic",
    "group": "Function",
    "description": "<p>设置范围内字体是否斜体(没参数默认设置当前选中范围的斜体)  WB.setFontItalic(R1,C1,R2,C2,BItalic,Index)</p> <ul> <li>WB.setFontItalic(1,1,1,1,true)  设置范围内字体为斜体</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "BItalic",
            "defaultValue": "true",
            "description": "<p>字是否斜体</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFontLineH",
    "name": "setFontLineH",
    "group": "Function",
    "description": "<p>设置范围内字体行高 WB.setFontLineH(R1,C1,R2,C2,NLineHeight,Index)</p> <ul> <li>WB.setFontLineH(1,1,1,1,20)  设置范围内字体行高为20px</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r2",
            "description": "<p>结束行(r2&gt;=0,r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c2",
            "description": "<p>结束列(c2&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "NLineHeight",
            "description": "<p>字体行高</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFontName",
    "name": "setFontName",
    "group": "Function",
    "description": "<p>设置范围内字体font-family WB.setFontName(R1,C1,R2,C2,PName,Index)</p> <ul> <li>WB.setFontName(0,0,2,2,'宋体')  设置范围内font-family</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r2",
            "description": "<p>结束行(r2&gt;=0,r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c2",
            "description": "<p>结束列(c2&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "PName",
            "description": "<p>字体</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFontSize",
    "name": "setFontSize",
    "group": "Function",
    "description": "<p>设置范围内字体大小 WB.setFontSize(R1,C1,R2,C2,NSize,Index)</p> <ul> <li>WB.setFontSize(1,1,1,1,20)  设置范围内字体大小为20px</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "NSize",
            "description": "<p>字体大小</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setFormatBrushStyle",
    "name": "setFormatBrushStyle",
    "group": "Function",
    "description": "<p>设置格式刷样式 WB.setFormatBrushStyle(R,C,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C1",
            "description": "<p>列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setHAlignment",
    "name": "setHAlignment",
    "group": "Function",
    "description": "<p>设置范围内单元格文本的水平对齐方式 WB.setHAlignment(R1,C1,R2,C2,HAlign,Index)</p> <ul> <li>WB.setHAlignment(0,0,0,0,2) 设置范围内单元格为左对齐</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r2",
            "description": "<p>结束行(r2&gt;=0 r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c2",
            "description": "<p>结束列(c2&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "hAlign",
            "defaultValue": "1",
            "description": "<ul> <li>1: 常规（根据单元格数据类型对齐）</li> <li>2：左对齐</li> <li>3：水平对齐</li> <li>4：右对齐</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setHeaderName",
    "name": "setHeaderName",
    "group": "Function",
    "description": "<p>改变列头行头的名称 WB.setHeaderName(R,C,headerName,Index)</p> <ul> <li>WB.setHeaderName(1,-1,'行头名')     第二行行头为 ‘行头名’</li> <li>WB.setHeaderName(-1,1,'列头名')     第二列列头为 ‘列头名’</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>改变行头,改变列头时R传-1 r&gt;=0</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>改变列头,改变行头时C传-1 c&gt;=0</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "setHeaderName",
            "description": "<p>名称</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setINnerBorder",
    "name": "setINnerBorder",
    "group": "Function",
    "description": "<p>设置范围内内边框(没参数默认设置当前选中范围的内边框) WB.setINnerBorder(R1,C1,R2,C2,NInner,CrInner,Index)</p> <ul> <li>WB.setINnerBorder(1,1,5,5)  设置范围内内边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setLeftBorder",
    "name": "setLeftBorder",
    "group": "Function",
    "description": "<p>设置范围内左边框(没参数默认设置当前选中范围的左边框) WB.setLeftBorder(R1,C1,R2,C2,NLeft,CrLeft,Index)</p> <ul> <li>WB.setLeftBorder(1,1,3,3)  设置范围内左边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "nLeft",
            "defaultValue": "1",
            "description": "<ul> <li>1:\tThin Line（细）</li> <li>2: Medium Line（中）</li> <li>5: Thick Line(粗)</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "crLeft",
            "defaultValue": "#000",
            "description": "<p>左边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setLock",
    "name": "setLock",
    "group": "Function",
    "description": "<p>设置范围内单元格是否锁定(没参数默认设置当前选中范围的单元格锁定,锁定的单元格不能编辑 没选中框 可选中) WB.setLock(R1,C1,R2,C2,bool,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "defaultValue": "true",
            "description": "<p>布尔值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setOutBorder",
    "name": "setOutBorder",
    "group": "Function",
    "description": "<p>设置范围内外边框(没参数默认设置当前选中范围的外边框) WB.setOutBorder(R1,C1,R2,C2,NOut,CrOut,Index)</p> <ul> <li>WB.setOutBorder(1,1,5,5)  设置范围内外边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setPrint",
    "name": "setPrint",
    "group": "Function",
    "description": "<p>打印设置参数obj 不传的项则使用已有设置</p> <ul> <li>marginTop 打印上边距</li> <li>marginBottom 打印下边距</li> <li>marginLeft 打印左边距</li> <li>marginRight 打印右边距</li> <li>paper 纸张类型</li> <li>printHeadings 是否打印行号列标</li> <li>printGridLine 是否打印网格线</li> <li>orientation 纸张方向</li> <li>printDirection 打印方向(1:先列后行呈N字形 2：先行后列呈Z字形)</li> <li>marginCopies 数据间隔 同张纸张两条数据(表)之间的间隔</li> <li>printer 打印机</li> <li>isshowfooterpageinfo 显示页码信息参数&quot;0&quot;, &quot;1&quot;, &quot;2&quot;, &quot;3&quot;, &quot;4&quot;, &quot;5&quot;, &quot;6&quot;对应:&quot;不显示&quot;, &quot;页脚居中&quot;, &quot;页脚居左&quot;, &quot;页脚居右&quot;, &quot;页眉居中&quot;, &quot;页眉居左&quot;, &quot;页眉居右</li> <li>footpagestyle 页码格式</li> <li>startR 打印区域开始row</li> <li>endR 打印区域结束row</li> <li>startC 打印区域开始col</li> <li>endC 打印区域结束col</li> <li>print 1打印当前工作表(默认) 2打印工作簿 3打印当前表当前选中区域</li> <li>printOnSamePaper 仅当print为2时生效 在纸张可以放下其他表的情况下 多张表是否放在同一张纸上 默认false</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Object",
            "optional": false,
            "field": "obj",
            "description": "<p>设置</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setRightBorder",
    "name": "setRightBorder",
    "group": "Function",
    "description": "<p>设置范围内右边框(没参数默认设置当前选中范围的右边框) WB.setRightBorder(R1,C1,R2,C2,NRight,CrRight,Index)</p> <ul> <li>WB.setRightBorder(1,1,3,3)  设置范围内右边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>右边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setRowColor",
    "name": "setRowColor",
    "group": "Function",
    "description": "<p>设置行的颜色（隔行变色,其实就相当于给单元格加填充)  WB.setRowColor(color,oddEven,start,Index)</p> <ul> <li>WB.setRowColor('rgb(255,255,224)',0)    开启隔行变色</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "color",
            "defaultValue": "rgb(255,255,224)",
            "description": "<p>有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "oddEven",
            "defaultValue": "0",
            "description": "<ul> <li>0 : 偶数行</li> <li>1 ：奇数行</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "start",
            "defaultValue": "0",
            "description": "<p>从哪里开始;</p> <ul> <li>0 ：从表格顶部开始</li> <li>1 ：从冻结行以下部分开始</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setRowHeight",
    "name": "setRowHeight",
    "group": "Function",
    "description": "<p>输入设置范围行高度(隐藏的行不能设置) : WB.setRowHeight(height,R1,R2,Index)</p> <ul> <li>WB.setRowHeight   (80,0,1)  第一二行设置行的高度为80</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Numer",
            "optional": false,
            "field": "height",
            "description": "<p>输入一个数值(height&lt;=0,隐藏)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setRowHidden",
    "name": "setRowHidden",
    "group": "Function",
    "description": "<p>设置行隐藏，如果鼠标拖拽行高为&lt;=0也会隐藏 WB.setRowHidden(R1,R2,BHidden,Index)</p> <ul> <li>WB.setRowHidden(1,2,true)  隐藏第二第三行</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R2",
            "description": "<p>结束行(r2&gt;=0,r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": false,
            "field": "BHidden",
            "description": "<p>是否隐藏</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setScrollPosition",
    "name": "setScrollPosition",
    "group": "Function",
    "description": "<p>设置滚动条的位置注意：有冻结的情况 这个函数要放在冻结函数之后 WB.setScrollPosition(x,y,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "x",
            "description": "<p>水平数值</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "y",
            "description": "<p>垂直数值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setTableSize",
    "name": "setTableSize",
    "group": "Function",
    "description": "<p>设置canvas以及容器的宽高  WB.setTableSize(width,height)</p> <ul> <li>WB.setTableSize(500,300)    设置canvas宽高（500，300）</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "width",
            "description": "<p>宽(width&gt;0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": false,
            "field": "height",
            "description": "<p>高(height&gt;0)</p>"
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
    "title": "setText",
    "name": "setText",
    "group": "Function",
    "description": "<p>设置单元格值(text)  WB.setText(R,C,Text,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>行(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>列(C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Text",
            "description": "<p>值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setTopBorder",
    "name": "setTopBorder",
    "group": "Function",
    "description": "<p>设置范围内上边框(没参数默认设置当前选中范围的上边框) WB.setTopBorder(R1,C1,R2,C2,NTop,CrTop,Index)</p> <ul> <li>WB.setTopBorder(1,1,3,3)  设置范围内上边框</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
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
            "description": "<p>上边框颜色，参数有效颜色值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setVAlignment",
    "name": "setVAlignment",
    "group": "Function",
    "description": "<p>设置范围内单元格文本的垂直对齐方式 WB.setVAlignment(R1,C1,R2,C2,VAlign,Index)</p> <ul> <li>WB.setVAlignment(0,0,0,0,1) 设置范围内单元格为顶部对齐</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r1",
            "description": "<p>开始行(r1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c1",
            "description": "<p>开始列(c1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "r2",
            "description": "<p>结束行(r2&gt;=0,r2&gt;=r1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "c2",
            "description": "<p>结束列(c2&gt;=0,c2&gt;=c1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "vAlign",
            "defaultValue": "2",
            "description": "<ul> <li>1：顶部对齐</li> <li>2：垂直居中对齐</li> <li>3：底部对齐</li> </ul>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "setValue",
    "name": "setValue",
    "group": "Function",
    "description": "<p>设置单元格value值(value text)  WB.setValue(R,C,Value,Index)</p> <ul> <li>WB.setValue(0,0,'统计',1)    设置第二个表（0，0）单元格的值value值为 统计</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "R",
            "description": "<p>行(R&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "C",
            "description": "<p>列(C&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Value",
            "description": "<p>值</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "sheetIndex",
    "name": "sheetIndex",
    "group": "Function",
    "description": "<p>设定或获取当前表索引: WB.sheetIndex(Index)</p> <ul> <li>WB.sheetIndex(1)   切换到tab栏的第二个表</li> <li>WB.sheetIndex()    获取当前表索引</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前索引",
            "description": "<p>表索引&gt;=0&amp;&amp;&lt;全部表的length.</p>"
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
    "title": "sheetName",
    "name": "sheetName",
    "group": "Function",
    "description": "<p>设定或者获取当前工作表名: sheetName（Index，Name）</p> <ul> <li>WB.sheetName(1,&quot;表二&quot;)   把第二个表的名字改成‘表二’</li> <li>WB.sheetName()   获取当前表名称</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前索引",
            "description": "<p>表索引&gt;=0&amp;&amp;&lt;全部表的length.</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": true,
            "field": "Name",
            "defaultValue": "当前表名",
            "description": "<p>表名.</p>"
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
    "title": "showChildCanvas",
    "name": "showChildCanvas",
    "group": "Function",
    "description": "<p>显示子canvas全部元素</p> <ul> <li>WB.showChildCanvas(&quot;container2&quot;)</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "id",
            "description": "<p>子canvas容器id</p>"
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
    "title": "showColHeading",
    "name": "showColHeading",
    "group": "Function",
    "description": "<p>是否显示列标题 WB.showColHeading(boolean,index)</p> <ul> <li>WB.showColHeading(false) 不显示列标题</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>是否显示列标题</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "showFixedLine",
    "name": "showFixedLine",
    "group": "Function",
    "description": "<p>是否显示当前表的冻结线 WB.showFixedLine(boolean,index)</p> <ul> <li>WB.showFixedLine(true)   显示冻结线</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "false",
            "description": "<p>是否显示</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "showRowHeading",
    "name": "showRowHeading",
    "group": "Function",
    "description": "<p>是否显示行标题 WB.showRowHeading(boolean,index)</p> <ul> <li>WB.showRowHeading(false) 不显示行标题</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>是否显示行标题</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "splitcolHeader",
    "name": "splitcolHeader",
    "group": "Function",
    "description": "<p>拆分列头、只支持拆分两行  WB.splitcolHeader(Col,Name,Count,Height,index)</p> <ul> <li>WB.splitcolHeader(4,&quot;名称&quot;)    设置第五列拆分列头名称为 &quot;名称&quot;,默认拆分两列，高度为列头高度的一半；</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": false,
            "field": "Col",
            "description": "<p>列(Col&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "String",
            "optional": false,
            "field": "Name",
            "description": "<p>名称</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Count",
            "defaultValue": "2",
            "description": "<p>拆分的列数</p>"
          },
          {
            "group": "Parameter",
            "type": "Number",
            "optional": true,
            "field": "Height",
            "defaultValue": "列头的一半高度",
            "description": "<p>Height&lt;该列头的高度（defaultH）</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "startPaint",
    "name": "startPaint",
    "group": "Function",
    "description": "<p>startPaint函数设置stopPaintedCount的值，每调用一次该值减1，直到为0，stopPaintedCount为0时重绘：startPaint(isAll)</p> <ul> <li>WB.startPaint()</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "bool",
            "description": "<p>布尔值：true stopPaintedCount直接赋值为0</p>"
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
    "description": "<p>stopEvent函数设置stopEventCount的值，每调用一次该值加1.</p> <ul> <li>WB.stopEvent()</li> </ul>",
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
    "description": "<p>stopPaint函数设置stopPaintCount的值，每调用一次该值加1.</p> <ul> <li>WB.stopPaint()</li> </ul>",
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  },
  {
    "type": "null",
    "url": "/null",
    "title": "useFormatBrushStyle",
    "name": "useFormatBrushStyle",
    "group": "Function",
    "description": "<p>单元格范围应用格式刷样式(没参数默认设置当前选中范围的单元格)  WB.useFormatBrushStyle(R1,C1,R2,C2,Index)</p>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
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
    "title": "wordWrap",
    "name": "wordWrap",
    "group": "Function",
    "description": "<p>设置为范围内单元格是否自动换行(没参数默认设置当前选中范围的自动换行) : WB.wordWrap(R1,C1,R2,C2,bool,Index)</p> <ul> <li>WB.wordWrap (0,0,1,1,true)    单元格（0,0）(0,1)(1,0)(1,1)  四个变成自动换行</li> </ul>",
    "parameter": {
      "fields": {
        "Parameter": [
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R1",
            "description": "<p>开始行(R1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C1",
            "description": "<p>开始列(C1&gt;=0)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "R2",
            "description": "<p>结束行(R2&gt;=0,R2&gt;=R1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "C2",
            "description": "<p>结束列(C2&gt;=0,C2&gt;=C1)</p>"
          },
          {
            "group": "Parameter",
            "type": "Boolean",
            "optional": true,
            "field": "boolean",
            "defaultValue": "true",
            "description": "<p>是否自动换行</p>"
          },
          {
            "group": "Parameter",
            "type": "Int",
            "optional": true,
            "field": "Index",
            "defaultValue": "当前表索引",
            "description": "<p>表索引</p>"
          }
        ]
      }
    },
    "version": "0.0.0",
    "filename": "src/js/table.js",
    "groupTitle": "Function"
  }
] });
