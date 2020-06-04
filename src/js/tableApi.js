//workbook

/**
 * @api {null} /null allowDelete
 * @apiName allowDelete
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to allow the delete key to delete the selected text(Defaults:true)
 * - WB.workbook.allowDelete = value(boolean)   
 */

/**
 * @api {null} /null allowResize
 * @apiName allowResize
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to adjust the column width and height by dragging(Defaults:true)
 * - WB.workbook.allowResize = value(boolean)   
 */

/**
 * @api {null} /null allowTabs
 * @apiName allowTabs
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to use tab key to switch cells(Defaults:true)
 * - WB.workbook.allowTabs = value(boolean)   
 */

 /**
 * @api {null} /null showTabs
 * @apiName showTabs
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to show the tabs column
 * - 0:Do not show(Defaults)
 * - 1:show
 * - 2:Button group
 * - WB.workbook.showTabs = value(Int)  
 */


/**
 * @api {null} /null width
 * @apiName width
 * @apiGroup Attribute:workbook
 * @apiDescription Workbook width(Defaults:Viewport width minus system scroll bar size)
 * - WB.workbook.width = value(number)  
 */

/**size
 * @api {null} /null height
 * @apiName height
 * @apiGroup Attribute:workbook
 * @apiDescription Workbook height(Defaults:Viewport height minus system scroll bar size)
 * - WB.workbook.height = value(number)  
 */

/**
 * @api {null} /null showHScrollBar
 * @apiName showHScrollBar
 * @apiGroup Attribute:workbook
 * @apiDescription How to display the horizontal scroll bar of the workbook(Defaults:1)
 * - 0:Do not show
 * - 1:Displayed when the workbook is wide(Defaults)
 * - WB.workbook.showHScrollBar = value  
 */

/**
 * @api {null} /null showVScrollBar
 * @apiName showVScrollBar
 * @apiGroup Attribute:workbook
 * @apiDescription How to display the vertical scroll bar of the workbook
 * - 0:Do not show
 * - 1:Displayed when the workbook is high(Defaults)
 * - WB.workbook.showVScrollBar = value
 */

/**
 * @api {null} /null behaviorMode
 * @apiName behaviorMode
 * @apiGroup Attribute:workbook
 * @apiDescription Workbook mode
 * - 1:sheet mode (Double click to edit)(Defaults)
 * - 2:grid mode (Click cell to enter edit box but no check box)
 * - 3:sheet click mode (Click to enter edit and check box)
 * - WB.workbook.behaviorMode = value
 */

/**
 * @api {null} /null scrollMode
 * @apiName scrollMode
 * @apiGroup Attribute:workbook
 * @apiDescription Scroll bar mode
 * - 1:The scroll bar starts from the top or left
 * - 2:The scroll bar is below or to the right of the freeze line(Defaults)
 * - WB.workbook.scrollMode = value
 */

/**
 * @api {null} /null showWorkBookBorder
 * @apiName showWorkBookBorder
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to display the border of the workbook(Defaults:true)
 * - WB.workbook.showWorkBookBorder = value(boolean)  
 */

 /**
 * @api {null} /null showRowArrow
 * @apiName showRowArrow
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to display the leftmost arrow when the current row is selected(Defaults:false)
 * - WB.workbook.showRowArrow = value(boolean)  
 */

  /**
 * @api {null} /null rowArrowColor
 * @apiName rowArrowColor
 * @apiGroup Attribute:workbook
 * @apiDescription The fill color of the arrow of the current line(Defaults:#000)
 * - WB.workbook.rowArrowColor = value(color)  
 */

/**
 * @api {null} /null workBookBorderColor
 * @apiName workBookBorderColor
 * @apiGroup Attribute:workbook
 * @apiDescription Set the border color of the workbook(Defaults:#ccc)
 * - WB.workbook.workBookBorderColor = value(color)  
 */

/**
 * @api {null} /null tabsOptions.fontColor
 * @apiName tabsOptions.fontColor
 * @apiGroup Attribute:workbook
 * @apiDescription Tab bar font color(Defaults:#444)
 * - WB.workbook.tabsOptions.fontColor = value(color)  
 */

/**
 * @api {null} /null tabsOptions.font
 * @apiName tabsOptions.font
 * @apiGroup Attribute:workbook
 * @apiDescription Tab bar font(Defaults:14px/20px 宋体)
 * - WB.workbook.tabsOptions.font = value(Font size/line height font type)  
 */

/**
 * @api {null} /null tabsOptions.fontSelColor
 * @apiName tabsOptions.fontSelColor
 * @apiGroup Attribute:workbook
 * @apiDescription Font color after table selection(Defaults:#CD853F)
 * - WB.workbook.tabsOptions.fontSelColor = value(color)  
 */

/**
 * @api {null} /null tabsOptions.selFillColor
 * @apiName tabsOptions.selFillColor
 * @apiGroup Attribute:workbook
 * @apiDescription Tab fill color when the current table is selected(Defaults:#fff)
 * - WB.workbook.tabsOptions.selFillColor = value(color)  
 */

/**
 * @api {null} /null tabsOptions.borderColor
 * @apiName tabsOptions.borderColor
 * @apiGroup Attribute:workbook
 * @apiDescription The border color of each table in the tab bar(Defaults:#DCDCDC)
 * - WB.workbook.tabsOptions.borderColor = value(color)  
 */

/**
 * @api {null} /null tabsOptions.fillColor
 * @apiName tabsOptions.fillColor
 * @apiGroup Attribute:workbook
 * @apiDescription Tabs bar background fill color(Defaults:#FDF5E6)
 * - WB.workbook.tabsOptions.fillColor = value(color)   
 */

 /**
 * @api {null} /null tabsOptions.showAdd
 * @apiName tabsOptions.showAdd
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to display the tabs column to add a table + sign(Defaults:true)
 * - WB.workbook.tabsOptions.showAdd = value(boolean) 
 */

/**
 * @api {null} /null showContextMenu
 * @apiName showContextMenu
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to display the right-click menu bar of the workbook(Defaults:false)
 * - WB.workbook.showContextMenu = value(boolean)  
 */

/**
 * @api {null} /null devScreenWidth
 * @apiName devScreenWidth
 * @apiGroup Attribute:workbook
 * @apiDescription Horizontal computer resolution at the time of development
 * - Whether to enable adaptation based on the adaption property, refer to the WB.adaption(bool, devW, devH) method
 * - WB.workbook.devScreenWidth = value(boolean)  
 */

/**
 * @api {null} /null devScreenHeight
 * @apiName devScreenHeight
 * @apiGroup Attribute:workbook
 * @apiDescription Vertical computer resolution at the time of development
 * - Whether to enable adaptation based on the adaption property, refer to the WB.adaption(bool, devW, devH) method
 * - WB.workbook.devScreenHeight = value(boolean)  
 */

/**
 * @api {null} /null adaption
 * @apiName adaption
 * @apiGroup Attribute:workbook
 * @apiDescription Whether to enable adaptive(Defaults:false)   
 * - WB.workbook.adaption = value(boolean)  
 */

/**
 * @api {null} /null print
 * @apiName print
 * @apiGroup Attribute:workbook
 * @apiDescription Print mode   
 * - 1:Print current worksheet(Defaults)
 * - 2:Print workbook
 * - 3:Print the currently selected area of the current table
 * - Method WB.setPrint(obj,Index) can set print settings
 * - WB.workbook.print = value(Int)  
 */

 /**
 * @api {null} /null printOnSamePaper
 * @apiName printOnSamePaper
 * @apiGroup Attribute:workbook
 * @apiDescription Whether multiple tables are printed on the same paper(Defaults:false) 
 * - Only effective when print is 2
 * - Method WB.setPrint(obj,Index) can set print settings
 * - WB.workbook.printOnSamePaper = value(Boolean)  
 */

 /**
 * @api {null} /null printer
 * @apiName printer
 * @apiGroup Attribute:workbook
 * @apiDescription printer
 * - Method WB.setPrint(obj,Index) can set print settings
 * - WB.workbook.printer = value(string)  
 */

/**
 * @api {null} /null marginCopies
 * @apiName marginCopies
 * @apiGroup Attribute:workbook
 * @apiDescription Print the data of two tables on the same sheet of paper The interval between the two data (tables)
 * - This attribute only takes effect when print is equal to 2 and printOnSamePaper is equal to true
 * - The unit is mm
 * - WB.workbook.marginCopies = value(int)  
 */

//activeSheet

/**
 * @api {null} /null sheetName
 * @apiName sheetName
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the worksheet name of the current table  
 * - To set the worksheet name, please use the WB.sheetName(Index, Name) method. WB is a new example
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.sheetName,activeSheet.sheetName  
 */

/**
 * @api {null} /null activeCol
 * @apiName activeCol
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the current column of the current table
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting 
 * - WB.activeSheet.activeCol,activeSheet.activeCol 
 */

/**
 * @api {null} /null activeRow
 * @apiName activeRow
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the current row of the current table 
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting 
 * - WB.activeSheet.activeRow,activeRow.activeRow
 */

/**
 * @api {null} /null rangeCol1
 * @apiName rangeCol1
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the start column of the selected range of the current table(==activeCol) 
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting 
 * - WB.activeSheet.rangeCol1,activeRow.rangeCol1
 */

/**
 * @api {null} /null rangeRow1
 * @apiName rangeRow1
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the start row of the selected range of the current table(==activeRow)  
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting 
 * - WB.activeSheet.rangeRow1,activeSheet.rangeRow1
 */

/**
 * @api {null} /null rangeCol2
 * @apiName rangeCol2
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the end column of the selected range of the current table  
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting 
 * - WB.activeSheet.rangeCol2,activeSheet.rangeCol2
 */

/**
 * @api {null} /null rangeRow2
 * @apiName rangeRow2
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the end row of the selected range of the current table  
 * - (r1, c1, r2, c2) These four values can be obtained by the method WB.getCellRC(Index)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.rangeRow2,activeSheet.rangeCol2  
 */

/**
 * @api {null} /null allowMoveRange
 * @apiName allowMoveRange
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether the current table allows dragging cells(Defaults:true) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.allowMoveRange = value(boolean),activeSheet.allowMoveRang = value(boolean)
 */

/**
 * @api {null} /null selectMode
 * @apiName selectMode
 * @apiGroup Attribute:activeSheet
 * @apiDescription The selected mode of the current table cell 
 * - 0:Freely select range(Defaults)
 * - 1:row mode,click the cell to select the row
 * - 2:Column mode,click the cell to select the column
 * - 3:Cell mode can not be dragged
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.selectMode = value(0|1|2|3),activeSheet.selectMode = value(0|1|2|3)
 */

 /**
 * @api {null} /null notRowSelectionNum
 * @apiName notRowSelectionNum
 * @apiGroup Attribute:activeSheet
 * @apiDescription Which row of the current table has no check box(Defaults:-1 All have check boxes) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.notRowSelectionNum = value(Int),activeSheet.notRowSelectionNum = value(Int) 
 */

/**
 * @api {null} /null notColSelectionNum
 * @apiName notColSelectionNum
 * @apiGroup Attribute:activeSheet
 * @apiDescription Which column of the current table has no check box(Defaults:-1 All have check boxes) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.notColSelectionNum = value(Int),activeSheet.notColSelectionNum = value(Int)  
 */

/**
 * @api {null} /null selHdrTopLeft
 * @apiName selHdrTopLeft
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether the current table has selected all in the upper left corner(Defaults:false) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.selHdrTopLeft = value(boolean),activeSheet.selHdrTopLeft = value(boolean)
 */

/**
 * @api {null} /null startCol
 * @apiName startCol
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the start column of the visible area of the current table
 * - Count from the next column of the frozen column(If there are two columns frozen in front, it will be listed as the third column)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.startCol,activeSheet.startCol
 */

/**
 * @api {null} /null startRow
 * @apiName startRow
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the start row of the visible area of the current table
 * - Count from the next line of the frozen line(If there are two lines frozen in front, it will be listed as the third line)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.startRow,activeSheet.startRow
 */

/**
 * @api {null} /null endCol
 * @apiName endCol
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the end column of the visible area of the current table
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.endCol,activeSheet.endCol
 */

/**
 * @api {null} /null endRow
 * @apiName endRow
 * @apiGroup Attribute:activeSheet
 * @apiDescription Get the end row of the visible area of the current table
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.endRow,activeSheet.endRow
 */

/**
 * @api {null} /null showFixedLine
 * @apiName showFixedLine
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to display the frozen line of the current table(Defaults:false) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.showFixedLine = value(boolean) ,activeSheet.showFixedLine = value(boolean)  
 */

/**
 * @api {null} /null showGridLines
 * @apiName showGridLines
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to display the grid lines of the current table(Defaults:true) 
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.showGridLines = value(boolean),activeSheet.showGridLines = valueboolean)    
 */

/**
 * @api {null} /null alwaysShowButton
 * @apiName alwaysShowButton
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to always display the button of the current table cell(Defaults:true)  
 * - In case of false value, btn will be displayed only when editing is entered
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - Set btn reference setCellType
 * - WB.activeSheet.alwaysShowButton = value(boolean),activeSheet.alwaysShowButton = value(boolean)    
 */

/**
 * @api {null} /null alwaysShowInArea
 * @apiName alwaysShowInArea
 * @apiGroup Attribute:activeSheet
 * @apiDescription If the current table cell btn is always displayed, is it displayed only after clicking in the cell(Defaults:true)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.alwaysShowInArea = value(boolean),activeSheet.alwaysShowInArea = value(boolean)   
 */

/**
 * @api {null} /null headeNotLineInFix
 * @apiName headeNotLineInFix
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to open the head mode of the current table(No grid lines above the frozen row)(Defaults:false)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.headeNotLineInFix = value(boolean),activeSheet.headeNotLineInFix = value(boolean)    
 */

/**
 * @api {null} /null canOpenLayer
 * @apiName canOpenLayer
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to click the current click cell can trigger the corresponding event(Defaults:true)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.canOpenLayer = value(boolean),activeSheet.canOpenLayer = value(boolean)  
 */

/**
 * @api {null} /null gridLinesColor
 * @apiName gridLinesColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The grid line color of the current table(Defaults:ccc)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.gridLinesColor = value(color),activeSheet.gridLinesColor = value(color)    
 */

 /**
 * @api {null} /null printSetting.printHeadings
 * @apiName printSetting.printHeadings
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to print the row and column labels of the current table(Defaults:false)  
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printHeadings = value(boolean),activeSheet.printSetting.printHeadings = value(boolean)  
 */

 /**
 * @api {null} /null printSetting.printGridLine
 * @apiName printSetting.printGridLine
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to print the grid lines of the current table (Defaults:false)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printGridLine = value(boolean),activeSheet.printSetting.printGridLine = value(boolean)  
 */

 /**
 * @api {null} /null printSetting.printDirection
 * @apiName printSetting.printDirection
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the printing direction of the current table
 * - 1:N-shaped(Defaults)
 * - 2:Z-shaped
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printDirection = value(int),activeSheet.printSetting.printDirection = value(int)  
 */

/**
 * @api {null} /null printSetting.orientation
 * @apiName printSetting.orientation
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the paper orientation when the current table is printed
 * - 1:Vertical(Defaults)
 * - 2:Horizontal
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.orientation = value(int),activeSheet.printSetting.orientation = value(int)   
 */

/**
 * @api {null} /null printSetting.isshowfooterpageinfo
 * @apiName printSetting.isshowfooterpageinfo
 * @apiGroup Attribute:activeSheet
 * @apiDescription Page number display information
 * - 0:Do not show(Defaults)
 * - 1:Footer centered
 * - 2:Footer left
 * - 3:Footer right
 * - 4:Header centered
 * - 5:Header left
 * - 6:Header right
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.isshowfooterpageinfo = value(Int),activeSheet.printSetting.isshowfooterpageinfo = value(Int)  
 */

/**
 * @api {null} /null printSetting.footpagestyle
 * @apiName printSetting.footpagestyle
 * @apiGroup Attribute:activeSheet
 * @apiDescription Print page number format of current table  (&p/&P) (Defaults:第 &p 页，共 &P 页)
 * - &p &P fixed(These two symbols will be replaced with the corresponding page and the total number of pages)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.footpagestyle = value(string),activeSheet.printSetting.footpagestyle = value(string)  
 */


 /**
 * @api {null} /null printSetting.paper
 * @apiName printSetting.paper
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the paper type for the current table printing(A3-A7) 
 * - A4(Defaults)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.paper = value(String),activeSheet.printSetting.paper = value(String)   
 */


/**
 * @api {null} /null printSetting.printArea.r1
 * @apiName printSetting.printArea.r1
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the row of the starting cell of the printing area of the current table page  
 * - ''Not set(Defaults)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printArea.r1 = value(Int),activeSheet.printSetting.printArea.r1 = value(Int)  
 */

/**
 * @api {null} /null printSetting.printArea.c1
 * @apiName printSetting.printArea.c1
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the column of the starting cell of the printing area of the current table page 
 * - ''Not set(Defaults)
 * -  Method WB.setPrint(obj,Index) can set print settings
 * -  Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printArea.c1 = value(Int),activeSheet.printSetting.printArea.c1 = value(Int)
 */

/**
 * @api {null} /null printSetting.printArea.r2
 * @apiName printSetting.printArea.r2
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the row of the end cell of the printing area of the current table page
 * - ''Not set(Defaults)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printArea.r2 = value(Int),activeSheet.printSetting.printArea.r2 = value(Int) 
 */

/**
 * @api {null} /null printSetting.printArea.c2
 * @apiName printSetting.printArea.c2
 * @apiGroup Attribute:activeSheet  
 * @apiDescription Set the column of the end cell of the printing area of the current table page  
 * - ''Not set(Defaults)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.printArea.c2 = value(Int),activeSheet.printSetting.printArea.c2 = value(Int) 
 */

/**
 * @api {null} /null printSetting.marginTop
 * @apiName printSetting.marginTop
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the top margin when printing the current table
 * - The unit is mm(Defaults 5mm)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.marginTop = value(Number),activeSheet.printSetting.marginTop = value(Number) 
 */

/**
 * @api {null} /null printSetting.marginBottom
 * @apiName printSetting.marginBottom
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the bottom margin when printing the current table
 * - The unit is mm(Defaults 5mm)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.marginBottom = value(Number),activeSheet.printSetting.marginBottom = value(Number)
 */

/**
 * @api {null} /null printSetting.marginLeft
 * @apiName printSetting.marginLeft
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the left margin when printing the current table
 * - The unit is mm(Defaults 5mm)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.marginLeft = value(Number),activeSheet.printSetting.marginLeft = value(Number)  
 */

/**
 * @api {null} /null printSetting.marginRight
 * @apiName printSetting.marginRight
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the right margin when printing the current table
 * - The unit is mm(Defaults 5mm)
 * - Method WB.setPrint(obj,Index) can set print settings
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.printSetting.marginRight = value(Number),activeSheet.printSetting.marginRight = value(Number)  
 */


/**
 * @api {null} /null cellPadding.left
 * @apiName cellPadding.left
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the left margin to get the cell contents of the current table(Defaults:2)
 * - You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.cellPadding.left = value(Number),activeSheet.cellPadding.left = value(Number)  
 */

/**
 * @api {null} /null cellPadding.right
 * @apiName cellPadding.right
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the right padding to get the cell content of the current table(Defaults:2)  
 * - You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.cellPadding.right = value(Number),activeSheet.cellPadding.right = value(Number)    
 */

/**
 * @api {null} /null cellPadding.top
 * @apiName cellPadding.top
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the top padding to get the cell content of the current table(Defaults:2)
 * - You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.cellPadding.top = value(Number),activeSheet.cellPadding.top = value(Number)
 */

/**
 * @api {null} /null cellPadding.bottom
 * @apiName cellPadding.bottom
 * @apiGroup Attribute:activeSheet
 * @apiDescription Set the bottom padding to get the cell content of the current table(Defaults:2)
 * - You can also use the method WB.setCellPadding(topSize, rightSize, bottomSize, leftSize, Index) to set
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.cellPadding.bottom = value(Number),activeSheet.cellPadding.bottom = value(Number)
 */

/**
 * @api {null} /null isSelectionHideBorder
 * @apiName isSelectionHideBorder
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to hide the check box of the selected cell of the current table(Defaults:false)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.isSelectionHideBorder = value(boolean),activeSheet.isSelectionHideBorder = value(boolean)  
 */

/**
 * @api {null} /null selectOption.selectBorderColor
 * @apiName selectOption.selectBorderColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription Checkbox color of current table(Defaults:green)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.selectOption.selectBorderColor  = value(color),activeSheet.selectOption.selectBorderColor = value(color)
 */

/**
 * @api {null} /null selectOption.selectFillColor
 * @apiName selectOption.selectFillColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The fill color of the check box of the current table(Defaults:rgba(0,0,245,0.1))
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.selectOption.selectFillColor  = value(color),activeSheet.selectOption.selectFillColor = value(color)
 */

/**
 * @api {null} /null rowHeaderData.defaultDataNode.style.fontColor
 * @apiName rowHeaderData.defaultDataNode.style.fontColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The font color of the header of the current table(Defaults:#fff)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.rowHeaderData.defaultDataNode.style.fontColor = value(color),activeSheet.rowHeaderData.defaultDataNode.style.fontColor = value(color)
 */

/**
 * @api {null} /null rowHeaderData.defaultDataNode.style.fillColor
 * @apiName rowHeaderData.defaultDataNode.style.fillColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The fill color of the header of the current table(Defaults:#008844)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.rowHeaderData.defaultDataNode.style.fillColor = value(color),activeSheet.rowHeaderData.defaultDataNode.style.fillColor = value(color) 
 */

/**
 * @api {null} /null rowHeaderData.showRowHeading
 * @apiName rowHeaderData.showRowHeading
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to display the header of the current table(Defaults:false) 
 * - Can be set using WB.showRowHeading(boolean, index) method
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.rowHeaderData.showRowHeading = value(boolean),activeSheet.rowHeaderData.showRowHeading = value(boolean)  
 */

/**
 * @api {null} /null rowHeaderData.defaultW
 * @apiName rowHeaderData.defaultW
 * @apiGroup Attribute:activeSheet
 * @apiDescription The width of the current table header(Defaults:30)  
 * - Can be set using WB.hdrWidth(width, index) method
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.rowHeaderData.defaultW = value(number)  
 */

/**
 * @api {null} /null colHeaderData.defaultDataNode.style.fontColor
 * @apiName colHeaderData.defaultDataNode.style.fontColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The font color of the column header of the current table(Defaults:#fff)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.colHeaderData.defaultDataNode.style.fontColor = value(color)，activeSheet.colHeaderData.defaultDataNode.style.fontColor = value(color)
 */

/**
 * @api {null} /null colHeaderData.defaultDataNode.style.fillColor
 * @apiName colHeaderData.defaultDataNode.style.fillColor
 * @apiGroup Attribute:activeSheet
 * @apiDescription The fill color of the current column header(Defaults:#008844)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.colHeaderData.defaultDataNode.style.fillColor = value(color)，activeSheet.colHeaderData.defaultDataNode.style.fillColor = value(color)
 */

/**
 * @api {null} /null colHeaderData.showColHeading
 * @apiName colHeaderData.showColHeading
 * @apiGroup Attribute:activeSheet
 * @apiDescription Whether to display the column header of the current table(Defaults:false) 
 * - Can be set using WB.showColHeading(boolean, index) method
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.colHeaderData.showColHeading = value(boolean)，activeSheet.colHeaderData.showColHeading = value(boolean)
 */

/**
 * @api {null} /null colHeaderData.defaultH
 * @apiName colHeaderData.defaultH
 * @apiGroup Attribute:activeSheet
 * @apiDescription The height of the column header of the current table(Defaults:30) 
 * - Can be set using WB.hdrHeight(height, index) method
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.colHeaderData.defaultH = value(Number)，activeSheet.colHeaderData.defaultH = value(Number)
 */

/**
 * @api {null} /null defaults.colWidth
 * @apiName defaults.colWidth
 * @apiGroup Attribute:activeSheet
 * @apiDescription The default column width of the current table(Defaults:60)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.defaults.colWidth = value(number)，activeSheet.defaults.colWidth = value(number)  
 */

/**
 * @api {null} /null defaults.rowHeight
 * @apiName defaults.rowHeight
 * @apiGroup Attribute:activeSheet
 * @apiDescription The default row height of the current table(Defaults:30)
 * - Please refer to (sheetIndex, getActiveSheet) for the current table acquisition and setting
 * - WB.activeSheet.defaults.rowHeight = value(number) ，activeSheet.defaults.rowHeight = value(number) 
 */






















