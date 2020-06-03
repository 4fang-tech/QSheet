this.defaultOptions = {
    "activeSheet":0,                                    //当前表索引（0开始）
    "allowDelete":true,                                 //是否允许删除键删除选中的文本
    "allowResize":true,                                 //是否允许通过拖动调整列宽行高
    "allowTabs":true,                                   //是否允许使用tab键切换单元格
    "showTabs":0,                                       //是否显示tab栏（0 不显示 1显示 2按钮组 默认:0）
    "defaultFontName":"微软雅黑",                        //工作簿的默认字体类型
    "defaultFontSize":14,                               //工作簿的默认字体大小
    "defaultLineHeight":20,                             //工作簿的字体默认行高
    "width":window.innerWidth-this.scrollbarWidth,      //工作簿宽度
    "height":window.innerHeight-this.scrollbarWidth,    //工作高度
    "showHScrollBar":1,                                 //显示水平滚动条的方式（0不显示 1工作簿较宽的时候显示）
    "showVScrollBar":1,                                 //显示垂直滚动条的方式（0不显示 1工作簿较宽的时候显示）
    "behaviorMode":1,                                   //工作簿模式  1  sheet模式   2 grid模式（点击单元格就有输入框、没有选中框） 3 sheet单击可编辑模式
    "scrollMode":2,                                     //滚动条的再冻结情况下的模式（1滚动条从顶部或者最左部开始  2滚动条再冻结线下方）
    "devScreenWidth":1280,                              //开发时候的电脑分辨率的宽 结合adaption属性是否开启自适应  
    "devScreenHeight":720,                              //开发时候的电脑分辨率的高 结合adaption属性是否开启自适应   
    "adaption":false,                                   //是否开启自适应
    "showWorkBookBorder":true,                          //是否显示工作簿的边框
    "showRowArrow":false,                               //是否显示当前行的箭头
    "rowArrowColor":"#000",                             //当前行箭头的填充色
    "print":1,                                          //1打印当前工作表(默认) 2打印工作簿 3打印当前表当前选中区域
    "printOnSamePaper":false,                           //仅当print为2时生效 在纸张可以放下其他表的情况下 多张表是否放在同一张纸上 默认false
    "printer":"undefined",                              //打印机
    "marginCopies":0,                                   //数据间隔 同张纸张两条数据之间的间隔
    "tabsOptions":{                                     //tabs栏参数
        "fontColor":"#444",                             //tab栏的字体颜色
        "font":"14px/20px 宋体",                        //tab栏字体
        "fontSelColor":"#CD853F",                       //tab栏选中之后的字体颜色
        "selFillColor":"#fff",                          //tab选中当前表的填充颜色
        "borderColor":"#DCDCDC" ,                       //tab栏每个表的边框颜色
        "fillColor":"#FDF5E6",                          //tabs栏的背景填充色
        "showAdd":true                                  //是否显示tabs栏添加一个表的的+号
    },
    "stopPaintedCount":1,                               //当这个值是否0的时候才发生重绘（每调用一次startPaint，值减1 每调用一次stopPainted,值加1）
    "stopEventCount":0,                                 //当这个值是否0的时候才发生执行on事件（每调用一次resumeEvent，值减1 每调用一次stopEvent,值加1）
    "showContextMenu":false,                            //是否显示右击菜单栏
    "sheetList":["Sheet1"],                             //tab栏的sheet列表
     sheets:{
        "Sheet1":{                                      //表  与sheetList对应
           "sheetName":"Sheet1",                        //表名
           "activeRow":0,                               //当前激活单元格所在行
           "activeCol":0,                               //当前激活单元格所在列
           "rangeCol1":0,                               //当前选中范围开始列   ==activeCol
           "rangeRow1":0,                               //当前选中范围开始行   ==activeRow   
           "rangeCol2":0,                               //当前选中范围结束列
           "rangeRow2":0,                               //当前选中范围结束行
           "allowMoveRange":true,                       //拖选
           "selectMode":0,                              //单元格选中模式 （0range模式，自由选择，1:row行模式，点击单元格选中行，2:col 列模式，点击单元格选中列，3:cell单元格模式   不可拖选）
           "selHdrTopLeft":false,                       //是否选中了左上角的全选
           "startCol":0,                                //可视区域开始列 若前面有冻结了两列  则开始列为第三列
           "startRow":0,                                //可视区域开始行 若前面有冻结了两行  则开始行为第三行
           "endCol":0,                                  //可视区域的结束列
           "endRow":0,                                  //可视区域的结束行
           "showFixedLine":false,                       //是否显示冻结线
           "showGridLines":true,                        //是否显示网格线
           "alwaysShowButton":true,                     //是否始终显示单元格的button
           "alwaysShowInArea":true,                     //始终显示单元格btn的情况下  是否只有在点击在该单元格再显示
           "headeNotLineInFix":false,                   //是否开始头模式（冻结行上方没有网格线,除冻结行最后行）
           "canOpenLayer":true,                         //是否点击单元格能触发相应事件
           "gridLinesColor":"#ccc",                     //表的网格线颜色
           "printSetting":{
                "printHeadings":false,                  //是否打印行号列标
                "printGridLine":false,                  //是否打印网格线
                "printDirection":1,                     //打印方向(1:先列后行呈N字形 2：先行后列呈Z字形) 
                "paper":'A4',                           //纸张类型
                "orientation":1,                        //纸张方向
                "isshowfooterpageinfo":0,               //显示页码信息参数"0", "1", "2", "3", "4", "5", "6"对应:"不显示", "页脚居中", "页脚居左", "页脚居右", "页眉居中", "页眉居左", "页眉居右
                "footpagestyle":"第 &p 页，共 &P 页",    //页码格式(&p &P注意区分大小写 这两个符号会替换成当前第几页 一共几页得数字)
                "printArea":{                           //打印区域
                    "r1":"",                            //r1c1开始单元格
                    "c1":"",
                    "r2":"",                            //r2c2结束单元格
                    "c2":""
                },
                "marginTop":5,                          //打印上边距
                "marginBottom":5,                       //打印下边距
                "marginLeft":5,                         //打印左边距
                "marginRight":5,                        //打印右边距
            },
           "cellPadding":{                              //表的单元格内容的内边距
                "left":2,
                "right":2,
                "top":2,
                "bottom":2
            },
            "isSelectionHideBorder":false,              //是否隐藏选中单元格的选中框
            "rows" : [] ,                               //行 size行高 （bHidden为true时，行隐藏）name:"name" name设置行头名  noDefaultH 为true 表示不是默认高度
            "columns" : [],                             //列 size列高 （bHidden为true时，列隐藏）name:"name" name设置列头名 noDefaultW 表示不是默认宽度
            "spans": [                                  //合并   
                // {
                //     "row":1,                         合并所在行
                //     "rowCount":1,                    合并总行数
                //     "col":1,                         合并所在列
                //     "colCount":5                     合并总列数
                // },   
            ],
            "data":{                                    //数据
                "dataTable":{  
                    // "1": {                                               //行
                    //     "1":{                                            //列
                    //         "value":-1000,                               //单元格的值
                    //         "text":"-1000",                              //单元格的值(tostring之后的值)
                    //         "style":{                                    //单元格的样式
                    //             "hAlign":3,                              //水平对齐方式
                    //             "vAlign":2,                              //垂直对齐方式
                    //             "font":"italic bold 16px/20px 楷体",     //单元格字体
                    //             "fontStrikeout":false,                   //是否删除线
                    //             "fontUnderline":false,                   //是否下划线
                    //             "wordWrap":false,                        //是否自动换行
                    //             "fontColor":"#666666",                   //字体颜色
                    //             "cellType":{"name":4,"list":['a','b']},  //单元格类型name（0无button 1button为三个点  2button为下箭头  3button为checkBox  4为下拉框）
                    //             "lock":false,                            //单元格是否锁定(没选中框，可选，不能编辑)
                    //             "bgColor":"pink",                        //单元格填充色
                    //             "canEdit":false,                         //单元格不能编辑(有选中框，可选，不能编辑)
                    //              "borderRight":{                         //单元格右边框（左上下同理）
                    //                   "color":"#fff",                    //边框颜色
                    //                   "style":2,                         //线条大小（1:	Thin Line,2: Medium Line,5: Thick Line）
                    //              },
                    //             "formatter":""                           //单元格格式
                    //         },
                    //         "formula":"=A1+B1"                           //单元格公式
                    //     }
                    // }
                }                        
            },
            "fixed":{
                "fixedRow":0,                           //冻结开始的行
                "fixedCol":0,                           //冻结开始的列
                "fixedRows":0,                          //冻结了多少行
                "fixedCols":0,                          //冻结了多少列
                "fixedH":0,                             //可视冻结高度
                "fixedW":0,                             //可视冻结宽度
                "hideH":0,                              //冻结行到开始行的隐藏高度
                "hideW":0,                              //冻结列到开始列的隐藏宽度
            },
            "rowStyle":{                                //隔行变色
                "color":'rgb(255,255,224)',             //颜色
                "oddEven":0,                            //奇偶
            },
            "selectOption":{              
                "selectBorderColor":'green',            //选中框颜色
                "selectFillColor":'rgba(0,0,245,0.1)',  //选中框多选的时候的填充色
            },
            "scrollOption":{                            //滚动条参数
                "coordY":0,                             //滚动垂直的距离
                "totolY":0,                             //垂直方向已经存在的距离
                "coordX":0,                             //滚动水平的距离
                "totolX":0,                             //水平方向已经存在的距离
                "offsetX":0,                            //水平偏移（滚动到最后一列   首列有该偏移量  最后列才能够刚刚显示完整）
                "offsetY":0,                            //水平偏移（滚动到最后一列   首列有该偏移量  最后列才能够刚刚显示完整）
            },
            "rowHeaderData":{                           //行头
                "defaultDataNode":{
                    "style":{  
                        "fontColor":"#fff",             //行头字体颜色
                        "fillColor":"#008844",          //行头填充色
                    }
                },
                "showRowHeading":true,                  //是否显示行头
                "defaultW":30,                          //行头的宽度
                
            },
            "colHeaderData":{                           //列头
                "defaultDataNode":{
                    "style":{
                        "fontColor":"#fff",             //列头字体颜色
                        "fillColor":"#008844",          //列头填充颜色
                    }    
                },
                "showColHeading":true,                  //是否显示列头
                "defaultH":30,                          //列头高度
                // "spans":[                            //是否有列头的拆分
                    // {
                    //     "col":1,                     所在列
                    //     "name":"拆分1",              名称
                    //     "colCount":2,                拆分列数
                    //     "height":25,                 高度（0<height<defaultH  默认时defaultH的一般高度）
                    // },
                // ]

            },
            "defaults":{                                //默认参数
                "colWidth":60,                          //列宽
                "rowHeight":30,                         //行高
            },
        }, 
    }
};