function Workbook(id,data){   
    this.workbook = data || {}; 
    this.boxId = id;    
    this.container = document.querySelector("#"+this.boxId);    
    this.container.className = 'clearfix';
    this.canvas = document.createElement('canvas');
    this.canvas.setAttribute('id',this.boxId+'Canvas');
    this.canvas.innerText='检测到您的浏览器版本过低,请升级您的浏览器,以获取更好的体验效果...';
    this.container.appendChild(this.canvas)
    this.paddingL = parseInt(this.getStyle(this.container,'paddingLeft')); //获取容器padding值(用于弹出层 输入框的等元素定位)
    this.paddingT = parseInt(this.getStyle(this.container,'paddingTop'));
    this.ctx = this.canvas.getContext('2d');
    this.ratio = this.getPixelRatio()   //获取屏幕缩放比例(缩放影响canvas模糊 canvas要缩放这个比的值以保证显示效果)
    this.scrollbarWidth = this.getScrollBarWidth(); //获取浏览器滚动条宽度
    this.totolHeight = 0;               //所有行(rows的总和)
    this.totolWidth = 0;    
    this.startTabW = 100;               //tabs栏表开始位置的初始距离(该值根据左边的箭头宽度改变)
    this.sheetItemExtraW = 10;          //tabs栏每一个表项增加的额外宽度
    this.tabArrowPadding = 15           //tabs栏四个箭头的padding
    this.scrollVMinHeight = 30;         //滚动条最小高度
    this.scrollHMinWidth = 30;          //滚动条最小宽度
    this.startSheet = 0;                //tabs栏第一个表
    this.endSheet = 0;  
    this._tabsHeight = 20;              //tab栏高度
    this.oncellbuttonclick =null;       //btn按钮事件
    this.onshoweditor = null;           //设置光标位置事件
    this.ownerdraw = null;              //用户自绘函数
    this.onclickcellopenlayer = null;   //单元格点击触发事件
    this.onopenpopup = null;            //弹出下拉框事件
    this.onvaluechange = null;          //单元格值改变事件
    this.onrelevantchange = null;       //事件
    this.onopenmorelist = null;         //下拉框最后一项更多事件
    this.onformularesult = null;        //公式结果事件
    this.onrowmodedblclick = null       //行选模式的双击事件
    this.onenteraddrow = null;          //绑定在回车的事件(增加行)
    this.onafteraddrow = null           //新增一行后的事件
    this.onequalsign = null;            //绑在输入框键入=(等号)事件
    this.documentHandle = null;         //绑定在documnet上的处理事件
    this.formatBrushStyle = {};         //初始化格式刷
    this.isHArrow = false               //辅助属性start  
    this.isVArrow = false;
    this.percentageV = 1;   
    this.percentageH = 1;
    this.tempValue = {};   
    this.scrollVLeft = 0;   
    this.eButton = 0;                   
    this.horizontalLine = 0;          
    this.verticalLine =0;   
    this.isShowBar = false;
    this.drawOne = true;
    this.tabSumW = 0;
    this.diffTabsH = 0;
    this.spansRowCount = 1;
    this.spansColCount = 1;
    this.keyDownCount = true;
    this.btnArr = [];
    this.checkBoxBtn = [];
    this.arrowOne = true;
    this.incanvasPath = true;
    this.addHeight = 0;
    this.defaultOptions = this.defaultBook()        //默认工作簿
    this.initWorkbook();
    this._defaultCellStyle = this.defaultCell()     //默认单元格样式
    this.container.style.position = 'relative'
    this.container.style.width = this.width+'px';
    this.container.style.height = this.height+'px';
    this.activeSheet = this.getActiveSheet(this.workbook.activeSheet);
    this.popup = (window.top.popup)?window.top.popup:(window.Popup)?new Popup():null;
    this.init() ;
    window.boxId = this.boxId;
}

//初始化
Workbook.prototype.init = function(){
    this.contextMenu();     //初始化右键菜单栏
    this.getStartSheet();   //获取tabs开始表
    this.getEndSheet();     //获取tabs结束表
    this.initTextArea()     //初始化输入
    this.bindEvent();       //事件绑定
};

//工作簿默认的参数
Workbook.prototype.defaultBook = function(){
    return {
        "activeSheet":0,                                    //当前表索引（0开始）
        "allowDelete":true,                                 //是否允许删除键删除选中的文本
        "allowResize":true,                                 //是否允许通过拖动调整列宽行高
        "allowTabs":true,                                   //是否允许使用tab键切换单元格
        "showTabs":0,                                       //是否显示tab栏（0不显示 1底部,2显示按钮）
        "defaultFontName":"宋体",                           //工作簿的默认字体类型
        "defaultFontSize":14,                               //工作簿的默认字体大小
        "defaultLineHeight":20,                             //工作簿的字体默认行高
        "width":window.innerWidth-this.scrollbarWidth,      //工作簿宽度(窗口的宽高减去滚动条的宽度)
        "height":window.innerHeight-this.scrollbarWidth,    //工作高度(窗口的宽高减去滚动条的宽度)
        "showHScrollBar":1,                                 //显示水平滚动条的方式（0不显示 1工作簿较宽的时候显示）
        "showVScrollBar":1,                                 //显示垂直滚动条的方式（0不显示 1工作簿较宽的时候显示）
        "behaviorMode":1,                                   //工作簿模式  1  sheet模式   2 grid模式（点击单元格就有输入框、没有选中框） 3 sheet单击可编辑模式
        "scrollMode":2,                                     //滚动条的再冻结情况下的模式（1滚动条从顶部或者最左部开始  2滚动条再冻结线下方）
        "devScreenWidth":1280,                              //开发时候的电脑分辨率的宽 结合adaption属性是否开启自适应  
        "devScreenHeight":720,                              //开发时候的电脑分辨率的高 结合adaption属性是否开启自适应   
        "adaption":false,                                   //是否开启自适应
        "showWorkBookBorder":true,                          //是否显示工作簿的边框
        "workBookBorderColor":"#ccc",                       //工作簿的边框颜色
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
            "fillColor":"#FFF",                             //tabs栏的背景填充色
            "showAdd":true                                  //是否显示tabs栏添加一个表的的+号
        },
        "stopPaintedCount":1,                               //当这个值是否0的时候才发生重绘（每调用一次startPaint，值减1 每调用一次stopPainted,值加1）
        "stopEventCount":0,                                 //当这个值是否0的时候才发生执行on事件（每调用一次resumeEvent，值减1 每调用一次stopEvent,值加1）
        "showContextMenu":false,                            //是否显示右击菜单栏
        "sheetList":["Sheet1"],                             //tab栏的sheet列表
         "sheets":{
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
                    "footpagestyle":"第 &p 页，共 &P 页",   //页码格式
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
                "spans": [],                                //合并   
                "data":{                                    //数据
                    "dataTable":{} 
                },
                "fixed":{
                    "fixedRow":0,                           //冻结开始的行
                    "fixedCol":0,                           //冻结开始的列
                    "fixedRows":0,                          //冻结了多少行
                    "fixedCols":0,                          //冻结了多少列
                    "fixedH":0,                             //可视冻结高度
                    "fixedW":0,                             //可视冻结宽度
                    "hideH":0,                              //冻结行到开始行的隐藏高度
                    "hideW":0,                              //true冻结列到开始列的隐藏宽度
                },
                "rowStyle":{                                //隔行变色
                    "color":'rgb(255,255,224)',             //颜色
                    "oddEven":0,                            //奇偶
                    "start":0,
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
                    "offsetY":0,                            //垂直偏移（滚动到最后一行   首行有该偏移量  最后行才能够刚刚显示完整）
                },
                "rowHeaderData":{                           //行头
                    "defaultDataNode":{
                        "style":{  
                            "themeFont":"Body",
                            "fontColor":"#fff",             //行头字体颜色
                            "fillColor":"#008844",          //行头填充色
                        }
                    },
                    "showRowHeading":false,                 //是否显示行头
                    "defaultW":30,                          //行头的宽度
                    
                },
                "colHeaderData":{                           //列头
                    "defaultDataNode":{
                        "style":{
                            "themeFont":"Body",    
                            "fontColor":"#fff",             //列头字体颜色
                            "fillColor":"#008844",          //列头填充颜色
                        }    
                    },
                    "showColHeading":false,                 //是否显示列头
                    "defaultH":30,                          //列头高度
                },
                "defaults":{                                //默认参数
                    "colWidth":60,                          //列宽
                    "rowHeight":30,                         //行高
                },
            }, 
        }
    };
} 

//单元格的默认样式
Workbook.prototype.defaultCell = function(){
    return {
        "fontColor":"#000",
        "hAlign":1,
        "cellType":{"name":0},
        "lock":false,
        "vAlign":2,
        "formatter":"",
        "font":this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName,
        "wordWrap":false,
    }
}

//初始化workbook
Workbook.prototype.initWorkbook = function(){
    var defaultOptions = JSON.parse(JSON.stringify(this.defaultOptions));
    this.addWorkBookAttr(defaultOptions,this.workbook);
    this.addWorkBookAttr(defaultOptions.tabsOptions,this.workbook.tabsOptions);  
    var sheetList = Object.keys(this.workbook.sheets);
    this.workbook.sheetList = sheetList;
    for(var i = 0;i<sheetList.length;i++){
       this.initSheet(sheetList[i],i);
    };
    return this.workbook
}

//初始化传进来的每一个表(把缺少的属性补全默认值)
Workbook.prototype.initSheet = function(sheetName,i){
    var sheet1 = JSON.parse(JSON.stringify(this.defaultOptions.sheets.Sheet1));
    var active = this.workbook.sheets[sheetName],_this = this;
    this.addWorkBookAttr(sheet1,active);
    active.sheetName = sheetName;
    var objKey = ['cellPadding','printSetting','printSetting.printArea','data','fixed','selectOption','scrollOption','rowHeaderData',
    'rowHeaderData.defaultDataNode','rowHeaderData.defaultDataNode.style','colHeaderData','colHeaderData.defaultDataNode',
    'colHeaderData.defaultDataNode.style','defaults']
    doAddWorkBookAttr(0)
    function doAddWorkBookAttr(idx){
        if(idx<objKey.length){
            _this.addWorkBookAttr(sheet1[objKey[idx]],active[objKey[idx]]);
            idx++;
            doAddWorkBookAttr(idx)
        }
    }
     var spans = active.spans;   //初始化检测合并改变初始化r2 c2的值
     spans.forEach(function(item){
         if(item.row==active.rangeRow1&&item.col==active.rangeCol1){
             active.rangeRow2 = item.row+item.rowCount-1;
             active.rangeCol2 = item.col+item.colCount-1;
         };
     });
     for(var a = 0;a<active.columns.length;a++){
         this.setNoDefaultW(a,i)
     }
     for(var b = 0;b<active.rows.length;b++){
         this.setNoDefaultH(b,i)
     }
     return active;
}

//补充workbook缺少属性
Workbook.prototype.addWorkBookAttr = function(obj1,obj2){
    for(var key in obj1){
        if(obj1.hasOwnProperty(key)===true&&obj2.hasOwnProperty(key)===false){
            obj2[key] = obj1[key]
        };  
    };
}

//事件绑定
Workbook.prototype.bindEvent = function(){         
    var menu = document.querySelector("#"+this.boxId+" div[menu=contextmenu]"),addrowElement = document.querySelector("#"+this.boxId+" .addrow"),
    addcolElement = document.querySelector("#"+this.boxId+" .addcol"),delrowElement = document.querySelector("#"+this.boxId+" .delrow"),
    delcolElement = document.querySelector("#"+this.boxId+" .delcol"),leftAlign = document.querySelector("#"+this.boxId+" .leftAlign"),
    centerAlign = document.querySelector("#"+this.boxId+" .centerAlign"),rightAlign = document.querySelector("#"+this.boxId+" .rightAlign"),
    topAlign = document.querySelector("#"+this.boxId+" .topAlign"),middleAlign = document.querySelector("#"+this.boxId+" .middleAlign"),
    bottomAlign = document.querySelector("#"+this.boxId+" .bottomAlign"),copy = document.querySelector("#"+this.boxId+" .copy"),
    cut =  document.querySelector("#"+this.boxId+" .cut"),child = document.querySelector('#'+this.boxId+" .child"),
    paste = document.querySelector("#"+this.boxId+" .paste"),moveVBar = false, moveHBar = false, moveCell = false,changHeight = false,
    changWidth = false,isArrow = false,isClickInCell = false,clickX = 0,clickY = 0,moveX = 0,moveY = 0,r=0,c=0;

    //滚轮事件
    this.addEvent(this.canvas,'mousewheel',this.throttle(this.scrollFun.bind(this),150,200));
    this.addEvent(this.canvas,'DOMMouseScroll',this.throttle(this.scrollFun.bind(this),150,200)); //火狐

    //canvas绑定mousedown事件
    this.addEvent(this.canvas,"mousedown",function(e){
        e = e || window.event;
        var pos = this.getMousePos(this.canvas,e),x = pos.x,y = pos.y,cellTypeName,index = this.workbook.activeSheet,
        showTabs = this.workbook.showTabs,showAdd = this.workbook.tabsOptions.showAdd,sheetList = this.workbook.sheetList,
        isClickVBar = this.isClickVBar(x,y),isClickHBar = this.isClickHBar(x,y),cellRange = this.getCellRC(),
        cellInfo,createDiv = 'vertical',arrowXInfo = this.getTabArrowX(),firstX = arrowXInfo.firstX,previousX = arrowXInfo.previousX,
        latterX = arrowXInfo.latterX,lastX = arrowXInfo.lastX,arrowW = this.tabArrowWidth,canOpenLayer = this.activeSheet.canOpenLayer,
        isClickInCell = this.isClickCell(x,y);
        clickX = x;clickY = y;
        this.eButton = e.button;
        this.incanvasPath=true;
        window.boxId = this.boxId
        //点击添加一个表区域：
        if(x>this.tabSumW+10&&x<this.tabSumW+20&&y>this.height-this.tabsHeight&&y<this.height&&showTabs==1&&showAdd){
            this.addSheet()
            this.startPaint(true)
        };

        //点击sheet列表(切换sheet表)：
        var rightDiff = this.startTabW;
        this.ctx.font = this.workbook.tabsOptions.font;   
        for(var i = this.startSheet;i<sheetList.length;i++){
            var itemW = this.ctx.measureText(sheetList[i]).width+this.sheetItemExtraW;
            if(x>rightDiff&&x<(rightDiff+itemW)&&y>(this.height-this.tabsHeight)&&y<this.height&&showTabs==1){
                this.sheetIndex(i);
                this.startPaint(true)
            }
            rightDiff += itemW;
        };

        moveVBar = (isClickVBar)?true:false;
        moveHBar = (isClickHBar)?true:false;

        //点击单元格区域   非滚动条区域
        if(isClickInCell){
            cellInfo = this.getActiveCell(x,y)   //传入坐标点得到对应的cell信息
            r = cellInfo.r;
            c = cellInfo.c;
            // var isLock = this.getLock(r,c);
            if(!(e.button == 2 &&(r>=cellRange.r1&&r<=cellRange.r2&&c>=cellRange.c1&&c<=cellRange.c2))){
                moveCell = true;
            }
            isArrow = (this.isHArrow || this.isVArrow)?true:false;   
        };
        
        //点击改变列宽行高区域的时候
        if(this.isHArrow&&this.workbook.allowResize){
            changWidth=true;
            createDiv = "vertical"
            this.createLine(createDiv)
        }
        if(this.isVArrow&&this.workbook.allowResize){
            createDiv  = "horizontal"
            changHeight=true;
            this.createLine(createDiv)
        };
      
        //点击tabs切换箭头
        if(x>firstX-3&&x<firstX+arrowW+3&&y>this.height-this.tabsHeight&&y<this.height&&showTabs==1){ 
            this.startSheet = 0;
            this.getEndSheet();
            this.drawTabs();
        };
        if(x>previousX-3&&x<previousX+arrowW+3&&y>this.height-this.tabsHeight&&y<this.height&&showTabs==1){    
           if(this.startSheet>=1){
               this.startSheet --;
               this.getEndSheet();
               this.drawTabs()
           }
        };
        if(x>latterX-3&&x<latterX+arrowW+3&&y>this.height-this.tabsHeight&&y<this.height&&showTabs==1){         
            if(this.endSheet<sheetList.length -1){
                this.endSheet ++;
                this.getStartSheet();
                this.drawTabs()
            }
        };
        if(x>lastX-3&&x<lastX+arrowW+3&&y>this.height-this.tabsHeight&&y<this.height&&showTabs==1){       
            this.endSheet = sheetList.length -1;
            this.getStartSheet();
            this.drawTabs()
        };

        //点击cellbutton;
        /**
         * @api {null} /null oncellbuttonclick
         * @apiName oncellbuttonclick
         * @apiGroup Event
         * @apiDescription Click the cell button to trigger the event 
         * - The function returns the table index and the current row and current column
         * @apiParam {Function} callback 
         * @apiExample {javascript} demo:
            WB.oncellbuttonclick=function(idx,r,c){
                alert("oncellbuttonclick"+idx+","+r+","+c)
            }
         */
        var f = this.oncellbuttonclick,isClickBtn = false,cellType = this.getCellType(r,c,r,c,index);
        if(cellType){
            cellTypeName = cellType.name;
        };
        if((cellTypeName==1||cellTypeName==2)&&e.button==0){
            isClickBtn = this.isClickBtn(r,c,x,y)
            if(typeof(f)=='function'&&this.workbook.stopEventCount==0&&isClickBtn){
                f(index,r,c);
            }
        }

        /**
         * @api {null} /null onopenpopup
         * @apiName onopenpopup
         * @apiGroup Event
         * @apiDescription Clicking on the cell with celltype==4 will open the drop-down window
         * - Need to import popup.js file
         * - Returns 10 values(w,h,x,y,list,r,c,child,fontSize,noSetIndex)
         * - noSetIndex(Special item, you can judge that the corresponding event occurs on the item)
         * @apiParam {Function} callback 
         * @apiExample {javascript} demo:
            WB.onopenpopup = function(w,h,x,y,list,r,c,child,fontSize,noSetIndex){
                var data =[];
                if(list){
                    for(var i =0;i<list.length;i++){
                        data.push({name:list[i]})
                    }
                };
                WB.popup._init({
                    shade: false,
                    area:[w+'px',list.length*h+"px"],
                    offset:[x+'px', y+'px'],
                    backgroundColor: '#dccccc',
                    rowHeight:34+"px",
                    data:data,
                    fontSize:fontSize,
                    success: function(index,value,listIndex) {
                        if(listIndex==noSetIndex){
                            WB.onopenmorelist = function(){
                                // alert("something")
                            };
                            if(typeof(WB.onopenmorelist)=='function'){
                                WB.onopenmorelist()
                            }
                        }else{
                            WB.setValue(r,c,value)
                            child.innerHTML = value
                            WB.startPaint(true)
                        }
                        this.close(index); 
                    },
                })
            }
         */

        /**
         * @api {null} /null onopenmorelist
         * @apiName onopenmorelist
         * @apiGroup Event
         * @apiDescription Refer to which item is triggered by onopenpopup, the function needs to be placed in onopenpopup success and called
         * @apiParam {Function} callback 
         */
        if(isClickInCell&&e.button==0&&r!=-1&&c!=-1&&cellTypeName==4){
           this.openPopup(r,c,index)  
        }

         /**
         * @api {null} /null onclickcellopenlayer
         * @apiName onclickcellopenlayer
         * @apiGroup Event
         * @apiDescription Click on the cell to trigger the event
         * - The function returns the table index and the current row and current column
         * @apiExample {javascript} demo:
            WB.onclickcellopenlayer=function(idx,r,c){
                alert("onclickcellopenlayer"+idx+","+r+","+c)
            }
         * @apiParam {Function} callback 
         */
        var OpenLayer = this.onclickcellopenlayer;
        if(typeof(OpenLayer)=='function'&&this.workbook.stopEventCount==0&&canOpenLayer&&isClickInCell&&e.button==0&&cellTypeName!=3&&cellTypeName!=4&&!isClickBtn){
            OpenLayer(index,r,c);
        }; 
       
        //点击checkbox
        if(cellTypeName==3&&e.button==0&&r!=-1&&c!=-1){
            var value = this.getValue(r,c),v = (value)?0:1;
            this.setValue(r,c,v)
        };  
    }.bind(this));

    //canvas mouseup 事件
    this.addEvent(this.canvas,"mouseup",this.openContextMenu.bind(this));

    //canvas dblclick 事件
    this.addEvent(this.canvas,'dblclick',function(e){
        e = e || window.event;
        var r = this.activeSheet.activeRow,c = this.activeSheet.activeCol,cellTypeName,selectMode = this.activeSheet.selectMode,
        mode = this.workbook.behaviorMode,index = this.workbook.activeSheet,cellType = this.getCellType(r,c,r,c),
        isLock = this.getLock(r,c,r,c),isCanEdit = this.getCanEdit(r,c,r,c),isClickBtn = false,f = this.onrowmodedblclick;
        if(cellType){
            cellTypeName = cellType.name;
        }
        if(isCanEdit!==false&&!isLock){
            this.isEdit = true;  
        }  
        if(mode==1&&this.isEdit){
            if((cellTypeName==1||cellTypeName==2)&&!this.activeSheet.alwaysShowButton){
                this.startPaint()
            }else{
                this.openInput(r,c)
            };
        }
        
        /**
         * @api {null} /null onrowmodedblclick
         * @apiName onrowmodedblclick
         * @apiGroup Event
         * @apiDescription Trigger event in double-click mode selection  
         * - The function returns the table index and the current row and current column
         * @apiParam {Function} callback 
         * @apiExample {javascript} demo:
            WB.onrowmodedblclick=function(idx,r,c){
                //do something...
            }
         */
        if(cellTypeName==1||cellTypeName==2){
            isClickBtn = this.isClickBtn(r,c,clickX,clickY);
        }
        if(typeof(f)=='function'&&this.workbook.stopEventCount==0&&selectMode==1&&e.button==0&&cellTypeName!=3&&cellTypeName!=4&&isClickInCell&&!isClickBtn){
            f(index,r,c);
        }; 
    }.bind(this));

    this.documentHandle = {
        //document mousedown事件
        mousedown:function(e){
            e = e || window.event;
            var pos = this.getMousePos(this.canvas,e),x = pos.x,y = pos.y;
            // if(x>0&&x<this.width&&y>0&&y<this.height){
            //     window.boxId = this.boxId;
            // }
            // if(window.boxId == this.boxId){
                if(menu)  menu.className = "hidden";
                if((x<0||x>this.horizontalLine||y<0||y>this.verticalLine)&&this.tempValue.r>=0&&this.tempValue.c>=0){
                    var child = document.querySelector('#'+this.boxId+' .child'),text = child.innerText,
                    value = this.getValue(this.activeSheet.activeRow,this.activeSheet.activeCol);
                    this.removeTextArea()
                    if(text!=value&&text!=''){
                        this.startPaint()
                    }
                }
                if(x<0||x>this.width||y<0||y>this.height){
                    this.incanvasPath = false;
                }
            // }
        }.bind(this),

        //document mousemove 事件
        mousemove:function(e){
            e = e || window.event;
            var _this = this
            var box = this.canvas.getBoundingClientRect();
            var x = e.clientX - box.left,y = e.clientY - box.top,isHasVBar,isHasHBar;
            moveX = x; moveY = y;
            //鼠标移入canvas表显示滚动条
            if(x<this.width&&y<this.height&&this.drawOne&&x>0&&y>0){
                this.drawOne = false;
                isHasVBar = this.isHasVBar();
                isHasHBar = this.isHasHBar();
                if(isHasVBar||isHasHBar){
                    this.isShowBar = true;
                    this.startPaint(true);
                }
            //鼠标移除cnavas表隐藏滚动条
            }else if((x>this.width||y>this.height||x<0||y<0)&&!this.drawOne&&!moveHBar&&!moveVBar){
                this.drawOne = true
                isHasVBar = this.isHasVBar();
                isHasHBar = this.isHasHBar();
                if(isHasVBar||isHasHBar){
                    this.isShowBar = false;
                    var t = null;
                    clearTimeout(t);         
                    t = setTimeout(function(){
                        _this.startPaint(true)
                    },1500);   
                }
            }
            //检测鼠标位置 移动到了行列的分界线显示东西或者南北箭头
            if((x>0&&x<this.width&&y>0&&y<this.numRowH&&this.numRowH>0)||(x>0&&x<this.numColW&&y>0&&y<this.height&&this.numRowH>0)){
                this.showArrow(x,y);
                this.arrowOne = true
            }else{
                if(this.arrowOne){
                    this.isHArrow = false;
                    this.isVArrow = false;
                    this.canvas.style.cursor = "default"
                    this.arrowOne = false
                }
            }
            if(changWidth&&this.workbook.allowResize){
                this.changeCellWidth(changWidth,clickX,x)
            }
            if(changHeight&&this.workbook.allowResize){
                this.changeCellHeight(changHeight,clickY,y)
            };
            //拖拽垂直滚动条
            if(moveVBar&&this.isShowBar){
                this.dragVBar(moveVBar,clickY,moveY) ;
            };
            //拖拽水平滚动条
            if(moveHBar&&this.isShowBar){
                this.dragHBar(moveHBar,clickX,moveX)
            }
            if(moveCell&&!isArrow){
                this.getSelection(moveX,moveY,moveCell);
            };   
            //鼠标移动到button上改变样式
            var btnW = this.cellBtnAreaWidth,isFill = true,mode = this.workbook.behaviorMode,startRow = this.activeSheet.startRow,
            startCol = this.activeSheet.startCol,fixedH = this.activeSheet.fixed.fixedH,fixedW = this.activeSheet.fixed.fixedW,
            r1 = this.activeSheet.rangeRow1,r2 = this.activeSheet.rangeRow2,c1 = this.activeSheet.rangeCol1,
            c2 = this.activeSheet.rangeCol2,selectMode = this.activeSheet.selectMode,spans = this.activeSheet.spans;
            spans.forEach(function(item){
                if(c1==item.col&&r1==item.row&&c2<item.col+item.colCount&&r2<item.row+item.rowCount){
                    isFill = false
                }
            });
            //celltype == 1或者2的hover样式
            this.btnArr.forEach(function(item){
                var btnL = item.x+item.w-btnW,btnR = item.w+item.x,btnT = item.y,btnB = item.y+item.h;
                if(x>btnL&&x<btnR&&y>btnT&&y<btnB&&btnB<_this.height-_this.tabsHeight&&!(item.row==startRow&&btnT<fixedH+_this.numRowH)&&!(item.col==startCol&&btnL<fixedW+_this.numColW)){   //btn区域
                    if(item.hover===false){
                        item.hover = true;
                        _this.drawBtn(item.row,item.col,item.x,item.y,item.w,item.h,item.hover);
                    };
                }else{
                    if(item.hover===true){
                        item.hover = false;
                        if(mode==1&&(selectMode==1||selectMode==2||((r1!=r2||c1!=c2)&&isFill))){
                            _this.startPaint(true);
                        }else{
                            _this.drawBtn(item.row,item.col,item.x,item.y,item.w,item.h,item.hover);
                        }
                    }   
                };
            });
            //celltype == 3的hover样式
            this.checkBoxBtn.forEach(function(item){     
                var checkL = item.x+item.w/2-_this.checkBoxW/2,checkR = _this.checkBoxW+item.x+item.w/2-_this.checkBoxW/2,
                checkT = item.y+item.h/2-_this.checkBoxH/2,checkB = item.y+item.h/2-_this.checkBoxH/2+_this.checkBoxH;
                if(x>checkL&&x<checkR&&y>checkT&&y<checkB&&checkB<_this.height-_this.tabsHeight&&!(item.row==startRow&&checkT<fixedH+_this.numRowH)&&!(item.col==startCol&&checkL<fixedW+_this.numColW)){ 
                    if(item.hover===false){
                        item.hover = true;
                        _this.drawCheckBox(item.row,item.col,item.x,item.y,item.w,item.h,item.hover);
                    }
                }else{
                    if(item.hover===true){
                        item.hover = false;
                        if(mode==1&&(selectMode==1||selectMode==2||((r1!=r2||c1!=c2)&&isFill))){
                            _this.startPaint(true);
                        }else{
                            _this.drawCheckBox(item.row,item.col,item.x,item.y,item.w,item.h,item.hover);
                        }
                    }
                };
            });
        }.bind(this),

        //document mouseup 事件
        mouseup:function(e){
            e = e || window.event;
            if(moveVBar&&this.isShowBar){
                moveVBar = false;
                this.dragVBar(moveVBar,clickY,moveY)
            }
            if(moveHBar&&this.isShowBar){
                moveHBar = false
                this.dragHBar(moveHBar,clickX,moveX)
            }   
            if(moveCell&&!isArrow){
                moveCell = false;
                this.getSelection(moveX,moveY,moveCell);
            };  
            if(changWidth&&this.workbook.allowResize){
                changWidth = false
                this.changeCellWidth(changWidth,clickX,moveX)
            }
            if(changHeight&&this.workbook.allowResize){
                changHeight = false
                this.changeCellHeight(changHeight,clickY,moveY)
            }
        }.bind(this),

         //右击菜单事件
        contextmenu:function(e){
            if(this.workbook.showContextMenu){
                e = e || window.event;
                this.stopDefault(e)
            }  
        }.bind(this),

        keydown:function(e){
            if(window.boxId == this.boxId){
                e = e || window.event;
                var r1 = this.activeSheet.rangeRow1,c1 = this.activeSheet.rangeCol1,mode = this.workbook.behaviorMode,r2 = this.activeSheet.rangeRow2,c2 = this.activeSheet.rangeCol2;
                var cellType = this.getCellType(r1,c1,r1,c1),cellTypeName,isCanEdit = this.getCanEdit(r,c,r,c),data = this.activeSheet.data.dataTable;
                if(cellType){
                    cellTypeName = cellType.name;
                }
                if(e.keyCode==9&&this.workbook.allowTabs){    //阻止tab键默认事件
                    this.stopDefault(e)
                };
                if(this.incanvasPath){
                    //可见字符才可以调起编辑
                    if(mode==1&&!(e.keyCode==13||e.keyCode==37||e.keyCode==38||e.keyCode==39||e.keyCode==40||e.keyCode== 46||e.ctrlKey||e.keyCode==9||e.keyCode==91||e.keyCode==18||
                        e.keyCode==20||e.keyCode==27||e.keyCode==18||e.keyCode==16||e.keyCode==45||e.keyCode==255||(e.keyCode>=112&&e.keyCode<=123))&&r1!=-1&&c1!=-1){
                        if(this.keyDownCount&&!this.isEdit){    //keyDownCount（当在选中单元格，非双击，进行了键盘输入的时候，可以调起输入函数，）        
                            this.clearText(r1,c1)
                            var f = this.getCellFormula(r1,c1);
                            if(f!==undefined){
                                data[r1][c1].formula = ''
                            }
                            if(isCanEdit!==false)   this.isEdit = true;  
                            this.keyDownCount = false         
                            if((cellTypeName==1||cellTypeName==2)&&!this.activeSheet.alwaysShowButton){
                                this.startPaint()
                            }else{
                                this.openInput(r1,c1)
                            }
                        };   
                    }else{
                        this.keyDownCount = true
                    };

                    if(e.keyCode==46&&!this.isEdit&&this.workbook.allowDelete){ //delete键删除文本                  
                        this.deleteText(r1,c1,r2,c2)
                        this.startPaint(true)
                    };

                    if(e.keyCode==13||e.keyCode==37||e.keyCode==38||e.keyCode==39||e.keyCode==40||(e.keyCode==9&&this.workbook.allowTabs)){           //上下左右回车（切换单元格）
                        if(!(e.altKey&&e.keyCode==13)){
                            if(!this.isEdit){
                               this.stopDefault(e)
                            }
                            if(this.popup){
                                this.popup.closeAll()
                            }
                            this.swapCell()  
                        } 
                    };

                    if(e.ctrlKey&&e.keyCode==67){   //快捷键复制操作   
                        this.editCopy()
                    };
                    if(e.ctrlKey&&e.keyCode==88){   //快捷键剪切操作
                        this.editCut()
                    };
                    if(e.ctrlKey&&e.keyCode==86){   //快捷键粘贴操作
                        this.editPaste()
                        this.startPaint(true)
                    };
                    if(e.altKey&&e.keyCode==13){    //alt+enter换行
                        this.insertN();
                    };
                    if(e.keyCode==27){  //Esc
                        this.tempValue = {}
                        this.removeTextArea()
                    }      
                };
            }
        }.bind(this)
    }

    //为document绑定事件
    for(var key in this.documentHandle){
        this.addEvent(this.document,key,this.documentHandle[key]);
    }
    
    //child mouseup
    this.addEvent(child,'mouseup',function(){
        var txt = window.getSelection?window.getSelection():document.selection.createRange().text()
        this.clipBoardTxt = txt.toString();      
    }.bind(this));

    //child keydown
    this.addEvent(child,'keydown',function(e){
        var _this = this
        setTimeout(function() {
             /**
             * @api {null} /null onequalsign
             * @apiName onequalsign
             * @apiGroup Event
             * @apiDescription Input box input = number trigger event
             * - The function returns the current row and current column
             * @apiParam {Function} callback 
             * @apiExample {javascript} demo:
                WB.onequalsign=function(r,c){
                    // do something...
                }
             */
            if(typeof(_this.onequalsign)=='function'&&child.innerText=='='){
                _this.onequalsign(_this.activeSheet.activeRow,_this.activeSheet.activeCol)
            }
        }, 0);
    }.bind(this))

    //插入整行
    this.addEvent(addrowElement,"mousedown",function(e){
        e = e || window.event;
        if(e.button == 0){
            this.editInsert("F1ShiftRows")
            this.startPaint(true)
        }
    }.bind(this))   

    //插入整列
    this.addEvent(addcolElement,"mousedown",function(e){  
        e = e || window.event;
        if(e.button == 0){
            this.editInsert("F1ShiftCols")
            this.startPaint(true)
        }
    }.bind(this))   

    //删除行
    this.addEvent(delrowElement,"mousedown",function(e){   
        e = e || window.event;
        if(e.button == 0){
            this.editDelete("F1ShiftRows");
            this.startPaint(true)
        }
    }.bind(this))   

    //删除列
    this.addEvent(delcolElement,"mousedown",function(e){   
        e = e || window.event;
        if(e.button == 0){
            this.editDelete("F1ShiftCols") 
            this.startPaint(true)
        }
    }.bind(this));

    //复制
    this.addEvent(copy,"mousedown",function(e){
        e = e || window.event;
        if(e.button == 0){
            this.editCopy()
        }
    }.bind(this))    

    //剪切
    this.addEvent(cut,"mousedown",function(e){
        e = e || window.event;
        if(e.button == 0){
            this.editCut()
            this.startPaint(true)
        }
    }.bind(this));   

    //粘贴
    this.addEvent(paste,"mousedown",function(e){         
        e = e || window.event;
        if(e.button == 0){
            this.editPaste()
            this.startPaint(true)
        }
    }.bind(this));
    
    //左对齐
    this.addEvent(leftAlign,"mousedown",function(e){        
        e = e || window.event;
        if(e.button == 0){
            var CellRC = this.getCellRC();
            this.setHAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,2)
            this.startPaint(true)
        }
    }.bind(this)) ;  

    //居中对齐
    this.addEvent(centerAlign,"mousedown",function(e){    
        e = e || window.event;
        if( e.button == 0){
            var CellRC = this.getCellRC();
            this.setHAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,3)
            this.startPaint(true)
        }
    }.bind(this)) ; 

    //右对齐
    this.addEvent(rightAlign,"mousedown",function(e){        
        e = e || window.event;
        if(e.button == 0){
            var CellRC = this.getCellRC();
            this.setHAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,4)
            this.startPaint(true)
        }
    }.bind(this)) ; 

    //顶部对齐
    this.addEvent(topAlign,"mousedown",function(e){     
        e = e || window.event;
        if(e.button == 0){
            var CellRC = this.getCellRC();
            this.setVAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,1)
            this.startPaint(true)
        }
    }.bind(this)) ; 

    //middle对齐
    this.addEvent(middleAlign,"mousedown",function(e){       
        e = e || window.event;
        if(e.button == 0){
            var CellRC = this.getCellRC();
            this.setVAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,2)
            this.startPaint(true)
        }
    }.bind(this)) ; 

    //底部对齐
    this.addEvent(bottomAlign,"mousedown",function(e){     
        e = e || window.event;
        if(e.button == 0){
            var CellRC = this.getCellRC();
            this.setVAlignment(CellRC.r1,CellRC.c1,CellRC.r2,CellRC.c2,3)
            this.startPaint(true)
        }
    }.bind(this)) ; 
}

//是否点击了竖向滚动条
Workbook.prototype.isClickVBar = function(x,y){
    var isHasVBar = this.isHasVBar(),scrollVInfo = this.getScrollVInfo(),y1 = scrollVInfo.y1,y2 = scrollVInfo.y2,isClickVBar = false;
    if(x<this.horizontalLine&&x>this.horizontalLine-this.scrollSize&&y<y2&&y>y1&&isHasVBar&&this.isShowBar){ 
        isClickVBar = true
    }
    return isClickVBar
}

//是否点击了横向滚动条
Workbook.prototype.isClickHBar = function(x,y){
    var isHasHBar = this.isHasHBar(),scrollHInfo = this.getScrollHInfo(),x1 = scrollHInfo.x1,x2 = scrollHInfo.x2,isClickHBar = false;
    if(x>x1&&x<x2&&y<this.verticalLine-this.diffTabsH&&y>this.verticalLine-this.scrollSize-this.diffTabsH&&isHasHBar){
        isClickHBar = true;
    }
    return isClickHBar
}

//是否点击在单元格区域范围
Workbook.prototype.isClickCell = function(x,y){
    var isClickVBar = this.isClickVBar(x,y),isClickHBar = this.isClickHBar(x,y),isClickInCell = false;
    if(!isClickVBar&&!isClickHBar&&y>0&&y<this.verticalLine-this.diffTabsH&&x<this.horizontalLine&&x>0){
        isClickInCell = true;
    }
    return isClickInCell
}

//是否有竖向滚动条
Workbook.prototype.isHasVBar = function(){
    var isHas = false;
    if(this.workbook.showVScrollBar==1&&this.totolHeight+this.numRowH>this.height&&this.percentageV<1){
        isHas = true;
    }
    return isHas
}

//是否有横向滚动条
Workbook.prototype.isHasHBar = function(){
    var isHas = false;
    if(this.workbook.showHScrollBar==1&&this.totolWidth+this.numColW>this.width&&this.percentageH<1){
        isHas = true;
    }
    return isHas;
}

//是否点击在celltype==1或者celltype==2的按钮上
Workbook.prototype.isClickBtn = function(r,c,x,y){
    var cellType = this.getCellType(r,c,r,c),cellTypeName;
    if(cellType){
        cellTypeName = cellType.name;
    };
    if((cellTypeName==1||cellTypeName==2)&&this.eButton==0){
        var selectsInfo = this.getSelectInfo(r,c),c1w,r1h,top,left,cutX = this.getBtnAndCutX(r,c).cutX,
        alwaysShowButton = this.activeSheet.alwaysShowButton,clickInBtn = false,btnW = this.cellBtnAreaWidth
        c1w =  selectsInfo.c1w,r1h = selectsInfo.r1h,top = selectsInfo.selectY,left = selectsInfo.selectX;
        if(x>(left+c1w-btnW-cutX)&&x<(c1w+left-cutX)&&y>top&&y<(top+r1h)&&(alwaysShowButton||(!alwaysShowButton&&this.isEdit))){
            clickInBtn = true
        }
    };
    return clickInBtn
}

//获取垂直滚动条的位置信息(y1,y2,x1)
Workbook.prototype.getScrollVInfo = function(){
    var fH = (this.workbook.scrollMode==2)?this.activeSheet.fixed.fixedH:0,coordY = this.activeSheet.scrollOption.coordY,
    scrollHeight = 0,y1 = 0,y2 = 0,x1 = 0;

    scrollHeight = Math.floor((this.height-this.numRowH-fH)*this.percentageV);
    if(scrollHeight<this.scrollVMinHeight){
        scrollHeight = this.scrollVMinHeight;
    }
    y1 = coordY+this.numRowH+fH+this.scrollSize/2;
    if(y1>this.height-this.tabsHeight-this.scrollVMinHeight){
        y1 = this.height-this.tabsHeight-this.scrollVMinHeight+this.scrollSize/2;
    }    
    y2 = scrollHeight+y1-this.scrollSize;
    x1 = this.horizontalLine-this.scrollSize+Math.floor(this.scrollSize/2);
    return {'y1':y1,'y2':y2,'x1':x1}
}

//获取水平滚动条的位置信息(x1,x2,y1)
Workbook.prototype.getScrollHInfo = function(){
    var fW = (this.workbook.scrollMode==2)?this.activeSheet.fixed.fixedW:0,coordX = this.activeSheet.scrollOption.coordX,
    scrollWidth = 0,x1 = 0,x2 = 0,y1 = 0;

    scrollWidth = Math.floor((this.width-this.numColW-fW)*this.percentageH)
    if(scrollWidth<this.scrollHMinWidth){
        scrollWidth = this.scrollHMinWidth;
    }
    x1 = coordX+this.numColW+fW+this.scrollSize/2;
    if(x1>this.width-this.scrollHMinWidth){
        x1 = this.width-this.scrollHMinWidth+this.scrollSize/2;
    }
    x2 = scrollWidth+x1-this.scrollSize;
    y1 = this.verticalLine-this.scrollSize+Math.floor(this.scrollSize/2)-this.diffTabsH;
    return {'x1':x1,'x2':x2,'y1':y1}
}

//根据传入的范围判断是否是一个合并的单元格
Workbook.prototype.isSpanCell = function(r1,c1,r2,c2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    spans = activeSheet.spans,isSpan = false;
    spans.forEach(function(item){
        if(r1>=item.row&&r1<item.row+item.rowCount&&r2>=item.row&&r2<item.row+item.rowCount&&c1>=item.col&&c1<
        item.col+item.colCount&&c2>=item.col&&c2<item.col+item.colCount){
            isSpan = true
        }
    })
    return isSpan
}

Workbook.prototype.getMousePos = function(canvas,event){     //获取鼠标在canvas的坐标点
    var box = canvas.getBoundingClientRect();
    var x = event.clientX - box.left ;
    var y = event.clientY - box.top ;
    return {"x":x,"y":y}
}

//绘制表
Workbook.prototype.drawTable = function(){   
    if(this.workbook.stopPaintedCount===0){     //当stopPaintedCount===0才发生重绘
        var activeSheet = this.getActiveSheet(),row = activeSheet.rows,col = activeSheet.columns;
        this.totolWidth = 0;this.totolHeight = 0;
        this.workbook.width = this.width;this.workbook.height = this.height;
        this.btnArr = [];this.checkBoxBtn = [];
        this.canvas.style.width = this.width + 'px';  //输出到浏览器 可见的/最终的大小
        this.canvas.style.height = this.height+ 'px';
        this.canvas.width = this.width*this.ratio ;   //画布真实大小，不可见
        this.canvas.height = this.height*this.ratio;
        this.ctx.scale(this.ratio, this.ratio);   //放大画布
        this.canvas.style.border = (this.workbook.showWorkBookBorder)?("1px solid "+this.workbook.workBookBorderColor):"1px solid transparent";         //是否显示canvas border
        var totolAndPercent = this.getTotolAndPercent();
        this.totolHeight = totolAndPercent.totolHeight;this.totolWidth = totolAndPercent.totolWidth;
        this.percentageV = totolAndPercent.percentageV;this.percentageH = totolAndPercent.percentageH; 
        this.getStartAndEnd()    //获取更新可视化区域的起始行列;
        var startRow = activeSheet.startRow,startCol = activeSheet.startCol,endRow = activeSheet.endRow,endCol = activeSheet.endCol,
        fixedRow = activeSheet.fixed.fixedRow,fixedCol = activeSheet.fixed.fixedCol,fixedRows = activeSheet.fixed.fixedRows,
        fixedCols = activeSheet.fixed.fixedCols,fixedH = activeSheet.fixed.fixedH,fixedW = activeSheet.fixed.fixedW,
        r = activeSheet.activeRow,c = activeSheet.activeCol,mode = this.workbook.behaviorMode,isShowLine = activeSheet.showGridLines,
        gridLinesColor = activeSheet.gridLinesColor,spans = activeSheet.spans;
        if(activeSheet.selHdrTopLeft){
            activeSheet.activeRow = -1;
            activeSheet.activeCol = -1;
            activeSheet.rangeRow1 = -1;
            activeSheet.rangeCol1 = -1;
            activeSheet.rangeRow2 = row.length-1;
            activeSheet.rangeCol2 = col.length-1;
        }
        spans.forEach(function(item){
            if(item.row==r&&item.col==c){
                activeSheet.rangeRow2 = item.row+item.rowCount-1;
                activeSheet.rangeCol2 = item.col+item.colCount-1;
            };
        });
        this.horizontalLine = totolAndPercent.horizontalLine;this.verticalLine = totolAndPercent.verticalLine;
        this.scrollVLeft = totolAndPercent.scrollVLeft;
        if(mode==2){    //grid模式默认增加上下左右的padding
            if(!activeSheet.cellPadding.left)   activeSheet.cellPadding.left = 5;
            if(!activeSheet.cellPadding.right)   activeSheet.cellPadding.right = 5;
            if(!activeSheet.cellPadding.top)    activeSheet.cellPadding.top = 5;
            if(!activeSheet.cellPadding.bottom)  activeSheet.cellPadding.bottom = 5;
        };
        
        var cellTypeName = this.getCellStatus(r,c).cellTypeName
        if(cellTypeName==4||cellTypeName==3){
            this.setCanEdit(r,c,r,c,false)
        }

        if(mode==2||mode==3){
            this.isEdit = true
        }

        //初始化填充白
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.fillStyle = "#FFF";
        this.ctx.fillRect(0,0,this.width,this.height)
        this.ctx.restore();

        //区域一(非冻结行非冻结列)
        this.drawBgFill(1);  
        //为不影响border
        this.ctx.save()
        this.ctx.beginPath();
        this.ctx.strokeStyle = gridLinesColor;
        if(fixedRows>0){
            this.ctx.moveTo(this.numColW+fixedW,this.numRowH+fixedH);
            this.ctx.lineTo(this.horizontalLine,this.numRowH+fixedH)
            this.ctx.stroke()
        }
        if(fixedCols>0){
            this.ctx.moveTo(this.numColW+fixedW,this.numRowH+fixedH);
            this.ctx.lineTo(this.numColW+fixedW,this.verticalLine)
            this.ctx.stroke()
        }
        this.ctx.restore()
        if(isShowLine)  this.drawLine(startRow,startCol,endRow,endCol);
        this.drawSpans(startRow,startCol,endRow,endCol)
        this.drawCellContent(startRow,startCol,endRow,endCol)
        //区域二(非冻结行冻结列区域)
        if(fixedCols>0){
            this.drawBgFill(2);  
            this.ctx.save()
            this.ctx.beginPath();
            this.ctx.strokeStyle = gridLinesColor;
            if(fixedRows>0){
                this.ctx.moveTo(this.numColW,this.numRowH+fixedH);
                this.ctx.lineTo(this.numColW+fixedW,this.numRowH+fixedH)
                this.ctx.stroke()
            }
            if(fixedCols>0&&!(activeSheet.headeNotLineInFix)){
                this.ctx.moveTo(this.numColW+fixedW,this.numRowH);
                this.ctx.lineTo(this.numColW+fixedW,this.numRowH+fixedH)
                this.ctx.stroke()
            }
            this.ctx.restore()
            if(isShowLine)  this.drawLine(startRow,fixedCol,endRow,fixedCols-1);
            this.drawSpans(startRow,fixedCol,endRow,fixedCols-1)
            this.drawCellContent(startRow,fixedCol,endRow,fixedCols-1)
        } 
        //区域三(冻结行非冻结列)
        if(fixedRows>0){   
            this.drawBgFill(3)  
            if(isShowLine)   this.drawLine(fixedRow,startCol,fixedRows-1,endCol);
            this.drawSpans(fixedRow,startCol,fixedRows-1,endCol)             
            this.drawCellContent(fixedRow,startCol,fixedRows-1,endCol);
        }
        //区域四(冻结行冻结列)
        if(fixedCols>0&&fixedRows>0){            
            this.drawBgFill(4);
            if(isShowLine)  this.drawLine(fixedRow,fixedCol,fixedRows-1,fixedCols-1);
            this.drawSpans(fixedRow,fixedCol,fixedRows-1,fixedCols-1);
            this.drawCellContent(fixedRow,fixedCol,fixedRows-1,fixedCols-1)
        };

        //最左最顶端的线
        this.drawLeftLine()
        this.drawTopLine()
        //选中框
        this.drawSelectBox()
        //行头
        if(activeSheet.rowHeaderData.showRowHeading){
            this.drawRowHeader(startRow,endRow)
            // 冻结部分
            if(fixedRows>0) this.drawRowHeader(fixedRow,fixedRows-1)   
        }
        //行箭头
        if(this.workbook.showRowArrow){
            this.drawRowArrow()
        }
        //列头
        if(activeSheet.colHeaderData.showColHeading){            //非冻结
            this.drawColHeader(startCol,endCol)
            //冻结
            if(fixedCols>0) this.drawColHeader(fixedCol,fixedCols-1)
        }
        //左上角全选三角
        if(activeSheet.rowHeaderData.showRowHeading&&activeSheet.colHeaderData.showColHeading&&(row.length>=1||col.length>=1)){
            this.drawTriangle(activeSheet.colHeaderData.defaultDataNode.style.fillColor,gridLinesColor)
        }
        //冻结窗格线条（首行/首列/拆分）
        this.drawFixLine(fixedRows,fixedCols,fixedH,fixedW,activeSheet.showFixedLine)
        //tabs栏
        this.drawTabs()
        this.openInput(r,c)
        //滚动条重绘
        this.scrollV()             
        this.scrollH()
    }
}

//输入
Workbook.prototype.openInput = function(r,c){
    var btnAndCutX = this.getBtnAndCutX(r,c),cellStatus = this.getCellStatus(r,c),selectsInfo = this.getSelectInfo(r,c),
    mode = this.workbook.behaviorMode,activeSheet = this.activeSheet,rows = activeSheet.rows,cols = activeSheet.columns,
    disBtn = btnAndCutX.disBtn,cutX = btnAndCutX.cutX,isCheckBox = cellStatus.isCheckBox,isCanEdit = cellStatus.isCanEdit,
    isLock = cellStatus.isLock;
    if(mode==2||mode==3){
        this.isEdit = true
    }
    if(r!=-1&&c!=-1&&!isCheckBox&&isCanEdit!==false&&!isLock&&this.isEdit&&this.eButton==0&&rows.length>=1&&cols.length>=1&&window.boxId==this.boxId&&this.incanvasPath){ 
        this.setCanvasInput(selectsInfo.c1w-disBtn-2-cutX, selectsInfo.r1h-2, selectsInfo.selectY+2+this.paddingT,selectsInfo.selectX+2+this.paddingL,r, c);
    }; 
}

//获取cell状态
Workbook.prototype.getCellStatus = function(r,c,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var isCheckBox = false,isCanEdit = true,isLock = false,cellTypeName;
    if(r>=0&&c>=0){
        var cellType = this.getCellType(r,c,r,c,index);
        if(cellType){
            if(cellType.name==3){
                isCheckBox = true;
            };
            cellTypeName = cellType.name;
        };
        isCanEdit = this.getCanEdit(r,c,r,c,index);
        isLock = this.getLock(r,c,r,c,index);
    }
    return {"isCheckBox":isCheckBox,"isCanEdit":isCanEdit,"isLock":isLock,"cellTypeName":cellTypeName}
}

//获取button在滚动条移进移出的偏移量以及双击输入的情况下输入框减少的宽度
Workbook.prototype.getBtnAndCutX = function(r,c,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,disBtn = 0,cutX = 0,cellTypeName,selectsInfo = this.getSelectInfo(r,c),
    selectX = selectsInfo.selectX,c1w = selectsInfo.c1w,cellType = this.getCellType(r,c,r,c,index),isHasHBar = this.isHasHBar()
    if(cellType){
        cellTypeName = cellType.name;
    }
    if(cellTypeName==1||cellTypeName==2||cellTypeName==4){
        disBtn = this.cellBtnAreaWidth            
        if(isHasHBar&&this.isShowBar&&selectX+c1w>this.scrollVLeft&&selectX+c1w-this.cellBtnAreaWidth<this.horizontalLine){
            cutX = selectX+c1w-this.scrollVLeft
        }
    };
    return {"disBtn":disBtn,"cutX":cutX}
}

/**
 * @api {null} /null getStartSheet
 * @apiName getStartSheet
 * @apiGroup Function
 * @apiDescription Get the starting table and index in the tab bar of the current workbook
 * - WB.getStartSheet ()  
 */
Workbook.prototype.getStartSheet = function(){         
    var startW = 0,startSumW = 0;
    this.ctx.font = this.workbook.tabsOptions.font;
    for(var i = this.endSheet ;i>=0;i--){
        startW = this.ctx.measureText(this.workbook.sheetList[i]).width+this.sheetItemExtraW;  
        startSumW += startW;
        if(startSumW>this.width- 300){
            this.startSheet = i;
            break;    
        }else{
            this.startSheet = 0;
        }
    }
    return {"startSheetName":this.workbook.sheetList[this.startSheet],"i":this.startSheet}
}

/**
 * @api {null} /null getEndSheet
 * @apiName getEndSheet
 * @apiGroup Function
 * @apiDescription Get end table and index in tab bar of current workbook
 * - WB.getEndSheet ()
 */
Workbook.prototype.getEndSheet = function(){                       
    var endW = 0,endSumW = 0;
    this.endSheet = this.workbook.sheetList.length - 1;
    this.ctx.font = this.workbook.tabsOptions.font;
    for(var i = this.startSheet ;i<=this.workbook.sheetList.length-1;i++){
        endW = this.ctx.measureText(this.workbook.sheetList[i]).width+this.sheetItemExtraW;  
        endSumW += endW;
        if(endSumW>this.width- 300){
            this.endSheet = i;
            break;
        }else{
            this.endSheet = this.workbook.sheetList.length - 1;
        }
    };
    return {"endSheetName":this.workbook.sheetList[this.endSheet],"i":this.endSheet}
}

//获取总行高列宽滚动条占比
Workbook.prototype.getTotolAndPercent = function(Index){
    this.getStartAndEnd()
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var row = activeSheet.rows,col = activeSheet.columns,hideH = activeSheet.fixed.hideH,hideW = activeSheet.fixed.hideW,
    startRow = activeSheet.startRow,endRow = activeSheet.endRow,startCol = activeSheet.startCol,endCol = activeSheet.endCol,
    fH = 0 , fW = 0,totolHeight = 0,totolWidth = 0,percentageV = 1,percentageH = 1,horizontalLine = activeSheet.fixed.fixedW+this.numColW,
    verticalLine = activeSheet.fixed.fixedH+this.numRowH,scrollVLeft = 0;
    if(this.workbook.scrollMode==2){
        fH = activeSheet.fixed.fixedH;
        fW = activeSheet.fixed.fixedW;
    };
    //总行高
    row.forEach(function(item){
        if(item.bHidden&&!item.tempSize){
            item.tempSize = item.size;
        }
        if(item.bHidden){
           item.size = 0;
        }
        totolHeight+=item.size; 
    });
    //总列宽
    col.forEach(function(item){
        if(item.bHidden&&!item.tempSize){
            item.tempSize = item.size;
        }
        if(item.bHidden){
            item.size = 0;
        }
        totolWidth+=item.size
    });
    for(var i = startCol;i <= endCol;i++){
        if(col[i])   horizontalLine+=col[i].size; 
    };
    for(var j = startRow;j <= endRow;j++){
        if(row[j])  verticalLine+=row[j].size;
    };

    horizontalLine = (horizontalLine>this.width)?this.width:horizontalLine;
    verticalLine = (verticalLine>this.height)?this.height:verticalLine;
    scrollVLeft = horizontalLine - this.scrollSize;

    if(verticalLine+this.addHeight<this.height-this.tabsHeight){
        this.diffTabsH = 0;
    }else{
        this.diffTabsH = verticalLine-(this.height-this.tabsHeight);
        if(this.diffTabsH>this.tabsHeight)   this.diffTabsH = this.tabsHeight;;
    }
    
    percentageV = (this.height-this.numRowH-fH-this.tabsHeight)/(totolHeight-hideH-fH);
    percentageH = (this.width-this.numColW-fW)/(totolWidth-hideW-fW);
    if(percentageH>1)    percentageH = 1;
    if(percentageV>1)    percentageV = 1
    percentageH =  percentageH.toFixed(5)-0;
    percentageV =  percentageV.toFixed(5)-0;
    return {"totolHeight":totolHeight,"totolWidth":totolWidth,"percentageV":percentageV,"percentageH":percentageH,"horizontalLine":horizontalLine,"verticalLine":verticalLine,"scrollVLeft":scrollVLeft}
}

//绘制tabs栏
Workbook.prototype.drawTabs = function(){
    if(this.workbook.showTabs==1||this.workbook.showTabs==2){
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.fillStyle = this.workbook.tabsOptions.fillColor;
        this.ctx.fillRect(0,this.height-this.tabsHeight,this.width,this.tabsHeight);   //tab栏
        this.ctx.fill();
        if(this.workbook.showTabs==1){
            var arrowH = (this.tabArrowHeight>this.tabsHeight)?this.tabsHeight:this.tabArrowHeight,arrowW = this.tabArrowWidth,
            arrowXInfo = this.getTabArrowX(),firstX = arrowXInfo.firstX,previousX = arrowXInfo.previousX,latterX = arrowXInfo.latterX,
            lastX = arrowXInfo.lastX;
            this.ctx.strokeStyle = '#666';
            this.ctx.fillStyle = '#666'
            //第一表箭头
            this.ctx.moveTo(firstX,this.height-this.tabsHeight/2);
            this.ctx.lineTo(firstX+arrowW,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(firstX+arrowW,this.height-this.tabsHeight/2+arrowH/2);
            this.ctx.closePath();
            this.ctx.moveTo(firstX,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(firstX,this.height-this.tabsHeight/2+arrowH/2);
            //前一个表箭头
            this.ctx.moveTo(previousX+arrowW,this.height-this.tabsHeight/2);  
            this.ctx.lineTo(previousX,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(previousX,this.height-this.tabsHeight/2+arrowH/2);
            this.ctx.closePath();
            //下一个表箭头
            this.ctx.moveTo(latterX+arrowW,this.height-this.tabsHeight/2);
            this.ctx.lineTo(latterX,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(latterX,this.height-this.tabsHeight/2+arrowH/2);
            this.ctx.closePath();
            //最后表箭头
            this.ctx.moveTo(lastX+arrowW,this.height-this.tabsHeight/2);
            this.ctx.lineTo(lastX,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(lastX,this.height-this.tabsHeight/2+arrowH/2);
            this.ctx.closePath();
            this.ctx.moveTo(lastX+arrowW,this.height-this.tabsHeight/2-arrowH/2);
            this.ctx.lineTo(lastX+arrowW,this.height-this.tabsHeight/2+arrowH/2);
            this.ctx.stroke();
            this.ctx.fill();
            this.ctx.restore();
    
            this.startTabW = lastX+20;
            this.ctx.save();                    
            this.ctx.beginPath();
            var tabSumW = this.startTabW,itemW=0,textColor = this.workbook.tabsOptions.fontColor;
            for(var i = this.startSheet;i<=this.endSheet;i++){
                this.ctx.font = this.workbook.tabsOptions.font;           //这一步一定要放在measuretext前面，不然会影响测量结果
                itemW = this.ctx.measureText(this.workbook.sheetList[i]).width+this.sheetItemExtraW;    
                this.ctx.strokeStyle = this.workbook.tabsOptions.borderColor;
                this.ctx.fillStyle = this.workbook.tabsOptions.selFillColor;    //选中的填充色
                this.ctx.textAlign = "center";
                this.ctx.textBaseline = "middle";
                var t = 0
                if(i!=this.workbook.activeSheet){
                    t = 3
                    this.ctx.moveTo(tabSumW,this.height-3)     //下边框      
                    this.ctx.lineTo(tabSumW+itemW,this.height-3);   
                }
                this.ctx.moveTo(tabSumW,this.height-this.tabsHeight+t);   //左边框
                this.ctx.lineTo(tabSumW,this.height-3);  
                this.ctx.moveTo(tabSumW+itemW,this.height-3)//右边框
                this.ctx.lineTo(tabSumW+itemW,this.height-this.tabsHeight+t);           
                if(i==this.workbook.activeSheet){
                    this.ctx.fillRect(tabSumW,this.height-this.tabsHeight,itemW,this.tabsHeight-3)
                    textColor = this.workbook.tabsOptions.fontSelColor
                }else{
                    textColor =  this.workbook.tabsOptions.fontColor
                }
                this.ctx.fillStyle = textColor ;
                this.ctx.textAlign = "center";
                this.ctx.textBaseline = "middle";
                this.ctx.fillText(this.workbook.sheetList[i],tabSumW+itemW/2,this.height-this.tabsHeight/2-2);
                tabSumW += itemW;  
            };
            this.ctx.stroke();
            this.ctx.restore();
            if((this.endSheet-this.startSheet+1)<this.workbook.sheetList.length){   //结尾竖线
                this.ctx.save()
                this.ctx.beginPath()
                this.ctx.lineWidth = 2
                this.ctx.strokeStyle = "#000"
                this.ctx.moveTo(tabSumW+1,this.height-this.tabsHeight)
                this.ctx.lineTo(tabSumW+1,this.height-3)
                this.ctx.stroke()
                this.ctx.restore()
            }
            //加号
            if(this.workbook.tabsOptions.showAdd){
                this.ctx.save();
                this.ctx.beginPath();
                this.ctx.strokeStyle = "#666";
                this.ctx.lineWidth = 2;
                this.ctx.moveTo(tabSumW+10,this.height-this.tabsHeight/2);
                this.ctx.lineTo(tabSumW+20,this.height-this.tabsHeight/2);
                this.ctx.moveTo(tabSumW+15,this.height-this.tabsHeight/2-5);
                this.ctx.lineTo(tabSumW+15,this.height-this.tabsHeight/2+5)
                this.ctx.stroke();
                this.ctx.restore();
            }
            this.tabSumW = tabSumW;
        }else if(this.workbook.showTabs==2&&this.bottomBtn.length>=1){
            var btnW = this.width/this.bottomBtn.length;
            var btnX = -btnW;
            var btnY = this.height-this.tabsHeight;
            for(var i = 0;i<this.bottomBtn.length;i++){
                btnX+=btnW;
                this.ctx.save();
                this.ctx.beginPath();
                this.ctx.strokeStyle = this.workbook.tabsOptions.borderColor;
                this.ctx.moveTo(btnX,btnY);
                this.ctx.lineTo(btnX+btnW,btnY);
                if(i!=this.bottomBtn.length-1){
                    this.ctx.moveTo(btnX+btnW,btnY);
                    this.ctx.lineTo(btnX+btnW,btnY+this.tabsHeight);
                };
                this.ctx.stroke();
                this.ctx.beginPath();
                this.ctx.textAlign = "center";
                this.ctx.textBaseline = "middle";
                this.ctx.font = this.workbook.tabsOptions.font; 
                this.ctx.fillStyle = this.workbook.tabsOptions.fontColor ;
                this.ctx.fillText(this.bottomBtn[i],btnX+btnW/2,btnY+this.tabsHeight/2);
                this.ctx.restore();
            }
        }  
    }
}

//获取tabs栏四个箭头的x开始位置
Workbook.prototype.getTabArrowX = function(){
    var firstX = 0,previousX = 0,latterX = 0,lastX = 0,arrowW = this.tabArrowWidth,startX = this.tabArrowStartX,
    arrowPd = this.tabArrowPadding;

    firstX = startX;
    previousX = firstX+arrowW+arrowPd;
    latterX = previousX+arrowW+arrowPd;
    lastX = latterX+arrowW+arrowPd;
    return {'firstX':firstX,'previousX':previousX,'latterX':latterX,'lastX':lastX}
}

//绘制行的箭头指向
Workbook.prototype.drawRowArrow = function(Index){
    if(this.workbook.showRowArrow){
       var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
       r = activeSheet.activeRow,c = activeSheet.activeCol,rows = activeSheet.rows,
        showRowHeader = activeSheet.rowHeaderData.showRowHeading;
        if(showRowHeader){
            var rowInfo = this.getSelectInfo(r,c)
            var x  = 0,y = rowInfo.selectY,h = (rows[r])?rows[r].size:0,selH = rowInfo.h;
            if(h>selH)  h = selH;
            if(h>5){
                this.ctx.save();
                this.ctx.beginPath();
                this.ctx.fillStyle = this.workbook.rowArrowColor;
                this.ctx.moveTo(x,y+h/2)
                this.ctx.lineTo(x,y+h/2+4)
                this.ctx.lineTo(x+6,y+h/2)
                this.ctx.lineTo(x,y+h/2-4)
                this.ctx.closePath()
                this.ctx.fill()
            }
        }
    }
}

//绘制选中框
Workbook.prototype.drawSelectBox = function(){
    var activeSheet = this.activeSheet,r1 = activeSheet.rangeRow1,r2 = activeSheet.rangeRow2,c1 = activeSheet.rangeCol1,
    c2 = activeSheet.rangeCol2,selectMode = activeSheet.selectMode,notColNum = activeSheet.notColSelectionNum,
    notRowNum = activeSheet.notRowSelectionNum,borderColor = activeSheet.selectOption.selectBorderColor,
    fillColor = activeSheet.selectOption.selectFillColor,spans = activeSheet.spans,mode = this.workbook.behaviorMode,
    isLock = this.getLock(r1,c1),selectsInfo = this.getSelectInfo(r1,c1,r2,c2),rangeH = selectsInfo.h,rangeW = selectsInfo.w;
    if(!activeSheet.isSelectionHideBorder&&(mode==1||mode==3)&&selectsInfo.h>0&&selectsInfo.w>0||r1==-1||c1==-1){
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.lineWidth = 0.5;
        this.ctx.strokeStyle = borderColor;
        this.ctx.fillStyle = fillColor;
        if(selectMode==1&&r1!=notRowNum){
            this.ctx.rect(this.numColW,selectsInfo.selectY,this.horizontalLine-this.numColW,rangeH)
            this.ctx.fill()
        }else if(selectMode==2&&c1!=notColNum){
            this.ctx.rect(selectsInfo.selectX,this.numRowH,rangeW,this.verticalLine-this.numRowH)
            this.ctx.fill()
        }else if((selectMode!=2&&selectMode!=1||r1==-1||c1==-1)&&!(isLock)&&!(r1==notRowNum&&r1==r2)&&!(c1==notColNum&&c1==c2)){
            var isFill = true;
            this.ctx.moveTo(selectsInfo.selectX,selectsInfo.selectY)
            this.ctx.lineTo(selectsInfo.selectX,selectsInfo.selectY+rangeH)
            this.ctx.lineTo(selectsInfo.selectX+rangeW,selectsInfo.selectY+rangeH)
            this.ctx.lineTo(selectsInfo.selectX+rangeW,selectsInfo.selectY) 
            this.ctx.closePath()
            spans.forEach(function(item){
                if(c1==item.col&&r1==item.row&&c2<item.col+item.colCount&&r2<item.row+item.rowCount){
                    isFill = false
                }
            });
            if((c1!=c2||r1!=r2)&&isFill){    //是否填充选的范围
                this.ctx.fill();
            } 
        }  
        this.ctx.stroke();
        this.ctx.restore();  
    };
}

//绘制左上角三角
Workbook.prototype.drawTriangle = function(fillColor,strokeColor){
    this.ctx.save();
    this.ctx.fillStyle = fillColor;
    this.ctx.fillRect(0,0,this.numColW,this.numRowH) 
    this.ctx.beginPath();
    this.ctx.fillStyle = "#fff";
    this.ctx.moveTo(this.numColW-20,this.numRowH-3);
    this.ctx.lineTo(this.numColW-3,this.numRowH-3);
    this.ctx.lineTo(this.numColW-3,this.numRowH-20);
    this.ctx.closePath()
    this.ctx.fill();
    this.ctx.beginPath();
    this.ctx.strokeStyle = strokeColor;
    this.ctx.moveTo(0,this.numRowH);
    this.ctx.lineTo(this.numColW,this.numRowH)
    this.ctx.moveTo(this.numColW,0);
    this.ctx.lineTo(this.numColW,this.numRowH)
    this.ctx.stroke()
    this.ctx.restore()
}

//绘制冻结线
Workbook.prototype.drawFixLine = function(fixedRows,fixedCols,fixedH,fixedW,showFixedLine){
    if(fixedRows>0&&showFixedLine){
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.strokeStyle = '#3C3C3C';
        this.ctx.lineWidth = 0.5
        this.ctx.moveTo(0,this.numRowH+fixedH);
        this.ctx.lineTo(this.horizontalLine,this.numRowH+fixedH)
        this.ctx.stroke();
        this.ctx.restore()
    }
    if(fixedCols>0&&showFixedLine){
        this.ctx.save();
        this.ctx.beginPath();   
        this.ctx.strokeStyle = "#3C3C3C";
        this.ctx.moveTo(this.numColW+fixedW,0);
        this.ctx.lineTo(this.numColW+fixedW,this.verticalLine);
        this.ctx.stroke();
        this.ctx.restore();
    };
}

//绘制行头
Workbook.prototype.drawRowHeader = function(rMin,rMax,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    if(activeSheet.rowHeaderData.showRowHeading){
        var rowHeaderColor = activeSheet.rowHeaderData.defaultDataNode.style.fillColor,
        rowHeaderFontColor = activeSheet.rowHeaderData.defaultDataNode.style.fontColor,
        row = activeSheet.rows,r = rMin,fixedRows = activeSheet.fixed.fixedRows,offsetY = activeSheet.scrollOption.offsetY,
        fixedH = activeSheet.fixed.fixedH,w = activeSheet.rowHeaderData.defaultW,
        gridLinesColor = activeSheet.gridLinesColor,isFixedH = (fixedRows>0&&r<fixedRows)?0:fixedH,
        isOffsetY = (fixedRows>0&&r<fixedRows)?0:offsetY,y = this.numRowH+isFixedH-isOffsetY,x = 0,height = 0,
        h = (fixedRows>0&&r<fixedRows)?fixedH+1:this.verticalLine-this.numRowH-fixedH,
        rectY = (fixedRows>0&&r<fixedRows)?this.numRowH:this.numRowH+fixedH;
        //背景
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.font= "14px Microsoft YaHei";
        this.ctx.textAlign="center";
        this.ctx.textBaseline="middle";
        this.ctx.fillStyle = rowHeaderColor;
        this.ctx.fillRect(x,rectY,w,h);
        //横线
        this.ctx.strokeStyle = gridLinesColor;
        for(var i=rMin;i<=rMax;i++){            
            if(row[i]&&(row[i].bHidden||row[i].size<=0)){
                continue
            }
            if(row[i]) height += row[i].size;
            if(row.length>1){
                this.ctx.moveTo(x,height+y);    
                this.ctx.lineTo(w,height+y);
            }
            //文字
            this.ctx.fillStyle = rowHeaderFontColor
            if(row[i]&&row[i].size>5){
                if(row[i].name){
                    this.ctx.fillText(row[i].name,w/2,height+y-(row[i].size/2))
                }else{
                    this.ctx.fillText(i+1,w/2,height+y-(row[i].size/2))
                } 
            }
        }
        this.ctx.stroke()
        this.ctx.restore();
    }
}

//绘制列头
Workbook.prototype.drawColHeader = function(cMin,cMax,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    if(activeSheet.colHeaderData.showColHeading){
        var colHeaderColor = activeSheet.colHeaderData.defaultDataNode.style.fillColor,
        colHeaderFontColor = activeSheet.colHeaderData.defaultDataNode.style.fontColor,
        col = activeSheet.columns,c = cMin,fixedCols = activeSheet.fixed.fixedCols,
        startCol = activeSheet.startCol,offsetX = activeSheet.scrollOption.offsetX,
        fixedW = activeSheet.fixed.fixedW,h =  activeSheet.colHeaderData.defaultH,colHeaderSpans = activeSheet.colHeaderData.spans,
        gridLinesColor = activeSheet.gridLinesColor,isFixedW = (fixedCols>0&&c<fixedCols)?0:fixedW,
        isOffsetX = (fixedCols>0&&c<fixedCols)?0:offsetX,x = this.numColW+isFixedW-isOffsetX,
        y = 0,width = 0,lineSplitH = 0,textSplitH = 0,textW = 0,
        w = (fixedCols>0&&c<fixedCols)?fixedW+1:this.horizontalLine-this.numColW-fixedW,
        rectX = (fixedCols>0&&c<fixedCols)?this.numColW:this.numColW+fixedW,name = '';
        //背景
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.fillStyle = colHeaderColor;
        this.ctx.fillRect(rectX,y,w,h);
        this.ctx.font= "14px Microsoft YaHei";
        this.ctx.textAlign="center";
        this.ctx.textBaseline="middle";
        // 竖线
        this.ctx.strokeStyle = gridLinesColor;
        for(var i=cMin;i<=cMax;i++){            
            if(col[i]&&(col[i].bHidden||col[i]<=0)){
                continue
            }
            if(col[i]) width += col[i].size;

            if(colHeaderSpans&&colHeaderSpans.length!=0){
                for(var j = 0;j<colHeaderSpans.length;j++){
                    //那些列的线条需要从拆分列头的高度开始
                    if(i>=colHeaderSpans[j].col&&i<colHeaderSpans[j].colCount+colHeaderSpans[j].col-1){
                        lineSplitH = colHeaderSpans[j].height;
                        if(fixedCols>0&&i==fixedCols-1&&colHeaderSpans[j].colCount+colHeaderSpans[j].col-1<startCol){
                            lineSplitH = 0
                        }
                    };
                    //那些列的文字需要从拆分列头的高度开始
                    if(i>=colHeaderSpans[j].col&&i<colHeaderSpans[j].colCount+colHeaderSpans[j].col){
                        textSplitH = colHeaderSpans[j].height;            
                    };
                    if(i==colHeaderSpans[j].col){
                        name = colHeaderSpans[j].name; 
                        var c1 = colHeaderSpans[j].col;
                        var c2 = colHeaderSpans[j].col+colHeaderSpans[j].colCount-1
                        var textCell = this.getSelectInfo(0,colHeaderSpans[j].col,0,colHeaderSpans[j].col)
                        var textX = textCell.selectX   
                    }
                }
            }
            if(lineSplitH>=this.numRowH) lineSplitH = this.numRowH;
            if(lineSplitH<=0) lineSplitH = 0;
            if(textSplitH>=this.numRowH) textSplitH = this.numRowH;
            if(textSplitH<=0) textSplitH= 0;

            this.ctx.beginPath()
            if(col.length>=1){
                this.ctx.moveTo(x+width,y+lineSplitH);    
                this.ctx.lineTo(x+width,h);
            }
            this.ctx.stroke()

            //拆分横线
            if(textSplitH>0&&textSplitH<this.numRowH&&col[i]){
                this.ctx.beginPath();
                this.ctx.moveTo(x+width,textSplitH);
                this.ctx.lineTo(x+width-col[i].size,textSplitH)
                this.ctx.stroke();
            }

            //拆分的文字
            if(name){
                this.ctx.beginPath();
                this.ctx.fillStyle = colHeaderFontColor;
                for(var a = c1;a<=c2;a++){
                   textW += this.getSelectInfo(0,a,0,a).w
                }
                this.ctx.fillText(name,textX+textW/2,textSplitH/2)
            }
           
            //文字
            this.ctx.beginPath()
            this.ctx.fillStyle = colHeaderFontColor;
            if(col[i]&&col[i].size>5){
                var letter = this.num2ExcelLetter(i+1)
                if(col[i].name&&textSplitH>=0&&textSplitH<this.numRowH){
                    this.ctx.fillText(col[i].name,width+x-col[i].size/2,(h-textSplitH)/2+textSplitH)
                }else if(!col[i].name&&textSplitH>=0&&textSplitH<this.numRowH){
                    this.ctx.fillText(letter,width+x-col[i].size/2,(h-textSplitH)/2+textSplitH)
                };
            } 
            lineSplitH = 0;textSplitH = 0;name = '';letter = '';textW = 0;
         }
         this.ctx.restore();
    }
}

//绘制网格线
Workbook.prototype.drawLine = function(rMin,cMin,rMax,cMax,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);    
    if(activeSheet.showGridLines){
        var row = activeSheet.rows,col = activeSheet.columns,headeNotLineInFix = activeSheet.headeNotLineInFix,
        fixedRow = activeSheet.fixed.fixedRow,fixedRows = activeSheet.fixed.fixedRows,fixedCols = activeSheet.fixed.fixedCols,
        fixedH = activeSheet.fixed.fixedH,fixedW = activeSheet.fixed.fixedW,offsetX = activeSheet.scrollOption.offsetX,
        offsetY = activeSheet.scrollOption.offsetY,gridLinesColor = activeSheet.gridLinesColor,spans = activeSheet.spans,
        r = rMin,c = cMin,addR = 0,rTop = 0;
        var isFixedH = (fixedRows>0&&r<fixedRows)?0:fixedH,isFixedW = (fixedCols>0&&c<fixedCols)?0:fixedW,
        isOffsetY = (fixedRows>0&&r<fixedRows)?0:offsetY,isOffsetX = (fixedCols>0&&c<fixedCols)?0:offsetX,
        x2 = (fixedCols>0&&c<fixedCols)?fixedW+this.numColW:this.horizontalLine,y2 = (fixedRows>0&&r<fixedRows)?fixedH+this.numRowH:this.verticalLine,
        x = this.numColW+isFixedW-isOffsetX,y = this.numRowH+isFixedH-isOffsetY,w = 0,h = 0,x1 = this.numColW+isFixedW,y1 = this.numRowH+isFixedH;
        if(headeNotLineInFix&&r<fixedRows&&fixedRows>0){
            addR = fixedRows -1;
            spans.forEach(function(item){
                if(addR>=item.row&&addR<item.rowCount+item.row){
                    addR = item.row
                }
            });
            for(var i = fixedRow;i<addR;i++){
                if(row[i])   rTop+=row[i].size
            };
        }
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.strokeStyle = gridLinesColor;
        for(var i = rMin;i<=rMax;i++){               
            if(row[i])  h += row[i].size;
            //横线
            if(!(headeNotLineInFix&&i<fixedRows-2)&&!(fixedRows>0&&i==fixedRows-1)){
                this.ctx.moveTo(x1,y+h);
                this.ctx.lineTo(x2,y+h)            
            }
        };
        for(var j=cMin;j<=cMax;j++){   
            if(col[j])  w += col[j].size;
            //竖线
            if(!(fixedCols>0&&j==fixedCols-1)){
                this.ctx.moveTo(w+x,y1+rTop);
                this.ctx.lineTo(w+x,y2)
            } 
        } 
        this.ctx.stroke();
        this.ctx.restore()
    }
}

//绘制最左线
Workbook.prototype.drawLeftLine = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var fixedRows = activeSheet.fixed.fixedRows,fixedRow = activeSheet.fixed.fixedRow,row = activeSheet.rows;
    if(activeSheet.showGridLines){
        var addR = 0,rTop = 0,spans = activeSheet.spans,gridLinesColor = activeSheet.gridLinesColor;
        if(activeSheet.headeNotLineInFix&&fixedRows>0){
            addR = fixedRows -1;
            spans.forEach(function(item){
                if(addR>=item.row&&addR<item.rowCount+item.row){
                    addR = item.row
                }
            });
            for(var i = fixedRow;i<addR;i++){
                if(row[i])   rTop+=row[i].size
            };
        }
        this.ctx.save()
        this.ctx.beginPath()
        this.ctx.strokeStyle = gridLinesColor;
        this.ctx.moveTo(this.numColW,this.numRowH+rTop);
        this.ctx.lineTo(this.numColW,this.verticalLine);
        this.ctx.stroke()
        this.ctx.restore()
    }
}

//绘制最顶线
Workbook.prototype.drawTopLine = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    if(activeSheet.showGridLines){
        var gridLinesColor = activeSheet.gridLinesColor;
        this.ctx.save()
        this.ctx.beginPath()
        this.ctx.strokeStyle = gridLinesColor;
        if(!activeSheet.headeNotLineInFix){
            this.ctx.moveTo(this.numColW,this.numRowH);
            this.ctx.lineTo(this.horizontalLine,this.numRowH);
        }
        this.ctx.stroke()
        this.ctx.restore()
    }
}

//绘制合并
Workbook.prototype.drawSpans = function(rMin,cMin,rMax,cMax,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var spans = activeSheet.spans,_this = this;
    this.ctx.save();
    this.ctx.beginPath();
    this.ctx.fillStyle = "#FFF"
    spans.forEach(function(item){
        var r1 = item.row,c1 = item.col,r2 = item.row+item.rowCount-1,c2 = item.col+item.colCount-1;
        if(r2>=rMin&&r1<=rMax&&c2>=cMin&&c1<=cMax){
            var selectInfo = _this.getSelectInfo(r1,c1,r2,c2)
            var w = selectInfo.w,h = selectInfo.h,x = selectInfo.selectX,y = selectInfo.selectY;
            if(h>0&&w>0){
                _this.ctx.fillRect(x+1,y+1,w-2,h-2);
            }
        }
    });
    this.ctx.restore()
}

//绘制背景填充(用于冻结)
Workbook.prototype.drawBgFill = function(area,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var fixedW = activeSheet.fixed.fixedW,fixedH = activeSheet.fixed.fixedH;
    var x = 0,y = 0, w = 0,h = 0;
    switch (area){
        case 1 :
            x = this.numColW+fixedW+1,y = this.numRowH+fixedH,w = this.horizontalLine-fixedW,h = this.verticalLine - fixedH;
            break;
        case 2 :
            x = this.numColW,y = this.numRowH+fixedH,w = fixedW-1,h = this.verticalLine - fixedH;
            break;
        case 3 :
            x = this.numColW+fixedW+1,y = this.numRowH,w = this.horizontalLine-fixedW,h = fixedH-1;
            break;
        case 4 :
            x = this.numColW,y = this.numRowH,w = fixedW-1,h = fixedH-1;
            break;
    }
    this.ctx.save();
    this.ctx.beginPath();
    this.ctx.fillStyle = "#FFF";
    this.ctx.fillRect(x,y,w,h)
    this.ctx.restore();
}

//绘制单元格内容
Workbook.prototype.drawCellContent = function(startR,startC,endR,endC,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var row = activeSheet.rows,col = activeSheet.columns,fixedRows = activeSheet.fixed.fixedRows,
    fixedCols = activeSheet.fixed.fixedCols,data = activeSheet.data.dataTable,allRow = 0,allCol = 0;
    var moneyReg = /^0*MR?\(?A?(\+(￥|¥|\$|€|￡|₣|₩))?$/; 
    var numReg = /^0*NR?\(?A?(\+thousands)?$/;           //数值格式（包含千位分隔符）
    for(var rKey in data){
        for(var j = startR;j<parseInt(rKey);j++){       
            if(row[j]){
                allRow+=row[j].size
            };
        };
        for(var cKey in data[rKey]){
            for(var h = startC;h<parseInt(cKey);h++){    
                if(col[h]){
                    allCol+=col[h].size;
                };
            }
            var mergeCount = this.getMergeCount(parseInt(rKey),parseInt(cKey),index),rowCount = (fixedRows!=0)?0:mergeCount.rowCount-1,
            colCount = (fixedCols!=0)?0:mergeCount.colCount-1;
            if(parseInt(rKey)+rowCount>=startR&&parseInt(rKey)<=endR&&parseInt(rKey)!=-1&&parseInt(cKey)+colCount>=startC&&parseInt(cKey)<=endC){
                var cell = data[rKey][cKey];
                if(JSON.stringify(cell.style)==='{}'||!cell.style){
                    cell.style = JSON.parse(JSON.stringify(this.defaultCellStyle))
                }
                var style = cell.style,formatter =  cell.style.formatter,isRed = false,R = parseInt(rKey),C = parseInt(cKey);
                var text = this.getText(R,C,R,C,index)
                if((moneyReg.test(formatter)||numReg.test(formatter))&&formatter.indexOf("R")!=-1&&Number(text)<0){
                    isRed = true
                }            
                if(formatter&&(cell.hasOwnProperty('text')||cell.hasOwnProperty('value'))){
                    text = this.getFormatValue(text,formatter);  
                };
                this.initValue(text,R,C,allRow+this.numRowH,allCol+this.numColW,style,isRed,index) 
            } 
            allCol = 0;   
        };
        allRow = 0;
    }
}

//重算percentage 
Workbook.prototype.recalculatePercentage = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var hideH = activeSheet.fixed.hideH,hideW = activeSheet.fixed.hideW,fixedH = activeSheet.fixed.fixedH,
    fixedW = activeSheet.fixed.fixedW,fH = 0 , fW = 0;
    if(this.workbook.scrollMode==2){
        fH = fixedH;
        fW = fixedW;
    }
    this.percentageV = ((this.height-this.numRowH-fH-this.tabsHeight)/(this.totolHeight-hideH-fH));
    this.percentageH = ((this.width-this.numColW-fW)/(this.totolWidth-hideW-fW));
    if(this.percentageV>1) this.percentageV = 1;
    if(this.percentageH>1) this.percentageH = 1;
    this.percentageV = this.percentageV.toFixed(5)-0;
    this.percentageH = this.percentageH.toFixed(5)-0
    this.dragVBar(false,0,0)
    this.dragHBar(false,0,0)
}

 /**
 * @api {null} WB.getSelectInfo(R1,C1,R2,C2,Index) getSelectInfo
 * @apiName getSelectInfo
 * @apiGroup Function
 * @apiDescription Get the coordinate width and height information of the selected cell
 * - Without parameters, the coordinate width and height of the currently selected range are obtained by default
 * - The width and height returned is the width and height seen
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.getSelectInfo = function(R1,C1,R2,C2,Index){             
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var rows = activeSheet.rows,cols = activeSheet.columns,startRow = activeSheet.startRow,startCol = activeSheet.startCol,
    fixedRow = activeSheet.fixed.fixedRow,fixedCol = activeSheet.fixed.fixedCol,fixedRows = activeSheet.fixed.fixedRows,
    fixedCols = activeSheet.fixed.fixedCols,offsetX = activeSheet.scrollOption.offsetX,offsetY = activeSheet.scrollOption.offsetY,
    fixedH = activeSheet.fixed.fixedH,fixedW = activeSheet.fixed.fixedW,spans = activeSheet.spans;
    var rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    var isFixedH = (fixedRows>0&&r1<fixedRows)?0:fixedH,isFixedW = (fixedCols>0&&c1<fixedCols)?0:fixedW,
    isOffsetY = (fixedRows>0&&r1<=startRow)?0:offsetY,isOffsetX = (fixedCols>0&&c1<=startCol)?0:offsetX,
    startR = (fixedRows>0&&r1<fixedRows)?fixedRow:startRow,startC = (fixedCols>0&&c1<fixedCols)?fixedCol:startCol,
    selectX = this.numColW+isFixedW-isOffsetX,selectY = this.numRowH+isFixedH-isOffsetY;
    spans.forEach(function(item){
        if(item.row==r1&&item.col==c1&&C2==undefined&&R2==undefined){
            c2 = item.col+item.colCount-1;
            r2 = item.row+item.rowCount-1;
        }
    });
    var fw = 0,w = 0,fh = 0 ,h = 0;
    //范围宽度
    var sc = (c1>startC)?c1:startC;
    var ec =  (c2>=fixedCols)?fixedCols-1:c2
    for(var i = sc ;i<=ec;i++){
        if(cols[i])   fw += cols[i].size 
    };
    var sc2 = (c1>startCol)?c1:startCol;
    for(var j = sc2 ;j<=c2;j++){
        if(cols[j])   w += cols[j].size 
    };
    //范围高度
    var sr = (r1>startR)?r1:startR;
    var er = (r2>=fixedRows)?fixedRows-1:r2
    for(var i = sr;i<=er;i++){
        if(rows[i])   fh += rows[i].size ;
    };
    var sr2 = (r1>startRow)?r1:startRow;
    for(var j = sr2 ;j<=r2;j++){
        if(rows[j])   h += rows[j].size 
    };
    //x坐标
    for(var a = startC;a < c1;a++){
        if(cols[a])  selectX += cols[a].size;
    };
    //y坐标
    for(var b = startR;b < r1;b++){
        if(rows[b])  selectY +=rows[b].size; 
    };

    w = ((c1>=fixedCols&&c1<=startCol&&fixedCols>0)||(c1<fixedCols&&c2>=fixedCols&&fixedCols>0)||(fixedCols==0&&c1<=startCol&&c2>=startCol))?w-offsetX:w;
    h = ((r1>=fixedRows&&r1<=startRow&&fixedRows>0)||(r1<fixedRows&&r2>=fixedRows&&fixedRows>0)||(fixedRows==0&&r1<=startRow&&r2>=startRow))?h-offsetY:h;
    if(w<=0)    w = 0;
    if(h<=0)    h = 0;
    var x = (selectX<this.numColW)?this.numColW:selectX;
    var y = (selectY<this.numRowH)?this.numRowH:selectY;
    var width = (fw+w<0)?0:fw+w;
    var height = (fh+h<0)?0:fh+h;
    return {"selectX":x,"selectY":y,"w":width,"h":height,"c1w":width,"r1h":height,"r1":r1,"c1":c1,"r2":r2,"c2":c2}
}

//判断r1 c1 r2 c2传参
Workbook.prototype.handleArg = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var r1 = R1,c1 = C1,r2 = R2,c2 = C2,rows = activeSheet.rows,cols = activeSheet.columns;
    r1 = (R1===undefined)?activeSheet.rangeRow1:R1;
    c1 = (C1===undefined)?activeSheet.rangeCol1:C1;
    r2 = (R2===undefined)?activeSheet.rangeRow2:R2;
    c2 = (C2===undefined)?activeSheet.rangeCol2:C2;
    if(R2===undefined&&R1!==undefined){
        r2 = r1;
    };
    if(C2===undefined&&C1!==undefined){
        c2 = c1
    }
    if(r1<0) r1 = 0;
    if(c1<0) c1 = 0;
    if(r2<0) r2 = 0;
    if(c2<0) c2 = 0;
    if(r1>rows.length-1) r1 = rows.length-1;
    if(c1>cols.length-1) c1 = cols.length-1;
    if(r2>rows.length-1) r2 = rows.length-1;
    if(c2>cols.length-1) c2 = cols.length-1;
    var tempR,tempC;
    if(r1>r2){
        tempR = r1;
        r1 = r2;
        r2 = tempR;
    };
    if(c1>c2){
        tempC = c1;
        c1 = c2;
        c2 = tempC
    }
    return {'r1':r1,'c1':c1,'r2':r2,'c2':c2}
}

//获取表的合并信息(包含合并单元格的宽高度)
Workbook.prototype.getSpansInfo = function(Index){              
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var spans = activeSheet.spans,spansInfo = [],spansHeight = 0,spansWidth = 0,toFixedW = 0,toFixedH = 0,
    row = activeSheet.rows,col = activeSheet.columns,fixedCols = activeSheet.fixed.fixedCols,fixedRows = activeSheet.fixed.fixedRows;
    spans.forEach(function(item){
        for(var i =item.row;i<item.rowCount+item.row;i++){    
            if(row[i])  spansHeight+=row[i].size; 
        };
        for(var i =item.col;i<item.colCount+item.col;i++){    
            if(col[i])  spansWidth+=col[i].size; 
        };
        for(var k = item.col;k<fixedCols;k++){
            if(col[k])  toFixedW+=col[k].size;
        };
        for(var k = item.row;k<fixedRows;k++){
            if(row[k])   toFixedH+=row[k].size;
        }
        spansInfo.push({"r":item.row,"c":item.col,"colCount":item.colCount,"rowCount":item.rowCount,"spansHeight":spansHeight,"spansWidth":spansWidth,"toFixedW":toFixedW,"toFixedH":toFixedH})
        spansHeight = 0;
        spansWidth = 0;
        toFixedH = 0;
        toFixedW = 0;
    });
    return spansInfo
}

//获取纵向最大的滚动距离
Workbook.prototype.getYMaxScroll = function(){
    var maxY = ((this.totolHeight+this.numRowH-this.height-this.activeSheet.fixed.hideH)*this.percentageV).toFixed(2)-0;
    if(maxY<0) maxY = 0;
    return maxY
}

//获取横向最大的滚动距离
Workbook.prototype.getXMaxScroll = function(){
    var maxX = ((this.totolWidth+this.numColW-this.width-this.activeSheet.fixed.hideW)*this.percentageH).toFixed(2)-0;
    if(maxX<0) maxX = 0;
    return maxX;
}

//拖动竖向滚动
Workbook.prototype.dragVBar = function(moveVBar,clickY,moveY){    
    var activeSheet = this.activeSheet,y = 0,maxY = this.getYMaxScroll(),
    oldStartRow = activeSheet.startRow,oldOffsetY = activeSheet.scrollOption.offsetY;
  
    y = moveY - clickY + activeSheet.scrollOption.totolY;
    if(this.popup){
        this.popup.closeAll()
    }
    if(y>=0){       
        activeSheet.scrollOption.coordY = (y<=maxY)?y:maxY;    
    }else if(y<0){
        activeSheet.scrollOption.coordY = 0
    };
    if(this.percentageV >= 1){
        activeSheet.scrollOption.coordY = 0
    }
    this.getStartAndEnd()
    var startRow = activeSheet.startRow,fixedRows = activeSheet.fixed.fixedRows,fixedRow = activeSheet.fixed.fixedRow,
    offsetY = activeSheet.scrollOption.offsetY,startR = (fixedRows>0)?fixedRows:fixedRow;
    if(activeSheet.scrollOption.coordY>0&&startRow==startR&&offsetY==0){
        activeSheet.scrollOption.coordY = 0
    }
    if(!moveVBar){
        activeSheet.scrollOption.totolY =  activeSheet.scrollOption.coordY ;
    }
    if((oldStartRow!=startRow)||(offsetY!=oldOffsetY)){
        this.startPaint(true)
    };
}

//拖动横向滚动条
Workbook.prototype.dragHBar = function(moveHBar,clickX,moveX){    
    var activeSheet = this.activeSheet,x,maxX = this.getXMaxScroll(),oldStartCol = activeSheet.startCol,
    oldOffsetX = activeSheet.scrollOption.offsetX;
    x = moveX - clickX + activeSheet.scrollOption.totolX;
    if(this.popup){
        this.popup.closeAll()
    }
    if(x>=0){           
        activeSheet.scrollOption.coordX = (x<=maxX)?x:maxX;
    }else{
        activeSheet.scrollOption.coordX= 0
    }  
    if(this.percentageH == 1){
        activeSheet.scrollOption.coordX = 0
    }         
    this.getStartAndEnd()
    var startCol = activeSheet.startCol,fixedCols = activeSheet.fixed.fixedCols,fixedCol = activeSheet.fixed.fixedCol,
    startC= (fixedCols>0)?fixedCols:fixedCol,offsetX = activeSheet.scrollOption.offsetX;
    if(activeSheet.scrollOption.coordX>0&&startCol==startC&&offsetX==0){
        activeSheet.scrollOption.coordX = 0
    }
    if(!moveHBar){
        activeSheet.scrollOption.totolX =  activeSheet.scrollOption.coordX 
    } 
    if((oldStartCol!=startCol)||(oldOffsetX!=offsetX)){
        this.startPaint(true)
    }
}

//绘制竖向滚动条
Workbook.prototype.scrollV = function(){  
    var isHasVBar = this.isHasVBar()
    if(isHasVBar&&this.isShowBar){
        var scrollVInfo = this.getScrollVInfo(),y1 = scrollVInfo.y1,y2 = scrollVInfo.y2,x1 = scrollVInfo.x1;
        this.ctx.save()
        this.ctx.beginPath();
        this.ctx.lineCap = "round";
        this.ctx.lineWidth = this.scrollSize;
        this.ctx.strokeStyle = "rgba(200,200,200,0.7)";        
        this.ctx.moveTo(x1,y1);
        this.ctx.lineTo(x1,y2)
        this.ctx.stroke()
        this.ctx.restore()
    }
}

//绘制水平滚动条
Workbook.prototype.scrollH = function(){      
    var isHasHBar = this.isHasHBar()
    if(isHasHBar&&this.isShowBar){
        var scrollHInfo = this.getScrollHInfo(),x1 = scrollHInfo.x1,x2 = scrollHInfo.x2,y1 = scrollHInfo.y1;
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.lineCap = "round";
        this.ctx.lineWidth = this.scrollSize;
        this.ctx.strokeStyle = "rgba(200,200,200,0.7)";
        this.ctx.moveTo(x1,y1);
        this.ctx.lineTo(x2,y1)
        this.ctx.stroke()
        this.ctx.restore()
    }
}

//滚轮函数
Workbook.prototype.scrollFun = function(e){    
    e = e || window.event;
    var isHasVBar = this.isHasVBar()
    if(isHasVBar&&this.isShowBar){
        var activeSheet = this.activeSheet,maxY = this.getYMaxScroll(),rows = activeSheet.rows,oldStartRow = activeSheet.startRow,
        oldOffsetY = activeSheet.scrollOption.offsetY,end = (activeSheet.startRow-1>=0)?activeSheet.startRow-1:0,
        start = (activeSheet.startRow-3>=0)?activeSheet.startRow-3:0;
        if(this.popup){
            this.popup.closeAll()
        }
        //设定每次滚动行数为3
        if(e.wheelDelta){  //IE  
            if(e.wheelDelta>0){  //上滚  正数
                for(var i = end;i>=start;i--){
                    if(rows[i]){
                        activeSheet.scrollOption.coordY -=(rows[i].size*this.percentageV)
                    };
                    if(i==0){
                        activeSheet.scrollOption.coordY = 0
                    }
                }
                activeSheet.scrollOption.coordY = activeSheet.scrollOption.coordY.toFixed(5)-0
                if( activeSheet.scrollOption.coordY<=0){
                    activeSheet.scrollOption.coordY = 0 ;
                    activeSheet.scrollOption.totolY = 0
                }
            }
            if(e.wheelDelta<0){  //下滚  负数                   
                for(var i = activeSheet.startRow;i<activeSheet.startRow+3;i++){
                    if(rows[i]){
                        activeSheet.scrollOption.coordY += (rows[i].size*this.percentageV)
                    }
                }
                activeSheet.scrollOption.coordY = activeSheet.scrollOption.coordY.toFixed(5)-0
                if(activeSheet.scrollOption.coordY> maxY){
                    activeSheet.scrollOption.coordY=maxY
                }
                if(this.percentageV >= 1||activeSheet.scrollOption.coordY<0){
                    activeSheet.scrollOption.coordY = 0
                }
            }   
        }else if(e.detail){  //FF
            if(e.detail>0){   //下滚  正数
                for(var i = activeSheet.startRow;i<activeSheet.startRow+3;i++){
                    if(rows[i]){
                        activeSheet.scrollOption.coordY += (rows[i].size*this.percentageV)
                    }
                }
                activeSheet.scrollOption.coordY = activeSheet.scrollOption.coordY.toFixed(5)-0
                if(activeSheet.scrollOption.coordY> maxY){
                    activeSheet.scrollOption.coordY=maxY
                }
                if(this.percentageV >= 1||activeSheet.scrollOption.coordY<0){
                    activeSheet.scrollOption.coordY = 0
                }
            }
            if(e.detail<0){   //上滚   负数
                for(var i = end;i>=start;i--){
                    if(rows[i]){
                        activeSheet.scrollOption.coordY -=(rows[i].size*this.percentageV)
                    }
                    if(i == 0){
                        activeSheet.scrollOption.coordY = 0;
                    }
                }
                activeSheet.scrollOption.coordY = activeSheet.scrollOption.coordY.toFixed(5)-0
                if(activeSheet.scrollOption.coordY<=0){
                    activeSheet.scrollOption.coordY = 0 ;
                    activeSheet.scrollOption.totolY = 0
                }
            }
        };       
        this.getStartAndEnd()
        var startRow = activeSheet.startRow,fixedRows = activeSheet.fixed.fixedRows,fixedRow = activeSheet.fixed.fixedRow,
        offsetY = activeSheet.scrollOption.offsetY,startR = (fixedRows>0)?fixedRows:fixedRow;
        if(activeSheet.scrollOption.coordY>=0&&startRow==startR&&offsetY==0){
            activeSheet.scrollOption.coordY = 0
        }
        activeSheet.scrollOption.totolY = activeSheet.scrollOption.coordY; 
        if((oldStartRow!=startRow)||(offsetY!=oldOffsetY)){
            this.startPaint(true)
        }
    }   
}

//绘制文字函数（文本单元格包括边框样式，填充等等信息）
Workbook.prototype.initValue = function(value,r,c,allRow,allCol,style,isRed,ind){    
    var activeSheet = this.getActiveSheet(ind),spansInfo = this.getSpansInfo(ind),startCol = activeSheet.startCol,
    startRow = activeSheet.startRow,endCol = activeSheet.endCol,endRow = activeSheet.endRow,fixedW = activeSheet.fixed.fixedW,
    fixedH = activeSheet.fixed.fixedH,fixedCols = activeSheet.fixed.fixedCols,fixedCol = activeSheet.fixed.fixedCol,
    fixedRow = activeSheet.fixed.fixedRow,fixedRows = activeSheet.fixed.fixedRows,offsetX = activeSheet.scrollOption.offsetX,
    offsetY = activeSheet.scrollOption.offsetY,data = activeSheet.data.dataTable,ar = activeSheet.activeRow,
    ac = activeSheet.activeCol,row = activeSheet.rows,col = activeSheet.columns,rowCount = 0,colCount=0,isDraw = false,
    selfRow,selfCol,_this=this,toFixedW=0,toFixedH=0,cutselfRow = 0,cutselfCol = 0,isOffsetX=0,isOffsetY=0,isFixedH = 0,
    isFixedW = 0,isCellWidth = 0;
    var re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;

    if(row[r])  selfRow = row[r].size;
    if(col[c])  selfCol = col[c].size;

    spansInfo.forEach(function(item){
        if(item.r==r&&item.c==c){   
            selfCol = item.spansWidth ;             
            selfRow = item.spansHeight ;
            toFixedW = item.toFixedW;
            toFixedH = item.toFixedH;
            rowCount = item.rowCount-1;
            colCount = item.colCount-1;
        };
    });

    if(fixedCols!=0){
        if(c<fixedCols){                           
            if(c+colCount<fixedCols){
                cutselfCol = 0;                     
            }else{
                for(var i = fixedCols;i<startCol;i++){
                    if(col[i])  cutselfCol += col[i].size;
                };
                selfCol = selfCol - cutselfCol - offsetX;
                selfCol = (selfCol<toFixedW)?toFixedW:selfCol
            } 
        }else{
            for(var i = c;i<startCol;i++){
                if(col[i])  cutselfCol += col[i].size;
            };
            isOffsetX = offsetX;isFixedW = fixedW; 
            selfCol = selfCol - cutselfCol ;
            selfCol = (selfCol<toFixedW)?toFixedW:selfCol;
        } ;
    }else{                             
        for(var i = c;i<startCol;i++){
            if(col[i])  cutselfCol += col[i].size;
        };
        isOffsetX = offsetX; selfCol = selfCol - cutselfCol;
        selfCol = (selfCol<toFixedW)?toFixedW:selfCol;
        isFixedW = 0
    };

    if(fixedRows!=0){
        if(r<fixedRows){                           
            if(r+rowCount<fixedRows){
                cutselfRow= 0;
            }else{
                for(var i = fixedRows;i<startRow;i++){
                    if(row[i])  cutselfRow += row[i].size;
                };
                selfRow = selfRow - cutselfRow - offsetY;
                selfRow =  (selfRow<toFixedH)?toFixedH:selfRow;
            }
        }else{
            for(var i = r;i<startRow;i++){
                if(row[i])  cutselfRow += row[i].size;
            };
            isOffsetY = offsetY; isFixedH = fixedH;  
            selfRow = selfRow - cutselfRow;
            selfRow = (selfRow<toFixedH)?toFixedH:selfRow
        };
    }else{                             
        for(var i = r;i<startRow;i++){
            if(row[i])  cutselfRow += row[i].size;
        };
        isOffsetY = offsetY; selfRow = selfRow - cutselfRow;
        selfRow = (selfRow<toFixedH)?toFixedH:selfRow;
        isFixedH = 0
    };

    var cellInfo = this.getCellStatus(r,c,ind),cellTypeName = cellInfo.cellTypeName,isCheckBox = cellInfo.isCheckBox,
    showBtn = activeSheet.alwaysShowButton,showBtnInArea = activeSheet.alwaysShowInArea;

    if(cellTypeName==1||cellTypeName==2){       
        if((showBtn&&!showBtnInArea)||(showBtn&&showBtnInArea&&r==ar&&c==ac)||(!showBtn&&r==ar&&c==ac&&this.isEdit)){   
            selfCol = selfCol - this.cellBtnAreaWidth;
            isCellWidth = this.cellBtnAreaWidth;
        }
    };

    if(cellTypeName==4){
        if((showBtn&&!showBtnInArea)||(showBtn&&showBtnInArea&&r==ar&&c==ac)||(!showBtn&&r==ar&&c==ac)){
            selfCol = selfCol - this.cellBtnAreaWidth;
            isCellWidth = this.cellBtnAreaWidth;
        } 
    }

    //文字坐标
    var hOffset = -isOffsetX+isFixedW,vOffset = -isOffsetY+isFixedH;
    var leftAlign = allCol+hOffset,centerAlign = allCol+selfCol/2+hOffset,rightAlign = allCol+selfCol+hOffset,       
    topAlign = allRow+vOffset,middleAlign = allRow+selfRow/2+vOffset,bottomAlign = allRow+selfRow+vOffset;

    //符合的范围才绘制（冻结行列可视行列，冻结行非冻结列可视行列，非冻结行冻结列可视行列，非冻结行非冻结列可视行列）
    if((((r+rowCount)>=startRow&&r<=endRow&&(c+colCount)>=startCol&&c<=endCol)||((r+rowCount)>=fixedRow&&r<fixedRows&&(c+colCount)>=fixedCol&&c<fixedCols)||
    ((r+rowCount)>=fixedRow&&r<fixedRows&&(c+colCount)>=startCol&&c<=endCol)|| ((r+rowCount)>=startRow&&r<=endRow&&(c+colCount)>=fixedCol&&c<fixedCols))
    ){ 
       isDraw = true;
    }

    var isdrawBG=false,isdrawValue=false,isOwerDrawed=false,isdrawBorder = false;
    if(style&&row[r]&&!row[r].bHidden&&col[c]&&!col[c].bHidden&&isDraw){
        isdrawBG=true;
    };

    if(style&&value!=undefined&&value!=''&&row[r]&&!row[r].bHidden&&col[c]&&!col[c].bHidden&&isDraw&&!isCheckBox){
        isdrawValue=true;
    };

    if(style&&row[r]&&!row[r].bHidden&&col[c]&&!col[c].bHidden&&isDraw){
        isdrawBorder = true;
    }

    //填充色
    this.ctx.save();
    if(isdrawBG&&style.bgColor){
        var border_bottom,border_right,border_left,border_top;
        if(row[r-1]){
            border_bottom = _this.getBorder(r-1,c,r-1,c,ind);
        };
        if(row[r+1]){
            border_top = _this.getBorder(r+1,c,r+1,c,ind);
        };
        if(col[c-1]){
            border_right = _this.getBorder(r,c-1,r,c-1,ind)
        };
        if(col[c+1]){
            border_left = _this.getBorder(r,c+1,r,c+1,ind)
        };
        var cutT = (border_bottom&&border_bottom.borderBottom)?1:0;
        var cutR = (border_right&&border_right.borderRight)?1:0;
        var cutL = (border_left&&border_left.borderLeft&&fixedCols>0&&c==fixedCols-1)?2:0;
        var cutB = (border_top&&border_top.borderTop&&fixedRows>0&&r==fixedRows-1)?2:0;
        this.ctx.beginPath();
        this.ctx.fillStyle = style.bgColor;
        this.ctx.fillRect(allCol-isOffsetX+isFixedW+cutR,allRow-isOffsetY+isFixedH+cutT,selfCol+1-cutR-cutL,selfRow+1-cutT-cutB);      //（x坐标  y坐标  宽  高）
    }
    this.ctx.restore()

    /**
    * @api {null} /null ownerdraw
    * @apiName ownerdraw
    * @apiGroup Event
    * @apiDescription Draw the contents of the cell yourself 
    * - The function returns this r c index
    * - If the self-drawn cell does not exist, please call WB.createData(r,c,index).Cell text, value has no value or no attributes
    * - Please don't call startPaint function inside the function
    * @apiParam {Function} callback 
    * @apiExample {javascript} demo:
        WB.createData(0,2,0)
        WB.ownerdraw=function(content,i,j,ind){
            if(i==0&&j==2&&ind==0){
                content.setLock(i,j,i,j,true)
                var selectInfo = content.getSelectInfo(i,j,i,j,0)
                var x = selectInfo.selectX;
                var y = selectInfo.selectY;
                var w = selectInfo.w;
                var h = selectInfo.h;
                content.ctx.arc(x+w/2,y+h/2,h/2-4,0,Math.PI*2,true)
                content.ctx.moveTo(x+w/2-2,y+h/2-3)
                content.ctx.arc(x+w/2-5,y+h/2-3,3,0,Math.PI*2,true)
                content.ctx.moveTo(x+w/2+8,y+h/2-3)
                content.ctx.arc(x+w/2+5,y+h/2-3,3,0,Math.PI*2,true)
                content.ctx.moveTo(x+w/2+7,y+h/2+2)
                content.ctx.arc(x+w/2,y+h/2+2,7,0,Math.PI,false)
                content.ctx.fillStyle = "#6495ED";
                content.ctx.strokeStyle = "white"
                content.ctx.fill()
                content.ctx.stroke()
            }
            return false;
        };
    */
    if(this.ownerdraw&&typeof(this.ownerdraw)=='function'&&this.workbook.stopEventCount==0){
       if((isdrawBG||isdrawValue)&&data[r]&&data[r][c]&&!(data[r][c].text||data[r][c].value)) {
            isOwerDrawed = this.ownerdraw(this,r,c,ind);
       }
    };

    //文字以及删除线下划线
    this.ctx.save()
    if(!isOwerDrawed&&isdrawValue){
        var x=0,y=0,textAlign,textBaseline,isFillText = true,text,re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
        this.ctx.font= style.font || this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+" "+ this.workbook.defaultFontName;
        this.ctx.fillStyle = style.fontColor || "#000";
        if(isRed) this.ctx.fillStyle = "red"
        var fontList = this.getFontList(r,c,ind),textSize = fontList.originSize,lineHeight = fontList.originHeight,
        textBold = fontList.originBold,textItalic = fontList.originItalic,paddingLeft = parseInt(activeSheet.cellPadding.left),
        paddingRight = parseInt(activeSheet.cellPadding.right),paddingTop = parseInt(activeSheet.cellPadding.top),
        paddingBottom = parseInt(activeSheet.cellPadding.bottom),twoPixels = (textBold||textItalic)?2:0;
        switch(style.hAlign){
            case 1:
                text = this.getValue(r,c,r,c,ind);
                x = re.test(text)?rightAlign-paddingRight:leftAlign+paddingLeft; textAlign = re.test(text)? "right":"left";
                break;
            case 2:
                x = leftAlign+paddingLeft+twoPixels; textAlign = "left";
                break;
            case 3:
                x = centerAlign; textAlign = "center";
                break;
            case 4:
                x = rightAlign-paddingRight-twoPixels; textAlign = "right";
                break;
            default:
                x = leftAlign+paddingLeft+twoPixels; textAlign = "left";
        }

        switch(style.vAlign){
            case 1:
                y = topAlign+paddingTop; textBaseline = "top";
                break;
            case 2:
                y = middleAlign; textBaseline = "middle";
                break;
            case 3:
                y = bottomAlign-paddingBottom; textBaseline = "bottom";
                break;
            default:
                y = middleAlign; textBaseline = "middle";
        }

        if(fixedCols!=0&&c<fixedCols&&c+colCount>=fixedCols){
            isFillText = (x<allCol)?false:true;
        };
        if(fixedRows!=0&&r<fixedRows&&r+rowCount>=fixedRows){
            isFillText =  (y<allRow)?false:true;
        };

        this.ctx.textAlign = textAlign;
        this.ctx.textBaseline = textBaseline;
        var middleText,breakHeight=0,lineH = 0,newResult=[],itemWidth = 0,x1=0,x2=0;
        var result = this.breakLine(value,selfCol-paddingLeft-paddingRight,this.ctx);   
        if(style.wordWrap){
            for(var i = 0;i<result.length;i++){
                lineH+=lineHeight;
                if(lineH<=selfRow-paddingBottom-paddingTop){    //即使是自动换行的，拿到该行高度能放下的一段或者多端文本保存到一个新的数组里面
                    newResult.push(result[i])
                }
            };

            if(style.vAlign==2){    //居中（根据行数是偶数还是奇数调整垂直上的偏移量）
                middleText = Math.floor(newResult.length/2);            
                if(newResult.length%2!=0){
                    breakHeight = middleText*lineHeight
                }else{
                    breakHeight =  middleText*lineHeight-lineHeight/2
                }
            }else if(style.vAlign==3){  //底部对齐
                breakHeight = (newResult.length-1)*lineHeight;
            }else{
                breakHeight =0  //顶部对齐不用调整偏移
            };
        }else{
            newResult.push(result[0])
        }
        newResult.forEach(function(item,ind){

            if(selfCol>=_this.cellBtnAreaWidth){
                var t = item
            }else{
                isFillText = false
            }

            if(isFillText){

                _this.ctx.beginPath()
                _this.ctx.fillText(t,x,y+(lineHeight*ind)-breakHeight); 
                itemWidth = _this.ctx.measureText(item).width;
                switch(style.hAlign){
                    case 1:
                        text = _this.getText(r,c,r,c,ind);
                        x1 = re.test(text)?-itemWidth:0; x2 = re.test(text)? 0:itemWidth;
                        break;
                    case 2:
                        x1 = 0;x2=itemWidth;
                        break;
                    case 3:
                        x1 = -itemWidth/2;x2=itemWidth/2;
                        break;
                    case 4:
                        x1 = -itemWidth;x2=0;
                        break;
                    default:
                        x1 = 0;x2=itemWidth;
                };
                var lineOffset = (style.vAlign==1)?textSize/2: (style.vAlign==2)?0:(style.vAlign==3)?-textSize/2:0;
                _this.ctx.beginPath()
                _this.ctx.strokeStyle = "#000"
                if(style.fontStrikeout){                                 
                    _this.ctx.moveTo(x+x1,y+(lineHeight*ind)-breakHeight+lineOffset);
                    _this.ctx.lineTo(x+x2,y+(lineHeight*ind)-breakHeight+lineOffset)
                }
                if(style.fontUnderline){
                    _this.ctx.moveTo(x+x1,y+(lineHeight*ind)-breakHeight+lineOffset+textSize/2);
                    _this.ctx.lineTo(x+x2,y+(lineHeight*ind)-breakHeight+lineOffset+textSize/2)
                };
                _this.ctx.stroke();
            }
        });
    };
    this.ctx.restore()

    var cutX = 0
    if((cellTypeName==1||cellTypeName==2)&&isDraw){
        cutX = this.getBtnAndCutX(r,c,ind).cutX; 
        this.drawBtn(r,c,leftAlign-cutX,topAlign,selfCol+isCellWidth,selfRow,false,ind);
        this.btnArr.push({"row":r,"col":c,"x":leftAlign-cutX,"y":topAlign,"w":selfCol+isCellWidth,"h":selfRow,"hover":false});
    }; 

    if(cellTypeName==3&&isDraw){
        this.drawCheckBox(r,c,leftAlign,topAlign,selfCol,selfRow,false,ind);
        this.checkBoxBtn.push({"row":r,"col":c,"x":leftAlign,"y":topAlign,"w":selfCol,"h":selfRow,"hover":false})
    };

    if(cellTypeName==4&&isDraw){
        cutX = this.getBtnAndCutX(r,c,ind).cutX;
        this.drawSelectBtn(r,c,leftAlign-cutX,topAlign,selfCol+isCellWidth,selfRow,ind);
    }

    //边框
    if(isdrawBorder){
        if(style.borderBottom){ 
            this.drawCellBorder(style.borderBottom,leftAlign,bottomAlign,rightAlign+isCellWidth,bottomAlign)
        }
        if(style.borderTop){   
            this.drawCellBorder(style.borderTop,leftAlign,topAlign,rightAlign+isCellWidth,topAlign)
        }
        if(style.borderLeft){   
            this.drawCellBorder(style.borderLeft,leftAlign,topAlign,leftAlign,bottomAlign)
        }
        if(style.borderRight){  
            this.drawCellBorder(style.borderRight,rightAlign+isCellWidth,topAlign,rightAlign+isCellWidth,bottomAlign)
        }
    };  
}

//绘制单元格边框
Workbook.prototype.drawCellBorder = function(border,x1,y1,x2,y2){
    this.ctx.save()
    var style =border.style;
    this.ctx.beginPath();
    var bW = (style==1)?1-0.5:(style==2)?2-0.5:(style==5)?3-0.5:1-0.5; 
    var bC =  (border.color || "#000").toLowerCase()
    this.ctx.strokeStyle = bC;
    var whiteRe = /^(white)$|^#(fff|ffffff)$/
    if(whiteRe.test(bC)){
        bW = 2
    }
    this.ctx.lineWidth = bW
    if(style!=0){  
        this.ctx.moveTo(x1,y1);
        this.ctx.lineTo(x2,y2);
    }
    this.ctx.stroke()
    this.ctx.restore()
}

//换行函数（把文本切割的函数）,要自动换行调用this.wordWrap(true)
Workbook.prototype.breakLine = function(t,w,context){           
    var newResult = []
    try{
        var chr = '';
        chr = t.toString().split("");
        var temp = "",result = [],text,patt = /[\n]/,pa = new RegExp(patt);
        for(var a = 0; a < chr.length; a++){
            if(context.measureText(temp).width < w && context.measureText(temp+(chr[a])).width <= w&&!pa.test(chr[a])&&!pa.test(chr[a-1])){   //当前文本再加一个字符是否超过单元格宽度
                temp += chr[a];
            }else{     //超过则把push到数组 result   temp重新赋值（值为下一个超过宽度的字符）
                result.push(temp); //.trim()
                if(pa.test(chr[a])){
                    temp = "<br>"
                }else{
                    temp = chr[a];                                  //去除首个回车符  
                }                
            }
        }
        result.push(temp);
        if(result[0]=='') result.shift();
        for(var i = 0;i<result.length;i++){
            if(result[i]=='<br>'&&(result[i-1]!='<br>'&&result[i-1]!==undefined)&&(result[i+1]!='<br>'&&result[i+1]!==undefined)){
                continue
            }
            text = result[i]
            if(result[i]=="<br>"){
                text = ""
            }
            newResult.push(text)
        }
    }catch(e){}  
    return newResult
}

//文字插入回车换行
Workbook.prototype.insertN = function(){           
    var activeSheet = this.activeSheet,r1 = activeSheet.rangeRow1,c1 = activeSheet.rangeCol1,
    child = document.querySelector("#"+this.boxId+" .child");
    this.endEdit()
    if(child){
        this.pasteHtmlAtCaret('<br>');
        this.wordWrap(r1,c1,r1,c1,true)
    };
}

//指定光标位置插入内容
Workbook.prototype.pasteHtmlAtCaret = function(html){          
    var sel, range;
    if (window.getSelection) {         
        sel = window.getSelection();    //文本范围或者光标位置
        if (sel.getRangeAt && sel.rangeCount) {
            range = sel.getRangeAt(sel.rangeCount-1);        //返回一个包含当前选区内容的区域对象。参数索引
            range.deleteContents();         //删除range内容
            var el = document.createElement("div");    
            el.innerHTML = html;
            var frag = document.createDocumentFragment(), node, lastNode;  //createDocumentFragment创建一个新的空白的文档片段
            while ( (node = el.firstChild) ) {
                lastNode = frag.appendChild(node);
            };
            range.insertNode(frag);   //插入空白片段
            if (lastNode) {
                range = range.cloneRange();        //克隆range对象
                range.setStartAfter(lastNode);
                range.collapse(true); //折叠range； 折叠后的 Range 为空
                sel.removeAllRanges(); //从当前selection对象中移除所有的range对象,取消所有的选择
                sel.addRange(range);  //向选区（Selection）中添加一个区域（Range）
            };
        };
    } else if (document.selection && document.selection.type != "Control") {
        // IE9以下
        document.selection.createRange().pasteHTML(html);
    };
}

 /**
 * @api {null} /null onenteraddrow
 * @apiName onenteraddrow
 * @apiGroup Event
 * @apiDescription Press Enter on the last line to add a line
 * - When the onenteraddrow function returns true, add a row
 * @apiParam {Function} callback 
 * @apiExample {javascript} demo:
    WB.onenteraddrow=function(r,c,v){
        return true
    }
 */

 /**
 * @api {null} /null onafteraddrow
 * @apiName onafteraddrow
 * @apiGroup Event
 * @apiDescription Enter the event after adding a new line
 * @apiParam {Function} callback 
 * @apiExample {javascript} demo:
    WB.onafteraddrow=function(r,c){
        //do something...
    }
 */

//键盘上下左右改变选中的单元格(37左  38上  39右  40下，回车13（下）)
Workbook.prototype.swapCell = function(e){      
    e = e || window.event;
    var activeSheet = this.activeSheet,r1 = activeSheet.rangeRow1,c1 = activeSheet.rangeCol1,r2 = activeSheet.rangeRow2,
    c2 = activeSheet.rangeCol2,fixedRows = activeSheet.fixed.fixedRows,fixedCols = activeSheet.fixed.fixedCols,
    endRow = activeSheet.endRow,endCol = activeSheet.endCol,startRow = activeSheet.startRow,startCol = activeSheet.startCol,
    spans = activeSheet.spans,child = document.querySelector("#"+this.boxId+" .child"),text='',isFocus = this.getFocus();  
    if(child){
        text = child.innerText;
    };  
    if(e.keyCode==37&&c1!=0){   //左
        if((window.getSelection().anchorOffset==0)||!isFocus){
            activeSheet.rangeCol1 -= 1;
            activeSheet.activeCol -= 1;
            activeSheet.rangeCol2 = activeSheet.rangeCol1;
            activeSheet.rangeRow2 = activeSheet.rangeRow1;
            this.removeTextArea()    //切换单元格的时候隐藏上一选中单元格的输入框
        } ;
    }else if(e.keyCode == 38&&r1!=0){   //上
        activeSheet.rangeRow1 -= 1;
        activeSheet.activeRow -= 1;
        activeSheet.rangeRow2 = activeSheet.rangeRow1;
        activeSheet.rangeCol2 = activeSheet.rangeCol1;
        this.removeTextArea()    
    }else if((e.keyCode==39||e.keyCode==9||(this.enterToRight&&e.keyCode==13))&&c2<activeSheet.columns.length-1){   //处于编辑状态则要判断是否到了最后一个字符才切换单元格并失去编辑焦点||未处于编辑状态的直接切换单元格
        if((window.getSelection().anchorOffset==text.length)||!isFocus||(this.enterToRight&&e.keyCode==13)){  
            activeSheet.rangeCol2 += 1;
            activeSheet.rangeCol1 = activeSheet.rangeCol2;
            activeSheet.activeCol = activeSheet.rangeCol1;
            activeSheet.rangeRow2 = activeSheet.rangeRow1;
            this.removeTextArea(e) 
        }
    }else if((e.keyCode==40||(e.keyCode==13&&!(e.altKey&&e.keyCode==13)&&((!this.enterToRight)||(this.enterToRight&&c2>=activeSheet.columns.length-1))))&&r2<activeSheet.rows.length-1){
        if(this.enterToRight&&c2>=activeSheet.columns.length-1&&e.keyCode==13&&!(e.altKey&&e.keyCode==13)){
            activeSheet.rangeCol1 = 0;
            activeSheet.activeCol = 0;
            activeSheet.scrollOption.coordX = 0;
        }
        activeSheet.rangeRow2 += 1 ;
        activeSheet.rangeRow1 = activeSheet.rangeRow2;
        activeSheet.activeRow = activeSheet.rangeRow1;
        activeSheet.rangeCol2 = activeSheet.rangeCol1;  
        this.removeTextArea(e)    
    };
    //最后一行增加行(enter键)
    if(e.keyCode==13&&!(e.altKey&&e.keyCode==13)&&r2>=activeSheet.rows.length-1){
        var f = this.onenteraddrow;
        var v = child.innerText
        if(typeof(f)=='function'&&e.keyCode==13&&f(r1,c1,v)===true){
            this.removeTextArea(e)
            this.maxRow(activeSheet.rows.length+1);
            activeSheet.rangeCol1 = 0;
            activeSheet.activeCol = 0;
            activeSheet.scrollOption.coordX = 0;
            activeSheet.rangeRow2 += 1 ;
            activeSheet.rangeRow1 = activeSheet.rangeRow2;
            activeSheet.activeRow = activeSheet.rangeRow1;
            activeSheet.rangeCol2 = activeSheet.rangeCol1;  
            if(typeof(this.onafteraddrow)=='function'){
                this.onafteraddrow(activeSheet.activeRow,activeSheet.activeCol)
            }
            this.setScrollPosition(activeSheet.scrollOption.coordX/this.percentageH,99999)
        }else{
            this.removeTextArea(e);
            if(text.replace(/\s/g,'')!=''){
                this.startPaint(true)
            }
        }
    };
    var rowCount = 1,colCount = 1;
    spans.forEach(function(item){                      //遇到合并的单元格
        if(item.row<= activeSheet.rangeRow1 && activeSheet.rangeRow1 <(item.row+item.rowCount)&&activeSheet.rangeCol1>=item.col&&activeSheet.rangeCol1<(item.colCount+item.col)){
            activeSheet.rangeRow1 = item.row;
            activeSheet.activeRow = item.row;
            activeSheet.rangeCol1 = item.col;
            activeSheet.activeCol = item.col;
            activeSheet.rangeRow2 = item.row+item.rowCount-1;
            activeSheet.rangeCol2 = item.col+item.colCount-1;
            rowCount = item.rowCount;
            colCount = item.colCount;
        };
    });  
    var maxY = this.getYMaxScroll(),maxX = this.getXMaxScroll(),coordY = activeSheet.scrollOption.coordY,
    coordX = activeSheet.scrollOption.coordX,rows = activeSheet.rows,cols = activeSheet.columns;

    //下方向键选中的单元格超过最后一行，开始行作相应改变
    if(coordY<maxY&&activeSheet.rangeRow2>=endRow-1&&this.percentageV<1&&this.percentageV>0&&(e.keyCode==40||(e.keyCode==13&&!(e.altKey&&e.keyCode==13)&&(!this.enterToRight)||(this.enterToRight&&c2>=activeSheet.columns.length-1)))){   
        var downH1 = 0;
        for(var i = startRow;i<startRow+rowCount;i++){
            if(rows[i]){
                downH1 += rows[i].size*this.percentageV;
            }
        }
        var downH2 = 0;
        for(var j = activeSheet.rangeRow1;j<=activeSheet.rangeRow2+1;j++){
            if(rows[j]){
                downH2 += rows[j].size*this.percentageV
            }
        }
        downH1 = downH1.toFixed(5)-0;
        downH2 = downH2.toFixed(5)-0;
        activeSheet.scrollOption.coordY += (downH1>=downH2)?downH1:downH2; 
        if(activeSheet.scrollOption.coordY>=maxY){
            activeSheet.scrollOption.coordY = maxY;            
        };
        activeSheet.scrollOption.totolY = activeSheet.scrollOption.coordY         //totolY   也要记录
    }
     //上方向键选中的单元格小于开始行，开始行作相应改变
    if(coordY>0&&activeSheet.rangeRow1<=startRow+1&&this.percentageV<1&&this.percentageV>0&&e.keyCode == 38){
        var upH1 = 0;
        for(var i = endRow;i>endRow-rowCount;i--){
            if(rows[i]){
                upH1 -= rows[i].size*this.percentageV;
            }
        }
        var upH2 = 0;
        for(var j = activeSheet.rangeRow2;j>=activeSheet.rangeRow1-1;j--){
            if(rows[j]){
                upH2 -= rows[j].size*this.percentageV
            }
        }
        upH1 = upH1.toFixed(5)-0;
        upH2 = upH2.toFixed(5)-0;
        activeSheet.scrollOption.coordY += (upH1>=upH2)?upH2:upH1; 
        if(activeSheet.scrollOption.coordY<=0){
            activeSheet.scrollOption.coordY = 0
        };
        activeSheet.scrollOption.totolY = activeSheet.scrollOption.coordY;
    };
    //右方向键选中的单元格超过最后一列，开始列作相应改变
    if(coordX<maxX&&activeSheet.rangeCol2>=endCol-1&&this.percentageH<1&&this.percentageH>0&&(e.keyCode==39||e.keyCode==9||(this.enterToRight&&e.keyCode==13))){   
        var rightW1 = 0
        for(var i = startCol;i<startCol+colCount;i++){
            if(cols[i]){
                rightW1 += cols[i].size*this.percentageH
            }
        }
        var rightW2 = 0;
        for(var j = activeSheet.rangeCol1;j<=activeSheet.rangeCol2+1;j++){
            if(cols[j]){
                rightW2 += cols[j].size*this.percentageH
            }
        }
        rightW1 = rightW1.toFixed(5)-0;
        rightW2 = rightW2.toFixed(5)-0;
        activeSheet.scrollOption.coordX += (rightW1>=rightW2)?rightW1:rightW2; 
        if( activeSheet.scrollOption.coordX>=maxX){
            activeSheet.scrollOption.coordX = maxX
        };
        activeSheet.scrollOption.totolX =  activeSheet.scrollOption.coordX ;
    };
    //左方向键选中的单元格小于开始列，开始列作相应改变
    if(coordX>0&&activeSheet.rangeCol1<=startCol+1&&this.percentageH<1&&this.percentageH>0&&e.keyCode==37){   
        var leftW1 = 0; 
        for(var i = endCol;i>endCol-colCount;i--){
            if(cols[i]){
                leftW1 -= cols[i].size*this.percentageH;
            }
        }
        var leftW2 = 0; 
        for(var j = activeSheet.rangeCol2;j>=activeSheet.rangeCol1-1;j--){
            if(cols[j]){
                leftW2 -= cols[j].size*this.percentageH
            }
        }
        leftW1 = leftW1.toFixed(5)-0;
        leftW2 = leftW2.toFixed(5)-0;
        activeSheet.scrollOption.coordX += (leftW1>=leftW2)?leftW2:leftW1; 
        if( activeSheet.scrollOption.coordX<=0){
            activeSheet.scrollOption.coordX = 0
        };
        activeSheet.scrollOption.totolX =  activeSheet.scrollOption.coordX ;
    };
    if(!(r1==activeSheet.rangeRow1&&c1==activeSheet.rangeCol1)||(startCol>fixedCols&&activeSheet.scrollOption.coordX==0)||
    (startRow>fixedRows&&activeSheet.scrollOption.coordY==0)||(endCol<cols.length-1&&activeSheet.scrollOption.coordX==maxX)||
    (endRow<rows.length-1&&activeSheet.scrollOption.coordY==maxY)){
        this.startPaint(true)
    }
}

//显示改变列宽行高的箭头
Workbook.prototype.showArrow = function(x,y){             
    var activeSheet = this.activeSheet,sumHeight = this.numRowH,sumWidth = this.numColW,startHeight = 0,startWidth = 0,
    row = activeSheet.rows,col = activeSheet.columns,startRCList = this.getStartRC(x,y),startR = startRCList.startR,
    startC = startRCList.startC,fy = startRCList.fy,fx = startRCList.fx,oy = startRCList.oy,ox = startRCList.ox,isFocus = this.getFocus();
   
    for(var i = 0;i<startR;i++){
        if(row[i])   startHeight+=row[i].size;;
    };
    for(var i = 0;i<startC;i++){
        if(col[i])  startWidth +=col[i].size;
    };
    var totolY = y + startHeight - fy + oy,totolX = x + startWidth - fx + ox;
    for(var k=0;k<row.length;k++){
        if(row[k]){
            sumHeight += row[k].size;
            if(totolY<=sumHeight&&totolY>this.numRowH){
                break;
            }
        }    
    }    
    for(var j=0;j<col.length;j++){
        if(col[j]){
            sumWidth += col[j].size;
            if(totolX<=sumWidth&&totolX>this.numColW){
                break;
            } 
        }             
    };
    if(x>this.numColW&&y>this.numRowH){
        this.canvas.style.cursor = "default"
        this.isHArrow=false
        this.isVArrow=false
    };
    if(y<=this.numRowH&&x>0&&y>0&&x<=this.width){
        if((totolX>sumWidth-6&&totolX<sumWidth||(x>this.numColW-6&&x<this.numColW&&y>0&&y<this.numRowH))&&!isFocus){
            this.canvas.style.cursor = "e-resize";
            this.isHArrow=true
        }else{
            this.canvas.style.cursor = "default"
            this.isHArrow=false
        }
        this.isVArrow=false
    } 
    if(x<this.numColW&&y>this.numRowH-4&&y<this.height-this.tabsHeight){
        if((totolY>sumHeight-6&&totolY<sumHeight||(y>this.numRowH-6&&y<this.numRowH&&x>0&&x<this.numColW))&&!isFocus){
            this.canvas.style.cursor = "n-resize"
            this.isVArrow=true;
        }else{
            this.canvas.style.cursor = "default"
            this.isVArrow=false
        }
        this.isHArrow=false
    } 
}

//点击的创建动态div
Workbook.prototype.createLine = function(createDiv){    
    var activeSheet = this.activeSheet,container = document.getElementById(this.boxId),col = activeSheet.activeCol,
    row = activeSheet.activeRow,sumWidth = this.numColW,sumHeight = this.numRowH,fixedRow = activeSheet.fixed.fixedRow,
    fixedCol = activeSheet.fixed.fixedCol,fixedRows = activeSheet.fixed.fixedRows,fixedCols = activeSheet.fixed.fixedCols,
    fixedW = 0,fixedH = 0,startC,startR,offsetX,offsetY,horLine,verLine,rows = activeSheet.rows,cols = activeSheet.columns,
    isHasHBar = this.isHasHBar(),isHasVBar = this.isHasVBar(),hBarSize = (isHasHBar&&this.isShowBar)?this.scrollSize:0,
    vBarSize = (isHasVBar&&this.isShowBar)?this.scrollSize:0;
    if(col>=fixedCols){
        fixedW = activeSheet.fixed.fixedW;
        startC = activeSheet.startCol;
        offsetX = activeSheet.scrollOption.offsetX
    }else{
        fixedW = 0;
        startC = fixedCol;
        offsetX = 0;
    };
    if(row>=fixedRows){
        fixedH = activeSheet.fixed.fixedH;
        startR = activeSheet.startRow;
        offsetY = activeSheet.scrollOption.offsetY;
    }else{
        fixedH = 0;
        startR = fixedRow;
        offsetY = 0;
    };
    for(var j = startC; j<=col ;j++){
        if(cols[j]) sumWidth += cols[j].size;
    }
    for(var j = startR; j<=row ;j++){
        if(rows[j]) sumHeight += rows[j].size;
    }
    
    if((this.totolWidth+this.numColW)>this.width){
        horLine = this.width - vBarSize;
    }else{
        horLine= this.totolWidth + this.numColW - vBarSize
    }
    if((this.totolHeight+this.numRowH)>this.height-this.tabsHeight){
        verLine = this.height - this.tabsHeight - hBarSize;
    }else{
        verLine= this.totolHeight + this.numRowH - hBarSize
    }
    var div = document.createElement("div");
    div.style.background = "green";
    div.style.position = "absolute";
    div.setAttribute("class","line");
    if(createDiv == "vertical"){       
        div.style.height = verLine+"px";
        div.style.top = 2+this.paddingT+"px";         
        div.style.left = sumWidth+fixedW-offsetX+this.paddingL+"px" 
        div.style.width = 1.5+"px"    
    }
    if(createDiv == "horizontal"){
        div.style.height = 1.5+"px";
        div.style.top = sumHeight+fixedH-offsetY+this.paddingT+"px";    
        div.style.left = 2+this.paddingL+"px"   
        div.style.width = horLine+"px"  
    }
    container.appendChild(div);
}

//改变列宽
Workbook.prototype.changeCellWidth = function(changWidth,clickX,canvasMoveX){  
    if(this.workbook.allowResize){
        var activeSheet = this.activeSheet,divLine = document.querySelector(".line"),sumWidth = this.numColW,
        col = activeSheet.activeCol,x1 = clickX,x2 = canvasMoveX,diffX = 0,fixedCol = activeSheet.fixed.fixedCol,
        fixedCols = activeSheet.fixed.fixedCols,fixedW = 0,startC=0,offsetX=0,cols = activeSheet.columns;
        if(col>=fixedCols){
            fixedW = activeSheet.fixed.fixedW;
            startC = activeSheet.startCol;
            offsetX = activeSheet.scrollOption.offsetX
        }else{
            fixedW = 0;
            startC = fixedCol;
            offsetX = 0;
        };
        for(var j = startC; j<=col ;j++){
            if(cols[j]){
                sumWidth += cols[j].size;
            }else if(col == -1){
                sumWidth = activeSheet.rowHeaderData.defaultW
            } 
        }
        diffX = x2 - x1 ;
        divLine.style.left = sumWidth+diffX+fixedW-offsetX+this.paddingL+"px";   
        var line = document.getElementsByClassName('line'),totolCut = 0;
        if(!changWidth){    //拖拽列宽 鼠标释放的时候
            this.canvas.style.cursor = "default";
            for(var i = line.length - 1; i >= 0; i--) { //清除线条（竖向div）
                line[i].parentNode.removeChild(line[i]); 
            }
            if(col!=-1&&cols[col]){
                if(Math.abs(diffX)>=cols[col].size&&diffX<=0){   //改变列宽
                    activeSheet.columns[col].bHidden = true;    //往左拖  超过diffx小于本身宽度  隐藏该列
                    totolCut = -cols[col].size
                    activeSheet.columns[col].tempSize = cols[col].size;
                    activeSheet.columns[col].size = 0;  
                    this.setNoDefaultW(col)
                }else{
                    if(cols[col+1]&&cols[col+1].bHidden&&diffX>0){
                        delete activeSheet.columns[col+1].bHidden;
                        activeSheet.columns[col+1].size =  diffX;
                        delete activeSheet.columns[col+1].tempSize
                        this.setNoDefaultW(col+1)
                    }else{
                        activeSheet.columns[col].size = cols[col].size + diffX;
                         this.setNoDefaultW(col)
                    }
                    totolCut = diffX;
                }
            }else if(col == -1){  
                if(Math.abs(diffX)>activeSheet.rowHeaderData.defaultW&&diffX<0){    //改变行头宽度
                    activeSheet.rowHeaderData.showRowHeading = false;  
                }else{
                    activeSheet.rowHeaderData.defaultW = activeSheet.rowHeaderData.defaultW + diffX;
                }
                activeSheet.selHdrTopLeft = false
            };
            this.totolWidth = this.totolWidth + totolCut;
            this.recalculatePercentage()
            this.startPaint(true);
        } 
    }
}

//是否是默认列宽
Workbook.prototype.setNoDefaultW = function(col,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    cols = activeSheet.columns,defaultW = activeSheet.defaults.colWidth;
    if(cols[col]){
        if(cols[col].size==defaultW){
            if(cols[col].hasOwnProperty('noDefaultW')){
                delete activeSheet.columns[col].noDefaultW
            }
        }else{
            activeSheet.columns[col].noDefaultW = true
        }
    }
}

//改变行高
Workbook.prototype.changeCellHeight = function(changHeight,clickY,canvasMoveY){  
    if(this.workbook.allowResize){
        var activeSheet = this.activeSheet,divLine = document.querySelector(".line"),sumHeight = this.numRowH,
        row = activeSheet.activeRow,y1 = clickY,y2 = canvasMoveY,diffY = 0,fixedRow = activeSheet.fixed.fixedRow,
        fixedRows = activeSheet.fixed.fixedRows,fixedH = 0,startR,offsetY,rows = activeSheet.rows;
        if(row>=fixedRows){
            fixedH = activeSheet.fixed.fixedH;
            startR = activeSheet.startRow;
            offsetY = activeSheet.scrollOption.offsetY
        }else{
            fixedH = 0;
            startR = fixedRow;
            offsetY = 0;
        };
        for(var j = startR; j<=row ;j++){
            if(rows[j]){
                sumHeight += rows[j].size;
            }else if(row == -1){
                sumHeight = activeSheet.colHeaderData.defaultH
            } 
        }
        diffY = y2 - y1  ;
        divLine.style.top = sumHeight+diffY+fixedH-offsetY+this.paddingT+"px"  
        var line = document.getElementsByClassName('line'),totolCut = 0;
        if(!changHeight){
            for(var i = line.length - 1; i >= 0; i--) { 
                line[i].parentNode.removeChild(line[i]); 
            }
            if(row !=-1&&rows[row]){                                  
                if(Math.abs(diffY)>=rows[row].size&&diffY<0){
                    activeSheet.rows[row].bHidden = true;
                    totolCut = -rows[row].size
                    activeSheet.rows[row].tempSize = rows[row].size;
                    activeSheet.rows[row].size = 0;  
                    this.setNoDefaultH(row)
                }else{
                    if(rows[row+1]&&rows[row+1].bHidden&&diffY>0){
                        delete activeSheet.rows[row+1].bHidden;
                        activeSheet.rows[row+1].size =  diffY;
                        delete  activeSheet.rows[row+1].tempSize
                        this.setNoDefaultH(row+1)
                    }else{
                        activeSheet.rows[row].size = rows[row].size + diffY;
                        this.setNoDefaultH(row)
                    }
                    totolCut = diffY
                }
            }else if(row == -1){                        
                if(Math.abs(diffY)> activeSheet.colHeaderData.defaultH&&diffY<0){
                    activeSheet.colHeaderData.showColHeading = false;  
                }else{
                    activeSheet.colHeaderData.defaultH = activeSheet.colHeaderData.defaultH + diffY;
                }
                activeSheet.selHdrTopLeft = false
            }
            this.totolHeight = this.totolHeight + totolCut;
            this.recalculatePercentage()
            this.startPaint(true)
        }
    }
}

//是否是默认行高
Workbook.prototype.setNoDefaultH = function(row,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rows = activeSheet.rows,defaultH = activeSheet.defaults.rowHeight;
    if(rows[row]){
        if(rows[row].size==defaultH){
            if(rows[row].hasOwnProperty('noDefaultH')){
                delete activeSheet.rows[row].noDefaultH
            }
        }else{
            activeSheet.rows[row].noDefaultH = true
        }
    }
}

//初始化右击菜单menu
Workbook.prototype.contextMenu = function(){    
    var container = document.getElementById(this.boxId),menu = document.createElement("div");
    menu.setAttribute("menu","contextmenu");
    menu.setAttribute("class","hidden")
    container.appendChild(menu);  
    var str = ""
    str += "<ul>";
    str += "<li><p>复制 剪切：</p><span class='copy'>复制&nbsp;</span><span class='cut'>剪切&nbsp;</span><span class='paste'>粘贴&nbsp;</span></li>"
    str += "<li><p>插入：</p><span class='addrow'>整行&nbsp;</span><span class='addcol'>整列&nbsp;</span></li>"
    str += "<li><p>删除：</p><span class='delrow'>整行&nbsp;</span><span class='delcol'>整列&nbsp;</span></li>"
    str += "<li><p>左右对齐：</p><span class='leftAlign'>左 &nbsp;</span><span class='centerAlign'>中 &nbsp;</span><span class='rightAlign'>右 &nbsp;</span></li>";
    str += "<li><p>上下对齐：</p><span class='topAlign'>上 &nbsp;</span><span class='middleAlign'>中 &nbsp;</span><span class='bottomAlign'>下 &nbsp;</span></li>";  
    str += "</ul>";
    menu.innerHTML = str
}

//打开右击菜单
Workbook.prototype.openContextMenu = function(e){   
    e = e || window.event;
    var btn = e.button,activeSheet = this.getActiveSheet();
    if(this.workbook.showContextMenu&&btn == 2){
        if(this.popup){
            this.popup.closeAll()
        }
        this.removeTextArea()
        var menu = document.querySelector("#"+this.boxId+" div[menu=contextmenu]");
        var pos = this.getContextPos(e)
        menu.style.left = pos.left + this.paddingL + "px";
        menu.style.top = pos.top + this.paddingT+ "px";
        menu.className = "show"
    }  
}

//获取右击菜单坐标
Workbook.prototype.getContextPos = function(e){   
    var menu = document.querySelector("#"+this.boxId+" div[menu=contextmenu]");
    var pos = this.getMousePos(this.canvas,e)
    var x = pos.x,y = pos.y,
    menuWidth = parseInt(this.getStyle(menu, 'width')),
    menuHeight = parseInt(this.getStyle(menu, 'height'));
    return{
        left:(x+menuWidth)>this.width?(this.width-menuWidth-10):x,    //考虑右击菜单是否超出canvas宽高
        top:(y+menuHeight)>this.height?(this.height-menuHeight-10):y
    }
}

 /**
 * @api {null} /null getActiveSheet
 * @apiName getActiveSheet
 * @apiGroup Function
 * @apiDescription Set or get the current table
 * - WB.getActiveSheet () 
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.getActiveSheet = function(Index){         
    var tag,activeSheet,index = (Index>=0)?Index:this.workbook.activeSheet;
    try{
        for(var i = 0;i<this.workbook.sheetList.length;i++){
            if(i == index){
                tag = this.workbook.sheetList[i];
            }
        };
        for(var k in this.workbook.sheets){
            if(tag == k){
                activeSheet = this.workbook.sheets[k];
            }
        }
    }catch(e){}
    return activeSheet;
}

 /**
 * @api {null} WB.sheetIndex(Index) sheetIndex
 * @apiName sheetIndex
 * @apiGroup Function
 * @apiDescription Set or get the current table index
 * @apiParam {Int} [Index=Current_table_index]  table index>=0&&<Total number of tables
 * @apiExample {javascript} demo:
    WB.sheetIndex(1)   //Switch to the second table in the tab bar
    WB.sheetIndex()    //Get the current table index
 */
Workbook.prototype.sheetIndex = function(Index){  
    if(Index>=0&&Index<this.workbook.sheetList.length){
        var activeSheet = this.getActiveSheet(Index);
        if(this.popup){
            this.popup.closeAll()
        }
        this.workbook.activeSheet = Index;
        this.tempValue = {}
        // this.removeTextArea();
        this.activeSheet = activeSheet;
    };
    return this.workbook.activeSheet
}

/**
 * @api {null} WB.sheetName() sheetName
 * @apiName sheetName
 * @apiGroup Function
 * @apiDescription Set or get the current worksheet name
 * @apiParam {Int} [Index=Current_table_index] table index>=0&&<Total number of tables
 * @apiParam {String} [Name=Current table name] Table Name
 * @apiExample {javascript} demo:
    WB.sheetName(1,"table two")   //Change the name of the second table to'table two'
    WB.sheetName()   //Get the current table name
 */
//设定或者获取当前工作表名
Workbook.prototype.sheetName = function(Index,Name){        
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    if(index>=0&&index<this.workbook.sheetList.length){
        this.workbook.activeSheet = index;
        var activeSheet = this.getActiveSheet(index);
        var name = Name || activeSheet.sheetName;
        var oldName = activeSheet.sheetName;
        activeSheet.sheetName = name;
        if(oldName!=name){             //无name参数 或者新name与旧nane一样则不做任何改变
            this.workbook.sheets[name] = activeSheet;
            delete this.workbook.sheets[oldName];
            for(var i = 0;i<this.workbook.sheetList.length;i++){
                if(oldName == this.workbook.sheetList[i]){
                    this.workbook.sheetList[i]=name
                };    
            }
            this.sheetIndex(index)
        }
        return activeSheet.sheetName;
    }
}

/**
 * @api {null} /null addSheet
 * @apiName addSheet
 * @apiGroup Function
 * @apiDescription Add a new table
 * - WB.addSheet(data)   Create a new table and pass in the data
 * @apiParam {Object} [data] The data of the table is optional
 */
Workbook.prototype.addSheet = function(data){
    var objLen = this.workbook.sheetList.length;
    this.endSheet++;
    if(this.typeObj(data)=='Object'){  
        var name = (data.sheetName)?data.sheetName:"Sheet"+(objLen+1)
        this.workbook.sheetList.push(name)    
        this.workbook.sheets[name] = data;
        this.initSheet(name)
    }else{
        this.workbook.sheetList.push("Sheet"+(objLen+1))    
        this.workbook.sheets["Sheet"+(objLen+1)] = this.newSheetFormat("Sheet"+(objLen+1));
    }
    this.getStartSheet()
    this.sheetIndex(objLen)
}

/**
 * @api {null} /null newWorkbook
 * @apiName newWorkbook
 * @apiGroup Function
 * @apiDescription Create a new workbook and clear the current workbook data
 * - WB.newWorkbook()   
 * @apiParam {Object} [template] Can be passed into the template
 */
Workbook.prototype.newWorkbook = function(template){
    if(this.typeObj(template)=='Object'){ 
        this.workbook = template;
        this.initWorkbook()
        this.activeSheet = this.getActiveSheet()
    }else{
        this.workbook = JSON.parse(JSON.stringify(this.defaultOptions))
        this.activeSheet = this.getActiveSheet()
    }
}

/**
 * @api {null} /null numSheets
 * @apiName numSheets
 * @apiGroup Function
 * @apiDescription Set the number of sheets in the workbook
 * @apiParam {Int} num 
 * - num>Total number of tables(Insert table after last table)
 * - num<Total number of tables&&num>=1(Delete table from behind)
 * - num<=0(Will not delete at least one table)
 * @apiExample {javascript} demo:
    WB.numSheets(3)   //The total number of tables is set to 3
 */
Workbook.prototype.numSheets = function(num){          
    var objLen = this.workbook.sheetList.length;
    if(num>=objLen){
        this.endSheet = num - 1;
        for(var i =1;i<=num-objLen;i++){
            this.workbook.sheetList.push("Sheet"+(i+objLen))    
            this.workbook.sheets["Sheet"+(i+objLen)] = this.newSheetFormat("Sheet"+(i+objLen));
            this.getStartSheet();
        }
        this.sheetIndex(num-1)
    }else if(num<objLen&&num>=1){
        this.endSheet = num - 1;
        for(var j = objLen;j>num;j--){
            this.workbook.sheetList.pop()
            delete this.workbook.sheets[this.workbook.sheetList[j-1]];
            this.getStartSheet();
        }
        if(this.workbook.activeSheet>num-1){
            this.sheetIndex(num-1)
        }else{
            this.sheetIndex(this.workbook.activeSheet)
        }
    }else {
        throw '至少留有一张表'
    };
}

/**
 * @api {null} /null deleteSheets
 * @apiName deleteSheets
 * @apiGroup Function
 * @apiDescription Delete worksheet
 * @apiParam {Int} [nSheet=Current index] Starting index
 * @apiParam {Int} [nSheets=1] Number, how many tables are deleted
 * @apiExample {javascript} demo:
    WB.deleteSheets(1,2)  //Delete the second and third tables (at least one table remains) 
 */
Workbook.prototype.deleteSheets = function(nSheet,nSheets){  
    var objLen =this.workbook.sheetList.length
    var index = (nSheet>=0)?nSheet:this.workbook.activeSheet;
    var num = nSheets || 1;
    if(num+index>objLen){
        num = objLen - index;
    }
    if(num>=objLen&&index == 0){
        throw "至少留有一张表"
    }else{
        var sheetName = ''
        for(var a = index;a<num+index;a++){
            sheetName += ('['+this.sheetName(a)+']')
        }
        this.endSheet = this.endSheet - num;
        for(var i = index;i < num+index;i++){   
            delete this.workbook.sheets[this.workbook.sheetList[i]];  
        };
        this.workbook.sheetList.splice(index,num);
        this.getStartSheet();
        if(index+num>=objLen){     
            this.sheetIndex(index-1)
        }else{
            this.sheetIndex(index)
        }
    };
}

/**
 * @api {null} /null insertSheets
 * @apiName insertSheets
 * @apiGroup Function
 * @apiDescription Insert n tables in front of the currently indexed table
 * @apiParam {Int} [nSheet=Current index] Starting index
 * @apiParam {Int} [nSheets=1] Number, how many tables to insert  nSheets>=1.
 * @apiExample {javascript} demo:
    WB.insertSheets(1,1)   //Insert a table before the second table
    WB.insertSheets()   //Insert a table before the current table
 */
Workbook.prototype.insertSheets = function(nSheet,nSheets){  
    var objLen =this.workbook.sheetList.length
    var index = (nSheet>=0)?nSheet:this.workbook.activeSheet;
    var num = nSheets ||1;
    this.endSheet = objLen + num - 1;
    for(var i = 1; i<=num;i++){
        this.workbook.sheetList.splice(index+(i-1),0,"Sheet"+(i+objLen))    
        this.workbook.sheets["Sheet"+(i+objLen)] = this.newSheetFormat("Sheet"+(i+objLen))
        this.getStartSheet();

    }   
    this.sheetIndex(index);
}

//+新建一个表的基本格式
Workbook.prototype.newSheetFormat = function(name){     
    var newSheet = JSON.parse(JSON.stringify(this.defaultOptions.sheets.Sheet1));
    newSheet.sheetName = name;
    newSheet.rowHeaderData.showRowHeading = true;
    newSheet.colHeaderData.showColHeading = true;
    var colWidth = newSheet.defaults.colWidth;
    var rowHeight = newSheet.defaults.rowHeight;
    for(var i = 0;i<this.newSheetRow;i++){
        newSheet.rows.push({"size":rowHeight})
    };
    for(var j = 0;j<this.newSheetCol;j++){
        newSheet.columns.push({"size":colWidth})
    }
    return newSheet
}

/**
 * @api {null} /null maxCol
 * @apiName maxCol
 * @apiGroup Function
 * @apiDescription Get or set the number of columns in the current worksheet
 * @apiParam {Int} [num=Current number of columns]  
 * - num>Current number of columns.If the incoming num is greater than the current number of columns, increase the number of columns
 * - 0<=num&&num<Current number of columns.Pass in num less than the current and greater than or equal to 0, then delete the number of columns (corresponding to merge, cell data clearing) corresponding to merge reduction (num<0,num=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.maxCol(10)        //The total number of columns is 10
    WB.maxCol()        //Get the total number of columns
 */
Workbook.prototype.maxCol = function(num,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var len = activeSheet.columns.length,data =  activeSheet.data.dataTable,_this = this,spans =  activeSheet.spans
    if(num<0) num = 0;
    if(num>=0){
        this.totolWidth = 0;
        if(num>len){    //传入num大于当前列数  则增加列数
            for(var i = 0;i<num-len;i++){
                activeSheet.columns.push({"size":  activeSheet.defaults.colWidth})
            }
        }else if(0<=num&&num<len){  //传入num小于当前  则删除列数（对应合并，单元格数据清除）
            for(var j = 0;j<len-num;j++){   
                activeSheet.columns.pop();
            };
            spans.forEach(function(item){
                if((item.col+item.colCount)>num&&item.col<=num){    //对应合并减少
                    item.colCount = num - item.col;   
                };       
            });
    
            for(var k = spans.length-1;k>=0;k--){
                if(spans[k].col>num||spans[k].colCount<1||spans[k].rowCount<1){ //找到符合删除的合并数据
                    activeSheet.spans.splice(k,1)
                };
            };

            var colHeaderSpans = activeSheet.colHeaderData.spans;
            if(colHeaderSpans){
                colHeaderSpans.forEach(function(item){
                    if((item.col+item.colCount)>num&&item.col<=num){    //对应合并减少
                        item.colCount = num - item.col;   
                    };       
                });
            }

            var rowkey = Object.keys(data);
            for(var q = 0 ;q<rowkey.length;q++){
                var colkey = Object.keys(data[rowkey[q]])
                colkey.forEach(function(item){              
                    if(parseInt(item)>=activeSheet.columns.length){ //行对应的列数据要删除
                        delete data[rowkey[q]][item]
                    };
                });
                var arr = Object.keys(data[rowkey[q]]);
                if(arr.length==0){  //如果数据行下是空的（没有列数据也删除）
                    delete data[rowkey[q]]
               }
            };
            var newAcCol = (num-1>=0)?num-1:0;
            var newAcRow = (activeSheet.activeRow>=0)?activeSheet.activeRow:0;
            this.col(newAcCol,index)
            this.row(newAcRow,index)
        }
        var col = activeSheet.columns;
        col.forEach(function(item){
            if(item.bHidden&&!item.tempSize){
                item.tempSize = item.size;
            }
            if(item.bHidden){
                item.size = 0;
            }
            _this.totolWidth+=item.size
        });
        this.tempValue = {}
        this.recalculatePercentage(index)
        if(num==0){
            this.removeTextArea()
        }
    }
    return activeSheet.columns.length;  
}

/**
 * @api {null} /null maxRow
 * @apiName maxRow
 * @apiGroup Function
 * @apiDescription Get or set the number of rows in the current worksheet
 * @apiParam {Int} [num=Current table rows]
 * - num>Current table rows.If the incoming num is greater than the current number of rows, the number of rows is increased
 * - 0<=num&&num<Current table rows.The incoming num is less than the current and greater than or equal to 0, then delete the number of rows (corresponding to the merge, the cell data is cleared) corresponding to the merge reduction (num<0,num=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.maxRow(9)       //The total number of rows is 9
    WB.maxRow()        //Get the total number of rows
 */
Workbook.prototype.maxRow = function(num,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    len = activeSheet.rows.length,data =  activeSheet.data.dataTable,spans =  activeSheet.spans,_this = this;
    if(num<0) num = 0;
    if(num>=0){
        this.totolHeight = 0
        if(num>len){
            for(var i = 0;i<num-len;i++){
                activeSheet.rows.push({"size":  activeSheet.defaults.rowHeight})
            }
        }else if(0<=num&&num<len){
            for(var j = 0;j<len-num;j++){   
                activeSheet.rows.pop();
            };
            spans.forEach(function(item){
                if((item.row+item.rowCount)>num&&item.row<=num){             //对应合并减少
                    item.rowCount = num - item.row;   
                };       
            });
            for(var k = spans.length-1;k>=0;k--){
                if(spans[k].row>num||spans[k].rowCount<1||spans[k].colCount<1){   //找到符合删除的合并数据
                    activeSheet.spans.splice(k,1)    
                };
            };
            var rowkey = Object.keys(data);                     //删除对应行的数据      
            for(var q = 0 ;q<rowkey.length;q++){
                if(parseInt(rowkey[q])>=activeSheet.rows.length){
                    delete data[rowkey[q]]
                }  
            }
            var newAcRow = (num-1>=0)?num-1:0;
            var newAcCol = (activeSheet.activeCol>=0)?activeSheet.activeCol:0;
            this.row(newAcRow,index)
            this.col(newAcCol,index)
        }
        var row = activeSheet.rows;
        row.forEach(function(item){
            if(item.bHidden&&!item.tempSize){
                item.tempSize = item.size;
            }
            if(item.bHidden){
               item.size = 0;
            }
            _this.totolHeight+=item.size;   //行总高
        });
        this.tempValue = {}
        this.recalculatePercentage(index)
        if(num==0){
            this.removeTextArea()
        }
    }
    return activeSheet.rows.length; 
}

/**
 * @api {null} WB.wordWrap(R1,C1,R2,C2,bool,Index) wordWrap
 * @apiName wordWrap
 * @apiGroup Function
 * @apiDescription Set whether the cells in the range are automatically wrapped (no parameters default to automatically wrap in the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} [boolean=true]  Whether to wrap
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.wordWrap (0,0,1,1,true)
 */
Workbook.prototype.wordWrap = function(R1,C1,R2,C2,bool,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,row = activeSheet.rows,b = (bool!==undefined)?bool:true;
    var rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.wordWrap = b  
            };
        };
    }
}

//自动换行并且是默认高度的行自动增加行高
Workbook.prototype.incrementRowheight = function(r,c,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var rows = activeSheet.rows,wordWrap = this.getWordWrap(r,c,r,c,index),isSpan = this.isSpanCell(r,c,r,c,index)
    if(rows[r]&&!rows[r].noDefaultH&&wordWrap&!isSpan){
        var cols = activeSheet.columns,fontList = this.getFontList(r,c,index),cellTypeName,lineHeight = fontList.originHeight,
        text = this.getText(r,c,r,c,index)||'',isCellWidth = 0,width = 0,defaultH = activeSheet.defaults.rowHeight,
        paddingLeft = parseInt(activeSheet.cellPadding.left),paddingRight = parseInt(activeSheet.cellPadding.right),
        paddingTop = parseInt(activeSheet.cellPadding.top),paddingBottom = parseInt(activeSheet.cellPadding.bottom);
        var cellType = this.getCellType(r,c,r,c);
        if(cellType){
            cellTypeName = cellType.name
        };
        if(cellTypeName==1||cellTypeName==2||cellTypeName==4){
            isCellWidth = this.cellBtnAreaWidth;
        };
        if(cols[c]) width = cols[c].size;
        this.ctx.font= this.getFont(r,c,r,c,index) || this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+" "+ this.workbook.defaultFontName;
        var result = this.breakLine(text,width-paddingRight-paddingLeft-isCellWidth,this.ctx);
        var resultLen = result.length||1;
        var h = (lineHeight*resultLen>defaultH)?lineHeight*resultLen+paddingTop+paddingBottom:defaultH;
        this.setRowHeight(h,r,r,index)
        this.recalculatePercentage(index)
    }
}

/**
 * @api {null} WB.getWordWrap(R1,C1,R2,C2,Index) getWordWrap
 * @apiName getWordWrap
 * @apiGroup Function
 * @apiDescription Get the status of whether the cells in the range are automatically wrapped (the default is to get the status of the word wrapping in the currently selected range without parameters)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getWordWrap(0,0) 
 */
Workbook.prototype.getWordWrap = function(R1,C1,R2,C2,Index){     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,wordWrapArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]&&data[i][j].style){
                var wordWrap;
                if(data[i][j].style.hasOwnProperty("wordWrap")){
                    wordWrap = data[i][j].style.wordWrap
                    wordWrapArr.push({"r":i,"c":j,"wordWrap":wordWrap})
                }
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&wordWrapArr.length==1){
        return wordWrapArr[0].wordWrap
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return wordWrapArr
    }
}

/**
 * @api {null} WB.createData(r,c,Index) createData
 * @apiName createData
 * @apiGroup Function
 * @apiDescription Initialize an empty cell (a cell with no value, no style, etc. will not be recorded, to be initialized within the range of existing rows and columns)
 * @apiParam {Int} r row(r>=0)
 * @apiParam {Int} c col(c>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.createData = function(i,j,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    spans = activeSheet.spans,data = activeSheet.data.dataTable;
    if(!data[i]){
        var rowObj ={}; 
            rowObj[j.toString()]={
                "style":JSON.parse(JSON.stringify(this.defaultCellStyle))
            };
            data[i.toString()] = rowObj;
    }else if(data[i]&&!data[i][j]){
        var colObj = {
            "style":JSON.parse(JSON.stringify(this.defaultCellStyle))
        };
        data[i.toString()][j.toString()]  = colObj
    }else if(data[i][j]&&!data[i][j].style){
        var styleObj = JSON.parse(JSON.stringify(this.defaultCellStyle))
        data[i.toString()][j.toString()].style  = styleObj
    };
    for(var k = 0;k<spans.length;k++){                  //删除多余的
        if(spans[k].row<=i&&i<(spans[k].row+spans[k].rowCount)&&spans[k].col<=j&&j<(spans[k].col+spans[k].colCount)){
            if(!(spans[k].row==i&&spans[k].col==j)){
                delete   activeSheet.data.dataTable[i][j]
            }
        }
    }
}

Workbook.prototype.getStartRC = function(x,y){
    var activeSheet = this.getActiveSheet(),offsetX = activeSheet.scrollOption.offsetX,offsetY = activeSheet.scrollOption.offsetY,
    startRow = activeSheet.startRow,startCol = activeSheet.startCol,fixedW = activeSheet.fixed.fixedW,fixedH= activeSheet.fixed.fixedH,
    fixedRow = activeSheet.fixed.fixedRow,fixedCol = activeSheet.fixed.fixedCol,fixedRows = activeSheet.fixed.fixedRows,
    fixedCols = activeSheet.fixed.fixedCols,startC=0,startR=0,fy=0,fx=0,oy=0,ox=0;
    if(fixedCols!=0&&x<(fixedW+this.numColW)&&x>0){
        startC = fixedCol;
        fx = 0;
        ox = 0;
    }else if(fixedCols!=0&&x>(fixedW+this.numColW)&&x<this.totolWidth){
        startC = startCol;
        fx = fixedW;
        ox = offsetX
    }else if(fixedCols == 0){
        startC = startCol;
        fx = 0;
        ox = offsetX;
    };

    if(fixedRows!=0&&y<(fixedH+this.numRowH)&&y>0){
        startR = fixedRow;
        fy = 0;
        oy = 0;
    }else if(fixedRows!=0&&y>(fixedH+this.numRowH)&&y<this.totolHeight){
        startR = startRow;
        fy = fixedH;
        oy = offsetY
    }else if(fixedRows == 0){
        startR = startRow;
        fy = 0;
        oy = offsetY;
    };
    return {"startR":startR,"startC":startC,"fy":fy,"fx":fx,"oy":oy,"ox":ox}
}

//根据鼠标点击位置取得rc
Workbook.prototype.getActiveCell = function(x,y){    
    var activeSheet = this.activeSheet,startRCList = this.getStartRC(x,y),startR = startRCList.startR,startC = startRCList.startC,
    r = 0,c= 0,sumY,sumX,sumHeight = this.numRowH,sumWidth = this.numColW,startHeight = 0,startWidth = 0,spans = activeSheet.spans,
    rows = activeSheet.rows,cols = activeSheet.columns,fy = startRCList.fy,fx = startRCList.fx,oy = startRCList.oy,ox = startRCList.ox;
    for(var i = 0;i<startR;i++){
        if(rows[i])  startHeight+=rows[i].size;
    };
    for(var i = 0;i<startC;i++){
        if(cols[i])   startWidth +=cols[i].size;
    }
    sumY = y + startHeight - fy + oy ;
    sumX = x + startWidth - fx + ox ;
    if(y>this.numRowH && x>this.numColW&&y<(this.totolHeight+this.numRowH)&&x<(this.totolWidth+this.numColW)){  //自定义选中单元格
        for(var q=0;q<rows.length;q++){
            if(rows[q]&&!rows[q].bHidden){
                sumHeight += rows[q].size;
                if(sumY<=sumHeight&&sumY>this.numRowH){
                    r = q;
                    break;
                }else{
                    r = activeSheet.rangeRow1;
                }
            }   
        }        
        for(var j=0;j<cols.length;j++){
            if(cols[j]){
                sumWidth += cols[j].size;
                if(sumX<=sumWidth&&sumX>this.numColW){
                    c = j;
                    break;
                }else{
                    c = activeSheet.rangeCol1;
                }  
            }                
        }
        spans.forEach(function(item){
            if(r>=item.row&&r<(item.row+item.rowCount)&&c>=item.col&&c<(item.col+item.colCount)&&r!=-1&&c!=-1){
                r = item.row;
                c = item.col;
            }
        });
        activeSheet.selHdrTopLeft = false;         
    }else if(y<this.numRowH && x>this.numColW&&x<=(this.totolWidth+this.numColW)){  //全选列
        for(var w=0;w<cols.length;w++){
            if(cols[w]){
                sumWidth += cols[w].size;
                if(sumX<=sumWidth&&sumX>this.numColW){
                    c = w;
                    break;
                }      
            }        
        }
        r = -1;
        activeSheet.selHdrTopLeft = false; 
    }else if(y>this.numRowH && x<this.numColW&&y<=(this.totolHeight+this.numRowH)){  //全选行
        for(var e=0;e<rows.length;e++){
            if(rows[e]){
                sumHeight += rows[e].size;
                if(sumY<=sumHeight&&sumY>this.numRowH){
                    r = e;
                    break;
                }
            }
        }    
        c = -1; 
        activeSheet.selHdrTopLeft = false; 
    }else if((y<this.numRowH&&x<this.numColW&&this.numRowH>0&&this.numColW>0)||activeSheet.selHdrTopLeft){  //表格全选
        c = -1;
        r = -1;
        activeSheet.selHdrTopLeft = true
    }else{  //默认（默认已选中的单元格）
        r = activeSheet.rangeRow1;
        c = activeSheet.rangeCol1;
    };
    activeSheet.activeRow = r;
    activeSheet.activeCol = c;          
    return {"r":r,"c":c}    
}

//选择范围
Workbook.prototype.getSelection = function(moveX,moveY,moveCell){  
    var activeSheet = this.activeSheet,r2 = activeSheet.rangeRow2 ,c2 = activeSheet.rangeCol2,r = activeSheet.activeRow,
    c = activeSheet.activeCol,spans = activeSheet.spans,sumY = 0,sumX = 0,sumHeight = this.numRowH,sumWidth = this.numColW,
    startHeight = 0,startWidth = 0,startRCList = this.getStartRC(moveX,moveY),startR = startRCList.startR,startC = startRCList.startC,
    fy = startRCList.fy,fx = startRCList.fx,oy = startRCList.oy,ox = startRCList.ox,rows = activeSheet.rows,cols = activeSheet.columns,
    mode = this.workbook.behaviorMode,selectMode = activeSheet.selectMode,cellTypeName,allowMoveRange = activeSheet.allowMoveRange;

    for(var i = 0;i<startR;i++){
        if(rows[i])  startHeight+=rows[i].size;
    };
    for(var i = 0;i<startC;i++){
        if(cols[i])  startWidth +=cols[i].size;
    }
    sumY = moveY + startHeight - fy + oy ;
    sumX = moveX + startWidth - fx+ ox ;

    if(moveX<this.horizontalLine&&moveY<this.verticalLine||activeSheet.selHdrTopLeft){
        for(var k=0;k<rows.length;k++){
            if(rows[k]){
                sumHeight +=rows[k].size;
            }
            if(r == -1){
                r2 = rows.length - 1
            }else{
                if(sumY<=sumHeight){
                    r2 = k; 
                    break;
                }
            }   
        }    
        for(var j=0;j<cols.length;j++){
            if(cols[j]){
                sumWidth += cols[j].size;
            }
            if(c == -1){
                c2 = cols.length -1 ;
            }else{
                if(sumX<=sumWidth){
                    c2 = j;  
                    break;
                }  
            }             
        }      
    };

    var cellStatus = this.getCellStatus(r,c),isCheckBox = cellStatus.isCheckBox,isCanEdit = cellStatus.isCanEdit,
    isLock = cellStatus.isLock,cellTypeName=cellStatus.cellTypeName,tempR,tempC;

    if(!(selectMode==3||mode==2||mode==3||cellTypeName==4||!allowMoveRange)){
        if(r>r2){
            tempR = r;r = r2;r2 = tempR;   
        }
        if(c>c2){
            tempC = c; c = c2; c2 = tempC;  
        };
    }

    //拖选遇到合并自动扩散
    for(var a = r;a<=r2;a++){
        for(var b = c;b<=c2;b++){
            spans.forEach(function(item){
                if(a>=item.row&&a<item.row+item.rowCount&&b>=item.col&&b<item.col+item.colCount&&c!=-1&r!=-1&&mode==1){
                    if(c>=item.col){
                        c = item.col;
                    };
                    if(c2<=item.col+item.colCount-1){
                        c2 = item.col+item.colCount-1
                    };
                    if(r>=item.row){
                        r = item.row;
                    };
                    if(r2<=item.row+item.rowCount-1){
                        r2 = item.row+item.rowCount-1
                    }
                }   
            })  
        }
    }

    //单元格选中模式 grid grid有选框模式 不能拖选
    if((selectMode==3||mode==2||mode==3||cellTypeName==4||!allowMoveRange)&&r!=-1&&c!=-1){
        r2 = r;
        c2 = c;
        spans.forEach(function(item){
            if(item.row<=r2&&r2<(item.row+item.rowCount)&&c2>=item.col&&c2<(item.colCount+item.col)&&r!=-1&&c!=-1){
                r2= item.row+item.rowCount-1;
                c2 = item.col+item.colCount-1;
            };
        })  ;
    };
    

    var or = activeSheet.rangeRow1,oc = activeSheet.rangeCol1,or2 = activeSheet.rangeRow2,oc2 = activeSheet.rangeCol2;
    activeSheet.rangeRow1 = r;activeSheet.rangeCol1 = c;activeSheet.rangeRow2 = r2;activeSheet.rangeCol2 = c2; 

    if(!(or==r&&oc==c&&or2==r2&&oc2==c2)||(cellTypeName==3&&r==r2&&c==c2&&!moveCell)||(JSON.stringify(this.tempValue)=='{}'
    &&(mode==2||mode==3)&&r!=-1&&c!=-1&&!isCheckBox&&isCanEdit!==false&&!isLock)){
        this.removeTextArea()
        this.startPaint(true)  
    };
    return {"r1":r,"c1":c,"r2":r2,"c2":c2}
}

 /**
 * @api {null} WB.editInsert(InsertType,R,C,N,Index) editInsert
 * @apiName editInsert
 * @apiGroup Function
 * @apiDescription Insert a new row or column in the current row or specified row
 * @apiParam {String} InsertType  
 * - F1ShiftRows:(Insert a line, move down the line)
 * - F1ShiftCols(Insert a column, Move the column to the right)
 * @apiParam {Int} R  Insert in front of which row, R>=0, R passes -1 when inserting column
 * @apiParam {Int} C  Insert in front of which column, C>=0, C insert -1 when inserting rows
 * @apiParam {Int} [N=1]  N>=1Number of rows or columns
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.editInsert ("F1ShiftRows",0,-1)  //Insert a line before the first line
    WB.editInsert ("F1ShiftCols",-1,2)  //Insert a column before the third column
 */
Workbook.prototype.editInsert = function(InsertType,R,C,N,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    defaultHeight = activeSheet.defaults.rowHeight || 30,defaultWidth = activeSheet.defaults.colWidth || 60,
    child = document.querySelector("#"+this.boxId+" .child"),newCol = {"size":defaultWidth},newRow = {"size":defaultHeight},
    textValue = activeSheet.data.dataTable,rowkey = Object.keys(textValue); 
    var r = 0 ,c = 0,n = 0 ,key,newKey;
    r = (R>=0)?R:activeSheet.rangeRow1;
    c = (C>=0)?C:activeSheet.rangeCol1;
    n = (N>=1)?N:1;
    if(InsertType=="F1ShiftRows"&&r!=-1){
        for(var j = 0;j<n;j++){
            activeSheet.rows.splice(r+j,0,newRow);  //当前行前面插入一行  
        }
        activeSheet.spans.forEach(function(item){
            if(r<=item.row){    //合并行如果在r之后  则要增加合并的开始行
                item.row += n
            }else if(r>item.row&&r<item.row+item.rowCount){
                item.rowCount += n
            }
        });
        for(var i = rowkey.length-1; i>=0;i--){ //从后面改变指向  
            key = rowkey[i]    
            if(r<=parseInt(key)){
                 newKey = (parseInt(key)+ n).toString()   
                 textValue[newKey]=textValue[key] ;       
            }
            if(r<=(parseInt(key))&&(parseInt(key))<parseInt(newKey)){   //删除新行的数据
                delete textValue[key]
             }
         };
         this.row(r,index);
         this.col(0,index)
         child.innerText =''
         this.totolHeight = this.totolHeight + defaultHeight*n;
         this.recalculatePercentage(index)
    }else if(InsertType=="F1ShiftCols"&&c!=-1){
        for(var j = 0;j<n;j++){
            activeSheet.columns.splice(c+j,0,newCol);              //当前行前面插入一行  
        }
        activeSheet.spans.forEach(function(item){
            if(c<=item.col){
                item.col += n
            }else if(c>item.col&&c<(item.col+item.colCount)){
                item.colCount += n
            }
        });

        var colHeaderSpans = activeSheet.colHeaderData.spans;
        if(colHeaderSpans){
            colHeaderSpans.forEach(function(item){
                if(c<=item.col){
                    item.col += n
                }else if(c>item.col&&c<(item.col+item.colCount)){
                    item.colCount += n
                }
            })
        }
        for(var k = rowkey.length-1; k>=0;k--){
            var colkey = Object.keys(textValue[rowkey[k]]) 
            for(var j = colkey.length-1;j>=0;j--){     //colkey = [ 1   4 ]
                key = colkey[j]   //    4  1
                if(c<=parseInt(key)){  
                    newKey = (parseInt(key)+n).toString()    //5   2
                    textValue[rowkey[k]][newKey]= textValue[rowkey[k]][key]  
                }
                if(c<=(parseInt(key))&&(parseInt(key))<parseInt(newKey)){  
                    delete textValue[rowkey[k]][key] 
                }
            }   
        };
        this.row(0,index)
        this.col(c,index)
        child.innerText =''
        this.totolWidth = this.totolWidth + defaultWidth*n;
        this.recalculatePercentage(index)
    }
}

 /**
 * @api {null} WB.editDelete(nShiftType,R,C,N,Index) editDelete
 * @apiName editDelete
 * @apiGroup Function
 * @apiDescription Delete the current or specified row or column
 * @apiParam {String} nShiftType  
 * - F1ShiftRows:(Delete row)
 * - F1ShiftCols(Delete column)
 * @apiParam {Int} R  Which row to delete, R>=0, when the column is deleted, R passes -1
 * @apiParam {Int} C  Which column to delete, C>=0, when deleting rows, C passes -1
 * @apiParam {Int} [N=1]  N>=1Number of rows or columns
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.editDelete  ("F1ShiftRows",0,-1)   //Delete the first line
    WB.editDelete  ("F1ShiftCols",-1,1)   //Delete the second column
 */
Workbook.prototype.editDelete = function(nShiftType,R,C,N,Index){ 
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    textValue = activeSheet.data.dataTable,rowkey = Object.keys(textValue);  
    var newkey,key,colkey, r = 0,c = 0,n = 0;
    r = (R>=0)?R:activeSheet.rangeRow1;
    c = (C>=0)?C:activeSheet.rangeCol1;
    n = (N>=1)?N:1;
    var spans = activeSheet.spans,row = activeSheet.rows,col = activeSheet.columns;
    if(r+n>row.length) n = row.length - r;
    if(c+n>col.length) c = col.length - c;
    if(nShiftType == "F1ShiftRows"&&R!=-1&&row.length>=n){
        for(var j = n+r-1;j>=r;j--){
            if(row[j]){
                this.totolHeight = this.totolHeight-row[j].size;  //删除行和重新计算总高度
                activeSheet.rows.splice(j,1)
            }
        }
        for(var i = spans.length-1 ;i>=0;i--){   //当删除行   合并行参数做相应改变
            for(var t = n+r-1;t>=r;t--){
                if(spans[i].row<=t&&t<=(spans[i].rowCount+spans[i].row-1)){
                    activeSheet.spans[i].rowCount -= 1;
                }else if(spans[i].row>t){
                    activeSheet.spans[i].row -= 1;
                }; 
                if(spans[i].rowCount<1||spans[i].colCount<1){
                    activeSheet.spans.splice(i,1)
                } 
            }      
        };

        for(var q = 0; q<rowkey.length;q++){
            key = rowkey[q]    
            for(var w = r;w<n+r;w++){
                if(w==(parseInt(key))&&textValue[key]){
                   delete textValue[key]
                }
            };
            if(r+n-1<parseInt(key)){      
                newkey = (parseInt(key)-n).toString()   
                textValue[newkey]=textValue[key] ;  
                delete textValue[key]
            } 
        };
        var newc = (activeSheet.activeCol>=0)?activeSheet.activeCol:0;
        var newr = (activeSheet.activeRow>=0)?activeSheet.activeRow:0;
        this.col(newc,index)
        this.row(newr,index)
        this.tempValue = {}
        this.recalculatePercentage(index);
        if(activeSheet.rows<=0){
            this.removeTextArea()
        }
    }else if(nShiftType == "F1ShiftCols"&&C!=-1&&col.length>=n){
        for(var j = n+c-1;j>=c;j--){
            if(col[j]){
                this.totolWidth = this.totolWidth - col[j].size;
                activeSheet.columns.splice(j,1);
            }
        };

        for(var w = spans.length-1 ;w>=0;w--){   
            for(var t = n+c-1;t>=c;t--){
                if(spans[w].col<=t&&t<=(spans[w].colCount+spans[w].col-1)){
                    activeSheet.spans[w].colCount -= 1;
                }else if(spans[w].col>t){
                    activeSheet.spans[w].col -= 1;
                }; 
                if(spans[w].colCount<1||spans[w].rowCount<1){
                    activeSheet.spans.splice(w,1)
                }  
            }       
        };

        var colHeaderSpans = activeSheet.colHeaderData.spans;
        if(colHeaderSpans){
            for(var d = colHeaderSpans.length-1;d>=0;d--){
                for(var s = n+c-1;s>=c;s--){
                    if(colHeaderSpans[d].col<=s&&s<=(colHeaderSpans[d].colCount+colHeaderSpans[d].col-1)){
                        activeSheet.colHeaderData.spans[d].colCount -= 1;
                    }else if(colHeaderSpans[d].col>t){
                        activeSheet.colHeaderData.spans[d].col -= 1;
                    }; 
                    if(colHeaderSpans[d].colCount<=1){
                        activeSheet.colHeaderData.spans.splice(d,1)
                    }  
                }
            }
        }
    
        for(var e = 0; e<rowkey.length;e++){
            colkey = Object.keys(textValue[rowkey[e]]) 
           for(var j = 0;j<colkey.length;j++){     
                key = colkey[j]      
               if((parseInt(key))>=c&&(parseInt(key))<c+n&&parseInt(key)){
                   delete textValue[rowkey[e]][key]
               }
               if(c+n<=parseInt(key)){      
                   newkey = (parseInt(key)-n).toString()   
                   textValue[rowkey[e]][newkey]=textValue[rowkey[e]][key] ;  
                   delete textValue[rowkey[e]][key]
               }
           }     
       };
        var newC = (activeSheet.activeCol>=0)?activeSheet.activeCol:0;
        var newR = (activeSheet.activeRow>=0)?activeSheet.activeRow:0;
        this.col(newC,index)
        this.row(newR,index)
        this.tempValue = {}
        this.recalculatePercentage(index)
        if(activeSheet.columns<=0){
            this.removeTextArea()
        }
    }
}	

//复制
Workbook.prototype.editCopy = function(R1,C1,R2,C2,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,spans = activeSheet.spans,copyData,r1 =0 ,c1=0,r2=0,c2=0,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,
    mode = this.workbook.behaviorMode;
    this.clipBoard = [];    //剪贴板  保存复制剪切的数据
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]){
                copyData =  data[i][j]
            }else{
                copyData = {}
            };
            this.clipBoard.push(copyData)
        }    
    };
    for(var k = 0;k<spans.length;k++){       
        if(spans[k].row>=r1&&spans[k].row<=r2&&spans[k].col>=c1&&spans[k].col<=c2){
            this.clipBoard.push(spans[k])
        }
    };
    var isFocus = (this.isEdit&&activeSheet.rangeRow1!=-1&&activeSheet.rangeCol1!=-1||mode==2||mode==3)?true:false;
    this.clipBoard.push({"rangeRow":r2-r1,"rangeCol":c2-c1,"c1":c1,"r1":r1,"c2":c2,"r2":r2,"operation":"copy","isFocus":isFocus})
    //rangeRow，rangeCol在剪贴板最后插入行列的范围用于填充粘贴
}

//剪切
Workbook.prototype.editCut = function(R1,C1,R2,C2,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,spans = activeSheet.spans,cutData,r1=0,c1=0,r2=0,c2=0,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,
    mode = this.workbook.behaviorMode;
    this.clipBoard = [];
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]){
                cutData =  data[i][j]
            }else{
                cutData = {}
            };
            this.clipBoard.push(cutData)
        }
    };

    for(var k = spans.length-1;k>=0;k--){      
        if(spans[k].row>=r1&&spans[k].row<=r2&&spans[k].col>=c1&&spans[k].col<=c2){
            this.clipBoard.push(spans[k]);
        }
    };
    var isFocus = (this.isEdit&&activeSheet.rangeRow1!=-1&&activeSheet.rangeCol!=-1||mode==2||mode==3)?true:false
    this.clipBoard.push({"rangeRow":r2-r1,"rangeCol":c2-c1,"c1":c1,"r1":r1,"c2":c2,"r2":r2,"operation":"cut","isFocus":isFocus})                //在剪贴板最后插入行列的范围用于填充粘贴
}

 //粘贴 
Workbook.prototype.editPaste = function(R,C,Index){      
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,r1=0,c1=0,rangeRow=0,rangeCol=0,k = 0,newSpansArray = [],
    spans = activeSheet.spans,child = document.querySelector("#"+this.boxId+" .child");
    r1 = (R>=0)?R:activeSheet.rangeRow1;
    c1 = (C>=0)?C: activeSheet.rangeCol1;
    if(c1<0)    c1 = 0;
    if(r1<0)    r1 = 0;
    var  isLock = this.getLock(r1,c1,r1,c1);
    if(isLock)  return;
    if(this.clipBoard&&this.clipBoard.length>=1){
        var isFocus = this.clipBoard[this.clipBoard.length-1].isFocus;
        if(!isFocus){                  //复制剪切操作是以整个单元格为单位的  粘贴的时候连带样式等等都要粘贴过来；
            rangeRow = this.clipBoard[this.clipBoard.length-1].rangeRow;
            rangeCol = this.clipBoard[this.clipBoard.length-1].rangeCol;
            for(var i = spans.length-1;i>=0;i--){        //粘贴前 判断是否有合并       先删除原有的合并
                if((r1+rangeRow)>=spans[i].row&&r1<(spans[i].row+spans[i].rowCount)&&(c1+rangeCol)>=spans[i].col&&c1<(spans[i].col+spans[i].colCount)&&(rangeRow!=0||rangeCol!=0)){
                    var spansIsLock = this.getLock(spans[i].row,spans[i].col)
                    if(!spansIsLock){
                        activeSheet.spans.splice(i,1)
                    }
                };
            }
            //取出复制或者剪切过来的合并项   -1-1    最后一项用于记录范围的数据不要
            var dataSum = (rangeRow+1)*(rangeCol+1)
            for(var j = this.clipBoard.length-1-1;j>dataSum-1;j--){
                newSpansArray.push(JSON.parse(JSON.stringify(this.clipBoard[j])));       //拷贝
            };
    
            var clipBoardR1 = this.clipBoard[this.clipBoard.length-1].r1;
            var clipBoardC1 = this.clipBoard[this.clipBoard.length-1].c1;
            var clipBoardR2 = this.clipBoard[this.clipBoard.length-1].r2;
            var clipBoardC2 = this.clipBoard[this.clipBoard.length-1].c2;
            var operation = this.clipBoard[this.clipBoard.length-1].operation;
            for(var c = spans.length-1;c>=0;c--){       //粘贴成功后删除原来剪切的合并
                if(spans[c].row>=clipBoardR1&&spans[c].row<=clipBoardR2&&spans[c].col>=clipBoardC1&&spans[c].col<=clipBoardC2&&operation=='cut'){
                    activeSheet.spans.splice(c,1)
                }
            };

            for(var q = clipBoardR1;q<=clipBoardR2;q++){                   //粘贴成功删除原来的剪切的数据项
                for(var e = clipBoardC1;e<=clipBoardC2;e++){
                    if(data[q]&&data[q][e]&&operation=='cut'){
                        if(!(data[q][e].style&&data[q][e].style.lock)){
                            delete data[q][e]
                        }
                    }
                }
                if(data[q]){               //如果整行剪切，该行数据被迁移   删除空行数据
                    if(Object.keys(data[q]).length ==0&&operation=='cut'){
                        delete data[q]
                    }
                }
            };
    
             //然后把复制剪切过来的合并项 push到  sheet的  spans中； 须在改变row col   
            for(var t = 0;t<newSpansArray.length;t++){ 
                newSpansArray[t].row = r1 +  newSpansArray[t].row -  clipBoardR1;
                newSpansArray[t].col = c1 +  newSpansArray[t].col -  clipBoardC1; 
                for(var i = newSpansArray[t].row+newSpansArray[t].rowCount-1;i>=newSpansArray[t].row;i--){
                    for(var j = newSpansArray[t].col+ newSpansArray[t].colCount;j>= newSpansArray[t].col;j--){
                        if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.lock){
                            newSpansArray[t].rowCount = (i-newSpansArray[t].col>=1)?i-newSpansArray[t].row:1;
                            newSpansArray[t].colCount = (j-newSpansArray[t].col>=1)?j-newSpansArray[t].col:1;
                        }
                    }
                };
                if(!(newSpansArray[t].colCount<=1&&newSpansArray[t].rowCount<=1)){
                    activeSheet.spans.push(newSpansArray[t])
                }
            } ;
    
            var rLen = 0,cLen = 0;
            ((r1+rangeRow)>activeSheet.rows.length-1)?rLen=activeSheet.rows.length-1:rLen=(r1+rangeRow);
            ((c1+rangeCol)>activeSheet.columns.length-1)?cLen=activeSheet.columns.length-1:cLen=(c1+rangeCol);
            for(var y = r1;y<=rLen;y++){            //rangeRow   rangeCol   复制剪切文本记录的范围
                 for(var u = c1;u<=cLen;u++){
                     if(data[y]&&data[y][u]&&data[y][u].style&&data[y][u].style.lock){
                         k+=1
                         continue       //锁定跳过
                     }
                    var rowObj ={}; 
                    rowObj[u]={} //没有记录行数据情况  用于新增   
                     if(this.clipBoard&&data[y]){
                         data[y][u] = JSON.parse(JSON.stringify(this.clipBoard[k]))
                     }else if(this.clipBoard&&!data[y]){
                         data[y] = rowObj;     //
                         data[y][u] = JSON.parse(JSON.stringify(this.clipBoard[k]))
                     };
                     k+=1;
                     if(JSON.stringify(data[y][u])==="{}"){
                         delete data[y][u]
                     }
                 } ;
                 if(data[y]){              
                    if(Object.keys(data[y]).length ==0){
                        delete data[y]
                    }
                }
            };
        }else{           //当剪切或者复制的只是单元格的部分文本的请款下  值粘贴该文本  样式等等不进行粘贴 
            if(!this.isEdit){
                this.setValue(r1,c1,this.clipBoardTxt)
            }
        }
    }
}

 /**
 * @api {null} WB.clearText(R1,C1,R2,C2,Index) clearText
 * @apiName clearText
 * @apiGroup Function
 * @apiDescription Clear the text of a range of cells(if there is text and value; text and value is set to empty)
 * - Cells that are locked/uneditable/checkbox cannot be deleted (no parameter defaults to clear the current range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.clearText (1,0)
 */
Workbook.prototype.clearText = function(R1,C1,R2,C2,Index){                     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,
    data = activeSheet.data.dataTable;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var cellStatus = this.getCellStatus(i,j),isCheckBox = cellStatus.isCheckBox,
            isCanEdit = cellStatus.isCanEdit,isLock = cellStatus.isLock;
            if(!isCheckBox&&!isLock&&isCanEdit!==false&&data[i]&&data[i][j]){
                var cell = data[i][j];
                if(cell.hasOwnProperty('value')){
                    cell.value = ''
                };
                if(cell.hasOwnProperty('text')){
                    cell.text = ''
                };
            }
        }
    }
};

 /**
 * @api {null} WB.deleteText(R1,C1,R2,C2,Index) deleteText
 * @apiName deleteText
 * @apiGroup Function
 * @apiDescription Delete the contents of a range of cells (if there are text and value attributes; these two attributes will be deleted, which is different from clearText)
 * - Cells that are locked/uneditable/checkbox cannot be deleted; (the default range is deleted without parameters)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.deleteText (1,0) 
 */
Workbook.prototype.deleteText = function(R1,C1,R2,C2,Index){                     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,
    data = activeSheet.data.dataTable;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var cellStatus = this.getCellStatus(i,j),isCheckBox = cellStatus.isCheckBox,
            isCanEdit = cellStatus.isCanEdit,isLock = cellStatus.isLock;
            if(!isCheckBox&&!isLock&&isCanEdit!==false&&data[i]&&data[i][j]){
                var cell = data[i][j];
                if(cell.hasOwnProperty('value')){
                    delete cell.value
                };
                if(cell.hasOwnProperty('text')){
                    delete cell.text
                };
            }
        }
    }
};

 /**
 * @api {null}  WB.deleteCell(R1,C1,R2,C2,Index) deleteCell
 * @apiName deleteCell
 * @apiGroup Function
 * @apiDescription Delete cell (including text and style, no parameter, delete current range by default)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.deleteCell (1,0)
 */
Workbook.prototype.deleteCell = function(R1,C1,R2,C2,Index){           
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var r1 = 0,c1 = 0,r2 = 0,c2 = 0;
    var rcArg = this.handleArg(R1,C1,R2,C2,index),
    r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2
    var textValue = activeSheet.data.dataTable;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(textValue[i]&&textValue[i][j]&&textValue[i][j]){
                delete textValue[i][j]
            }
        };
        if(textValue[i]){               //该行数据被迁移   删除空行数据
            if(Object.keys(textValue[i]).length ==0){
                delete textValue[i]
            }
        }
    }
};

/**
 * @api {null} WB.setRowHeight(height,R1,R2,Index) setRowHeight
 * @apiName setRowHeight
 * @apiGroup Function
 * @apiDescription Set the height of the lines within the range (hidden lines cannot be set)
 * @apiParam {Numer} height  Value (height<=0, hidden)
 * @apiParam {Int} R1  Start row(R1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=r1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
     WB.setRowHeight(80,0,1)  //The first and second lines set the height of the line to 80
 */
Workbook.prototype.setRowHeight = function(height,R1,R2,Index){ 
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    if(R2===undefined){
        R2 = R1
    };
    if(R1!=-1){
        for(var i = R1;i<=R2;i++){
            if(activeSheet.rows[i]&&!activeSheet.rows[i].bHidden){
                this.totolHeight -=  activeSheet.rows[i].size;
                if(height<=0){
                    activeSheet.rows[i].size = 0;
                    activeSheet.rows[i].bHidden = true
                }else{
                    activeSheet.rows[i].size = height;
                };
                this.totolHeight +=  activeSheet.rows[i].size
                this.setNoDefaultH(i,index)
            }
       }
    }
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.setColWidth(width,c1,c2,index) setColWidth
 * @apiName setColWidth
 * @apiGroup Function
 * @apiDescription Set the width of the columns in the range (hidden columns cannot be set)
 * @apiParam {Numer} width  Value (width<=0, hidden)
 * @apiParam {Int} c1  Start column(c1>=0)
 * @apiParam {Int} [c2]  End column(c1>=0,c2>=c1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setColWidth(80,0,0)  //Set the width of the first column to 80
 */
Workbook.prototype.setColWidth = function(width,C1,C2,Index){    //输入设置列宽度
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    if(C2==undefined){
        C2 = C1
    };
    if(C1!=-1){
        for(var i = C1;i<=C2;i++){
            if(activeSheet.columns[i]&&!activeSheet.columns[i].bHidden){
                this.totolWidth -=  activeSheet.columns[i].size;   //先把原来的减掉
                if(width<=0){
                    activeSheet.columns[i].size=0;
                    activeSheet.columns[i].bHidden = true;
                }else{
                    activeSheet.columns[i].size = width 
                }
            };
            this.totolWidth +=  activeSheet.columns[i].size
            this.setNoDefaultW(i,index)
       }
    };
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.setHAlignment(R1,C1,R2,C2,HAlign,Index) setHAlignment
 * @apiName setHAlignment
 * @apiGroup Function
 * @apiDescription Set the horizontal alignment of cell text in the range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [hAlign=1]  
 * - 1: Regular (aligned according to cell data type)
 * - 2：left
 * - 3：center
 * - 4：Right
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setHAlignment(0,0,0,0,2) //Set the cell text in the range to be left aligned
 */
Workbook.prototype.setHAlignment = function(R1,C1,R2,C2,HAlign,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    hAlign = HAlign || 1 ,data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.hAlign = hAlign
            }
        }   
    }
}

/**
 * @api {null} WB.setVAlignment(R1,C1,R2,C2,VAlign,Index) setVAlignment
 * @apiName setVAlignment
 * @apiGroup Function
 * @apiDescription Set the vertical alignment of cell text in a range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [vAlign=2]
 * - 1：top
 * - 2：middle
 * - 3：bottom
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setVAlignment(0,0,0,0,1)
 */
Workbook.prototype.setVAlignment = function(R1,C1,R2,C2,VAlign,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    vAlign = VAlign || 2,data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.vAlign = vAlign
            }
        }   
    }
}

/**
 * @api {null} WB.getAlignment(R1,C1,R2,C2,Index) getAlignment
 * @apiName getAlignment
 * @apiGroup Function
 * @apiDescription Get the alignment of the cells in the range (no parameter defaults to get the alignment of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getAlignment(0,0) 
 */
Workbook.prototype.getAlignment = function(R1,C1,R2,C2,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,alignArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),
    r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var alignObj = {}
            if(data[i]&&data[i][j]&&data[i][j].style){
                alignObj.r = i;
                alignObj.c = j;
                if(data[i][j].style.hAlign){
                    alignObj.hAlign =  data[i][j].style.hAlign;
                };
                if(data[i][j].style.vAlign){
                    alignObj.vAlign = data[i][j].style.vAlign ;
                }
                alignArr.push(alignObj)
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&alignArr.length==1){
        return alignArr[0]
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return alignArr
    }
}

/**
 * @api {null} WB.fontColor(R1,C1,R2,C2,Color,Index) fontColor
 * @apiName fontColor
 * @apiGroup Function
 * @apiDescription Set the font color of the cell text in the range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {String} [color=#000] 
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fontColor(0,0,2,2,'pink')   //Set the font to pink
 */
Workbook.prototype.fontColor = function(R1,C1,R2,C2,Color,Index){     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    color = Color || "black",data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.fontColor = color;
            }
        }   
    }
};

/**
 * @api {null} WB.getFontColor(R1,C1,R2,C2,Index) getFontColor
 * @apiName getFontColor
 * @apiGroup Function
 * @apiDescription Get the font color of the cells in the range (no parameter defaults to get the font color of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getFontColor(0,0) 
 */
Workbook.prototype.getFontColor = function(R1,C1,R2,C2,Index){      
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,fontColorArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var fontcolor;
            if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.fontColor){
                fontcolor = data[i][j].style.fontColor;
                fontColorArr.push({"r":i,"c":j,"fontColor":fontcolor})
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&fontColorArr.length==1){
        return fontColorArr[0].fontColor
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return fontColorArr
    }
}

/**
 * @api {null} WB.setFontItalic(R1,C1,R2,C2,BItalic,Index) setFontItalic
 * @apiName setFontItalic
 * @apiGroup Function
 * @apiDescription Whether the font of the cells in the set range is italic (no parameter defaults to set the italic of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean}  [BItalic=true] 
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFontItalic(1,1,1,1,true)
 */
Workbook.prototype.setFontItalic = function(R1,C1,R2,C2,BItalic,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,b = (BItalic!==undefined)?BItalic:true,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.font){
                    var fontList = this.getFontList(i,j,index)
                    var originName = fontList.originName,originSize = fontList.originSize,originHeight = fontList.originHeight,
                    originBold = fontList.originBold;
                    if(originBold&&b){
                        data[i][j].style.font = 'italic'+' '+'bold'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(originBold&&!b){
                        data[i][j].style.font = 'bold'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(!originBold&&b){
                        data[i][j].style.font = 'italic'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(!originBold&&!b){
                        data[i][j].style.font = originSize+"px/"+originHeight+'px '+originName;
                    }
                }
            }
        }
    } ; 
}

/**
 * @api {null} WB.setFontBold(R1,C1,R2,C2,BBold,Index) setFontBold
 * @apiName setFontBold
 * @apiGroup Function
 * @apiDescription Whether the font of the cells in the setting range is bold (no parameter defaults to setting the bold of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean}  [BBold=true]  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFontBold(1,1,1,1,true)  
 */
Workbook.prototype.setFontBold = function(R1,C1,R2,C2,BBold,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,b = (BBold!==undefined)?BBold:true,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.font){
                    var fontList = this.getFontList(i,j,index)
                    var originName = fontList.originName,originSize = fontList.originSize,originHeight = fontList.originHeight,
                    originItalic = fontList.originItalic;
                    if(b&&originItalic){
                        data[i][j].style.font = 'italic'+' '+'bold'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(b&&!originItalic){
                        data[i][j].style.font = 'bold'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(!b&&originItalic){
                        data[i][j].style.font = 'italic'+' '+originSize+"px/"+originHeight+'px '+originName;
                    }else if(!b&&!originItalic){
                        data[i][j].style.font = originSize+"px/"+originHeight+'px '+originName;
                    }
                }
            }
        }
    } 
}

/**
 * @api {null} WB.setFontLineH(R1,C1,R2,C2,NLineHeight,Index) setFontLineH
 * @apiName setFontLineH
 * @apiGroup Function
 * @apiDescription Set the font line height of the cells in the range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int}  NLineHeight  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFontLineH(1,1,1,1,20)  //Set the font line height of the cells in the range to 20px
 */
Workbook.prototype.setFontLineH = function(R1,C1,R2,C2,NLineHeight,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,rows = activeSheet.rows, paddingBottom = parseInt(activeSheet.cellPadding.bottom),
    paddingTop = parseInt(activeSheet.cellPadding.top),defaultH = activeSheet.defaults.rowHeight;
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.font){
                    var fontList = this.getFontList(i,j,index),
                    originName = fontList.originName,originSize = fontList.originSize,originHeight = fontList.originHeight,
                    originBold = fontList.originBold,originItalic = fontList.originItalic,nLineHeight = NLineHeight || originHeight;
                    if(originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+'bold'+' '+originSize+"px/"+nLineHeight+'px '+originName;
                    }else if(originBold&&!originItalic){
                        data[i][j].style.font = 'bold'+' '+originSize+"px/"+nLineHeight+'px '+originName;
                    }else if(!originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+originSize+"px/"+nLineHeight+'px '+originName;
                    }else if(!originBold&&!originItalic){
                        data[i][j].style.font = originSize+"px/"+nLineHeight+'px '+originName;
                    }
                    if(rows[i]&&!rows[i].noDefaultH){
                        var wordWrap = this.getWordWrap(i,j,i,j,index);
                        if(wordWrap){
                            this.incrementRowheight(i,j,index)
                        }else{
                            var h = (nLineHeight>defaultH)?nLineHeight+paddingBottom+paddingTop:defaultH;
                            this.setRowHeight(h,i,i,index)
                        }
                    };
                }
            }
        }
    }  
}

/**
 * @api {null} WB.setFontSize(R1,C1,R2,C2,NSize,Index)   setFontSize
 * @apiName setFontSize
 * @apiGroup Function
 * @apiDescription Set the font size of the cells in the range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int}  NSize  font size
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFontSize(1,1,1,1,20)  //Set the cell font size in the range to 20px
 */
Workbook.prototype.setFontSize = function(R1,C1,R2,C2,NSize,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rows = activeSheet.rows,data = activeSheet.data.dataTable,nSize = parseInt(NSize),defaultH = activeSheet.defaults.rowHeight,
    paddingBottom = parseInt(activeSheet.cellPadding.bottom),paddingTop = parseInt(activeSheet.cellPadding.top);
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.font){
                    var fontList = this.getFontList(i,j,index)
                    var originName = fontList.originName,originSize = fontList.originSize,originHeight = fontList.originHeight,
                    originBold = fontList.originBold,originItalic = fontList.originItalic,nSize = NSize || originSize;
                    if(rows[i]&&!rows[i].noDefaultH){
                        originHeight += nSize-originSize;
                        // if(originHeight<defaultH){
                        //     originHeight = defaultH
                        // }
                    };
                    if(originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+'bold'+' '+nSize+"px/"+originHeight+'px '+originName;
                    }else if(originBold&&!originItalic){
                        data[i][j].style.font = 'bold'+' '+nSize+"px/"+originHeight+'px '+originName;
                    }else if(!originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+nSize+"px/"+originHeight+'px '+originName;
                    }else if(!originBold&&!originItalic){
                        data[i][j].style.font = nSize+"px/"+originHeight+'px '+originName;
                    };
                    if(rows[i]&&!rows[i].noDefaultH){
                        var wordWrap = this.getWordWrap(i,j,i,j,index);
                        if(wordWrap){
                            this.incrementRowheight(i,j,index)
                        }else{
                            var h = (originHeight>defaultH)?originHeight+paddingBottom+paddingTop:defaultH;
                            this.setRowHeight(h,i,i,index)
                        }
                    };
                }
            }
        }
    }     
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.setFontName(R1,C1,R2,C2,PName,Index) setFontName
 * @apiName setFontName
 * @apiGroup Function
 * @apiDescription Set the font type of the cells in the range
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {String}  PName  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFontName(0,0,2,2,'宋体')  
 */
Workbook.prototype.setFontName = function(R1,C1,R2,C2,PName,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){
        for(var j = C1;j<=C2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.font){
                    var fontList = this.getFontList(i,j,index)
                    var originName = fontList.originName,originSize = fontList.originSize,originHeight = fontList.originHeight,
                    originBold = fontList.originBold,originItalic = fontList.originItalic,pName = PName || originName;
                    if(originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+'bold'+' '+originSize+"px/"+originHeight+'px '+pName;
                    }else if(originBold&&!originItalic){
                        data[i][j].style.font = 'bold'+' '+originSize+"px/"+originHeight+'px '+pName;
                    }else if(!originBold&&originItalic){
                        data[i][j].style.font = 'italic'+' '+originSize+"px/"+originHeight+'px '+pName;
                    }else if(!originBold&&!originItalic){
                        data[i][j].style.font = originSize+"px/"+originHeight+'px '+pName;
                    }
                }
            }
        }
    }     
}

/**
 * @api {null} WB.getFontList(r,c,Index) getFontList
 * @apiName getFontList
 * @apiGroup Function
 * @apiDescription Get every item of the font
 * @apiParam {Int} r  row(r>=0)
 * @apiParam {Int} c  column(c>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.getFontList = function(i,j,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,originItalic,originBold,originSize,originHeight;
    if(data[i]&&data[i][j]&&data[i][j].style){
        if(data[i][j].style.font){
            try{
                var fontList = data[i][j].style.font.trim().split(/\s+/)
                var originName = fontList[fontList.length-1]||this.workbook.defaultFontName
                var sizeAndHeight = fontList[fontList.length-2].split('/')
                originSize = parseFloat(sizeAndHeight[0])||this.workbook.defaultFontSize
                originHeight = parseFloat(sizeAndHeight[1])||this.workbook.defaultLineHeight
                if(fontList[fontList.length-3]=='italic'|| fontList[fontList.length-4]=='italic'){
                    originItalic = 'italic';
                };
                if(fontList[fontList.length-3]=='bold'){
                    originBold = 'bold'
                }
            }catch(e){}
        }
    }
    return {"originName":originName,"originSize":originSize,"originHeight":originHeight,"originBold":originBold,"originItalic":originItalic}
}

/**
 * @api {null} WB.getFont(R1,C1,R2,C2,Index) getFont
 * @apiName getFont
 * @apiGroup Function
 * @apiDescription Get the font of the selected cell (no parameter defaults to get the font of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getFont(0,0) 
 */
Workbook.prototype.getFont = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,fontArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var font;
            if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.font){
                font = data[i][j].style.font;
                fontArr.push({"r":i,"c":j,"font":font})
            }
        }
    };
    if(((r1==r2&&c1==c2)||isSpan)&&fontArr.length==1){
        return fontArr[0].font
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return fontArr
    }
}

/**
 * @api {null} WB.clearAllBorder(R1,C1,R2,C2,Index) clearAllBorder
 * @apiName clearAllBorder
 * @apiGroup Function
 * @apiDescription Clear all the borders of the cells in the range (without parameters, the current range is cleared by default)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.clearAllBorder(2,2,5,3) 
 */
Workbook.prototype.clearAllBorder = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,rcArg = this.handleArg(R1,C1,R2,C2,index),
    r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.borderLeft){
                    delete data[i][j].style.borderLeft
                };
                if(data[i][j].style.borderRight){
                    delete data[i][j].style.borderRight
                };
                if(data[i][j].style.borderBottom){
                    delete  data[i][j].style.borderBottom
                };
                if(data[i][j].style.borderTop){
                    delete data[i][j].style.borderTop
                }
            };
        }
    }   
}

/**
 * @api {null} WB.setOutBorder(R1,C1,R2,C2,NOut,CrOut,Index) setOutBorder
 * @apiName setOutBorder
 * @apiGroup Function
 * @apiDescription Set the outer border of the cells in the range (no parameter defaults to the outer border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NOut=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrOut='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setOutBorder(1,1,5,5)  
 */
Workbook.prototype.setOutBorder = function(R1,C1,R2,C2,NOut,CrOut,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crOut = CrOut || "black",nOut = NOut || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2; 
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                var merge = this.getMergeCount(i,j,index),rowCount = merge.rowCount-1,colCount = merge.colCount-1
                if(j==c1){
                    data[i][j].style.borderLeft = { "color":crOut,"style":nOut}; 
                };
                if(j+colCount==c2){
                    data[i][j].style.borderRight = { "color":crOut,"style":nOut}; 
                };
                if(i==r1){
                    data[i][j].style.borderTop = { "color":crOut,"style":nOut}; 
                };
                if(i+rowCount==r2){
                    data[i][j].style.borderBottom = { "color":crOut,"style":nOut}; 
                };
            };
        }
    }
}


/**
 * @api {null} WB.setINnerBorder(R1,C1,R2,C2,NInner,CrInner,Index) setINnerBorder
 * @apiName setINnerBorder
 * @apiGroup Function
 * @apiDescription Set the inner border of the cell in the range (no parameter defaults to the inner border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NInner=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrInner='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setINnerBorder(1,1,5,5)
 */
Workbook.prototype.setINnerBorder = function(R1,C1,R2,C2,NInner,CrInner,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crInner = CrInner || "black",nInner = NInner || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2; 
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                var merge = this.getMergeCount(i,j,index),rowCount = merge.rowCount-1,colCount = merge.colCount-1
                if(j!=c1){
                    data[i][j].style.borderLeft = { "color":crInner,"style":nInner}; 
                };
                if(j+colCount!=c2){
                    data[i][j].style.borderRight = { "color":crInner,"style":nInner}; 
                };
                if(i!=r1){
                    data[i][j].style.borderTop = { "color":crInner,"style":nInner}; 
                };
                if(i+rowCount!=r2){
                    data[i][j].style.borderBottom = { "color":crInner,"style":nInner}; 
                };
            };
        }
    }
}


/**
 * @api {null} WB.setAllBorder(R1,C1,R2,C2,NAll,CrALL,Index) setAllBorder
 * @apiName setAllBorder
 * @apiGroup Function
 * @apiDescription Set the full border of the cells in the range (no parameter defaults to the full border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NAll=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrALL='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setAllBorder(2,2,5,3) 
 */
Workbook.prototype.setAllBorder = function(R1,C1,R2,C2,NAll,CrALL,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,craLL = CrALL || "black",nAll = NAll || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.borderLeft = { "color":craLL,"style":nAll}; 
                data[i][j].style.borderRight = { "color":craLL,"style":nAll}; 
                data[i][j].style.borderBottom = { "color":craLL,"style":nAll}; 
                data[i][j].style.borderTop = { "color":craLL,"style":nAll}; 
            };        
        }
    } 
}

/**
 * @api {null} WB.setBottomBorder(R1,C1,R2,C2,NBottom,CrBottom,Index) setBottomBorder
 * @apiName setBottomBorder
 * @apiGroup Function
 * @apiDescription Set the bottom border of the cells in the range (no parameters default to the bottom border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NBottom=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrBottom='#000']    Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setBottomBorder(1,1,3,3)  
 */
Workbook.prototype.setBottomBorder = function(R1,C1,R2,C2,NBottom,CrBottom,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crBottom = CrBottom || "black",nBottom = NBottom || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index);
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.borderBottom = { "color":crBottom,"style":nBottom}; 
            };
        }
    }
}

/**
 * @api {null}  WB.setTopBorder(R1,C1,R2,C2,NTop,CrTop,Index) setTopBorder
 * @apiName setTopBorder
 * @apiGroup Function
 * @apiDescription Set the top border of the cells in the range (no parameter defaults to the top border of the currently selected range) 
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NTop=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrTop='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setTopBorder(1,1,3,3)
 */
Workbook.prototype.setTopBorder = function(R1,C1,R2,C2,NTop,CrTop,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crTop = CrTop || "black",nTop = NTop || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index);
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.borderTop = { "color":crTop,"style":nTop}; 
            };
        }
    }
}

/**
 * @api {null} WB.setRightBorder(R1,C1,R2,C2,NRight,CrRight,Index) setRightBorder
 * @apiName setRightBorder
 * @apiGroup Function
 * @apiDescription Set the right border of the cells in the range (no parameter defaults to the right border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [NRight=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [CrRight='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setRightBorder(1,1,3,3)
 */
Workbook.prototype.setRightBorder = function(R1,C1,R2,C2,NRight,CrRight,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crRight = CrRight || "black",nRight = NRight || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index);
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.borderRight = { "color":crRight,"style":nRight}; 
            };
        }
    }
}

/**
 * @api {null} WB.setLeftBorder(R1,C1,R2,C2,NLeft,CrLeft,Index) setLeftBorder
 * @apiName setLeftBorder
 * @apiGroup Function
 * @apiDescription Set the left border of the cells in the range (no parameter defaults to the left border of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [nLeft=1]  
 * - 1:	Thin Line
 * - 2: Medium Line
 * - 5: Thick Line
 * @apiParam {String} [crLeft='#000']   Border color
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setLeftBorder(1,1,3,3) 
 */
Workbook.prototype.setLeftBorder = function(R1,C1,R2,C2,NLeft,CrLeft,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,crLeft = CrLeft || "black",nLeft = NLeft || 1,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index);
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.borderLeft = { "color":crLeft,"style":nLeft}; 
            };
        }
    }
}

/**
 * @api {null} WB.getBorder(R1,C1,R2,C2,Index) getBorder
 * @apiName getBorder
 * @apiGroup Function
 * @apiDescription Get the border of the cell in the range (without parameters, the border of the currently selected range is obtained by default)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getBorder(0,0)
 */
Workbook.prototype.getBorder = function(R1,C1,R2,C2,Index){   //获取单元格边框
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,borderArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var borderObj = {}
            if(data[i]&&data[i][j]&&data[i][j].style){
                borderObj.r = i;
                borderObj.c = j;
                if(data[i][j].style.borderLeft){
                    borderObj.borderLeft =  data[i][j].style.borderLeft;
                };
                if(data[i][j].style.borderRight){
                    borderObj.borderRight = data[i][j].style.borderRight;
                };
                if(data[i][j].style.borderTop){
                    borderObj.borderTop = data[i][j].style.borderTop;
                };
                if(data[i][j].style.borderBottom){
                    borderObj.borderBottom = data[i][j].style.borderBottom;
                };
                if(borderObj.borderLeft||borderObj.borderRight||borderObj.borderTop|| borderObj.borderBottom){
                    borderArr.push(borderObj)
                }
            };
        }
    }    
    if(((r1==r2&&c1==c2)||isSpan)&&borderArr.length==1){
        return borderArr[0]
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return borderArr
    }
}

/**
 * @api {null} WB.col(column,index) col
 * @apiName col
 * @apiGroup Function
 * @apiDescription Set or return the active column of the current worksheet
 * @apiParam {Int} [column]  column(column>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.col()   //Get current active column
    WB.col(2)  //Set the current active column to the third column
 */
Workbook.prototype.col = function(column,Index){  
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),c = parseInt(column);
    if(c>=-1){
        activeSheet.activeCol = c;
        activeSheet.rangeCol1 = c;
        activeSheet.rangeCol2 = activeSheet.rangeCol1;
        activeSheet.spans.forEach(function(item){                      
            if(item.row<= activeSheet.rangeRow1 && activeSheet.rangeRow1 <(item.row+item.rowCount)&&activeSheet.rangeCol1>=item.col&&activeSheet.rangeCol1<(item.colCount+item.col)){
                activeSheet.rangeCol1 = item.col;
                activeSheet.activeCol = item.col;
                activeSheet.rangeCol2 = item.col+item.colCount-1;
                activeSheet.rangeRow1 = item.row;
                activeSheet.activeRow = item.row;
                activeSheet.rangeRow2 = item.row+item.rowCount-1;
            }
        });
        if(activeSheet.rangeCol1==-1){
            activeSheet.rangeCol2 = activeSheet.columns.length-1
        }
    }
    return activeSheet.activeCol;
}

/**
 * @api {null} WB.row(row,index) row
 * @apiName row
 * @apiGroup Function
 * @apiDescription Set or return the active row of the current worksheet
 * @apiParam {Int} [row]  row(row>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.row()   //Get current active line
    WB.row(2)  //Set the current active column to the third row
 */
Workbook.prototype.row = function(row,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),r = parseInt(row);
    if(r>=-1){
        activeSheet.activeRow = r;
        activeSheet.rangeRow1 = r;
        activeSheet.rangeRow2 = activeSheet.rangeRow1;
        activeSheet.spans.forEach(function(item){                     
            if(item.row<= activeSheet.rangeRow1 && activeSheet.rangeRow1 <(item.row+item.rowCount)&&activeSheet.rangeCol1>=item.col&&activeSheet.rangeCol1<(item.colCount+item.col)){
                activeSheet.rangeRow1 = item.row;
                activeSheet.activeRow = item.row;
                activeSheet.rangeRow2 = item.row+item.rowCount-1;
                activeSheet.rangeCol2 = item.col+item.colCount-1;
                activeSheet.activeCol = item.col;
                activeSheet.rangeCol1 = item.col;
            }
        });
        if(activeSheet.rangeRow1==-1){
            activeSheet.rangeRow2 = activeSheet.rows.length-1
        }
    }   
    return activeSheet.activeRow;
}

/**
 * @api {null} WB.setActiveEditCell(R,C) setActiveEditCell
 * @apiName setActiveEditCell
 * @apiGroup Function
 * @apiDescription Open the specified edit cell
 * @apiParam {Int} R  row(R>=0)
 * @apiParam {Int} C  column(C>=0)
 * @apiExample {javascript} demo:
    WB.setActiveEditCell(1,1)     
 */
Workbook.prototype.setActiveEditCell = function(R,C){
    var r = parseInt(R),c = parseInt(C);
    if(r>=0&&c>=0){
        this.row(r);
        this.col(c);
        this.tempValue = {}
        this.isEdit=true;this.incanvasPath=true,this.eButton=0
        this.startPaint()
    }
}

/**
 * @api {null} WB.fontStrikeout(R1,C1,R2,C2,bool,Index) fontStrikeout
 * @apiName fontStrikeout
 * @apiGroup Function
 * @apiDescription Strikethrough of the cells in the set range (no parameter defaults to the strikethrough of the currently selected cell)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} [boolean=true]  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fontStrikeout(0,0,2,2,true)
 */
Workbook.prototype.fontStrikeout = function(R1,C1,R2,C2,bool,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable, b = (bool!==undefined)?bool:true,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.fontStrikeout = b;
            }
        }   
    }
}

/**
 * @api {null} WB.getFontStrikeout(R1,C1,R2,C2,Index) getFontStrikeout
 * @apiName getFontStrikeout
 * @apiGroup Function
 * @apiDescription Whether to get the strikethrough of the cells in the range (no parameter defaults to get the strikethrough status of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getFontStrikeout(0,0)
 */
Workbook.prototype.getFontStrikeout = function(R1,C1,R2,C2,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,fontStrikeoutArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var fontStrikeout;
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.hasOwnProperty("fontStrikeout")){
                    fontStrikeout = data[i][j].style.fontStrikeout
                    fontStrikeoutArr.push({"r":i,"c":j,"fontStrikeout":fontStrikeout})
                }
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&fontStrikeoutArr.length==1){
        return fontStrikeoutArr[0].fontStrikeout
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return fontStrikeoutArr
    }
}

/**
 * @api {null} WB.fontUnderline(R1,C1,R2,C2,bool,Index) fontUnderline
 * @apiName fontUnderline
 * @apiGroup Function
 * @apiDescription Underline of cells in the setting range (no parameter defaults to underlining of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} [boolean=true]   
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fontUnderline(0,0,2,2) 
 */
Workbook.prototype.fontUnderline = function(R1,C1,R2,C2,bool,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable, b = (bool!==undefined)?bool:true,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.fontUnderline = b;
            }
        }   
    }
}

/**
 * @api {null} WB.getFontUnderline(R1,C1,R2,C2,Index) getFontUnderline
 * @apiName getFontUnderline
 * @apiGroup Function
 * @apiDescription Get whether the cells in the range are underlined (no parameter defaults to the underline state of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getFontUnderline(0,0)
 */
Workbook.prototype.getFontUnderline = function(R1,C1,R2,C2,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,fontUnderlineArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var fontUnderline
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.hasOwnProperty("fontUnderline")){
                    fontUnderline = data[i][j].style.fontUnderline
                    fontUnderlineArr.push({"r":i,"c":j,"fontUnderline":fontUnderline})
                }
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&fontUnderlineArr.length==1){
        return fontUnderlineArr[0].fontUnderline
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return fontUnderlineArr
    }
}

/**
 * @api {null} WB.getValue(R1,C1,R2,C2,Index) getValue
 * @apiName getValue
 * @apiGroup Function
 * @apiDescription Get the value of the cell (no parameter defaults to get the value of the currently selected cell)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getValue(0,0)
 */
Workbook.prototype.getValue = function(R1,C1,R2,C2,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,valueArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var value;
            if(data[i]&&data[i][j]){
                value = data[i][j].value
                valueArr.push({"r":i,"c":j,"value":value})
            }
        }
    };
    if(((r1==r2&&c1==c2)||isSpan)&&valueArr.length==1){
        return valueArr[0].value
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return valueArr
    }
}

/**
 * @api {null} WB.getText(R1,C1,R2,C2,Index) getText
 * @apiName getText
 * @apiGroup Function
 * @apiDescription Get the value of the text attribute of the active cell, if there is no text attribute, get the value of the value attribute(no parameter defaults to the value of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getText(0,0) 
 */
Workbook.prototype.getText = function(R1,C1,R2,C2,Index){   //获取激活单元格文本
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,textArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var text;
            if(data[i]&&data[i][j]){
                var cell = data[i][j];
                if(cell.hasOwnProperty('text')){
                    text = data[i][j].text
                }else if(!cell.hasOwnProperty('text')&&cell.hasOwnProperty('value')){
                    var value =  cell.value;
                    if(typeof(value)=='object'){
                        text = JSON.stringify(value)
                    }else{
                        if(value !== null && value !== undefined){
                            text = value.toString()
                        }
                    }
                } 
                textArr.push({"r":i,"c":j,"text":text})
            }
        }
    };
    if(((r1==r2&&c1==c2)||isSpan)&&textArr.length==1){
        return textArr[0].text
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return textArr
    }
}

/**
 * @api {null} WB.hdrHeight(height,index) hdrHeight
 * @apiName hdrHeight
 * @apiGroup Function
 * @apiDescription Set the height of the column header
 * @apiParam {Number} height  height<=0  Will hide the column header and the height is 0
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.hdrHeight(40)  //Set the height of the column header to 40px
 */
Workbook.prototype.hdrHeight = function(height,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    if(height>0){
        activeSheet.colHeaderData.defaultH = height;
    }else{
        activeSheet.colHeaderData.showColHeading=false;
        activeSheet.colHeaderData.defaultH = 0;
    }
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.hdrWidth(width,index) hdrWidth
 * @apiName hdrWidth
 * @apiGroup Function
 * @apiDescription Set the width of the row header
 * @apiParam {Number} width  width<=0  Will hide the row header and the width is 0
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.hdrWidth(40)  //Set the width of the row header to 40px
 */
Workbook.prototype.hdrWidth = function(width,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    if(width>0){
        activeSheet.rowHeaderData.defaultW = width
    }else{
        activeSheet.rowHeaderData.showRowHeading=false;
        activeSheet.rowHeaderData.defaultW = 0
    }
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.showColHeading(boolean,index) showColHeading
 * @apiName showColHeading
 * @apiGroup Function
 * @apiDescription Whether to display column headers
 * @apiParam {Boolean} boolean
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.showColHeading(false) 
 */
Workbook.prototype.showColHeading = function(bool,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    activeSheet.colHeaderData.showColHeading = bool ;
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.showRowHeading(boolean,index) showRowHeading
 * @apiName showRowHeading
 * @apiGroup Function
 * @apiDescription Whether to display row headers
 * @apiParam {Boolean} boolean
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.showRowHeading(false)
 */
Workbook.prototype.showRowHeading  = function(bool,Index){   
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    activeSheet.rowHeaderData.showRowHeading  = bool  ;
    this.recalculatePercentage(index)
}

/**
 * @api {null} /null setDefaultFontName
 * @apiName setDefaultFontName
 * @apiGroup Function
 * @apiDescription Set the default font of the workbook  
 * @apiParam {string}  fontname Font name 
 * @apiExample {javascript} demo:
    WB.setDefaultFontName('楷体') 
 */
Workbook.prototype.setDefaultFontName = function(fontname){
    this.workbook.defaultFontName = fontname;
    this._defaultCellStyle.font = this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName;
}

/**
 * @api {null} /null setDefaultFontSize
 * @apiName setDefaultFontSize
 * @apiGroup Function
 * @apiDescription Set the default font size of the workbook 
 * @apiParam {int}  fontsize font size
 * @apiExample {javascript} demo:
    WB.setDefaultFontSize(16)
 */
Workbook.prototype.setDefaultFontSize = function(fontsize){
    this.workbook.defaultFontSize = fontsize;
    this._defaultCellStyle.font = this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName;
}

/**
 * @api {null} /null setDefaultLineHeight
 * @apiName setDefaultLineHeight
 * @apiGroup Function
 * @apiDescription Set the default font line height of the workbook  
 * @apiParam {int}  lineheight Font line height
 * @apiExample {javascript} demo:
    WB.setDefaultLineHeight(20) 
 */
Workbook.prototype.setDefaultLineHeight = function(lineheight){
    this.workbook.defaultLineHeight = lineheight;
    this._defaultCellStyle.font = this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName;
}

/**
 * @api {null} WB.setRowHidden(R1,R2,BHidden,Index) setRowHidden
 * @apiName setRowHidden
 * @apiGroup Function
 * @apiDescription Set the line to hide, if the mouse drags the line height is less than 0 will also hide
 * @apiParam {Int} R1  Start row(R1>=0)
 * @apiParam {Int} R2  End row(R2>=0,R2>=R1)
 * @apiParam {Boolean} BHidden 
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setRowHidden(1,2,true)  //Hide the second and third lines
 */
Workbook.prototype.setRowHidden = function(R1,R2,BHidden,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rows = activeSheet.rows;
    if(BHidden){
        for(var i = R1;i<=R2;i++){
            if(rows[i]&&!rows[i].bHidden){
                activeSheet.rows[i].bHidden = true;
                activeSheet.rows[i].tempSize = rows[i].size;
                this.totolHeight = this.totolHeight - rows[i].size;
                activeSheet.rows[i].size = 0;
            }
        }
    }else{
        for(var j = 0;j<=rows.length-1;j++){
            if(rows[j].bHidden&&j<=R2&&j>=R1){
                delete activeSheet.rows[j].bHidden 
                activeSheet.rows[j].size = rows[j].tempSize;
                this.totolHeight = this.totolHeight + rows[j].size
                delete activeSheet.rows[j].tempSize
            }
        }
    }
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.setColHidden(C1,C2,BHidden,Index) setColHidden
 * @apiName setColHidden
 * @apiGroup Function
 * @apiDescription Set the column to hide, if the mouse drags the column width is less than 0 will also hide
 * @apiParam {Int} C1  Start column(C1>=0)
 * @apiParam {Int} C2  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} BHidden 
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setColHidden(1,2,true)  //Hide the second and third columns
 */
Workbook.prototype.setColHidden = function(C1,C2,BHidden,Index){    //设置列隐藏
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    cols = activeSheet.columns;
    if(BHidden){
        for(var i = C1;i<=C2;i++){
            if(cols[i]&&!cols[i].bHidden){
                activeSheet.columns[i].bHidden = true;
                activeSheet.columns[i].tempSize = cols[i].size;
                this.totolWidth = this.totolWidth - cols[i].size;
                activeSheet.columns[i].size = 0;
            }
        }
    }else{
        for(var j = 0;j<=cols.length -1;j++){
            if(cols[j].bHidden&&j<=C2&&j>=C1){
                delete activeSheet.columns[j].bHidden;
                activeSheet.columns[j].size = cols[j].tempSize;
                this.totolWidth = this.totolWidth + cols[j].size;
                delete activeSheet.columns[j].tempSize;
            }

        }
    };
    this.recalculatePercentage(index)
}

/**
 * @api {null} WB.mergeCells(R1,C1,R2,C2,Index) mergeCells
 * @apiName mergeCells
 * @apiGroup Function
 * @apiDescription Merge cells (no parameter defaults to the merge of the currently selected range)
 * - The merge will change the existing merged data, and only the content of the first cell will be retained after the merge
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.mergeCells(2,1,3,2)   
 */
Workbook.prototype.mergeCells = function(R1,C1,R2,C2,Index){     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var rowCount,colCount,spans = activeSheet.spans;
    var  rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    rowCount = r2 - r1 + 1;
    colCount = c2 - c1 + 1;
    //删除原来的合并
    for(var i = r2;i>=r1;i--){                 //从后面遍历开始删除spans里的元素
        for(var j = c2;j>=c1;j--){
            for(var k = spans.length-1;k>=0;k--){
                if(spans[k].row==i&&spans[k].col==j){
                    activeSheet.spans.splice(k,1)
                 }
            }  
        }
    };
    if(r1 != r2 || c1 != c2){                     //先删除原来的再在spans中插入新的合并数据
        activeSheet.spans.push({"row":r1,"rowCount":rowCount,"col":c1,"colCount":colCount});
        for(var q = r1;q<r1+rowCount;q++){      //合并后只保留左上角单元格的值
            for(var w = c1 ;w<c1+colCount;w++){
                if(!(q==r1&&w==c1)&&activeSheet.data.dataTable[q]&&activeSheet.data.dataTable[q][w]){
                    delete activeSheet.data.dataTable[q][w]
                }
            }
        }
    };
}

/**
 * @api {null} WB.removeMergeCells(R1,C1,R2,C2,Index) removeMergeCells
 * @apiName removeMergeCells
 * @apiGroup Function
 * @apiDescription Cancel merged cells (no parameters default to cancel the merge of the current range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.removeMergeCells(2,1,3,2)
 */
Workbook.prototype.removeMergeCells = function(R1,C1,R2,C2,Index){    //取消合并单元格
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    spans = activeSheet.spans,rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r2;i>=r1;i--){                 //从后面遍历开始删除spans里的元素
        for(var j = c2;j>=c1;j--){
            for(var k = spans.length-1;k>=0;k--){
                if(spans[k].row==i&&spans[k].col==j){
                    activeSheet.spans.splice(k,1)
                 }
            }  
        }
    };
}

/**
 * @api {null}  WB.fixedRowAndCol(R,C,startR,startC,Index） fixedRowAndCol
 * @apiName fixedRowAndCol
 * @apiGroup Function
 * @apiDescription Freeze split pane
 * - fixedRows (How many rows are frozen and which rows start to freeze)
 * - fixedRow (What is the top row of the view after freezing)
 * - fixedCols (How many columns are frozen and which one starts to freeze)
 * - fixedRow (What is the leftmost column of the view after freezing;
 * @apiParam {Int} R In which line to start freezing (the freezing line is below this line) (R>=0)
 * @apiParam {Int} C In which column to start freezing (C>=0)
 * @apiParam {Int} [startR=0]  Start row(startR>=0,startR<R)
 * @apiParam {Int} [startC=0]  Start column(startC>=0,startC<C)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fixedRowAndCol(1,1)   //Freeze the split cell at the position of cell (1,1)
 */
Workbook.prototype.fixedRowAndCol = function(R,C,startR,startC,Index){   //冻结拆分窗格
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var rows = activeSheet.rows,cols = activeSheet.columns;
    if(R>rows.length) R = rows.length;
    if(C>cols.length) C = cols.length;
    var startr = startR || 0;
    var startc = startC || 0;
    if(startr>rows.length-1) startr = rows.length-1;
    if(startc>cols.length-1) startc = cols.length-1;
    if(activeSheet.fixedRows>0||activeSheet.fixedCols>0){
        this.removeFixed(index)
    }
    if(R>=0&&C>=0){
        var fixedRows = (R>0)?R:0;
        var fixedRow = (R>0)?startr:0;
        var fixedCols = (C>0)?C:0;
        var fixedCol = (C>0)?startc:0;
        var hideH = (R>0)?this.getHideWH(startr,0,index).hideH:0;
        var hideW = (C>0)?this.getHideWH(0,startc,index).hideW:0;
        if((R>0&&fixedRows<=fixedRow)||(C>0&&fixedCols<=fixedCol)){
            return
        }
        activeSheet.fixed.fixedRows = fixedRows ;     
        activeSheet.fixed.fixedCols = fixedCols;
        activeSheet.fixed.fixedRow = fixedRow;
        activeSheet.fixed.fixedCol = fixedCol;
        activeSheet.fixed.hideH = hideH;
        activeSheet.fixed.hideW = hideW;
    }
    this.getStartAndEnd(index)
}

/**
 * @api {null} WB.fixedFirstRow(startR,Index) fixedFirstRow
 * @apiName fixedFirstRow
 * @apiGroup Function
 * @apiDescription Freeze the first row of the visible area of the active table
 * - Call this function and the existing frozen column will be removed
 * @apiParam {Int} [startR=0]  Set which line is the start row of the visual view
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fixedFirstRow()  
 */
Workbook.prototype.fixedFirstRow = function(startR,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var startr = startR || 0,rows = activeSheet.rows;
    if(startr>rows.length-1) startr = rows.length-1
    var hideH = this.getHideWH(startr,0,index).hideH;
    activeSheet.fixed.fixedRows = startr  + 1;
    activeSheet.fixed.fixedRow = startr;
    activeSheet.fixed.fixedCols = 0;         
    activeSheet.fixed.fixedCol = 0;
    activeSheet.fixed.hideW = 0;
    activeSheet.fixed.fixedW = 0;
    activeSheet.fixed.hideH = hideH;
    this.getStartAndEnd(index)
}

/**
 * @api {null} WB.fixedFirstCol(startC,Index) fixedFirstCol
 * @apiName fixedFirstCol
 * @apiGroup Function
 * @apiDescription Freeze the first column of the visible area of the active table
 * - Call this function and the existing frozen rows will be removed
 * @apiParam {Int} [startC=0]  Set which column is the starting column of the visual view
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.fixedFirstCol() 
 */
Workbook.prototype.fixedFirstCol = function(startC,Index){  
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var cols = activeSheet.columns,startc = startC || 0;
    if(startc>cols.length-1) startc = cols.length-1
    var hideW = this.getHideWH(0,startc,index).hideW;
    activeSheet.fixed.fixedCols = startc  + 1;
    activeSheet.fixed.fixedCol = startc;
    activeSheet.fixed.fixedRows = 0      
    activeSheet.fixed.fixedRow = 0 ;
    activeSheet.fixed.hideH = 0;
    activeSheet.fixed.fixedH = 0;
    activeSheet.fixed.hideW = hideW;
    this.getStartAndEnd(index)
}

/**
 * @api {null} WB.removeFixed(Index) removeFixed
 * @apiName removeFixed
 * @apiGroup Function
 * @apiDescription Unfreeze pane 
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.removeFixed ()  // Remove all freezes
 */
Workbook.prototype.removeFixed = function(Index){       
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    activeSheet.fixed.fixedRows = 0;
    activeSheet.fixed.fixedCols = 0;
    activeSheet.fixed.fixedRow = 0;
    activeSheet.fixed.fixedCol = 0;
    activeSheet.fixed.fixedH = 0;
    activeSheet.fixed.hideH = 0;
    activeSheet.fixed.fixedW = 0;
    activeSheet.fixed.hideW = 0;
    this.getStartAndEnd(index)
}

//获取冻结的时候0行或者0列到开始冻结行或者列隐藏的宽高
Workbook.prototype.getHideWH = function(startR,startC,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var startc = startC || activeSheet.fixed.fixedCol,startr = startR || activeSheet.fixed.fixedRow;
    var row = activeSheet.rows,col = activeSheet.columns,hideH = 0,hideW = 0;
    for(var i = 0;i<startr;i++){
        if(row[i])   hideH += row[i].size;
    };
    for(var j = 0;j<startc;j++){
        if(col[j])  hideW += col[j].size;
    };
    return {"hideH":hideH,"hideW":hideW}
}

//获取冻结线到startCol或者startRow距离
Workbook.prototype.getFixToStart = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var fixedRows = activeSheet.fixed.fixedRows,fixedCols = activeSheet.fixed.fixedCols,
    startRow = activeSheet.startRow,startCol = activeSheet.startCol,row = activeSheet.rows,
    col = activeSheet.columns,fHideH = 0,fHideW = 0;
    if(fixedRows!=0){
        for(var i = fixedRows;i<startRow;i++){        
            if(row[i])   fHideH += row[i].size;
        };
    };
    if(fixedCols!=0){
        for(var i = fixedCols;i<startCol;i++){
            if(col[i])  fHideW += col[i].size;
        }
    }
    return {"fHideH":fHideH,"fHideW":fHideW}
}

//获取冻结区域的宽高度；
Workbook.prototype.getFixedWH = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var fixedRow = activeSheet.fixed.fixedRow,fixedRows = activeSheet.fixed.fixedRows,fixedCol = activeSheet.fixed.fixedCol,
    fixedCols = activeSheet.fixed.fixedCols,row = activeSheet.rows,col = activeSheet.columns,fixedH = 0,fixedW = 0;
    for(var i = fixedRow;i< fixedRows;i++){         
        if(row[i])   fixedH += row[i].size;
    };
    for(var j = fixedCol;j<fixedCols;j++){              
        if(col[j])  fixedW += col[j].size;
    };
    activeSheet.fixed.fixedH = fixedH;
    activeSheet.fixed.fixedW = fixedW;
    return {"fixedH":activeSheet.fixed.fixedH,"fixedW":activeSheet.fixed.fixedW}
}

/**
 * @api {null} WB.showFixedLine(boolean,index) showFixedLine
 * @apiName showFixedLine
 * @apiGroup Function
 * @apiDescription Whether to display the frozen line of the current table
 * @apiParam {Boolean}  boolean
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.showFixedLine(true)   //show
 */
Workbook.prototype.showFixedLine = function(bool,Index){        
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index)
    activeSheet.showFixedLine = bool ;
}

/**
 * @api {null}  WB.setHeaderName(R,C,headerName,Index) setHeaderName
 * @apiName setHeaderName
 * @apiGroup Function
 * @apiDescription Change the names of column headers and row headers
 * @apiParam {Int} R   Change the row header, change the column header when R passes -1(r>=0)
 * @apiParam {Int} C   Change the column header, change the row header when C passes -1 (c>=0)
 * @apiParam {String} setHeaderName  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setHeaderName(1,-1,'name')
    WB.setHeaderName(-1,1,'name')  
 */
Workbook.prototype.setHeaderName = function(R,C,headerName,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var row = activeSheet.rows,col = activeSheet.columns;
    if(C==-1&&R>=0){
        if(row[R])   row[R].name = headerName 
    };
    if(R==-1&&C>=1){
        if(col[C])  col[C].name = headerName 
    }
}

/**
 * @api {null} WB.deleteHeaderNames(R,C,Index) deleteHeaderNames
 * @apiName deleteHeaderNames
 * @apiGroup Function
 * @apiDescription Delete the name of the column header and row header
 * @apiParam {Int} R   When deleting the row header, C passes -1 (remove both the column header and the row header name R>=0 C>=0) R>=0
 * @apiParam {Int} C   R pass -1 (C>=0) when deleting column header
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.deleteHeaderNames(1,-1)
    WB.deleteHeaderNames(-1,1) 
 */
Workbook.prototype.deleteHeaderNames = function(R,C,Index){         //删除列行头name  
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var row = activeSheet.rows,col = activeSheet.columns;
    if(R>=0){
        if(row[R]&&row[R].hasOwnProperty('name')){
            delete row[R].name 
        }  
    };
    if(C>=0){
        if(col[C]&&col[C].hasOwnProperty('name')){
            delete col[C].name
        }
    }
}

/**
 * @api {null} WB.getStartAndEnd(index) getStartAndEnd
 * @apiName getStartAndEnd
 * @apiGroup Function
 * @apiDescription Get the start row、start column、end row and end column of the view
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getStartAndEnd()
 */
Workbook.prototype.getStartAndEnd = function(Index){       
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var coordY = activeSheet.scrollOption.coordY,coordX = activeSheet.scrollOption.coordX,sumWidth = 0,sumHeight = 0,
    sumWidth2 = 0,sumHeight2 = 0,cols = activeSheet.columns,rows = activeSheet.rows,hideW = activeSheet.fixed.hideW,
    hideH = activeSheet.fixed.hideH;
    var scrollHeight = (coordY/this.percentageV).toFixed(2)-0;   //滚动条滚动了多少高度（往上滚动了多少行）
    var scrollWidth = (coordX/this.percentageH).toFixed(2)-0;
    var fixedWH = this.getFixedWH(index),fixedW = fixedWH.fixedW,fixedH = fixedWH.fixedH,
    maxY =  this.getYMaxScroll(),maxX = this.getXMaxScroll();
    //计算竖向的偏移（用于开始行做一边点偏移，使组后一行显示完毕）
     if(coordY>=maxY&&(this.totolHeight+this.numRowH)>this.height){                                                                                
        this.addHeight =  this.tabsHeight;
    }else{
        this.addHeight = 0;
    };

    for(var k = 0;k<cols.length;k++){   //计算可视视图的开始列
        sumWidth += cols[k].size;
        if(scrollWidth+fixedW+hideW<sumWidth){
            activeSheet.startCol = k ;
            break;
        };
    };
    for(var q = 0;q<rows.length;q++){   // 计算可视视图的开始行
        sumHeight += rows[q].size;
        if(scrollHeight+fixedH+hideH+this.addHeight<sumHeight){
            activeSheet.startRow = q ;
            break;
        }
    };
    for(var w = 0;w<cols.length;w++){   //计算可视视图的结束列
        sumWidth2 += cols[w].size;
        if(this.width-this.numColW+scrollWidth+hideW<=sumWidth2){
            activeSheet.endCol = w;
            break;
        }else{
            activeSheet.endCol = cols.length -1 ;
        }
    };
    for(var e = 0;e<rows.length;e++){   //计算可视视图的结束行
        sumHeight2 += rows[e].size;
        if((this.height-this.numRowH+scrollHeight+hideH)<=sumHeight2){
            activeSheet.endRow = e;
            break;
        }else{  
            activeSheet.endRow = rows.length-1;
        }
    };

    var startCol = activeSheet.startCol,startRow = activeSheet.startRow,endCol = activeSheet.endCol,endRow = activeSheet.endRow,
    sumX = 0,sumY = 0;
    //计算竖向的偏移（用于开始行做一边点偏移，使最后一行显示完毕）
    if(coordY>=maxY&&(this.totolHeight+this.numRowH)>this.height&&coordY>0){
        for(var i = startRow;i<=endRow;i++){
            if(rows[i])  sumY += rows[i].size;
        };    
        activeSheet.scrollOption.offsetY = (sumY + this.numRowH + fixedH - this.height ) + this.tabsHeight;
    }else{
        activeSheet.scrollOption.offsetY = 0;
    };

    //计算水平的偏移
    if(coordX>=maxX&&(this.totolWidth+this.numColW)>this.width&&coordX>0){
        for(var j = startCol;j<=endCol;j++){
            if(cols[j])  sumX += cols[j].size;
        };
        activeSheet.scrollOption.offsetX = sumX + this.numColW + fixedW- this.width;
    }else{
        activeSheet.scrollOption.offsetX = 0;
    };
    return {"startRow":activeSheet.startRow,"startCol":activeSheet.startCol,"endRow":activeSheet.endRow,"endCol":activeSheet.endCol}
}

/**
 * @api {null} WB.setScrollPosition(x,y,Index) setScrollPosition
 * @apiName setScrollPosition
 * @apiGroup Function
 * @apiDescription  Set the position of the scroll bar
 * - In case of freezing, this function should be placed after the freezing function
 * @apiParam {Number} x  
 * @apiParam {Number} y  
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setScrollPosition = function(x,y,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var totolAndPercent = this.getTotolAndPercent();
    this.totolHeight = totolAndPercent.totolHeight;
    this.totolWidth = totolAndPercent.totolWidth;
    this.percentageV = totolAndPercent.percentageV;
    this.percentageH = totolAndPercent.percentageH;
    activeSheet.scrollOption.coordY = (y*this.percentageV).toFixed(5)-0;
    activeSheet.scrollOption.totolY = (y*this.percentageV).toFixed(5)-0;
    activeSheet.scrollOption.coordX = (x*this.percentageH).toFixed(5)-0;
    activeSheet.scrollOption.totolX = (x*this.percentageH).toFixed(5)-0;
    var maxY = this.getYMaxScroll(),maxX = this.getXMaxScroll()
    if(activeSheet.scrollOption.coordY>=maxY){  
        activeSheet.scrollOption.coordY = maxY;
        activeSheet.scrollOption.totolY = maxY;
    }
    if(activeSheet.scrollOption.coordY<0||this.percentageV >= 1){
        activeSheet.scrollOption.coordY = 0;
        activeSheet.scrollOption.totolY = 0;
    }
     if(activeSheet.scrollOption.coordX>=maxX){
        activeSheet.scrollOption.coordX = maxX;
        activeSheet.scrollOption.totolX = maxX;
    }
    if(activeSheet.scrollOption.coordX<0||this.percentageH >= 1){
        activeSheet.scrollOption.coordX = 0;
        activeSheet.scrollOption.totolX = 0;
    }
}


/**
 * @api {null}  WB.setFillColor(color,R1,C1,R2,C2,Index) setFillColor
 * @apiName setFillColor
 * @apiGroup Function
 * @apiDescription Set the fill color of the cells in the range (no parameter defaults to the fill color of the currently selected range)
 * @apiParam {String} color  
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setFillColor('pink',0,0,2,2)   //Set the fill color of the cells in the range to pink
 */
Workbook.prototype.setFillColor = function(color,R1,C1,R2,C2,Index){      
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    if(color){
        for(var i = r1;i<=r2;i++){         
            for(var j = c1;j<=c2;j++){    
                this.createData(i,j,index)
                if(data[i]&&data[i][j]&&data[i][j].style){
                    data[i][j].style.bgColor = color ; 
                };
            }       
        };
    }
};

/**
 * @api {null}  WB.getFillColor(R1,C1,R2,C2,Index) getFillColor
 * @apiName getFillColor
 * @apiGroup Function
 * @apiDescription Get the fill color of the cells in the range (no parameter defaults to the fill color of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getFillColor(0,0)  
 */
Workbook.prototype.getFillColor = function(R1,C1,R2,C2,Index){       
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,fillColorArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var bgColor;
            if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.bgColor){
                bgColor = data[i][j].style.bgColor;
                fillColorArr.push({"r":i,"c":j,"bgColor":bgColor})
            }
        }
    }

    if(((r1==r2&&c1==c2)||isSpan)&&fillColorArr.length==1){
        return fillColorArr[0].bgColor
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return fillColorArr
    }
}

/**
 * @api {null} WB.setRowColor(color,oddEven,start,Index) setRowColor
 * @apiName setRowColor
 * @apiGroup Function
 * @apiDescription Set row color（Interlace color change, equivalent to setting the fill color of the cell) 
 * @apiParam {String} color=rgb(255,255,224)  
 * @apiParam {Number} [oddEven=0]  
 * - 0 : Even rows
 * - 1 ：Odd rows
 * @apiParam {Number} [start=0] From which row to start;
 * - 0 ：Start at the top of the table
 * - 1 ：Start from the following part of the frozen row
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setRowColor('rgb(255,255,224)',0)
 */
Workbook.prototype.setRowColor = function(color,oddEven,start,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,r2 = activeSheet.rows.length - 1,c2 = activeSheet.columns.length - 1,
    startRow = activeSheet.startRow;
    activeSheet.rowStyle = {
        "color":color || 'rgb(255,255,224)',
        "oddEven":oddEven || 0 ,
        "start":start || 0
    };
    var startR = (activeSheet.rowStyle.start==0)?0:startRow
    if(activeSheet.rowStyle.oddEven==0||activeSheet.rowStyle.oddEven==1){
        for(var i = startR;i<=r2;i++){ 
            if(i%2!=0&& activeSheet.rowStyle.oddEven==0){    //偶数行变色，跳过奇数行
                continue
            }else if(i%2==0&& activeSheet.rowStyle.oddEven==1){    //奇数行变色，跳过偶数行
                continue
            }        
            for(var j = 0;j<= c2;j++){    
                this.createData(i,j,index)
                if(data[i]&&data[i][j]&&data[i][j].style){
                    data[i][j].style.bgColor = activeSheet.rowStyle.color ; 
                };
            }       
        };
    }
}

/**
 * @api {null} WB.removeRowColor(Index) removeRowColor
 * @apiName removeRowColor
 * @apiGroup Function
 * @apiDescription Cancel interlace discoloration
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.removeRowColor = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,r2 = activeSheet.rows.length - 1,c2 = activeSheet.columns.length - 1,
    startRow = activeSheet.startRow,startR = (activeSheet.rowStyle.start==0)?0:startRow
    for(var i = startR;i<=r2;i++){ 
        if(i%2!=0&&activeSheet.rowStyle.oddEven==0){     //跳过奇数行
            continue
        }else if(i%2==0&&activeSheet.rowStyle.oddEven==1){   //跳过偶数行
            continue
        }        
        for(var j = 0;j<= c2;j++){ 
            if(data[i]&&data[i][j]&&data[i][j].style){
                delete   data[i][j].style.bgColor
            }
        }
    }
}

/**
 * @api {null} WB.deleteFillColor(R1,C1,R2,C2,Index) deleteFillColor
 * @apiName deleteFillColor
 * @apiGroup Function
 * @apiDescription Delete the fill color of the cells in the range (no parameter clears the fill color in the range by default)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.deleteFillColor(0,0,2,2)   
 */
Workbook.prototype.deleteFillColor = function(R1,C1,R2,C2,Index){           
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index)
    var data = activeSheet.data.dataTable,rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]&&data[i][j].style){
                delete data[i][j].style.bgColor
            }
        }   
    }
}

//输入框
Workbook.prototype.setCanvasInput = function (width, height, top, left, r, c) {      
    var activeSheet = this.activeSheet,data =activeSheet.data.dataTable,textarea = document.querySelector("#"+this.boxId+" .textarea"),
    textAlign = 'left',middleAlign = 'middle',Textline,re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    var font = this.getFont(r,c,r,c)||this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName;
    var align = this.getAlignment(r,c,r,c);
    if(align&&align.hAlign){
        var hAlign = align.hAlign;
        textAlign =  (hAlign==2)? "left":(hAlign==3)?"center":(hAlign==4)?"right":"left";
        var v = this.getValue(r,c,r,c)
        if(hAlign==1&&re.test(v)){
            textAlign = "right";
        }else if(hAlign==1&&!re.test(v)){
            textAlign = "left";
        }
    };
    if(align&&align.vAlign){
        var vAlign =  align.vAlign;
        middleAlign = (vAlign==1)?"top":(vAlign==2)?"middle":(vAlign==3)?"bottom":"middle";
    }
    var fontColor = this.getFontColor(r,c,r,c)||"#000",bgColor = this.getFillColor(r,c,r,c)||"#FFF",
    wordWrap = this.getWordWrap(r,c,r,c),fontUnderline = this.getFontUnderline(r,c,r,c),
    fontStrikeout = this.getFontStrikeout(r,c,r,c);
    if(fontUnderline&&!fontStrikeout){
        Textline = 'underline';
    }else if(!fontUnderline&&fontStrikeout){
        Textline = 'line-through';
    }else if(fontUnderline&&fontStrikeout){
        Textline = 'underline line-through';
    }else{
        Textline = 'none'
    };     

    var isBreakLine = (wordWrap)?"word-break:break-all":"white-space:nowrap",isHasHBar = this.isHasHBar(),isHasVBar = this.isHasVBar();
    var scrollHS = (isHasHBar&&this.isShowBar)?this.scrollSize:0;
    var scrollVS = (isHasVBar&&this.isShowBar)?this.scrollSize:0;
    var tabH = (this.verticalLine<this.height)?0:this.tabsHeight

    if((top+height)>=(this.verticalLine-tabH-scrollHS)){      //针对最后行列的textarea框超出canvas或盖住滚动条的情况，height/width取剩余宽高
        height = (this.verticalLine-tabH-scrollHS) - top  ;
    };

    if((left+width)>(this.horizontalLine-scrollVS)){     
        width = (this.horizontalLine-scrollVS) - left;
    };

    //cssText(添加多个style ie使用cssText)
    if (textarea) {         //+boxT    
        textarea.style.cssText = 'top:'+top+'px;left:'+left+'px;max-height:'+height+'px;max-width:'+width+'px;position:absolute;overflow:hidden;'
        textarea.setAttribute('r', r)
        textarea.setAttribute('c', c)
        document.querySelector("#"+this.boxId+" .child").style.cssText= isBreakLine+';min-width:'+width+'px;max-width:'+width+'px;height:'+height+'px;font:'+font+
        ';display:table-cell;overflow:hidden;vertical-align:'+ middleAlign+';background:'+bgColor+';text-align:'+textAlign+";color:"+fontColor+";text-decoration:"+
        Textline+";box-sizing:border-box;user-select:text;-moz-user-select:text;-webkit-user-select:text;-ms-user-select:text";
    }
    if((top>this.height-this.tabsHeight-this.scrollSize)||(left>this.width-this.scrollSize)||height<=0||width<=0){    //超出隐藏textarea框
        textarea.style.display = "none";
    };
    var child = document.querySelector("#"+this.boxId+" .child");

    child.focus() //双击单元格时取得焦点
    var row = textarea.getAttribute('r'),col = textarea.getAttribute('c');
    var text = '';
    if(this.tempValue.r>=0&&this.tempValue.c>=0){
        text = child.innerText
    }else{
        if(data[row]&&data[row][col]){
            var cell = data[row][col]
            if(cell.text){  
                text =  cell.text;
            }else if(!cell.hasOwnProperty('text')&&cell.hasOwnProperty('value')){
                if(typeof(cell.value)=='object'){
                    text = JSON.stringify(cell.value)                
                }else{
                    if(cell.value !== null && cell.value !== undefined){
                        text = cell.value.toString();
                    }
                }
            }
        }
    }

    var textLength = text.length   //取没替换之前的不然一个空格的长度等于&nbsp;字符的长度
    var formula = this.getCellFormula(row,col,row,col)
    var formatter = this.getCellFormat(row,col,row,col);
    //编辑状态中  若单元格有公式  则编辑状态显示的内容为公式
    if(formula&&formatter!='text'){
        text = formula;
    }
    child.innerText = text

    /**
     * @api {null} /null onshoweditor
     * @apiName onshoweditor
     * @apiGroup Event
     * @apiDescription Select text according to the set editSelStart (start position of the cursor) and editSelEnd (end position of the cursor)
     * - If there is a callback but these two parameters are not set, all text is selected by default
     * - If there is no callback, the cursor is at the start position by default
     * - child is a text edit box
     * @apiParam {Function} callback 
     * @apiExample {javascript} demo:
     *  WB.onshoweditor=function(child){
    var text = child.innerHTML;
    WB.editSelStart = text.length;
    WB.editSelEnd = text.length;
 }
     */
    if(typeof(this.onshoweditor)=='function'&&this.workbook.stopEventCount==0){
        this.editSelStart=-1;
        this.onshoweditor(child)
    }
     
    if(this.editSelStart<0){
        this.editSelStart = 0;
        this.editSelEnd = textLength;
    }else if(this.editSelEnd<0){
        this.editSelEnd = textLength;
    }

    var temp;
    if(this.editSelStart>this.editSelEnd){
        temp = this.editSelEnd;
        this.editSelEnd = this.editSelStart;
        this.editSelStart = temp;
    };
   
    this.editSelStart = (this.editSelStart>textLength)?textLength:this.editSelStart;
    this.editSelEnd = (this.editSelEnd>textLength)?textLength:this.editSelEnd;
    var textNode = child.firstChild,isOverflow = false;
    if(textNode&&textNode.nodeType==3&&(this.editSelStart>textNode.nodeValue.length||this.editSelEnd>textNode.nodeValue.length)){
        isOverflow = true;
    };
    if(textNode&&textNode.nodeType==1){
        isOverflow = true;
    }

    if(window.getSelection&&textNode){
        var sel = window.getSelection();
        var range = document.createRange();
        range = range.cloneRange()
       //如果结束节点类型是 Text， Comment, or CDATASection之一, 那么 endOffset指的是从结束节点算起字符的偏移量。 对于其他 Node 类型节点， endOffset是指从结束结点开始算起子节点的偏移量。
       if(typeof(this.onshoweditor)=='function'&&!isOverflow){
            range.setStart(textNode,this.editSelStart);         //<div>批<br>示</div>
            range.setEnd(textNode,this.editSelEnd);
       }else{    //没有指定范围怎把光标移至最后
            range.selectNodeContents(textNode)    //Range 的 startContainer 属性和 endContainer 属性保存的 Node 将被设置为 referenceNode。 startOffset 为 0,  endOffset 将被设置为 referenceNode 的字符数或子节点个数
            range.collapse(false)           //false折叠但range的end节点 
       }    
        sel.removeAllRanges()
        sel.addRange(range);  
    }else if(document.selection&&textNode){
        var end = this.editSelEnd - textLength;
        var range = document.selection.createRange();
        range.moveToElementText(child);
        if(typeof(this.onshoweditor)=='function'&!isOverflow){
             //基本单位量可以是字符(character),单词(word),句子(sentence)
            range.moveStart("character",this.editSelStart);   // 第二个参数正数开始点向前移动几个位置
            range.moveEnd("character",end);    //  第二个位置 正数，终点向后移动几个位置，负数向前移动几个位置，0终点再文字最后 不移动
        }else{
            range.collapse(false);//光标移至最后
        }
        range.select();
    }
    this.tempValue = {"r":r,"c":c};
};

//移除输入框并完成输入
Workbook.prototype.removeTextArea = function(e){  
    var textarea = document.querySelector("#"+this.boxId+" .textarea"),activeSheet = this.getActiveSheet(),
    r = activeSheet.rangeRow1,c = activeSheet.rangeCol1,child = document.querySelector("#"+this.boxId+" .child"),
    mode = this.workbook.behaviorMode;
    this.keyDownCount = true
    this.endEdit(e) //失去焦点完成输入
    if(textarea){
        textarea.style.cssText =  'top:' +  0 + 'px;left:' +0+'px;max-height:'+0+'px;max-width:'+0+'px;overflow:hidden;position:absolute'
        textarea.setAttribute('r',r);
        textarea.setAttribute('c',c)
    }
    if(child){
        child.focus();
        child.style.cssText = 'height:'+0+'px;width:'+0+'px;overflow:hidden';
    };
    if(mode == 2|| mode == 3){
        this.isEdit = true ;  
        this.isFocus = true;
    }else{
        this.isEdit = false ;  
        this.isFocus = false;
    }
    this.tempValue = {};
}

//完成输入
Workbook.prototype.endEdit = function(e){
    var activeSheet = this.getActiveSheet(),child = document.querySelector("#"+this.boxId+" .child"),
    dataTable = activeSheet.data.dataTable,formulaReg = /^=.+/,rows = activeSheet.rows;
    if(e&&e.keyCode==13){     
        this.stopDefault(e)
    };
    if(this.tempValue.r>=0&&this.tempValue.c>=0){
        var row = this.tempValue.r,col = this.tempValue.c,text = child.innerText,value = this.getValue(row,col);
        if(text!=value){
            if(child.innerText.length!=0){
                this.createData(row,col)
            };
            //如果时自动换行完成输入时自动增加行高
            if(dataTable[row]&&dataTable[row][col]&& dataTable[row][col].style){
                if(rows[row]&&!rows[row].noDefaultH){
                    var wordWrap = this.getWordWrap(row,col,row,col);
                    if(wordWrap){
                        this.incrementRowheight(row,col);
                    }
                } 
                var formatter = this.getCellFormat(row,col)
                var hAlign = this.getAlignment(row,col).hAlign;    
                var finalResult = this.checkInputType(formatter,child.innerText,hAlign,row,col);
                if(finalResult.v!==undefined&&finalResult.v!==null){
                    this.setValue(row,col,finalResult.v)
                }
                if(finalResult.t){
                    this.setText(row,col,finalResult.t)
                }
                if(finalResult.f){
                    this.setCellFormula(row,col,finalResult.f)
                };
                if(finalResult.format){
                    this.setCellFormat(row,col,row,col,finalResult.format)
                };
                if(finalResult.hAlign){
                    this.setHAlignment(row,col,row,col,finalResult.hAlign)
                };
                //原来单元格有公式 编辑删除之后 不符合正则匹配则把公式属性去除  若删除了所有文本则 value为空
                if(child.innerText.length==0||!formulaReg.test(child.innerText)&&dataTable[row][col].formula){
                    delete dataTable[row][col].formula;
                };
            }
        }
    }
}

//检查输入类型
Workbook.prototype.checkInputType = function(formatter,value,align,r,c){
    var v,t,formula,format,hAlign,formulaReg = /^=.+/;
    var re =/^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/,f = this.onformularesult;
    if(formatter=='text'){
        v = value
    }else{
        if(formulaReg.test(value)){
            formula = value;
            if(typeof(f)=='function'&&this.workbook.stopEventCount==0){
                v = f(formula,r,c);
            }
            if(v=="conflict"){
                v = '';
                formula = '';
            }
        }else{
            v = value
            if(re.test(value)){
                v = Number(value);
            };
            if(value=='false'){
                v = false
            }
            if(value=='true'){
                v = true
            };
            var dateValue = this.getFullData(value);    
            if(this.typeObj(dateValue)=='Array'){   //判断是否符合日期格式
                var type = dateValue[3];
                var ft = dateValue[4] || '/'; 
                v = parseInt(this.jsDateToExcelDate(new Date(dateValue[0]+'-'+dateValue[1]+'-'+dateValue[2])));
                t = this.excelDateToJSDate(Number(v),ft)
                format = (formatter)?formatter:type;
                if(!(align==2||align==3)){
                    hAlign = 4;
                }
            }
        }
    }
    return {"v":v,"t":t,"f":formula,"format":format,"hAlign":hAlign}
}

//将Excel中的日期序列号转换为日期（字符串）
Workbook.prototype.excelDateToJSDate = function(numb, format){
    var offset = new Date().getTimezoneOffset()* 60 * 1000;
    var time = new Date((numb-25569)*3600000*24+offset);
    time.setYear(time.getFullYear());
    var year = time.getFullYear();
    var month = time.getMonth() + 1;
    var date = time.getDate();
    if (!format) {
        format='-';
    }
    return year + format + month + format + date
}

// 将日期（字符串）转换为Excel中的日期序列号
Workbook.prototype.jsDateToExcelDate = function(inDate){
    return  (inDate.getTime()- (new Date(1899,11,30)).getTime())/1000/60/60/24;
}
/**
 * @api {null} WB.setCellFormula(R1,C1,Formula,Index) setCellFormula
 * @apiName setCellFormula
 * @apiGroup Function
 * @apiDescription Set cell formula
 * @apiParam {Int} r1  row(r1>=0)
 * @apiParam {Int} c1  column(c1>=0)
 * @apiParam {String} Formula(Must start with = and at least one character after)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setCellFormula(0,0,'=A1+B1')   //Set the cell formula to =A1+B1
 */

 /**
 * @api {null} /null onformularesult
 * @apiName onformularesult
 * @apiGroup Event
 * @apiDescription The event that returns the result of the formula
 * - Function return (formula r c) 
 * @apiParam {Function} callback 
 * @apiExample {javascript} demo:
    WB.onformularesult = function(f,r,c){
        //do something...
    }
 */
Workbook.prototype.setCellFormula = function(R1,C1,Formula,Index){     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,formulaReg = /^=.+/,s = '';
    if(formulaReg.test(Formula)){
        this.createData(R1,C1,index)
        if(typeof(this.onvaluechange)=='function'&&this.workbook.stopEventCount==0){
            s = this.onvaluechange(R1,C1,Formula,data[R1][C1].formula);
            if(s=='-1') return
        }
        var f = this.onformularesult;
        if(data[R1]&&data[R1][C1]&&this.workbook.stopEventCount==0){
            data[R1][C1].formula = Formula;
            var formula =  data[R1][C1].formula;
            var formatter = data[R1][C1].style.formatter;
            if(formatter!='text'){               //不是文本格式 则输出公式结果
                if(typeof(f)=='function'){
                    var result = f(formula,R1,C1);
                    this.setValue(R1,C1,result)
                }
            }
        }
        if(s=='1'&&typeof(this.onrelevantchange)=='function'){
            this.onrelevantchange(R1,C1)
        }
    } 
};

/**
 * @api {null} WB.getCellFormula(R1,C1,R2,C2,Index) getCellFormula
 * @apiName getCellFormula
 * @apiGroup Function
 * @apiDescription Get the formula of the cell (the default is to get the formula of the currently selected range without parameters)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getCellFormula(1,1) 
 */
Workbook.prototype.getCellFormula = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,cellFormulaArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var formula;
            if(data[i]&&data[i][j]&&data[i][j].formula){
                formula = data[i][j].formula
                cellFormulaArr.push({"r":i,"c":j,"formula":formula})
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&cellFormulaArr.length==1){
        return cellFormulaArr[0].formula
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return cellFormulaArr
    }
}

 //初始化一个输入框
Workbook.prototype.initTextArea = function(){                     
    var hiddenTextarea = document.createElement('div')
    , textareaChild = document.createElement('div');
    hiddenTextarea.style.cssText  =  'top:' +  0 + 'px;left:' +0+'px;max-height:'+0+'px;max-width:'+0+'px;overflow:hidden;';
    hiddenTextarea.setAttribute('class', 'textarea')
    hiddenTextarea.setAttribute('r', 0)
    hiddenTextarea.setAttribute('c', 0)
    document.getElementById(this.boxId).appendChild(hiddenTextarea);
    textareaChild.setAttribute('class','child')
    textareaChild.setAttribute('contenteditable', 'true')
    textareaChild.style.cssText = 'height:'+0+'px;width:'+0+'px;overflow:hidden;'
    document.querySelector("#"+this.boxId+" .textarea").appendChild(textareaChild);
    var child = document.querySelector("#"+this.boxId+" .child");
    child.focus() ;   
}

//判断是当前单元格是否有光标聚焦
Workbook.prototype.getFocus = function(){               
    var textarea = document.querySelector("#"+this.boxId+" .textarea");
    if(textarea){
        var top = parseInt(textarea.style.top);
        var left = parseInt(textarea.style.left);
        var maxW = parseInt(textarea.style.maxWidth);
        var maxH = parseInt(textarea.style.maxHeight);
        if(!(top==0&&left===0&&maxW==0&&maxH==0)){
            this.isFocus = true;
        }else{
            this.isFocus = false
        }
        return this.isFocus
    }
}

//获取输入框定位的额外距离
Workbook.prototype.getBoxDiff = function(){                                        
    var scrollT = document.documentElement.scrollTop;
    var scrollL = document.documentElement.scrollLeft;
    var box = this.canvas.getBoundingClientRect();
    var boxLeft =  box.left ;    //
    var boxTop  =   box.top ;
    var boxT = scrollT+boxTop;
    var boxL = scrollL+boxLeft;
    return({"boxT":boxT,"boxL":boxL})
}

//绘制celltype==1 || celltype==2 的btn
Workbook.prototype.drawBtn = function(r,c,x,y,w,h,bool,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    ar = activeSheet.activeRow,ac = activeSheet.activeCol,
    showBtn = activeSheet.alwaysShowButton,showBtnInArea = activeSheet.alwaysShowInArea,isOverflow = false,cellTypeName,
    rows = activeSheet.rows,cols = activeSheet.columns;
    var cellType = this.getCellType(r,c,r,c);
    if(cellType){
        cellTypeName = cellType.name
    };
    if(((showBtn&&!showBtnInArea)||(showBtn&&showBtnInArea&&r==ar&&c==ac)||(!showBtn&&r==ar&&c==ac&&this.isEdit))&&
    !isOverflow&&rows[r]&&!rows[r].bHidden&&cols[c]&&!cols[c].bHidden&&w>=this.cellBtnAreaWidth){
        if(cellTypeName==1||cellTypeName==2){
            //背景色
            this.ctx.save();
            this.ctx.beginPath();
            this.ctx.fillStyle = (bool)?this.btnFillColorH:this.btnFillColor;
            this.ctx.strokeStyle = (bool)?this.btnBorderColorH:this.btnBorderColor;
            //填充
            this.ctx.fillRect(x+w-this.cellBtnAreaWidth,y+1,this.cellBtnAreaWidth-1,h-2)
            //border
            this.ctx.lineWidth = 0.5
            this.ctx.moveTo(x+w-this.cellBtnAreaWidth,y+1);
            this.ctx.lineTo(x+w-this.cellBtnAreaWidth,y+h-1);
            this.ctx.lineTo(x+w-1,y+h-1);
            this.ctx.lineTo(x+w-1,y+1);
            this.ctx.closePath();
            this.ctx.stroke()
            this.ctx.restore();   
        };
        if(cellTypeName==2){     
            //三角箭头
            this.ctx.save();
            this.ctx.beginPath();
            this.ctx.strokeStyle = "#000";
            this.ctx.fillStyle = "#000";
            this.ctx.moveTo(x+w-this.cellBtnWidth,y+h/2-this.cellBtnHeight/2);
            this.ctx.lineTo(x+w-this.cellBtnWidth/2-2,y+h/2+this.cellBtnHeight/2-6);
            this.ctx.lineTo(x+w-4,y+h/2-this.cellBtnHeight/2);
            this.ctx.stroke();
            this.ctx.restore(); 
        }else if(cellTypeName==1){
            //三个小圆点
            this.ctx.save();
            this.ctx.beginPath();
            this.ctx.fillStyle = "#585858";
            this.ctx.arc(x+w-this.cellBtnWidth-1,y+h/2-1.5,1.5,Math.PI*2,0,true);
            this.ctx.arc(x+w-this.cellBtnWidth-1+4,y+h/2-1.5,1.5,Math.PI*2,0,true)
            this.ctx.arc(x+w-this.cellBtnWidth-1+8,y+h/2-1.5,1.5,Math.PI*2,0,true)
            this.ctx.fill();
            this.ctx.restore();
        }
    }  
};

//绘制celltype==4 的btn
Workbook.prototype.drawSelectBtn = function(r,c,x,y,w,h,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rows = activeSheet.rows,cols = activeSheet.columns,ar = activeSheet.activeRow,
    ac = activeSheet.activeCol,isOverflow = false,btnW = 6,btnH = 6, showBtn = activeSheet.alwaysShowButton,
    showBtnInArea = activeSheet.alwaysShowInArea;
    if((y>this.height-this.tabsHeight-this.scrollSize)||(x>this.width-this.scrollSize)){    
        isOverflow=true
    };
    if(((showBtn&&!showBtnInArea)||(showBtn&&showBtnInArea&&r==ar&&c==ac)||(!showBtn&&r==ar&&c==ac&&this.isEdit))&&
    !isOverflow&&rows[r]&&!rows[r].bHidden&&cols[c]&&!cols[c].bHidden&&w>=this.cellBtnAreaWidth){
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.fillStyle = "#000"
        this.ctx.moveTo(x+w-this.cellBtnAreaWidth+btnW/2,y+h/2-btnH/2);
        this.ctx.lineTo(x+w-this.cellBtnAreaWidth+btnW,y+h/2+btnH/2);
        this.ctx.lineTo(x+w-this.cellBtnAreaWidth+btnW+btnW/2,y+h/2-btnH/2)
        this.ctx.closePath();
        this.ctx.fill()
        this.ctx.restore();
    }
}

//绘制celltype==3 的btn
Workbook.prototype.drawCheckBox = function(r,c,x,y,w,h,bool,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    rows = activeSheet.rows,cols = activeSheet.columns,v = this.getValue(r,c);
    if(rows[r]&&!rows[r].bHidden&&cols[c]&&!cols[c].bHidden&&w>=this.checkBoxW){
        this.ctx.fillStyle = (bool)?this.checkBoxFillColorH:this.checkBoxFillColor;
        this.ctx.strokeStyle = (bool)?this.checkBoxBorderColorH:this.checkBoxBorderColor;
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.rect(x+w/2-this.checkBoxW/2,y+h/2-this.checkBoxH/2,this.checkBoxW,this.checkBoxH)
        this.ctx.stroke();
        this.ctx.fill()
        this.ctx.restore();
        if(v){
           //勾选：
           this.ctx.save();
           this.ctx.beginPath();
           this.ctx.lineWidth = 2;
           this.ctx.strokeStyle = "green";
           this.ctx.moveTo(x+w/2-this.checkBoxW/2+2,y+h/2);
           this.ctx.lineTo(x+w/2,y+h/2+this.checkBoxH/2-2);
           this.ctx.lineTo(x+w/2+this.checkBoxW/2,y+h/2-this.checkBoxH/2)
           this.ctx.stroke();
           this.ctx.restore();
        }
    }
}

/**
 * @api {null} /null getCellRC
 * @apiName getCellRC
 * @apiGroup Function
 * @apiDescription Get the range of rows and columns of the current table
 * - WB.getCellRC()
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.getCellRC = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var r1 = activeSheet.rangeRow1;
    var c1 = activeSheet.rangeCol1;
    var r2 = activeSheet.rangeRow2;
    var c2 = activeSheet.rangeCol2;
    return {"r1":r1,"c1":c1,"r2":r2,"c2":c2}
}

/**
 * @api {null} WB.setLock(R1,C1,R2,C2,bool,Index) setLock
 * @apiName setLock
 * @apiGroup Function
 * @apiDescription Set whether the cells in the range are locked (no parameter defaults to the locking of the cells in the currently selected range)
 * - The locked cell cannot be edited, no check box is selected
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} [bool=true]  
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setLock = function(R1,C1,R2,C2,bool,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,b = (bool!==undefined)?bool:true,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){         
        for(var j = c1;j<=c2;j++){      
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.lock = b ; 
            };
        }       
    };
};

/**
 * @api {null} WB.getLock(R1,C1,R2,C2,Index) getLock
 * @apiName getLock
 * @apiGroup Function
 * @apiDescription Get whether the cell is locked (no parameter defaults to get the locked state of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getLock(0,0)
 */
Workbook.prototype.getLock = function(R1,C1,R2,C2,Index){    
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,lockArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var lock
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.hasOwnProperty("lock")){
                   lock = data[i][j].style.lock
                   lockArr.push({"r":i,"c":j,"lock":lock})
                }
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&lockArr.length==1){
        return lockArr[0].lock
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return lockArr
    }
}

/**
 * @api {null} WB.setCanEdit(R1,C1,R2,C2,bool,Index) setCanEdit
 * @apiName setCanEdit
 * @apiGroup Function
 * @apiDescription Set whether the cell can be edited (no parameter default setting currently selected range can not be edited)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Boolean} [bool=false]
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setCanEdit(1,1,1,1,false) 
 */
Workbook.prototype.setCanEdit = function(R1,C1,R2,C2,bool,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,b = (bool!==undefined)?bool:false,
    rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){         
        for(var j = c1;j<=c2;j++){      
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.canEdit = b ; 
            };
        }       
    };
}

/**
 * @api {null} WB.getCanEdit(R1,C1,R2,C2,Index) getCanEdit
 * @apiName getCanEdit
 * @apiGroup Function
 * @apiDescription Get whether the cells in the range can be edited (no parameter defaults to the uneditable state of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getCanEdit(0,0)
 */
Workbook.prototype.getCanEdit = function(R1,C1,R2,C2,Index){       
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,canEditArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var canEdit;
            if(data[i]&&data[i][j]&&data[i][j].style){
                if(data[i][j].style.hasOwnProperty("canEdit")){
                    canEdit = data[i][j].style.canEdit
                    canEditArr.push({"r":i,"c":j,"canEdit":canEdit})
                }
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&canEditArr.length==1){
        return canEditArr[0].canEdit
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return canEditArr
    }
}

/**
 * @api {null} WB.setCellFormat(R1,C1,R2,C2,type,Index) setCellFormat
 * @apiName setCellFormat
 * @apiGroup Function
 * @apiDescription Format cells in a range
 * - date：yyyy年m月d日 ,yyyy年m月 ,m月d日,yyyy-m-d ,yyyy-m,m-d ,yyyy-m-d:uppercase ,yyyy-m:uppercase ,星期n ,yyyy/m/d ,yyyy/m,m/d
 * - Scientific notation：E  (0E:Keep a decimal,00E:Keep two decimal places,The number of 0 is the number of reserved decimals)
 * - Numerical format：NR(A+thousands  (Remain as many decimals as there are 0s in front of N, use thousands separators after +thousands after N, and do not use thousands separators if there are none) R(A reference currency
 * - currency：MR(A+$   (Remain as many decimals as there are 0s in front of M, and set the color to red for negative numbers with R, yes (there are parentheses (only when A)), and A represents a positive number display, R(A There is no limit to the three character positions, but it must be placed after M. If there is a + currency symbol behind M, the currency symbol is added in front of it. If not, the currency symbol is not used Yuan display is almost ¥), US dollar, Euro, British pound, French franc, Korean won)
 * - percentage：P+%  (Remain as many decimals as there are 0s in front of P)     
 * - Amount of capital：capitalMoney(Numbers are converted to upper-case amounts, and decimals are reserved to the point)
 * - text：Text format (return as is, including formulas)  
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {String}  type  Valid cell format value
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setCellFormat(0,0,2,c2,"yyyy年m月d日")   //The format of the cells in the set range is the date ("yyyy年m月d日")
 */
Workbook.prototype.setCellFormat = function(R1,C1,R2,C2,type,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){         
        for(var j = C1;j<=C2;j++){      
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style.formatter = type.trim() ;
                if(type!='text'&&data[i][j].style.hAlign==1){
                    data[i][j].style.hAlign = 4;
                }
            };
        }       
    };
};

/**
 * @api {null} WB.getCellFormat(R1,C1,R2,C2,Index) getCellFormat
 * @apiName getCellFormat
 * @apiGroup Function
 * @apiDescription Get the format of the cell (without parameters, the format of the currently selected cell is obtained by default)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getCellFormat(1,1)
 */
Workbook.prototype.getCellFormat = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,CellFormatArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var formatter;
            if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.formatter){
                formatter = data[i][j].style.formatter
                CellFormatArr.push({"r":i,"c":j,"formatter":formatter})
            }
        }
    }    
    if(((r1==r2&&c1==c2)||isSpan)&&CellFormatArr.length==1){
        return CellFormatArr[0].formatter
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return CellFormatArr
    }
}

Workbook.prototype.getFormatValue = function(Value,Type){
    var fullData = this.getFullData(Value);
    var re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    if(re.test(Value)){
        var JSDate = this.excelDateToJSDate(Number(Value), '-')
        var JSDateArr = JSDate.split("-")
    }
    switch(Type){
        case "yyyy年m月d日":         //2019年10月10日    data[2019,10,10]
           if(this.typeObj(fullData)=='Array'){
                Value =  fullData[0]+"年"+fullData[1]+"月"+fullData[2]+"日"
           }else if(re.test(Value)){
                Value =  JSDateArr[0]+"年"+JSDateArr[1]+"月"+JSDateArr[2]+"日"
            }
            break;
        case "yyyy年m月":           //2019年10月
           if(this.typeObj(fullData)=='Array'){
                Value = fullData[0]+"年"+fullData[1]+"月"
           }else if(re.test(Value)){
                Value =  JSDateArr[0]+"年"+JSDateArr[1]+"月"
            }
           break;
        case "m月d日":           //10月10日
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[1]+"月"+fullData[2]+"日"
            }else if(re.test(Value)){
                Value = JSDateArr[1]+"月"+JSDateArr[2]+"日"
            }
            break;
        case "yyyy-m-d":            //2019-10-10
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[0]+"-"+fullData[1]+"-"+fullData[2]
            }else if(re.test(Value)){
                Value = JSDate
            }
            break;
        case "yyyy-m":              //2019-10
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[0]+"-"+fullData[1]
            }else if(re.test(Value)){
                Value = JSDateArr[0]+"-"+JSDateArr[1]
            }
            break;
        case "m-d":              //10-10
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[1]+"-"+fullData[2]
            }else if(re.test(Value)){
                Value = JSDateArr[1]+"-"+JSDateArr[2]
            }
            break;
        case "yyyy-m-d:uppercase":   //
            if(this.typeObj(fullData)=='Array'){
                var year = this.converToUpper(fullData[0].toString());
                var month = this.converToUpper(fullData[1].toString());
                var d = this.converToUpper(fullData[2].toString());
                Value = year+"年"+month+"月"+d+"日"
            }else if(re.test(Value)){
                var year = this.converToUpper(JSDateArr[0].toString());
                var month = this.converToUpper(JSDateArr[1].toString());
                var d = this.converToUpper(JSDateArr[2].toString());
                Value = year+"年"+month+"月"+d+"日"
            }
            break;
        case "yyyy-m:uppercase":   //
            if(this.typeObj(fullData)=='Array'){
                var year = this.converToUpper(fullData[0].toString());
                var month = this.converToUpper(fullData[1].toString());
                Value = year+"年"+month+"月"
            }else if(re.test(Value)){
                var year = this.converToUpper(JSDateArr[0].toString());
                var month = this.converToUpper(JSDateArr[1].toString());
                Value = year+"年"+month+"月"
            }
            break;
        case "星期n":              //根据年月日  转换成星期几
            var weekNum,day = '',m;
            if(this.typeObj(fullData)=='Array'){
                m = new Date(fullData[0],fullData[1]-1,fullData[2]);    //月份记得-1
                weekNum = m.getDay();
            }else if(re.test(Value)){
                m = new Date(JSDateArr[0],JSDateArr[1]-1,JSDateArr[2]);    //月份记得-1
                weekNum = m.getDay();
            }
            switch(weekNum){
                case 0:
                    day = "星期日"
                    break;
                case 1:
                    day = "星期一"
                    break;
                case 2:
                    day = "星期二"
                    break;
                case 3:
                    day = "星期三"
                    break;
                case 4:
                    day = "星期四"
                    break;
                case 5:
                    day = "星期五"
                    break;
                case 6:
                    day = "星期六"
                    break;
            };
            Value = day
            break;
        case "yyyy/m/d":
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[0]+"/"+fullData[1]+"/"+fullData[2]
            }else if(re.test(Value)){
                Value = JSDateArr[0]+"/"+JSDateArr[1]+"/"+JSDateArr[2]
            }
            break;
        case "yyyy/m":
            if(this.typeObj(fullData)=='Array'){
                Value= fullData[0]+"/"+fullData[1]
            }else if(re.test(Value)){
                Value= JSDateArr[0]+"/"+JSDateArr[1]
            }
            break;
        case "m/d":              //10/10
            if(this.typeObj(fullData)=='Array'){
                Value = fullData[1]+"/"+fullData[2]
            }else if(re.test(Value)){
                Value = JSDateArr[1]+"/"+JSDateArr[2]
            }
            break;
        case "capitalMoney":
            Value = this.getCapitalMoney(Value);
            break;
        case "text":
            Value = Value;
            break;
        default :
            Value = this.getRegFormat(Value,Type)
    };
    return Value
}

Workbook.prototype.getRegFormat = function(Value,Type){
    var scicentReg = /^0*E$/;        //匹配科学计数法的正则；（E,0E,00E,....）;
    var numReg = /^0*NR?\(?A?(\+thousands)?$/;           //数值格式（包含千位分隔符）
    var moneyReg = /^0*MR?\(?A?(\+(￥|¥|\$|€|￡|₣|₩))?$/;   //匹配货币正则（人民币、日元、美元、欧元、英镑、法郎、韩元）
    var percentage = /^0*P(\+%)$/               //匹配百分比正则
    switch(true){
        case scicentReg.test(Type):
            Value = this.getScient(Value,Type);
            break;
        case numReg.test(Type):
            Value = this.getNumFormat(Value,Type);
            break;
        case moneyReg.test(Type):
            Value = this.getMoneyFormat(Value,Type);
            break;
        case percentage.test(Type):
            Value = this.getPercentageFormat(Value,Type)
            break;
    };
    return Value
}

//金额大写（保留到角分,整数48位）
Workbook.prototype.getCapitalMoney = function(num){
    var re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/      //用来判断值是否为数字
    if(re.test(num)){
        var fraction = ['角','分'];                                   //小数部分转换
        var digit = ['零','壹','贰','叁','肆','伍','陆','柒','捌','玖'];    //正数部分转换
        var unit = [                                            //单位
            ['元','万','亿','兆','京','垓','秭','穰','沟','涧','正','载'],             //内层循环(暂时支持到48位)
            ['','拾','佰','仟']           //外层循环
        ];
        var finalStr = ''
        var v = Number(num);
        if(re.test(v)){
            var head = v < 0?'欠':'';          //v小于0则在大写金额前面加上欠字
            v = Math.abs(v);
            //处理小数部分
            for(var i = 0;i<fraction.length;i++){    //保留到分（先取角在取分,角*10%10  分*100%10）  
                finalStr+=digit[Math.floor((v*10).toFixed(2)*Math.pow(10,i))%10]+fraction[i] ;//toFixed(2)   防止00.8*10 = 0.79999999的小数不精确问题   
            };
            finalStr = finalStr.replace(/零./g,'');   //把零角或者零分的情况都去掉
            finalStr = finalStr || "整" ;        //小数部分没有值  小数部分字符为   整字
            v = Math.floor(v);   
            var vL = v.toString().length
            if(vL>48){      //大于48位  返回原数据
                return num
            }
            //处理正数部分
            for(var i = 0;i<unit[0].length&&v>0;i++){           //正数取余不用*10   依次/10取每一位，v<0则退出
                var digitStr = '';
                for(var j = 0;j<unit[1].length&&v>0;j++){
                    digitStr = digit[v % 10] + unit[1][j] + digitStr;   
                    v = Math.floor(v/10);
                };
                  //贰 亿 零仟零佰零拾零 万 零仟零佰零拾壹 元 玖角捌分  ——>贰 亿 零 万 零仟零佰零拾壹 元 玖角捌分
                  //贰仟零佰零拾零 万 零仟零佰零拾壹 元 玖角捌分  ->贰仟万 零仟零佰零拾壹 元 玖角捌分
                  //.replace(/^$/,'零')digitStr(零仟零佰零拾零 ""只有是既是它开头又是它结尾则替换零(空)；贰仟零佰零拾零，贰仟"" 不替换零)
                digitStr = digitStr.replace(/(零.)*零$/,'').replace(/^$/,'零')
                finalStr = digitStr + unit[0][i] + finalStr ;         //
            };
            finalStr = finalStr.replace(/(零.)*零元/,'元').replace(/(零.)+/g, '零').replace(/^整$/,'零元整')
            return head + finalStr;
        }else{            //非数字原样输出
            finalStr = num
        }
        return finalStr
    }else{
        return num
    }
    
}
//百分比格式（包含保留多少位小数点）
Workbook.prototype.getPercentageFormat = function(num,type){
    var re = /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    if(re.test(num)){
        var d,v = Number(num);
        var dot = (type.split("0")).length-1;   //求0出现了几次,用于保留几位小数;
        if(type.split("+").length==2&&re.test(v)){           //判断使用有一个千位的分隔符,用+来判断
            v = (v*100).toFixed(dot)
            d = v+"%"
        }else{
            d = num
        }
        return d
    }else{
        return num
    }
}

//货币格式（保留多少位小数以及使用什么货币符号）
Workbook.prototype.getMoneyFormat = function(num,type){
    var re = /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    if(re.test(num)){
        var d,intPart,dotPart,v = Number(num);
        var dot = (type.split("0")).length-1;   //求0出现了几次,用于保留几位小数;
        v = v.toFixed(dot)
        intPart = v.split(".")[0];       //获取整数部分
        dotPart = v.split(".")[1];  
        var isCurrency='',isDotPart='';
        if(type.split("+").length==2&& re.test(v)){
            isCurrency = type.split("+")[1]
        };
        if(dotPart){
            isDotPart = "."+dotPart
        };
        if(Number(v)<0&&re.test(v)){
            if(type.indexOf("(")!=-1){
                d = "("+isCurrency+Math.abs(Number(intPart)).toLocaleString().split(".")[0]+isDotPart+")";
            }else{
                if(type.indexOf("A")!=-1){
                    d = isCurrency+Math.abs(Number(intPart)).toLocaleString().split(".")[0]+isDotPart;
                }else{
                    d = isCurrency+Number(intPart).toLocaleString().split(".")[0]+isDotPart;
                }
            }
        }else if(Number(v)>=0&&re.test(v)){
            d = isCurrency+Number(intPart).toLocaleString().split(".")[0]+isDotPart
        }else{
            d = num
        }
        return d
    }else{
        return num
    }
}

//数值格式（保留多少位小数以及是否使用千位分隔符）
Workbook.prototype.getNumFormat = function(num,type){
    var re = /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    if(re.test(num)){
        var d,intPart,dotPart,v = Number(num);
        var dot = (type.split("0")).length-1;   //求0出现了几次,用于保留几位小数;
        v = v.toFixed(dot)
        intPart = v.split(".")[0];       //获取整数部分
        dotPart = v.split(".")[1];  
        var isDotPart='',excelDateNum;
        var fullData = this.getFullData(num);
        if(this.typeObj(fullData)=='Array'){
            excelDateNum = fullData[0]+"-"+fullData[1]+"-"+fullData[2]
        }else{
            excelDateNum = num
        }
        var excelDate = this.jsDateToExcelDate(new Date(excelDateNum))
        if(dotPart){
            isDotPart = "."+dotPart
        };
        if(Number(v)<0&&re.test(v)){
            if(type.indexOf("A")!=-1){
                intPart =  Math.abs(intPart);
            };
            if(type.split("+").length==2&&re.test(v)){
                if(type.indexOf("(")!=-1){
                    d = "("+Math.abs(Number(intPart)).toLocaleString().split(".")[0]+isDotPart+")";
                }else{
                    d = Number(intPart).toLocaleString().split(".")[0]+isDotPart;
                }
            }else{
                if(type.indexOf("(")!=-1){
                    d = "("+Math.abs(Number(intPart))+isDotPart+")";
                }else{
                    d = intPart+isDotPart
                }
            }
        }else if(Number(v)>=0&&re.test(v)){
            if(type.split("+").length==2&& re.test(v)){
                d = Number(intPart).toLocaleString().split(".")[0]+isDotPart;
            }else{
                d = v;
            }
        }else if(!Number(num)&&!re.test(num)&&excelDate){
            d = Math.floor(excelDate).toFixed(dot)
        }else{
            d = num
        }
        return d
    }else{
        return num
    }
    
}  

//重写toFixed函数防止小数不精确
Number.prototype.toFixed = function (s) {
    var _this = this, changenum, index;
    if (this < 0) {
        _this = -_this;
    }
    changenum = (parseInt(_this * Math.pow(10, s) + 0.5) / Math.pow(10, s)).toString();
    index = changenum.indexOf(".");
    if (index < 0 && s > 0) {
        changenum = changenum + ".";
        for (var i = 0; i < s; i++) {
            changenum = changenum + "0";
        }
    } else {
        var decimal =  (changenum.length -1) - index  //字符总长度除去小数点得出数字的长度  减去index位置（整数长度）就可以得出现在有多少位小数（index位置可知道前面有多少位整数 index=2  前面就有2位）
        for (var i = 0; i < s - decimal; i++) {           //传进来的长度减去现有长度
            changenum = changenum + "0";
        }
    }
    if (this < 0) {
        return "-"+changenum;
    } else {
        return changenum;
    }
}

//数字转科学计数法
Workbook.prototype.getScient = function(num,type){
    var re = /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/;
    if(re.test(num)){
        var v = Math.abs(Number(num)),d;
        var p = Math.floor(Math.log(v)/Math.LN10);   //获取几次方 Math.LN10=Math.log(10)
        //1000 1000000这种数由于精度问题却获得的是2.99999   5.999999
        var vStr = v.toString()
        var zeroRe = /^1(000)+$/
        if(zeroRe.test(vStr)){
            var zeroN = vStr.split('0').length-1;
            p = zeroN
        }
        
        var dot=(type.split("0")).length-1;   //求0出现了几次,用于保留几位小数
        var n = (v * Math.pow(10, -p)).toFixed(dot)
        if(p<10&&p>=0){
            p = "0"+p;
        };

        if(v>=10&&Number(num)>0){
            d = n + 'E' + "+" + p;
        }else if(v>=1&&v<10&&Number(num)>0){
            d = n + 'E' +"+00"
        }else if(v<1&&v>0&&Number(num)>0){
            d = n + 'E' +p;
        }else if(Number(num)<0&&v>=10){
            d = "-" + n + 'E' + "+" + p;
        }else if(Number(num)<0&&v>=1&&v<10){
            d = "-" + n + 'E' +"+00"
        }else if(Number(num)<0&&v<1&&v>0){
            d =  "-" + n + 'E' + p;
        }else{
            d = num
        }
        return d
    }else{
        return num
    }
}

//转大写的日期格式
Workbook.prototype.converToUpper = function(str){
    var chineseArr = ["〇","一","二","三","四","五","六","七","八","九"];
    var converNum = [];   
    for(var i = 0; i < str.length; i++){
        converNum.push(chineseArr[str.charAt(i)]);
    };
    var chineseNum = converNum.join('');
    if(chineseNum.length==2){      //月份  日  作处理   年份原样输出
       if(chineseNum[0]=="一"){
           if(chineseNum[1]=="〇"){
                chineseNum = "十"
           }else{
                chineseNum = "十"+chineseNum[1]
           }
       }else if(chineseNum[0]!="一"){
            if(chineseNum[1]=="〇"){
                chineseNum = chineseNum[0]+"十";
            }else{
                chineseNum = chineseNum[0]+"十"+chineseNum[1]
            }
       }
    }
    return chineseNum
}

//根绝单元格值 获取年月日
Workbook.prototype.getFullData = function(Value){
    if(Value!==undefined){
        Value = Value.toString();
        var myDate = new Date();
        var nowYear = myDate.getFullYear(); //获取当前年
        var dateReg = /年|月|日/g;
        if(dateReg.test(Value)){
            var yearDate = Value.split('年');
            var yearArray;
            if(yearDate.length==2){
                var monthDate = yearDate[1].split('月');
                if(monthDate.length==2){
                   var dayDate = monthDate[1].split('日');
                   if(dayDate.length==2){
                        if(Number(yearDate[0])%1===0&&Number(yearDate[0])>=1900&&Number(yearDate[0])<=9999&&Number(monthDate[0])%1===0&&Number(monthDate[0])>=1&&Number(monthDate[0])<=12&&Number(dayDate[0])%1===0&&Number(dayDate[0])>=1&&Number(dayDate[0])<=31){
                            yearArray = [Number(yearDate[0]),Number(monthDate[0]),Number(dayDate[0]),"yyyy年m月d日"]  //2019年10月10日  ->  2019 10  10
                        }
                   }else if(dayDate.length==1){
                        if(Number(yearDate[0])%1===0&&Number(yearDate[0])>=1900&&Number(yearDate[0])<=9999&&Number(monthDate[0])%1===0&&Number(monthDate[0])>=1&&Number(monthDate[0])<=12){
                            yearArray = [Number(yearDate[0]),Number(monthDate[0]),1,"yyyy年m月"]  //2019年10月  ->  2019 10  1
                        }
                   }else{
                        yearArray = Value
                   }
               }else{
                    yearArray = Value
               }
            }else if(yearDate.length==1){
                var monthDate = yearDate[0].split('月');
                if(monthDate.length==2){
                    var dayDate =  monthDate[1].split('日');
                    if(dayDate.length==2&&Number(monthDate[0])%1===0&&Number(monthDate[0])>=1&&Number(monthDate[0])<=12&&Number(dayDate[0])%1===0&&Number(dayDate[0])>=1&&Number(dayDate[0])<=31){
                        yearArray = [Number(nowYear),Number(monthDate[0]),Number(dayDate[0]),"m月d日"]  //10月10日  ->  2019 10  10
                   }else{
                        yearArray = Value
                   }
                }else{
                    yearArray = Value
                }
            }else{
                yearArray = Value
            }
            return yearArray
        }else{
            var newDateValue = Value.split('/');
            var dataArray ;      // /或者-分割的情况
            if(newDateValue.length==3){   //2019/10/10  ->  2019  10 10 
                if(Number(newDateValue[0])%1===0&&Number(newDateValue[0])>=1900&&Number(newDateValue[0])<=9999&&Number(newDateValue[1])%1===0&&Number(newDateValue[1])>=1&&Number(newDateValue[1])<=12&&Number(newDateValue[2])%1===0&&Number(newDateValue[2])>=1&&Number(newDateValue[2])<=31){
                    dataArray = [Number(newDateValue[0]),Number(newDateValue[1]),Number(newDateValue[2]),"yyyy/m/d",'/']
                }else{
                    dataArray =  Value
                }
            }else if(newDateValue.length==2){
                var date1 = newDateValue[0].split("-");  
                var date2 = newDateValue[1].split('-');
                if(date1.length==2&&date2.length==1&&Number(date1[0])%1===0&&Number(date1[0])>=1900&&Number(date1[0])<=9999&&Number(date1[1])%1==0&&Number(date1[1])>=1&&Number(date1[1])<=12&&Number(date2[0])%1===0&&Number(date2[0])>=1&&Number(date2[0])<=31){
                    dataArray = [Number(date1[0]),Number(date1[1]),Number(date2[0]),"yyyy/m/d",'/']//2019-10/10  ->2019 10 10 
                }else if(date1.length==1&&date2.length==2&&Number(date1[0])%1===0&&Number(date1[0])>=1900&&Number(date1[0])<=9999&&Number(date2[0])%1==0&&Number(date2[0])>=1&&Number(date2[0])<=12&&Number(date2[1])%1===0&&Number(date2[1])>=1&&Number(date2[1])<=31){
                    dataArray = [Number(date1[0]),Number(date2[0]),Number(date2[1]),"yyyy/m/d",'/'] //2019/10-10  ->2019 10 10
                }else if(date1.length==1&&date2.length==1&&Number(newDateValue[0])%1===0&&Number(newDateValue[0])>=1900&&Number(newDateValue[0])<=9999&&Number(newDateValue[1])%1===0&&Number(newDateValue[1])>=1&&Number(newDateValue[1])<=12){
                    dataArray = [Number(newDateValue[0]),Number(newDateValue[1]),1,"yyyy/m",'/']//2019/10  ->2019 10 1
                }else if(date1.length==1&&date2.length==1&&Number(newDateValue[0])%1===0&&Number(newDateValue[0])>=1&&Number(newDateValue[0])<=12&&Number(newDateValue[1])%1===0&&Number(newDateValue[1])>=1&&Number(newDateValue[1])<=31){
                    dataArray = [Number(nowYear),Number(newDateValue[0]),Number(newDateValue[1]),"m/d",'/'] //10/10  ->2019 10 10 
                }else{
                    dataArray = Value
                }
            }else if(newDateValue.length==1){
                var date =  newDateValue[0].split("-");       //只有-分割的情况
                if(date.length==3&&Number(date[0])%1===0&&Number(date[0])>=1900&&Number(date[0])<=9999&&Number(date[1])%1==0&&Number(date[1])>=1&&Number(date[1])<=12&&Number(date[2])%1===0&&Number(date[2])>=1&&Number(date[2])<=31){
                    dataArray = [Number(date[0]),Number(date[1]),Number(date[2]),"yyyy-m-d",'-']//2019-10-10  ->2019 10 10 
                }else if(date.length==2&&Number(date[0])%1===0&&Number(date[0])>=1900&&Number(date[0])<=9999&&Number(date[1])%1==0&&Number(date[1])>=1&&Number(date[1])<=12){
                    dataArray = [Number(date[0]),Number(date[1]),1,"yyyy-m",'-']//2019-10  ->2019 10 1 
                }else if(date.length==2&&Number(date[0])%1==0&&Number(date[0])>=1&&Number(date[0])<=12&&Number(date[1])%1===0&&Number(date[1])>=1&&Number(date[1])<=31){
                    dataArray = [Number(nowYear),Number(date[0]),Number(date[1]),"m-d",'-'] //10-10  ->2019 10 10 
                }
                // else if(newDateValue.length==1&&Number(newDateValue[0])%1===0&&Number(newDateValue[0])>=1&&Number(newDateValue[0])<=31){
                //     dataArray = [1900,1,Number(date[0])]//10  ->2019 1 10 
                // }
                else {
                    dataArray = Value
                }
            }else{
                dataArray = Value
            }
            return dataArray
        }
    }
    
}

/**
 * @api {null} WB.setCellType(R1,C1,R2,C2,obj,Index) setCellType
 * @apiName setCellType
 * @apiGroup Function
 * @apiDescription Set the cellType of the cells in the range
 * For obj parameter format, please refer to cellTypeContent method(var obj = WB.cellTypeContent(4,["apple","banana"]))
 * - WB.setCellType(0,0,0,0,obj)    
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {int} obj  
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setCellType = function(R1,C1,R2,C2,obj,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable;
    for(var i = R1;i<=R2;i++){         
        for(var j = C1;j<=C2;j++){      
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                activeSheet.data.dataTable[i][j].style.cellType = obj ; 
            };
        }       
    };
}

/**
 * @api {null} WB.getCellType(R1,C1,R2,C2,Index) getCellType
 * @apiName getCellType
 * @apiGroup Function
 * @apiDescription Get the celltype of a range of cells (no parameter defaults to the celltype of the currently selected range)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.getCellType(0,0)
 */
Workbook.prototype.getCellType = function(R1,C1,R2,C2,Index){     
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
    data = activeSheet.data.dataTable,cellTypeArr = [],rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2,isSpan = this.isSpanCell(r1,c1,r2,c2);
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            var cellType;
            if(data[i]&&data[i][j]&&data[i][j].style&&data[i][j].style.cellType){
                cellType = data[i][j].style.cellType
                cellTypeArr.push({"r":i,"c":j,"cellType":cellType})
            }
        }
    }
    if(((r1==r2&&c1==c2)||isSpan)&&cellTypeArr.length==1){
        return cellTypeArr[0].cellType
    }else if((r1!=r2||c1!=c2)&&!isSpan){
        return cellTypeArr
    }
}

/**
 * @api {null} WB.cellTypeContent(n,list,i) cellTypeContent
 * @apiName cellTypeContent
 * @apiGroup Function
 * @apiDescription Set the cell type to return an object format
 * @apiParam {int} n  
 * - 0：No button
 * - 1：The button type is three dots
 * - 2：Button type is subscript arrow
 * - 3：checkBox(The value of this cell is not drawn and cannot be edited)
 * - 4: select(Drop-down box)
 * @apiParam {Array} [list]
 * @apiParam {Int} [i]  Special item (reference onopenpopup)
 * @apiExample {javascript} demo:
    WB.cellTypeContent(4,["a","b","c"]) 
 */
Workbook.prototype.cellTypeContent = function(n,list,i){
    return {"name":n,"list":list,"notSetIndex":i}
}

/**
 * @api {null} WB.setTableSize(width,height) setTableSize
 * @apiName setTableSize
 * @apiGroup Function
 * @apiDescription Set the width and height of the canvas and container
 * @apiParam {Number} width  width(width>0)
 * @apiParam {Number} height height(height>0)
 * @apiExample {javascript} demo:
    WB.setTableSize(500,300) 
 */
Workbook.prototype.setTableSize = function(width,height){
    this.workbook.width = width || this.workbook.width;
    this.workbook.height = height || this.workbook.height;
    this.container.style.height = this.workbook.height+'px'
    this.container.style.width = this.workbook.width+'px'
    this.recalculatePercentage()
}

/**
 * @api {null} /null drawPrint
 * @apiName drawPrint
 * @apiGroup Function
 * @apiDescription print
 * - Receive a callback and return an array (base64)
 * @apiParam {Function} callback
 * @apiParam {Int} [Index=Current_table_index]  table index 
 * @apiExample {javascript} demo:
    WB.drawPrint(function(arr){
        console.log(arr)
    })
 */
Workbook.prototype.drawPrint = function(callback,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var oShowTabs = this.workbook.showTabs,oShowHScrollBar = this.workbook.showHScrollBar,
    oShowVScrollBar = this.workbook.showVScrollBar,oIndex = this.workbook.activeSheet,
    oWidth = this.workbook.width,oHeight = this.workbook.height;
    var totolW = [],totolH = 0,sheetList = this.workbook.sheetList,dataArr = [],sheetArr = [];
    this.workbook.showTabs = 0;
    this.workbook.showHScrollBar = 0;
    this.workbook.showVScrollBar = 0;

    for(var i = 0;i<sheetList.length;i++){
        var activeSheet = this.getActiveSheet(i),rows = activeSheet.rows,cols = activeSheet.columns,
        oShowRowHeading = activeSheet.rowHeaderData.showRowHeading,oShowColHeading = activeSheet.colHeaderData.showColHeading,
        oIsSelectionHideBorder =  activeSheet.isSelectionHideBorder,oFixedRow = activeSheet.fixed.fixedRow,
        oFixedRows = activeSheet.fixed.fixedRows,oFixedCol = activeSheet.fixed.fixedCol,oFixedCols = activeSheet.fixed.fixedCols,
        oCoordY = activeSheet.scrollOption.coordY,oTotolY = activeSheet.scrollOption.totolY,
        oCoordX = activeSheet.scrollOption.coordX,oTotolX = activeSheet.scrollOption.totolX,
        oOffsetX = activeSheet.scrollOption.offsetX,oOffsetY = activeSheet.scrollOption.offsetY,
        oShowGridLines = activeSheet.showGridLines;
        var colHeader = activeSheet.colHeaderData,rowHeader = activeSheet.rowHeaderData;
        var numRowH = (activeSheet.printSetting.printHeadings)?colHeader.defaultH:0,
        numColW = (activeSheet.printSetting.printHeadings)?rowHeader.defaultW:0;
        var maxRowAndCol = this.getMaxRowAndCol(i)  //获取最大有单元格的行与列 以及该总行高与总列宽
        totolH+=maxRowAndCol.totolHeight+numRowH;
        totolW.push(maxRowAndCol.totolWidth+numColW);
        var gridLinesColor = activeSheet.gridLinesColor;
        var totolWidth = maxRowAndCol.totolWidth+numColW;
        var totolHeight = maxRowAndCol.totolHeight+numRowH;
        var maxRow = maxRowAndCol.maxRow,maxCol = maxRowAndCol.maxCol;
        //重置sheet参数
        this.workbook.width = totolWidth;
        this.workbook.height = totolHeight;
        activeSheet.rowHeaderData.showRowHeading = (activeSheet.printSetting.printHeadings)?true:false;
        activeSheet.colHeaderData.showColHeading = (activeSheet.printSetting.printHeadings)?true:false;
        activeSheet.showGridLines = (activeSheet.printSetting.printGridLine)?true:false;
        activeSheet.isSelectionHideBorder = true;
        activeSheet.scrollOption.coordX = 0;
        activeSheet.scrollOption.totolX = 0;
        activeSheet.scrollOption.coordY = 0;
        activeSheet.scrollOption.totolY = 0;
        activeSheet.scrollOption.offsetX = 0;
        activeSheet.scrollOption.offsetY = 0;
        this.removeFixed(i)
        this.sheetIndex(i)
        this.startPaint(true);
        this.canvas.width = this.width*this.ratio;      
        this.canvas.height = this.height*this.ratio;
        this.ctx.scale(this.ratio,this.ratio);
        if(activeSheet.printSetting.printGridLine){
            this.drawLine(0,0,maxRow,maxCol,i)
            this.drawTopLine(i)
            this.drawLeftLine(i)
        }
        this.drawSpans(0,0,maxRow,maxCol,i)
        this.drawCellContent(0,0,maxRow,maxCol,i)
        if(activeSheet.printSetting.printHeadings){
            this.drawRowHeader(0,maxRow,i)
            this.drawRowArrow(i)
        }
        if(activeSheet.printSetting.printHeadings){
            this.drawColHeader(0,maxCol,i)
        }
        if(activeSheet.printSetting.printHeadings&&activeSheet.printSetting.printHeadings&&(rows.length>1||cols.length>1)){
            this.drawTriangle(activeSheet.colHeaderData.defaultDataNode.style.fillColor,gridLinesColor)
        }

        var dataURL = this.canvas.toDataURL();
        dataArr.push({'index':i,'h':totolHeight,'dataUrl':dataURL})
        sheetArr.push({'index':i,'h':totolHeight,'w':totolWidth})
        activeSheet.rowHeaderData.showRowHeading = oShowRowHeading;
        activeSheet.colHeaderData.showColHeading = oShowColHeading;
        activeSheet.showGridLines = oShowGridLines;
        activeSheet.isSelectionHideBorder = oIsSelectionHideBorder;
        this.fixedRowAndCol(oFixedRows,oFixedCols,oFixedRow,oFixedCol,i)
        activeSheet.scrollOption.coordY = oCoordY;
        activeSheet.scrollOption.totolY = oTotolY;
        activeSheet.scrollOption.coordX = oCoordX;
        activeSheet.scrollOption.totolX = oTotolX;
        activeSheet.scrollOption.offsetX = oOffsetX;
        activeSheet.scrollOption.offsetY = oOffsetY;
    }
    totolW = Math.max.apply(null,totolW)
    //恢复原来设置的参数
    this.workbook.width = oWidth;
    this.workbook.height = oHeight;
    this.workbook.showTabs = oShowTabs;
    this.workbook.showHScrollBar = oShowHScrollBar;
    this.workbook.showVScrollBar = oShowVScrollBar;
    this.sheetIndex(oIndex)
    this.startPaint(true)
    
    var canvas = document.createElement('canvas');
    var ctx = canvas.getContext('2d')
    canvas.width = totolW*this.ratio;      
    canvas.height = totolH*this.ratio;
    var imgArr = [],_this = this
    for(var t = 0;t<dataArr.length;t++){
        var img = new Image();
        img.src = dataArr[t].dataUrl;
        imgArr.push({'img':img,'h':dataArr[t].h})
    }
    doDraw(0,this.print)
    var y = 0;
    function doDraw(idx,cb){
        if(imgArr[idx]&&idx<imgArr.length){
            imgArr[idx]['img'].onload = function(){
                ctx.drawImage(imgArr[idx]['img'],0,Math.floor(y*_this.ratio))
                y += dataArr[idx]['h']
                idx++;
                doDraw(idx,cb)
                return true
            }
        }else{
            cb.call(_this,canvas.toDataURL(),sheetArr,callback,index)
        }
    }
}

//获取表最大的打印行列
Workbook.prototype.getMaxRowAndCol = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),rows = activeSheet.rows,
    cols = activeSheet.columns,spans = activeSheet.spans,data = activeSheet.data.dataTable,
    spansRowArr = [],spansColArr = [],maxRow = 0,maxCol = 0;
    spans.forEach(function(item){
        spansRowArr.push(item.row+item.rowCount-1);
        spansColArr.push(item.col+item.colCount-1);
    })

    if(spansRowArr.length>=1){
        maxRow = Math.max.apply(null,spansRowArr);
    }
    if(spansColArr.length>=1){
        maxCol = Math.max.apply(null,spansColArr)
    }

    var rowkey = Object.keys(data),lastRow = 0,lastCol = 0;
    if(rowkey.length>=1){
        lastRow = Math.max.apply(null,rowkey)
    }

    maxRow = (maxRow>=lastRow)?maxRow:lastRow;
    
    var maxColArr = [];
    for(var t = 0;t<=maxRow;t++){
        if(data[t]){
            var colkey = Object.keys(data[t]);
            if(colkey.length>=1){
                lastCol = Math.max.apply(null,colkey)
            }
            maxColArr.push(lastCol)
        }
    };
    if(maxColArr.length>=1){
        lastCol = Math.max.apply(null,maxColArr)   
    }

    maxCol = (maxCol>=lastCol)?maxCol:lastCol;

    var totolHeight = 0,totolWidth = 0;
    for(var i = 0;i<=maxRow;i++){
        if(rows[i]){
            totolHeight+=rows[i].size; 
        }
    }

    for(var j = 0;j<=maxCol;j++){
        if(cols[j]){
            totolWidth+=cols[j].size;
        }
    }

    return {'maxRow':maxRow,'maxCol':maxCol,'totolHeight':totolHeight,'totolWidth':totolWidth}
}

Workbook.prototype.print = function(allSheetUrl,sheetArr,callback,Index){
    var print = (this.workbook.print>=1&&this.workbook.print<=3)?this.workbook.print:1,
    printOnSamePaper = this.workbook.printOnSamePaper;
    var marginCopies = parseInt((this.workbook.marginCopies)/25.4*this.dpi),_this = this;
    var canvas = document.createElement('canvas');
    var ctx = canvas.getContext('2d')
    var img = new Image(),pageArr = [];
    img.src = allSheetUrl;
    //切片分页 画页码
    img.onload = function(){
        var totolH = 0;
        for(var i = 0;i<sheetArr.length;i++){
            var w = sheetArr[i].w,h = sheetArr[i].h;
            totolH += h;
            if(((print==1||print==3)&&sheetArr[i].index==Index)||print==2){   
                var index = i,activeSheet = _this.getActiveSheet(index),printSetting = activeSheet.printSetting;
                var printDirection = parseInt(printSetting.printDirection);
                printDirection = (printDirection>=1&&printDirection<=2)?printDirection:1;
                var orientation = parseInt(printSetting.orientation);
                orientation = (orientation>=1&&orientation<=2)?orientation:1;
                var paperInfo = _this.getPaperSize(index),printW = paperInfo.paperW,printH = paperInfo.paperH,
                marginRight = paperInfo.marginRight,marginLeft = paperInfo.marginLeft,marginTop = paperInfo.marginTop,
                marginBottom = paperInfo.marginBottom;
                var isshowfooterpageinfo = parseInt(printSetting.isshowfooterpageinfo);
                var footpagestyle =  printSetting.footpagestyle||"第 &p 页，共 &P 页"
                isshowfooterpageinfo = (isshowfooterpageinfo>=0&&isshowfooterpageinfo<=6)?isshowfooterpageinfo:0;
                var canvasW = (orientation==1)?printW:printH,canvasH = (orientation==1)?printH:printW;
                var colHeader = activeSheet.colHeaderData,rowHeader = activeSheet.rowHeaderData;
                var numRowH = (printSetting.printHeadings)?colHeader.defaultH:0,
                numColW = (printSetting.printHeadings)?rowHeader.defaultW:0;
                canvas.width = canvasW
                canvas.height = canvasH
                var printAreaPosition = _this.getPrintAreaPosition(index),startX = printAreaPosition.startX+numColW,
                startY = printAreaPosition.startY+numRowH,endX = printAreaPosition.endX+numColW,endY = printAreaPosition.endY+numRowH,
                width = printAreaPosition.width+numColW,height = printAreaPosition.height+numRowH;
                var contentW = (print==1||print==2)?w*_this.ratio:width*_this.ratio;
                var contentH = (print==1||print==2)?h*_this.ratio:height*_this.ratio;
                var endH = (print==1||print==2)?Math.floor(totolH*_this.ratio):Math.floor((totolH-h+endY)*_this.ratio);
                var endW = (print==1||print==2)?Math.floor(w*_this.ratio):Math.floor(endX*_this.ratio);
                var startH = (print==1||print==2)?0:Math.floor(startY*_this.ratio);
                var startW = (print==1||print==2)?0:Math.floor(startX*_this.ratio);
                var cutNum = (print==1||print==2)?1:0;
                var vPageNum = Math.ceil(contentH/canvasH),hPageNum = Math.ceil(contentW/canvasW);
                vPageNum = Math.ceil((vPageNum*(marginTop+marginBottom+numRowH)+contentH-numRowH)/canvasH);
                hPageNum = (print==2&&printOnSamePaper)?1:Math.ceil((hPageNum*(marginLeft+marginRight+numColW)+contentW-numColW)/canvasW);
                var page = 0,row = 0,col = 0,marginLR = 0,cutL = 0,sx = 0,marginTB = 0,cutT = 0,sy = 0;
                var outEnd = (printDirection==1)?hPageNum:(printDirection==2)?vPageNum:hPageNum;
                var innerEnd = (printDirection==1)?vPageNum:(printDirection==2)?hPageNum:vPageNum;
                for(var a = 0;a<outEnd;a++){
                    if(printDirection==1){
                        marginLR = (a>0)?(marginLeft+marginRight):0;
                        cutL = (a>1||(a>0&&print==3))?Math.floor(numColW*_this.ratio):0;
                        sx = startW+a*(canvasW-marginLR)-(a-cutNum)*cutL
                    }else if(printDirection==2){
                        marginTB = (a>0)?(marginTop+marginBottom):0;
                        cutT = (a>1||(a>0&&print==3))?Math.floor(numRowH*_this.ratio):0;
                        sy = Math.floor((totolH-h)*_this.ratio)+startH+a*(canvasH-marginTB)-(a-cutNum)*cutT
                    }
                    for(var b = 0;b<innerEnd;b++){ 
                        page += 1;
                        canvas.width = canvasW
                        canvas.height = canvasH
                        if(printDirection==1){
                            marginTB = (b>0)?(marginTop+marginBottom):0;
                            cutT = (b>1||(print==3&&b>0))?Math.floor(numRowH*_this.ratio):0;
                            sy = Math.floor((totolH-h)*_this.ratio)+startH+b*(canvasH-marginTB)-(b-cutNum)*cutT;
                            row = b;
                            col = a;
                        }else if(printDirection==2){
                            marginLR = (b>0)?(marginLeft+marginRight):0;
                            cutL = (b>1||(print==3&&b>0))?numColW*_this.ratio:0;
                            sx = startW+b*(canvasW-marginLR)-(b-cutNum)*cutL
                            row = a;
                            col = b; 
                        }
                        var cutNumRowH = (row>0||print==3)?Math.floor(numRowH*_this.ratio):0;
                        var cutNumColW = (col>0||print==3)?Math.floor(numColW*_this.ratio):0;
                        var sWidth = (canvasW-marginRight-marginLeft-cutNumColW);
                        sWidth = (sWidth>=endW-sx)?(endW-sx):sWidth;
                        var sHeight = (canvasH-marginBottom-marginTop-cutNumRowH);
                        sHeight = (sHeight>=endH-sy)?(endH-sy):sHeight;
                        var dx = marginLeft+cutNumColW,dy = marginTop+cutNumRowH,dWidth = sWidth,dHeight = sHeight;
                        ctx.drawImage(img,sx,sy,sWidth,sHeight,dx,dy,dWidth,dHeight)
                        //绘制每一个页的行列标(除了顶部的所有页绘制列标 除了最左的所有页绘制行标)
                        if(printSetting.printHeadings&&!(row==0&&col==0)||print==3){
                            if(row>0&&col==0&&(print==1||print==2)){
                                ctx.drawImage(img,sx,Math.floor((totolH-h)*_this.ratio),sWidth,Math.floor(numRowH*_this.ratio),marginLeft,marginTop,sWidth,Math.floor(numRowH*_this.ratio))
                            };
                            if(col>0&&row==0&&(print==1||print==2)){
                                ctx.drawImage(img,0,sy,Math.floor(numColW*_this.ratio),sHeight,marginLeft,marginTop,Math.floor(numColW*_this.ratio),sHeight)
                            };
                            if(row>0&&col>0||print==3){
                                var cutW = Math.floor(numColW*_this.ratio)
                                var cutH = Math.floor(numRowH*_this.ratio)
                                ctx.drawImage(img,sx,Math.floor((totolH-h)*_this.ratio),sWidth,Math.floor(numRowH*_this.ratio),marginLeft+cutW,marginTop,sWidth,Math.floor(numRowH*_this.ratio))
                                ctx.drawImage(img,0,sy,Math.floor(numColW*_this.ratio),sHeight,marginLeft,marginTop+cutH,Math.floor(numColW*_this.ratio),sHeight)
                                ctx.drawImage(img,0,Math.floor((totolH-h)*_this.ratio),Math.floor(numColW*_this.ratio),Math.floor(numRowH*_this.ratio),marginLeft,marginTop,Math.floor(numColW*_this.ratio),Math.floor(numRowH*_this.ratio))
                            }
                        }
                        var base64 = canvas.toDataURL();
                        pageArr.push({'page':page,'base64':base64,'sx':sx,'sWidth':sWidth,'sy':Math.floor(sy-(totolH-h)*_this.ratio),'sHeight':sHeight,'row':row,'col':col,'index':i,'canvasW':canvasW,'canvasH':canvasH,'isshowfooterpageinfo':isshowfooterpageinfo,
                        'footpagestyle':footpagestyle,'marginT':marginTop,'marginB':marginBottom,'marginL':marginLeft,'marginR':marginRight})
                    }
                }
            }
        };
        pageArr = _this.deleteEmptyPage(pageArr);
        if(pageArr.length>=1){
            var newPageArr = [],newH = 0,newW = 0,marginC = 0;
            for(var t = 0;t<pageArr.length;t++){
                var image = new Image()
                image.src = pageArr[t].base64;
                pageArr[t]['image'] = image
                newH+=pageArr[t].sHeight;   //当打印在同一张纸上
                newW = pageArr[0]['canvasW']-pageArr[0]['marginL']-pageArr[0]['marginR'];
                marginC+=marginCopies;
            }
        
            if(printOnSamePaper&&print==2){
                drawSamePage(0,function(arr){
                    if(typeof(callback)=='function'){
                        callback(arr)
                    }
                })
                var onSamePaperArr = [];
                var firstPaperInfo = _this.getPaperSize(0),fMarginRight = firstPaperInfo.marginRight,fMarginLeft = firstPaperInfo.marginLeft,
                fMarginTop = firstPaperInfo.marginTop,fMarginBottom = firstPaperInfo.marginBottom,newDy = 0;
                newH+=+marginC-marginCopies;
                canvas.width = newW;
                canvas.height = newH;
                function drawSamePage(idx,cb){
                    if(pageArr[idx]&&idx<pageArr.length){
                        pageArr[idx]['image'].onload = function(){
                            ctx.drawImage(pageArr[idx]['image'],pageArr[idx]['marginL'],pageArr[idx]['marginT'],pageArr[idx]['sWidth'],pageArr[idx]['sHeight'],0,newDy,pageArr[idx]['sWidth'],pageArr[idx]['sHeight'])
                            newDy+=pageArr[idx]['sHeight']+marginCopies;
                            idx++;
                            drawSamePage(idx,cb)
                            return true
                        }
                    }else{
                        var samePaper = new Image();
                        samePaper.src = canvas.toDataURL();
                        samePaper.onload = function(){
                            var vNum = Math.ceil(newH/pageArr[0]['canvasH']);
                            vNum = Math.ceil((vNum*(fMarginTop+fMarginBottom)+newH)/pageArr[0]['canvasH']);
                            var sameSy = 0,cutLR = fMarginLeft+fMarginRight,cutTB = fMarginBottom+fMarginTop,numPage = 0;
                            for(var a = 0;a<vNum;a++){
                                numPage+=1;
                                canvas.width = pageArr[0]['canvasW'];
                                canvas.height = pageArr[0]['canvasH'];
                                ctx.drawImage(samePaper,0,sameSy,pageArr[0]['canvasW']-cutLR,pageArr[0]['canvasH']-cutTB,fMarginLeft,fMarginTop,pageArr[0]['canvasW']-cutLR,pageArr[0]['canvasH']-cutTB);
                                sameSy+=pageArr[0]['canvasH'];
                                var numImage = new Image()
                                numImage.src = canvas.toDataURL();
                                onSamePaperArr.push({'base64':canvas.toDataURL(),'image':numImage,'page':numPage,'canvasW':pageArr[0]['canvasW'],'canvasH':pageArr[0]['canvasH'],'isshowfooterpageinfo':pageArr[0]['isshowfooterpageinfo'],
                                'footpagestyle':pageArr[0]['footpagestyle']})
                            }
                            drawPageNum(0,onSamePaperArr,function(arr){
                                cb(arr)
                            });
                        }
                    }
                }
            }else{
                drawPageNum(0,pageArr,function(arr){
                    if(typeof(callback)=='function'){
                        callback(arr)
                    }
                });
            }
    
            function drawPageNum(idx,numArr,cb){   //画页码
                if(numArr[idx]&&idx<numArr.length){
                    numArr[idx]['image'].onload = function(){
                        canvas.width = numArr[idx]['canvasW'];
                        canvas.height = numArr[idx]['canvasH'];
                        ctx.drawImage(numArr[idx]['image'],0,0)
                        if(numArr[idx]['isshowfooterpageinfo']!=0){
                            ctx.font = '14px 楷体'
                            var pageNumPosition = _this.getPageNumPosition(numArr[idx]['canvasW'],numArr[idx]['canvasH'],numArr[idx]['isshowfooterpageinfo']);
                            ctx.textAlign = pageNumPosition.align;
                            ctx.textBaseline="middle";
                            ctx.fillStyle = '#000';
                            var text = numArr[idx]['footpagestyle'].replace('&p',numArr[idx].page).replace('&P',numArr.length)
                            ctx.fillText(text,pageNumPosition.x,pageNumPosition.y)
                        }
                        newPageArr.push({'base64':canvas.toDataURL()})
                        idx++;
                        drawPageNum(idx,numArr,cb)
                        return true
                    }
                }else{
                    cb(newPageArr)
                }
            }
        }
       
    }
}

//获取打印区域的位置信息
Workbook.prototype.getPrintAreaPosition = function(Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),printSetting = activeSheet.printSetting;
    var rows = activeSheet.rows,cols = activeSheet.columns;
    var printR1 = parseInt(printSetting.printArea.r1)>=0?parseInt(printSetting.printArea.r1):activeSheet.rangeRow1;
    var printC1 = parseInt(printSetting.printArea.c1)>=0?parseInt(printSetting.printArea.c1):activeSheet.rangeCol1;
    var printR2 = parseInt(printSetting.printArea.r2)>=0?parseInt(printSetting.printArea.r2):activeSheet.rangeRow2;
    var printC2 = parseInt(printSetting.printArea.c2)>=0?parseInt(printSetting.printArea.c2):activeSheet.rangeCol2;
    var maxRowAndCol = this.getMaxRowAndCol(index)  //获取最大有单元格的行与列 以及该总行高与总列宽
    var maxRow = maxRowAndCol.maxRow,maxCol = maxRowAndCol.maxCol;
    if(printR1>maxRow){
        printR1 = maxRow;
    };
    if(printR2>maxRow){
        printR2 = maxRow
    };
    if(printC1>maxCol){
        printC1 = maxCol
    };
    if(printC2>maxCol){
        printC2 = maxCol;
    }
    var startX = 0,startY = 0,endX = 0,endY = 0,width = 0,height = 0;
    for(var i = 0;i<printR1;i++){
        if(rows[i]){
            startY += rows[i].size;
        }
    };
    for(var i = 0;i<printC1;i++){
        if(cols[i]){
            startX += cols[i].size
        }
    };
    for(var i = 0;i<=printR2;i++){
        if(rows[i]){
            endY += rows[i].size;
        }
    };
    for(var i = 0;i<=printC2;i++){
        if(cols[i]){
            endX += cols[i].size;
        }
    }
    width = endX - startX;
    height = endY - startY;
    
    return {'startX':startX,'startY':startY,'endX':endX,'endY':endY,'width':width,'height':height}
}

//删除空页
Workbook.prototype.deleteEmptyPage = function(pageArr){
   if(pageArr){
       for(var i = 0;i<pageArr.length;i++){
            var index = (pageArr[i].index>=0)?pageArr[i].index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index),
            rows = activeSheet.rows,cols = activeSheet.columns,data = activeSheet.data.dataTable,
            spans = activeSheet.spans,printSetting = activeSheet.printSetting,numColW = 0,numRowH = 0;
            var printDirection = parseInt(printSetting.printDirection);
            printDirection = (printDirection>=1&&printDirection<=2)?printDirection:1;
            if(printSetting.printHeadings){
                numColW = activeSheet.rowHeaderData.defaultW;
                numRowH = activeSheet.colHeaderData.defaultH;
            }
            var sy = pageArr[i].sy,sHeight = pageArr[i].sHeight,sx = pageArr[i].sx,sWidth = pageArr[i].sWidth,
            ey = sy+sHeight,ex = sx+sWidth;
            var r1 = 0,c1 = 0,r2 = 0, c2 = 0,sxSumX = numColW,sySumY = numRowH,exSumX = numColW,eySumY = numRowH;
            for(var k = 0;k<rows.length;k++){
               if(rows[k])  sySumY+=rows[k].size;
               if(sy<=sySumY*this.ratio){
                   r1 = k;
                   break;
               }
            }
            for(var k = 0;k<rows.length;k++){
                if(rows[k])  eySumY+=rows[k].size;
                if(ey<=eySumY*this.ratio){
                    r2 = k;
                    break;
                }
            }
            for(var j = 0;j<cols.length;j++){
                if(cols[j])  sxSumX+=cols[j].size;
                if(sx<=sxSumX*this.ratio){
                    c1 = j;
                    break;
                }
            }
            for(var j = 0;j<cols.length;j++){
                if(cols[j])  exSumX+=cols[j].size;
                if(ex<=exSumX*this.ratio){
                    c2 = j;
                    break;
                }
            }

            for(var n = r1;n<=r2;n++){
                for(var v = c1;v<=c2;v++){
                    if(data[n]&&data[n][v]){
                        pageArr[i].ety = false
                    }
                }
            };
            spans.forEach(function(item){
                var row1 = item.row,col1 = item.col,row2 = item.row+item.rowCount-1,col2 = item.col+item.colCount-1;
                if(row2>=r1&&row1<=r2&&col2>=c1&&col1<=c2){
                    pageArr[i].ety = false
                }
            }); 
            // delete pageArr[i].sy;
            // delete pageArr[i].sx;
            // delete pageArr[i].sHeight;
            // delete pageArr[i].sWidth;
        }
        var diviR,diviC
        for(var i = pageArr.length-1;i>=0;i--){
            if(pageArr[i].ety!==false){
                pageArr.splice(i,1)
            }else if(pageArr[i].ety===false&&printDirection==1){
                diviR = pageArr[i].row;
                break;
            }else if(pageArr[i].ety===false&&printDirection==2){
                diviC = pageArr[i].col;
                break;
            }  
        };
    
        for(var j = pageArr.length-1;j>=0;j--){
            if(pageArr[j].row>diviR&&printDirection==1){
                if(pageArr[j].ety!==false){
                    pageArr.splice(j,1)
                }
            }else if(pageArr[j].col>diviC&&printDirection==2){
                if(pageArr[j].ety!==false){
                    pageArr.splice(j,1)
                }
            }
        }
            
        //重置page
        for(var k = 0;k<pageArr.length;k++){
            pageArr[k].page = k+1;
            // delete pageArr[k].row;
            // delete pageArr[k].col;
            // delete pageArr[k].ety;
            // delete pageArr[k].index
        }  
        
        return pageArr  
    }
}  

//获取页码位置信息 
Workbook.prototype.getPageNumPosition = function(canvasW,canvasH,isshowfooterpageinfo){
    //"0", "1", "2", "3", "4", "5", "6"对应:"不显示", "页脚居中", "页脚居左", "页脚居右", "页眉居中", "页眉居左", "页眉居右
    var x = 0,y = 0,align = 'center';
    switch(isshowfooterpageinfo){
        case 1:
            x = canvasW/2;y = canvasH-10;align = 'center';
            break;
        case 2:
            x = 0;y = canvasH -10;align = 'left';
            break;
        case 3:
            x = canvasW;y = canvasH - 10;align = 'right';
            break;
        case 4:
            x = canvasW/2;y = 10;align = 'center';
            break;
        case 5:
            x = 0;y = 10;align = 'left';
            break;
        case 6:
            x = canvasW;y = 10;align = 'right'
    }
    return {'x':x,'y':y,'align':align}
}

/**
 * @api {null} /null setPrint
 * @apiName setPrint
 * @apiGroup Function
 * @apiDescription Set to print, the default value is used for parameters that are not passed
 * - marginTop 
 * - marginBottom 
 * - marginLeft 
 * - marginRight 
 * - paper (Paper type)
 * - printHeadings (Whether to print row and column labels)
 * - printGridLine (Whether to print grid lines)
 * - orientation (Paper orientation)
 * - printDirection (Print direction:)(1:N shape 2：Z shape) 
 * - marginCopies (The interval between two data (tables) on the same paper)
 * - printer (printer)
 * - isshowfooterpageinfo (Display page number information parameters "0", "1", "2", "3", "4", "5", "6" correspond to: "not displayed", "footer centered", "footer left" "Header right", "Header centered", "Header left", "Header right")
 * - footpagestyle (Page number format)
 * - startR (Start row of print area)
 * - endR (end row of print area)
 * - startC (Start column of print area)
 * - endC (end column of print area)
 * - print (1 Print the current worksheet (default) 2 Print the workbook 3 Print the currently selected area of the current table)
 * - printOnSamePaper (Only effective when print is 2, whether multiple sheets are placed on the same sheet, default false)
 * @apiParam {Object} obj  
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setPrint = function(obj,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var printSetting = activeSheet.printSetting;
    if(obj){
        printSetting.marginTop = (obj.marginTop>=0)?obj.marginTop:printSetting.marginTop;
        printSetting.marginBottom = (obj.marginBottom>=0)?obj.marginBottom:printSetting.marginBottom;
        printSetting.marginLeft = (obj.marginLeft>=0)?obj.marginLeft:printSetting.marginLeft;
        printSetting.marginRight = (obj.marginRight>=0)?obj.marginRight:printSetting.marginRight;
        printSetting.paper = obj.paper||printSetting.paper;
        printSetting.printHeadings = (obj.printHeadings===undefined)?printSetting.printHeadings:(obj.printHeadings)?true:false;
        printSetting.printOnSamePaper = (obj.printOnSamePaper===undefined)?printSetting.printOnSamePaper:(obj.printOnSamePaper)?true:false;
        printSetting.printGridLine = (obj.printGridLine===undefined)?printSetting.printGridLine:(obj.printGridLine)?true:false;
        printSetting.orientation = obj.orientation||printSetting.orientation;
        printSetting.printDirection = obj.printDirection||printSetting.printDirection;
        this.workbook.marginCopies = (obj.marginCopies>=0)?obj.marginCopies:this.workbook.marginCopies;
        printSetting.printer = obj.printer||printSetting.printer;
        printSetting.isshowfooterpageinfo = (obj.isshowfooterpageinfo>=0&&obj.isshowfooterpageinfo<=6)?obj.isshowfooterpageinfo:printSetting.isshowfooterpageinfo;
        printSetting.footpagestyle = obj.footpagestyle||printSetting.footpagestyle;
        printSetting.printArea.r1 = (obj.startR>=0)?obj.startR:printSetting.printArea.r1;
        printSetting.printArea.r2 = (obj.endR>=0)?obj.endR:printSetting.printArea.r2;
        printSetting.printArea.c1 = (obj.startC>=0)?obj.startC:printSetting.printArea.c1;
        printSetting.printArea.c2 = (obj.endC>=0)?obj.endC:printSetting.printArea.c2;
        this.workbook.print = (obj.print>=1&&obj.print<=3)?obj.print:this.workbook.print;
    }
}

Workbook.prototype.getPaperSize = function(Index){
    //一英寸等于2.54cm=25.4mm
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var printSetting = activeSheet.printSetting;
    var marginTop = printSetting.marginTop,
        marginBottom =printSetting.marginBottom,
        marginLeft = printSetting.marginLeft,
        marginRight = printSetting.marginRight,
        paper = printSetting.paper,
        dpi = this.dpi;
    var paperW = 0,paperH = 0;
    switch(paper.toUpperCase()){
        case 'A3':
            paperW = 297;paperH = 420;
            break;
        case 'A4':
            paperW = 210;paperH = 297;
            break;
        case 'A5':
            paperW = 148;paperH = 210;
            break;
        case 'A6':
            paperW = 105;paperH = 148;
            break;
        default :
            paperW = 210;paperH = 297;
    };
    //毫米转px
    paperW = parseInt(paperW/25.4*dpi);
    paperH = parseInt(paperH/25.4*dpi);
    marginTop = parseInt(marginTop/25.4*dpi)||0
    marginBottom = parseInt(marginBottom/25.4*dpi)||0
    marginLeft = parseInt(marginLeft/25.4*dpi)||0
    marginRight = parseInt(marginRight/25.4*dpi)||0
    return {'paperW':paperW,'paperH':paperH,'marginTop':marginTop,'marginBottom':marginBottom,'marginLeft':marginLeft,'marginRight':marginRight}
}

/**
 * @api {null} WB.splitcolHeader(Col,Name,Count,Height,index) splitcolHeader
 * @apiName splitcolHeader
 * @apiGroup Function
 * @apiDescription Split column header, only support splitting two rows
 * @apiParam {Int} Col  col(Col>=0)
 * @apiParam {String} Name  
 * @apiParam {Int} [Count=2] Number of columns split
 * @apiParam {Number} [Height=Half the height of the column head]  Height<The height of the column head
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.splitcolHeader(4,"name") //Set the name of the split column header of the fifth column to "name", split the two columns by default, and the height is half the height of the column header
 */
Workbook.prototype.splitcolHeader = function(Col,Name,Count,Height,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var colHeaderData = activeSheet.colHeaderData;
    if(colHeaderData.showColHeading){
        if(!colHeaderData.spans) activeSheet.colHeaderData.spans = [];
        var colHeaderSpans =  colHeaderData.spans;
        for(var i = colHeaderSpans.length-1;i>=0;i--){
            if(Col>=colHeaderSpans[i].col&&Col<colHeaderSpans[i].col+colHeaderSpans[i].colCount){
                colHeaderSpans.splice(i,1)
            }
        }
        var count = Count || 2;
        var height = Height || activeSheet.colHeaderData.defaultH/2;
        activeSheet.colHeaderData.spans.push({'col':Col,'name':Name,'colCount':count,'height':height});
    }
}

/**
 * @api {null} WB.removeSplitcolHeader(Col,index) removeSplitcolHeader
 * @apiName removeSplitcolHeader
 * @apiGroup Function
 * @apiDescription Delete split column header
 * @apiParam {Int} Col  col(Col>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.removeSplitcolHeader(4)   //Delete the split column header of the fifth column
 */
Workbook.prototype.removeSplitcolHeader = function(Col,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var colSpans = activeSheet.colHeaderData.spans;
    if(colSpans){
        for(var i = colSpans.length-1;i>=0;i--){
            if(colSpans[i].col==Col){
                activeSheet.colHeaderData.spans.splice(i,1)
            }
        }
    }
}

/**
 * @api {null} WB.setValue(R,C,Value,Index) setValue
 * @apiName setValue
 * @apiGroup Function
 * @apiDescription Set the cell's value (value text)
 * @apiParam {Int} R  row(R>=0)
 * @apiParam {Int} C  column(C>=0)
 * @apiParam {String} Value  
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setValue(0,0,'test',1)
 */

 /**
 * @api {null} /null onvaluechange
 * @apiName onvaluechange
 * @apiGroup Event
 * @apiDescription The function returns r, c, new value, old value
 * - Call this function to return if it is -1 do not perform setValue and setCellFormula operations If it is 1, then call the onrelevantchange event!
 * @apiParam {Function} callback 
 * @apiExample {javascript} demo:
    WB.onvaluechange=function(idx,r,c){
        //do something...
    };
 */
Workbook.prototype.setValue = function(R,C,Value,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,s = '',text = '',cellTypeName,child = document.querySelector('#'+this.boxId+' .child'),
    cellType = this.getCellType(R,C,R,C,index),formula = this.getCellFormula(R,C,R,C),formatter = this.getCellFormat(R,C,R,C),
    rows = activeSheet.rows;
    if(cellType){
        cellTypeName = cellType.name
    };
    if(R>=0&&C>=0){
        this.createData(R,C,index);
        if(data[R]&&data[R][C]){
            if(typeof(this.onvaluechange)=='function'&&this.workbook.stopEventCount==0){
                s = this.onvaluechange(R,C,Value,data[R][C].value);
                if(s=='-1') return
            };
            if(cellTypeName==3){
                var v = (Value)?1:0;
                data[R][C].value = v ; 
                data[R][C].text = (v).toString()
            }else{
                data[R][C].value = Value;
                if(typeof(Value)=='object'){
                    data[R][C].text = JSON.stringify(Value)
                }else{
                    if(Value!==undefined&&Value!==null){
                        data[R][C].text = Value.toString()
                    }
                };
                text = data[R][C].text
                if(formula&&formatter!='text'){
                    text = formula;
                }
                if(this.tempValue.r==R&&this.tempValue.c==C){
                    child.innerText = text
                }
                if(rows[R]&&!rows[R].noDefaultH){
                    var wordWrap = this.getWordWrap(R,C,R,C,index);
                    if(wordWrap){
                        this.incrementRowheight(R,C,index);
                    }
                } 
            }
        }  
    }
    if(s=='1'&&typeof(this.onrelevantchange)=='function'){
        this.onrelevantchange(R,C)
    }
}

/**
 * @api {null} WB.setText(R,C,Text,Index) setText
 * @apiName setText
 * @apiGroup Function
 * @apiDescription Set the text value of the cell
 * @apiParam {Int} R  row(R>=0)
 * @apiParam {Int} C  col(C>=0)
 * @apiParam {String} Text  
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setText = function(R,C,Text,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable;
    this.createData(R,C,index);
    if(typeof(Text)=='object'){
        data[R][C].text = JSON.stringify(Text)
    }else{
        if(Text!==undefined&&Text!==null){
            data[R][C].text = Text.toString()
        }
    };
}

/**
 * @api {null} WB.setCellPadding(topSize,rightSize,bottomSize,leftSize,Index) setCellPadding
 * @apiName setCellPadding
 * @apiGroup Function
 * @apiDescription Set padding of table cells
 * @apiParam {Number} topSize  padding-top
 * @apiParam {Number} rightSize  padding-right值
 * @apiParam {Number} bottomSize  padding-bottom
 * @apiParam {Number} leftSize  padding-left
 * @apiParam {Int} [Index=Current_table_index]  table index
 * @apiExample {javascript} demo:
    WB.setCellPadding(0,0,0,5)    //Set the right padding of the cell to 5px  
 */
Workbook.prototype.setCellPadding = function(topSize,rightSize,bottomSize,leftSize,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    activeSheet.cellPadding.top = topSize;
    activeSheet.cellPadding.right = rightSize;
    activeSheet.cellPadding.bottom = bottomSize;
    activeSheet.cellPadding.left = leftSize;
}

//打开下拉弹层
Workbook.prototype.openPopup = function(r,c,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,child = document.querySelector("#"+this.boxId+" .child"),
    cellType = this.getCellType(r,c,r,c,index),cellTypeName,cellTypeList,notSetIndex;
    if(cellType){
        cellTypeName = cellType.name;cellTypeList = cellType.list;notSetIndex = cellType.notSetIndex;
    };
    if(cellTypeName==4){
        var onopenpopup = this.onopenpopup,boxOffset = this.getBoxDiff(),fontSize = this.getFontList(r,c).originSize,
        selectsInfo = this.getSelectInfo(r,c),c1w =  selectsInfo.c1w,r1h = selectsInfo.r1h,top = selectsInfo.selectY,
        left = selectsInfo.selectX;
        if((window.top.popup||this.popup)&&typeof(onopenpopup)=='function'&&this.workbook.stopEventCount==0){
            onopenpopup(c1w,r1h,boxOffset.boxL+left+1,boxOffset.boxT+top+r1h,cellTypeList,r,c,child,fontSize,notSetIndex); 
        }   
    }
}

/**
 * @api {null} /null startPaint
 * @apiName startPaint
 * @apiGroup Function
 * @apiDescription The startPaint function sets the value of stopPaintedCount.Each time the value is decremented by 1 until it is 0.redraw when stopPaintedCount is 0.
 * - WB.startPaint()    
 * @apiParam {Boolean} [bool]  If true stopPaintedCount is directly assigned a value of 0 
 */
Workbook.prototype.startPaint = function(isAll){
    if(isAll){
        this.workbook.stopPaintedCount = 0;
    }
    if(this.workbook.stopPaintedCount>0){
        this.workbook.stopPaintedCount--
    }
    if(this.workbook.stopPaintedCount==0){
        this.drawTable()  
    }
}

/**
 * @api {null} /null stopPaint
 * @apiName stopPaint
 * @apiGroup Function
 * @apiDescription The stopPaint function sets the value of stopPaintCount, which is increased by 1 every time it is called.
 * - WB.stopPaint()    
 */
Workbook.prototype.stopPaint = function(){
    this.workbook.stopPaintedCount++;
}

/**
 * @api {null} /null resumeEvent
 * @apiName resumeEvent
 * @apiGroup Function
 * @apiDescription The resumeEvent function sets the value of stopEventCount. The value is decremented by 1 each time it is called until it is 0. stopEventCount is 0 and the on event is not executed.
 * - WB.resumeEvent()    
 * @apiParam {Boolean} [bool]  If true stopEventCount is directly assigned a value of 0 
 */
Workbook.prototype.resumeEvent = function(isAll){
    if(isAll){
        this.workbook.stopEventCount = 0;
    }
    if(this.workbook.stopEventCount>0){
        this.workbook.stopEventCount--
    }
    if(this.workbook.stopEventCount<0){
        this.workbook.stopEventCount = 0
    }
}

/**
 * @api {null} /null stopEvent
 * @apiName stopEvent
 * @apiGroup Function
 * @apiDescription The stopEvent function sets the value of stopEventCount, which is increased by 1 each time it is called.
 * - WB.stopEvent()    
 */
Workbook.prototype.stopEvent = function(){
    this.workbook.stopEventCount++;
}

Workbook.prototype.getMergeCount = function(r,c,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet;
    var activeSheet = this.getActiveSheet(index);
    var spans = activeSheet.spans;
    var rowCount = 1,colCount = 1;
    spans.forEach(function(item){
        if(item.row==r&&item.col==c){
            rowCount = item.rowCount;
            colCount = item.colCount;
        }
    });
    return {"rowCount":rowCount,"colCount":colCount}
}

/**
 * @api {null} WB.clearAllStyle(R1,C1,R2,C2,Index) clearAllStyle
 * @apiName clearAllStyle
 * @apiGroup Function
 * @apiDescription Restore the default style of the cells in the range (no parameters default to the currently selected range of cells)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.clearAllStyle = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable,rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style = JSON.parse(JSON.stringify(this.defaultCellStyle))
            }
        }
    }
}

/**
 * @api {null} WB.setFormatBrushStyle(R,C,Index) setFormatBrushStyle
 * @apiName setFormatBrushStyle
 * @apiGroup Function
 * @apiDescription Format brush style
 * @apiParam {Int} R1  row(R1>=0)
 * @apiParam {Int} C1  column(C1>=0)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.setFormatBrushStyle = function(R,C,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable;
    if(data[R]&&data[R][C]&&data[R][C].style){
        this.formatBrushStyle = data[R][C].style;
    }
};

/**
 * @api {null} /null clearFormatBrushStyle
 * @apiName clearFormatBrushStyle
 * @apiGroup Function
 * @apiDescription Clear format brush style 
 * - WB.clearFormatBrushStyle()
 */
Workbook.prototype.clearFormatBrushStyle = function(){
    this.formatBrushStyle = {};
}

/**
 * @api {null} WB.useFormatBrushStyle(R1,C1,R2,C2,Index) useFormatBrushStyle
 * @apiName useFormatBrushStyle
 * @apiGroup Function
 * @apiDescription Apply format brush styles within the range (no parameters default to the currently selected range of cells)
 * @apiParam {Int} [R1]  Start row(R1>=0)
 * @apiParam {Int} [C1]  Start column(C1>=0)
 * @apiParam {Int} [R2]  End row(R2>=0,R2>=R1)
 * @apiParam {Int} [C2]  End column(C2>=0,C2>=C1)
 * @apiParam {Int} [Index=Current_table_index]  table index
 */
Workbook.prototype.useFormatBrushStyle = function(R1,C1,R2,C2,Index){
    var index = (Index>=0)?Index:this.workbook.activeSheet,activeSheet = this.getActiveSheet(index);
    var data = activeSheet.data.dataTable, rcArg = this.handleArg(R1,C1,R2,C2,index),r1 = rcArg.r1,
    c1 = rcArg.c1,r2 = rcArg.r2,c2 = rcArg.c2;
    for(var i = r1;i<=r2;i++){
        for(var j = c1;j<=c2;j++){
            this.createData(i,j,index)
            if(data[i]&&data[i][j]&&data[i][j].style){
                data[i][j].style = this.formatBrushStyle;
            }
        }
    }
}

/**
 * @api {null} /null setDocument
 * @apiName setDocument
 * @apiGroup Function
 * @apiDescription Set the document that triggered the click event  
 * - WB.setDocument(D)
 * @apiParam {obj} d  
 * @apiExample {javascript} demo:
    WB.setDocument(window.top.document)   //Set to top
 */
Workbook.prototype.setDocument = function(D){
    for(var key in this.documentHandle){
        this.removeEvent(this.document,key,this.documentHandle[key]);
        this.document = D || this.document;
        this.addEvent(this.document,key,this.documentHandle[key])
    }
}

/**
 * @api {null} /null adaption
 * @apiName adaption
 * @apiGroup Function
 * @apiDescription Adaptive function
 * - WB.adaption(bool,devW,devH)  You can also turn on adaptation by setting the adaption attribute
 * @apiParam {Boolean} bool   
 * @apiParam {Number} devW The horizontal resolution of the computer screen during development
 * @apiParam {Number} devH The vertical resolution of the computer screen during development
 * @apiExample {javascript} demo:
    WB.adaption(true,1280,720)   //Turn on adaptive
 */
Workbook.prototype.adaption = function(bool,devW,devH){
    this.workbook.devScreenWidth = devW || this.workbook.devScreenWidth;
    this.workbook.devScreenHeight = devH || this.workbook.devScreenHeight;
    this.workbook.adaption = bool;
    if(this.workbook.adaption){
        var oldWidth = this.width;
        var oldHeight = this.height
        var newW = this.workbook.width/this.workbook.devScreenWidth*window.innerWidth
        var newH = this.workbook.height/this.workbook.devScreenHeight*window.innerHeight
        this.setTableSize(newW,newH)
        for(var k in this.workbook.sheets){
            var activeSheet = this.workbook.sheets[k];
            var cols = activeSheet.columns,rows = activeSheet.rows;
            activeSheet.colHeaderData.defaultH = activeSheet.colHeaderData.defaultH/oldHeight*this.height
            activeSheet.rowHeaderData.defaultW = activeSheet.rowHeaderData.defaultW/oldWidth*this.width
            for(var i = 0;i<cols.length;i++){
                if(cols[i]&&cols[i].bHidden){
                    activeSheet.columns[i].tempSize = activeSheet.columns[i].tempSize/oldWidth*this.width
                }else if(cols[i]&&!cols[i].bHidden){
                    activeSheet.columns[i].size = activeSheet.columns[i].size/oldWidth*this.width
                }
            }; 
            for(var j = 0;j<rows.length;j++){
                if(rows[j]&&rows[j].bHidden){
                    activeSheet.rows[j].tempSize =  activeSheet.rows[j].tempSize/oldHeight*this.height
                }else if(rows[j]&&!rows[j].bHidden){
                    activeSheet.rows[j].size = activeSheet.rows[j].size/oldHeight*this.height
                }
            }
        } 
    }
};


/**
 * @api {null} /null addChildCanvas
 * @apiName addChildCanvas
 * @apiGroup Function
 * @apiDescription Add a child canvas element to insert into the page
 * - The function is mainly used to pop up another layer in the existing canvas table, which is another workbook
 * - WB.addChildCanvas(options)   
 * @apiParam {Object} options Sub-canvas related parameters
 * - id(Child container id) This parameter needs to be passed to specify the container of the child canvas
 * - parentId(Parent container id)
 * - width(Defaults:250)
 * - height(Defaults:300)
 * - x(Defaults:0)
 * - y(Defaults:0)
 * - shade(Defaults:false) 
 * @apiExample {javascript} demo:
    WB.addChildCanvas({
        width:250,  //The width of the child canvas and the width of the container
        height:300, //The height of the child canvas and the height of the container
        x:280,  //Relative to the x coordinate of the parent container
        y:90,   //Relative to the parent container's y coordinate
        parentId:"container",   //Parent container default document.body
        id:"container2",    //Container id of child canvan
        shade:true, //Whether to show the shade
    });

    var WB2 = new Workbook("container2")

    WB2.setTableSize(250,300)
    WB2.maxCol(2)
    WB2.maxRow(3)
    WB2.setColWidth (125,0,1)
    WB2.showColHeading(true) 
    WB2.setHeaderName(-1,0,'project')
    WB2.setHeaderName(-1,1,'Numerical value')
    WB2.activeSheet.colHeaderData.defaultDataNode.style.fillColor = "#A9A9A9"
    WB2.workbook.tabsOptions.fontColor = '#1E90FF'
    WB2.workbook.showTabs = 2
    WB2.bottomBtn=['yes','cancel']
    WB2.tabsHeight = 30;
    WB2.workbook.behaviorMode = 3
    WB.setActiveCanvas('container');
    WB2.setValue(0,1,"Hidden bomb",0)
    WB2.setLock(0,1,0,1,true)
    WB2.setValue(0,0,"data",0)
    WB2.startPaint()
    WB.onclickcellopenlayer = function(ind,r,c){
        if(r==0&&c==1){
            WB.showChildCanvas("container2")   //show
            WB.setActiveCanvas('container2');
        }
    }
    WB2.onclickcellopenlayer = function(ind,r,c){
        if(r==0&&c==1){
            WB2.hiddenChildCanvas("container2")   //hide
            WB.setActiveCanvas('container');
        }
    }
 */
Workbook.prototype.addChildCanvas = function(options){
    var width = options.width || 250;
    var height = options.height || 300;
    var x = options.x || 0;
    var y = options.y || 0;
    var id = options.id;
    var parentId = options.parentId
    var shade = options.shade || false;
    var parentEle = document.getElementById(parentId) || document.body
    var shadeB = document.createElement('div');
    shadeB.setAttribute('id',id+'shade')
    if(shade){
        shadeB.style.cssText = 'height:100%;width:100%;position:absolute;top:0;left:0;background-color:rgba(0,0,0,.3);display:none';
        parentEle.appendChild(shadeB);
    }
    var divB = document.createElement('div')
    divB.setAttribute('id',id);
    divB.setAttribute('class','clearfix')
    divB.style.cssText = 'height:'+height+'px;width:'+width+'px;position:absolute;top:'+y+'px;left:'+x+'px;z-index:99;display:none'
    parentEle.appendChild(divB);
}

/**
 * @api {null} /null hiddenChildCanvas
 * @apiName hiddenChildCanvas
 * @apiGroup Function
 * @apiDescription Hide all elements of child canvas
 * - WB2.hiddenChildCanvas("container2")  
 * @apiParam {String} id Id of the child canvas container
 */
Workbook.prototype.hiddenChildCanvas = function(id){
    var divBEle = document.getElementById(id);
    var shadeBEle = document.getElementById(id+'shade');
    if(divBEle)     divBEle.style.display = 'none';
    if(shadeBEle)   shadeBEle.style.display = 'none';
    this.removeTextArea()
}

/**
 * @api {null} /null setActiveCanvas
 * @apiName setActiveCanvas
 * @apiGroup Function
 * @apiDescription Set the active canvas when there are multiple canvases on the same page (the default is the last new canvas, click on the corresponding canvas activecanvas will automatically switch)
 * @apiParam {String} id 
 * @apiExample {javascript} demo:
    WB.setActiveCanvas("container2")  
 */
Workbook.prototype.setActiveCanvas = function(id){
    window.boxId = id
}

/**
 * @api {null} /null showChildCanvas
 * @apiName showChildCanvas
 * @apiGroup Function
 * @apiDescription Show all elements of child canvas
 * @apiParam {String} id 
 * @apiExample {javascript} demo:
    WB.showChildCanvas("container2")  
 */
Workbook.prototype.showChildCanvas = function(id){
    var divBEle = document.getElementById(id);
    var shadeBEle = document.getElementById(id+'shade');
    if(divBEle){
        divBEle.style.display = 'block';
        divBEle.style.position = 'absolute'
    } 
    if(shadeBEle)   shadeBEle.style.display = 'block';
    this.removeTextArea()
}


/**
 * @api {null} WB.exportFileAsJson() exportFileAsJson
 * @apiName exportFileAsJson
 * @apiGroup Function
 * @apiDescription Save the current workbook as a json file and download
 */
Workbook.prototype.exportFileAsJson = function(){
    try{
        var hasBlob=false
        try{hasBlob=!!Blob}catch(ex){};
        if(hasBlob){
            var blob = new Blob([JSON.stringify(this.workbook)], {type: ""});
            saveAs(blob, "workbook.json"); 
        }else{
            if(window.ActiveXObject){
                var data = JSON.stringify(this.workbook);
                var fso = new ActiveXObject("Scripting.FileSystemObject");
                var mygetfolder=new ActiveXObject("WScript.shell");
                if(mygetfolder.SpecialFolders("Desktop")){
                    var savePath = mygetfolder.SpecialFolders("Desktop");  
                    savePath = savePath.replace(/\\/g,'\\\\');
                    var f;
                    f = fso.CreateTextFile(savePath+'\\'+'workbook.json',true,true); 
                    f.write(data);
                    f.Close();
                }
            }
        }
    }catch(e){}
}

/**
 * @api {null} WB.exportFileAsXml()  exportFileAsXml
 * @apiName exportFileAsXml
 * @apiGroup Function
 * @apiDescription Save the current workbook as an xml file and download it (each table generates an xml)
 */
Workbook.prototype.exportFileAsXml = function(){
    var _this = this
    this.json2Xml(function(xmlDoc,filename){   
        try{
            var hasBlob=false
            try{hasBlob=!!Blob}catch(ex){};
            if(hasBlob){
                var blob = new Blob([_this.xml2String(xmlDoc)], {type: ""});  // saveAs(Blob/File/Url)  ie10+ 支持blob  
                saveAs(blob, filename+'.xml'); 
            }else{
                if(window.ActiveXObject){
                    var data = _this.xml2String(xmlDoc);
                    var fso = new ActiveXObject("Scripting.FileSystemObject");
                    var mygetfolder = new ActiveXObject("WScript.shell");     
                    if(mygetfolder.SpecialFolders("Desktop")){     
                        var savePath = mygetfolder.SpecialFolders("Desktop");  
                        savePath = savePath.replace(/\\/g,'\\\\');
                        var f;
                        f = fso.CreateTextFile(savePath+'\\'+filename+".xml",true,true);   //文件 是否覆盖 编码格式
                        f.write(data);
                        f.Close();
                    }   
                }
            }
        }catch(e){}
    })
}

Workbook.prototype.exportFileAsZip = function(){
   // to do
}

/**
 * @api {null} /null importFile
 * @apiName importFile
 * @apiGroup Function
 * @apiDescription File import (the file only supports .json, .xml, .xlsx, excel.zip)
 * - WB.exportFileAsXml()  f
 * @apiParam {Object} f file
 * @apiParam {Function} callback 
 */
Workbook.prototype.importFile = function(f,callback){
    try{
        var dIndex,ext,_this = this;
        var hasFileRead = false;
        try{hasFileRead=!!FileReader}catch(ex){};
        if(hasFileRead){
            var reader = new FileReader();  //ie10+
            dIndex = f.name.lastIndexOf('.');
            ext = f.name.slice(dIndex+1);
            if(ext == 'json'){
                reader.readAsText(f)
                reader.onload = function(){
                    _this.workbook = {}
                    _this.workbook = JSON.parse(this.result)
                    _this.workbook = _this.initWorkbook()
                    afterImport(_this.workbook.activeSheet,callback)
                }
            }else if((ext == 'xlsx')||(ext=="zip")){
                this.excel2Json(f,function(){
                    afterImport(_this.workbook.activeSheet,callback)
                })
            }else if(ext == 'xml'){
                reader.readAsText(f)
                reader.onload = function(){
                    _this.xml2Json(this.result,f,function(){
                        afterImport(_this.workbook.activeSheet,callback)
                    })
                }
            }
        }else{
            if(window.ActiveXObject){
                //to do
            }
        };
        function afterImport(i,cb){
            _this.sheetIndex(i)
            _this.getStartSheet()
            _this.getEndSheet()
            _this.startPaint(true)
            if(typeof(cb)=='function'){
                cb()
            }
        }
    }catch(e){}
}

//导进来的xml文件转成json
Workbook.prototype.xml2Json = function(result,f,callback){
    var xmlData = this.string2XML(result)
    this.workbook = {};
    this.workbook = this.initWorkbook();
    this.workbook.showTabs = 1;  
    this.workbook.activeSheet = 0;
    this.workbook.sheetList = [];
    this.workbook.sheets = {}
    var fileName = f.name;
    fileName = fileName.slice(0,fileName.lastIndexOf('.'))
    var tableTag = xmlData.getElementsByTagName('TABLE')
    for(var i = 0;i<tableTag.length;i++){
        var sheetName;
        if(tableTag.length==1){
            sheetName = fileName
        }else{
            sheetName = fileName+'_'+(i+1)
        }
        this.workbook.sheetList.push(sheetName)
        this.workbook.sheets[sheetName]= {}
        this.initSheet(sheetName,i); 
        var activeSheet = this.getActiveSheet(i)
        var bodyTag = tableTag[i].getElementsByTagName('TBODY')[0]||tableTag[i];
        var tableClass = tableTag[i].getAttribute('class');
        var pagesetup =  tableTag[i].getAttribute('pagesetup');
        var trTag = bodyTag.getElementsByTagName('TR');
        if(tableTag[0]){
            var defaultFontSize = this.getXmlStyleValue(tableTag[0],'FONT-SIZE') || this.workbook.defaultFontSize;
            defaultFontSize = this.pt2px(parseFloat(defaultFontSize))   //磅转px
            var defaultFontFamily = this.getXmlStyleValue(tableTag[0],'FONT-FAMILY') || this.workbook.defaultFontName;
            this.workbook.defaultFontSize = defaultFontSize;
            this.workbook.defaultFontName = defaultFontFamily;
            this.defaultCellStyle={"font":this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName}  
        }
        //打印相关设置
        var paper = this.getXmlPageSetUpValue(pagesetup,'paper');
        var marginTop = this.getXmlPageSetUpValue(pagesetup,'marginTop');
        var marginBottom = this.getXmlPageSetUpValue(pagesetup,'marginBottom');
        var marginLeft = this.getXmlPageSetUpValue(pagesetup,'marginLeft');
        var marginRight = this.getXmlPageSetUpValue(pagesetup,'marginRight');
        var gridLine = this.getXmlPageSetUpValue(pagesetup,'gridline');
        var orientation = this.getXmlPageSetUpValue(pagesetup,'orientation');
        var marginCopies = this.getXmlPageSetUpValue(pagesetup,'marginCopies');
        var printer = this.getXmlPageSetUpValue(pagesetup,'printer');
        var isshowfooterpageinfo = this.getXmlPageSetUpValue(pagesetup,'isshowfooterpageinfo');
        var footpagestyle = this.getXmlPageSetUpValue(pagesetup,'footpagestyle')
        if(paper){
            activeSheet.printSetting.paper = paper
        }
        if(marginTop){
            activeSheet.printSetting.marginTop = parseInt(marginTop);
        };
        if(marginBottom){
            activeSheet.printSetting.marginBottom = parseInt(marginBottom);
        };
        if(marginLeft){
            activeSheet.printSetting.marginLeft = parseInt(marginLeft);
        };
        if(marginRight){
            activeSheet.printSetting.marginRight = parseInt(marginRight);
        };
        activeSheet.printSetting.printGridLine = parseInt(gridLine)?true:false;
        if(orientation){
            activeSheet.printSetting.orientation = parseInt(orientation)
        };
        if(marginCopies){
            this.workbook.marginCopies = parseInt(marginCopies)
        };
        if(printer){
           this.workbook.printer = printer
        }
        if(isshowfooterpageinfo){
            activeSheet.printSetting.isshowfooterpageinfo = isshowfooterpageinfo
        }
        if(footpagestyle){
            activeSheet.printSetting.footpagestyle = footpagestyle;
        }

        if(trTag){
            var rowLen = (trTag.length-1>0)?trTag.length-1:0;
            this.maxRow(rowLen,i)   //row
            var firstRow = trTag[0];
            var firstRowTd = firstRow.getElementsByTagName('TD');
            var colLen = (firstRowTd.length-1>0)?firstRowTd.length-1:0;
            this.maxCol(colLen,i);  //columns
            if(firstRowTd){
                for(var a = 1;a<firstRowTd.length;a++){
                    var tdW = this.getXmlStyleValue(firstRowTd[a],'WIDTH') || activeSheet.defaults.colWidth
                    tdW = parseInt(tdW);
                    this.setColWidth(tdW,a-1,a-1,i)   //columns width
                }
            };
            for(var a = 1;a<trTag.length;a++){
                var row = a-1
                var trH = this.getXmlStyleValue(trTag[a],'HEIGHT')
                trH = parseInt(trH);
                this.setRowHeight(trH,row,row,i)  //row height
                var tdTag = trTag[a].getElementsByTagName('TD');
                if(tdTag){
                    var colSpanCount = 0 ,rowSpanCount = 0;
                    for(var c = 1;c<tdTag.length;c++){
                        var col = c-1+colSpanCount
                        var colSpan = parseInt(tdTag[c].getAttribute('colSpan'))||1;
                        var rowSpan = parseInt(tdTag[c].getAttribute('rowSpan'))||1;
                        if(colSpan>1){
                            colSpanCount += colSpan-1;
                        }
                        if(rowSpan>1&&c){
                            rowSpanCount += rowSpan-1;
                        }
                        if(colSpan>1||rowSpan>1){
                            this.mergeCells(row,col,row+rowSpan-1,col+colSpan-1,i)  //spans
                        }
                        activeSheet.spans.forEach(function(item){
                            if(row>item.row&&row<item.row+item.rowCount&&col>=item.col&&col<item.col+item.colCount){
                                col += item.colCount;
                                colSpanCount += item.colCount
                            }
                        })

                        this.createData(row,col,i);
                        if(tdTag[c].firstChild){
                            var value = tdTag[c].firstChild.nodeValue;
                            // if(re.test(value))  value = Number(value)
                            this.setValue(row,col,value,i)    //value
                        };

                        //class 样式转换
                        var classAttr = tdTag[c].getAttribute('class');
                        if(classAttr&&tableClass&&tableClass.indexOf('tblGenFixed')!=-1){
                            if(classAttr.indexOf('rg')!=-1){
                                this.fontColor(row,col,row,col,'blue',i)
                            };
                            if(classAttr.indexOf('rv')!=-1){
                                this.setFillColor('#D8F173',row,col,row,col,i)
                            };
                            if(classAttr.indexOf('rs')!=-1){
                                this.setFillColor('#BFDB4D',row,col,row,col,i)
                            };
                            if(classAttr.indexOf('s1')!=-1){
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-RIGHT',i)
                                this.setChangeBorder(row,col,'white .5pt solid','BORDER-BOTTOM',i)
                            };
                            if(classAttr.indexOf('s2')!=-1){
                                this.setChangeBorder(row,col,'white .5pt solid','BORDER-RIGHT',i)
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-BOTTOM',i)
                            };
                            if(classAttr.indexOf('s3')!=-1){
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-RIGHT',i)
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-BOTTOM',i)
                            };
                            if(classAttr.indexOf('sl')!=-1){
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-LEFT',i)
                            };
                            if(classAttr.indexOf('st')!=-1){
                                this.setChangeBorder(row,col,'black .5pt solid','BORDER-TOP',i)
                            };
                            if(classAttr.indexOf('fb')!=-1){
                                this.setFontBold(row,col,row,col,true,i)
                            };
                            if(classAttr.indexOf('wb')!=-1){
                                this.wordWrap(row,col,row,col,true,i)
                            };
                            if(classAttr.indexOf('fi')!=-1){
                                this.setFontItalic(row,col,row,col,true,i)
                            };
                            if(classAttr.indexOf('fu')!=-1){
                                this.fontUnderline(row,col,row,col,true,i)
                            };
                            if(classAttr.indexOf('fd')!=-1){
                                this.fontStrikeout(row,col,row,col,true,i)
                            };
                            if(classAttr.indexOf('a0')!=-1){
                                this.setVAlignment(row,col,row,col,2,i)
                                this.setHAlignment(row,col,row,col,2,i)
                            };
                            if(classAttr.indexOf('a1')!=-1){
                                this.setVAlignment(row,col,row,col,2,i)
                                this.setHAlignment(row,col,row,col,4,i)
                            };
                            if(classAttr.indexOf('a2')!=-1){
                                this.setVAlignment(row,col,row,col,2,i)
                                this.setHAlignment(row,col,row,col,3,i)
                            };
                            if(classAttr.indexOf('a3')!=-1){
                                this.setVAlignment(row,col,row,col,1,i)
                                this.setHAlignment(row,col,row,col,2,i)
                            };
                            if(classAttr.indexOf('a4')!=-1){
                                this.setVAlignment(row,col,row,col,1,i)
                                this.setHAlignment(row,col,row,col,4,i)
                            };
                            if(classAttr.indexOf('a5')!=-1){
                                this.setVAlignment(row,col,row,col,1,i)
                                this.setHAlignment(row,col,row,col,3,i)
                            };
                            if(classAttr.indexOf('a6')!=-1){
                                this.setVAlignment(row,col,row,col,3,i)
                                this.setHAlignment(row,col,row,col,2,i)
                            };
                            if(classAttr.indexOf('a7')!=-1){
                                this.setVAlignment(row,col,row,col,3,i)
                                this.setHAlignment(row,col,row,col,4,i)
                            };
                            if(classAttr.indexOf('a8')!=-1){
                                this.setVAlignment(row,col,row,col,3,i)
                                this.setHAlignment(row,col,row,col,3,i)
                            };
                        }

                        //style行内样式转换
                        var style = tdTag[c].getAttribute('style');
                        if(style){
                            var fontSize = this.getXmlStyleValue(tdTag[c],'FONT-SIZE')  //fontsize
                            if(fontSize){
                                fontSize = this.pt2px(parseFloat(fontSize))  
                                this.setFontSize(row,col,row,col,fontSize,i)
                            };
        
                            var fontFamily = this.getXmlStyleValue(tdTag[c],'FONT-FAMILY')  //fontname
                            if(fontFamily){
                                this.setFontName(row,col,row,col,fontFamily,i)
                            }
        
                            var fontWeight = this.getXmlStyleValue(tdTag[c],'FONT-WEIGHT') //fontweight
                            if(fontWeight){
                                this.setFontBold(row,col,row,col,true,i)
                            }
        
                            var fontStyle = this.getXmlStyleValue(tdTag[c],'FONT-STYLE')    //fontstyle
                            if(fontStyle){
                                this.setFontItalic(row,col,row,col,true,i)
                            };
        
                            var fontColor = this.getXmlStyleValue(tdTag[c],'COLOR');    //fontcolor
                            if(fontColor){
                                this.fontColor(row,col,row,col,fontColor,i)
                            }
        
                            var formula = tdTag[c].getAttribute('x:fmla')  //formula
                            if(formula){
                                this.setCellFormula(row,col,formula,i)
                            }
        
                            var bgcolor = this.getXmlStyleValue(tdTag[c],'BACKGROUND-COLOR') //gbcolor
                            if(bgcolor){
                                this.setFillColor(bgcolor,row,col,row,col,i)
                            }
        
                            var hAlign = this.getXmlStyleValue(tdTag[c],'TEXT-ALIGN')   //halign
                            if(hAlign){
                                hAlign = (hAlign=='left')?2:(hAlign=='center')?3:(hAlign=='right')?4:1;
                                this.setHAlignment(row,col,row,col,hAlign,i)
                            }
        
                            var vAlign = this.getXmlStyleValue(tdTag[c],'VERTICAL-ALIGN')   //valian
                            if(vAlign){
                                vAlign = (vAlign=='top')?1:(vAlign=='center')?2:3;
                                this.setVAlignment(row,col,row,col,vAlign,i)
                            }
        
                            var decoration = this.getXmlStyleValue(tdTag[c],'TEXT-DECORATION') // underline line-through
                            if(decoration){
                                if(decoration.indexOf('underline')!=-1){
                                    this.fontUnderline(row,col,row,col,true,i)
                                };
                                if(decoration.indexOf('line-through')!=-1){
                                    this.fontStrikeout(row,col,row,col,true,i)
                                }
                            }
        
                            var wordBreak = this.getXmlStyleValue(tdTag[c],'WORD-BREAK')    //word-break
                            if(wordBreak){
                                this.wordWrap(row,col,row,col,true,i)
                            }

                            var whiteSpace = this.getXmlStyleValue(tdTag[c],'WHITE-SPACE')    //word-break
                            if(whiteSpace=='nowrap'){
                                this.wordWrap(row,col,row,col,false,i)
                            }
        
                            var border = this.getXmlStyleValue(tdTag[c],'BORDER')    //border  取颜色以及width(对应style：thin medium thick)
                            this.setChangeBorder(row,col,border,'BORDER',i)
        
                            var borderLeft = this.getXmlStyleValue(tdTag[c],'BORDER-LEFT')    
                            this.setChangeBorder(row,col,borderLeft,'BORDER-LEFT',i)
        
                            var borderRight = this.getXmlStyleValue(tdTag[c],'BORDER-RIGHT')    
                            this.setChangeBorder(row,col,borderRight,'BORDER-RIGHT',i)
        
                            var borderTop = this.getXmlStyleValue(tdTag[c],'BORDER-TOP')    
                            this.setChangeBorder(row,col,borderTop,'BORDER-TOP',i)
        
                            var borderBottom = this.getXmlStyleValue(tdTag[c],'BORDER-BOTTOM')    
                            this.setChangeBorder(row,col,borderBottom,'BORDER-BOTTOM',i)
                            
        
                            var format = this.getXmlStyleValue(tdTag[c],'MSO-NUMBER-FORMAT') //format
                            if(format){
                                this.xmlFormat2JsonFormat(row,col,format,i) 
                            }
                        } 
                    }
                };
            };
            this.row(0,i)  //  设置活动单元格
            this.col(0,i)
            this.setTableSize(window.innerWidth-this.scrollbarWidth,window.innerHeight-this.scrollbarWidth)    //workbook (width height)
        }

    this.showColHeading(true,i)
    this.showRowHeading(true,i)
    }
    callback()
}

//json格式转成xml文档
Workbook.prototype.json2Xml = function(callback){
    for(var i = 0;i<this.workbook.sheetList.length;i++){
        var xmlDoc = this.createXmlDoc()
        var activeSheet = this.getActiveSheet(i),data = activeSheet.data.dataTable,
        cols = activeSheet.columns,rows = activeSheet.rows,totolW = 0;
        var statement = xmlDoc.createProcessingInstruction("xml","version=\"1.0\" encoding=\"UTF-8\"");  
        xmlDoc.appendChild(statement); 
        var table = xmlDoc.createElement('TABLE');
        var tbody = xmlDoc.createElement('TBODY')
        //打印设置
        var paper = activeSheet.printSetting.paper;
        paper = (paper)?"paper:'"+paper+"',":'';
        var orientation = activeSheet.printSetting.orientation;
        orientation = (orientation)?"orientation:'"+orientation+"',":'';
        var printGridLine = (activeSheet.printSetting.printGridLine)?1:0;
        printGridLine =  "gridline:'"+printGridLine+"',";
        var marginCopies= this.workbook.marginCopies;
        marginCopies =  (parseInt(marginCopies)>=0)?"marginCopies:'"+marginCopies+"mm',":'';
        var printer = this.workbook.printer;
        printer =  (printer)?"printer:'"+printer+"',":'';
        var isshowfooterpageinfo = activeSheet.printSetting.isshowfooterpageinfo;
        isshowfooterpageinfo = (parseInt(isshowfooterpageinfo)>=0)?"isshowfooterpageinfo:'"+isshowfooterpageinfo+"',":'';
        var footpagestyle = activeSheet.printSetting.footpagestyle;
        footpagestyle = (footpagestyle)?"footpagestyle:'"+footpagestyle+"',":'';
        var marginTop = activeSheet.printSetting.marginTop;
        marginTop = (parseInt(marginTop)>=0)?"marginTop:'"+marginTop+"mm',":'';
        var marginBottom = activeSheet.printSetting.marginBottom;
        marginBottom = (parseInt(marginBottom)>=0)?"marginBottom:'"+marginBottom+"mm',":'';
        var marginLeft = activeSheet.printSetting.marginLeft;
        marginLeft = (parseInt(marginLeft)>=0)?"marginLeft:'"+marginLeft+"mm',":'';
        var marginRight = activeSheet.printSetting.marginRight;
        marginRight = (parseInt(marginRight)>=0)?"marginRight:'"+marginRight+"mm',":'';
        table.setAttribute('pagesetup',"{"+paper+orientation+printGridLine+marginCopies+printer+isshowfooterpageinfo+footpagestyle+marginTop+marginBottom+marginLeft+marginRight+"}")
        table.setAttribute('class','tblGenFixed')
        for(var a = 0;a<cols.length;a++){
            if(cols[a]) totolW+=cols[a].size;
        };
        totolW = (totolW)?'WIDTH:'+totolW+'px;':'';
        var tableFontSize = (this.workbook.defaultFontSize)?'FONT-SIZE:'+(this.px2pt(this.workbook.defaultFontSize)+'pt;'):'';
        table.setAttribute('style',totolW+tableFontSize)
        xmlDoc.appendChild(table);
        var tableEle = xmlDoc.getElementsByTagName('TABLE')[0];
        tableEle.appendChild(tbody);
        var tbodyEle = xmlDoc.getElementsByTagName('TBODY')[0]||tableEle;
        //找到有单元格最大的行数和列数
        var rowKey = Object.keys(data),maxRow = 0;
        if(rowKey.length>=1){
            maxRow = Math.max.apply(null,rowKey)
        }
        var maxColArr = [],maxCol = 0
        for(var t = 0;t<=maxRow;t++){
            if(data[t]){
                var colkey = Object.keys(data[t]);
                if(colkey.length>=1){
                    maxCol = Math.max.apply(null,colkey)
                }
                maxColArr.push(maxCol)
            }
        };
        if(maxColArr.length>=1){
            maxCol = Math.max.apply(null,maxColArr)   
        }

        if(rows.length>0){  //firstRow
            var firstTr = xmlDoc.createElement('TR');
            firstTr.setAttribute('class','firstrow')
            firstTr.setAttribute('style','HEIGHT:1px');
            tbodyEle.appendChild(firstTr);
            var firstTrEle = tbodyEle.getElementsByTagName('TR')[0];
            for(var k = -1;k<=maxCol;k++){
                var firstTrTd = xmlDoc.createElement('TD');
                var firstTrTdW = activeSheet.defaults.colWidth
                if(k==-1){
                    firstTrTdW = 1
                }else{
                    if(cols[k]) firstTrTdW = cols[k].size
                };
                firstTrTd.setAttribute('style','WIDTH:'+firstTrTdW+'px');
                firstTrEle.appendChild(firstTrTd)
            }
        };
        for(var a = 0;a<=maxRow;a++){   //每一行的style
            var tr = xmlDoc.createElement('TR');
            var trH = rows[a].size || activeSheet.defaults.rowHeight;
            var trStyle = 'DISPLAY: block; HEIGHT:'+trH+'px';
            tr.setAttribute('style',trStyle);
            tbodyEle.appendChild(tr);            
        };
        var trEle = tbodyEle.getElementsByTagName('TR')||[];
        var row,col,spans = activeSheet.spans;
        for(var a = 1;a<trEle.length;a++){  //每一行下的每一个td
            row = a-1;
            var firstTd = xmlDoc.createElement('TD');
            trEle[a].appendChild(firstTd);  //每行(除第一行) 前面插入一个空的TD
            for(var b = 0;b<=maxCol;b++){   //插入有样式的TD
                col = b;
                var td = xmlDoc.createElement('TD');
                var colCount = 0,rowCount = 0;
                spans.forEach(function(item){
                    if(row==item.row&&col==item.col){
                        colCount = item.colCount;
                        rowCount = item.rowCount;
                        if(colCount>1) td.setAttribute('colSpan',colCount);
                        if(rowCount>1) td.setAttribute('rowSpan',rowCount)
                    }
                });
                trEle[a].appendChild(td);
            }
        };

        for(var rKey in data){
            var r = parseInt(rKey)
            for(var cKey in data[rKey]){
                var c = parseInt(cKey);
                var cell = trEle[r+1].getElementsByTagName('TD')[c+1];
                var value = this.getValue(r,c,r,c,i)   //innerHtml  table元素ie6-10不支持
                if(value){
                    cell.appendChild(xmlDoc.createTextNode(value))
                }   
                var align = this.getAlignment(r,c,r,c,i)   //对齐
                if(align){
                    var hAlign = align.hAlign,vAlign = align.vAlign;
                    hAlign = (hAlign==2)?'left':(hAlign==3)?'center':(hAlign==4)?'right':'';
                    if(hAlign) hAlign = 'TEXT-ALIGN:'+hAlign+';';
                    vAlign = (vAlign==1)?'top':(vAlign==3)?'bottom':'';
                    if(vAlign) vAlign = 'VERTICAL-ALIGN:'+vAlign+';';
                };

                var formula = this.getCellFormula(r,c,r,c,i);   //公式
                if(formula){
                    cell.setAttribute('x:fmla',formula)
                }

                var fontList = this.getFontList(r,c,i)  
                var fontName = fontList.originName||''  //font-family
                if(fontName!=this.workbook.defaultFontName){
                    fontName = 'FONT-FAMILY:'+fontName+';'
                }else{
                    fontName = ''
                }

                var fontItailc = fontList.originItalic||''  //font-style
                if(fontItailc){
                    fontItailc = 'FONT-STYLE:'+fontItailc+';'
                }

                var fontBold = fontList.originBold||''  //font-weight
                if(fontBold){
                    fontBold = 'FONT-WEIGHT:'+fontBold+';'
                }

                var fontSize = fontList.originSize||'';  //font-size
                if(fontSize!=this.workbook.defaultFontSize){
                    fontSize = 'FONT-SIZE:'+this.px2pt(fontSize)+'pt;'
                }else{
                    fontSize = ''
                }

                var fontColor = this.getFontColor(r,c,r,c,i)||''    //color
                if(fontColor){
                    fontColor = 'COLOR:'+fontColor+';'
                }

                var bgColor = this.getFillColor(r,c,r,c,i)||''; //background-color
                if(bgColor){
                    bgColor = 'BACKGROUND-COLOR:'+bgColor+';'
                }

                var underline = this.getFontUnderline(r,c,r,c,i)||''    //text-decoration
                var strikeout = this.getFontStrikeout(r,c,r,c,i)||''
                var text_decoration = ''
                if(underline&&strikeout){
                    text_decoration = 'TEXT-DECORATION:underline line-through;'
                }else if(underline&&!strikeout){
                    text_decoration = 'TEXT-DECORATION:underline;'
                }else if(!underline&&strikeout){
                    text_decoration = 'TEXT-DECORATION:line-through;'
                }

                var word_break = this.getWordWrap(r,c,r,c,i)||''    //WORD-BREAK
                if(word_break){
                    word_break = 'WORD-BREAK:break-all;'
                }

                var border = this.getBorder(r,c,r,c,i);
                var borderRight='',borderLeft='',borderBottom='',borderTop='';
                if(border){
                    borderRight = border.borderRight||'',rStyle=1,rColor = '#000';  //border-right
                    if(borderRight){
                        rStyle = (borderRight.style==1)?'thin':(borderRight.style==2)?'medium':(borderRight.style==5)?'thick':'thin';
                        rColor =  borderRight.color||'#000';
                        borderRight = 'BORDER-RIGHT:'+rStyle+' solid '+rColor+';';
                    }
    
                    borderLeft = border.borderLeft||'',lStyle=1,lColor = '#000';    //border-left
                    if(borderLeft){
                        lStyle = (borderLeft.style==1)?'thin':(borderLeft.style==2)?'medium':(borderLeft.style==5)?'thick':'thin';
                        lColor =  borderLeft.color||'#000';
                        borderLeft = 'BORDER-LEFT:'+lStyle+' solid '+lColor+';';
                    }
    
                    borderBottom = border.borderBottom||'',bStyle=1,bColor = '#000';    //border-bottom
                    if(borderBottom){
                        bStyle = (borderBottom.style==1)?'thin':(borderBottom.style==2)?'medium':(borderBottom.style==5)?'thick':'thin';
                        bColor =  borderBottom.color||'#000';
                        borderBottom = 'BORDER-BOTTOM:'+bStyle+' solid '+bColor+';';
                    }
    
                    borderTop = border.borderTop||'',tStyle=1,tColor = '#000';    //border-top
                    if(borderTop){
                        tStyle = (borderTop.style==1)?'thin':(borderTop.style==2)?'medium':(borderTop.style==5)?'thick':'thin';
                        tColor =  borderTop.color||'#000';
                        borderTop = 'BORDER-TOP:'+tStyle+' solid '+tColor+';';
                    }
                }

                var format = this.getCellFormat(r,c,r,c,i)||''  //MSO-NUMBER-FORMAT
                if(format){
                    var newFormat = this.jsonFormat2XmlFormat(r,c,format,i)
                    format = newFormat;
                    format = 'MSO-NUMBER-FORMAT:'+format+';'
                }

                var tdStyle = hAlign+vAlign+fontName+fontItailc+fontBold+fontSize+fontColor+bgColor+text_decoration+word_break+
                borderRight+borderLeft+borderBottom+borderTop+format

                if(tdStyle){
                    cell.setAttribute('style',tdStyle)
                }
            }
        };

        //删除因合并产生多余的项
        for(var s = trEle.length-1;s>=1;s--){
            var tdEle = trEle[s].getElementsByTagName('TD')||[]
            for(var j = tdEle.length-1;j>=1;j--){
                spans.forEach(function(item){
                    if((item.row==s-1&&j-1>item.col&&j-1<item.col+item.colCount)||(item.col==j-1&&s-1>item.row&&s-1<item.row+item.rowCount)||
                        (s-1>item.row&&s-1<item.row+item.rowCount&&j-1>item.col&&j-1<item.col+item.colCount)){
                        trEle[s].removeChild(tdEle[j])
                    }
                }); 
            }
        }

        if(typeof(callback)=='function'){
            callback(xmlDoc,activeSheet.sheetName)
        } 

    }
   
}

Workbook.prototype.excel2Json = function(f,callback){
    this.workbook = {};
    this.workbook = this.initWorkbook()
    var _this = this,reSheet = /^(xl\/worksheets\/sheet)[0-9]+(\.xml)$/;
    var sArr = [],cellStyleArr = [],re =  /^\-[0-9]+\.?[0-9]*$|^[0-9]+\.?[0-9]*$/,m = {},themeXml,stylesXml;
    JSZip.loadAsync(f).then(function(zip){
        var task=[];
        zip.forEach(function(relativePath){
            task.push(relativePath);
        })
        doIt(0,function(){
            if(m["xl/theme/theme1.xml"]){
                themeXml = _this.string2XML(m["xl/theme/theme1.xml"])
            };
            if(m["xl/styles.xml"]){
                stylesXml = _this.string2XML(m["xl/styles.xml"]);
            }
            if(m["xl/workbook.xml"]){
                var workbookXmlData =  _this.string2XML(m["xl/workbook.xml"]);
                //tabs栏表格
                var sheetsTag = workbookXmlData.getElementsByTagName("sheets")[0],sheetsChild = sheetsTag.childNodes;
                _this.workbook.sheetList = [];
                _this.workbook.sheets = {}
                for (i = 0; i < sheetsChild.length ;i++) {
                    _this.workbook.showTabs = 1
                    //sheetList
                    var sheetName = sheetsChild[i].getAttribute("name")
                    _this.workbook.sheetList.push(sheetName)    
                    var sheetId = sheetsChild[i].getAttribute("sheetId")
                    _this.workbook.sheets[sheetName]= {}
                    ////sheets
                    _this.initSheet(sheetName,i); 
                    _this.workbook.sheets[sheetName]["sheetId"] = sheetId;  //临时存下的id
                }
                //width height
                _this.setTableSize(window.innerWidth-_this.scrollbarWidth,window.innerHeight-_this.scrollbarWidth)
                //活动表索引 activeSheet
                var workbookView = workbookXmlData.getElementsByTagName("workbookView")[0]
                if(workbookView){
                    var activeTab = workbookView.getAttribute("activeTab")||0
                    _this.workbook.activeSheet = parseInt(activeTab);
                    //水平垂直滚动条是否显示
                    var showVerticalScroll = (workbookView.getAttribute('showVerticalScroll')=='0')?0:1;
                    var showHorizontalScroll = (workbookView.getAttribute('showHorizontalScroll')=='0')?0:1;
                    _this.workbook.showVScrollBar = showVerticalScroll;
                    _this.workbook.showHScrollBar = showHorizontalScroll;
                }

                //默认字体
                var styleFonts = stylesXml.getElementsByTagName('fonts')[0];
                if(styleFonts){
                    var styleFontTag = styleFonts.getElementsByTagName('font');
                    if(styleFontTag&&styleFontTag.length>=1){
                        var defSzTag = styleFontTag[0].getElementsByTagName('sz')[0];
                        var defColorTag = styleFontTag[0].getElementsByTagName('color')[0];
                        var defNameTag = styleFontTag[0].getElementsByTagName('name')[0];
                        var defSize,defColor,defName;
                        if(defSzTag){ 
                            var defSize = defSzTag.getAttribute('val');
                            if(defSize){    //excel字体单位是磅
                                defSize = _this.pt2px(defSize)
                                _this.workbook.defaultFontSize = defSize;
                            }
                        };
                        if(defColorTag){
                            var defColorRgb = defColorTag.getAttribute('rgb');
                            var defColorTheme = defColorTag.getAttribute('theme');
                            var defColorTint = defColorTag.getAttribute('tint')||0;
                            if(defColorRgb){
                                defColor = '#'+defColorRgb.slice(2);
                            }else if(defColorTheme){
                                defColor = _this.getExcelThemeColor(themeXml,Number(defColorTheme),Number(defColorTint))
                            }
                            if(defColor){
                                _this.defaultCellStyle={"fontColor":defColor}
                            }
                        };
                        if(defNameTag){
                            var defNameVal = defNameTag.getAttribute('val');
                            if(defNameVal){
                                defName = defNameVal;
                                _this.workbook.defaultFontName = defName
                            } 
                        };
                        _this.defaultCellStyle={"font":_this.workbook.defaultFontSize+"px/"+_this.workbook.defaultLineHeight+"px "+ _this.workbook.defaultFontName}  
                    }
                };

            };
            //worksheets 下的所有sheet
            for(var key in m){
                if(reSheet.test(key)){  
                    var extInd =  key.lastIndexOf('.');
                    var splitInd = key.lastIndexOf('t')
                    var id = key.substring(splitInd+1,extInd);
                    for(var i = 0;i<_this.workbook.sheetList.length;i++){
                        if(id == _this.workbook.sheets[_this.workbook.sheetList[i]]["sheetId"]){
                            var activeSheet =  _this.workbook.sheets[_this.workbook.sheetList[i]];
                            var activeSheetXmlData = _this.string2XML(m[key])
                            delete activeSheet.sheetId  //使用完删除该属性
                            //每一个表的xml转json
                            //showGridLines
                            var sheetViewTag = activeSheetXmlData.getElementsByTagName('sheetView')[0];
                            if(sheetViewTag){
                                var showGridLines = sheetViewTag.getAttribute('showGridLines')
                                activeSheet.showGridLines = (showGridLines==0)?false:true;
                                //列头行头显示状态
                                var showRowColHeaders = sheetViewTag.getAttribute('showRowColHeaders');
                                activeSheet.rowHeaderData.showRowHeading = (showRowColHeaders==0)?false:true;
                                activeSheet.colHeaderData.showColHeading = (showRowColHeaders==0)?false:true;
                            }
                            //spans
                            var mergeCellsTag = activeSheetXmlData.getElementsByTagName('mergeCells')[0];
                            if(mergeCellsTag){
                                var mergeCellsTagChild = mergeCellsTag.childNodes;
                                for(var a = 0;a<mergeCellsTagChild.length;a++){
                                    var ref = mergeCellsTagChild[a].getAttribute('ref').split(':')
                                    var ref1 = ref[0];
                                    var ref2 = ref[1];
                                    var row = Number(ref1.match(/[0-9]+$/g).join(''))-1;    //excel 从1开始 json定义从0开始
                                    var col = _this.excelLetter2Num(ref1.match(/^[A-Z]+/g).join(''))-1;
                                    var rowCount = Number(ref2.match(/[0-9]+$/g).join(''))-row;
                                    var colCount = _this.excelLetter2Num(ref2.match(/^[A-Z]+/g).join(''))-col;
                                    activeSheet.spans.push({"row":row,"rowCount":rowCount,"col":col,"colCount":colCount})
                                }
                            };
                            //rows  columns
                            var sheetFormatPrTag = activeSheetXmlData.getElementsByTagName('sheetFormatPr')[0];
                            var defaultRowHeight = 14,defaultColWidth = 8.09;
                            if(sheetFormatPrTag){
                                defaultRowHeight = Math.floor(sheetFormatPrTag.getAttribute('defaultRowHeight'))||14;
                                defaultColWidth = Math.floor(sheetFormatPrTag.getAttribute('defaultColWidth'))||8.09;
                            }
                            //一英寸 = 72磅  excel行高单位为磅 列宽单位为0.1英寸
                            activeSheet.defaults.rowHeight = _this.pt2px(defaultRowHeight)
                            _this.workbook.defaultLineHeight = _this.pt2px(defaultRowHeight)
                            activeSheet.defaults.colWidth = Math.floor(defaultColWidth/10*_this.dpi)
                            _this.defaultCellStyle={"font":_this.workbook.defaultFontSize+"px/"+_this.workbook.defaultLineHeight+"px "+ _this.workbook.defaultFontName}  
                            var rowsTag = activeSheetXmlData.getElementsByTagName('row')
                            var cNumArr = []
                            if(rowsTag){
                                var lastRow = parseInt(rowsTag[rowsTag.length-1].getAttribute('r'));
                                _this.maxRow(lastRow,i)      
                                for(var b = 0;b<lastRow;b++){
                                    if(rowsTag[b]){
                                        var cTag = rowsTag[b].getElementsByTagName('c');
                                        if(cTag){
                                            for(var c = 0;c<cTag.length;c++){
                                                if(cTag[c]){
                                                    var cTagR = cTag[c].getAttribute('r');
                                                    if(cTagR){
                                                        cNumArr.push(_this.excelLetter2Num(cTagR.match(/^[A-Z]+/g).join(''))-1)
                                                    }
                                                }
                                            }
                                        }
                                        var ht = rowsTag[b].getAttribute('ht');
                                        var h = (ht)? _this.pt2px(ht):activeSheet.defaults.rowHeight;       
                                        _this.setRowHeight(h,b,b,i)                          
                                    };
                                };
                            }
                            var colTag = activeSheetXmlData.getElementsByTagName('col')
                            var colWidthArr = []
                            if(colTag){
                                for(var j = 0;j<colTag.length;j++){
                                    var width = Number(colTag[j].getAttribute('width'))-1;
                                    var max =  Number(colTag[j].getAttribute('max'))-1;
                                    var min =  Number(colTag[j].getAttribute('min'))-1;
                                    for(var k = min;k<=max;k++){
                                        colWidthArr.push({"c":k,'width':width})
                                    }
                                }
                            };
                            var lastCol = 0
                            if(cNumArr.length>=1){
                                lastCol = Math.max.apply(null,cNumArr);
                                _this.maxCol(lastCol+1,i)
                                for(var k = 0;k<=lastCol;k++){
                                    var w = activeSheet.defaults.colWidth
                                    for(var j = 0;j<colWidthArr.length;j++){
                                        if(k==colWidthArr[j].c){
                                            w = Math.floor(colWidthArr[j].width/10*_this.dpi)
                                        }
                                    }
                                    _this.setColWidth(w,k,k,i)
                                }
                            }; 
                            //fixeds 冻结
                            var paneTag = activeSheetXmlData.getElementsByTagName('pane')[0];
                            if(paneTag){
                                var state = paneTag.getAttribute('state');
                                if(state=='frozen'){
                                    var topLeftCell = paneTag.getAttribute('topLeftCell');
                                    var fixedRows = Number(topLeftCell.match(/[0-9]+$/g).join(''))-1
                                    var fixedCols = _this.excelLetter2Num(topLeftCell.match(/^[A-Z]+/g).join(''))-1;
                                    var ySplit = paneTag.getAttribute('ySplit');
                                    var xSplit = paneTag.getAttribute('xSplit');
                                    var fixedRow = fixedRows-Number(ySplit);
                                    var fixedCol = fixedCols-Number(xSplit);
                                    _this.fixedRowAndCol(fixedRows,fixedCols,fixedRow,fixedCol,i)
                                }
                            }
                            //activeRow activeCol rangeRow1 rangeCol1 rangeRow2 rangeCol2
                            var selectionTag = activeSheetXmlData.getElementsByTagName('selection')[0];
                            if(selectionTag){
                                var activeCell = selectionTag.getAttribute('activeCell');
                                var activeCol=0,activeRow=0
                                if(activeCell){
                                    activeCol = _this.excelLetter2Num(activeCell.match(/^[A-Z]+/g).join(''))-1;
                                    activeRow = Number(activeCell.match(/[0-9]+$/g).join(''))-1
                                }
                                _this.row(activeRow,i);
                                _this.col(activeCol,i)
                            }       
                            //value text formula
                            var ctag = activeSheetXmlData.getElementsByTagName('c');
                            if(ctag){
                                for(var s = 0;s<ctag.length;s++){
                                    var cellR,cellC,value,formula
                                    var cellRAttr = ctag[s].getAttribute('r');
                                    var cellTAttr = ctag[s].getAttribute('t');
                                    var cellSAttr = ctag[s].getAttribute('s');
                                    if(cellRAttr){
                                        cellR = Number(cellRAttr.match(/[0-9]+$/g).join(''))-1;
                                        cellC = _this.excelLetter2Num(cellRAttr.match(/^[A-Z]+/g).join(''))-1;
                                    };
                                    var vTag = ctag[s].getElementsByTagName('v')[0];
                                    var fTag = ctag[s].getElementsByTagName('f')[0];
                                    if(vTag){
                                        value = vTag.firstChild.nodeValue;
                                        if(cellR!==undefined&&cellC!==undefined&&value!==undefined&&cellTAttr!='s'){
                                            if(re.test(value)) value = Number(value);
                                            _this.setValue(cellR,cellC,value,i)
                                        }
                                        if(cellTAttr=='s'){
                                            sArr.push({'r':cellR,'c':cellC,'v':value,'i':i})
                                        }
                                    };
                                    if(fTag){
                                        formula = fTag.firstChild.nodeValue;
                                        if(formula){
                                            _this.setCellFormula(cellR,cellC,'='+formula,i)
                                        }
                                    }
                                    if(cellSAttr){
                                        cellStyleArr.push({'r':cellR,'c':cellC,'s':cellSAttr,'i':i})
                                    }  
                                }
                            }
                            //sheetProtection 
                            var sheetProtectionTag = activeSheetXmlData.getElementsByTagName('sheetProtection')[0];
                            var rows = activeSheet.rows,cols = activeSheet.columns;
                            if(sheetProtectionTag){
                                var selectLockedCells = sheetProtectionTag.getAttribute('selectLockedCells')
                                if(selectLockedCells=='1'){  //1 不能选中锁定单元格
                                    _this.setLock(0,0,rows.length-1,cols.length-1,true,i)
                                }else{
                                    _this.setCanEdit(0,0,rows.length-1,cols.length-1,false,i)
                                }
                            }
                        }
                    };
                };
            }
            //v值的文本对照
            if(m["xl/sharedStrings.xml"]){
                var sharedStringXml = _this.string2XML(m["xl/sharedStrings.xml"]);
                var tTag = sharedStringXml.getElementsByTagName('t');
                if(tTag){
                    for(var e = 0;e<sArr.length;e++){
                        var sValue = tTag[Number(sArr[e].v)].firstChild.nodeValue;
                        _this.setValue(sArr[e].r,sArr[e].c,sValue,sArr[e].i)
                    }
                }        
            };
            //单元格样式
            if(m["xl/styles.xml"]){
                var cellXfs = stylesXml.getElementsByTagName('cellXfs')[0];
                var fills = stylesXml.getElementsByTagName('fills')[0];
                var borders = stylesXml.getElementsByTagName('borders')[0];
                var fonts = stylesXml.getElementsByTagName('fonts')[0];
                var numFmts = stylesXml.getElementsByTagName('numFmts')[0];
                if(cellXfs){
                    var xfTag = cellXfs.getElementsByTagName('xf');
                    if(xfTag){
                        for(var j = 0;j<cellStyleArr.length;j++){
                            var cellXf = xfTag[Number(cellStyleArr[j].s)],cellStyleR = cellStyleArr[j].r,
                            cellStyleC = cellStyleArr[j].c,cellStyleI = cellStyleArr[j].i;
                            //对齐方式
                            var alignment = cellXf.getElementsByTagName('alignment')[0],
                                protection = cellXf.getElementsByTagName('protection')[0],
                                applyAlignment = cellXf.getAttribute('applyAlignment'),
                                applyFill = cellXf.getAttribute('applyFill'),
                                applyBorder = cellXf.getAttribute('applyBorder'),
                                applyFont = cellXf.getAttribute('applyFont'),
                                applyNumberFormat = cellXf.getAttribute('applyNumberFormat'),
                                applyProtection = cellXf.getAttribute('applyProtection'),
                                fillId = cellXf.getAttribute('fillId'),
                                borderId = cellXf.getAttribute('borderId'),
                                fontId = cellXf.getAttribute('fontId')||'0',
                                numFmtId = cellXf.getAttribute('numFmtId');
                            var horizontalAttr,verticalAttr,wrapText
                            //excel单元格默认底部对齐
                            if(alignment&&applyAlignment=='1'){
                                horizontalAttr = alignment.getAttribute('horizontal');
                                verticalAttr = alignment.getAttribute('vertical');   
                                wrapText = alignment.getAttribute('wrapText');   
                                horizontalAttr = (horizontalAttr=='left')?2:(horizontalAttr=='center')?3:(horizontalAttr=='right')?4:1;
                                verticalAttr = (verticalAttr=='top')?1:(verticalAttr=='center')?2:3;
                                if(wrapText=='1'){
                                    _this.wordWrap(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                }
                            }else{
                                horizontalAttr = 1;
                                verticalAttr = 3
                            };
                            _this.setHAlignment(cellStyleR,cellStyleC,cellStyleR,cellStyleC,horizontalAttr,cellStyleI);
                            _this.setVAlignment(cellStyleR,cellStyleC,cellStyleR,cellStyleC,verticalAttr,cellStyleI);
                            //!lock
                            if(protection&&applyProtection=='1'){
                                var lockAttr = protection.getAttribute('locked');
                                if(lockAttr=='0'){
                                    _this.setLock(cellStyleR,cellStyleC,cellStyleR,cellStyleC,false,cellStyleI)
                                    _this.setCanEdit(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                }
                            }
                            //填充色
                            if(applyFill=='1'&&Number(fillId)>=2&&fills){
                                var fillTag = fills.getElementsByTagName('fill');
                                if(fillTag){
                                    for(var k = 0;k<fillTag.length;k++){
                                        if(Number(fillId)==k){
                                            var fgColorTag = fillTag[k].getElementsByTagName('fgColor')[0]
                                            if(fgColorTag){
                                                var rgb = fgColorTag.getAttribute('rgb');
                                                var theme = fgColorTag.getAttribute('theme');
                                                var tint = fgColorTag.getAttribute('tint')||0;
                                                var fillColor
                                                if(rgb){
                                                    fillColor = '#'+rgb.slice(2);
                                                }else if(theme){
                                                    fillColor = _this.getExcelThemeColor(themeXml,Number(theme),Number(tint))
                                                }
                                                _this.setFillColor(fillColor,cellStyleR,cellStyleC,cellStyleR,cellStyleC,cellStyleI)
                                            }
                                        }
                                    }
                                }
                            }
                            //border
                            if(applyBorder=='1'&&Number(borderId)>=1&&borders){
                                var borderTag = borders.getElementsByTagName('border');
                                if(borderTag){
                                    for(var k = 0;k<borderTag.length;k++){
                                        if(Number(borderId)==k){
                                            var leftTag = borderTag[k].getElementsByTagName('left')[0];
                                            var rightTag = borderTag[k].getElementsByTagName('right')[0];
                                            var topTag = borderTag[k].getElementsByTagName('top')[0];
                                            var bottomTag = borderTag[k].getElementsByTagName('bottom')[0];
                                            if(leftTag){
                                                _this.excelBorder2canvas(themeXml,leftTag,'left',cellStyleR,cellStyleC,cellStyleI);
                                            };
                                            if(rightTag){
                                                _this.excelBorder2canvas(themeXml,rightTag,'right',cellStyleR,cellStyleC,cellStyleI);
                                            }
                                            if(topTag){
                                                _this.excelBorder2canvas(themeXml,topTag,'top',cellStyleR,cellStyleC,cellStyleI);
                                            }
                                            if(bottomTag){
                                                _this.excelBorder2canvas(themeXml,bottomTag,'bottom',cellStyleR,cellStyleC,cellStyleI);
                                            }
                                        }
                                    }
                                }
                            }
                            //font
                            if(fonts){
                                var fontTag = fonts.getElementsByTagName('font');
                                if(fontTag&&applyFont=='1'){
                                    for(var k = 0;k<fontTag.length;k++){
                                        if(Number(fontId)==k){
                                            var fontSzTag = fontTag[k].getElementsByTagName('sz')[0],
                                                fontColorTag = fontTag[k].getElementsByTagName('color')[0],
                                                fontNameTag = fontTag[k].getElementsByTagName('name')[0],
                                                fontBTag = fontTag[k].getElementsByTagName('b')[0],
                                                fontITag = fontTag[k].getElementsByTagName('i')[0],
                                                fontStrikeTag = fontTag[k].getElementsByTagName('strike')[0],
                                                fontUTag = fontTag[k].getElementsByTagName('u')[0];
                                            var fontSize,fontColor,fontName;
                                            if(fontSzTag){ 
                                                var fontVal = fontSzTag.getAttribute('val');
                                                if(fontVal){    //excel字体单位是磅
                                                    fontSize = _this.pt2px(fontVal);
                                                }
                                            };
                                            if(fontColorTag){
                                                var fontColorRgb = fontColorTag.getAttribute('rgb');
                                                var fontColorTheme = fontColorTag.getAttribute('theme');
                                                var fontColorTint = fontColorTag.getAttribute('tint')||0;
                                                if(fontColorRgb){
                                                    fontColor = '#'+fontColorRgb.slice(2);
                                                }else if(fontColorTheme){
                                                    fontColor = _this.getExcelThemeColor(themeXml,Number(fontColorTheme),Number(fontColorTint))
                                                }
                                            };
                                            if(fontNameTag){
                                                var fontNameVal = fontNameTag.getAttribute('val');
                                                if(fontNameVal){
                                                    fontName = fontNameVal;
                                                } 
                                            };
                                            if(fontBTag){
                                                _this.setFontBold(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                            }    
                                            if(fontITag){
                                                _this.setFontItalic(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                            };
                                            if(fontStrikeTag){
                                                _this.fontStrikeout(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                            };
                                            if(fontUTag){
                                                _this.fontUnderline(cellStyleR,cellStyleC,cellStyleR,cellStyleC,true,cellStyleI)
                                            }
                                            if(fontColor){
                                                _this.fontColor(cellStyleR,cellStyleC,cellStyleR,cellStyleC,fontColor,cellStyleI)
                                            };
                                            if(fontName){
                                                _this.setFontName(cellStyleR,cellStyleC,cellStyleR,cellStyleC,fontName,cellStyleI)
                                            };
                                            if(fontSize){
                                                _this.setFontSize(cellStyleR,cellStyleC,cellStyleR,cellStyleC,fontSize,cellStyleI)
                                            }
                                        }
                                    }
                                }
                            }
                            //format
                            if(applyNumberFormat=='1'&&numFmtId>=1&&numFmts){
                                var numFmtTag = numFmts.getElementsByTagName('numFmt');
                                if(numFmtTag){
                                    for(var k = 0;k<numFmtTag.length;k++){
                                        var fmtId = numFmtTag[k].getAttribute('numFmtId');
                                        if(fmtId==numFmtId){
                                            var formatCode = numFmtTag[k].getAttribute('formatCode');
                                            _this.xmlFormat2JsonFormat(cellStyleR,cellStyleC,formatCode,cellStyleI)                                            
                                        }
                                    }
                                }
                            }
        
                        }
                    }
                }
        
                
            };
            callback();
        })

        function doIt(idx,cb){
            if(task[idx]&&idx<task.length){
                var relativePath = task[idx];
                zip.file(relativePath).async("string").then(function (data){
                    m[relativePath] = data;
                    idx++;
                    doIt(idx,cb);
                    return true;
                })
            }else{
                cb()
            }
        }
    })
}

//获取excel主题色
Workbook.prototype.getExcelThemeColor = function(theme1Xml,theme,tint){
    var clrScheme = theme1Xml.getElementsByTagName('a:clrScheme')[0]
    if(clrScheme){
        var clrSchemeChild = clrScheme.childNodes;
        if(clrSchemeChild&&clrSchemeChild.length>=1){
            //0=1, 1=0, 2=3, 3=2 switch
            var t = theme;
            if(theme==0) t = 1;
            if(theme==1) t = 0;
            if(theme==2) t = 3;
            if(theme==3) t = 2;
            var clrSchemeItem = clrSchemeChild[t];
            var clrSchemeItemChild = clrSchemeItem.childNodes[0];
            if(clrSchemeItemChild){
                var colorVal
                if(t==0||t==1){
                    colorVal ='#'+clrSchemeItemChild.getAttribute('lastClr');
                }else{
                    colorVal = '#'+clrSchemeItemChild.getAttribute('val');
                }
                //处理主题色深色淡色
                if(tint!=0){ 
                    var rgbColor = this.colorHex2Rgb(colorVal)
                    var colorArr = rgbColor.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
                    var rColor = Math.floor(Number(colorArr[0])*(1+tint));
                    var gColor = Math.floor(Number(colorArr[1])*(1+tint));
                    var bColor = Math.floor(Number(colorArr[2])*(1+tint));
                    if(rColor>255) rColor = 255;
                    if(gColor>255) gColor = 255;
                    if(bColor>255) bColor = 255;
                    var newColor = 'rgb('+rColor+','+gColor+','+bColor+')';
                    colorVal = newColor
                }
                return colorVal
            }
        }
    }
}

//excel border -> canvas border
Workbook.prototype.excelBorder2canvas = function(themeXml,tag,dir,r,c,i){
    var style = tag.getAttribute('style');
    style = (style=='thin')?1:(style=='meduim')?2:(style=='thick')?5:1;
    var colorTag = tag.getElementsByTagName('color')[0];
    var color = '#000'
    if(colorTag){
        var colorRgb =  colorTag.getAttribute('rgb');
        var colorTheme = colorTag.getAttribute('theme');
        var colorTint = colorTag.getAttribute('tint')||0;
        if(colorRgb){
            color = '#'+colorRgb.slice(2);
        }else if(colorTheme){
            color = this.getExcelThemeColor(themeXml,Number(colorTheme),Number(colorTint))
        }
        if(dir=='left'){
            this.setLeftBorder(r,c,r,c,style,color,i)
        };
        if(dir=='right'){
            this.setRightBorder(r,c,r,c,style,color,i)
        }
        if(dir=='top'){
            this.setTopBorder(r,c,r,c,style,color,i)
        }
        if(dir=='bottom'){
            this.setBottomBorder(r,c,r,c,style,color,i)
        }
    }
}

//json单元格格式转xml单元格格式
Workbook.prototype.jsonFormat2XmlFormat = function(r,c,format,i){
    var newFormat = format;
    var value = Number(this.getValue(r,c,r,c,i))
    if(format=='yyyy年m月d日'){
        newFormat = 'yyyy\"年\"m\"月\"d\"日\"'
    };
    if(format=='yyyy年m月'){
        newFormat = 'yyyy\"年\"m\"月\"'
    };
    if(format=='m月d日'){
        newFormat = 'm\"月\"d\"日\"'
    };
    if(format=='星期n'){
        newFormat = '\[\$-804\]aaaa'
    };
    if(format=='m/d'){
        newFormat = 'm\/d'
    };
    var perRe = /^0*(P\+%)$/
    if(perRe.test(format)){
        var pIndex = format.indexOf('P');
        var zeroNum = format.slice(0,pIndex);
        var zeroLength = 0;
        if(zeroNum) zeroLength = zeroNum.length;
        newFormat = '0.'+Array(zeroLength+1).join('0')+'%'
    };
    var numRe = /^0*NR?\(?A?(\+thousands)?/
    if(numRe.test(format)){
        var nIndex = format.indexOf('N');
        var nZero = format.slice(0,nIndex);
        var nZeroLen = 0;
        if(nZero)   nZeroLen = nZero.length;
        var nZeroStr = '.'+Array(nZeroLen+1).join('0')
        var thousands = (format.indexOf('+thousands')!=-1)?'#,##':'';
        if(value>=0){
            newFormat = thousands+'0'+nZeroStr+'_'
        }else{
            var red = (format.indexOf('R')!=-1)?'[Red]':'';
            var brackets = format.indexOf('(');
            var abs = format.indexOf('A')
            if(brackets!=-1){   //有括号
                newFormat = thousands+'0'+nZeroStr+'_)\;'+red+'\\('+thousands+'0'+nZeroStr+'\\)';
            }else{
                if(abs!=-1){
                    newFormat = thousands+'0'+nZeroStr+'\;'+red+thousands+'0'+nZeroStr
                }else{
                    if(red){
                        newFormat = thousands+'0'+nZeroStr+'_\;'+red+'\\'+'-'+'0'+nZeroStr+'\\'
                    }else{
                        newFormat = thousands+'0'+nZeroStr+'_'
                    }
                }
            }
        }
    };
    var menoyRe = /^0*MR?\(?A?(\+[¥|$|€|￡|₣|₩])?/;
    if(menoyRe.test(format)){
        var mIndex = format.indexOf('M');
        var mZero = format.slice(0,mIndex);
        var mZeroLen = 0;
        if(mZero)   mZeroLen = mZero.length;
        var mZeroStr = '.'+Array(mZeroLen+1).join('0')
        var mSign = format.match(/[¥|$|€|￡|₣|₩]/g)||'';
        if(mSign)   mSign = mSign[0];
        var mBrackets = format.indexOf('(');
        var mAbs = format.indexOf('A')
        var mRed = (format.indexOf('R')!=-1)?'[Red]':'';
        if(mBrackets!=-1){
            newFormat = '\"'+mSign+'\"'+'#,##0'+mZeroStr+'_)\;'+mRed+'\\('+'\"'+mSign+'\"'+'#,##0'+mZeroStr+'\\)'
        }else{
            if(value<0&&mRed&&mAbs!=-1){
                newFormat = '\"'+mSign+'\"'+'#,##0'+mZeroStr+'\;'+mRed+'\"'+mSign+'\"'+'#,##0'+mZeroStr;
            }else if(value>=0||(value<0&&mAbs==-1)){
                newFormat = '\"'+mSign+'\"'+'#,##0'+mZeroStr+'\;'+mRed+'\"'+mSign+'\"'+'-#,##0'+mZeroStr;
            }
        }
    }
    return newFormat
}

//xml单元格格式转json单元格格式
Workbook.prototype.xmlFormat2JsonFormat = function(r,c,format,i){
    //excel日期相关格式后面带;@
    var ymdRe = /^(yyyy\"年\"m\"月\"d\"日\")(;@)?$/;
    var ymRe = /^(yyyy\"年\"m\"月\")(;@)?$/;
    var mdRe = /^(m\"月\"d\"日\")(;@)?$/;
    var mdSlashRe = /^(m\/d)(;@)?$/;
    var dayRe = /^(\[\$-804\]aaaa)(;@)?$/;
    var perRe = /^(0\.)0*%$/
    var numRe = /^((#,##)?0\.?0*)(_?\s*\)?;?(\[Red\])?\\?\(?-?(#,##)?0?\.?0*\\?\)?\s*)$/;
    var menoyRe = /^(("[¥|$|€|￡|₣|₩]")(#,##)?0\.?0*)(_?\s*\)?;?(\[Red\])?("[¥|$|€|￡|₣|₩]")?\\?\(?-?("[¥|$|€|￡|₣|₩]")?(#,##)?0?\.?0*\\?\)?\s*)$/;
    if(ymdRe.test(format)){ //yyyy年m月d日
        this.setCellFormat(r,c,r,c,'yyyy年m月d日',i)
        var ymdVCn = this.getValue(r,c,r,c,i)
        if(ymdVCn!==undefined&&ymdVCn!==null){
            var ymdTCn = this.excelDateToJSDate(ymdVCn,'/')
            if(ymdTCn)  this.setText(r,c,ymdTCn,i)
        }  
    };
    if(ymRe.test(format)){
        this.setCellFormat(r,c,r,c,'yyyy年m月',i)
        var ymVCn = this.getValue(r,c,r,c,i)
        if(ymVCn!==undefined&&ymVCn!==null){
            var ymTCn = this.excelDateToJSDate(ymVCn,'/')
            if(ymTCn)  this.setText(r,c,ymTCn,i)
        }
    };
    if(mdRe.test(format)){
        this.setCellFormat(r,c,r,c,'m月d日',i)
        var mdVCn = this.getValue(r,c,r,c,i)
        if(mdVCn!==undefined&&mdVCn!==null){
            var mdTCn = this.excelDateToJSDate(mdVCn,'/')
            if(mdTCn)  this.setText(r,c,mdTCn,i)
        }
    };
    if(numRe.test(format)){
        var numbV = this.getValue(r,c,r,c,i)
        var thousands = (format.indexOf('#,##')!=-1)?'+thousands':'';
        var red = (format.indexOf('[Red]')!=-1&&Number(numbV)<0&&numbV<0)?'R':'';
        var bracketsStr = format.match(/\(-?(#,##)?0?\.?0*\\?\)/g)
        var zeroStr = format.match(/0\.?0*/g)
        var zeroLen = 0;
        if(zeroStr) zeroStr = zeroStr[0];
        if(zeroStr){
            zeroLen = zeroStr.slice(zeroStr.indexOf('.')+1).length
        }
        if(bracketsStr) bracketsStr.join('')
        var brackets = (format.indexOf(bracketsStr)!=-1&&Number(numbV)<0&&numbV<0)?'(':'';
        var abs = ((format.indexOf('-')==-1&&Number(numbV)<0&&!((format.indexOf('-')==-1)&&/^((#,##)?0\.?0*)_\s*$/.test(format))))?'A':'';
        var numbType = Array(zeroLen+1).join('0')+'N'+red+brackets+abs+thousands;
        this.setCellFormat(r,c,r,c,numbType,i)
    }
    if(menoyRe.test(format)){
        var moneyV = this.getValue(r,c,r,c,i);
        var moneySymbol = format.slice(1,2)
        var moneyRed = (format.indexOf('[Red]')!=-1&&typeof(moneyV)=='number'&&moneyV<0)?'R':'';
        var moneyBracketsStr = format.match(/\(-?("[¥|$|€|￡|₣|₩]")?(#,##)?0?\.?0*\\?\)/g);
        if(moneyBracketsStr) moneyBracketsStr.join('');
        var moneyBrackets = (format.indexOf(moneyBracketsStr)!=-1&&typeof(moneyV)=='number'&&moneyV<0)?'(':'';
        var moneyAbs = (format.indexOf('-')==-1&&typeof(numbV)=='number'&&numbV<0)?'A':'';
        var moneyZeroStr = format.match(/0\.?0*/g)
        var moneyZeroLen = 0;
        if(moneyZeroStr) moneyZeroStr = moneyZeroStr[0];
        if(moneyZeroStr){
            moneyZeroLen = moneyZeroStr.slice(moneyZeroStr.indexOf('.')+1).length
        }
        var moneyType = Array(moneyZeroLen+1).join('0')+'M'+moneyRed+moneyBrackets+moneyAbs+'+'+moneySymbol
        this.setCellFormat(r,c,r,c,moneyType,i)
    }
    if(mdSlashRe.test(format)){    //m/d
        this.setCellFormat(r,c,r,c,'m/d',i)
        var mdV = this.getValue(r,c,r,c,i)
        if(mdV!==undefined&&mdV!==null){
            var mdT = this.excelDateToJSDate(mdV,'/')
            if(mdT)   this.setText(r,c,mdT,i)
        }
    };
    if(dayRe.test(format)){    //星期几
        this.setCellFormat(r,c,r,c,'星期n',i)
        var dayV = this.getValue(r,c,r,c,i)
        if(dayV!==undefined&&dayV!==null){
            var dayT = this.excelDateToJSDate(dayV,'/')
            if(dayT)  this.setText(r,c,dayT,i)
        }
    };
    if(perRe.test(format)){   //百分比
        var zeroStart = format.indexOf('.');
        var zeroEnd = format.indexOf('%');
        var zeroNum = format.slice(zeroStart+1,zeroEnd);
        var zeroLength = 0;
        if(zeroNum) zeroLength = zeroNum.length;
        var pType = Array(zeroLength+1).join('0')+'P+%';
        this.setCellFormat(r,c,r,c,pType,i)
    };  
}

//xml border转json border格式
Workbook.prototype.setChangeBorder = function(row,col,border,attr,index){
    if(border&&border.trim()!=''){
        var attr = attr.split('-'),border =  border.trim().split(/\s+/);
        var borderStyleRe = /(none|hidden|dotted|dashed|solid|double|groove|ridge|inset|outset|inherit)/g;
        var borderWeightRe = /^[0-9]*\.?[0-9]*(px|pt)?$|^(thin|medium|thick)$/g
        var borderStyle = 'solid';
        var borderWeight = 1;
        var borderColor = '#000';
        for(var i = 0;i<border.length;i++){
            if(!borderStyleRe.test(border[i])&&!borderWeightRe.test(border[i])){
                borderColor = border[i]
            };
            if(border[i].match(borderStyleRe)){
                borderStyle = border[i]
            };
            if(border[i].match(borderWeightRe)){
                borderWeight = border[i]
            }
        }
        borderWeight = (borderWeight=='thin')?1:(borderWeight=='medium')?2:(borderWeight=='thick')?5:borderWeight;
        if(/^\./.test(borderWeight)){
            borderWeight = '0'+borderWeight
        };
        if(/(pt)$/.test(borderWeight)){
            borderWeight = this.pt2px(borderWeight.replace(/pt/,''))||1
        };
        borderWeight = parseInt(borderWeight)
        borderWeight = (borderWeight>=5)?5:(borderWeight>=3&&borderWeight<=4)?2:1;
        if(attr[1]=='LEFT'||attr=='BORDER') this.setLeftBorder(row,col,row,col,borderWeight,borderColor,index)
        if(attr[1]=='RIGHT'||attr=='BORDER') this.setRightBorder(row,col,row,col,borderWeight,borderColor,index)
        if(attr[1]=='TOP'||attr=='BORDER') this.setTopBorder(row,col,row,col,borderWeight,borderColor,index)
        if(attr[1]=='BOTTOM'||attr=='BORDER') this.setBottomBorder(row,col,row,col,borderWeight,borderColor,index)
    }
}

//工具函数

//字符串转xml
Workbook.prototype.string2XML = function(xmlString){
    // for IE
    if (window.ActiveXObject) {
        var xmlobject = new ActiveXObject("Microsoft.XMLDOM");
        xmlobject.async = "false";
        xmlobject.loadXML(xmlString);
        return xmlobject;
    }
    // for other browsers
    else {
        var parser = new DOMParser();
        var xmlobject = parser.parseFromString(xmlString, "text/xml");
        return xmlobject;
    }
}

//xml转字符串
Workbook.prototype.xml2String = function(xmlObject){
 // for IE
    if (window.ActiveXObject) {      
        return xmlObject.xml;
    }
    // for other browsers
    else {       
        return (new XMLSerializer()).serializeToString(xmlObject);
    }
}

//创建xml文档
Workbook.prototype.createXmlDoc = function(){
    var xmldoc = null;
    if(window.ActiveXObject){
        xmldoc = new ActiveXObject('Microsoft.XMLDOM')
    }else if(document.implementation&&document.implementation.createDocument){
        xmldoc = document.implementation.createDocument('','',null)
    }
    return xmldoc
}

//获取xml style某个属性值
Workbook.prototype.getXmlStyleValue = function(ele,attr){
    // style="FONT-SIZE: 9pt;FONT-FAMILY: 微软雅黑;MSO-NUMBER-FORMAT: #,##0.00_);[Red]\(#,##0.00\);MSO-PROTECTION:UNLOCKED VISIBLE;"
    if(!ele) return
    var style = ele.getAttribute('style');
    var result
    if(style){
        result = style.split(attr+':')[1];
        if(result){
            result = result.split(':')[0].trim();
            var lastIndex = result.lastIndexOf(';');
            if(lastIndex!=-1){
                result = result.slice(0,lastIndex).trim()
            }
        }
    }    
    return result
}

Workbook.prototype.getXmlPageSetUpValue = function(pagesetup,attr){
    var result;
    if(pagesetup){
        result = pagesetup.split(attr+':')[1];
        if(result){
            result = result.split(',')[0].trim();
            result = result.match(/[^"']/g)
            if(result){
                result = result.join('')
            }
        }
    }
    return result
}

//磅转px
Workbook.prototype.pt2px = function(size){
    //一英寸等于72磅
    return Math.floor(Number(size)/72*this.dpi);
}

//px转磅
Workbook.prototype.px2pt = function(size){
    return (Number(size)/this.dpi*72).toFixed(2)-0
}

//阻止默认事件
Workbook.prototype.stopDefault = function(e){
    if(e&&e.preventDefault){
        e.preventDefault(); 
    }else{
        window.event.returnValue = false
    }
}

//阻止冒泡
Workbook.prototype.stopBubble = function(e){
    if(e&&e.stopPropagation){
        e.stopPropagation(); 
    }else{
        window.event.cancelBubble = true; 
    } 
}

//是否对象
Workbook.prototype.typeObj = function(obj){
    var type = Object.prototype.toString.call(obj);
    if(type=='[object Array]'){
      return 'Array';
    }else if(type=='[object Object]'){
      return 'Object';
    }else{
      return "obj is not object or array"
    }
}

//获取样式
Workbook.prototype.getStyle = function(obj,attr){
    if (obj.currentStyle){  //针对IE浏览器
        return obj.currentStyle[attr];
    }else{                 //针对Firefox
        return getComputedStyle(obj,false)[attr];  
    }
}

//节流函数  减少拖拽滚动过程的单元格绘制频率
Workbook.prototype.throttle = function(method,delay,duration){ 
    var timer = null, begin = new Date(),_this = this;
    return function(e){
       var isHasVBar = _this.isHasVBar();
        if(isHasVBar&&_this.isShowBar){
            _this.stopDefault(e);
            _this.stopBubble(e)
        }
        var context = this,args = arguments;
        var current = new Date();
        if(current-begin>=duration){    //当上一次执行的时间与当前的时间差大于设置的执行间隔时长的话，就主动执行一次
            method.apply(context,args)
            begin = current;
        }else {
            clearTimeout(timer);
            timer = setTimeout(function(){  //每一个delay时间只执行一次
                method.apply(context,args)
            },delay)
        }
    }
}

//解决文字模糊  获取设备像素比(IE10及以下不支持)
Workbook.prototype.getPixelRatio = function(){  
    var  backingStore = this.ctx.backingStorePixelRatio ||
        this.ctx.webkitBackingStorePixelRatio ||
        this.ctx.mozBackingStorePixelRatio ||
        this.ctx.msBackingStorePixelRatio ||
        this.ctx.oBackingStorePixelRatio ||
        this.ctx.backingStorePixelRatio || 1;
        return (window.devicePixelRatio || 1) / backingStore;      
}

//增加事件监听
Workbook.prototype.addEvent = function(element, type, callback){
    if(element.addEventListener){
        element.addEventListener(type, callback, false); 
    } else if(element.attachEvent){
        element.attachEvent('on' + type, callback);
    } else {
        element['on' + type] = callback;
    }
}

//移除事件监听
Workbook.prototype.removeEvent = function(element, type, callback){
    if(element.removeEventListener){
        element.removeEventListener(type, callback, false);
    }else if(element.detachEvent){
        element.detachEvent("on"+type,callback);
    }else{
        element["on"+type] = null;
    }
}

//列头数字转换成字母表示(3->C 27->AA)  num>=1
Workbook.prototype.num2ExcelLetter = function(num){
    var letterArr = [];
    var num2Letter = function(_num){
        var n = _num-1
        var a = parseInt(n/26)
        var b = n%26
        letterArr.push(String.fromCharCode(64+parseInt(b+1)));
        if(a>0){
            num2Letter(a)
        }
    }
    num2Letter(num)
    return letterArr.reverse().join('')
}

//列头字母转换成数字(C->3 AA->27)   
Workbook.prototype.excelLetter2Num = function(letter){
    var str = letter.toLowerCase().split("");
    var al = str.length;
    var getCharNumber = function (charx) {
        return charx.charCodeAt() - 96;
    };
    var numout = 0;
    var charnum = 0;
    for (var i = 0; i < al; i++) {
        charnum = getCharNumber(str[i]);
        numout += charnum * Math.pow(26, al - i - 1);
    };
    return numout;
}

//十六进制转rgb
Workbook.prototype.colorHex2Rgb = function(c){
    var reg = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
    if(reg.test(c)){
        var color = c.toLowerCase();
        if(color.length==4){
            var newColor = '#'
            for(var i = 1;i<4;i++){
                newColor += color.slice(i,i+1).concat(color.slice(i,i+1));
            }
            color = newColor;
        };
        var hexColorArr = [];
        for(var j = 1;j<7;j+=2){
            hexColorArr.push(parseInt(color.slice(j,j+2),16))   //把值当做16进制的值 转换为10进制
        }
        return 'rgb('+hexColorArr.join(',')+')'
    }else{
        return c
    }
}

//rgb转16
Workbook.prototype.colorRgb2Hex = function(color){
   // RGB颜色值的正则
  var reg = /^(rgb|RGB)/;
  if (reg.test(color)) {
    var strHex = "#";
    // 把RGB的3个数值变成数组
    var colorArr = color.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
    // 转成16进制
    for (var i = 0; i < colorArr.length; i++) {
      var hex = Number(colorArr[i]).toString(16);
      if (hex.length==1) {
        hex = '0'+hex;
      }
      strHex += hex;
    }
    return strHex;
  } else {
    return String(color);
  }
}

//滚动条w
Workbook.prototype.getScrollBarWidth = function(){
    var odiv = document.createElement('div'),//创建一个div
    styles = {
        width: '100px',
        height: '100px',
        overflowY: 'scroll'//让他有滚动条
    }, i, scrollbarWidth;
    for (i in styles) odiv.style[i] = styles[i];
    document.body.appendChild(odiv);//把div添加到body中
    scrollbarWidth = odiv.offsetWidth - odiv.clientWidth;//相减
    document.body.removeChild(odiv);//移除创建的div
    return scrollbarWidth;//返回滚动条宽度
}


//属性

/**
 * @api {null} /null checkBoxBorderColor
 * @apiName checkBoxBorderColor
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the border color of the check box button(Defaults:#C5C1AA)
 * - WB.checkBoxBorderColor="color"   
 */
Object.defineProperty(Workbook.prototype, "checkBoxBorderColor", {
    value: "#C5C1AA",
    writable:true    
});

/**
 * @api {null} /null checkBoxFillColor
 * @apiName checkBoxFillColor
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the fill color of the check box button(Defaults:#F0F0F0)
 * - WB.checkBoxFillColor="color"   
 */
Object.defineProperty(Workbook.prototype, "checkBoxFillColor", {
    value: "#F0F0F0",
    writable:true    
});


/**
 * @api {null} /null checkBoxBorderColorH
 * @apiName checkBoxBorderColorH
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the border color of the check box button (hover)(Defaults:rgba(30,144,255,.4))
 * - WB.checkBoxBorderColorH="color"   
 */
Object.defineProperty(Workbook.prototype, "checkBoxBorderColorH", {
    value: "rgba(30,144,255,.4)",
    writable:true    
});

/**
 * @api {null} /null checkBoxFillColorH
 * @apiName checkBoxFillColorH
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the fill color of the checkbox button (hover)(Defaults:rgba(240,248,255))
 * - WB.checkBoxFillColorH="color"   
 */
Object.defineProperty(Workbook.prototype, "checkBoxFillColorH", {
    value: "rgba(240,248,255)",
    writable:true    
});

/**
 * @api {null} /null editSelStart
 * @apiName editSelStart
 * @apiGroup Attribute:prototype
 * @apiDescription Set the start position of the selected text when triggering input(Defaults:0)
 * - WB.editSelStart=1   Start with the second word
 */
Object.defineProperty(Workbook.prototype, "editSelStart", {
    value: 0,
    writable:true    
});


/**
 * @api {null} /null editSelEnd
 * @apiName editSelEnd
 * @apiGroup Attribute:prototype
 * @apiDescription Set the end position of the selected text when triggering input(Defaults:0)
 * - WB.editSelEnd=3   End of the third word
 */
Object.defineProperty(Workbook.prototype, "editSelEnd", {
    value: 0,
    writable:true     
});

Object.defineProperty(Workbook.prototype, "document", {
    value: document,
    writable:true     
});

Object.defineProperty(Workbook.prototype, "checkBoxW", {
    value: 13,
    writable:false    
});

Object.defineProperty(Workbook.prototype, "checkBoxH", {
    value: 13,
    writable:false    
});

Object.defineProperty(Workbook.prototype, "numColW", { 
    get : function () {
        return (this.activeSheet.rowHeaderData.showRowHeading)?this.activeSheet.rowHeaderData.defaultW:0
    },
});

Object.defineProperty(Workbook.prototype, "numRowH", { 
    get : function () {
        return (this.activeSheet.colHeaderData.showColHeading)?this.activeSheet.colHeaderData.defaultH:0
    },
});


/**
 * @api {null} /null defaultCellStyle
 * @apiName defaultCellStyle
 * @apiGroup Attribute:prototype
 * @apiDescription Get or set the default property value of the cell. The value set is an object. The default value is used without changing a certain property. The default values are as follows
 * - "fontColor":"#000"
 * - "hAlign":1
 * - "cellType":{"name":0}
 * - "lock":false
 * - "vAlign":2
 * - "formatter":""
 * - "font":this.workbook.defaultFontSize+"px/"+this.workbook.defaultLineHeight+"px "+ this.workbook.defaultFontName
 * - "wordWrap":false
 * - WB.defaultCellStyle={"fontColor":"red"}  Set the default font color of the cell to red
 */
Object.defineProperty(Workbook.prototype, "defaultCellStyle", { 
    get : function () {
        return this._defaultCellStyle;
    },   
    set : function(value){
        this._defaultCellStyle.fontColor = value.fontColor || this._defaultCellStyle.fontColor;
        this._defaultCellStyle.hAlign = value.hAlign || this._defaultCellStyle.hAlign;
        this._defaultCellStyle.cellType = value.cellType || this._defaultCellStyle.cellType;
        this._defaultCellStyle.lock = value.lock || this._defaultCellStyle.lock;
        this._defaultCellStyle.vAlign = value.vAlign || this._defaultCellStyle.vAlign;
        this._defaultCellStyle.formatter = value.formatter || this._defaultCellStyle.formatter;
        this._defaultCellStyle.font = value.font || this._defaultCellStyle.font;
        this._defaultCellStyle.wordWrap = value.wordWrap || this._defaultCellStyle.wordWrap;
    }   
});

/**
 * @api {null} /null scrollSize
 * @apiName scrollSize
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the scroll bar size(Defaults:7)
 * - WB.scrollSize=8  The scroll bar size is 8px
 */
Object.defineProperty(Workbook.prototype, "scrollSize", {
    value: 7,
    writable:true    
});

/**
 * @api {null} /null bottomBtn
 * @apiName bottomBtn
 * @apiGroup Attribute:prototype
 * @apiDescription Show btn button group when showtabs display mode is set to 2
 * - WB.bottomBtn=['yes','cancel']  
 */
Object.defineProperty(Workbook.prototype, "bottomBtn", {
    value: [],
    writable:true    
});



/**
 * @api {null} /null tabsHeight
 * @apiName tabsHeight
 * @apiGroup Attribute:prototype
 * @apiDescription Set the height of the tabs bar
 * - WB.tabsHeight=30  The height of the tabs bar is 30px
 */
Object.defineProperty(Workbook.prototype, "tabsHeight", {
    get : function () {
        return (this.workbook.showTabs==0)?0:this._tabsHeight
    }, 
    set : function(value){
        this._tabsHeight = value;
    }  
});

/**
 * @api {null} /null tabArrowWidth
 * @apiName tabArrowWidth
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the arrow width of the tabs switch table(Defaults:5)
 * - WB.tabArrowWidth=5  Tabs bar switch table arrow width is 5px
 */
Object.defineProperty(Workbook.prototype, "tabArrowWidth", {
    value: 5,
    writable:true    
});


/**
 * @api {null} /null tabArrowHeight
 * @apiName tabArrowHeight
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the arrow height of the tabs switch table(Defaults:8)
 * - WB.tabArrowHeight=8  Tabs bar switch table arrow height is 8px
 */
Object.defineProperty(Workbook.prototype, "tabArrowHeight", {
    value: 8,
    writable:true    
});

/**
 * @api {null} /null tabArrowStartX
 * @apiName tabArrowStartX
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get the position of the tabs column switch table arrow (the first arrow) from the left side of the tabs column(Defaults:15)
 * - WB.tabArrowStartX=15  The position of the arrow of the tabs bar switching table is 15px from the left side of the tabs bar
 */
Object.defineProperty(Workbook.prototype, "tabArrowStartX", {
    value: 15,
    writable:true    
});

/**
 * @api {null} /null isEdit
 * @apiName isEdit
 * @apiGroup Attribute:prototype
 * @apiDescription Set whether table initialization enters edit state(Defaults:false)
 * - WB.isEdit=true  table initialization into edit
 */
Object.defineProperty(Workbook.prototype, "isEdit", {
    value: false,
    writable:true    
});

/**
 * @api {null} /null isFocus
 * @apiName isFocus
 * @apiGroup Attribute:prototype
 * @apiDescription Can judge whether it is in editing according to isFocus or isEdit(Defaults:false)
 */
Object.defineProperty(Workbook.prototype, "isFocus", {
    value: false,
    writable:true    
});

Object.defineProperty(Workbook.prototype, "clipBoard", {
    value: [],
    writable:true,
    enumerable:true,
});

Object.defineProperty(Workbook.prototype, "clipBoardTxt", {
    value: '',
    writable:true,
});

Object.defineProperty(Workbook.prototype, "cellBtnWidth", {
    value: 10,
});

Object.defineProperty(Workbook.prototype, "cellBtnHeight", {
    value: 10,
});

Object.defineProperty(Workbook.prototype, "cellBtnAreaWidth", {
    value: 13,
});

Object.defineProperty(Workbook.prototype, "isStopPainted", {
    get : function () {
        return this.workbook.stopPaintedCount==0?false:true;
    }
});

/**
 * @api {null} /null btnBorderColor
 * @apiName btnBorderColor
 * @apiGroup Attribute:prototype
 * @apiDescription The border color of the cell button (celltype ==1 celltype==2)(Defaults: #C5C1AA)
 * - WB.btnBorderColor = value(color)  
 */
Object.defineProperty(Workbook.prototype, "btnBorderColor", {
    value: "#C5C1AA",
    writable:true    
});

/**
 * @api {null} /null enterToRight
 * @apiName enterToRight
 * @apiGroup Attribute:prototype
 * @apiDescription Whether the cell moves to the right after pressing the Enter key (Defaults: false), when the property is changed to true, the carriage return moves the cell to the last column and jumps to the next row
 * - WB.enterToRight = value(bool)  
 */
Object.defineProperty(Workbook.prototype, "enterToRight", {
    value: false,
    writable:true    
});

/**
 * @api {null} /null btnFillColor
 * @apiName btnFillColor
 * @apiGroup Attribute:prototype
 * @apiDescription The fill color of the cell button (celltype ==1 celltype==2) (Defaults: #F0F0F0)
 * - WB.btnFillColor = value(color)  
 */
Object.defineProperty(Workbook.prototype, "btnFillColor", {
    value: "#F0F0F0",
    writable:true    
});

/**
 * @api {null} /null btnBorderColorH
 * @apiName btnBorderColorH
 * @apiGroup Attribute:prototype
 * @apiDescription The border color of the cell button (hover) (celltype ==1 celltype==2) (Defaults: rgba(30,144,255,.4))
 * - WB.btnBorderColorH = value(color)  
 */
Object.defineProperty(Workbook.prototype, "btnBorderColorH", {
    value: "rgba(30,144,255,.4)",
    writable:true    
});

/**
 * @api {null} /null btnFillColorH
 * @apiName btnFillColorH
 * @apiGroup Attribute:prototype
 * @apiDescription The fill color of the cell button (hover) (celltype==1 celltype==2) (Defaults: rgba(240,248,255))
 * - WB.btnFillColorH = value(color)  
 */
Object.defineProperty(Workbook.prototype, "btnFillColorH", {
    value: "rgba(240,248,255)",
    writable:true    
});

/**
 * @api {null} /null newSheetRow
 * @apiName newSheetRow
 * @apiGroup Attribute:prototype
 * @apiDescription Set the number of rows when creating a new table(Defaults: 8)
 * - WB.newSheetRow = value(Int)  
 */
Object.defineProperty(Workbook.prototype, "newSheetRow", {
    value: 8,
    writable:true    
});

/**
 * @api {null} /null newSheetCol
 * @apiName newSheetCol
 * @apiGroup Attribute:prototype
 * @apiDescription Set the number of columns when creating a new table(Defaults: 8)
 * - WB.newSheetCol = value(Int)  
 */
Object.defineProperty(Workbook.prototype, "newSheetCol", {
    value: 8,
    writable:true    
});

/**
 * @api {null} /null dpi
 * @apiName dpi
 * @apiGroup Attribute:prototype
 * @apiDescription Set or get dpi (Defaults 96)
 * - WB.dpi = 值(Int)  
 */
Object.defineProperty(Workbook.prototype, "dpi", {
    value: 96,
    writable:true    
});


Object.defineProperty(Workbook.prototype, "widthDiff", {
    get : function () {
        return window.screen.width - this.workbook.devScreenWidth 
    }, 
});

Object.defineProperty(Workbook.prototype, "heightDiff", {
    get : function () {
        return window.screen.height - this.workbook.devScreenHeight 
    },  
});

Object.defineProperty(Workbook.prototype, "width", {
    get : function () {
        return this.workbook.width 
    },  
});

Object.defineProperty(Workbook.prototype, "height", {
    get : function () {
        return this.workbook.height 
    },  
});













































