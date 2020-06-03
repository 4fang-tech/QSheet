var workbook = {
    "width":765,            
    "height":430,              
    "behaviorMode":1,
    "showTabs":1,                                      
    "sheetList":["出纳数据","请假单"],
    "sheets":{
        "出纳数据":{
            "sheetName":"出纳数据",
            "rows" : [{"size": 34,}, {"size": 49}, {"size": 30}, {"size": 34}, {"size": 34}, {"size": 34}, 
            {"size": 34}, {"size": 34}, {"size": 31}, {"size": 34}, {"size": 34}, {"size": 34}, {"size": 34}, {"size": 34}, {"size": 34}, {"size": 34}] ,//行高
            "columns" : [{"size": 38}, {"size": 90}, {"size": 100}, {"size": 17}, 
            {"size": 103}, {"size": 254}, {"size": 90,"name":"列头"},{"size": 90}, {"size": 90}, {"size": 90}],//列宽
            "spans": [
                {
                    "row":0,
                    "rowCount":1,
                    "col":3,
                    "colCount":3
                },
                {
                    "row":1,
                    "rowCount":1,
                    "col":1,
                    "colCount":5
                },
                {
                    "row":10,
                    "rowCount":1,
                    "col":2,
                    "colCount":4
                },
                {
                    "row":11,
                    "rowCount":1,
                    "col":2,
                    "colCount":4
                },
                {
                    "row":12,
                    "rowCount":2,
                    "col":0,
                    "colCount":7
                }, 
                {
                    "row":14,
                    "rowCount":2,
                    "col":0,
                    "colCount":5
                },
            ],
            "data":{
                "dataTable":{  
                   "1": {   //row
                        "1": {
                            "value":-1000,
                            "text":"-1000",
                            "style":{
                                "fontColor":"#666666",
                                "hAlign":3,
                                "vAlign":2,
                                "font":"14px/20px 楷体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":"00E"
                            },
                        },
                    },
                   "2": {   //row2
                        "1": {
                            "value":"编号",
                            "text":"编号",
                            "style":{
                                "fontColor":"#666666",
                                "hAlign":4,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":0},
                                "lock":true,
                                "formatter":"",
                                "bgColor":"pink"
                            }
                        },
                        "2":{
                            "value":1,
                            "text":"1",
                            "style":{
                                "hAlign":1,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":3},
                                "lock":false,
                                "canEdit":false,
                                "formatter":""
                            }
                        },
                        "3":{
                            "value":"",
                            "text":"",
                            "style":{
                                "hAlign":1,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            }
                        },
                        "4":{
                            "value":"收据编号",
                            "text":"收据编号",
                            "style":{
                                "fontColor":"#666666",
                                "hAlign":4,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":"",
                            }
                        },
                        "5":{
                            "value":"0X13409090909",
                            "text":"0X13409090909",
                            "style":{
                                "fontColor":"#666666",
                                "hAlign":2,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "cellType":{"name":4,"list":['0X13409090909','0X13509090909','0X13609090909'],"notSetIndex":2},
                                "wordWrap":false,
                                "lock":false,
                                "formatter":""
                            }
                        },
                    },
                   "3": {   //row3
                        "1":{
                                "value":"日期",
                                "text":"日期",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"italic bold 14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":"",
                                }
                            },
                        "2": {
                                "value":"",
                                "text":"",
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                            "3":{
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                            "4":{
                                "value":"楼盘号",
                                "text":"楼盘号",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "font":"14px/20px 宋体",
                                    "cellType":{"name":0},
                                    "lock":false,
                                    // "canEdit":false,
                                    "formatter":""
                                }
                            },
                            "5":{
                                "value":"0033",
                                "text":"0033",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":2,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":4,"list":["0033","0034","0035"]},
                                    "lock":false,
                                    "formatter":""
                                }
                            }
                    },
                    "4":{   //row4
                        "1": {
                                "value":"业务",
                                "text":"业务",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":true, 
                                    "fontUnderline":true,
                                    "wordWrap":true,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                        "2":  {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"yyyy-m",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                        "3":{  
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                        "4": {
                                "value":"楼盘",
                                "text":"楼盘",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":true,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                        "5":  {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":2},
                                    "lock":false,
                                }
                            }
                    },
                    "5":{    //row5
                        "1": {
                                "value":"账户",
                                "text":"账户",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                        "2":  {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "3":{
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "4":{
                                "value":"房号",
                                "text":"房号",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "5":{
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            }
                    },
                    "6":{   //row6
                        "1":   {
                                "value":"转入账户",
                                "text":"转入账户",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                        "2":    {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "3":   {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "4":    {
                                "value":"业主姓名",
                                "text":"业主姓名",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "5":    {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            }
                    },
                    "7": {    //row7
                        "1":   {
                                "value":"金额",
                                "text":"金额",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":2},
                                    "lock":false,
                                    "formatter":"",
                                }
                            },
                            "2":    {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                            "3":    {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                            "4":  {
                                "value":"收据金额",
                                "text":"收据金额",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "5":  {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":2},
                                    "lock":false,
                                }
                            }
                    },
                    "8": {   //row8
                        "1": {
                                "value":"摘要",
                                "text":"摘要",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    // "bgColor":'pink',
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "2":   {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                }
                            },
                            "3":  {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "4":  {
                                "value":"物业编号",
                                "text":"物业编号",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    // "borderRight":{
                                    //     "color":"red",
                                    //     "style":1
                                    // },
                                    "formatter":""
                                }
                            },
                            "5": {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":1},
                                    "lock":false
                                }
                            }
                    },
                    "9":{   //row9
                        "1":  {
                                "value":"所属业务部",
                                "text":"所属业务部",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "2":     {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "formatter":"",
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false
                                }
                            },
                            "3":    {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "font":"14px/20px 宋体",
                                    "formatter":""
                                }
                            },
                            "4":   {
                                "value":"凭证号",
                                "text":"凭证号",
                                "style":{
                                    "fontColor":"#666666",
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px Arial",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "5": {
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    // "borderBottom":{
                                    //     "color":"Black",
                                    //     "style":1
                                    // },
                                    "wordWrap":false,
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "font":"14px/20px 宋体",
                                    "formatter":""
                                }
                            }
                    },
                    "10":{    //row10
                        "1": {
                                "value":"所属公司",
                                "text":"所属公司",
                                "style":{
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                        "4":{
                            "style":{
                                "hAlign":1,
                                "vAlign":2,
                                "font":"14px/20px 宋体",
                                "fontStrikeout":false, 
                                "fontUnderline":false,
                                "wordWrap":false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            }
                        }
                    },
                    "11":{    //row11
                        "1": {
                                "value":"回填信息",
                                "text":"回填信息",
                                "style":{
                                    "hAlign":4,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                }
                            },
                            "4":  {
                                "value":"",
                                "text":"",
                                "style":{
                                    "hAlign":1,
                                    "vAlign":2,
                                    "font":"14px/20px 宋体",
                                    "fontStrikeout":false, 
                                    "fontUnderline":false,
                                    "wordWrap":false,
                                    "cellType":{"name":0},
                                    "lock":false,
                                    "formatter":""
                                },
                                // "formula":"IF(OR(C8=0,C8="",F3=""),"",TEXT(C4,"YYYY-MM-DD")+C9+",金额"+C8)"
                            }
                    }
                }                            
            },
            "rowHeaderData":{
                "showRowHeading":true,
           },
           "colHeaderData":{
                "showColHeading":true,
           },
        }, 
        "请假单":{
            "sheetName":"请假单",
            "rows" : [{"size": 68},{"size": 38},{"size": 30},{"size": 41},{"size": 78},{"size": 54},{"size": 54}] ,//行高
             "columns" : [{"size": 85},{"size": 60},{"size": 60},{"size": 100},{"size": 60},{"size": 60},{"size": 88},{"size": 180}],//列宽
             "spans": [
                {
                    "col": 0,
                    "colCount": 8,
                    "row":0,
                    "rowCount": 1
                },
                {
                    "col": 1,
                    "colCount": 2,
                    "row":1,
                    "rowCount": 1
                },
                {
                    "col": 4,
                    "colCount": 2,
                    "row":1,
                    "rowCount": 1
                },
                {
                    "col": 0,
                    "colCount": 1,
                    "row":2,
                    "rowCount": 2
                },
                {
                    "col": 1,
                    "colCount": 7,
                    "row":4,
                    "rowCount": 1
                },
                {
                    "col": 0,
                    "colCount": 1,
                    "row":5,
                    "rowCount": 2
                },
                {
                    "col": 1,
                    "colCount": 5,
                    "row":5,
                    "rowCount": 2
                },
             ],
             "data":{
                 "dataTable":{  
                    "0": {   
                        "0":{
                            "value":"请   假   单",
                            "text":"请   假   单",
                            "style":{
                                "font": "18px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                    },
                    "1": {   //row
                        "0":{
                            "value":"请假人姓名",
                            "text":"请假人姓名",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "3":{
                            "value":"所属部门",
                            "text":"所属部门",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "formatter":"",
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "borderBottom":{
                                    "color":"Black",
                                    "style":1,
                                },
                                "borderTop":{
                                    "color":"pink",
                                    "style":1,
                                },
                                "borderLeft":{
                                    "color":"yellow",
                                    "style":1,
                                },
                                "borderRight":{
                                    "color":"green",
                                    "style":1,
                                },
                                "lock":false
                            },
                        },
                        "6":{
                            "value":"职别",
                            "text":"职别",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                    },
                    "2": {   //row
                        "0":{
                            "value":"类\n别",
                            "text":"类\n别",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": true,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "1":{
                            "value":"事假",
                            "text":"事假",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "2":{
                            "value":"病假",
                            "text":"病假",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "3":{
                            "value":"婚假",
                            "text":"婚假",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":"",
                                "borderLeft":{
                                    "color": "yellow",
                                    "style": 1
                                }
                            },
                        },
                        "4":{
                            "value":"年假",
                            "text":"年假",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "5":{
                            "value":"补休",
                            "text":"补休",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "6":{
                            "value":"调休",
                            "text":"调休",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "7":{
                            "value":"其它",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        
                        
                    },
                    "3": {   //row
                        "1":{
                            "value":"√",
                            "text":"√",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": false,
                                "bgColor":"pink",
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""

                            },
                        },
                    },
                    "4": {   //row
                        "0":{
                            "value":"请  假\n时  间",
                            "text":"请  假\n时  间",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": true,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                    },
                    "5": {   //row
                        "0":{
                            "value":"请假事由（事假写明有什么事）",
                            "text":"请假事由（事假写明有什么事）",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": true,
                                "fontUnderline":true,
                                "fontStrikeout":true,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                        "6":{
                            "value":"主管部门\n意见",
                            "text":"主管部门\n意见",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": true,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                    },
                    "6": {   //row
                        "6":{
                            "value":"批\n示",
                            "text":"批\n示",
                            "style":{
                                "font": "14px/20px 宋体",
                                "fontColor": "#000",
                                "hAlign": 3,
                                "vAlign": 2,
                                "wordWrap": true,
                                "cellType":{"name":0},
                                "lock":false,
                                "formatter":""
                            },
                        },
                    },
                 }                            
             },
            "rowHeaderData":{
                 "showRowHeading":true,
            },
            "colHeaderData":{
                 "showColHeading":true,
            },
        },
    }
}
