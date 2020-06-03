if (!Function.prototype.bind) (function(){
    var slice = Array.prototype.slice;
    Function.prototype.bind = function() {
      var thatFunc = this, thatArg = arguments[0];
      var args = slice.call(arguments, 1);
      if (typeof thatFunc !== 'function') {
        // closest thing possible to the ECMAScript 5
        // internal IsCallable function
        throw new TypeError('Function.prototype.bind - ' +
               'what is trying to be bound is not callable');
      }
      return function(){
        var funcArgs = args.concat(slice.call(arguments))
        return thatFunc.apply(thatArg, funcArgs);
      };
    };
})();

function Popup(options){
    // this._init(options);
    this.zIndex=29891014;                             //当前弹窗的层次
    this.index=0;                                     //当前弹窗id
    this.onblur = false;
    this.blur = this.blur.bind(this);
}

Popup.prototype._init = function(options) {
    this.shade = options.shade || false;              //是否开启遮罩层，treu 或者 false
    this.border = options.border || "1px solid #000"; //边框样式
    this.padding = options.padding || {};             //缩进,例：{top:2px,bottom:2px}
    this.boxShadow = options.boxShadow || '3px 3px 6px #009688'; //阴影样式
    this.backgroundColor = options.backgroundColor || '#9E9E9E'; //弹出框背景样式
    this.activeBackgroundColor = options.activeBackgroundColor || '#ffffff'; //内容行选中样式
    this.area = options.area || ['300px','200px'];      //弹窗大小,不能传auto,不然后面计算不了
    this.offset = options.offset || ['50%','50%'];    //弹窗位置,默认页面垂直居中显示,传参不要传百分比
    this.rowHeight = options.rowHeight || '26';     //每一行内容的高度
    this.fontSize = options.fontSize || "16";       //字体大小
    this.target = options.target || false;            //目标,为true不处理弹窗关闭
    this.data = options.data || [];                   //内容树
    this.success = options.success;                   //点击选项行的回调函数,提供3个参数：弹窗id、选中返回值,返回值在内容树的位置,ps:需要手动关闭弹窗
    this.callbackFun = options.callbackFun;           //自定义内容的回调函数, 提供一个参数：一个弹窗骨架，最后要 return 弹窗骨架(原生dom结构);

    var defaultData = [
        {
            name: 'test',
            icon: "",
            children: [
                {
                    name: 'test1',
                    icon:""
                }
            ]
            
        },
        {
            name: '测试',
            icon: "",
            href: "",
            children: [
                {
                    name: '测试1',
                    icon:"",
                    href: "",
                },
                {
                    name: '测试1',
                    icon:"",
                    href: "",
                }
            ] 
        }
    ];

    if(!this.data) return false;
    if(this.onblur)  return false;
    this._createShade(this.shade);
    if(this.callbackFun){
        var skevaron = this.callbackFun(this._createTop(),this.data);
        if(!skevaron) {
            // console.log("`returns`: can not be empty");
            return false;
        }
        document.body.appendChild(skevaron);
    } else {
        this._createMain(this.data);
    }
    this.index++;
    if(!this.onblur) {
        var _this = this;
        setTimeout(function() {
            _this._addEvent(document, 'click', _this.blur);
        }, 200);
        this.onblur = true;
    }
};

Popup.prototype._createTop = function(){
    //计算并返回滚动条宽度
    var hasScrollbar = function() {
        var odiv = document.createElement('div'),
            styles = {
                width: '100px',
                height: '100px',
                overflowY: 'scroll'
            }, i, scrollbarWidth;
        for (i in styles) odiv.style[i] = styles[i];
        document.body.appendChild(odiv);
        scrollbarWidth = odiv.offsetWidth - odiv.clientWidth;
        odiv.remove();
        return scrollbarWidth;
    }
    var style = {
        position: 'absolute',
        top: this.offset[1],
        left: this.offset[0],
        width: this.area[0],
        height: this.area[1],
        zIndex: this.zIndex,
        border: this.border,
        'overflow-y': 'auto',
        'user-select': 'none',
        'box-shadow': this.boxShadow,
        'box-sizing': 'border-box',
        backgroundColor: this.backgroundColor,
        'font-size': parseInt(this.fontSize) + "px"
    };

    var newWidth = parseInt(this.offset[1])+parseInt(this.area[1]);
    if(newWidth > window.innerHeight && parseInt(this.area[1])<window.innerHeight) {
        style['top'] = window.innerHeight-parseInt(this.area[1])+"px";
    } else if(parseInt(this.area[1])>window.innerHeight) {
        style['top'] = '0px';
        style['height'] = window.innerHeight+"px";
    }

    var newHeight = parseInt(this.offset[0])+parseInt(this.area[0]);
    if(newHeight>window.innerWidth && parseInt(this.area[0])<window.innerWidth){
        style['left'] = window.innerWidth-parseInt(this.area[0])+'px';
    } else if(parseInt(this.area[0])>window.innerWidth) {
        style['left'] = '0px';
        style['width'] = window.innerWidth+'px';
    }


    for(var item in this.padding){
        style['padding-'+item]=parseInt(this.padding[item])+'px';
        if(item === 'top' || item === 'bottom'){
            style['height'] = parseInt(style['height'])+parseInt(this.padding[item])+'px';
        }
    }
    var Div = this._createElement({
        id: 'popup-main'+this.index,
        "class": 'popup-main',
        style: style
    },"div");
    this.zIndex++;
    return Div;
}

Popup.prototype._createMain = function(arr) {
    if(arr.length == 0) {
        // console.log("`options.data`: Is an empty array");
        return false;
    };
    var onmouseover = function(bgColor,event){
        event.stopPropagation()
        event.preventDefault();
        var target = event.target;
        if(target.parentNode.className == 'item-list') {
            target.parentNode.style.backgroundColor = bgColor;
            return false;
        }
        if(target.className == 'item-list'){
            target.style.backgroundColor = bgColor;
        }
    };
    var onmouseout = function(bgColor,event) {
        event.stopPropagation()
        event.preventDefault();
        var target = event.target;
        if(target.parentNode.className == 'item-list') {
            target.parentNode.style.backgroundColor = bgColor;
            return false;
        }
        if(target.className == 'item-list'){
            target.style.backgroundColor = bgColor;
        }
    };
    var dom = this._createTop();
    var len = arr.length;
    for(var i =0; i<len; i++) {
        var flag = false;
        if(arr[i].children) {flag = true;}
        var style = {
            width: 'auto',
            'overflow-y': 'hidden',
            // height: this.rowHeight+'px',
            'overflow-wrap': "break-word",
            'word-break': 'break-all',
            'border-bottom': i != len-1 ? this.border : "",
            'line-height': parseInt(this.rowHeight)-3+'px',
            'box-sizing': 'border-box',
            'padding-bottom': '3px'
        };
        if(String(parseInt(this.area[1])) !== "NaN" && i == len-1 && parseInt(this.area[1]) - parseInt(this.rowHeight)*(i+1) >10){
            style['border-bottom'] = this.border;
        }
        var list = this._createElement({
            "class": 'item-list',
            style: style,
            href: arr[i].href,
            onclick: this.success.bind(this,this.index,arr[i].name,i),
            onmouseover: onmouseover.bind(this,this.activeBackgroundColor),
            onmouseout: onmouseout.bind(this, this.backgroundColor)
        },'div');
        if(arr[i].icon) {
            var icon = this._createElement({
                "class":'icon',
                style:{
                    background: 'url('+arr[i].icon+') no-repeat',
                    width: '16px',
                    height: '16px',
                    display: 'inline-block',
                    margin: "0px 5px",
                    "vertical-align":'middle'
                }
            },'a');
            icon.innerHTML = "";
            list.appendChild(icon);
        }
       
        var content = this._createElement({
            class: 'content',
            style: {
                // height:'100%',
                display: "inline-block",
                margin: "0px 5px",
                "vertical-align":'middle',
            }
        },'a');
        content.innerHTML = arr[i].name || '&nbsp;';
        list.appendChild(content);
        
        if(flag){
            var childIcon = this._createElement({
                "class": 'child-icon',
                style: {
                    float: 'right',
                    height: '100%',
                }
            },'a');
            childIcon.innerHTML = ">";
            list.appendChild(childIcon);
        }
        dom.appendChild(list);
    }
    document.body.appendChild(dom);
}

/**
 * 是否创建遮罩层
 * @param {Boolean} shade 布尔值
 */
Popup.prototype._createShade = function(shade) {
    if(!shade) return false;
    var dom=this._createElement({
        id:'popup-shade'+this.index,
        "class": 'popup-shade',
        style:{
            position: 'absolute',
            top: '0%',
            left: '0%',
            zIndex:this.zIndex,
            width:'100%',
            height:'100%',
            opacity:0.1,
            backgroundColor:'#000000',
        },
        onclick: this.close.bind(this,this.index),
        oncontextmenu:function(){return false}
    },'div');
    this.zIndex++;
    document.body.appendChild(dom);
}

/**
* 关闭遮罩层
*  @param {number} id 该层的index
*/
Popup.prototype.shadeClose=function(id){
   var shade=this._getId('popup-shade'+Number(id));
   shade.parentNode.removeChild(shade);
   shade=null;
}

/**
* 关闭
* @param {number} id 该层的index
*/
Popup.prototype.close = function(id){
    var iframe=this._getId('popup-main'+id);
    try{ 
        iframe.src = 'about:blank';
        iframe.contentWindow.document.write(''); 
        iframe.contentWindow.document.clear(); 
    }catch(e){} 
    //把iframe从页面移除 
    iframe ? document.body.removeChild(iframe):''; 
    iframe=null;

    if(this.onblur) {
        this._removeEvent(document, 'click', this.blur);
        this.onblur = false;
        this.target = false;
    }
    //如果有，就移除遮罩层
    try{
        this.shadeClose(id);
    }catch(e){}

}

Popup.prototype.closeAll = function(){
    this.onblur = false;
    this.target = false;
    this._removeEvent(document, 'click', this.blur);
    var mainList = this._getClassName('popup-main');
    // console.log("mainList", mainList, mainList[i]);
    for(var i=0;i<mainList.length;i++) {
        document.body.removeChild(mainList[i]);
    };
    var shadeList = this._getClassName('popup-shade');
    for(var j=0;j<shadeList.length;j++) {
        document.body.removeChild(shadeList[j]);
    };
}

/**
 * 没有遮罩层时，点击页面关闭弹窗
 */
Popup.prototype.blur = function(event) {
    var target = event.target;
    if(target.parentNode && target.parentNode.className !== 'item-list' && target.className !== 'item-list' && !this.target) {
        var list = this._getClassName('popup-main');
        for(var i=0;i<list.length;i++) {
            list[i].parentNode.removeChild(list[i]);
        }

        if(this.shade) {
            var shadeList = this._getClassName('popup-shade');
            for(var j=0;j<shadeList.length;j++) {
                document.body.removeChild(shadeList[j]);
            };
        }

        if(this.onblur) {
            this._removeEvent(document, 'click', this.blur);
            this.onblur = false;
        }
    }
}

//工具函数
Popup.prototype._getId=function(name){
    return document.getElementById(name);
},
Popup.prototype._getClassName = function(name,dom){
    if(dom) return dom.getElementsByClassName(name);
    return document.getElementsByClassName(name);
},
Popup.prototype._setAttribute=function(dom,name,value){
    //console.log('name',name,typeof name);
    if( /on\w+/.test(name) ){
        name = name.toLowerCase();
        dom[name]=value || "";
    }else if(name === 'style'){
        if(!value ||typeof value ==='string'){
            dom.style.cssText =value || '';
        }else if( value && typeof value === 'object'){
            for(var name in value){
                //可以通过style={witdh :20} 这种形式设置样式,可以省略掉单位px
                dom.style[name] = value[name];
            }
        }
    }else{
        if(name in dom){
            dom[name] =value || '';
        }
        if(value){
            dom.setAttribute(name,value);
        }else{
            dom.removeAttribute(name);
        }
    }
},
Popup.prototype._createElement=function(vnode,container){
    if( typeof vnode ==='string'){
        return document.createTextNode(vnode);
    }
    var dom=document.createElement(container);
    for(var key in vnode) {
        this._setAttribute(dom,key,vnode[key]);
    }
    return dom;
}
Popup.prototype._addEvent = function (elem, type, callback) {
    elem.addEventListener ?
    elem.addEventListener(type, callback, false) :
    elem.attachEvent('on' + type, callback);
}
Popup.prototype._removeEvent = function (elem, type, callback) {
    elem.removeEventListener ?
    elem.removeEventListener(type, callback, false) :
    elem.detachEvent('on' + type, callback);
}
