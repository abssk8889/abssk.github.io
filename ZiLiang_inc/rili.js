
var $ = function (id) { 
return "string" == typeof id ? document.getElementById(id) : id; 
}; 
var Class = { 
create: function() { 
return function() { 
this.initialize.apply(this, arguments); 
} 
} 
} 
Object.extend = function(destination, source) { 
for (var property in source) { 
destination[property] = source[property]; 
} 
return destination; 
} 
var Calendar = Class.create(); 
Calendar.prototype = { 
initialize: function(container, options) { 
this.Container = $(container);//table结构容器
this.Days = [];//日期列表 
this.SetOptions(options); 
this.Year = this.options.Year; 
this.Month = this.options.Month; 
this.SelectDay = this.options.SelectDay ? new Date(this.options.SelectDay) : null; 
this.onSelectDay = this.options.onSelectDay; 
this.onToday = this.options.onToday; 
this.onFinish = this.options.onFinish; 
this.Draw(); 
}, 
 
SetOptions: function(options) { 
this.options = {//默认值 
Year: new Date().getFullYear(), 
Month: new Date().getMonth() + 1, 
SelectDay: null,//选择日期 
onSelectDay: function(){}, 
onToday: function(){}, 
onFinish: function(){} 
}; 
Object.extend(this.options, options || {}); 
}, 
//上月 
PreMonth: function() { 
//取得上月日期对象 
var d = new Date(this.Year, this.Month - 2, 1); 
//设置属性 
this.Year = d.getFullYear(); 
this.Month = d.getMonth() + 1; 
//重绘日历 
this.Draw(); 
}, 
//下一个月 
NextMonth: function() { 
var d = new Date(this.Year, this.Month, 1); 
this.Year = d.getFullYear(); 
this.Month = d.getMonth() + 1; 
this.Draw(); 
}, 

Draw: function() { 
//保存日期列表 
var arr = []; 
//用当月第一天在一周中的日期值作为当月离第一天的天数 
for(var i = 1, firstDay = new Date(this.Year, this.Month - 1, 1).getDay(); i <= firstDay; i++){ arr.push(" "); } 
//用当月最后一天在一个月中的日期值作为当月的天数 
for(var i = 1, monthDay = new Date(this.Year, this.Month, 0).getDate(); i <= monthDay; i++){ arr.push(i); } 
// /
var frag = document.createDocumentFragment(); 
this.Days = []; 
while(arr.length > 0){ 
//每个星期插入一个tr
var row = document.createElement("tr"); 
//星期 
for(var i = 1; i <= 7; i++){ 
var cell = document.createElement("td"); 
cell.innerHTML = " "; 
if(arr.length > 0){ 
var d = arr.shift(); 
cell.innerHTML = d; 
if(d > 0){ 
this.Days[d] = cell; 
//获取今日 
if(this.IsSame(new Date(this.Year, this.Month - 1, d), new Date())){ this.onToday(cell); } 
//判断用户是否作了选择
if(this.SelectDay && this.IsSame(new Date(this.Year, this.Month - 1, d), this.SelectDay)){ this.onSelectDay(cell); } 
} 
} 
row.appendChild(cell); 
} 
frag.appendChild(row); 
} 
//此先清空然后再插入(ie的table不能用innerHTML) 
while(this.Container.hasChildNodes()){ this.Container.removeChild(this.Container.firstChild); } 
this.Container.appendChild(frag); 
this.onFinish(); 
}, 
//判断是否同一日
IsSame: function(d1, d2) { 
return (d1.getFullYear() == d2.getFullYear() && d1.getMonth() == d2.getMonth() && d1.getDate() == d2.getDate()); 
} 
}
//导航图片切换
function a(i){
document.getElementById(i).className="s";
}
function b(i){
document.getElementById(i).className="";
}


//点击复制

function copyurl1(j) 
{ 
var url1=document.getElementById(j);//对象
url1.select(); //选择对象 
document.execCommand("Copy"); //执行浏览器复制命令 
alert("复制成功,直接发送给好友!"); 
} 
//验证码
function sjs(){
   var a=msg.num1.value=(Math.floor(Math.random()*100+1)); 
   var b=msg.num2.value=(Math.floor(Math.random()*100+1));	
    }
function tj()
		 {
		var w=msg.num1.value;
		var z=msg.num2.value;
		var y=w+z
	   if(msg.num3.value!=y){
		 alert('验证码错误');
		 msg.num3.focus();
		 msg.num3.value=""
		 return false;
		 }
		}
		