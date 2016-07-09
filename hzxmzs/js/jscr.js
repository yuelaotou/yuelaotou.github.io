// JavaScript Document
//获取当前分辨率
function streen(){
	document.write("当前分辨率为"+screen.width+"*"+screen.height)
	
}

// ***控制标题样式JavaScript Document
var pltsPop=null;
var pltsoffsetX = 10;   // ?????????????????;3-12 ??
var pltsoffsetY = 15;  // ?????????????;3-12 ??
var pltsPopbg="#FFFFEE"; //???
var pltsPopfg="#000000"; //???
var pltsTitle="cvvvv";
document.write('<div id=pltsTipLayer style="display: none;position: absolute; z-index:10001"></div>');
function pltsinits()
{document.onmouseover=plts;
document.onmousemove=moveToMouseLoc;}
function plts()
{var o=event.srcElement;
if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
pltsPop=o.dypop;
if(pltsPop!=null&&pltsPop!=""&&typeof(pltsPop)!="undefined")
{pltsTipLayer.style.left=-1000;
pltsTipLayer.style.display='';
var Msg=pltsPop.replace(/\n/g,"<br>");
Msg=Msg.replace(/\0x13/g,"<br>");
var re=/\{(.[^\{]*)\}/ig;
if(!re.test(Msg))pltsTitle="???????!";
else{
re=/\{(.[^\{]*)\}(.*)/ig;
pltsTitle=Msg.replace(re,"$1")+"&nbsp;";
re=/\{(.[^\{]*)\}/ig;
Msg=Msg.replace(re,"");
Msg=Msg.replace("<br>","");}
var content =
'<table style="background:#E3F1FA;border:1px solid #444444" id=toolTipTalbe><tr id=pltsPoptop><th id=topleft><p id=topright style="display:none"></th></tr><tr><td  style="padding:4px;line-height:135%">'+Msg+'</td></tr>'+
'<tr id=pltsPopbot style="display:none"><th><p id=botleft></p><p id=botright align=rightstyle="display:none"></th></tr></table>';
pltsTipLayer.innerHTML=content;
toolTipTalbe.style.width=Math.min(pltsTipLayer.clientWidth,document.body.clientWidth/2.2);
moveToMouseLoc();
return true;}else{pltsTipLayer.innerHTML='';
pltsTipLayer.style.display='none';
return true;}}
function moveToMouseLoc()
{if(pltsTipLayer.innerHTML=='')return true;
var MouseX=event.x;
var MouseY=event.y;
//window.status=event.y;
var popHeight=pltsTipLayer.clientHeight;
var popWidth=pltsTipLayer.clientWidth;
if(MouseY+pltsoffsetY+popHeight>document.body.clientHeight)
{popTopAdjust=-popHeight-pltsoffsetY*1.5;
pltsPoptop.style.display="none";
pltsPopbot.style.display="";}else{
popTopAdjust=0;pltsPoptop.style.display="";pltsPopbot.style.display="none";}
if(MouseX+pltsoffsetX+popWidth>document.body.clientWidth){
popLeftAdjust=-popWidth-pltsoffsetX*2;topleft.style.display="none";
botleft.style.display="none";
topright.style.display="";
botright.style.display="";}else
{popLeftAdjust=0;topleft.style.display="";botleft.style.display="";
topright.style.display="none";botright.style.display="none";}
pltsTipLayer.style.left=MouseX+pltsoffsetX+document.body.scrollLeft+popLeftAdjust;
pltsTipLayer.style.top=MouseY+pltsoffsetY+document.body.scrollTop+popTopAdjust;
return true;}
pltsinits();