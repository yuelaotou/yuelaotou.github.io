//服务项目的脚本实现
/*function service(){
	var len;
	var showerObj;
	var listObj;
	var showerWidth=200;
	var showerHeight=160;
	var r;
	var cR=0;
	var ccR=0;
	var timer=0;
	window.onload=function(){
		showerObj=document.getElementById("show");
		listObj=showerObj.getElementsByTagName("div");
		len=listObj.length;
		r=Math.PI/180*360/len;
		for(var i=0;i<len;i++){
			var item=listObj[i];
			item.style.top=showerHeight/2+Math.sin(r*i)*showerWidth/2+"px";
			item.style.left=showerWidth/2+Math.cos(r*i)*showerWidth/2+"px";
			item.rotate=(r*i+2*Math.PI)%(2*Math.PI);
			item.onclick=function(){
				cR=Math.PI/2-this.rotate;
				timer || (timer=setInterval(rotate,10));
			}
		}
		rotate();
		var rX=showerObj.offsetLeft+showerWidth/2;
		var ry=showerObj.offsetTop+showerHeight/2;
		function rotate(){
			ccR=(ccR+2*Math.PI)%(2*Math.PI);
			if(cR-ccR<0) cR=cR+2*Math.PI;
			if(cR-ccR<Math.PI){
				ccR=ccR+(cR-ccR)/19;
			}else{
				ccR=ccR-(2*Math.PI+ccR-cR)/19;
			}
		
			if(Math.abs((cR+2*Math.PI)%(2*Math.PI)-(ccR+2*Math.PI)%(2*Math.PI))<Math.PI/720){
				ccR=cR;
				clearInterval(timer);
				timer=0;
			}
		
			for(var i=0;i<len;i++){
				var item=listObj[i];
				var w,h;
				var sinR=Math.sin(r*i+ccR);
				var cosR=Math.cos(r*i+ccR);
				w=40+0.6*40*sinR;
				h=(40+0.6*40*sinR);
				item.style.cssText +=";width:"+w+"px;height:"+h+"px;top:"+parseInt(showerHeight/2+sinR*showerWidth/2/3-w/2+10)+"px;left:"+parseInt(showerWidth/2+cosR*showerWidth/2-h/2+17)+"px;z-index:"+parseInt(showerHeight/2+sinR*showerWidth/2/3-w/2)+";";
			}
		}
		document.getElementById("l").onclick=function(){
			cR=(cR+r+2*Math.PI)%(2*Math.PI);
			timer || (timer=setInterval(rotate,10));
		}
		document.getElementById("r").onclick=function(){
			cR=(cR-r+2*Math.PI)%(2*Math.PI);
			timer || (timer=setInterval(rotate,10));
		}
	}
}
*/


//图片，文字连续滚动方法

//容器ID
//向上滚动(0向上 1向下 2向左 3向右)
//滚动的步长
//容器可视宽度
//容器可视高度
//定时器 数值越小，滚动的速度越快(1000=1秒)
//间歇停顿时间(0为不停顿,1000=1秒)
//开始时的等待时间(0为不等待,1000=1秒)
//间歇滚动间距(可选)
function Marquee(){
	this.ID=document.getElementById(arguments[0]);
	this.Direction=arguments[1];
	this.Step=arguments[2];
	this.Width=arguments[3];
	this.Height=arguments[4];
	this.Timer=arguments[5];
	this.WaitTime=arguments[6];
	this.StopTime=arguments[7];
	if(arguments[8]){
		this.ScrollStep=arguments[8];
	}else{
		this.ScrollStep=this.Direction>1?this.Width:this.Height;
	}
	this.CTL=this.StartID=this.Stop=this.MouseOver=0;
	this.ID.style.overflow="hidden";
	this.ID.noWrap=true;
	this.ID.style.width=this.Width;
	this.ID.style.height=this.Height;
	this.ClientScroll=this.Direction>1?this.ID.scrollWidth:this.ID.scrollHeight;
	this.ID.innerHTML+=this.ID.innerHTML;
	this.Start(this,this.Timer,this.WaitTime,this.StopTime);
}
Marquee.prototype.Start=function(msobj,timer,waittime,stoptime){
	msobj.StartID=function(){
		msobj.Scroll();
	}
	msobj.Continue=function(){
		if(msobj.MouseOver==1){
			setTimeout(msobj.Continue,waittime);
		}else{
			clearInterval(msobj.TimerID);
			msobj.CTL=msobj.Stop=0;
			msobj.TimerID=setInterval(msobj.StartID,timer);
		}
	}
	msobj.Pause=function(){
		msobj.Stop=1; 
		clearInterval(msobj.TimerID); 
		setTimeout(msobj.Continue,waittime);
	}
	msobj.Begin=function(){
		msobj.TimerID=setInterval(msobj.StartID,timer);
		msobj.ID.onmouseover=function(){
			msobj.MouseOver=1; 
			clearInterval(msobj.TimerID);
		}
		msobj.ID.onmouseout=function(){
			msobj.MouseOver=0; 
			if(msobj.Stop==0){
				clearInterval(msobj.TimerID); 
				msobj.TimerID=setInterval(msobj.StartID,timer);
			}
		}
	}
	setTimeout(msobj.Begin,stoptime);
}
Marquee.prototype.Scroll=function(){
	switch(this.Direction){
		case 0:
			this.CTL+=this.Step;
			if(this.CTL>=this.ScrollStep&&this.WaitTime>0){
				this.ID.scrollTop+=this.ScrollStep+this.Step-this.CTL;
				this.Pause();
				return;
			}else{
				if(this.ID.scrollTop>=this.ClientScroll){
					this.ID.scrollTop-=this.ClientScroll;
				}
				this.ID.scrollTop+=this.Step;
			}
			break;
		case 1:
			this.CTL+=this.Step;
			if(this.CTL>=this.ScrollStep&&this.WaitTime>0){
				this.ID.scrollTop-=this.ScrollStep+this.Step-this.CTL;
				this.Pause();
				return;
			}else{
				if(this.ID.scrollTop<=0){
					this.ID.scrollTop+=this.ClientScroll;
				}
				this.ID.scrollTop-=this.Step;
			}
			break;
		case 2:
			this.CTL+=this.Step;
			if(this.CTL>=this.ScrollStep&&this.WaitTime>0){
				this.ID.scrollLeft+=this.ScrollStep+this.Step-this.CTL;
				this.Pause();
				return;
			}else{
				if(this.ID.scrollLeft>=this.ClientScroll){
					this.ID.scrollLeft-=this.ClientScroll;
				}
				this.ID.scrollLeft+=this.Step;
			}
			break;
		case 3:
			this.CTL+=this.Step;
			if(this.CTL>=this.ScrollStep&&this.WaitTime>0){
				this.ID.scrollLeft-=this.ScrollStep+this.Step-this.CTL;
				this.Pause();
				return;
			}else{
				if(this.ID.scrollLeft<=0){
					this.ID.scrollLeft+=this.ClientScroll;
				}
				this.ID.scrollLeft-=this.Step;
			}
			break;
	}
}
//图片拖拽翻页
/*
function id(obj){
	return document.getElementById(obj);
}
var page;
var lm,mx;
var md=false;
var sh=0;
var en=false;
function picpage(id,tag){
	var picpage = document.getElementById(id);
	page = picpage.getElementsByTagName(tag);
	if(page.length>0){
		page[0].style.zIndex=2;
	}

	for(i=0;i<page.length;i++){
		page[i].className="page";
		page[i].id="page"+i;
		page[i].i=i;
		page[i].onmousedown=function(e){
			if(!en){
				if(!e){e=e||window.event;}
				lm=this.offsetLeft;
				mx=(e.pageX)?e.pageX:e.x;
				this.style.cursor="w-resize";
				md=true;
				if(document.all){
					this.setCapture();
				}else{
					window.captureEvents(Event.MOUSEMOVE|Event.MOUSEUP);
				}
			}
		}
		page[i].onmousemove=function(e){
			if(md){
				en=true;
				if(!e){e=e||window.event;}
				var ex=(e.pageX)?e.pageX:e.x;
				this.style.left=ex-(mx-lm)+350;
				if(this.offsetLeft<75){
					var cu=(this.i==0)?page.length-1:this.i-1;
					page[sh].style.zIndex=0;
					page[cu].style.zIndex=1;
					this.style.zIndex=2;
					sh=cu;
				}
			
				if(this.offsetLeft>75){
					var cu=(this.i==page.length-1)?0:this.i+1;
					page[sh].style.zIndex=0;
					page[cu].style.zIndex=1;
					this.style.zIndex=2;
					sh=cu;
				}
			}
		}
		page[i].onmouseup=function(){
			this.style.cursor="default";
			md=false;
			if(document.all){
				this.releaseCapture();
			}else{
				window.releaseEvents(Event.MOUSEMOVE|Event.MOUSEUP);
			}
			flyout(this);
		}
	}
}
function flyout(obj){
	if(obj.offsetLeft<75){
		if((obj.offsetLeft + 350 - 20) > -275 ){
			obj.style.left=obj.offsetLeft + 350 - 20;
			window.setTimeout("flyout(id('"+obj.id+"'));",0);
		}else{
			obj.style.left=-275;
			obj.style.zIndex=0;
			flyin(id(obj.id));
		}
	}

	if(obj.offsetLeft>75){
		if((obj.offsetLeft + 350 + 20) < 1125 ){
			obj.style.left=obj.offsetLeft + 350 + 20;
			window.setTimeout("flyout(id('"+obj.id+"'));",0);
		}else{
			obj.style.left=1125;
			obj.style.zIndex=0;
			flyin(id(obj.id));
		}
	}
}

function flyin(obj){
	if(obj.offsetLeft<75){
		if((obj.offsetLeft + 350 + 20) < 425  ){
			obj.style.left=obj.offsetLeft + 350 + 20;
			window.setTimeout("flyin(id('"+obj.id+"'));",0);
		}else{
			obj.style.left=425;
			en=false;
		}
	}
	if(obj.offsetLeft>75){
		if((obj.offsetLeft + 350 - 20) > 425  ){
			obj.style.left=obj.offsetLeft + 350 - 20;
			window.setTimeout("flyin(id('"+obj.id+"'));",0);
		}else{
			obj.style.left=425;
			en=false;
		}
	}
}
*/