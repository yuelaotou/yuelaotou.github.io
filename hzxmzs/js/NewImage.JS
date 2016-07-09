//自适应图片大小的弹出窗口
//onclick="newimg(this.src);"
function newimg(imgurl)
{ 
newwin=window.open('about:blank','','width=0,height=0'); //先新打开一个小窗口
newwin.document.write('<body leftmargin=0 topmargin=0><img id=img1 src='+imgurl+' onload="self.resizeTo(this.offsetWidth + 10,this.offsetHeight + 50);">'); //把窗口的大小调整到与图片一样大
} 
//onload="javascript:DrawImage(this);"
//图片的实际大小
function DrawImage(ImgD){ 
	var image=new Image(); 
	image.src=ImgD.src; 
	if(image.width>0 && image.height>0){ 
		flag=true; 
		//if(image.width>=518){ 
			//ImgD.width=518; 
			//ImgD.height=(image.height*518)/image.width; 
		//}else{ 
			ImgD.width=image.width; 
			ImgD.height=image.height; 
		//}  
	} 
} 
