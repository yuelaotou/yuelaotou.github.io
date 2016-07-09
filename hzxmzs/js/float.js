function f_animate_quickBoard()
{		
	
	obj_layer = document.all["ly_quickBoard"];

	v_menuY_min = 380
	v_menuY_max_fromBottom = 0;
	
	/*v_menuX = 919;*/
	v_menuY = v_menuY_min + document.documentElement.scrollTop;
	
	v_nowY_str = obj_layer.style.top;
	v_nowY_num = parseInt(v_nowY_str.substring(0, v_nowY_str.length - 2));
	
	v_layerHeightStr = document.all["ly_quickBoard"].style.height
	v_layerHeightNum = parseInt(v_layerHeightStr.substring(v_layerHeightStr, v_layerHeightStr.length - 2));
	
	if((document.documentElement.scrollTop > v_menuY_min) || (v_nowY_num > v_menuY_min))
	{
		v_menuY = document.documentElement.scrollTop + 100
		
		v_menuY_max = document.documentElement.clientHeight + document.documentElement.scrollTop - 
		v_layerHeightNum - v_menuY_max_fromBottom;
		
		if(v_menuY < v_menuY_min) v_menuY = v_menuY_min;
		if(v_menuY > v_menuY_max) v_menuY = v_menuY_max;
			
		f_move_quickBoard(obj_layer, /*v_menuX, */v_menuY);
	}
	
	setTimeout(f_animate_quickBoard, 10);
	
}

function f_move_quickBoard(p_layer,/* p_x,*/ p_y)
{
	v_speedX = 10;
	v_speedY = v_speedX;
	
/*	p_layer.style.left = parseInt(p_layer.style.left) - ((parseInt(p_layer.style.left) - p_x) / v_speedX);*/
	p_layer.style.top = parseInt(p_layer.style.top) - ((parseInt(p_layer.style.top) - p_y) / v_speedY);

	if(document.all.top_div)
	{
		document.all.top_div.style.left = 0;
		document.all.top_div.style.top = parseInt(p_layer.style.top) + 36;
	}
}

f_animate_quickBoard();

function autoBlur(){ 

	try{
		if(event.srcElement.tagName=="A"||event.srcElement.tagName=="IMG") document.documentElement.focus(); 
	}catch(e) {}
}
document.onfocusin=autoBlur;
