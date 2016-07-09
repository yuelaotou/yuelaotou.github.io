//////////////////////////////////////////////////////////
function KeyDown(){ 
//alert(event.keyCode);
//ÆÁ±Î Alt+F4
if ((event.altKey == true)&&(event.keyCode==115)) {
	window.showModelessDialog("about:blank","","dialogWidth:1px;dialogheight:1px"); 
	return false;} 
//ÆÁ±Î Shift
//if (event.shiftKey == true){alert("Shift¼üÒÑÆÁ±Î")} 	
//ÆÁ±Î Ctrl + R
if (event.ctrlKey == true && event.keyCode==82)
	{event.keyCode=0;return false;} 
//ÆÁ±Î F3¡¢F4¡¢F5 ¡¢F6¡¢F11
if ((event.keyCode==114)||(event.keyCode==115)||(event.keyCode==116)||(event.keyCode==117)||(event.keyCode==122))
	{event.keyCode=0;return false;} 
//ÆÁ±Î Alt+·½Ïò¼ü
if ((event.altKey == true)&&((event.keyCode==37)||(event.keyCode==39))) {   
	return false;}
//ÆÁ±Î Ctrl+N
if ((event.ctrlKey == true)&&(event.keyCode==78)) {return false;}
//ÆÁ±Î shift+F10
if ((event.shiftKey == true)&&(event.keyCode==121)) {return false;}
} 
document.onkeydown = KeyDown;
//********************************************
//var currentpos,timer;
//function initialize()
//{
//timer=setInterval ("scrollwindow ()",30);
//}
//function sc()
//{
//clearInterval(timer);
//}
//function scrollwindow()
//{
//currentpos=document.body.scrollTop;
//window.scroll(0,++currentpos);
//if (currentpos !=document.body.scrollTop)
//sc();
//}
//document.onmousedown=sc
//document.ondblclick=initialize

