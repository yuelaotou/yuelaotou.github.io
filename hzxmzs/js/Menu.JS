function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="Images/Buttom/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="Images/Buttom/menudown.gif";
}
}
function exit(){
if(confirm('您确定要退出系统吗？'))
 {
 parent.location="Cookies.asp?Session_ID=Exit"
 }
}
