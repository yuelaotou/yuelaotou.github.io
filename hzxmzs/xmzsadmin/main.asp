<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>


<!--打开关闭全屏的样式开始-->
<STYLE type=text/css>
.navPoint {
	FONT-SIZE: 12px; CURSOR: hand; COLOR: white; FONT-FAMILY: Webdings
}
</STYLE>
<!--打开关闭全屏的样式结束-->
<SCRIPT language="JavaScript">

<!--//这是实现框架的伸缩
function switchBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("leftFrame").style.display="none"
}else{
switchPoint.innerText=3
document.all("leftFrame").style.display=""
}}
//-->
</SCRIPT>
<!--这段代码实现在状态栏上出现停留时间-->
<script language="JavaScript" >
    var Temp;
    var TimerId = null;
    var TimerRunning = false;

    Seconds = 0
    Minutes = 0
    Hours = 0

    function showtime() 
    {
      if(Seconds >= 59) 
      {
        Seconds = 0
        if(Minutes >= 59) 
        {
          Minutes = 0
          if(Hours >= 23) 
          {
            Seconds = 0
            Minutes = 0
            Hours = 0
          } 
          else { 
            ++Hours 
          }
        } 
        else { 
          ++Minutes 
        }
      } 
      else { 
        ++Seconds 
      }

      if(Seconds != 1) { var ss="s" } else { var ss="" }
      if(Minutes != 1) { var ms="s" } else { var ms="" }
      if(Hours != 1) { var hs="s" } else { var hs="" }

      Temp = '停留时间 '+Hours+' 小时'+', '+Minutes+' 分'+', '+Seconds+' 秒'+''
      window.status = Temp;
      TimerId = setTimeout("showtime()", 1000);
      TimerRunning = true;
    }
  
    var TimerId = null;
    var TimerRunning = false;

    function stopClock() {
      if(TimerRunning) 
        clearTimeout(TimerId);
        TimerRunning = false;
    }

    function startClock() {
      stopClock();
      showtime();
    }

    function stat(txt) {
      window.status = txt;
      setTimeout("erase()", 2000);
    }

    function erase() {
      window.status = "";
    }

</SCRIPT>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body style="MARGIN: 0px" scroll=no onLoad="loadBar(0)">
<NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT>
<!--这是调用状态栏停留时间-->
<script type="text/javascript" language="javascript">//startClock();</script>
<script language="javascript">

function loadBar(fl)
//fl is show/hide flag
{
  var x,y;
  if (self.innerHeight)
  {// all except Explorer
    x = self.innerWidth;
    y = self.innerHeight;
  }
  else 
  if (document.documentElement && document.documentElement.clientHeight)
  {// Explorer 6 Strict Mode
   x = document.documentElement.clientWidth;
   y = document.documentElement.clientHeight;
  }
  else
  if (document.body)
  {// other Explorers
   x = document.body.clientWidth;
   y = document.body.clientHeight;
  }

    var el=document.getElementById('loader');
	if(null!=el)
	{
		var top = (y/2) - 50;
		var left = (x/2) - 100;
		if( left<=0 ) left = 10;
		el.style.visibility = (fl==1)?'visible':'hidden';
		el.style.display = (fl==1)?'block':'none';
		el.style.left = left + "px"
		el.style.top = top + "px";
		el.style.zIndex = 2;
	}
}
</script>
<div id="loader" style="Z-INDEX: 100;position:absolute;display:none">
<table style="FILTER: Alpha(opacity=90);" border="0" cellpadding="5" cellspacing="1" bgcolor="#D7DEF8" onClick="loadBar(0)"><tr>
      <td bgcolor="#FFFFFF" align="left"><img src="../sysImages/load.gif" alt="请等待" width="94" height="17" align="absmiddle" style="margin:3px">&nbsp;&nbsp;数据载入中...&nbsp;<br>
</td></tr></table>
</div>
<script type="text/javascript" language="javascript">loadBar(1);</script>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top" id=leftFrame name="leftFrame">
	
	<iframe height="100%" width="200" name="Left_main" src="left.asp" frameborder="0" scrolling="yes"></iframe>
	
	
	
	</td>
    <td onClick="switchBar()" class="main">
	
	<FONT style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
	<BR>
 <BR>
 <BR><BR>
 <BR>
            <BR>
            <BR>
            <BR>
           <SPAN class=navPoint id=switchPoint title=打开/关闭全屏>3</SPAN><BR>
           <BR>
      <BR><BR><BR><BR>屏幕切换    </FONT>
	
	
	
	</td>
    <td valign="top" style="WIDTH: 100%"><iframe height="100%" width="100%" name="main" src="User_Show.asp" frameborder="0" scrolling="yes"></iframe></td>
  </tr>
</table>

</body>
</html>
