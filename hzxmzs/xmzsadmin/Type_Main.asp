<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<BODY style="MARGIN: 0px" scroll=no oncontextmenu="return false">
<NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="50%">
		<IFRAME id=carnoc style="WIDTH: 450; HEIGHT: 100%" src="Type_list.asp" frameBorder=0 scrolling=yes name="Type_left">
    </IFRAME></td>
    <td width="50%">
	<IFRAME id=carnoc style="WIDTH: 100%; HEIGHT: 100%" name=Type_Main src="Type_Insert.asp" frameBorder=0 scrolling=yes>
    </IFRAME>
	</td>
  </tr>
</table>
</body>
</html>
