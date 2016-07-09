<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)
	title=Trim(Request.Form("title"))
	url=Trim(Request.Form("url"))
	locks=Trim(Request.Form("lock"))
	sql = "INSERT INTO Module_ID(title,url,lock) VALUES('"&title&"','"&url&"',"&locks&")"
	conn.execute(sql)
	call alert("增加成功","module.asp")
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>

<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr>
    <td><form id="form1" name="form1" method="post" action="">
      <fieldset style="padding: 5" class="body">
      <legend><strong> 模块频道</strong></legend>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#AAEF92">
        <tr>
          <td width="100" height="23" align="center" bgcolor="#FFFFFF">标题:</td>
          <td bgcolor="#FFFFFF"><label>
            <input name="title" type="text" id="title" />
          </label></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">文件:</td>
          <td bgcolor="#FFFFFF"><label>
            <input name="url" type="text" id="url" />
          </label></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">锁定:</td>
          <td bgcolor="#FFFFFF"><label>
            是
                <input type="radio" name="lock" value="1" />
                否
          </label>
            <label>
            <input type="radio" name="lock" value="0" />
            </label></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30" align="center"><label>
            <input name="Submit_OK" type="submit" id="Submit_OK" value="添加" />
            <input name="Action" type="hidden" id="Action" value="Insert" />
            &nbsp;
            <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1);" />
          </label></td>
        </tr>
		
      </table>
      </fieldset>
        </form>
    </td>
  </tr>
</table>
<% 
call Rs_Close()
call DB_Close()
%>
</body>
</html>
