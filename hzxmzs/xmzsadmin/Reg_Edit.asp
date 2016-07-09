<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/MD5.Asp" -->
<%
IF Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'验证用户权限
	ID=Trim(Request.QueryString("ID"))
	pwd=Trim(Request.Form("pwd"))
	pwd=md5(pwd)
	real_name=Trim(Request.Form("real_name"))
	email=Trim(Request.Form("email"))
	qq=Trim(Request.Form("qq"))
	tel=Trim(Request.Form("tel"))
	url=Trim(Request.Form("url"))
	address=Trim(Request.Form("address"))
	Sql = "update Reg set pwd='"&pwd&"',real_name='"&real_name&"',email='"&email&"',qq='"&qq&"',tel='"&tel&"',url='"&url&"',address='"&address&"' where id="&id
	conn.execute(sql)
	call alert("用户修改成功","Reg.asp")



End if
'''''''''''''''''''''''''''''''''''
id=Trim(Request.QueryString("id"))
sql="select * from Reg where id="&id
set rs=conn.execute(sql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong>用户修改</strong></legend>
      <form name="form1" method="post" action="">
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr>
            <td width="100" height="23" align="center" bgcolor="#FFFFFF">帐号</td>
            <td bgcolor="#FFFFFF"><label>
              <input name="UID" type="text" id="UID" size="20" value="<%=rs("UID")%>" disabled="disabled" />
              *&nbsp;帐号不能修改</label></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">密码</td>
            <td bgcolor="#FFFFFF"><input name="PWD" type="password" id="PWD" size="20" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">姓名</td>
            <td bgcolor="#FFFFFF"><label>
              <input name="Real_Name" type="text" id="Real_Name" size="20" value="<%=RS("Real_Name")%>"/>
            </label></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">邮件</td>
            <td bgcolor="#FFFFFF"><input name="Email" type="text" id="Email" size="30" value="<%=rs("Email")%>" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">qq</td>
            <td bgcolor="#FFFFFF"><input name="qq" type="text" id="qq" size="20" value="<%=rs("qq")%>" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">电话</td>
            <td bgcolor="#FFFFFF"><input name="Tel" type="text" id="Tel" size="20" value="<%=rs("Tel")%>" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">网址</td>
            <td bgcolor="#FFFFFF"><input name="url" type="text" id="url" size="30" value="<%=rs("URL")%>" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">地址</td>
            <td bgcolor="#FFFFFF"><input name="address" type="text" id="address" size="30" value="<%=RS("address")%>" /></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="添加">
            &nbsp;
              <input type="button" name="reset" value="返回" onClick="javascript:history.go(-1);" />
              <input name="Action" type="hidden" id="Action" value="Edit" /></td>
          </tr>
        </table>
      </form>
    </fieldset></td>
  </tr>
</table>
<%
call rs_close()
call db_close()
%>
</body>
</html>
