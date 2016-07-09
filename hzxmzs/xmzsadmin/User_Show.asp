<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<%
ID=Trim(Request.QueryString("ID"))
IF ID = Empty Then
	Str = "Uid='"&Session("Uid")&"'"
Else
	Str = "id="&id
End IF
SQL = "Select * from admin where "&Str&""
set rs=server.CreateObject("ADODB.RecordSet")
rs.open SQL,conn,1,1

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>

<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body oncontextmenu="return false">
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong>基本资料</strong></legend>
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td bgcolor="#FFFFFF"><br>
              <table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#D7DEF8">
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">头&nbsp;&nbsp;像：</td>
                  <td><img src="../sysImages/head/<%= Rs("Head") %>.gif" name="qq" width="32" height="32" align="absmiddle"> </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">类&nbsp;&nbsp;型：</td>
                  <td><% If RS("Flag") = 0 Then %>
                    普通用户
                    <% Else %>
                    <font color="#FF0000">超级管理员</font>
                    <% End If %>
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td width="100" height="22" align="center">用户名：</td>
                  <td>&nbsp;<%= Rs("UID") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="23" align="center">姓&nbsp;&nbsp;名：</td>
                  <td>&nbsp;<%= Rs("Name") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">性&nbsp;&nbsp;别：</td>
                  <td>&nbsp;<%= RS("sex") %> </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">年&nbsp;&nbsp;龄：</td>
                  <td>&nbsp;<%= RS("age") %> </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">电&nbsp;&nbsp;话：</td>
                  <td>&nbsp;<%= RS("tel") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">QQ：</td>
                  <td>&nbsp;<%= RS("qq") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">邮&nbsp;&nbsp;件：</td>
                  <td>&nbsp;<%= RS("Email") %> </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">网&nbsp;&nbsp;址：</td>
                  <td>&nbsp;<%= RS("URL") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">登&nbsp;&nbsp;录：</td>
                  <td>&nbsp;<%= RS("Counts") %>&nbsp;次</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">登录IP：</td>
                  <td>&nbsp;<%= RS("IP") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">简&nbsp;&nbsp;介：</td>
                  <td>&nbsp;<%= rs("Content") %></td>
                </tr>
              </table>
            <br>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30" align="center"><input type="Reset" name="reset2" value="刷新" onClick="javascript:location.reload();">
            &nbsp;
            <input type="button" name="reset" value="返回" onClick="javascript:history.go(-1);">
          </td>
        </tr>
      </table>
    </fieldset></td>
  </tr>
</table>
</body>
</html>
