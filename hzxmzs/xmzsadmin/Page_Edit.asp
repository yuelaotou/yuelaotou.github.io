<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'验证用户权限
	Dim Rs,Sql
	SS_ID = Trim(Request.Form("SS_ID"))
	ID = Trim(Request.Form("ID"))
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from Page where ID="&ID
	Rs.open sql,conn,1,3
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("Rank") = Trim(Request.Form("Rank"))
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('修改成功!');location='Page_List.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ID = Trim(Request.QueryString("ID"))
set Rs=server.createobject("adodb.recordset")
Sql="select * from Page where ID="&ID
Rs.open sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
</head>
<body oncontextmenu="return false">
<form name="Form1" method="post" action="" onSubmit="return input_ok(this)">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td bgcolor="#F0E7CF"> <fieldset style="padding: 5">
        <legend><strong>多栏单页</strong></legend>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
              频道类别：</td>
            <td width="150"> <strong><font color="#FF0000">
			<% 
Dim SS_ID
SS_ID = Request.QueryString("SS_ID")
set Rc=server.createobject("adodb.recordset")
Sql="select * from Type Where ID="&SS_ID&""
Rc.open sql,conn,1,1
Response.Write(Rc("Title"))
Rc.close
SET Rc=nothing
%>

			
			
			</font></strong></td>
            <td align="center">&nbsp;</td>
          </tr>
        </table>	
		
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="30" align="center" bgcolor="#FFFFFF">标题：</td>
            <td bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" value="<%= RS("Title") %>" size="30">&nbsp;&nbsp;排序：
              <input name="Rank" type="text" id="Rank" size="4" value="<%= RS("Rank") %>" maxlength="4">
              [正整数]</td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF">内容：</td>
            <td bgcolor="#FFFFFF">
			
			<textarea name="Content" style="display:none"><%= RS("Content") %></textarea> 
        <iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width="650" HEIGHT="500"></iframe>
			
			
			
			</td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"> <input type="submit" name="Submit_OK" value="修改(E)" accesskey="E"> 
              &nbsp; <input type="button" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
			  <input name="Action" type="hidden" id="Action" value="Edit">
			 <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>">
			 <input name="ID" type="hidden" id="SS_ID" value="<%= ID %>">
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
call rs_close()
call db_close()
%>