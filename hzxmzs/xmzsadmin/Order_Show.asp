<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<% 
ID = Trim(Request.QueryString("ID"))
set Rs=server.createobject("adodb.recordset")
Sql="select * from [Order] where ID="&ID&""
Rs.open sql,conn,1,1
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>

</head>

<body oncontextmenu="return false">

  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td> 
        <fieldset style="padding: 5" class="body">
        <legend><strong>在线预约</strong>		</legend>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
            频道类别：</td>
          <td> <strong><font color="#FF0000">查看</font></strong> </td>
        </tr>
      </table>
	  
        
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr> 
          <td width="100" height="22" align="center" bgcolor="#FFFFFF">姓名：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("UserName") %></td>
        </tr>
        <tr> 
          <td height="22" align="center" bgcolor="#FFFFFF">地址：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("Address") %></td>
        </tr>
        <tr> 
          <td height="22" align="center" bgcolor="#FFFFFF">电话：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("Tel") %></td>
        </tr>
        <tr> 
          <td height="22" align="center" bgcolor="#FFFFFF">Email：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("Email") %></td>
        </tr>
        <tr> 
          <td height="22" align="center" bgcolor="#FFFFFF">留言：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("Content") %></td>
        </tr>
        <tr> 
          <td height="22" align="center" bgcolor="#FFFFFF">订购日期：</td>
          <td bgcolor="#FFFFFF">&nbsp;<%= rs("DateTime") %></td>
        </tr>
      </table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="30" align="center"> <input type="Reset" name="reset2" value="刷新(R)" accesskey="R" onClick="javascript:location.reload();"> 
            &nbsp; <input type="button" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
          </td>
        </tr>
      </table>
      </fieldset>
</td>
  </tr>
</table>

</body>
</html>
<%
call rs_close()
call db_close()

%>



