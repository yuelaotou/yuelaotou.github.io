<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<% 
'删除单个
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'验证用户权限
	ID = Trim(Request.QueryString("ID"))
	Call Del("[Order]",ID,"?")
End if 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'发货
If Trim(Request.QueryString("Action"))="Lock" Then
	Call Check_Right(5)	'验证用户权限
	ID = Trim(Request.QueryString("ID"))
	Flag = Trim(Request.QueryString("Flag"))
	ReCord = Trim(Request.QueryString("ReCord"))
	Call Lock_OK("[Order]",ID,ReCord,Flag,"?")
End if 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Rs,Sql
SS_ID=request.QueryString("SS_ID")
set Rs=server.createobject("adodb.recordset")
Sql="select * from [Order] where SS_ID="&SS_ID&" order by id desc"
Rs.open sql,conn,1,1
Rs.pagesize=23
If Len(Request.QueryString("Page")) = 0 Then 
	Page = 1 
Else 
	Page=Cint(Request.Querystring("Page")) 
End If 
Rs.AbsolutePage=Page
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
        <legend><strong>在线预约</strong></legend>
		
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
            频道类别：</td>
          <td> <strong><font color="#FF0000">所有</font></strong> </td>
        </tr>
      </table>
	  
        
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr bgcolor="#FFFFFF"> 
          <td width="50" height="22" align="center">编号</td>
          <td width="80" align="center">姓名</td>
          <td width="353" align="center">地址</td>
          <td width="150" align="center">邮箱</td>
          <td width="126" align="center">电话</td>
          <td width="123" align="center" bgcolor="#FFFFFF">日期</td>
          <td width="30" align="center">详细</td>
          <td width="30" align="center">删除</td>
        </tr>
        <%
dim I
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
        <tr bgcolor="#FFFFFF" style="cursor:hand" OnMouseOver="this.bgColor='#F0F0F0'" OnMouseOut="this.bgColor='#FFFFFF'"> 
          <td height="22" align="center"><%= rs("id") %></td>
          <td align="center">&nbsp;<%= rs("UserName") %></td>
          <td>&nbsp;<%= rs("Address") %></td>
          <td align="center"><%= rs("Email") %></td>
          <td align="center"><%= rs("Tel") %></td>
          <td align="center"><%= rs("DateTime") %></td>
          <td align="center"><a href="Order_Show.asp?ID=<%= rs("ID") %>"><img src="../sysImages/Search.gif" width="20" height="20" border="0" align="absmiddle"></a></td>
          <td align="center"> <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&SS_ID=<%= SS_ID %>&ID=<%= rs("id") %>';return true;}return false;}">          </td>
        </tr>
        <% 
rs.MoveNext
NEXT
%>
      </table>
<% Call RS_Empty()	'没有记录 %>
  
        
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="50%" height="30" align="center" valign="middle"> <% call Image_Page("") %>  </td>
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


