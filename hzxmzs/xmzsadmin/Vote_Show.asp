<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'''''''''''''''''''''''''''''''''''''''''''
if request("title")<>"" then
set rs=conn.execute("insert into Vote (title,Click,SS_ID)values('"&request("title")&"',"&request("Click")&",1)")
response.Redirect("Vote_Show.asp")
end if
if request("send")="editp" then
set rs=conn.execute("update Vote set title='"&request("titlet")&"',Click="&request("Click")&" where id="&request("id"))
response.Redirect("Vote_Show.asp")
end if
if request("name")<>"" then
if request("titlesub")="" then
sql="insert into Vote (title,SS_ID)values('"&request("name")&"',0)"
else
sql="update Vote set title='"&request("name")&"' where id="&request("titlesub")
end if
set rs=conn.execute(sql)
response.Redirect("Vote_Show.asp")
end if
IF Trim(Request.QueryString("Action"))="Del" Then
	ID=Trim(Request.QueryString("ID"))
    set rs=conn.execute("delete from Vote where id="&ID)
    response.Redirect("Vote_Show.asp")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.Form("Action")) = "Del_ALL" Then
	Call Check_Right(4)	'验证用户权限
	Dim ID
	ID = Trim(Request.Form("ID"))
	If ID = Empty then
		Response.write("<script>alert('系统提示:\n\n请选择删除文件！');location='?';</script>")
		Response.End
	Else
		Call Del_ALL("Vote",ID,"?")
	End if
End if 
''''''''''''''''''''''''''''''''''''''''''''
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>

</head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong>网站调查</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
		  <%
		  set rs=conn.execute("select * from Vote where SS_ID=0")
		  %>
            <td width="110" height="22" align="center">投票标题:</td>
			<form name="form2" method="post" action="">
            <td width="375" align="left">
              <input name="name" type="text" size="50" value=<%=rs("title")%>>            </td>
            <td width="85" align="center">
			<input type="submit" name="Submit2" value="修改标题">
			<input name="titlesub" type="hidden" value="<%=rs("id")%>">
			</td>
			</form>
            <td width="381">&nbsp;</td>
          </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
              <tr bgcolor="#FFFFFF">
                <td width="50" height="22" align="center">编号</td>
                <td width="578" align="center">标题</td>
                <td width="288" align="center">得票数</td>
                <td width="30" align="center">删除</td>
              </tr>
              <%
Sql = "select * from Vote where SS_ID=1"
set rs=server.CreateObject("ADODB.RecordSet")
rs.open Sql,conn,1,1
dim I
do while not rs.eof
    IF rs.eof THEN EXIT do
 %>
              <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#F0F0F0'" onMouseOut="this.bgColor='#FFFFFF'">
                <td height="22" align="center"><%= rs("id") %></td>
                <td align="center">&nbsp;<a href="?send=edit&id=<%=rs("id")%>"><%= rs("title") %></a></td>
                <td align="center"><%=rs("Click")%></td>
                <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
              </tr>
              <% 
rs.MoveNext
loop
%>
            </table>
                <% Call RS_Empty()	'没有记录 %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%if request("send")<>"edit" then %>
                  <tr>
                    <td width="9%" height="30" align="center" valign="middle">添加投票</td>
					<form name="form1" method="post" action="">
                    <td width="22%" align="left" valign="middle">
                      <input type="text" name="title">                    </td> 
                    <td width="8%" align="center" valign="middle"><input name="Click" type="text" id="Click" value="0" size="5"></td>
                    <td width="61%" align="left"><input type="submit" name="Submit" value="添加"></td>
					</form>
                  </tr>
				  <%
				  else
				  id=request("id")
				  set rs=conn.execute("select * from Vote where id="&id)
				  %>
				  
                  <tr>
                    <td width="9%" height="30" align="center" valign="middle">添加投票</td>
					<form name="form1" method="post" action="?send=editp&id=<%=rs("id")%>">
                    <td width="22%" align="left" valign="middle">
                      <input type="text" name="titlet" value="<%=rs("title")%>"></td> 
                    <td width="8%" align="center" valign="middle">
					<input name="Click" type="text" id="Click" value="<%=rs("Click")%>" size="5"></td>
                    <td width="61%" align="left"><input type="submit" name="Submit" value="修改"></td>
					</form>
                  </tr>
				  <%end if%>
              </table></td>
          </tr>
      </table>
    </fieldset></td>
  </tr>
</table>
</body>
</html>
