<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.QueryString("Action"))="del" Then
	Call Check_Right(4)
	ID=Trim(Request.QueryString("id"))
	call del("Reg",id,"Reg.asp")
	call alert("删除成功","?")
End if
''''''''''''锁定'''''''''''''''''''
IF Trim(Request.QueryString("Action"))="lock" then
	Call Check_Right(5)
	ID=Trim(Request.QueryString("id"))
	locks=cint(Request.QueryString("locks"))
	call Lock_Url("Reg",id,locks,"?")
	call alert("操作成功","?")
End if
'''''''''''''''''''''''''''''''''''
SQL="select * from Reg order by id desc"
set rs=server.CreateObject("ADODB.RecordSet")
rs.open SQL,conn,1,1
rs.pagesize=10
If Len(Request.QueryString("Page")) = 0 Then 
	Page = 1 
Else 
	Page=Cint(Request.Querystring("Page")) 
End If 
Rs.AbsolutePage=Page

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>

<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong><img src="../SysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" />会员管理</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %></td>
          <td align="center"><label>
            <input name="Submit_OK" type="button" id="Submit_OK" value="添加" onClick="location='Reg_Insert.asp'" />
            &nbsp;
            <input name="reset" type="button" id="reset" value="刷新" onClick="javascript:location.reload();" />
          </label></td>
        </tr>
      </table>
      <form name="form1" method="post" action="">
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr>
            <td width="50" height="22" align="center" bgcolor="#FFFFFF">编号</td>
            <td width="100" align="center" bgcolor="#FFFFFF">用户名</td>
            <td width="100" align="center" bgcolor="#FFFFFF">真实姓名</td>
            <td width="150" align="center" bgcolor="#FFFFFF">邮箱</td>
            <td width="80" align="center" valign="middle" bgcolor="#FFFFFF">QQ</td>
            <td align="center" bgcolor="#FFFFFF">电话</td>
            <td width="30" align="center" bgcolor="#FFFFFF">修改</td>
            <td width="30" align="center" bgcolor="#FFFFFF">锁定</td>
            <td width="30" align="center" bgcolor="#FFFFFF">删除</td>
          </tr>
          <% FOR I=1 TO RS.pagesize 
		if rs.eof then exit for
		%>
          <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#FOFOFO'" onMouseOut="this.bgColor='#FFFFFF'">
            <td height="22" align="center"><%= RS("ID") %></td>
            <td align="center"><%= RS("Uid") %></td>
            <td align="center"><%= RS("Real_Name") %></td>
            <td align="center"><%= RS("Email") %></td>
            <td align="center"><a target="blank" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=<%= rs("qq") %>&amp;Site=本站&amp;Menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%= rs("qq") %>:6" alt="联系技术部" border="0" align="absmiddle" /></a></td>
            <td align="center"><%=rs("tel")%></td>
            <td align="center"><a href="Reg_Edit.asp?id=<%=rs("id")%>"><img src="../sysImages/Edit.gif" width="12" height="12" border="0" alt="修改"></a></td>
            <td align="center"><%IF rs("Lock")=0 Then%>
                <a href="?Action=lock&id=<%=rs("id")%>&locks=1"><img src="../sysImages/Unlock.gif" width="11" height="12" border="0" alt="锁定"></a>
                <%else%>
                <a href="?Action=lock&id=<%=rs("id")%>&locks=0"><img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="不锁定"></a>
                <%End if%>            </td>
            <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确认要删除吗?')){location='?Action=del&id=<%=rs("id")%>';return true;}return flase;}"></td>
          </tr>
          <%rs.movenext
		next
		%>
        </table>
        <label></label>
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
