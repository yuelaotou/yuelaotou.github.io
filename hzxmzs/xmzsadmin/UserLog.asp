<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'循环删除
If Trim(Request.Form("Action"))="Del_Loop" Then
	Call Check_Right(4)	'验证用户权限
	Dim ID
	ID = Trim(Request.Form("ID"))
		Call Del_all("UserLog","?")
End if 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除单个
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'验证用户权限
	ID = Trim(Request.QueryString("ID"))
	Call Del("UserLog",ID,"?")
End if 


set Rs=server.createobject("adodb.recordset")
Sql="select * from UserLog order by Id desc"
Rs.open sql,conn,1,1
Rs.pagesize=24
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
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>

</head>

<body>
<form name="form1" method="post" action="">
  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td align="center"><fieldset style="padding: 5" class="body">
        <legend><strong>系统日志</strong></legend>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr bgcolor="#FFFFFF">
            <td width="50" height="22" align="center">编号</td>
            <td width="60" align="center">姓名</td>
            <td width="60" align="center">用户名</td>
            <td width="60" align="center">登录</td>
            <td width="120" align="center">登录时间</td>
            <td width="100" align="center">IP</td>
            <td align="center">日志内容</td>
            <td width="30" align="center">删除</td>
            <td width="30" align="center"><input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox" title="全选" /></td>
          </tr>
          <%
dim I
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
          <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#F0F0F0'" onMouseOut="this.bgColor='#FFFFFF'">
            <td height="22" align="center"><%= rs("id") %></td>
            <td align="center"><%= rs("User_Name") %></td>
            <td align="center"><%= rs("Login") %></td>
            <td align="center"><%= rs("Counts") %></td>
            <td align="center"><%= rs("Datetime") %></td>
            <td align="center"><%= rs("ip") %></td>
            <td>&nbsp;<%= rs("content") %></td>
            <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
            <td align="center"><input name="ID" type="checkbox" id="ID" value="<%= rs("id") %>" /></td>
          </tr>
          <% 
rs.MoveNext
NEXT
%>
        </table>
        <% Call RS_Empty()	'没有记录 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %>
            </td>
            <td align="center"><input type="submit" name="Submit_OK" value="删除" onClick="{if(confirm('确定删除吗?')){this.document.form1.submit();return true;}return false;}" />
              &nbsp;
              <input type="button" name="reset" value="刷新" onClick="javascript:location.reload();" />
              <input name="Action" type="hidden" value="Del_Loop" />
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
