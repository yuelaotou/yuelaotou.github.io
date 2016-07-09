<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'验证用户权限
	ID=Trim(Request.QueryString("id"))
	IF Request.Cookies("ID") = ID Then
		Response.Write "<script>alert('系统提示:\n\n不能删除正在登录用户!');history.go(-1)</script>"
		Response.End
	Else
	call del("Admin",id,"User.asp")
	call alert("删除成功","?")
	End if
End if

If Trim(Request.Form("Action"))="Del" then
	Call Check_Right(4)	'验证用户权限
	call del_all("Admin","?")
End if

IF Trim(Request.QueryString("Action"))="Lock" then
	Call Check_Right(5)	'验证用户权限
	ID=Trim(Request.QueryString("id"))
	locks=cint(Request.QueryString("Locks"))
	IF Request.Cookies("ID") = ID Then
		Response.Write "<script>alert('系统提示:\n\n不能锁定正在登录用户!');history.go(-1)</script>"
		Response.End
	Else
	call Lock_Url("Admin",id,locks,"?")
	call alert("操作成功","?")
	End if
End if

sql="select * from Admin order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
rs.pagesize=5
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
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong><img src="../SysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" />
	  
	  
	  
	  </strong>
	  <% 
	  SS_ID=Trim(Request.QueryString("SS_ID"))
	  CALL Type_Title(SS_ID)
	   %>
</legend>
	   
	   <table width="100%" border="0" cellspacing="1" cellpadding="4">
         <tr>
           <td width="50%" align="center"><%Call Image_Page("")%></td>
           <td align="center"><label>
             <input type="button" name="Submit" value="添加" onClick="location='User_Insert.asp'">
             &nbsp;
             <input type="button" name="Submit2" value="刷新" onClick="javascript:location.reload();">
           </label></td>
         </tr>
       </table>
	   		  <form name="form2" method="post" action=""> 
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DBF9D0">
 
        <tr bgcolor="#FFFFFF">
          <td width="50" height="22" align="center">编号</td>
          <td width="120" align="center">用户名</td>
          <td width="100" align="center">姓名</td>
          <td width="80" align="center">登录次数</td>
          <td align="center">用户说明</td>
          <td width="30" align="center">查看</td>
          <td width="30" align="center">权限</td>
          <td width="30" align="center">菜单</td>
          <td width="30" align="center">修改</td>
          <td width="30" align="center">锁定</td>
          <td width="30" align="center">删除</td>
          <td width="30" align="center"><label>
            <input type="checkbox" name="chkAll" value="checkbox" onClick="CheckAll(this.form)" >
          </label></td>
        </tr>
		
<%

For i=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
		
        <tr bgcolor="#FFFFFF" style="cursor:hand" OnMouseOver="this.bgColor='#F0F0F0'" OnMouseOut="this.bgColor='#FFFFFF'">
          <td height="22" align="center"><%= rs("id") %></td>
          <td align="center">
		  <%IF RS("Flag")=1 Then%>
		  <img src="../sysImages/Admin.gif" width="16" height="16" align="absmiddle" border="0" alt="超级管理员">
		  <%Else%>
		  <img src="../sysImages/User.gif" width="16" height="16" align="absmiddle" border="0" alt="普通管理员">
		  <%End if%>
		  </td>
          <td align="center"><%= rs("name") %></td>
          <td align="center"><%=rs("Counts")%></td>
          <td align="center"><%=rs("Content")%></td>
          <td align="center"><a href="User_Show.asp?id=<%=rs("id")%>"><img src="../sysImages/Search.gif" width="20" height="20" align="absmiddle" border="0" alt="查看会员信息"></a></td>
          <td align="center">
		  <% IF rs("Flag")=0 Then %>
		  <a href="Right.asp?id=<%=rs("id")%>"><img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="授权"></a>		  
		  <% Else %>
		  <% IF RS("ID")=Cint(Request.cookies("id")) Then %>
		  <a href="Right.asp?id=<%=rs("id")%>"><img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="授权"></a>
		  <% Else %>
		  <img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="提示:不能给系统管理员授权!" onClick="javascript:window.alert('系统提示:\n\n不能给系统管理员授权！')">
		  <% End if %>
		  <% End if %>
		  </td>
		  
          <td align="center">
		  
		  <% IF rs("Flag")=0 Then %>
		  <a href="User_menu.asp?ID=<%= rs("id") %>"><img src="../sysImages/Menu.gif" width="16" height="16" align="absmiddle" border="0"></a> 
		  <% Else %>
		  <% IF RS("ID")=Cint(request.Cookies("ID")) Then %>
		  <a href="User_menu.asp?ID=<%= rs("id") %>"><img src="../sysImages/Menu.gif" width="16" height="16" align="absmiddle" border="0"></a>
		  <% Else %>
		  <img src="../sysImages/Menu.gif" width="16" height="16" align="absmiddle" border="0" alt="提示:不能给系统管理员授权菜单！" onClick="javascript:window.alert('系统提示:\n\n不能给系统管理员授权菜单！')">
		  <% End if %>
		  <% End if %>
		  
		  </td>
		  
		  
		  
		  
          <td align="center">
		  <% IF RS("Flag")=0 Then %>
		  <a href="User_Edit.asp?ID=<%=rs("ID")%>"><img src="../sysImages/Edit.gif" width="12" height="12" align="absmiddle" border="0" alt="修改用户信息"></a>
		  <% Else %>
		  <% IF rs("id")=Cint(request.cookies("id")) Then %>
		  <a href="User_Edit.asp?ID=<%=rs("ID")%>"><img src="../sysImages/Edit.gif" width="12" height="12" align="absmiddle" border="0" alt="修改用户信息"></a>
		  <% Else %>
		  <img src="../sysImages/Edit.gif" alt="提示:不能修改系统管理员帐号！" width="12" height="12" border="0" onClick="javascript:window.alert('系统提示:\n\n不能修改系统管理员帐号！')">
		  <% End if %>
		  <%End if%>
		  </td>
          
		  
		  
		  
		  
		  <td align="center">
		 <% IF Rs("Flag") = 0 Then%> 
		<% If Rs("Lock") = 0 Then %> 
		<a href="?Action=Lock&ID=<%= rs("ID") %>&Locks=1"> 
            <img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="锁定">            </a> 
			<% Else %>
			 <a href="?Action=Lock&ID=<%= rs("ID") %>&Locks=0"> 
            <img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="解锁">            </a> 
			<% End If %><% Else %>
			<img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="提示:不能锁定系统管理员帐号！" onClick="javascript:window.alert('系统提示:\n\n不能锁定系统管理员帐号！')">
			<% End if %>
			</td>
					  
					  
					  
          <td align="center">
		  <% IF rs("Flag")=0 Then %>
		  <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&ID=<%= rs("id") %>';return true;}return false;}"> 
		  <% Else %>
		  <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="javascript:window.alert('系统提示:\n\n不能删除系统管理员帐号！')" alt="提示:不能删除系统管理员帐号！">
		  <% End if %>
		  
		  </td>
          <td align="center"><label>
            <input name="ID" type="checkbox" id="ID" value="<%= rs("id") %>">
          </label></td>
        </tr>
		
        <% 
rs.MoveNext
NEXT
%>
      </table>
	  <label>
	  <input type="submit" name="Submit2" value="删除">
	  </label>
		  <input name="Action" type="hidden" id="Action" value="Del">
		  </form>
    </fieldset>
	   
    </td>
  </tr>
</table>

</body>
</html>
