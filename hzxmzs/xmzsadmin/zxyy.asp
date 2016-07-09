<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'验证用户权限
	ID=Trim(Request.QueryString("ID"))
	Call Del("message",ID,"?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Lock" Then
ID=Trim(Request.QueryString("ID"))
Locks=Trim(Request.QueryString("Locks"))
Call Lock_url("Books",id,Locks,"?")
Call alert("锁定成功","?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Check" Then
ID=Trim(Request.QueryString("ID"))
Call Check_url("message",id,"1","?")
Call alert("审核通过","?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="NCheck" Then
ID=Trim(Request.QueryString("ID"))
Call Check_url("message",id,"0","?")
Call alert("撤销审核","?")
End if
''''''''''''''''''''''''''''''''''''''''''''
Sql = "select * from message order by id desc "
set rs=server.CreateObject("ADODB.RecordSet")
rs.open Sql,conn,1,1
rs.pagesize=10
if Len(Request.QueryString("page"))=0 then
	page=1
else
	page=Cint(Request.QueryString("page"))
end if
rs.Absolutepage=page
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
    <td>
      <fieldset style="padding: 5" class="body">
      <legend><strong>客户留言</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> 服务类别：</td>
          <td width="150"><strong><font color="#FF0000">管理留言</font></strong> </td>
          <td width="100" align="center">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form action="" method="post" name="form1" id="form1">
          <tr>
            <td>
              <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">编号</td>
                  <td align="center">姓名</td>
                  <td align="center">电话</td>
                  <td align="center">标题</td>
                  <td align="center">内容</td>
                  <td align="center">浔美回复</td>
                  <td align="center">是否审核</td>
                  <td align="center" bgcolor="#FFFFFF">添加时间</td>
                  <td align="center" bgcolor="#FFFFFF">查看</td>
                  <td align="center" bgcolor="#FFFFFF">审核</td>
                  <td align="center">删除</td>
                </tr>
                <%
			i=1
			do while not rs.eof
		  	i=i+1
 %>
                <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#F0F0F0'" onMouseOut="this.bgColor='#FFFFFF'">
                  <td width="4%" align="center"><%= rs("id") %></td>
                  <td width="5%" align="center"><%= rs("nickname") %></td>
                  <td width="10%" align="center"><%= rs("tel") %></td>
                  <td width="12%" align="center">
                    <%
					Response.Write newleft(rs("title"),16)
				%>
                  </td>
                  <td width="16%" align="center">
                    <%
					Response.Write newleft(rs("content"),24)
				%>
                  </td>
                  <td width="16%" align="center">
                    <%
					Response.Write newleft(rs("restore"),22)
				%>
                  </td>
                  <td width="6%" align="center">
                    <%
					if rs("status")=1 then
					Response.Write "是"
					else
					Response.Write "否"
					end if
				%>
                  </td>
                  <td width="10%" align="center"><%= rs("DateTime") %></td>
                  <td width="5%" align="center"><a href="showzxyy.asp?id=<%=rs("id")%>"><img src="../sysImages/Search.gif" width="20" height="20" border="0"></a></td>
                  <td width="5%" align="center">
				  <%
					if rs("status")=0 then
					%>
					<img src="../sysImages/lock.gif" alt="审核通过" width="16" height="16" border="0" onClick="{if(confirm('确定审核通过?')){location='?Action=Check&amp;ID=<%= rs("id") %>';return true;}return false;}" /> 
					<%
					else
					%>
					<img src="../sysImages/Unlock.gif" alt="撤销审核" width="16" height="16" border="0" onClick="{if(confirm('确定撤销审核?')){location='?Action=NCheck&amp;ID=<%= rs("id") %>';return true;}return false;}" /> 
					<%
					end if
				%>
				  
				  </td>
				  
				  
                  <td width="5%" align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
                </tr>
                <% 
				'if i>rs.pagesize then exit do
			  rs.movenext
			  loop

%>
              </table>
              <% Call RS_Empty()	'没有记录 %>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="30" align="center" valign="middle">
                    <% call Image_Page("") %>
                  </td>
                  <td align="center">&nbsp;
                    <input type="button" name="reset" value="刷新" onClick="javascript:location.reload();" />
                    <input name="Action" type="hidden" value="Del_ALL" />
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </form>
      </table>
      </fieldset>
    </td>
  </tr>
</table>
</body>
</html>
