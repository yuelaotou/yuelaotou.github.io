<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.Form("Action"))="Rights" Then
	ID=Trim(Request.Form("ID"))
	Rights=Trim(Request.Form("Rights"))
	SQL = "update admin set Rights='"&Rights&"' where id="&id
	conn.execute(SQL)
	Call alert("授权成功","User.asp")

End if
'''''''''''''''
ID = Trim(Request.QueryString("ID"))
sql_menu="select * from admin where ID="&ID
Set Rc=server.CreateObject("ADODB.RecordSet")
Rc.open sql_menu,conn,1,1
if rc("Rights")<>Empty then
	Rights = split(Rc("Rights"),",")
end if


''''''''''''''''''''''''
SQL="select * from Rights"
Set Rs=server.CreateObject("ADODB.RecordSet")
Rs.open SQL,conn,1,1
Rs.pagesize=22
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

<title>无标题文档</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong><img src="../SysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" />用户授权</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %></td>
          <td align="center"><label>&nbsp;
            <input name="reset" type="button" id="reset" value="返回" onClick="javascript:history.go(-1);" />
          </label></td>
        </tr>
      </table>
      <form name="form1" method="post" action="">
        <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EBEBEB">
          <tr>
            <td width="50" align="center" bgcolor="#FFFFFF">编号</td>
            <td align="center" bgcolor="#FFFFFF">标题</td>
            <td width="30" align="center" bgcolor="#FFFFFF"><input type="checkbox" name="chkAll" value="checkbox" onClick="CheckAll(this.form)" ></td>
          </tr>
		 <%
		 For i=1 to rs.pagesize
		 if rs.eof then exit for
		 %> 
          <tr>
            <td align="center" bgcolor="#FFFFFF"><%= rs("ID") %></td>
            <td align="center" bgcolor="#FFFFFF"><%= rs("Title") %></td>
            <td align="center" bgcolor="#FFFFFF">
			<% 
		
	Flag = 0
	if Rights<>Empty then
	

    For J = LBound(Rights) To UBound(Rights)
      IF Trim(Rs("ID"))=Trim(Rights(J)) Then
	     Flag = 1
	  End If
    Next	
	end if
 %>  
			<input name="Rights" type="checkbox" id="Rights" value="<%= rs("ID")%>" <% If Flag = 1 Then %>checked<% End If %> title="编号:<%=rs("id")%>"></td>
          </tr>
		  <%
		  rs.movenext
		  Next
		  %>
        </table>
        <label>
          <input type="submit" name="Submit" value="授权">
        </label>
        <input name="Action" type="hidden" id="Action" value="Rights">
        <input name="ID" type="hidden" id="ID" value="<%= ID %>">
      </form>
    </fieldset></td>
  </tr>
</table>
<%
Call RS_Close()
Call RC_Close()
Call DB_Close()
%>

</body>
</html>
