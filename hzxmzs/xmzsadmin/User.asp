<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	ID=Trim(Request.QueryString("id"))
	IF Request.Cookies("ID") = ID Then
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n����ɾ�����ڵ�¼�û�!');history.go(-1)</script>"
		Response.End
	Else
	call del("Admin",id,"User.asp")
	call alert("ɾ���ɹ�","?")
	End if
End if

If Trim(Request.Form("Action"))="Del" then
	Call Check_Right(4)	'��֤�û�Ȩ��
	call del_all("Admin","?")
End if

IF Trim(Request.QueryString("Action"))="Lock" then
	Call Check_Right(5)	'��֤�û�Ȩ��
	ID=Trim(Request.QueryString("id"))
	locks=cint(Request.QueryString("Locks"))
	IF Request.Cookies("ID") = ID Then
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n�����������ڵ�¼�û�!');history.go(-1)</script>"
		Response.End
	Else
	call Lock_Url("Admin",id,locks,"?")
	call alert("�����ɹ�","?")
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
             <input type="button" name="Submit" value="���" onClick="location='User_Insert.asp'">
             &nbsp;
             <input type="button" name="Submit2" value="ˢ��" onClick="javascript:location.reload();">
           </label></td>
         </tr>
       </table>
	   		  <form name="form2" method="post" action=""> 
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DBF9D0">
 
        <tr bgcolor="#FFFFFF">
          <td width="50" height="22" align="center">���</td>
          <td width="120" align="center">�û���</td>
          <td width="100" align="center">����</td>
          <td width="80" align="center">��¼����</td>
          <td align="center">�û�˵��</td>
          <td width="30" align="center">�鿴</td>
          <td width="30" align="center">Ȩ��</td>
          <td width="30" align="center">�˵�</td>
          <td width="30" align="center">�޸�</td>
          <td width="30" align="center">����</td>
          <td width="30" align="center">ɾ��</td>
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
		  <img src="../sysImages/Admin.gif" width="16" height="16" align="absmiddle" border="0" alt="��������Ա">
		  <%Else%>
		  <img src="../sysImages/User.gif" width="16" height="16" align="absmiddle" border="0" alt="��ͨ����Ա">
		  <%End if%>
		  </td>
          <td align="center"><%= rs("name") %></td>
          <td align="center"><%=rs("Counts")%></td>
          <td align="center"><%=rs("Content")%></td>
          <td align="center"><a href="User_Show.asp?id=<%=rs("id")%>"><img src="../sysImages/Search.gif" width="20" height="20" align="absmiddle" border="0" alt="�鿴��Ա��Ϣ"></a></td>
          <td align="center">
		  <% IF rs("Flag")=0 Then %>
		  <a href="Right.asp?id=<%=rs("id")%>"><img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="��Ȩ"></a>		  
		  <% Else %>
		  <% IF RS("ID")=Cint(Request.cookies("id")) Then %>
		  <a href="Right.asp?id=<%=rs("id")%>"><img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="��Ȩ"></a>
		  <% Else %>
		  <img src="../sysImages/Menu.gif" border="0" align="absmiddle" alt="��ʾ:���ܸ�ϵͳ����Ա��Ȩ!" onClick="javascript:window.alert('ϵͳ��ʾ:\n\n���ܸ�ϵͳ����Ա��Ȩ��')">
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
		  <img src="../sysImages/Menu.gif" width="16" height="16" align="absmiddle" border="0" alt="��ʾ:���ܸ�ϵͳ����Ա��Ȩ�˵���" onClick="javascript:window.alert('ϵͳ��ʾ:\n\n���ܸ�ϵͳ����Ա��Ȩ�˵���')">
		  <% End if %>
		  <% End if %>
		  
		  </td>
		  
		  
		  
		  
          <td align="center">
		  <% IF RS("Flag")=0 Then %>
		  <a href="User_Edit.asp?ID=<%=rs("ID")%>"><img src="../sysImages/Edit.gif" width="12" height="12" align="absmiddle" border="0" alt="�޸��û���Ϣ"></a>
		  <% Else %>
		  <% IF rs("id")=Cint(request.cookies("id")) Then %>
		  <a href="User_Edit.asp?ID=<%=rs("ID")%>"><img src="../sysImages/Edit.gif" width="12" height="12" align="absmiddle" border="0" alt="�޸��û���Ϣ"></a>
		  <% Else %>
		  <img src="../sysImages/Edit.gif" alt="��ʾ:�����޸�ϵͳ����Ա�ʺţ�" width="12" height="12" border="0" onClick="javascript:window.alert('ϵͳ��ʾ:\n\n�����޸�ϵͳ����Ա�ʺţ�')">
		  <% End if %>
		  <%End if%>
		  </td>
          
		  
		  
		  
		  
		  <td align="center">
		 <% IF Rs("Flag") = 0 Then%> 
		<% If Rs("Lock") = 0 Then %> 
		<a href="?Action=Lock&ID=<%= rs("ID") %>&Locks=1"> 
            <img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="����">            </a> 
			<% Else %>
			 <a href="?Action=Lock&ID=<%= rs("ID") %>&Locks=0"> 
            <img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="����">            </a> 
			<% End If %><% Else %>
			<img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="��ʾ:��������ϵͳ����Ա�ʺţ�" onClick="javascript:window.alert('ϵͳ��ʾ:\n\n��������ϵͳ����Ա�ʺţ�')">
			<% End if %>
			</td>
					  
					  
					  
          <td align="center">
		  <% IF rs("Flag")=0 Then %>
		  <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&ID=<%= rs("id") %>';return true;}return false;}"> 
		  <% Else %>
		  <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="javascript:window.alert('ϵͳ��ʾ:\n\n����ɾ��ϵͳ����Ա�ʺţ�')" alt="��ʾ:����ɾ��ϵͳ����Ա�ʺţ�">
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
	  <input type="submit" name="Submit2" value="ɾ��">
	  </label>
		  <input name="Action" type="hidden" id="Action" value="Del">
		  </form>
    </fieldset>
	   
    </td>
  </tr>
</table>

</body>
</html>
