<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	ID=Trim(Request.QueryString("ID"))
	Call Del("Books",ID,"?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.Form("Action")) = "Del_ALL" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	Dim ID
	ID = Trim(Request.Form("ID"))
	If ID = Empty then
		Response.write("<script>alert('ϵͳ��ʾ:\n\n��ѡ��ɾ���ļ���');location='?';</script>")
		Response.End
	Else
		Call Del_ALL("Books",ID,"?")
	End if
End if 

''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Lock" Then
ID=Trim(Request.QueryString("ID"))
Locks=Trim(Request.QueryString("Locks"))
Call Lock_url("Books",id,Locks,"?")
Call alert("�����ɹ�","?")
End if
''''''''''''''''''''''''''''''''''''''''''''
SS_ID=request.QueryString("SS_ID")
Sql = "select * from Books where SS_ID="&SS_ID&" order by id desc"
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
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>

</head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong>�ͻ�����</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> �������</td>
            <td width="150"><strong><font color="#FF0000">��������</font></strong> </td>
            <td width="100" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form action="" method="post" name="form1" id="form1">
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
              <tr bgcolor="#FFFFFF">
                <td width="50" height="22" align="center">���</td>
                <td width="100" align="center">����</td>
                <td align="center">����</td>
                <td width="150" align="center" bgcolor="#FFFFFF">����ʱ��</td>
                <td width="40" align="center">�ѻظ�</td>
                <td width="30" align="center">�ظ�</td>
                <td width="30" align="center">����</td>
                <td width="30" align="center">ɾ��</td>
                </tr>
              <%
dim I
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
              <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#F0F0F0'" onMouseOut="this.bgColor='#FFFFFF'">
                <td height="22" align="center"><%= rs("id") %></td>
                <td align="center">&nbsp;<%= rs("Name") %></td>
                <td align="center">&nbsp;<%= rs("Email") %></td>
                <td align="center"><%= rs("DateTime") %></td>
                <td align="center"><% If Rs("Flag") = 0 Then %>
                      <img src="../sysImages/Err.gif" width="12" height="11" border="0" align="absmiddle" />
                      <% Else %>
                      <img src="../sysImages/Ok.gif" width="16" height="16" border="0" align="absmiddle" />
                      <% End If %>                </td>
                <td align="center"><a href="Books_Reply.asp?ID=<%= rs("ID") %>"><img src="../sysImages/Edit.gif" width="12" height="12" border="0" alt="�ظ�����" /></a> </td>
                <td align="center"><% If Rs("Lock") = 0 Then %>
                      <a href="?Action=Lock&amp;ID=<%= rs("ID") %>&amp;ReCord=Lock&amp;Locks=1"> <img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="����" /> </a>
                      <% Else %>
                      <a href="?Action=Lock&amp;ID=<%= rs("ID") %>&amp;ReCord=Lock&amp;Locks=0"> <img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="����" /> </a>
                      <% End If %>                </td>
                <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
                </tr>
              <% 
rs.MoveNext
NEXT
%>
            </table>
                <% Call RS_Empty()	'û�м�¼ %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="50%" height="30" align="center" valign="middle">
					<% call Image_Page("") %>
                    </td>
                    <td align="center">&nbsp;
                      <input type="button" name="reset" value="ˢ��" onClick="javascript:location.reload();" />
                      <input name="Action" type="hidden" value="Del_ALL" />
                    </td></tr>
              </table></td>
          </tr>
        </form>
      </table>
    </fieldset></td>
  </tr>
</table>
</body>
</html>
