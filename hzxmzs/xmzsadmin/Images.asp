<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
'ɾ������
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	ID = Trim(Request.QueryString("ID"))
	SS_ID = Request.Form("SS_ID")
	Call Del("Images",ID,"?SS_ID="&SS_ID&"")
End if 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��������
If Trim(Request.QueryString("Action"))="Lock" Then
	Call Check_Right(5)	'��֤�û�Ȩ��
	ID = Trim(Request.QueryString("ID"))
	SS_ID = Request.QueryString("SS_ID")
	Locks = Trim(Request.QueryString("Locks"))
	ReCord = Trim(Request.QueryString("ReCord"))
	Call Lock_OK("Images",ID,ReCord,Locks,"?SS_ID="&SS_ID&"")
End if 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SS_ID = Request.QueryString("SS_ID")
set Rs=server.createobject("adodb.recordset")
Sql="select * from Images Where SS_ID="&SS_ID&" order by id desc"
Rs.open sql,conn,1,1
Rs.pagesize=30
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
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong>ͼƬ����</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> ͼƬ���</td>
            <td width="150"><strong><font color="#FF0000">
              <% 
set Rc=server.createobject("adodb.recordset")
Sql="select * from Type Where ID="&SS_ID&""
Rc.open sql,conn,1,1
Response.Write(Rc("Title"))
Rc.close
SET Rc=nothing
%>
            </font></strong> </td>
            <td width="100" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>

      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr bgcolor="#FFFFFF">
          <td width="50" height="22" align="center">���</td>
          <td width="88" align="center">ͼƬ</td>
          <td align="center">����</td>
          <td width="80" align="center">����ʱ��</td>
          <td width="50" align="center">���</td>
          <td width="30" align="center" bgcolor="#FFFFFF">ͼƬ</td>
          <td width="30" align="center">�޸�</td>
          <td width="30" align="center">����</td>
          <td width="30" align="center">ɾ��</td>
        </tr>
        <%
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
        <tr bgcolor="#FFFFFF" style="cursor:hand" OnMouseOver="this.bgColor='#F0F0F0'" OnMouseOut="this.bgColor='#FFFFFF'">
          <td height="22" align="center"><%= rs("id") %></td>
          <td align="center"><% If Rs("Img")<>Empty Then %>
              <img src="../UploadFile/<%= Rs("Img") %>" alt="" width="200" height="120">
          <% End If %></td>
          <td width="100">&nbsp;<%= rs("Title") %></td>
          <td width="80" align="center"><%= rs("DateTime") %></td>
          <td align="center"><%= rs("Counts") %></td>
          <td align="center"><% If Rs("Img")<>Empty Then %>
              <img src="../sysImages/img.gif" width="20" height="20" border="0" align="absmiddle">
              <% End If %>          </td>
          <td align="center"><a href="Images_Edit.asp?SS_ID=<%= SS_ID %>&ID=<%= rs("ID") %>"><img src="../sysImages/Edit.gif" width="12" height="12" border="0" alt="�޸�"></a> </td>
          <td align="center"><% If Rs("Lock") = 0 Then %>
              <a href="?Action=Lock&SS_ID=<%= SS_ID %>&ID=<%= rs("ID") %>&ReCord=Lock&Locks=1"> <img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="����"> </a>
              <% Else %>
              <a href="?Action=Lock&SS_ID=<%= SS_ID %>&ID=<%= rs("ID") %>&ReCord=Lock&Locks=0"> <img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="����"> </a>
              <% End If %>          </td>
          <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&SS_ID=<%= SS_ID %>&ID=<%= rs("id") %>';return true;}return false;}"> </td>
        </tr>
        <% 
rs.MoveNext
NEXT
%>
      </table>
      <% Call RS_Empty()	'û�м�¼ %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %></td>
          <td align="center"><input type="button" name="Submit_OK" value="���" onClick="location='Images_Insert.asp?SS_ID=<%= SS_ID %>'">
            &nbsp;
            <input type="button" name="reset" value="ˢ��" onClick="javascript:location.reload();">
          </td>
        </tr>
      </table>
    </fieldset></td>
  </tr>
</table>
</body>
</html>

