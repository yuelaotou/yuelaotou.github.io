<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'ѭ��ɾ��
If Trim(Request.Form("Action"))="Del_Loop" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	Dim ID
	ID = Trim(Request.Form("ID"))
		Call Del_all("UserLog","?")
End if 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ɾ������
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
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
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>

</head>

<body>
<form name="form1" method="post" action="">
  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td align="center"><fieldset style="padding: 5" class="body">
        <legend><strong>ϵͳ��־</strong></legend>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr bgcolor="#FFFFFF">
            <td width="50" height="22" align="center">���</td>
            <td width="60" align="center">����</td>
            <td width="60" align="center">�û���</td>
            <td width="60" align="center">��¼</td>
            <td width="120" align="center">��¼ʱ��</td>
            <td width="100" align="center">IP</td>
            <td align="center">��־����</td>
            <td width="30" align="center">ɾ��</td>
            <td width="30" align="center"><input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox" title="ȫѡ" /></td>
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
            <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
            <td align="center"><input name="ID" type="checkbox" id="ID" value="<%= rs("id") %>" /></td>
          </tr>
          <% 
rs.MoveNext
NEXT
%>
        </table>
        <% Call RS_Empty()	'û�м�¼ %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %>
            </td>
            <td align="center"><input type="submit" name="Submit_OK" value="ɾ��" onClick="{if(confirm('ȷ��ɾ����?')){this.document.form1.submit();return true;}return false;}" />
              &nbsp;
              <input type="button" name="reset" value="ˢ��" onClick="javascript:location.reload();" />
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
