<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
'''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	ID=Trim(Request.QueryString("ID"))
	Call Del("message",ID,"?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Lock" Then
ID=Trim(Request.QueryString("ID"))
Locks=Trim(Request.QueryString("Locks"))
Call Lock_url("Books",id,Locks,"?")
Call alert("�����ɹ�","?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="Check" Then
ID=Trim(Request.QueryString("ID"))
Call Check_url("message",id,"1","?")
Call alert("���ͨ��","?")
End if
'''''''''''''''''''''''''''''''''''''''''''''
IF Trim(Request.QueryString("Action"))="NCheck" Then
ID=Trim(Request.QueryString("ID"))
Call Check_url("message",id,"0","?")
Call alert("�������","?")
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
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>
      <fieldset style="padding: 5" class="body">
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
            <td>
              <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
                <tr bgcolor="#FFFFFF">
                  <td height="22" align="center">���</td>
                  <td align="center">����</td>
                  <td align="center">�绰</td>
                  <td align="center">����</td>
                  <td align="center">����</td>
                  <td align="center">����ظ�</td>
                  <td align="center">�Ƿ����</td>
                  <td align="center" bgcolor="#FFFFFF">���ʱ��</td>
                  <td align="center" bgcolor="#FFFFFF">�鿴</td>
                  <td align="center" bgcolor="#FFFFFF">���</td>
                  <td align="center">ɾ��</td>
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
					Response.Write "��"
					else
					Response.Write "��"
					end if
				%>
                  </td>
                  <td width="10%" align="center"><%= rs("DateTime") %></td>
                  <td width="5%" align="center"><a href="showzxyy.asp?id=<%=rs("id")%>"><img src="../sysImages/Search.gif" width="20" height="20" border="0"></a></td>
                  <td width="5%" align="center">
				  <%
					if rs("status")=0 then
					%>
					<img src="../sysImages/lock.gif" alt="���ͨ��" width="16" height="16" border="0" onClick="{if(confirm('ȷ�����ͨ��?')){location='?Action=Check&amp;ID=<%= rs("id") %>';return true;}return false;}" /> 
					<%
					else
					%>
					<img src="../sysImages/Unlock.gif" alt="�������" width="16" height="16" border="0" onClick="{if(confirm('ȷ���������?')){location='?Action=NCheck&amp;ID=<%= rs("id") %>';return true;}return false;}" /> 
					<%
					end if
				%>
				  
				  </td>
				  
				  
                  <td width="5%" align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&amp;ID=<%= rs("id") %>';return true;}return false;}" /> </td>
                </tr>
                <% 
				'if i>rs.pagesize then exit do
			  rs.movenext
			  loop

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
