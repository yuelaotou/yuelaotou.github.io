<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<% 
'ɾ������
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'��֤�û�Ȩ��
	ID = Trim(Request.QueryString("ID"))
	jobid = Trim(Request.QueryString("jobid"))
	Call Del("Works",ID,"?jobid="&jobid&"")
End if 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Rs,Sql
Jobid = request.QueryString("Jobid")
set Rs=server.createobject("adodb.recordset")
Sql="select * from Works where Jobid="&Jobid&" order by id desc"
Rs.open sql,conn,1,1
Rs.pagesize=20
If Len(Request.QueryString("Page")) = 0 Then 
	Page = 1 
Else 
	Page=Cint(Request.Querystring("Page")) 
End If 
Rs.AbsolutePage=Page
%>



<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>

</head>

<body oncontextmenu="return false">

  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td> 
        <fieldset style="padding: 5" class="body">
        <legend><strong>ӦƸ��Ϣ</strong>		</legend>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
            Ƶ�����</td>
          <td> <strong><font color="#FF0000">����ӦƸ��</font></strong> </td>
        </tr>
      </table>
	  
        
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr bgcolor="#FFFFFF"> 
          <td width="50" height="22" align="center">���</td>
          <td width="80" align="center">����</td>
          <td align="center">��ַ</td>
          <td width="100" align="center">�绰</td>
          <td width="70" align="center" bgcolor="#FFFFFF">ӦƸʱ��</td>
          <td width="30" align="center">��ϸ</td>
          <td width="30" align="center">ɾ��</td>
        </tr>
        <%
dim I
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
        <tr bgcolor="#FFFFFF" style="cursor:hand" OnMouseOver="this.bgColor='#F0F0F0'" OnMouseOut="this.bgColor='#FFFFFF'"> 
          <td height="22" align="center"><%= rs("id") %></td>
          <td align="center">&nbsp;<%= rs("Name") %></td>
          <td>&nbsp;<%= rs("Adress") %></td>
          <td align="center"><%= rs("Tel") %></td>
          <td align="center"><%= rs("DateTime") %></td>
          <td align="center"><a href="Work_Show.asp?id=<%= rs("ID") %>"><img src="../sysImages/Search.gif" width="20" height="20" border="0" align="absmiddle"></a></td>
          <td align="center"> <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=Del&id=<%= rs("id") %>&jobid=<%= jobid %>';return true;}return false;}">          </td>
        </tr>
        <% 
rs.MoveNext
NEXT
%>
      </table>
<% Call RS_Empty()	'û�м�¼ %>
  
        
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="50%" height="30" align="center" valign="middle"> <% call Image_Page("") %>  </td>
        </tr>
      </table>
</fieldset>
</td>
  </tr>
</table>

</body>
</html>

<%
call rs_close()
call db_close()
%>


