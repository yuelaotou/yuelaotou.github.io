<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'��֤�û�Ȩ��
	Rs,Sql
	ID = Trim(Request.Form("ID"))
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from [Job] Where ID="&ID&""
	Rs.open sql,conn,1,3
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("Explain") = Trim(Request.Form("Explain"))
	Rs("Address") = Trim(Request.Form("Address"))
	Rs("EndDate") = Trim(Request.Form("EndDate"))
	Rs("Number") = Trim(Request.Form("Job_Number"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("Rank") = Trim(Request.Form("Rank"))
	Rs("DateTime") = Date()
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('�޸ĳɹ�!');location='Job_List.asp';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF Request.QueryString("Action")="DelImage" then
	'Call Check_Right(193)	'��֤�û�Ȩ��
	ID = Trim(Request.QueryString("ID"))
	Call Del_File("[Job]",ID,"Image","Job_List.asp")
End IF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ID = Trim(Request.QueryString("ID"))
set Rs=server.createobject("adodb.recordset")
Sql="select * from [Job] where ID="&ID
Rs.open sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/Date.JS"></script>

</head>

<body oncontextmenu="return false">
<form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td> <fieldset style="padding: 5" class="body">
        <legend><strong> �˲���Ƹ </strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
              Ƶ�����</td>
            <td> <strong><font color="#FF0000">��ƸӢ��</font></strong> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">��λ���ƣ�</td>
            <td bgcolor="#FFFFFF">
			<input name="Title" type="text" id="Title" size="30" value="<%= Rs("Title") %>"> 
            </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">�������ڣ�</td>
            <td bgcolor="#FFFFFF"> 
<script language="JavaScript">arrowtag("EndDate",'<%= Rs("EndDate") %>');</script></td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">���ʴ�����</td>
            <td bgcolor="#FFFFFF">
<input name="Explain" type="text" id="Explain" size="30" value="<%= Rs("Explain") %>"> 
            </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">�����ص㣺</td>
            <td bgcolor="#FFFFFF">
<input name="Address" type="text" id="Address" size="30" value="<%= Rs("Address") %>"> 
            </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">��Ƹ������</td>
            <td bgcolor="#FFFFFF">
<input name="Job_Number" type="text" id="Job_Number" size="20" value="<%= Rs("Number") %>">
              [������]</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">����</td>
            <td bgcolor="#FFFFFF">
<input name="Rank" type="text" id="Rank" size="4" maxlength="4" value="<%= Rs("Rank") %>">
              [������]</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">���ݣ�</td>
            <td bgcolor="#FFFFFF">
<textarea name="Content" cols="60" rows="10"><%= Rs("Content") %></textarea> 
              &nbsp;<img src="../sysImages/sizeminus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows-=2" alt="��С�༭��"> 
              <img src="../sysImages/sizeplus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows+=2" alt="���߱༭��"> 
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"> <input type="submit" name="Submit_OK" value="�޸�(E)" accesskey="e"> 
              &nbsp; <input type="button" name="reset" value="����(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
			  <input name="Action" type="hidden" id="Action" value="Edit">
              <input name="ID" type="hidden" id="ID" value="<%= ID %>">
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
<% 
call rs_close()
call rc_close()
call db_close()
 %>
