<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'��֤�û�Ȩ��
	ID = Trim(Request.Form("ID"))
	SS_ID = Trim(Request.Form("SS_ID"))
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from Movie where id="&id
	Rs.open sql,conn,1,3
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("Author") = Trim(Request.Form("Author"))
	Rs("Img") = Trim(Request.Form("Img"))
	Rs("address") = Trim(Request.Form("address"))
	Rs("BeginTime") = Trim(Request.Form("BeginTime"))
	Rs("Movie") = Trim(Request.Form("Upfile"))
	Rs("Flag") = Trim(Request.Form("Flag"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("Rank") = Trim(Request.Form("Rank"))
	Rs("SS_ID") = SS_ID
	Rs("DateTime") = Date()
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('�޸ĳɹ�!');location='Movie.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ID = Request.QueryString("ID")
Sql = "Select * from Movie where id="&id
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open Sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewImage.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../js/Date.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body oncontextmenu="return false">
<script language="JavaScript">
function input_ok(){
    var text="";

//��֤�û���
    var M1 = theForm.Title.value; 
    if (M1 == ""){text=text+"[��Ӱ����] ����Ϊ�գ�\n"} 
	
//��֤�û���
    var M2 = theForm.Author.value; 
    if (M2 == ""){text=text+"[������] ����ѡ��\n"} 

//��֤����
    var M3 = theForm.address.value; 
    if (M3 == ""){text=text+"[��Ӱ����] ����Ϊ�գ�\n"} 

//��֤����
    var M4 = theForm.BeginTime.value; 
    if (M4 == ""){text=text+"[��ӳʱ��] ����Ϊ�գ�\n"} 
	
//��֤�û���
    var M5 = theForm.Upfile.value; 
    if (M5 == ""){text=text+"[��Ƶ] �����ϴ���\n"} 

//��֤�绰
   var M6 = theForm.Content.value;
   if (M6 == ""){text=text+"[����] ����Ϊ�գ�\n"} 

//��ҳ�淴����Ϣ	
   if (text == "") {
	   //document.theForm.submit();
	   }
   else {
    alert("������ʾ��\n\n" + text);
    return false;
    }           
}


</script>

<form name="theForm" method="post" action="" onSubmit="return input_ok()">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td> <fieldset style="padding: 5" class="body">
        <legend><strong>������Ƶ</strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
              Ƶ�����</td>
            <td width="150"> <strong><font color="#FF0000">
			<% 
Dim SS_ID
SS_ID = Request.QueryString("SS_ID")
set Rc=server.createobject("adodb.recordset")
Sql="select * from Type Where ID="&SS_ID&""
Rc.open sql,conn,1,1
Response.Write(Rc("Title"))
Rc.close
SET Rc=nothing
%>

			
			
			</font></strong></td>
            <td align="center">&nbsp;</td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">���ƣ�</td>
            <td bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" size="30" value="<%= RS("Title") %>"> 
              &nbsp;&nbsp; </td>
            <td width="200" rowspan="9" align="center" valign="middle" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">��Ƶ��</td>
            <td bgcolor="#FFFFFF"> <input name="Upfile" type="text" id="Upfile" size="35"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF">���ݣ�</td>
            <td bgcolor="#FFFFFF"> <textarea name="Content" cols="45" rows="5"><%= rs("Content") %></textarea> 
              &nbsp; <img src="../sysImages/sizeminus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows-=2" alt="��С�༭��"> 
              <img src="../sysImages/sizeplus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows+=2" alt="���߱༭��">            </td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"> <input type="submit" name="Submit_OK" value="���(A)" accesskey="a"> 
              &nbsp; <input type="button" name="reset" value="����(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
			  <input name="Action" type="hidden" id="Action" value="Edit">
			 <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>">
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
