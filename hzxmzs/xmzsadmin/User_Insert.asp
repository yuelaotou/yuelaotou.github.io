<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/MD5.Asp" -->
<%
IF Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'��֤�û�Ȩ��
	set RS=Server.createObject("adodb.recordset")
	Sql="select * from Admin where UID='"&Trim(Request.Form("UID"))&"'"
	RS.open sql,conn,1,3
	IF RS.BOF And RS.EoF Then
		RS.AddNew
		RS("UID")=Request.Form("UID")
		PWD = Request.Form("PWD")
		RS("PWD")=Md5(PWD)
		RS("Name")=Request.Form("Name")
		RS("Head")=Request.Form("Change_head")
		RS("sex")=Request.Form("sex")
		RS("age")=Request.Form("age")
		RS("tel")=Request.Form("tel")
		RS("qq")=Request.Form("qq")
		RS("Email")=Request.Form("Email")
		RS("url")=Request.Form("url")
		RS("Content")=Request.Form("Content")
		rs("ip")=Request.ServerVariables("REMOTE_ADDR")
		RS.update
		RS.requery
		RS.close
		set RS=nothing
		Call alert("�û���ӳɹ�","User.asp")
	Else
		RS.close
		set RS=nothing
		Call OutScript("ϵͳ��ʾ:\n\n�û�����ռ��")
	End if
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<script language="JavaScript" type="text/javascript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript">
function input_ok(theForm){
    var text="";
	
//��֤�û���
    var M1 = theForm.UID.value; 
    if (M1 == ""){text=text+"[�û���] ����Ϊ�գ�\n"} 
//��֤�û���
    var M2 = theForm.PWD.value; 
    if (M2 == ""){text=text+"[����] ����Ϊ�գ�\n"} 
//��֤�û���
    var M3 = theForm.sex.value; 
    if (M3 == ""){text=text+"[�Ա�] ����Ϊ�գ�\n"} 

//��֤����
    var M4 = theForm.Name.value; 
    if (M4 == ""){text=text+"[����] ����Ϊ�գ�\n"} 

//��֤����
    var M5 = theForm.age.value; 
    if (M5 == "") {
		text=text+"[����] ����Ϊ�գ�\n"} 
	else {
		if (isNaN(M5)) {text=text+"[����] ����Ϊ��ֵ��\n"}
	}

//��֤�绰
    var M6 = theForm.tel.value; 
    if (M6 == ""){text=text+"[�绰] ����Ϊ�գ�\n"} 
	

//��֤Email
    var M8 = theForm.Email.value; 
    if (M8 == "")
		{text=text+"[Email] ����Ϊ�գ�\n"} 
	else
		{var p=M8.indexOf('@');
		if (p<1 || p==(M8.length-1)) {text=text+"[Email] ��д���淶��\n";}
		}
//��ҳ�淴����Ϣdocument.theForm.submit();	
   if (text == "") {}
   else {
    alert("������ʾ��\n\n" + text);
    return false;
    }           
}
</script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body oncontextmenu="return false">
<form name="form1" method="post" action="" onSubmit="return input_ok(this)">

<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
    <td align="center"><fieldset style="padding: 5" class="body">
      <legend><strong>���ӹ���Ա</strong></legend>
      <table width="60%" border="0" cellpadding="5" cellspacing="1" bgcolor="#F5F5F5">
        <tr>
          <td width="100" height="22" align="center" bgcolor="#FFFFFF">�û�����</td>
          <td bgcolor="#FFFFFF"><input name="UID" type="text" id="UID" size="12" title="����Ϊ��" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;�룺</td>
          <td bgcolor="#FFFFFF"><input name="PWD" type="password" id="PWD" size="12" title="����Ϊ��" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;����</td>
          <td bgcolor="#FFFFFF"><input name="Name" type="text" id="Name" size="12" title="����Ϊ��" />
            <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">ͷ&nbsp;&nbsp;��</td>
          <td bgcolor="#FFFFFF"><table width="120" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td><select name="Change_head" size="1" id="Change_head" onChange="showimage()">
                <option value="0">��ѡ��</option>
                <%
		  Dim I
		   for i=1 to 16 
		   %>
                <option value="<%= i %>"><%= i %></option>
                <% next %>
              </select></td>
              <td height="40" align="center"><img src="../sysImages/head/0.gif" name="qq" width="32" height="32" align="absmiddle" id="qq" />
                    <script language="JavaScript" type="text/javascript">
function showimage()
{document.images.qq.src="../sysImages/head/"+document.form1.Change_head.options[document.form1.Change_head.selectedIndex].value+".gif";
document.form1.head.value=document.form1.Change_head.options[document.form1.Change_head.selectedIndex].value;
}
  </script>
                    <input name="head" type="hidden" id="head" value="0" /></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;��</td>
          <td bgcolor="#FFFFFF"><select name="sex" id="sex">
            <option value="" selected="selected">�Ա�</option>
            <option value="��">��</option>
            <option value="Ů">Ů</option>
          </select>
          <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;�䣺</td>
          <td bgcolor="#FFFFFF"><input name="age" type="text" id="age" size="2" maxlength="2" title="����Ϊ��" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;����</td>
          <td bgcolor="#FFFFFF"><input name="tel" type="text" id="tel" size="20" title="����Ϊ��" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">QQ��</td>
          <td bgcolor="#FFFFFF"><input name="qq" type="text" id="qq" size="20" /></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;����</td>
          <td bgcolor="#FFFFFF"><input name="Email" type="text" id="Email" size="35" title="����Ϊ��" />
            <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;ַ��</td>
          <td bgcolor="#FFFFFF"><input name="URL" type="text" id="URL" title="����Ϊ��" value="http://" size="35" /></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">��&nbsp;&nbsp;�飺</td>
          <td bgcolor="#FFFFFF"><textarea name="Content" cols="35" rows="5" id="Content"></textarea></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30" align="center"><input type="submit" name="Submit_OK" value="���" />
            &nbsp;
            <input type="button" name="reset" value="����" onClick="javascript:history.go(-1);" />
            <input name="Action" type="hidden" id="Action" value="Insert" /></td>
        </tr>
      </table>
    </fieldset></td>
  </tr>
</table>
</form>
</body>
</html>
