<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/MD5.Asp" -->

<%
IF Trim(Request.Form("Action"))="Pass" Then
	Call Check_Right(1)
	Old_Pass=Trim(Request.Form("Password"))
	Old_Pass=md5(Old_Pass)
	Pass1=Trim(Request.Form("Pass1"))
	Pass1=md5(Pass1)
	UID=Session("UID")
	Flag=conn.execute("select count(*) from admin where uid='"&uid&"' and pwd='"&Old_Pass&"'")(0)
	if Flag=0 then
		Call OutScript("ԭ�������")
	else
		sql="update admin set Pwd='"&pass1&"' where UID='"&uid&"'"
		conn.execute(sql)
		call alert("�����޸ĳɹ�","?")
	end if

End if
%>




<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<script language="JavaScript">
<!--
function input_ok(theForm){
    var text="";

	//��֤�û���
    var M1 = theForm.Pass.value; 
	if (M1 == "") {text=text+"[ԭ����] ����Ϊ�գ�\n"}
	
	//��֤�û���
    var M2 = theForm.Pass1.value;
	var L2 = M2.length; 
	if (M2 == "") {text=text+"[������] ����Ϊ�գ�\n"}
	
//��֤ȷ������
    var M3 = theForm.Pass2.value; 
    if (M3 == ""){text=text+"[ȷ������] ����Ϊ�գ�\n"} 
	else {
		if (M2 != M3){text=text+"[����] �� [ȷ������]����ͬ��\n"}
	}
	
//��֤��֤��

	//��ҳ�淴����Ϣ	
   if (text == ""){return true;}
   else{
    alert(text);
    return false;
    }           
}
-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<fieldset style="padding: 5" class="body">
              <legend><strong>�޸�����</strong></legend>
	          <form name="form1" method="post" action="" onSubmit="return input_ok(this)">
	            <table width="40%" border="0" align="center" cellpadding="5" cellspacing="1">
                  <tr>
                    <td width="100" align="center">ԭ �� ��:</td>
                    <td><label>
                      <input name="Password" type="password" id="Password" size="20">
                    </label></td>
                  </tr>
                  <tr>
                    <td align="center">�� �� ��:</td>
                    <td><label>
                      <input name="Pass1" type="password" id="Pass1" size="20">
                    </label></td>
                  </tr>
                  <tr>
                    <td align="center">ȷ������:</td>
                    <td><label>
                      <input name="pass2" type="password" id="pass2" size="20">
                    </label></td>
                  </tr>
                  <tr>
                    <td colspan="2" align="center"><input type="submit" name="Submit" value="�ύ">
                      <input name="Action" type="hidden" id="Action" value="Pass">
                      &nbsp;
                      <input type="reset" name="Submit2" value="����"></td>
                  </tr>
                </table>
      </form>
      </fieldset>
	</td>
  </tr>
</table>
</body>
</html>
