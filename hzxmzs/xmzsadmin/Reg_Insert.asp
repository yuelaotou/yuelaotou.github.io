<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/MD5.Asp" -->

<% 
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'验证用户权限
	UID = Trim(Request.Form("UID"))
	Flag=conn.execute("select count(*) from Reg where UID='"&UID&"'")(0)
	if flag <>0 then
		call OutScript("用户名已存在")
	else
		PWD = Trim(Request.Form("PWD"))
		PWD = md5(PWD)
		Real_name = Trim(Request.Form("Real_name"))
		email = Trim(Request.Form("email"))
		qq = Trim(Request.Form("qq"))
		tel = Trim(Request.Form("tel"))
		url = Trim(Request.Form("url"))
		address = Trim(Request.Form("address"))
		
		sql = "INSERT INTO reg(UID,PWD,Real_name,email,qq,tel,url,address) VALUES('"&UID&"','"&PWD&"','"&Real_name&"','"&email&"','"&qq&"','"&tel&"','"&url&"','"&address&"')"
		conn.execute(sql)
		call alert("用户增加成功","reg.asp")
	end if
End If

 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/check.JS"></script>

</head>

<body>


<form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td> <fieldset style="padding: 5" class="body">
        <legend><strong> 模块频道</strong></legend>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="23" align="center" bgcolor="#FFFFFF">帐号</td>
            <td bgcolor="#FFFFFF"><label>
              <input name="UID" type="text" id="UID" size="20">
            </label></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">密码</td>
            <td bgcolor="#FFFFFF"><input name="PWD" type="password" id="PWD" size="20"></td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">姓名</td>
            <td bgcolor="#FFFFFF"><label>
              <input name="Real_Name" type="text" id="Real_Name" size="20">
            </label></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">邮件</td>
            <td bgcolor="#FFFFFF"><input name="Email" type="text" id="Email" size="30"></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">qq</td>
            <td bgcolor="#FFFFFF"><input name="qq" type="text" id="qq" size="20"></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">电话</td>
            <td bgcolor="#FFFFFF"><input name="Tel" type="text" id="Tel" size="20"></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">网址</td>
            <td bgcolor="#FFFFFF"><input name="url" type="text" id="url" size="30"></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">地址</td>
            <td bgcolor="#FFFFFF"><input name="address" type="text" id="address" size="30"></td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="添加" >
              &nbsp;
              <input type="button" name="reset" value="返回" onClick="javascript:history.go(-1);">
              <input name="Action" type="hidden" id="Action" value="Insert"></td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
