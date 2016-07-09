<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/MD5.Asp" -->
<%
IF Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'验证用户权限
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
		Call alert("用户添加成功","User.asp")
	Else
		RS.close
		set RS=nothing
		Call OutScript("系统提示:\n\n用户名已占用")
	End if
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<script language="JavaScript" type="text/javascript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript">
function input_ok(theForm){
    var text="";
	
//验证用户名
    var M1 = theForm.UID.value; 
    if (M1 == ""){text=text+"[用户名] 不能为空；\n"} 
//验证用户名
    var M2 = theForm.PWD.value; 
    if (M2 == ""){text=text+"[密码] 不能为空；\n"} 
//验证用户名
    var M3 = theForm.sex.value; 
    if (M3 == ""){text=text+"[性别] 不能为空；\n"} 

//验证姓名
    var M4 = theForm.Name.value; 
    if (M4 == ""){text=text+"[姓名] 不能为空；\n"} 

//验证年龄
    var M5 = theForm.age.value; 
    if (M5 == "") {
		text=text+"[年龄] 不能为空；\n"} 
	else {
		if (isNaN(M5)) {text=text+"[年龄] 必须为数值；\n"}
	}

//验证电话
    var M6 = theForm.tel.value; 
    if (M6 == ""){text=text+"[电话] 不能为空；\n"} 
	

//验证Email
    var M8 = theForm.Email.value; 
    if (M8 == "")
		{text=text+"[Email] 不能为空；\n"} 
	else
		{var p=M8.indexOf('@');
		if (p<1 || p==(M8.length-1)) {text=text+"[Email] 填写不规范；\n";}
		}
//向页面反馈信息document.theForm.submit();	
   if (text == "") {}
   else {
    alert("出错提示：\n\n" + text);
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
      <legend><strong>增加管理员</strong></legend>
      <table width="60%" border="0" cellpadding="5" cellspacing="1" bgcolor="#F5F5F5">
        <tr>
          <td width="100" height="22" align="center" bgcolor="#FFFFFF">用户名：</td>
          <td bgcolor="#FFFFFF"><input name="UID" type="text" id="UID" size="12" title="不能为空" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">密&nbsp;&nbsp;码：</td>
          <td bgcolor="#FFFFFF"><input name="PWD" type="password" id="PWD" size="12" title="不能为空" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="23" align="center" bgcolor="#FFFFFF">姓&nbsp;&nbsp;名：</td>
          <td bgcolor="#FFFFFF"><input name="Name" type="text" id="Name" size="12" title="不能为空" />
            <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">头&nbsp;&nbsp;像：</td>
          <td bgcolor="#FFFFFF"><table width="120" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td><select name="Change_head" size="1" id="Change_head" onChange="showimage()">
                <option value="0">请选择</option>
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
          <td height="22" align="center" bgcolor="#FFFFFF">性&nbsp;&nbsp;别：</td>
          <td bgcolor="#FFFFFF"><select name="sex" id="sex">
            <option value="" selected="selected">性别</option>
            <option value="男">男</option>
            <option value="女">女</option>
          </select>
          <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">年&nbsp;&nbsp;龄：</td>
          <td bgcolor="#FFFFFF"><input name="age" type="text" id="age" size="2" maxlength="2" title="不能为空" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">电&nbsp;&nbsp;话：</td>
          <td bgcolor="#FFFFFF"><input name="tel" type="text" id="tel" size="20" title="不能为空" />
            <font color="#FF0000">*</font></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">QQ：</td>
          <td bgcolor="#FFFFFF"><input name="qq" type="text" id="qq" size="20" /></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">邮&nbsp;&nbsp;件：</td>
          <td bgcolor="#FFFFFF"><input name="Email" type="text" id="Email" size="35" title="不能为空" />
            <font color="#FF0000">*</font> </td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">网&nbsp;&nbsp;址：</td>
          <td bgcolor="#FFFFFF"><input name="URL" type="text" id="URL" title="不能为空" value="http://" size="35" /></td>
        </tr>
        <tr>
          <td height="22" align="center" bgcolor="#FFFFFF">简&nbsp;&nbsp;介：</td>
          <td bgcolor="#FFFFFF"><textarea name="Content" cols="35" rows="5" id="Content"></textarea></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30" align="center"><input type="submit" name="Submit_OK" value="添加" />
            &nbsp;
            <input type="button" name="reset" value="返回" onClick="javascript:history.go(-1);" />
            <input name="Action" type="hidden" id="Action" value="Insert" /></td>
        </tr>
      </table>
    </fieldset></td>
  </tr>
</table>
</form>
</body>
</html>
