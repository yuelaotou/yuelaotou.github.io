<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'验证用户权限
	SS_ID = Trim(Request.Form("SS_ID"))
	name = Trim(Request.Form("name"))
	Img = Trim(Request.Form("Img"))
	zhiwei =Trim(Request.Form("zhiwei"))
	zili = Trim(Request.Form("zili"))
	jianjie = Trim(Request.Form("jianjie"))
	linian = Trim(Request.Form("linian")) 
	zuopin = Trim(Request.Form("zuopin")) 
	embed = Trim(Request.Form("embed")) 
	set Rs = Server.CreateObject("adodb.recordset")
	Sql="select * from Team"
	Rs.open sql,conn,1,3
	Rs.Addnew
	Rs("SS_ID") = SS_ID
	Rs("name") = name
	Rs("Img") = Img
	Rs("zhiwei") = zhiwei
	Rs("zili") = zili
	RS("jianjie") = jianjie
	Rs("linian") = linian
	Rs("zuopin") = zuopin
	Rs("embed") = embed
	Rs("date") = date()
	Rs("lock") = "0"
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('添加成功!');location='Team.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>设计团队添加</title>
<script language="JavaScript" type="text/javascript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/NewImage.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/Date.JS"></script>


</head>

<body oncontextmenu="return false">
<form name="theForm" method="post" action="">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td bgcolor="#F0E7CF"><fieldset style="padding: 5">
        <legend><strong>信息展示</strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> 频道类别：</td>
            <td width="150"><strong><font color="#FF0000">
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
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">设计师名字：</td>
            <td bgcolor="#FFFFFF"><input name="name" type="text" id="name" size="30" />
              &nbsp;&nbsp; </td>
            <td width="170" rowspan="6" align="center" valign="middle" bgcolor="#FFFFFF"><img name="logo" width="120" height="120" border="0" id="logo" style="CURSOR: hand" title="提示:查看图片实际大小" onClick="newimg(this.src);" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">设计师照片：</td>
            <td bgcolor="#FFFFFF"><input name="Img" type="text" id="Img" size="35" />
              <input name="Submit_Open" type="button" id="Submit_Open" value="上传" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">设计师职位：</td>
            <td bgcolor="#FFFFFF"><input name="zhiwei" type="text" id="zhiwei" size="30" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">设计师资历：</td>
            <td bgcolor="#FFFFFF"><input name="zili" type="text" id="zili" size="30" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">设计师简介：</td>
            <td bgcolor="#FFFFFF"><textarea name="jianjie" cols="40" rows="8" id="jianjie"></textarea></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">设计师理念：</td>
            <td bgcolor="#FFFFFF"><textarea name="linian" cols="40" rows="8" id="linian"></textarea></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">设计师作品：</td>
            <td bgcolor="#FFFFFF"><input name="zuopin" type="text" id="zuopin" size="30" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">外连作品：</td>
            <td bgcolor="#FFFFFF"><textarea name="embed" cols="40" rows="8" id="embed"></textarea></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="添加" />
              &nbsp;
              <input type="button" name="reset" value="返回"onClick="javascript:history.go(-1);" />
              <input name="Action" type="hidden" id="Action" value="Insert" />
              <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>" />
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
