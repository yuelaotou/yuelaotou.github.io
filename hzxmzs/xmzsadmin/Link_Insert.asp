<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'验证用户权限
	Rs,Sql
	SS_ID = Trim(Request.Form("SS_ID"))
	Title=Trim(Request.Form("Title"))
	img=Trim(Request.Form("Img"))
	URL = Trim(Request.Form("URL"))
	Content = Trim(Request.Form("Content")) 
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from Link"
	Rs.open sql,conn,1,3
	Rs.Addnew
	Rs("Title") = Title
	Rs("Img") = img
	Rs("URL") = URL
	Rs("Content") = Content
	Rs("SS_ID") = SS_ID
	Rs("Counts")=Rs("Counts")+1
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('添加成功!');location='link.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<script language="JavaScript" type="text/javascript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/NewImage.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/Date.JS"></script>


</head>

<body oncontextmenu="return false">
<form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td bgcolor="#F0E7CF"><fieldset style="padding: 5">
        <legend><strong> 友情链接</strong></legend>
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
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">标题：</td>
            <td bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" size="30" />
              &nbsp;&nbsp; </td>
            <td width="170" rowspan="4" align="center" valign="middle" bgcolor="#FFFFFF"><img name="logo" width="88" height="31" border="0" id="logo" style="CURSOR: hand" title="提示:查看图片实际大小" onClick="newimg(this.src);" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">图片：</td>
            <td bgcolor="#FFFFFF"><input name="Img" type="text" id="Img" size="35" />
                <input name="Submit_Open" type="button" id="Submit_Open" value="上传" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;" />
              【88×31】</td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">链接：</td>
            <td bgcolor="#FFFFFF"><input name="URL" type="text" id="URL" size="35" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">描述：</td>
            <td bgcolor="#FFFFFF"><textarea name="Content" cols="40" rows="8" id="Content"></textarea></td>
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
