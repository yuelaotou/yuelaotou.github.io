<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(2)	'验证用户权限
	ID = Trim(Request.Form("ID"))
	SS_ID= Trim(Request.Form("SS_ID"))
	SS_Path = Trim(Request.Form("SS_Path"))
	ImgLock = Trim(Request.Form("ImgLock"))
	NewsImage = Trim(Request.Form("NewsImage"))
	Title = Trim(Request.Form("Title"))
	DateTime =  Trim(Request.Form("DateTime"))
	Counts= Trim(Request.Form("Counts"))
	Content =  Trim(Request.Form("Content"))
	author = Trim(Request.Form("author"))
	NewsSource =  Trim(Request.Form("NewsSource"))
	'''''''''''''''''''''''''''''''''''''''''''''''''
	SS_ID = Trim(Request.Form("SS_ID"))
	SS_Path = Trim(Request.Form("SS_Path"))
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from News where ID="&ID
	Rs.open sql,conn,1,3
	'Rs.Addnew
	Rs("SS_ID") = SS_ID
	Rs("SS_Path") = SS_Path
	Rs("Img") = Trim(Request.Form("NewsImage"))
	If ImgLock = "" Then
		Rs("ImgLock") = 0
	Else
		Rs("ImgLock") = 1
	End If 
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("DateTime") = Trim(Request.Form("DateTime"))
	Rs("Counts") = Trim(Request.Form("Counts"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("author") = Trim(Request.Form("author"))
	Rs("NewsSource") = Trim(Request.Form("NewsSource"))
	Rs("TitleColor")=Trim(Request.Form("TitleColor"))

	Rs.Update
	Response.Write "<script>alert('修改成功!');location='News.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
	RS.close
	Set RS = nothing
	
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF Request.QueryString("Action")="Del_Img" then
	'Call Check_Right(78)	'验证用户权限
	ID = Request.QueryString("ID")
	SS_ID = Request.QueryString("SS_ID")
	Call Del_File("News",ID,"Img","News.asp?SS_ID="&SS_ID&"")
End IF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SS_ID = Request.QueryString("SS_ID")
ID = Request.QueryString("ID")
set Rs=server.createobject("adodb.recordset")
Sql="select * from News where ID="&ID&""
Rs.open sql,conn,1,1
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>

<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewImage.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/Date.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/JavaScript">
// 显示无模式对话框
function eShowDialog(url, width, height, optValidate) {
	if (optValidate) {
		if (!validateMode()) return;
	}
	var arr = showModalDialog(url, window, "dialogWidth:" + width + "px;dialogHeight:" + height + "px;help:no;scroll:no;status:no");
}
// 显示帮助
function openwin(url,width,height){
	eShowDialog(url,width,height);
	return false;
}

</script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
    <tr>
      <td><fieldset style="padding: 5" class="body">
        <legend><strong>文章频道┠</strong><font color="#FF0000">
          <% 
SS_ID = Request.QueryString("SS_ID")
set Rc=server.createobject("adodb.recordset")
Sql="select * from Type Where ID="&SS_ID&""
Rc.open sql,conn,1,1
SS_Path = rc("SS_Path")
Response.Write(Rc("Title"))
Rc.close
SET Rc=nothing
%>
          </font> </legend>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr>
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">文章性质：</td>
            <td bgcolor="#FFFFFF"><input name="ImgLock" type="checkbox" id="ImgLock" <% If rs("ImgLock") = 1 Then Response.Write("checked") End If%> title="提示:图片链接时在首页显示图片新闻">
              图片链接
              <input name="Counts" type="text" id="Counts" value="<%= rs("Counts") %>" size="4" maxlength="4">
              点击</td>
            <td width="145" rowspan="4" align="center" valign="middle" bgcolor="#FFFFFF">
			<% if rs("img")<>empty then %>
			<img src="../UploadFile/<%= rs("img") %>" name="logo" width="120" height="120" id="logo" onClick="newimg(this.src);" style="CURSOR: hand">
			<% else %>
			<img src="../sysImages/NoImage.gif" name="logo" width="120" height="120" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" title="提示:查看图片实际大小">
			<% end if %>
			<% IF RS("Img")<>Empty Then %>
<a href="#" onClick="{if(confirm('确定删除吗?')){location='?Action=Del_Img&ID=<%= ID %>&SS_ID=<%= SS_ID %>';;return true;}return false;}">【删除图片】</a>
                          <% End IF %>		    </td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">链接图片：</td>
            <td bgcolor="#FFFFFF"><input name="NewsImage" type="text" id="Img" size="20" readonly="0" title="提示:请选择上传" value="<%= rs("Img") %>">
                <input name="Submit_Open" type="button" id="Submit_Open" value="上传" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;">
              【250×150】</td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">文章标题：</td>
            <td bgcolor="#FFFFFF">
			<input name="Title" type="text" id="Title" size="50" value="<%= rs("Title") %>">
                <input name="TitleColor" type="hidden">
			 <input type="button" name="Color" value="颜色" onClick="openwin('../Include/color.htm',300,260)" title="">		    </td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">发布人：</td>
            <td bgcolor="#FFFFFF"><% 
			
			set RC=server.createobject("adodb.recordset")
			Sql="select * from Admin where uid='"& Session("Uid") &"'"
			RC.open sql,conn,1,1

			 %>
              <input name="author" type="text" id="author" value="<%= Rc("name") %>" size="8">
              <input type="button" name="Submit4" value="浔美装饰" onClick="document.theForm.author.value='浔美装饰'">
                <% 
			  RC.close
			  set Rc=Nothing
			   %>
              来源：
              <input name="NewsSource" type="text" id="NewsSource" value="<%= rs("NewsSource") %>" size="10">
              <input type="button" name="Submit4" value="本站原创"onclick="document.theForm.NewsSource.value='本站原创'">
              发布时间：
              <script language="JavaScript">arrowtag("DateTime",'<%= date() %>');</script>            </td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">文章内容：</td>
            <td colspan="2" bgcolor="#FFFFFF"><textarea name="Content" style="display:none"><%= Rs("Content") %></textarea>
                <iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width="600" HEIGHT="350"></iframe></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="修改(A)" accesskey="a">
              &nbsp;
              <input type="button" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);">
              <input name="Action" type="hidden" id="Action" value="Edit">
              <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>">
            <input name="ID" type="hidden" id="ID" value="<%= ID %>"></td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </form>
</table>
</body>
</html>
