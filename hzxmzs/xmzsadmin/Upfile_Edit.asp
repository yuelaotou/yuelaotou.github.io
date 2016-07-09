<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
Dim RS,Sql
IF Trim(Request.Form("Action"))="Edit" then
	Call Check_Right(3)	'验证用户权限
	set RS=Server.createObject("adodb.recordset")
	Sql="select Top 1 * from UpFile"
	RS.open sql,conn,1,3
		RS("JpegPath")=Request.Form("JpegPath")
		RS("JpegType")=Request.Form("JpegType")
		RS("Jpegsize")=Request.Form("Jpegsize")
		RS("JpegWidth")=Request.Form("JpegWidth")
		RS("JpegHeight")=Request.Form("JpegHeight")
		RS("ImageWidth")=Request.Form("ImageWidth")
		RS("ImageHeight")=Request.Form("ImageHeight")
		RS("ImageLogo")=Request.Form("Img")
		RS("ImageAlpha")=Request.Form("ImageAlpha")
		RS("FontColor")=Right(Request.Form("FontColor"),6)
		RS("FontFamily")=Request.Form("FontFamily")
		RS("fontsize")=Request.Form("fontsize")
		RS("fontWidth")=Request.Form("fontWidth")
		RS("fontHeight")=Request.Form("fontHeight")
		RS("fontText")=Request.Form("fontText")
		RS("JpegLock")=Request.Form("JpegLock")
		RS("LogoLock")=Request.Form("LogoLock")
		RS("fontLock")=Request.Form("fontLock")
		RS.update
		RS.requery
		RS.close
		set RS=nothing
		Response.Write "<script>alert('修改成功!');location='?';</script>"
		Response.End
End IF

IF Request.QueryString("Action")="DelImage" then
	Call Check_Right(4)	'验证用户权限
	Call Del_File("UpFile",1,"ImageLogo","?")
End IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sql="select Top 1 * from UpFile"
set RS=Server.createObject("adodb.recordset")
RS.open sql,conn,1,1

 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" src="../JS/SelectColor.JS"></script>
<script language="JavaScript" src="../JS/NewImage.JS"></script>
<script language="JavaScript" src="Check.JS"></script>
</head>

<body oncontextmenu="return false">

<form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
      <td> 
        <fieldset style="padding: 5" class="body">
        <legend><strong>上传设置</strong></legend>
      
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>路径格式</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">上传图片路径：</td>
                  <td><input name="JpegPath" type="text" id="JpegPath" value="<%= RS("JpegPath") %>" size="30"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">上传图片格式：</td>
                  <td><input name="JpegType" type="text" id="JpegType" size="30" value="<%= RS("JpegType") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">上传图片大小：</td>
                  <td><input name="Jpegsize" type="Jpegsize" id="File" value="<%= RS("Jpegsize") %>" size="10" maxlength="10"> &nbsp;10485760[10M]</td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table>
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>缩略图</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">缩略图宽度：</td>
                  <td><input name="JpegWidth" type="text" id="JpegWidth" size="4" maxlength="3" value="<%= RS("JpegWidth") %>"> 
                    &nbsp;[正整数]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">缩略图高度：</td>
                  <td><input name="JpegHeight" type="text" id="JpegHeight" size="4" maxlength="3" value="<%= RS("JpegHeight") %>"> 
                    &nbsp;[正整数]</td>
                </tr>
                <tr>
                  <td height="22" align="center">使用缩略图：</td>
                  <td>是
<input name="JpegLock" type="radio" value="1" <% IF RS("JpegLock")=1 Then %>checked<% End IF %>>
                    否
<input type="radio" name="JpegLock" value="0"<% IF RS("JpegLock")=0 Then %>checked<% End IF %>></td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table>
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>水印</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">水印在图片上的X坐标：</td>
                  <td width="200"> <input name="ImageWidth" type="text" id="ImageWidth" size="4" maxlength="3" value="<%= RS("ImageWidth") %>"> 
                    &nbsp;[正整数]</td>
                  <td rowspan="7" align="center"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                      <tr> 
                        <td align="center"> <% IF RS("ImageLogo")<>Empty Then %>
                          <img src="<%= RS("JpegPath") %>/<%= RS("ImageLogo") %>" name="logo" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" width="120" height="120" onload="javascript:DrawImage(this);"> 
                          <% Else %>
                          <img src="../sysImages/NoImage.gif" name="logo" width="120" height="120" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" title="提示:查看图片实际大小">	
                          <% End IF %> </td>
                      </tr>
                      <tr> 
                        <td height="20" align="center"><% IF RS("ImageLogo")<>Empty Then %>
                          <a href="#" onClick="{if(confirm('确定删除吗?')){location='?Action=DelImage';;return true;}return false;}">【删除图片】</a>
                          <% End IF %></td>
                      </tr>
                    </table> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">水印在图片上的Y坐标：</td>
                  <td><input name="ImageHeight" type="text" id="ImageHeight" size="4" maxlength="3" value="<%= RS("ImageHeight") %>"> 
                    &nbsp;[正整数]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">水印图标：</td>
                  <td><input name="Img" type="text" id="Img" size="18" value="<%= RS("ImageLogo") %>" readonly="0">
<input name="Submit_Open" type="button" id="Submit_Open" value="上传" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;"> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">水印透明度：</td>
                  <td> <input name="ImageAlpha" type="text" id="ImageAlpha" size="4" maxlength="3" value="<%= RS("ImageAlpha") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">使用水印：</td>
                  <td>是 
                    <input name="LogoLock" type="radio" value="1" <% IF RS("LogoLock")=1 Then %>checked<% End IF %>>
                    否 
                    <input type="radio" name="LogoLock" value="0"<% IF RS("LogoLock")=0 Then %>checked<% End IF %>></td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table>
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td>		
  <fieldset style="padding: 5">
              <legend><strong>文本</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="22" align="center">文字在图片上的X坐标：</td>
                  <td><input name="fontWidth" type="text" id="fontWidth" size="4" maxlength="3" value="<%= RS("fontWidth") %>"> 
                    &nbsp;[正整数]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">文字在图片上的Y坐标：</td>
                  <td><input name="fontHeight" type="text" id="fontHeight" size="4" maxlength="3" value="<%= RS("fontHeight") %>"> 
                    &nbsp;[正整数]</td>
                </tr>
                <tr> 
                  <td width="150" height="22" align="center">文字颜色：</td>
                  <td><input name="FontColor" type="text" id=d_bgcolor size="7" maxlength="7" value="<%= RS("FontColor") %>"> 
                    <img src="../sysImages/Rect.gIF" alt="标题颜色" name="s_bgcolor" width="18" height="17" border="0" align="absmiddle" id=s_bgcolor style="cursor:hand" onClick="SelectColor('bgcolor')"> 
                    <input type="button" name="Clear" value="清除"onclick="document.form1.FontColor.value=''" title="清除标题颜色"> 
                  </td>
                </tr>
                <tr> 
                  <td height="22" align="center">文字字体：</td>
                  <td> <select name="FontFamily" id="FontFamily">
                      <option value="宋体" <% IF RS("FontFamily")="宋体" Then %>selected<% End IF %>>宋体</option>
                      <option value="黑体" <% IF RS("FontFamily")="黑体" Then %>selected<% End IF %>>黑体</option>
                      <option value="幼圆" <% IF RS("FontFamily")="幼圆" Then %>selected<% End IF %>>幼圆</option>
                      <option value="楷体_GB2312" <% IF RS("FontFamily")="楷体_GB2312" Then %>selected<% End IF %>>楷体</option>
                      <option value="华文行楷" <% IF RS("FontFamily")="华文行楷" Then %>selected<% End IF %>>华文行楷</option>
                      <option value="华文彩云" <% IF RS("FontFamily")="华文彩云" Then %>selected<% End IF %>>华文彩云</option>
                    </select> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">文字字号：</td>
                  <td>
<select name="fontsize" id="fontsize">
<% for i=1 to 30 %>
           <option value="<%= i %>" <% If Cint(RS("fontsize"))=i Then %>selected<% End IF %>><%= i %></option>
					  <% next %>
</select>
				  </td>
                </tr>
                <tr> 
                  <td height="22" align="center">文&nbsp;&nbsp;&nbsp;&nbsp;本：</td>
                  <td><input name="fontText" type="text" id="fontText" size="30" value="<%= RS("fontText") %>"></td>
                </tr>
                <tr>
                  <td height="22" align="center">使用文本：</td>
                  <td>是 
                    <input name="fontLock" type="radio" value="1" <% IF RS("fontLock")=1 Then %>checked<% End IF %>>
                    否 
                    <input type="radio" name="fontLock" value="0"<% IF RS("fontLock")=0 Then %>checked<% End IF %>></td>
                </tr>
              </table>
</fieldset>
</td>
          </tr>
        </table>		
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
            <td height="30" align="center"> 
<input type="submit" name="Submit_OK" value="修改(E)" accesskey="E">
              &nbsp; <input type="Reset" name="reset" value="刷新(R)" accesskey="R" onClick="javascript:location.reload();"> 
              <input name="Action" type="hidden" id="Action" value="Edit">

		    </td>
  </tr>
</table>

</fieldset>
      </td>
  </tr>
</table>
</form>


</body>
</html>
<%
call DB_CLOSE()
CALL RS_CLOSE()
%>
