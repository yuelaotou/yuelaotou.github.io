<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
Dim RS,Sql
IF Trim(Request.Form("Action"))="Edit" then
	Call Check_Right(3)	'��֤�û�Ȩ��
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
		Response.Write "<script>alert('�޸ĳɹ�!');location='?';</script>"
		Response.End
End IF

IF Request.QueryString("Action")="DelImage" then
	Call Check_Right(4)	'��֤�û�Ȩ��
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
        <legend><strong>�ϴ�����</strong></legend>
      
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>·����ʽ</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">�ϴ�ͼƬ·����</td>
                  <td><input name="JpegPath" type="text" id="JpegPath" value="<%= RS("JpegPath") %>" size="30"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�ϴ�ͼƬ��ʽ��</td>
                  <td><input name="JpegType" type="text" id="JpegType" size="30" value="<%= RS("JpegType") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�ϴ�ͼƬ��С��</td>
                  <td><input name="Jpegsize" type="Jpegsize" id="File" value="<%= RS("Jpegsize") %>" size="10" maxlength="10"> &nbsp;10485760[10M]</td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table>
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>����ͼ</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">����ͼ��ȣ�</td>
                  <td><input name="JpegWidth" type="text" id="JpegWidth" size="4" maxlength="3" value="<%= RS("JpegWidth") %>"> 
                    &nbsp;[������]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">����ͼ�߶ȣ�</td>
                  <td><input name="JpegHeight" type="text" id="JpegHeight" size="4" maxlength="3" value="<%= RS("JpegHeight") %>"> 
                    &nbsp;[������]</td>
                </tr>
                <tr>
                  <td height="22" align="center">ʹ������ͼ��</td>
                  <td>��
<input name="JpegLock" type="radio" value="1" <% IF RS("JpegLock")=1 Then %>checked<% End IF %>>
                    ��
<input type="radio" name="JpegLock" value="0"<% IF RS("JpegLock")=0 Then %>checked<% End IF %>></td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table>
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>ˮӡ</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">ˮӡ��ͼƬ�ϵ�X���꣺</td>
                  <td width="200"> <input name="ImageWidth" type="text" id="ImageWidth" size="4" maxlength="3" value="<%= RS("ImageWidth") %>"> 
                    &nbsp;[������]</td>
                  <td rowspan="7" align="center"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                      <tr> 
                        <td align="center"> <% IF RS("ImageLogo")<>Empty Then %>
                          <img src="<%= RS("JpegPath") %>/<%= RS("ImageLogo") %>" name="logo" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" width="120" height="120" onload="javascript:DrawImage(this);"> 
                          <% Else %>
                          <img src="../sysImages/NoImage.gif" name="logo" width="120" height="120" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" title="��ʾ:�鿴ͼƬʵ�ʴ�С">	
                          <% End IF %> </td>
                      </tr>
                      <tr> 
                        <td height="20" align="center"><% IF RS("ImageLogo")<>Empty Then %>
                          <a href="#" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=DelImage';;return true;}return false;}">��ɾ��ͼƬ��</a>
                          <% End IF %></td>
                      </tr>
                    </table> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">ˮӡ��ͼƬ�ϵ�Y���꣺</td>
                  <td><input name="ImageHeight" type="text" id="ImageHeight" size="4" maxlength="3" value="<%= RS("ImageHeight") %>"> 
                    &nbsp;[������]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">ˮӡͼ�꣺</td>
                  <td><input name="Img" type="text" id="Img" size="18" value="<%= RS("ImageLogo") %>" readonly="0">
<input name="Submit_Open" type="button" id="Submit_Open" value="�ϴ�" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;"> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">ˮӡ͸���ȣ�</td>
                  <td> <input name="ImageAlpha" type="text" id="ImageAlpha" size="4" maxlength="3" value="<%= RS("ImageAlpha") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">ʹ��ˮӡ��</td>
                  <td>�� 
                    <input name="LogoLock" type="radio" value="1" <% IF RS("LogoLock")=1 Then %>checked<% End IF %>>
                    �� 
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
              <legend><strong>�ı�</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="22" align="center">������ͼƬ�ϵ�X���꣺</td>
                  <td><input name="fontWidth" type="text" id="fontWidth" size="4" maxlength="3" value="<%= RS("fontWidth") %>"> 
                    &nbsp;[������]</td>
                </tr>
                <tr> 
                  <td height="22" align="center">������ͼƬ�ϵ�Y���꣺</td>
                  <td><input name="fontHeight" type="text" id="fontHeight" size="4" maxlength="3" value="<%= RS("fontHeight") %>"> 
                    &nbsp;[������]</td>
                </tr>
                <tr> 
                  <td width="150" height="22" align="center">������ɫ��</td>
                  <td><input name="FontColor" type="text" id=d_bgcolor size="7" maxlength="7" value="<%= RS("FontColor") %>"> 
                    <img src="../sysImages/Rect.gIF" alt="������ɫ" name="s_bgcolor" width="18" height="17" border="0" align="absmiddle" id=s_bgcolor style="cursor:hand" onClick="SelectColor('bgcolor')"> 
                    <input type="button" name="Clear" value="���"onclick="document.form1.FontColor.value=''" title="���������ɫ"> 
                  </td>
                </tr>
                <tr> 
                  <td height="22" align="center">�������壺</td>
                  <td> <select name="FontFamily" id="FontFamily">
                      <option value="����" <% IF RS("FontFamily")="����" Then %>selected<% End IF %>>����</option>
                      <option value="����" <% IF RS("FontFamily")="����" Then %>selected<% End IF %>>����</option>
                      <option value="��Բ" <% IF RS("FontFamily")="��Բ" Then %>selected<% End IF %>>��Բ</option>
                      <option value="����_GB2312" <% IF RS("FontFamily")="����_GB2312" Then %>selected<% End IF %>>����</option>
                      <option value="�����п�" <% IF RS("FontFamily")="�����п�" Then %>selected<% End IF %>>�����п�</option>
                      <option value="���Ĳ���" <% IF RS("FontFamily")="���Ĳ���" Then %>selected<% End IF %>>���Ĳ���</option>
                    </select> </td>
                </tr>
                <tr> 
                  <td height="22" align="center">�����ֺţ�</td>
                  <td>
<select name="fontsize" id="fontsize">
<% for i=1 to 30 %>
           <option value="<%= i %>" <% If Cint(RS("fontsize"))=i Then %>selected<% End IF %>><%= i %></option>
					  <% next %>
</select>
				  </td>
                </tr>
                <tr> 
                  <td height="22" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                  <td><input name="fontText" type="text" id="fontText" size="30" value="<%= RS("fontText") %>"></td>
                </tr>
                <tr>
                  <td height="22" align="center">ʹ���ı���</td>
                  <td>�� 
                    <input name="fontLock" type="radio" value="1" <% IF RS("fontLock")=1 Then %>checked<% End IF %>>
                    �� 
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
<input type="submit" name="Submit_OK" value="�޸�(E)" accesskey="E">
              &nbsp; <input type="Reset" name="reset" value="ˢ��(R)" accesskey="R" onClick="javascript:location.reload();"> 
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
