<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="UpFile_Submit.asp" -->
<html>
<head>
<title>上传文件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">

</head>
<body bgcolor="#D7DEF8" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false;" onselectstart="return false;">
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set Rs = Server.CreateObject( "ADODB.Recordset" )
Sql="select Top 1 * from UpFile"
Rs.open sql,Conn,1,1
''''''''''''''''''''''''''''''''''''''''''''''''
dim upload,file,formName,formPath
dim i,l,fileType,newfilename,filenamelist
''''''''''''''''''''''''''''''''''''''''''''''''
formPath = RS("JpegPath")
''''''''''''''''''''''''''''''''''''''''''''''''
newfilename=makefilename()
set upload=new upload_5xsoft
formPath=server.mappath(formPath)&"/"
for each formName in upload.objFile
	set file=upload.file(formName)
		fileType=file.FileName
		i=instr(fileType,".")
		l=len(fileType)
		IF i>0 Then
		  fileType=right(fileType,l-i+1)
		End if
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		UpFileType=lcase(fileType)
		arrUpFileType=split(RS("JpegType"),"|")
		For i=0 to ubound(arrUpFileType)
			IF UpFileType = trim(arrUpFileType(i)) Then
				EnableUpload = true
				Exit for
			End if
		Next
	IF EnableUpload=false then
		FoundErr=true
		Response.Write("<script language='javascript'>alert('系统提示:\n\n文件类型不允许!\n[只允许"&RS("JpegType")&"]');history.go(-1)</script>")
		Response.End
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	ElseIF File.FileSize > RS("Jpegsize") then
		Response.Write("<script language='javascript'>alert('系统提示:\n\n上传文件过大!\n[最大"&File.FileSize&"KB]');history.go(-1)</script>")
		Response.End
	Else
		newfilenamea=newfilename & "a" & fileType
		filenamelista=formPath & newfilenamea
		newfilename=newfilename & fileType
		filenamelist=formPath & newfilename
		file.SaveAs filenamelist
		set file=nothing
	End IF
next
set upload=nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'缩略图
If RS("JpegLock") = 1 Then
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.Open filenamelist
	Jpeg.Width = RS("JpegWidth")
	Jpeg.Height = RS("JpegHeight")
	Jpeg.Save filenamelista
	Set Jpeg = Nothing
End IF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'添加图片水印
If RS("LogoLock") = 1 Then
	Set Image = Server.CreateObject("Persits.Jpeg") 
	Image.Open filenamelist
	Set Logo = Server.CreateObject("Persits.Jpeg") 
	LogoPath = Server.MapPath(RS("JpegPath")&"/"&RS("ImageLogo"))
	Logo.Open LogoPath
	Width = Image.Width - RS("ImageWidth")
	Height = Image.Height - RS("ImageHeight")
	Image.DrawImage Width, Height, Logo,RS("ImageAlpha")
	Image.Save filenamelist
	Set Image = Nothing
	Set Logo = Nothing
End IF	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 添加文字
If RS("fontLock") = 1 Then
	dim Font
	Set Font = Server.CreateObject("Persits.Jpeg")
	Font.Open filenamelist
	Font.Canvas.Font.Color = "&h"&RS("fontColor")
	Font.Canvas.Font.Family = ""& RS("fontFamily") &""
    Font.canvas.font.size= ""& RS("fontsize") &""
	Width = font.Width - RS("fontWidth")
	Height = font.Height - RS("fontHeight")
	Font.Canvas.Print Width,Height,""& RS("fontText") &""
	Font.Save filenamelist
	Set Font = Nothing
End IF
 %> 


<table width="300" border="0" align="center" cellpadding="10" cellspacing="10">
  <tr>
            <td>		
  <fieldset style="padding: 5" class="body">
              
      <legend><strong>上传成功</strong></legend>
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td height="50" align="center"> 
<script language="JavaScript">
//两个表单
function Img_logo()
{
	window.opener.document.theForm.Img.value='<%= newfilename %>'
	window.opener.document.theForm.logo.src='<%= RS("JpegPath") %>/<%= newfilename %>'
	opener=null;
	close();
}
//一个表单
function Upfile()
{
	window.opener.document.theForm.Upfile.value='<%= newfilename %>'
	opener=null;
	close();
}

</script>
<%
Dim UpfileType
UpfileType = Request.QueryString("UpfileType")
Select Case UpfileType
	Case "Img_logo"
Response.Write("<input type='button' value='关闭窗口' onClick='javascript:Img_logo();'  name=button>")
	Case "Upfile"
Response.Write("<input type='button' value='关闭窗口' onClick='javascript:Upfile();'  name=button>")
End Select
%>
          </td>
        </tr>
      </table>
      </fieldset>
        </td>
  </tr>
</table>
</body>
</html>

