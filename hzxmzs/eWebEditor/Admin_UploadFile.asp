<!--#include file = "Include/Startup.asp"-->
<!--#include file = "admin_private.asp"-->
<!--#include file="../Admin/Include/Function.asp" -->
<%

Dim sStyleID, sUploadDir, sCurrDir, sDir

sPosition = sPosition & "�ϴ��ļ�����"

Call Header()
Call Content()
Call Footer()


Sub Content()
	If IsObjInstalled("Scripting.FileSystemObject") = False Then
		Response.Write "�˹���Ҫ�������֧���ļ�ϵͳ����FSO�������㵱ǰ�ķ�������֧�֣�"
		Exit Sub
	End If

	' ��ʼ���������
	Call InitParam()

	Select Case sAction
	Case "DELALL"			' ɾ�������ļ�
		Call DoDelAll()
	Case "DEL"				' ɾ��ָ���ļ�
		Call DoDel()
	Case "DELFOLDER"		' ɾ���ļ���
		Call DoDelFolder()
	End Select

	' ��ʾ�ļ��б�
	Call ShowList()
End Sub

' UploadFileĿ¼�µ������ļ��б�
Sub ShowList()

	Response.Write "<BR><table width='100%' border=0 align=center cellpadding=0 cellspacing=1 bgcolor='#D7DEF8'><tr>" & _
		"<form action='?' method=post name=queryform><td height='22' align='right' bgcolor='#FFFFFF'>" & _
		"ѡ����ʽĿ¼��<select name='id' size=1 onchange=""location.href='?id='+this.options[this.selectedIndex].value"">" & InitSelect(sStyleID, "select ('��ʽ��'+S_Name+'---Ŀ¼��'+S_UploadDir),S_ID from eWebEditor_Style order by S_ID asc", "ѡ��...") & "</select>" & _
		"</td></form></tr></table><BR>"
	
	If sCurrDir = "" Then Exit Sub
	
	Response.Write "<table width=100% height=22 border=0 cellpadding=0 cellspacing=1 bgcolor=#D7DEF8>" & _
		"<form action='?id=" & sStyleID & "&action=del' method=post name=myform>" & _
		"<tr>" & _
			"<td width=50 height=25 align=center bgcolor=#FFFFFF>����</td>" & _
			"<td bgcolor=#FFFFFF align=center> �ļ���ַ </td>" & _
			"<td width=100 align=center bgcolor=#FFFFFF> ��С </td>" & _
			"<td width=100 align=center bgcolor=#FFFFFF> ������ </td>" & _
			"<td  align=center bgcolor=#FFFFFF> �ϴ����� </td>" & _
			"<td width=50 align=center bgcolor=#FFFFFF> ɾ�� </td>" & _
		"</tr>"

	Dim sCurrPage, nCurrPage, nFileNum, nPageNum, nPageSize
	sCurrPage = Trim(Request("page"))
	nPageSize = 14
	If sCurrpage = "" Or Not IsNumeric(sCurrPage) Then
		nCurrPage = 1
	Else
		nCurrPage = CLng(sCurrPage)
	End If

	Dim oFSO, oUploadFolder, oUploadFiles, oUploadFile, sFileName

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	Set oUploadFolder = oFSO.GetFolder(Server.MapPath(sCurrDir))
	If Err.Number>0 Then
		Response.Write "</table>��Ч��Ŀ¼��"
		Exit Sub
	End If

	
	If sDir <> "" Then
		Response.Write "<tr align=center style='cursor:hand' OnMouseOver=this.bgColor='#F0F0F0' OnMouseOut=this.bgColor='#FFFFFF'>" & _
			"<td><img border=0 src='sysimage/file/folderback.gif'></td>" & _
			"<td align=left colspan=5><a href=""?id=" & sStyleID & "&dir="
		If InstrRev(sDir, "/") > 1 Then
			Response.Write Left(sDir, InstrRev(sDir, "/") - 1)
		End If
		Response.Write """>������һ��Ŀ¼</a></td></tr>"
	End If

	Dim oSubFolder
	For Each oSubFolder In oUploadFolder.SubFolders
		Response.Write "<tr align=center>" & _
			"<td><img border=0 src='sysimage/file/folder.gif'></td>" & _
			"<td align=left colspan=4><a href=""?id=" & sStyleID & "&dir="
		If sDir <> "" Then
			Response.Write sDir & "/"
		End If
		Response.Write oSubFolder.Name & """>" & oSubFolder.Name & "</a></td>" & _
			"<td><a href='?id=" & sStyleID & "&dir=" & sDir & "&action=delfolder&foldername=" & oSubFolder.Name & "'>ɾ��</a></td></tr>"
	Next

	
	Set oUploadFiles = oUploadFolder.Files

	nFileNum = oUploadFiles.Count
	nPageNum = Int(nFileNum / nPageSize)
	If nFileNum Mod nPageSize > 0 Then
		nPageNum = nPageNum+1
	End If
	If nCurrPage > nPageNum Then
		nCurrPage = 1
	end If

	Dim i
	i = 0
	For Each oUploadFile In oUploadFiles
		i = i + 1
		If i > (nCurrPage - 1) * nPageSize And i <= nCurrPage * nPageSize Then
			sFileName = oUploadFile.Name
			
			Response.Write "<tr align=center>" & _
				"<td height=30 bgcolor=#FFFFFF>" & FileName2Pic(sFileName) & "</td>" & _
				"<td bgcolor=#FFFFFF>"
Dim Image,Ico
Image = Right(sFileName,3)
Select case Image
	case "JPG":Ico=1
	case "jpg":Ico=1
	case "GIF":Ico=1
	case "gif":Ico=1
	case "PNG":Ico=1
	case "png":Ico=1
	case "BMP":Ico=1
	case "bmp":Ico=1
End Select
If Ico = 1 Then
Response.Write "<a href=../Admin/Folder/Folder_Image_Main.asp?UpFilesPath=../"&sUploadDir&"&FileName="& sFileName &">" & sFileName & "</a>"

ElseIF Right(sFileName,3) = "swf" Then
Response.Write "<a href=../Admin/Folder/Folder_Flash.asp?UpFilesPath=../"&sUploadDir&"&FileName="& sFileName &">" & sFileName & "</a>"
Else				
Response.Write "<a href="&sUploadDir&""&sFileName &">" & sFileName & "</a>"
End IF
				Response.Write "</td>" & _
				"<td bgcolor=#FFFFFF>" & oUploadFile.size & " BK </td>" & _
				"<td bgcolor=#FFFFFF>" & oUploadFile.datelastaccessed & "</td>" & _
				"<td bgcolor=#FFFFFF>" & oUploadFile.datecreated & "</td>" & _
				"<td bgcolor=#FFFFFF><input type=checkbox name=delfilename value=""" & sFileName & """></td></tr>"
		Elseif i > nCurrPage * nPageSize Then
			Exit For
		End If
	Ico=0
	Next
		Set oUploadFolder = Nothing
	Set oUploadFiles = Nothing

	If nFileNum <= 0 Then
		Response.Write "<tr><td height=30 bgcolor=#FFFFFF colspan='6' align=center><font color='#FF0000'>ָ��Ŀ¼�����ڻ�û���ļ���</font></td></tr>"
	End If
	Response.Write "</table><BR>"

	If nFileNum > 0 Then
		' ��ҳ
		Response.Write "<table width=100% height=30 border=0 cellpadding=0 cellspacing=1 bgcolor=#D7DEF8 ><tr><td align=center bgcolor=#FFFFFF>"
		If nCurrPage > 1 Then
			Response.Write "<a href='?page=1&ID="&Request.QueryString("ID")&"'>��ҳ</a>&nbsp;&nbsp;<a href='?page="& nCurrPage - 1 & "&ID="&Request.QueryString("ID")&"'>��һҳ</a>&nbsp;&nbsp;"
		Else
			Response.Write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
		End If
		If nCurrPage < i / nPageSize Then
			Response.Write "<a href='?page="&nCurrPage + 1&"&ID="&Request.QueryString("ID")&"'>��һҳ</a>&nbsp;&nbsp;<a href='?page=" & nPageNum & "&ID="&Request.QueryString("ID")&"'>βҳ</a>"
		Else
			Response.Write "��һҳ&nbsp;&nbsp;βҳ"
		End If
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;��<b>" & nFileNum & "</b>��&nbsp;&nbsp;ҳ��:<b>" & nCurrPage & "/" & nPageNum & "</b>&nbsp;&nbsp;<b>" & nPageSize & "</b>���ļ�/ҳ"
		Response.Write "</td><td width=200 align=center bgcolor=#FFFFFF>"
		Response.Write "<input type=submit name=b value='ɾ��'> &nbsp;&nbsp; &nbsp;&nbsp;<input type=button name=b1 value='���' onclick=""javascript:if (confirm('��ȷ��Ҫ��������ļ���')) {location.href='admin_uploadfile.asp?id=" & sStyleID & "&action=delall';}"">"
	
		Response.Write "</td></tr></table></form>"
	End If

End Sub

' ɾ��ָ�����ļ�
Sub DoDel()
	On Error Resume Next
	Dim sFileName, oFSO, sMapFileName
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	For Each sFileName In Request.Form("delfilename")
		sMapFileName = Server.MapPath(sCurrDir & sFileName)
		If oFSO.FileExists(sMapFileName) Then
			oFSO.DeleteFile(sMapFileName)
		End If
	Next
	Set oFSO = Nothing
End Sub

' ɾ�����е��ļ�
Sub DoDelAll()
	On Error Resume Next
	Dim sFileName, oFSO, sMapFileName, oFolder, oFiles, oFile
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFSO.GetFolder(Server.MapPath(sCurrDir))
	Set oFiles = oFolder.Files
	For Each oFile In oFiles
		sFileName = oFile.Name
		sMapFileName = Server.MapPath(sCurrDir & sFileName)
		If oFSO.FileExists(sMapFileName) Then
			oFSO.DeleteFile(sMapFileName)
		End If
	Next
	Set oFile = Nothing
	Set oFolder = Nothing
	Set oFSO = Nothing
End Sub

' ɾ���ļ���
Sub DoDelFolder()
	On Error Resume Next
	Dim sFolderName, oFSO, sMapFolderName
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sFolderName = Trim(Request("foldername"))
	sMapFolderName = Server.Mappath(sCurrDir & sFolderName)
	If oFSO.FolderExists(sMapFolderName) = True Then
		oFSO.DeleteFolder(sMapFolderName)
	End If
	Set oFSO = Nothing
End Sub

' ���������Ƿ�֧��ĳһ����
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function


' ���ļ���ȡͼ
Function FileName2Pic(sFileName)
	Dim sExt, sPicName
	sExt = UCase(Mid(sFileName, InstrRev(sFileName, ".")+1))
	Select Case sExt
	Case "TXT"
		sPicName = "txt.gif"
	Case "CHM", "HLP"
		sPicName = "hlp.gif"
	Case "DOC"
		sPicName = "doc.gif"
	Case "PDF"
		sPicName = "pdf.gif"
	Case "MDB"
		sPicName = "mdb.gif"
	Case "GIF"
		sPicName = "gif.gif"
	Case "JPG"
		sPicName = "jpg.gif"
	Case "BMP"
		sPicName = "bmp.gif"
	Case "PNG"
		sPicName = "pic.gif"
	Case "ASP", "JSP", "JS", "PHP", "PHP3", "ASPX"
		sPicName = "code.gif"
	Case "HTM", "HTML", "SHTML"
		sPicName = "htm.gif"
	Case "ZIP"
		sPicName = "zip.gif"
	Case "RAR"
		sPicName = "rar.gif"
	Case "EXE"
		sPicName = "exe.gif"
	Case "AVI"
		sPicName = "avi.gif"
	Case "MPG", "MPEG", "ASF"
		sPicName = "mp.gif"
	Case "RA", "RM"
		sPicName = "rm.gif"
	Case "MP3"
		sPicName = "mp3.gif"
	Case "MID", "MIDI"
		sPicName = "mid.gif"
	Case "WAV"
		sPicName = "audio.gif"
	Case "XLS"
		sPicName = "xls.gif"
	Case "PPT", "PPS"
		sPicName = "ppt.gif"
	Case "SWF"
		sPicName = "swf.gif"
	Case Else
		sPicName = "unknow.gif"
	End Select
	FileName2Pic = "<img border=0 src='sysimage/file/" & sPicName & "'>"
End Function

' ===============================================
' ��ʼ��������
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
' ===============================================
Function InitSelect(v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = ""
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	oRs.Open s_Sql, oConn, 0, 1
	Do While Not oRs.Eof
		InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
		If CStr(oRs(1)) = CStr(v_InitValue) Then
			InitSelect = InitSelect & " selected"
		End If
		InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
		oRs.MoveNext
	Loop
	oRs.Close
End Function

' ===============================================
' ��ʼ���������
' ===============================================
Function InitParam()
	sStyleID = Trim(Request("id"))
	sUploadDir = ""
	If IsNumeric(sStyleID) = True Then
		sSql = "select S_UploadDir from eWebEditor_Style where S_ID=" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			sUploadDir = oRs(0)
		End If
		oRs.Close
	End If
	If sUploadDir = "" Then
		sStyleID = ""
	Else
		sUploadDir = Replace(sUploadDir, "\", "/")
		If Right(sUploadDir, 1) <> "/" Then
			sUploadDir = sUploadDir & "/"
		End If
	End If
	sCurrDir = sUploadDir

	' ��ʽ�µ�Ŀ¼
	sDir = Trim(Request("dir"))
	If sDir <> "" Then
		If CheckValidDir(Server.Mappath(sUploadDir & sDir)) = True Then
			sCurrDir = sUploadDir & sDir & "/"
		Else
			sDir = ""
		End If
	End If
End Function

' ===============================================
' ���Ŀ¼����Ч��
' ===============================================
Function CheckValidDir(s_Dir)
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	CheckValidDir = oFSO.FolderExists(s_Dir)
	Set oFSO = Nothing	
End Function
%>