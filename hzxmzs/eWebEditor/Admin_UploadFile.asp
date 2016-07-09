<!--#include file = "Include/Startup.asp"-->
<!--#include file = "admin_private.asp"-->
<!--#include file="../Admin/Include/Function.asp" -->
<%

Dim sStyleID, sUploadDir, sCurrDir, sDir

sPosition = sPosition & "上传文件管理"

Call Header()
Call Content()
Call Footer()


Sub Content()
	If IsObjInstalled("Scripting.FileSystemObject") = False Then
		Response.Write "此功能要求服务器支持文件系统对象（FSO），而你当前的服务器不支持！"
		Exit Sub
	End If

	' 初始化传入参数
	Call InitParam()

	Select Case sAction
	Case "DELALL"			' 删除所有文件
		Call DoDelAll()
	Case "DEL"				' 删除指定文件
		Call DoDel()
	Case "DELFOLDER"		' 删除文件夹
		Call DoDelFolder()
	End Select

	' 显示文件列表
	Call ShowList()
End Sub

' UploadFile目录下的所有文件列表
Sub ShowList()

	Response.Write "<BR><table width='100%' border=0 align=center cellpadding=0 cellspacing=1 bgcolor='#D7DEF8'><tr>" & _
		"<form action='?' method=post name=queryform><td height='22' align='right' bgcolor='#FFFFFF'>" & _
		"选择样式目录：<select name='id' size=1 onchange=""location.href='?id='+this.options[this.selectedIndex].value"">" & InitSelect(sStyleID, "select ('样式：'+S_Name+'---目录：'+S_UploadDir),S_ID from eWebEditor_Style order by S_ID asc", "选择...") & "</select>" & _
		"</td></form></tr></table><BR>"
	
	If sCurrDir = "" Then Exit Sub
	
	Response.Write "<table width=100% height=22 border=0 cellpadding=0 cellspacing=1 bgcolor=#D7DEF8>" & _
		"<form action='?id=" & sStyleID & "&action=del' method=post name=myform>" & _
		"<tr>" & _
			"<td width=50 height=25 align=center bgcolor=#FFFFFF>类型</td>" & _
			"<td bgcolor=#FFFFFF align=center> 文件地址 </td>" & _
			"<td width=100 align=center bgcolor=#FFFFFF> 大小 </td>" & _
			"<td width=100 align=center bgcolor=#FFFFFF> 最后访问 </td>" & _
			"<td  align=center bgcolor=#FFFFFF> 上传日期 </td>" & _
			"<td width=50 align=center bgcolor=#FFFFFF> 删除 </td>" & _
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
		Response.Write "</table>无效的目录！"
		Exit Sub
	End If

	
	If sDir <> "" Then
		Response.Write "<tr align=center style='cursor:hand' OnMouseOver=this.bgColor='#F0F0F0' OnMouseOut=this.bgColor='#FFFFFF'>" & _
			"<td><img border=0 src='sysimage/file/folderback.gif'></td>" & _
			"<td align=left colspan=5><a href=""?id=" & sStyleID & "&dir="
		If InstrRev(sDir, "/") > 1 Then
			Response.Write Left(sDir, InstrRev(sDir, "/") - 1)
		End If
		Response.Write """>返回上一级目录</a></td></tr>"
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
			"<td><a href='?id=" & sStyleID & "&dir=" & sDir & "&action=delfolder&foldername=" & oSubFolder.Name & "'>删除</a></td></tr>"
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
		Response.Write "<tr><td height=30 bgcolor=#FFFFFF colspan='6' align=center><font color='#FF0000'>指定目录下现在还没有文件！</font></td></tr>"
	End If
	Response.Write "</table><BR>"

	If nFileNum > 0 Then
		' 分页
		Response.Write "<table width=100% height=30 border=0 cellpadding=0 cellspacing=1 bgcolor=#D7DEF8 ><tr><td align=center bgcolor=#FFFFFF>"
		If nCurrPage > 1 Then
			Response.Write "<a href='?page=1&ID="&Request.QueryString("ID")&"'>首页</a>&nbsp;&nbsp;<a href='?page="& nCurrPage - 1 & "&ID="&Request.QueryString("ID")&"'>上一页</a>&nbsp;&nbsp;"
		Else
			Response.Write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
		End If
		If nCurrPage < i / nPageSize Then
			Response.Write "<a href='?page="&nCurrPage + 1&"&ID="&Request.QueryString("ID")&"'>下一页</a>&nbsp;&nbsp;<a href='?page=" & nPageNum & "&ID="&Request.QueryString("ID")&"'>尾页</a>"
		Else
			Response.Write "下一页&nbsp;&nbsp;尾页"
		End If
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;共<b>" & nFileNum & "</b>个&nbsp;&nbsp;页次:<b>" & nCurrPage & "/" & nPageNum & "</b>&nbsp;&nbsp;<b>" & nPageSize & "</b>个文件/页"
		Response.Write "</td><td width=200 align=center bgcolor=#FFFFFF>"
		Response.Write "<input type=submit name=b value='删除'> &nbsp;&nbsp; &nbsp;&nbsp;<input type=button name=b1 value='清空' onclick=""javascript:if (confirm('你确定要清空所有文件吗？')) {location.href='admin_uploadfile.asp?id=" & sStyleID & "&action=delall';}"">"
	
		Response.Write "</td></tr></table></form>"
	End If

End Sub

' 删除指定的文件
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

' 删除所有的文件
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

' 删除文件夹
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

' 检测服务器是否支持某一对象
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


' 按文件名取图
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
' 初始化下拉框
'	v_InitValue	: 初始值
'	s_Sql		: 从数据库中取值时,select name,value from table
'	s_AllName	: 空值的名称,如:"全部","所有","默认"
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
' 初始化传入参数
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

	' 样式下的目录
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
' 检测目录的有效性
' ===============================================
Function CheckValidDir(s_Dir)
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	CheckValidDir = oFSO.FolderExists(s_Dir)
	Set oFSO = Nothing	
End Function
%>