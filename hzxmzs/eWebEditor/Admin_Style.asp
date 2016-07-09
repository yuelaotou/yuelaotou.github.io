<!--#include file = "Include/Startup.asp"-->
<!--#include file = "admin_private.asp"-->

<%

Dim sStyleID, sStyleName, sStyleDir, sStyleCSS, sStyleUploadDir, sStyleWidth, sStyleHeight, sStyleMemo, nStyleIsSys, sStyleStateFlag, sStyleDetectFromWord, sStyleInitMode, sStyleBaseUrl, sStyleUploadObject, sStyleAutoDir, sStyleBaseHref, sStyleContentPath, sStyleAutoRemote, sStyleShowBorder
Dim sStyleFileExt, sStyleFlashExt, sStyleImageExt, sStyleMediaExt, sStyleRemoteExt, sStyleFileSize, sStyleFlashSize, sStyleImageSize, sStyleMediaSize, sStyleRemoteSize
Dim sToolBarID, sToolBarName, sToolBarOrder, sToolBarButton

sPosition = sPosition & "样式管理"

If sAction = "STYLEPREVIEW" Then
	' 样式预览
	Call InitStyle()
	Call ShowStylePreview()
	Response.End
End If


Call Header()
Call Content()
Call Footer()


Sub Content()
	Select Case sAction
	Case "COPY"
		' 拷贝一标准样式
		Call InitStyle()
		Call DoCopy()
		Call ShowStyleList()
	Case "STYLEADD"
		' 新增样式表单
		Call ShowStyleForm("ADD")
	Case "STYLESET"
		' 样式设置，修改或查看
		Call InitStyle()
		Call ShowStyleForm("SET")
	Case "STYLEADDSAVE"
		' 样式新增保存
		Call CheckStyleForm()
		Call DoStyleAddSave()
	Case "STYLESETSAVE"
		' 样式设置修改保存
		Call CheckStyleForm()
		Call DoStyleSetSave()
		Call RemoveApplication()
	Case "STYLEDEL"
		' 样式删除
		Call InitStyle()
		Call DoStyleDel()
		Call ShowStyleList()
		Call RemoveApplication()
	Case "CODE"
		' 显示引用代码
		Call InitStyle()
		Call ShowStyleCode()
	Case "TOOLBAR"
		' 显示工具栏表单
		Call InitStyle()
		Call ShowToolBarList()
	Case "TOOLBARADD"
		' 新增工具栏
		Call InitStyle()
		Call DoToolBarAdd()
		Call ShowToolBarList()
		Call RemoveApplication()
	Case "TOOLBARMODI"
		' 修改工具栏
		Call InitStyle()
		Call DoToolBarModi()
		Call ShowToolBarList()
		Call RemoveApplication()
	Case "TOOLBARDEL"
		' 删除工具栏
		Call InitStyle()
		Call DoToolBarDel()
		Call ShowToolBarList()
		Call RemoveApplication()
	Case "BUTTONSET"
		' 按钮设置
		Call InitStyle()
		Call InitToolBar()
		Call ShowButtonList()
	Case "BUTTONSAVE"
		' 按钮设置保存
		Call InitStyle()
		Call InitToolBar()
		Call DoButtonSave()
		Call RemoveApplication()
	Case Else
		' 当前所有样式列表
		Call ShowStyleList()
	End Select
End Sub


' 当前所有样式列表
Sub ShowStyleList()

	Response.Write "<BR><table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#D7DEF8'>"& _
  "<tr>"& _
    "<td width='80' height='22' align='center' bgcolor='#FFFFFF'>样式名称</td>"& _
    "<td width='60' align='center' bgcolor='#FFFFFF'>最佳宽度</td>"& _
    "<td width='60' align='center' bgcolor='#FFFFFF'>最佳高度</td>"& _
   "<td align='center' bgcolor='#FFFFFF'>说明</td>"& _
    "<td width='180' align='center' bgcolor='#FFFFFF'>管理</td>"& _
  "</tr>"

	Dim sManage
	sSql = "select * from ewebeditor_style"
	oRs.Open sSql, oConn, 0, 1
	Do While Not oRs.Eof
		sManage = "<a href='?action=stylepreview&id=" & oRs("S_ID") & "'>预览</a>|<a href='?action=code&id=" & oRs("S_ID") & "'>代码</a>|<a href='?action=styleset&id=" & oRs("S_ID") & "'>设置</a>|<a href='?action=toolbar&id=" & oRs("S_ID") & "'>工具栏</a>"
		If oRs("S_IsSys") = 1 Then
			sManage = sManage & "|<a href='?action=copy&id=" & oRs("S_ID") & "'>拷贝</a>"
		Else
			sManage = sManage & "|<a href='?action=styledel&id=" & oRs("S_ID") & "'>删除</a>"
		End If
		Response.Write "<tr style='cursor:hand'>" & _
			"<td height='22' align='center' bgcolor='#FFFFFF'>" & outHTML(oRs("S_Name")) & "</td>" & _
		"<td align='center' bgcolor='#FFFFFF'>" & oRs("S_Width") & "</td>" & _
		"<td align='center' bgcolor='#FFFFFF'>" & oRs("S_Height") & "</td>" & _
		"<td bgcolor='#FFFFFF'>&nbsp;" & outHTML(oRs("S_Memo")) & "</td>" & _
		"<td align='center' bgcolor='#FFFFFF'>" & sManage & "</td>" & _
		"</tr>"
		oRs.MoveNext
	Loop
	oRs.Close
	
	Response.Write "</table><BR>"


End Sub

' 拷贝一标准样式
Sub DoCopy()
	' 只有系统样式才有拷贝功能
	If nStyleIsSys <> 1 Then
		Exit Sub
	End If
	' 自动取有效的样式名，从1......到...
	Dim i, b, sNewID, sNewName
	b = False
	i = 0
	Do While b = False
		i = i + 1
		sNewName = sStyleName & i
		sSql = "select s_id from ewebeditor_style where s_name='" & sNewName & "'"
		oRs.Open sSql, oConn, 0, 1
		If oRs.Eof And oRs.Bof Then
			b = True
		End If
		oRs.Close
	Loop
	' 插入一与标准样式相同的记录，名不同
	sSql = "insert into ewebeditor_style(s_name,s_dir,s_css,s_uploaddir,s_width,s_height,s_memo,s_issys,s_fileext,s_flashext,s_imageext,s_mediaext, s_remoteext,s_filesize,s_flashsize,s_imagesize,s_mediasize,s_remotesize,s_stateflag,s_detectfromword,s_initmode,s_baseurl,s_uploadobject,s_basehref,s_contentpath) select '" & sNewName & "',s_dir,s_css,s_uploaddir,s_width,s_height,s_memo,0,s_fileext,s_flashext,s_imageext,s_mediaext,s_remoteext,s_filesize,s_flashsize,s_imagesize,s_mediasize,s_remotesize,s_stateflag,s_detectfromword,s_initmode,s_baseurl,s_uploadobject,s_basehref,s_contentpath from ewebeditor_style where s_id=" & sStyleID
	oConn.Execute sSql
	' 取新样式的ID
	sSql = "select s_id from ewebeditor_style where s_name='" & sNewName & "'"
	oRs.Open sSql, oConn, 0, 1
	sNewID = oRs(0)
	oRs.Close
	' 拷贝工具栏
	sSql = "insert into ewebeditor_toolbar(s_id,t_name,t_order,t_button) select " & sNewID & ",t_name,t_order,t_button from ewebeditor_toolbar where s_id=" & sStyleID
	oConn.Execute sSql
End Sub

' 样式表单
Sub ShowStyleForm(sFlag)
	Dim s_Title, s_Button, s_Action
	Dim s_FormStateFlag, s_FormDetectFromWord, s_FormInitMode, s_FormBaseUrl, s_FormUploadObject, s_FormAutoDir, s_FormAutoRemote, s_FormShowBorder
	
	If sFlag = "ADD" Then
		sStyleID = ""
		sStyleName = ""
		sStyleDir = "standard"
		sStyleCSS = "office"
		sStyleUploadDir = "UploadFile/"
		sStyleBaseHref = "http://Localhost/eWebEditor/"
		sStyleContentPath = "UploadFile/"
		sStyleWidth = "600"
		sStyleHeight = "400"
		sStyleMemo = ""
		nStyleIsSys = 0
		s_Title = "新增样式"
		s_Action = "StyleAddSave"
		sStyleFileExt = "rar|zip|exe|doc|xls|chm|hlp"
		sStyleFlashExt = "swf"
		sStyleImageExt = "gif|jpg|jpeg|bmp"
		sStyleMediaExt = "rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov"
		sStyleRemoteExt = "gif|jpg|bmp"
		sStyleFileSize = "500"
		sStyleFlashSize = "100"
		sStyleImageSize = "100"
		sStyleMediaSize = "100"
		sStyleRemoteSize = "100"
		sStyleStateFlag = "1"
		sStyleAutoRemote = "1"
		sStyleShowBorder = "0"
		sStyleUploadObject = "0"
		sStyleAutoDir = "0"
		sStyleDetectFromWord = "true"
		sStyleInitMode = "EDIT"
		sStyleBaseUrl = "0"
	Else
		sStyleName = inHTML(sStyleName)
		sStyleDir = inHTML(sStyleDir)
		sStyleCSS = inHTML(sStyleCSS)
		sStyleUploadDir = inHTML(sStyleUploadDir)
		sStyleBaseHref = inHTML(sStyleBaseHref)
		sStyleContentPath = inHTML(sStyleContentPath)
		sStyleMemo = inHTML(sStyleMemo)
		s_Title = "设置样式"
		s_Action = "StyleSetSave"
	End If

	s_FormStateFlag = InitSelect("d_stateflag", Split("显示|不显示", "|"), Split("1|0", "|"), sStyleStateFlag, "", "")
	s_FormAutoRemote = InitSelect("d_autoremote", Split("自动上传|不自动上传", "|"), Split("1|0", "|"), sStyleAutoRemote, "", "")
	s_FormShowBorder = InitSelect("d_showborder", Split("默认显示|默认不显示", "|"), Split("1|0", "|"), sStyleShowBorder, "", "")
	s_FormUploadObject = InitSelect("d_uploadobject", Split("无惧无组件上传类|ASPUpload上传组件|SA-FileUp上传组件|LyfUpload上传组件", "|"), Split("0|1|2|3", "|"), sStyleUploadObject, "", "")
	s_FormAutoDir = InitSelect("d_autodir", Split("不使用|年目录|年月目录|年月日目录", "|"), Split("0|1|2|3", "|"), sStyleAutoDir, "", "")
	s_FormDetectFromWord = InitSelect("d_detectfromword", Split("自动检测有提示|不自动检测", "|"), Split("true|false", "|"), sStyleDetectFromWord, "", "")
	s_FormInitMode = InitSelect("d_initmode", Split("代码模式|编辑模式|文本模式|预览模式", "|"), Split("CODE|EDIT|TEXT|VIEW", "|"), sStyleInitMode, "", "")
	s_FormBaseUrl = InitSelect("d_baseurl", Split("相对路径|绝对根路径|绝对全路径", "|"), Split("0|1|2", "|"), sStyleBaseUrl, "", "")
	
	s_Button = "<tr><td align=center colspan=4><input type='submit' name='Submit' value='提交'>&nbsp;<input type='reset' name='Submit2' value='重置'></td></tr>"
	

	Response.Write "<BR><table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#CCCCCC'>"& _
  "<tr>"& _
    "<td height='22' bgcolor='#FFFFFF'>"& _
	"<table border=0 cellpadding=0 cellspacing=0 align=center>" & _
			"<form action='?action=" & s_Action & "&id=" & sStyleID & "' method=post name=myform>" & _
			"" & _
			"<tr><td width='15%' align=right>样式名称：</td><td width='35%'><input type=text  size=20 name=d_name title='引用此样式的名字，不要加特殊符号，最大50个字符长度' value=""" & sStyleName & """> *</td><td width='15%' align=right>初始模式：</td><td width='35%'>" & s_FormInitMode & " *</td></tr>" & _
			"<tr><td width='15%' align=right>上传组件：</td><td width='35%'>" & s_FormUploadObject & " *</td><td width='15%' align=right>自动目录：</td><td width='35%'>" & s_FormAutoDir & " *</td></tr>" & _
			"<tr><td width='15%' align=right>图片目录：</td><td width='35%'><input type=text  size=20 name=d_dir title='存放此样式图片文件的目录名，必须在ButtonImage下，最大50个字符长度' value=""" & sStyleDir & """> *</td><td width='15%' align=right>样式目录：</td><td width='35%'><input type=text  size=20 name=d_css title='存放此样式css文件的目录名，必须在CSS下，最大50个字符长度' value=""" & sStyleCSS & """> *</td></tr>" & _
			"<tr><td width='15%' align=right>最佳宽度：</td><td width='35%'><input type=text  name=d_width size=20 title='最佳引用效果的宽度，数字型' value='" & sStyleWidth & "'> *</td><td width='15%' align=right>最佳高度：</td><td width='35%'><input type=text  name=d_height size=20 title='最佳引用效果的高度，数字型' value='" & sStyleHeight & "'> *</td></tr>" & _
			"<tr><td width='15%' align=right>状 态 栏：</td><td width='35%'>" & s_FormStateFlag & " *</td><td width='15%' align=right>Word粘贴：</td><td width='35%'>" & s_FormDetectFromWord & " *</td></tr>" & _
			"<tr><td width='15%' align=right>远程文件：</td><td width='35%'>" & s_FormAutoRemote & " *</td><td width='15%' align=right>指导方针：</td><td width='35%'>" & s_FormShowBorder & " *</td></tr>" & _
			"<tr><td colspan=4>&nbsp;&nbsp;&nbsp;上传文件及系统文件路径相关设置（只有在使用相对路径模式时，才要设置显示路径和内容路径）：</td></tr>" & _
			"<tr><td width='15%' align=right>路径模式：</td><td width='35%'>" & s_FormBaseUrl & " *</td><td width='15%' align=right>上传路径：</td><td width='35%'><input type=text  size=20 name=d_uploaddir title='上传文件所存放路径，相对eWebEditor根目录文件的路径，最大50个字符长度' value=""" & sStyleUploadDir & """> *</td></tr>" & _
			"<tr><td width='15%' align=right>显示路径：</td><td width='35%'><input type=text  size=20 name=d_basehref title='显示内容页所存放路径，必须以&quot;/&quot;开头，最大50个字符长度' value=""" & sStyleBaseHref & """></td><td width='15%' align=right>内容路径：</td><td width='35%'><input type=text  size=20 name=d_contentpath title='实际保存在内容中的路径，相对显示路径的路径，不能以&quot;/&quot;开头，最大50个字符长度' value=""" & sStyleContentPath & """></td></tr>" & _
			"<tr><td colspan=4>&nbsp;&nbsp;&nbsp;允许上传文件类型及文件大小设置（文件大小单位为KB，0表示没有限制）：</td></tr>" & _
			"<tr><td width='15%' align=right>图片类型：</td><td width='35%'><input type=text  name=d_imageext size=20 title='用于图片相关的上传，最大250个字符长度' value='" & sStyleImageExt & "'></td><td width='15%' align=right>图片限制：</td><td width='35%'><input type=text  name=d_imagesize size=20 title='数字型，单位KB' value='" & sStyleImageSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>Flash类型：</td><td width='35%'><input type=text  name=d_flashext size=20 title='用于插入Flash动画，最大250个字符长度' value='" & sStyleFlashExt & "'></td><td width='15%' align=right>Flash限制：</td><td width='35%'><input type=text  name=d_flashsize size=20 title='数字型，单位KB' value='" & sStyleFlashSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>媒体类型：</td><td width='35%'><input type=text  name=d_mediaext size=20 title='用于插入媒体文件，最大250个字符长度' value='" & sStyleMediaExt & "'></td><td width='15%' align=right>媒体限制：</td><td width='35%'><input type=text  name=d_mediasize size=20 title='数字型，单位KB' value='" & sStyleMediaSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>其它类型：</td><td width='35%'><input type=text  name=d_fileext size=20 title='用于插入其它文件，最大250个字符长度' value='" & sStyleFileExt & "'></td><td width='15%' align=right>其它限制：</td><td width='35%'><input type=text  name=d_filesize size=20 title='数字型，单位KB' value='" & sStyleFileSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>远程类型：</td><td width='35%'><input type=text  name=d_remoteext size=20 title='用于自动上传远程文件，最大250个字符长度' value='" & sStyleRemoteExt & "'></td><td width='15%' align=right>远程限制：</td><td width='35%'><input type=text  name=d_remotesize size=20 title='数字型，单位KB' value='" & sStyleRemoteSize & "'></td></tr>" & _
			"<tr><td align=right>备注说明：</td><td colspan=3><textarea rows=2 cols=58 name=d_memo title='此样式的说明，更有利于调用'>" & sStyleMemo & "</textarea></td></tr>" & s_Button & _
			"</form>" & _
			"</table>"& _
"</td>"& _
 "</tr>"& _
"</table><BR>"

	Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#CCCCCC'>"& _
  "<tr>" & _
    "<td height='22' bgcolor='#FFFFFF'>&nbsp;路径模式设置说明：</td>"& _
 "</tr>"& _
  "<tr>" & _
    "<td height='22' bgcolor='#FFFFFF'>&nbsp;<b>相对路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""UploadFile/...""或""../UploadFile/...""形式呈现，当使用此模式时，显示路径和内容路径必填，显示路径必须以""/""开头和结尾，内容路径设置中不能以""/""开头。</td>"& _
  "</tr>"& _
  "<tr>"& _
    "<td height='22' bgcolor='#FFFFFF'>&nbsp;<b>绝对根路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""/eWebEditor/UploadFile/...""这种形式呈现，当使用此模式时，显示路径和内容路径不必填。</td>"& _
  "</tr>"& _
  "<tr>"& _
    "<td height='22' bgcolor='#FFFFFF'>&nbsp;<b>绝对全路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""http://xxx.xxx.xxx/eWebEditor/UploadFile/...""这种形式呈现，当使用此模式时，显示路径和内容路径不必填。</td>"& _
  "</tr>"& _
"</table><BR>"

End Sub

' 初始化样式表数据
Sub InitStyle()
	Dim b
	b = False
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		sSql = "select * from ewebeditor_style where s_id=" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			sStyleName = oRs("S_Name")
			sStyleDir = oRs("S_Dir")
			sStyleCSS = oRs("S_CSS")
			sStyleUploadDir = oRs("S_UploadDir")
			sStyleBaseHref = oRs("S_BaseHref")
			sStyleContentPath = oRs("S_ContentPath")
			sStyleWidth = CStr(oRs("S_Width"))
			sStyleHeight = CStr(oRs("S_Height"))
			sStyleMemo = oRs("S_Memo")
			nStyleIsSys = oRs("S_IsSys")
			sStyleFileExt = oRs("S_FileExt")
			sStyleFlashExt = oRs("S_FlashExt")
			sStyleImageExt = oRs("S_ImageExt")
			sStyleMediaExt = oRs("S_MediaExt")
			sStyleRemoteExt = oRs("S_RemoteExt")
			sStyleFileSize = oRs("S_FileSize")
			sStyleFlashSize = oRs("S_FlashSize")
			sStyleImageSize = oRs("S_ImageSize")
			sStyleMediaSize = oRs("S_MediaSize")
			sStyleRemoteSize = oRs("S_RemoteSize")
			sStyleStateFlag = CStr(oRs("S_StateFlag"))
			sStyleAutoRemote = CStr(oRs("S_AutoRemote"))
			sStyleShowBorder = CStr(oRs("S_ShowBorder"))
			sStyleUploadObject = CStr(oRs("S_UploadObject"))
			sStyleAutoDir = CStr(oRs("S_AutoDir"))
			sStyleDetectFromWord = oRs("S_DetectFromWord")
			sStyleInitMode = oRs("S_InitMode")
			sStyleBaseUrl = oRs("S_BaseUrl")
			b = True
		End If
		oRs.Close
	End If
	If b = False Then
		Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
	End If
End Sub

' 检测样式表单提交的有效性
Sub CheckStyleForm()
	sStyleName = Trim(Request("d_name"))
	sStyleDir = Trim(Request("d_dir"))
	sStyleCSS = Trim(Request("d_css"))
	sStyleUploadDir = Trim(Request("d_uploaddir"))
	sStyleBaseHref = Trim(Request("d_basehref"))
	sStyleContentPath = Trim(Request("d_contentpath"))
	sStyleWidth = Trim(Request("d_width"))
	sStyleHeight = Trim(Request("d_height"))
	sStyleMemo = Request("d_memo")
	sStyleImageExt = Request("d_imageext")
	sStyleFlashExt = Request("d_flashext")
	sStyleMediaExt = Request("d_mediaext")
	sStyleRemoteExt = Request("d_remoteext")
	sStyleFileExt = Request("d_fileext")
	sStyleImageSize = Request("d_imagesize")
	sStyleFlashSize = Request("d_flashsize")
	sStyleMediaSize = Request("d_mediasize")
	sStyleRemoteSize = Request("d_remotesize")
	sStyleFileSize = Request("d_filesize")
	sStyleStateFlag = Request("d_stateflag")
	sStyleAutoRemote = Request("d_autoremote")
	sStyleShowBorder = Request("d_showborder")
	sStyleUploadObject = Request("d_uploadobject")
	sStyleAutoDir = Request("d_autodir")
	sStyleDetectFromWord = Request("d_detectfromword")
	sStyleInitMode = Request("d_initmode")
	sStyleBaseUrl = Request("d_baseurl")

	sStyleUploadDir = Replace(sStyleUploadDir, "\", "/")
	sStyleBaseHref = Replace(sStyleBaseHref, "\", "/")
	sStyleContentPath = Replace(sStyleContentPath, "\", "/")
	If Right(sStyleUploadDir, 1) <> "/" Then sStyleUploadDir = sStyleUploadDir & "/"
	If Right(sStyleBaseHref, 1) <> "/" Then sStyleBaseHref = sStyleBaseHref & "/"
	If Right(sStyleContentPath, 1) <> "/" Then sStyleContentPath = sStyleContentPath & "/"

	If sStyleName = "" Or Get_TrueLen(sStyleName) > 50 Then
		Go_Error "样式名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleName) = False Then
		Go_Error "样式名请勿包含特殊字符！"
	End If
	If sStyleDir = "" Or Get_TrueLen(sStyleDir) > 50 Then
		Go_Error "按钮图片目录名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleDir) = False Then
		Go_Error "按钮图片目录名请勿包含特殊字符！"
	End If
	If sStyleCSS = "" Or Get_TrueLen(sStyleCSS) > 50 Then
		Go_Error "样式CSS目录名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleCSS) = False Then
		Go_Error "样式CSS目录名请勿包含特殊字符！"
	End If

	If sStyleUploadDir = "" Or Get_TrueLen(sStyleUploadDir) > 50 Then
		Go_Error "上传路径不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleUploadDir) = False Then
		Go_Error "上传路径请勿包含特殊字符！"
	End If
	Select Case sStyleBaseUrl
	Case "0"
		If sStyleBaseHref = "" Or Get_TrueLen(sStyleBaseHref) > 50 Then
			Go_Error "当使用相对路径模式时，显示路径不能为空，且不大于50个字符长度！"
		End If
		If IsSafeStr(sStyleBaseHref) = False Then
			Go_Error "当使用相对路径模式时，显示路径请勿包含特殊字符！"
		End If
		If Left(sStyleBaseHref, 1) <> "/" Then
			Go_Error "当使用相对路径模式时，显示路径必须以&quot;/&quot;开头！"
		End If

		If sStyleContentPath = "" Or Get_TrueLen(sStyleContentPath) > 50 Then
			Go_Error "当使用相对路径模式时，内容路径不能为空，且不大于50个字符长度！"
		End If
		If IsSafeStr(sStyleContentPath) = False Then
			Go_Error "当使用相对路径模式时，内容路径请勿包含特殊字符！"
		End If
		If Left(sStyleContentPath, 1) = "/" Then
			Go_Error "当使用相对路径模式时，内容路径不能以&quot;/&quot;开头！"
		End If
	Case "1", "2"
		sStyleBaseHref = ""
		sStyleContentPath = ""
	End Select
	
	If IsNumeric(sStyleWidth) = False Then
		Go_Error "请填写有效的最佳引用宽度！"
	End If
	If IsNumeric(sStyleHeight) = False Then
		Go_Error "请填写有效的最佳引用高度！"
	End If

	If Get_TrueLen(sStyleImageExt) > 250 Then
		Go_Error "图片文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleFlashExt) > 250 Then
		Go_Error "Flash文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleMediaExt) > 250 Then
		Go_Error "媒体文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleFileExt) > 250 Then
		Go_Error "其它文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleRemoteExt) > 250 Then
		Go_Error "远程文件类型不能大于250个字符长度！"
	End If

	If IsNumeric(sStyleImageSize) = False Then
		Go_Error "请填写有效的图片限制大小！"
	End If
	If IsNumeric(sStyleFlashSize) = False Then
		Go_Error "请填写有效的Flash限制大小！"
	End If
	If IsNumeric(sStyleMediaSize) = False Then
		Go_Error "请填写有效的媒体文件限制大小！"
	End If
	If IsNumeric(sStyleFileSize) = False Then
		Go_Error "请填写有效的其它文件限制大小！"
	End If
	If IsNumeric(sStyleRemoteSize) = False Then
		Go_Error "请填写有效的远程文件限制大小！"
	End If

End Sub

' 样式新增保存
Sub DoStyleAddSave()
	sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "'"
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		Go_Error "此样式名已经存在，请用另一个样式名！"
	End If
	oRs.Close

	sSql = "select * from ewebeditor_style where s_id=0"
	oRs.Open sSql, oConn, 1, 3
	oRs.AddNew
	oRs("S_Name") = sStyleName
	oRs("S_Dir") = sStyleDir
	oRs("S_CSS") = sStyleCSS
	oRs("S_UploadDir") = sStyleUploadDir
	oRs("S_BaseHref") = sStyleBaseHref
	oRs("S_ContentPath") = sStyleContentPath
	oRs("S_Width") = sStyleWidth
	oRs("S_Height") = sStyleHeight
	oRs("S_Memo") = sStyleMemo
	oRs("S_ImageExt") = sStyleImageExt
	oRs("S_FlashExt") = sStyleFlashExt
	oRs("S_MediaExt") = sStyleMediaExt
	oRs("S_FileExt") = sStyleFileExt
	oRs("S_RemoteExt") = sStyleRemoteExt
	oRs("S_ImageSize") = sStyleImageSize
	oRs("S_FlashSize") = sStyleFlashSize
	oRs("S_MediaSize") = sStyleMediaSize
	oRs("S_FileSize") = sStyleFileSize
	oRs("S_RemoteSize") = sStyleRemoteSize
	oRs("S_StateFlag") = sStyleStateFlag
	oRs("S_AutoRemote") = sStyleAutoRemote
	oRs("S_ShowBorder") = sStyleShowBorder
	oRs("S_UploadObject") = sStyleUploadObject
	oRs("S_AutoDir") = sStyleAutoDir
	oRs("S_DetectFromWord") = sStyleDetectFromWord
	oRs("S_InitMode") = sStyleInitMode
	oRs("S_BaseUrl") = sStyleBaseUrl
	oRs.Update
	sStyleID = oRs("S_ID")
	oRs.Close
	Response.Write "<script>alert('样式增加成功！');location='Admin_Style.asp';</script>"
	Response.end
End Sub


' 样式修改保存
Sub DoStyleSetSave()
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		' 是否存在同名
		sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "' and s_id<>" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			Go_Error "此样式名已经存在，请用另一个样式名！"
		End If
		oRs.Close
		
		sSql = "select * from ewebeditor_style where s_id=" & sStyleID
		oRs.Open sSql, oConn, 1, 3
		If Not oRs.Eof Then
			'If oRs("S_IsSys") = 1 Then
				'Go_Error "系统样式，不允许修改！"
			'End If
		Else
			Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
		End If
		oRs("S_Name") = sStyleName
		oRs("S_Dir") = sStyleDir
		oRs("S_CSS") = sStyleCSS
		oRs("S_UploadDir") = sStyleUploadDir
		oRs("S_BaseHref") = sStyleBaseHref
		oRs("S_ContentPath") = sStyleContentPath
		oRs("S_Width") = sStyleWidth
		oRs("S_Height") = sStyleHeight
		oRs("S_Memo") = sStyleMemo
		oRs("S_ImageExt") = sStyleImageExt
		oRs("S_FlashExt") = sStyleFlashExt
		oRs("S_MediaExt") = sStyleMediaExt
		oRs("S_FileExt") = sStyleFileExt
		oRs("S_RemoteExt") = sStyleRemoteExt
		oRs("S_ImageSize") = sStyleImageSize
		oRs("S_FlashSize") = sStyleFlashSize
		oRs("S_MediaSize") = sStyleMediaSize
		oRs("S_RemoteSize") = sStyleRemoteSize
		oRs("S_FileSize") = sStyleFileSize
		oRs("S_StateFlag") = sStyleStateFlag
		oRs("S_AutoRemote") = sStyleAutoRemote
		oRs("S_ShowBorder") = sStyleShowBorder
		oRs("S_UploadObject") = sStyleUploadObject
		oRs("S_AutoDir") = sStyleAutoDir
		oRs("S_DetectFromWord") = sStyleDetectFromWord
		oRs("S_InitMode") = sStyleInitMode
		oRs("S_BaseUrl") = sStyleBaseUrl
		oRs.Update
		oRs.Close
	Else
		Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
	End If
	Response.Write "<script>alert('样式修改成功！');location='Admin_Style.asp';</script>"
	Response.end
End Sub

' 样式删除
Sub DoStyleDel()
	If nStyleIsSys = 0 Then
		sSql = "delete from ewebeditor_style where s_id=" & sStyleID
		oConn.Execute sSql
		sSql = "delete from ewebeditor_toolbar where s_id=" & sStyleID
		oConn.Execute sSql
	End If
End Sub

' 样式预览
Sub ShowStylePreview()
	Response.Expires = -1 
	Response.ExpiresAbsolute = Now() - 1 
	Response.cachecontrol = "no-cache"
	'网页不会被缓存
	Response.Write "<html><head>" & _
		"<title>样式预览</title>" & _
		"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"<link href='Include/style.css' rel='stylesheet' type='text/css'>"&_
		"</head><body oncontextmenu='return false'>" & _
"<table width='" & sStyleWidth & "' border='0' align='center' cellpadding='10' cellspacing='1'>"&_
    "<tr><td bgcolor='#D7DEF8'><fieldset style='padding: 5'>"&_
      "<legend><strong>样式预览</strong>&nbsp;<a href='#' onClick='history.go(-1)'>返回菜单</a></legend>"&_
	    "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"&_
		  "<tr bgcolor='#FFFFFF'><td>"&_
		"<input type=hidden name=content1  value=''>" & _
		"<iframe ID='eWebEditor1' src='ewebeditor.asp?id=content1&style=" & sStyleName & "' frameborder=0 scrolling=no width='" & sStyleWidth & "' HEIGHT='" & sStyleHeight & "'></iframe>" & _
    "</td></tr></table></fieldset></td></tr></table></body></html>"
End Sub

' 显示引用代码
Sub ShowStyleCode()
	Response.Write "<BR><table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#D7DEF8'>"&_
  "<tr>"&_
    "<td height='22' bgcolor='#FFFFFF'>样式（<b>" & outHTML(sStyleName) & "</b>）的最佳调用代码如下（其中XXX按实际关联的表单项进行修改）：</td>"&_
  "</tr>"&_
    "<tr>"&_
    "<td height='22' bgcolor='#FFFFFF'><textarea rows=5 cols=65 style='width:100%'><IFRAME ID=""eWebEditor1"" SRC=""ewebeditor.asp?id=XXX&style=" & sStyleName & """ FRAMEBORDER=""0"" SCROLLING=""no"" WIDTH=""" & sStyleWidth & """ HEIGHT=""" & sStyleHeight & """></IFRAME></textarea></td>"&_
  "</tr>"&_
"</table><br>"
End Sub

' 显示工具栏表单列表
Sub ShowToolBarList()
	Response.Write "<BR><table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#CCCCCC'><tr>" & _
		"<td height='22' bgcolor='#FFFFFF'>&nbsp;样式（" & outHTML(sStyleName) & "）下的工具栏管理：</td>" & _
		"<tr>"& _
		"<td height='22' bgcolor='#FFFFFF'>"

	Dim s_AddForm, s_ModiForm

	' 增加表单
	If nStyleIsSys = 1 Then
		s_AddForm = ""
	Else
		' 取当前最大排序号
		Dim nMaxOrder
		sSql = "select max(T_order) from ewebeditor_toolbar where s_id=" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If IsNull(oRs(0)) Then
			nMaxOrder = 1
		Else
			nMaxOrder = oRs(0) + 1
		End If
		oRs.Close

		s_AddForm = "<hr width='95%' align=center size=1><table border=0 cellpadding=4 cellspacing=0 align=center>" & _
		"<form action='?id=" & sStyleID & "&action=toolbaradd' name='addform' method=post>" & _
		"<tr><td>工具栏名：<input type=text name=d_name size=20  value='工具栏" & nMaxOrder & "'> 排序号：<input type=text name=d_order size=5 value='" & nMaxOrder & "' > <input type=submit name=b1 value='新增工具栏'></td></tr>" & _
		"</form></table><hr width='95%' align=center size=1>"
	End If

	' 修改表单
	Dim s_Manage, s_SubmitButton
	s_ModiForm = "<form action='?id=" & sStyleID & "&action=toolbarmodi' name=modiform method=post>" & _
		"<table border=0 cellpadding=0 cellspacing=0 align=center class=list1>" & _
		"<tr align=center>ID工具栏名排序号操作</tr>"
	sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID & " order by t_order asc"
	oRs.Open sSql, oConn, 0, 1
	Do While Not oRs.Eof
		s_Manage = "<a href='?id=" & sStyleID & "&action=buttonset&toolbarid=" & oRs("t_id") & "'>按钮设置</a>"
		If nStyleIsSys <> 1 Then
			s_Manage = s_Manage & "|<a href='?id=" & sStyleID & "&action=toolbardel&delid=" & oRs("t_id") & "'>删除</a>"
		End If
		s_ModiForm = s_ModiForm & "<tr align=center>" & _
			"<td>" & oRs("t_id") & "</td>" & _
			"<td><input type=text name='d_name" & oRs("t_id") & "' value=""" & inHTML(oRs("t_name")) & """ size=30 ></td>" & _
			"<td><input type=text name='d_order" & oRs("t_id") & "' value='" & oRs("t_order") & "' size=5 ></td>" & _
			"<td>" & s_Manage & "</td>" & _
			"</tr>"
		oRs.MoveNext
	Loop
	oRs.Close
	s_SubmitButton = "<input type=submit name=b1 value=' 修改 '>"
	s_ModiForm = s_ModiForm & "</table>"& _
	"<BR><table width='100%' border='0' cellpadding='0' cellspacing='0'>"& _
  "<tr>"& _
   "<td height='22' bgcolor='#FFFFFF' align='center'>" & s_SubmitButton & "</td>"& _
  "</tr>"& _
"</table></form></td></tr></table><br>"

	' 输出表单
	Response.Write s_AddForm & s_ModiForm
End Sub

' 新增工具栏
Sub DoToolBarAdd()
	Dim s_Name, s_Order
	s_Name = Trim(Request("d_name"))
	s_Order = Trim(Request("d_order"))
	If s_Name = "" Or Get_TrueLen(s_Name) > 50 Then
		Go_Error "工具栏名不能为空，且长度不能大于50个字符长度！"
	End If
	If IsNumeric(s_Order) = False Then
		Go_Error "无效的工具栏排序号，排序号必须为数字！"
	End If
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式下的工具栏，不允许新增！"
	End If
	sSql = "select * from ewebeditor_toolbar where 1=0"
	oRs.Open sSql, oConn, 1, 3
	oRs.AddNew
	oRs("s_id") = sStyleID
	oRs("t_name") = s_Name
	oRs("t_order") = s_Order
	oRs("t_button") = ""
	oRs.Update
	oRs.Close
	Response.Write "<script>alert(""工具栏（" & outHTML(s_Name) & "）增加操作成功！"");</script>"
End Sub

' 修改工具栏
Sub DoToolBarModi()
	Dim s_Name, s_Order
	sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID
	oRs.Open sSql, oConn, 1, 3
	Do While Not oRs.Eof
		s_Name = Trim(Request("d_name" & oRs("t_id")))
		s_Order = Trim(Request("d_order" & oRs("t_id")))
		If s_Name <> "" And IsNumeric(s_Order) = True Then
			If s_Name <> oRs("t_name") Or s_Order <> CStr(oRs("t_Order")) Then
				oRs("t_name") = s_Name
				oRs("t_order") = s_Order
				oRs.Update
			End If
		End If
		oRs.MoveNext
	Loop
	oRs.Close
	Response.Write "<script>alert('工具栏修改操作成功！');location='Admin_Style.asp';</script>"
	Response.end
End Sub

' 删除工具栏
Sub DoToolBarDel()
	Dim s_DelID
	s_DelID = Trim(Request("delid"))
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式下的工具栏，不允许删除！"
	End If
	If IsNumeric(s_DelID) = True Then
		sSql = "delete from ewebeditor_toolbar where s_id=" & sStyleID & " and t_id=" & s_DelID
		oConn.Execute sSql
		Response.Write "<script language=javascript>alert('工具栏（ID：" & s_DelID & "）删除操作成功！');</script>"
	End If
End Sub

' 初始化工具栏
Sub InitToolBar()
	Dim b
	b = False
	sToolBarID = Trim(Request("toolbarid"))
	If IsNumeric(sToolBarID) = True Then
		sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID & " and t_id=" & sToolBarID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			sToolBarName = oRs("T_Name")
			sToolBarOrder = oRs("T_Order")
			sToolBarButton = oRs("T_Button")
			b = True
		End If
		oRs.Close
	End If
	If b = False Then
		Go_Error "无效的工具栏ID号，请通过页面上的链接进行操作！"
	End If
End Sub

' 按钮设置
Sub ShowButtonList()
	Dim i, n


	' 取所有按钮
	Dim aButtonCode(), aButtonTitle(), aButtonImage()
	sSql = "select * from ewebeditor_button where b_allowselect=1 order by b_order asc"
	oRs.Open sSql, oConn, 0, 1
	i = 0
	Do While Not oRs.Eof
		i = i + 1
		Redim Preserve aButtonCode(i)
		Redim Preserve aButtonTitle(i)
		Redim Preserve aButtonImage(i)
		aButtonCode(i) = oRs("B_Code")
		aButtonTitle(i) = oRs("B_Title")
		aButtonImage(i) = oRs("B_Image")
		oRs.MoveNext
	Loop
	oRs.Close


	' 取可选列表
	Dim s_Option1
	s_Option1 = ""
	For i = 1 To UBound(aButtonCode)
		s_Option1 = s_Option1 & "<option value='" & aButtonCode(i) & "'>" & aButtonTitle(i) & "</option>"
	Next

	' 取已选列表
	Dim aButton, s_Option2, s_Temp
	aButton = Split(sToolBarButton, "|")
	s_Option2 = ""
	For i = 0 To UBound(aButton)
		s_Temp = Code2Title(aButton(i), aButtonCode, aButtonTitle)
		If s_Temp <> "" Then
			s_Option2 = s_Option2 & "<option value='" & aButton(i) & "'>" & s_Temp & "</option>"
		End If
	Next


	'以下为客户端操作选择脚本
	'''''''''''''''''''''''''''''''''''
%>

<script language=javascript>
// 加入已选
function Add() {
	var sel1=document.myform.d_b1;
	var sel2=document.myform.d_b2;
	if (sel1.selectedIndex<0) {
		alert("请选择一个待选按钮！");
		return;
	}
	sel2.options[sel2.length]=new Option(sel1.options[sel1.selectedIndex].innerHTML,sel1.options[sel1.selectedIndex].value);
}

// 从已选中删除
function Del() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if (nLen<1) return;
	if (nIndex<0) {
		alert("请选择一个已选按钮！");
		return;
	}
	for (var i=nIndex;i<nLen-1;i++) {
		sel.options[i].value=sel.options[i+1].value;
		sel.options[i].innerHTML=sel.options[i+1].innerHTML;
	}
	sel.length=nLen-1;
}

// 排序：向上移动
function Up() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==0)) return;
	if (nIndex<0) {
		alert("请选择一个要移动的已选按钮！");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex-1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex-1].innerHTML;
	sel.options[nIndex-1].value=sValue;
	sel.options[nIndex-1].innerHTML=sHTML;
	sel.selectedIndex=nIndex-1;
}

// 排序：向下移动
function Down() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==nLen-1)) return;
	if (nIndex<0) {
		alert("请选择一个要移动的已选按钮！");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex+1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex+1].innerHTML;
	sel.options[nIndex+1].value=sValue;
	sel.options[nIndex+1].innerHTML=sHTML;
	sel.selectedIndex=nIndex+1;
}

// 提交处理
function checkform() {
	var sel=document.myform.d_b2;
	var nLen = sel.length;
	var str="";
	for (var i=0;i<nLen;i++) {
		if (i>0) str+="|";
		str+=sel.options[i].value;
	}
	document.myform.d_button.value=str;
	return true;
}

</script>

<%
	'''''''''''''''''''''''''''''''''''


	' 选择设置表单
	Dim s_SubmitButton
		s_SubmitButton = "<input type=submit name=b value=' 保存设置 '>"
	Response.Write "<BR><table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#D7DEF8'>"& _
  "<tr>"& _
    "<td height='22' align='center' bgcolor='#FFFFFF'>"	
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 align=center>" & _
		"<form action='?action=buttonsave&id=" & sStyleID & "&toolbarid=" & sToolBarID & "' method=post name=myform onsubmit='return checkform()'>" & _
		"<tr align=center><td>可选按钮</td><td></td><td>已选按钮</td><td></td></tr>" & _
		"<tr align=center>" & _
			"<td><select name='d_b1' size=20 style='width:200px' ondblclick='Add()'>" & s_Option1 & "</select></td>" & _
			"<td><input type=button name=b1 value=' → ' onclick='Add()'><br><br><input type=button name=b1 value=' ← ' onclick='Del()'></td>" & _
			"<td><select name='d_b2' size=20 style='width:200px' ondblclick='Del()'>" & s_Option2 & "</select></td>" & _
			"<td><input type=button name=b3 value='↑' onclick='Up()'><br><br><br><input type=button name=b4 value='↓' onclick='Down()'></td>" & _
		"</tr>" & _
		"<input type=hidden name='d_button' value=''>" & _
		"<tr><td colspan=4 align=right>" & s_SubmitButton & "</td></tr>" & _
		"</form></table>"

Response.Write "</td></tr></table>"

	' 显示图片对照表
	Response.Write "<BR>以下是按钮图片对照表（部分下拉框或特殊按钮可能没图）：<Br>"
	Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#D7DEF8'>"
	n = 0
	For i = 1 To UBound(aButtonCode)
		n = i Mod 4
		If n = 1 Then
			Response.Write "<tr>"
		End If
		Response.Write "<td height='22' align='center' bgcolor='#FFFFFF'>"
		If aButtonImage(i) <> "" Then
			Response.Write "<img border=0 align=absmiddle src='ButtonImage/standard/" & aButtonImage(i) & "'>"
		End If
		Response.Write aButtonTitle(i)
		Response.Write "</td>"
		If n = 0 Then
			Response.Write "</tr>"
		End If
	Next
	If n > 0 Then
		For i = 1 To 4 - n
			Response.Write "<td bgcolor='#FFFFFF'>&nbsp;</td>"
		Next
		Response.Write "</tr>"
	End if
	Response.Write "</table><br><br>"
End Sub

' 由按钮代码得到按钮标题
Function Code2Title(s_Code, a_ButtonCode, a_ButtonTitle)
	Dim i
	Code2Title = ""
	For i = 1 To UBound(a_ButtonCode)
		If UCase(a_ButtonCode(i)) = UCase(s_Code) Then
			Code2Title = a_ButtonTitle(i)
			Exit Function
		End If
	Next
End Function

' 按钮设置保存
Sub DoButtonSave()
	Dim s_Button
	s_Button = Trim(Request("d_button"))
	sSql = "select * from ewebeditor_toolbar where t_id=" & sToolBarID
	oRs.Open sSql, oConn, 1, 3
	If Not oRs.Eof Then
		oRs("T_Button") = s_Button
		oRs.Update
	End If
	oRs.Close
	Response.Write "<script>alert('工具栏按钮设置保存成功！');location='Admin_Style.asp';</script>"
	Response.end
End Sub

Sub RemoveApplication()
	Dim aApplicationName, i
	aApplicationName = Application("eWebEditor_ApplicationName")
	If IsArray(aApplicationName) = True Then
		For i = 1 To UBound(aApplicationName)
			Application.Contents.Remove(aApplicationName(i))
		Next
		Application.Contents.Remove("eWebEditor_ApplicationName")
		Application.Contents.Remove("eWebEditor_ApplicationUrl")
	End If
End Sub

%>