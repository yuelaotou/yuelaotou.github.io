
<%
If Request.Cookies("eWebEditor") = Empty Then
	Response.Write "<script>alert('系统提示:\n\n仅限系统管理员查看！');history.go(-1)</script>"
	Response.End
End If

' 执行每天只需处理一次的事件
Call BrandNewDay()

' 初始化数据库连接
Call DBConnBegin()

' 公用变量
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))


' ********************************************
' 以下为页面公用区函数
' ********************************************
' ============================================
' 输出每页公用的顶部内容
' ============================================
Sub Header()
	Response.Expires = -1 
	Response.ExpiresAbsolute = Now() - 1 
	Response.cachecontrol = "no-cache"
	'网页不会被缓存

	Response.Write "<html><head>"
	
	' 输出 meta 标记
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
	
	' 输出标题
	Response.Write "<title>eWebSoft在线文本编辑器 - 后台管理</title>"
	
    ' 输出每页都使用的基本样式表
	Response.Write "<link href='Include/style.css' rel='stylesheet' type='text/css'>"
	' 输出每页都使用的基本客户端脚本
	Response.Write "<script language='javaScript' SRC='Include/private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body bgcolor=#ffffff  oncontextmenu='return false;'>"
	
	' 输出页面顶部(Header1)
Response.Write "<table width='100%' border='0' align='center' cellpadding='10' cellspacing='1'>"&_
  "<tr>"&_
    "<td bgcolor='#D7DEF8'><fieldset style='padding: 5'>"&_
      "<legend><strong>在线文本编辑器</strong>&nbsp;<a href='admin_uploadfile.asp?ID=29'>上传管理</a>&nbsp;|&nbsp;<a href='admin_style.asp'>样式管理</a>&nbsp;|&nbsp;<a href='admin_decode.asp'>获取函数</a></legend>"
	  Response.Write("<table width='100%' border='0' cellspacing='0' cellpadding='0' bgcolor='#FFFFFF'><tr><td>")
End Sub

' ============================================
' 输出每页公用的底部内容
' ============================================
Sub Footer()
	' 释放数据连接对象
	Call DBConnEnd()

	' 底部导航
	Response.Write("</td></tr></table>")
	Response.Write "</fieldset></td></tr></table>"
	Response.Write "</body></html>"
End Sub

' ===============================================
' 初始化下拉框
'	s_FieldName	: 返回的下拉框名	
'	a_Name		: 定值名数组
'	a_Value		: 定值值数组
'	v_InitValue	: 初始值
'	s_Sql		: 从数据库中取值时,select name,value from table
'	s_AllName	: 空值的名称,如:"全部","所有","默认"
' ===============================================
Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function
%>
