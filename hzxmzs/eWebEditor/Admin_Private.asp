
<%
If Request.Cookies("eWebEditor") = Empty Then
	Response.Write "<script>alert('ϵͳ��ʾ:\n\n����ϵͳ����Ա�鿴��');history.go(-1)</script>"
	Response.End
End If

' ִ��ÿ��ֻ�账��һ�ε��¼�
Call BrandNewDay()

' ��ʼ�����ݿ�����
Call DBConnBegin()

' ���ñ���
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))


' ********************************************
' ����Ϊҳ�湫��������
' ********************************************
' ============================================
' ���ÿҳ���õĶ�������
' ============================================
Sub Header()
	Response.Expires = -1 
	Response.ExpiresAbsolute = Now() - 1 
	Response.cachecontrol = "no-cache"
	'��ҳ���ᱻ����

	Response.Write "<html><head>"
	
	' ��� meta ���
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
	
	' �������
	Response.Write "<title>eWebSoft�����ı��༭�� - ��̨����</title>"
	
    ' ���ÿҳ��ʹ�õĻ�����ʽ��
	Response.Write "<link href='Include/style.css' rel='stylesheet' type='text/css'>"
	' ���ÿҳ��ʹ�õĻ����ͻ��˽ű�
	Response.Write "<script language='javaScript' SRC='Include/private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body bgcolor=#ffffff  oncontextmenu='return false;'>"
	
	' ���ҳ�涥��(Header1)
Response.Write "<table width='100%' border='0' align='center' cellpadding='10' cellspacing='1'>"&_
  "<tr>"&_
    "<td bgcolor='#D7DEF8'><fieldset style='padding: 5'>"&_
      "<legend><strong>�����ı��༭��</strong>&nbsp;<a href='admin_uploadfile.asp?ID=29'>�ϴ�����</a>&nbsp;|&nbsp;<a href='admin_style.asp'>��ʽ����</a>&nbsp;|&nbsp;<a href='admin_decode.asp'>��ȡ����</a></legend>"
	  Response.Write("<table width='100%' border='0' cellspacing='0' cellpadding='0' bgcolor='#FFFFFF'><tr><td>")
End Sub

' ============================================
' ���ÿҳ���õĵײ�����
' ============================================
Sub Footer()
	' �ͷ��������Ӷ���
	Call DBConnEnd()

	' �ײ�����
	Response.Write("</td></tr></table>")
	Response.Write "</fieldset></td></tr></table>"
	Response.Write "</body></html>"
End Sub

' ===============================================
' ��ʼ��������
'	s_FieldName	: ���ص���������	
'	a_Name		: ��ֵ������
'	a_Value		: ��ֵֵ����
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
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
