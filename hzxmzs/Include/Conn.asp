<%
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1 
Response.ExpiresAbsolute = Now() - 1 
Response.cachecontrol = "no-cache"
'��ҳ���ᱻ����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Database
Dim conn				'�������Ӷ���
Database = 0		'0:Access,1:SQLserver
Select Case Database
	Case 0
		'ASP����ACCESS���ݿ� 
		DBpath = Server.MapPath("../Database/#xmzs.mdb")
		'DBpath = "D:\site\admin\Database\db.mdb"
		on error resume next
		
		Set conn=Server.CreateObject("ADODB.Connection")
		Conn_String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBpath
		Conn.Open Conn_String
		If Err.Number<>0 Then
			Response.Write("<div align='left'>")
			Response.Write("������룺"&Err.Number&"<br>")
			Response.Write("�������"&Err.Source&"<br>")	
			Response.Write("����������"&Err.Description&"")
			Response.Write("</div>")
		End If
	Case 1
		'Asp����MS SQL server 2000 ���ݿ�
		Dim Server_Name			'��������
		Dim User_ID				'��¼�ʺ�
		Dim Password 			'��¼����
		Dim DB_Name				'���ݿ���
		Server_Name = "(local)"		'��������
		User_ID = "sa"				'��¼�ʺ�
		Password = "sa" 				'��¼����
		DB_Name = "ahqcedu"			'���ݿ���
		'�����ݿ�
		On Error Resume Next
		dim conn_String			'SQL Server�����ַ���
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn_String = "Driver={SQL Server};Server="&Server_Name&";Uid="&User_ID&";Pwd="&Password&";Database="&DB_Name&""
		Conn.Open Conn_String
		If Err.Number<>0 Then	'���Ҵ���
			Response.Write("<div align='left'>")
			Response.Write("������룺"&Err.Number&"<br>")
			Response.Write("�������"&Err.Source&"<br>")	
			Response.Write("����������"&Err.Description&"")
			Response.Write("</div>")
		End If
End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�رռ�¼��
Sub Rs_Close()
	Rs.close
	Set Rs = Nothing
End Sub
'�رռ�¼��
Sub RC_Close()
	RC.close
	Set RC = Nothing
End Sub
'�ر����ݿ�
Sub DB_Close()
	Conn.Close
	Set Conn = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>