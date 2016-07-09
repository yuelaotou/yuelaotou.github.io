<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1 
Response.ExpiresAbsolute = Now() - 1 
Response.cachecontrol = "no-cache"
'网页不会被缓存
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Database
Dim conn				'定义连接对象
Database = 0		'0:Access,1:SQLserver
Select Case Database
	Case 0
		'ASP连接ACCESS数据库 
		DBpath = Server.MapPath("../Database/#xmzs.mdb")
		'DBpath = "D:\site\admin\Database\db.mdb"
		on error resume next
		
		Set conn=Server.CreateObject("ADODB.Connection")
		Conn_String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBpath
		Conn.Open Conn_String
		If Err.Number<>0 Then
			Response.Write("<div align='left'>")
			Response.Write("错误代码："&Err.Number&"<br>")
			Response.Write("错误对象："&Err.Source&"<br>")	
			Response.Write("错误描述："&Err.Description&"")
			Response.Write("</div>")
		End If
	Case 1
		'Asp连接MS SQL server 2000 数据库
		Dim Server_Name			'服务器名
		Dim User_ID				'登录帐号
		Dim Password 			'登录密码
		Dim DB_Name				'数据库名
		Server_Name = "(local)"		'服务器名
		User_ID = "sa"				'登录帐号
		Password = "sa" 				'登录密码
		DB_Name = "ahqcedu"			'数据库名
		'打开数据库
		On Error Resume Next
		dim conn_String			'SQL Server连接字符串
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn_String = "Driver={SQL Server};Server="&Server_Name&";Uid="&User_ID&";Pwd="&Password&";Database="&DB_Name&""
		Conn.Open Conn_String
		If Err.Number<>0 Then	'查找错误
			Response.Write("<div align='left'>")
			Response.Write("错误代码："&Err.Number&"<br>")
			Response.Write("错误对象："&Err.Source&"<br>")	
			Response.Write("错误描述："&Err.Description&"")
			Response.Write("</div>")
		End If
End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'关闭记录集
Sub Rs_Close()
	Rs.close
	Set Rs = Nothing
End Sub
'关闭记录集
Sub RC_Close()
	RC.close
	Set RC = Nothing
End Sub
'关闭数据库
Sub DB_Close()
	Conn.Close
	Set Conn = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>