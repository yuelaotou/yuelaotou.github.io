<%
IF Session("UID")=Empty Then
	Response.Write "<script>"
	Response.Write "alert('�Ƿ��û�,���½!');parent.location='index.asp';"
	Response.Write "</script>"
	Response.end
End if

%>
