<%
IF Session("UID")=Empty Then
	Response.Write "<script>"
	Response.Write "alert('非法用户,请登陆!');parent.location='index.asp';"
	Response.Write "</script>"
	Response.end
End if

%>
