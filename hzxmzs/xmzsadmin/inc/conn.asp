<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Response.Buffer = True
On Error Resume Next
set conn=Server.CreateObject("ADODB.Connection") 
conn.open "driver={SQL Server};database=zhende918;Server=203.158.16.22;uid=zhende918_f;pwd=dtyh306997" 
If Err Then
err.Clear
set conn=nothing
response.Write"���ݿ����ӳ������������ִ�!"
response.End
End If
sub login()
dim myurl,outurl
myurl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if myurl="" then
response.write "<script>alert('��ǰ��������ֹ��');location.href='index.asp';</script>"
Response.End
end if
if session("adminame")="" then
response.write "<script>top.location.href='index.asp';</script>"
Response.End
end if
end sub
%>
