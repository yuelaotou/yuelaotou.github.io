<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Include/Function.asp" -->
<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/Config.asp"-->

<%
if request("send")="add" then
nickname=request.Form("nickname")
email=request.Form("email")
tel=request.Form("tel")
qq=request.Form("qq")
title=request.Form("title")
content=request.Form("content")

Set rs=server.CreateObject("ADODB.RecordSet")
    rs.open "select * from message",conn,1,3
	rs.addnew
	rs("nickname")=nickname
	rs("email")=email
	rs("tel")=tel
	rs("qq")=qq
	rs("title")=title
	response.Write myReplace(content)
	rs("content")=content
	rs("datetime")=now()   
	rs.update
	rs.close  
	set rs=nothing  
response.Write "<script>alert('留言成功，敬请等待我们的回复！');location='../message.htm'</script>"
end if
%>