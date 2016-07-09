<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.Form("Action")) = "Reply" Then
	ID=Trim(Request.Form("ID"))
	set rs=server.CreateObject("ADODB.RecordSet")
	SQL="select * from Books where id="&id
	rs.open SQL,conn,1,3
	rs("Reply")=Trim(Request.Form("Content"))
	rs("ReplyTime")=Now()
	rs("Flag")=1
	rs.update
	rs.requery
	rs.close()
	response.Write "<script>alert('回复成功');location='Books.asp?SS_ID="&rs("SS_ID")&"</script>"
End if

''''''''''''''''''''''''''''''''''''''''''
id = Trim(Request.QueryString("id"))
set rs=server.CreateObject("adodb.recordset")
SQL = "Select * from Books where id="&id
rs.open SQL,conn,1,1
%>
<html">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="post" action="">
  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td><fieldset style="padding: 5">
        <legend><strong>服务频道</strong> </legend>
        <BR>
        <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><fieldset style="padding: 5" class="body">
              <legend><font color="#FF0000">回复留言</font></legend>
              <table width="100%" height="100" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
                <tr bgcolor="#FFFFFF">
                  <td width="150" height="25" align="center">编号：<%= rs("id") %>&nbsp; </td>
                  <td height="25">&nbsp;&nbsp;留言时间：<%= rs("DateTime") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="20" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="center">&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="center"><img src="../sysImages/head/<%= rs("head") %>.gif" width="32" height="32"></td>
                      </tr>
                      <tr>
                        <td align="center">&nbsp;</td>
                      </tr>
                  </table></td>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp; <%= rs("Content") %></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" align="center"><font color=red>姓名：<%= rs("Name") %></font></td>
                  <td height="25" align="right"><table width="400" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="100"><% if rs("Tel")<>Empty then %>
                          电话：<%= rs("Tel") %>
                          <% end if %></td>
                        <td width="100"><% if rs("qq")<>Empty then %>
                            <img src="../sysImages/QQ.GIF" width="16" height="16" align="absmiddle"> <a href="http://wpa.qq.com/msgrd?V=1&Uin=<%= rs("qq") %>" target="_blank" title="点击可以立即给我发送消息"><%= rs("qq") %></a>
                            <% end if %>
                        </td>
                        <td width="100"><img src="../sysImages/MAIL.GIF" width="16" height="16" align="absmiddle">
                            <% if rs("Email")<>Empty then %>
                            <a href="mailto:<%= rs("Email") %>">电邮</a>
                            <% else %>
                          电邮
                          <% end if %>
                        </td>
                        <td width="100">&nbsp;<%= rs("ip") %>&nbsp;</td>
                      </tr>
                  </table></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" colspan="2"><textarea name="content" id="content" style="display:none"><%=Server.HTMLEncode(rs("Reply"))%></textarea>
                      <iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width="650" HEIGHT="350"></iframe></td>
                </tr>
              </table>
            </fieldset></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="回复">
              &nbsp;
              <input type="button" name="reset" value="返回" onClick="javascript:history.go(-1);">
              <input name="Action" type="hidden" id="Action" value="Reply">
              <input name="ID" type="hidden" id="Action" value="<%= ID %>">
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
