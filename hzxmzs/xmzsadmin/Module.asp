<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
IF Trim(Request.QueryString("Action"))="del" Then
	Call Check_Right(4)
	ID=Trim(Request.QueryString("id"))
	Flag=conn.execute("select count(*) from Type where Module_ID="&id)(0)
	IF Flag = 0 then
		call del("Module_ID",id,"?")
	Else
		call alert("请先删除关联类型","?")
	End if
End if

If Trim(Request.Form("Action"))="Del" then
	Call Check_Right(4)
	call del_all("Module_id","?")
End if

IF Trim(Request.QueryString("Action"))="lock" then
	Call Check_Right(5)
	ID=Trim(Request.QueryString("id"))
	locks=cint(Request.QueryString("locks"))
	call Lock_Url("Module_id",id,locks,"?")
	call alert("操作成功","?")
End if

sql="select * from Module_id order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
rs.pagesize=20
If Len(Request.QueryString("Page")) = 0 Then 
	Page = 1 
Else 
	Page=Cint(Request.Querystring("Page")) 
End If 
Rs.AbsolutePage=Page
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr>
    <td><fieldset style="padding: 5" class="body">
      <legend><strong><img src="../SysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" />模块管理</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" height="30" align="center" valign="middle"><% call Image_Page("") %></td>
          <td align="center"><label>
            <input name="Submit_OK" type="button" id="Submit_OK" value="添加" onClick="location='Module_insert.asp'" />
            &nbsp;
            <input name="reset" type="button" id="reset" value="刷新" onClick="javascript:location.reload();" />
          </label></td>
        </tr>
      </table>
	  <form name="form1" method="post" action="">
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr>
          <td width="50" height="22" align="center" bgcolor="#FFFFFF">编号</td>
          <td width="120" align="center" bgcolor="#FFFFFF">标题</td>
          <td align="center" bgcolor="#FFFFFF">文件地址</td>
          <td width="100" align="center" bgcolor="#FFFFFF">窗口目标</td>
          <td width="30" align="center" bgcolor="#FFFFFF">修改</td>
          <td width="30" align="center" bgcolor="#FFFFFF">锁定</td>
          <td width="30" align="center" bgcolor="#FFFFFF">删除</td>
          <td width="30" align="center" bgcolor="#FFFFFF"><label>
            <input name="chkAll" type="checkbox" value="checkbox" onClick="CheckAll(this.form)">
          </label></td>
        </tr>
		<% FOR I=1 TO RS.pagesize 
		if rs.eof then exit for
		%>
        <tr bgcolor="#FFFFFF" style="cursor:hand" onMouseOver="this.bgColor='#FOFOFO'" onMouseOut="this.bgColor='#FFFFFF'">
          <td height="22" align="center"><%=rs("id")%></td>
          <td align="center"><%=rs("title")%></td>
          <td align="center"><%=rs("url")%></td>
          <td align="center"><%=rs("target")%></td>
          <td align="center"><a href="Module_Edit.asp?id=<%=rs("id")%>"><img src="../sysImages/Edit.gif" width="12" height="12" border="0" alt="修改"></a></td>
          <td align="center">
		  <%IF rs("Lock")=0 Then%>
		  <a href="?Action=lock&id=<%=rs("id")%>&locks=1"><img src="../sysImages/Unlock.gif" width="11" height="12" border="0" alt="锁定"></a>
		  <%else%>
		  <a href="?Action=lock&id=<%=rs("id")%>&locks=0"><img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="不锁定"></a>
		  <%End if%>			</td>
          <td align="center"><img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确认要删除吗?')){location='?Action=del&id=<%=rs("id")%>';return true;}return flase;}"></td>
          <td align="center"><label>
            <input name="id" type="checkbox" id="id" value="<%=rs("id")%>">
          </label></td>
        </tr>
		<%rs.movenext
		next
		%>
      </table>
      
      <label>
      <input type="submit" name="Submit" value="删除">
      </label>
	  <input name="Action" type="hidden" id="Action" value="Del">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center"><% call Image_Page("") %></td>
        </tr>
      </table>
	  </form>
    </fieldset></td>
  </tr>
</table>
<% 
call Rs_Close()
call DB_Close()
%>
</body>
</html>
