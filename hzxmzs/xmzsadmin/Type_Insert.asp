<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
SS_Path=Trim(Request.QueryString("SS_Path"))
Parent_id=Trim(Request.QueryString("Parent_id"))
if SS_Path = empty then
	SS_Path = "/"
end if
if Parent_id = empty then
	Parent_id = 0
end if
if Trim(Request.Form("Action"))="insert" then
	Call Check_Right(2)	'验证用户权限
	SS_Path=Trim(Request.Form("SS_Path"))
	Title=Trim(Request.Form("Title"))
	Rank=cint(Request.Form("Rank"))
	Module_id=cint(Request.Form("Module_id"))
	Locks=Trim(Request.Form("Lock"))
	sql="insert into Type(Parent_id,Title,Rank,Module_id,Lock) values("&Parent_id&",'"&Title&"',"&Rank&","&Module_id&","&Locks&")"
	conn.execute(sql)
	'查询,更新
	sql = "select id,SS_Path from Type"
	Set Rs=Server.CreateObject("ADODB.RecordSet")
	Rs.Open sql,conn,1,3
	Rs.Movelast
	Id = rs("id")
	IF SS_Path = "/" then
		SS_Path = ID
	else
		SS_Path = SS_Path & "-" & ID
	End if
		Rs("SS_Path") = SS_Path
		Rs.update
		Rs.requery
		rs.close
		'起刷新左边窗口的作用
		Response.Write("<script>")
		Response.Write("parent.Type_left.window.location='Type_List.asp';")
		Response.Write("location='Type_Insert.asp';")
		Response.Write("</script>")
end if
%>



<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>

<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td>
	<fieldset style="padding: 5" class="body">
	 <legend><strong>添加模块</strong> </legend>
	
	
	
	
	 
	   <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#AAEF92">
         <form name="form1" method="post" action="">
		 <tr>
           <td width="100" align="center" bgcolor="#FFFFFF">路径:</td>
           <td bgcolor="#FFFFFF"><label>
             <input name="SS_Path" type="text" id="SS_Path" size="20" readonly="0" value="<%=SS_Path%>">
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">标题:</td>
           <td bgcolor="#FFFFFF"><label>
             <input name="Title" type="text" id="Title" size="20">
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">排序:</td>
           <td bgcolor="#FFFFFF"><label>
             <input name="Rank" type="text" id="Rank" size="5">
           [正整数]</label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">模块:</td>
           <td bgcolor="#FFFFFF"><label>
             <select name="Module_id" size="1">
     			<option>选择模块</option>
				
				 <%
			   sql="select * from Module_id"
			   set rs=conn.execute(sql)
			   do until rs.eof
			   %>
               <option value="<%=rs("id")%>"><%=rs("Title")%></option>
			  
			  <%rs.movenext
			  loop
			  %> 
             </select>
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">锁定:</td>
           <td bgcolor="#FFFFFF"><label>
             <select name="lock" size="1" id="lock">
               <option selected="selected">选择锁定</option>
               <option value="1">锁定</option>
               <option value="0">不锁定</option>
             </select>
           </label></td>
         </tr>
         <tr>
           <td colspan="2" align="center" bgcolor="#FFFFFF"><label>
             <input type="submit" name="Submit" value="提交">
             &nbsp;
             <input name="Action" type="hidden" value="insert">
             <input type="reset" name="Submit2" value="重置">
             <input name="Parent_ID" type="hidden" value="<%= Parent_id %>">
           </label></td>
         </tr>
	     </form>
      </table>
        
	 </fieldset>
	</td>
  </tr>
</table>
</body>
</html>
<% 
call Rs_Close()
call DB_Close()
 %>