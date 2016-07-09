<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
if Trim(Request.Form("Action"))="Edit" then
	Call Check_Right(3)	'验证用户权限
	ID=Trim(Request.Form("id"))
	Title=Trim(Request.Form("Title"))
	Rank=cint(Request.Form("Rank"))
	Module_id=cint(Request.Form("Module_id"))
	Locks=Trim(Request.Form("Lock"))
	sql="Update Type set Title='"&Title&"',Rank="&Rank&",Module_id="&Module_id&",Lock="&Locks&" where id="&id
	conn.execute(sql)
	Response.Write("<script>")
	Response.Write("parent.Type_left.window.location='Type_List.asp';")
	Response.Write("location='Type_Insert.asp';")
	Response.Write("</script>")
end if
'.............................................................	
	id=Trim(Request.QueryString("id"))
	sql = "select * from Type where id="&id
	Set Rs=conn.execute(sql)
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
	 <legend><strong>修改模块</strong> </legend>
	
	
	
	
	 
	   <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#AAEF92">
         <form name="form1" method="post" action="">
		 <tr>
           <td width="100" align="center" bgcolor="#FFFFFF">路径:</td>
           <td bgcolor="#FFFFFF"><label>
             <input name="SS_Path" type="text" id="SS_Path" size="20" readonly="0" value="<%=rs("SS_Path")%>" disabled="disabled">
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">标题:</td>
           <td bgcolor="#FFFFFF"><label>
           <input name="Title" type="text" id="Title" size="20" value="<%=rs("Title")%>">
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">排序:</td>
           <td bgcolor="#FFFFFF"><label>
             <input name="Rank" type="text" id="Rank" size="5" value="<%=rs("Rank")%>">
           [正整数]</label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">模块:</td>
           <td bgcolor="#FFFFFF"><label>
             <select name="Module_id" size="1">
               <option>选择模块</option>
			   <%
			   sql="select * from Module_id"
			   set rc=conn.execute(sql)
			   do until rc.eof
			   %>
			   <option value="<%=rc("id")%>" <%if rc("id")=rs("Module_id") then%> selected="selected"<%end if%>><%=rc("Title")%></option>
			  <%rc.movenext
			  loop
			  %> 
             </select>
           </label></td>
         </tr>
         <tr>
           <td align="center" bgcolor="#FFFFFF">锁定:</td>
           <td bgcolor="#FFFFFF"><label>
             <select name="lock" size="1" id="lock">
               <option value="1"<%if rs("lock")=1 then%>selected="selected"<%end if%>>锁定</option>
               <option value="0"<%if rs("lock")=0 then%>selected="selected"<%end if%>>不锁定</option>
             </select>
           </label></td>
         </tr>
         <tr>
           <td colspan="2" align="center" bgcolor="#FFFFFF"><label>
             <input type="submit" name="Submit" value="提交">
             &nbsp;
             <input name="Action" type="hidden" value="Edit">
             <input type="reset" name="Submit2" value="重置">
			 <input name="ID" type="hidden" value="<%=RS("ID")%>">
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