<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
if Trim(Request.Form("Action"))="Menu" then
	Call Check_Right(9)	'验证用户权限
	id=Trim(Request.Form("id"))
	Menu=Trim(Request.Form("Menu"))
	Sql = "update admin set Menu='"&Menu&"' where id="&id
	conn.execute(Sql)
	Response.Write("<script>")
	Response.Write("parent.Left_main.window.location='Left.asp';")
	Response.Write("location='User.asp';")
	Response.Write("</script>")
	Call alert("授权成功","User.asp")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script language="JavaScript" src="../JS/Alert.JS"></script>
<link href="../css/Tree.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/Tree.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	
	<!-- 用户菜单 -->
<form name="form1" method="post" action="">
	  <table width="400" border="0" cellpadding="10" cellspacing="10">
        <tr>
          <td>
		  <fieldset style="padding: 5" class="body">
              <legend><strong>授权菜单</strong>&nbsp;&nbsp;全选&nbsp;<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" title="全选">&nbsp;&nbsp;</legend>
		  
		  
		  
		      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2"><img src="../sysImages/Station.gIF" width="24" height="22" align="absmiddle">根目录</td>
                </tr>
                <tr>
                  <td width="10">&nbsp;</td>
                  <td>
				  <%
				  ID=Trim(Request.QueryString("id"))
				  Menu_SQL="select * from admin where id="&id
				  Set rc=server.CreateObject("ADODB.RecordSet")
				  rc.open Menu_SQL,conn,1,1
				  if rc("Menu")<>"" then
				  Menus=split(rc("Menu"),",")
				  end if
				  
				  
				  
		  Function Tree(Parent_id)
		  sql = "select * from type where Parent_id="&Parent_id&" order by Rank desc"
		  set rs=server.CreateObject("ADODB.RecordSet")
		  rs.open sql,conn,1,1
		  %>
              <table border="0" cellspacing="0" cellpadding="0">
                <%
				i = 1
				do until RS.eof
				ChildCount=conn.execute("select count(*) from Type where Parent_ID="&RS("ID"))(0)
				
				if ChildCount=0 then
					if i=RS.recordcount then
						menutype="file1"
					else
						menutype="file"
					end if
					onmouseup=""
				else
					if i=RS.recordcount then
						menutype="menu3"
						listtype="list1"
						onmouseup=" onMouseUp=change1('a"&RS("ID")&"','b"&RS("ID")&"');"
					else
						menutype="menu1"
						listtype="list"
						onmouseup=" onMouseUp=change2('a"&RS("ID")&"','b"&RS("ID")&"');"
					end if
				end if
				if rs("Parent_id")<>0 then
				Module_sql="select * from Module_id where id="&rs("Module_id")
				set rc=conn.execute(Module_sql)
				module_Title ="模块:"&rc("Title")
				rc.close()
				end if
             %>
                <tr>
                  <td class='<%= menutype %>' id='b<%= RS("ID") %>'<%= onmouseup %> height="25"><a title="<%= module_Title %>"><%= rs("Title")%></a> 
						<% 
	Flag = 0
	J=0
    For J = LBound(menus) To UBound(menus)
      IF Trim(Rs("ID"))=Trim(menus(J)) Then
	     Flag = 1
	  End If
    Next	
 %>		  
				        <input name="Menu" type="checkbox" id="Menu" value="<%=rs("id")%>" <%if flag=1 then%>checked<%end if%> title="编号:<%=rs("id")%>">				  </td>
                </tr>

                <% 
 			  IF ChildCount>0 then
  			  %>
                <tr id='a<%=rs("id")%>' style="display:none">
                  <td class='<%= listtype %>'><% tree(RS("ID")) %>
                  </td>
                </tr>
                <% 
			end if
			rs.movenext	
			i=i+1	
			loop		

  			 %>
              </table>
          <%
		End Function
		call Tree(0)
		%>
			<br>
			  
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center">
<input type="submit" name="Submit_OK" value="菜单(M)" accesskey="M" >
              &nbsp; <input type="Reset" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
              <input name="Action" type="hidden" value="Menu"> 
			  <input name="ID" type="hidden" value="<%= id %>"> 
			  </td>
          </tr>
        </table>	  
				  
				  
				  </td>
                </tr>
              </table>
		  </fieldset>
		  </td>
        </tr>
      </table>
	
    
	
</form>	
	
    </td>
  </tr>
</table>
</body>
</html>
