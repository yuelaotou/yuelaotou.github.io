<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/Alert.JS"></script>
<link href="../css/Tree.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/Tree.JS"></script>
<script language="JavaScript" src="../JS/Clock.JS"></script>
<script language="JavaScript">

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
	eval("submenu" + sid + ".style.display=\"\";");
	imgmenu.background="../sysImages/menuup.gif";
}
else
{
	eval("submenu" + sid + ".style.display=\"none\";");
	imgmenu.background="../sysImages/menudown.gif";
}
}

</script>
<style type="text/css">
<!--
.td_line {	border: 1px solid #FFFFFF;
}
.STYLE1 {	color: #333300;
	font-weight: bold;
}
-->
</style>
</head>

<body  onLoad="startClock()" bgcolor="#56C1F5">
<table width="150" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25" align="center" background="../sysImages/ltitlebg.gif"><span class="STYLE1">管理系统</span></td>
  </tr>
  <tr>
    <td height="25" align="center" background="../sysImages/title_bg_quit.gif"><a href="Exit.asp" title="退出系统">退出</a>&nbsp;&nbsp;<a href="javascript:location.reload();"title="刷新菜单">刷新</a></td>
  </tr>
</table><br />
<table width="150" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25" align="center" background="../sysImages/title_bg_quit.gif">
		<a href="CreateAll.asp" target="main" title="生成首页">生成首页</a>
	</td>
  </tr>
  <tr height="2"></tr>
  <tr>
    <td height="25" align="center" background="../sysImages/title_bg_quit.gif">
		<a href="CreateList.asp" target="main" title="生成栏目页">生成栏目页</a>
	</td>
  </tr>
</table><br />
<table width="150" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25" align="center" background="../sysImages/menudown.gif" class="menu_title" id="imgmenu0" style="cursor:hand" onClick="showsubmenu(0)">系统菜单</td>
  </tr>
  <tr class="body">
    <td colspan="2" align="center" class="td_line" id="submenu0" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20">&nbsp;</td>
      </tr>
      <tr>
        <td height="20">&nbsp;<img src="../sysImages/Station.gIF" width="24" height="22" align="absmiddle" />根目录</td>
      </tr>
      <tr>
        <td height="20">
		<%
			UID = Session("UID")
		  Menu_SQL="select * from admin where UID='"&UID&"'"
		  Set rc=server.CreateObject("ADODB.RecordSet")
		  rc.open Menu_SQL,conn,1,1
		  if rc("Menu")<>Empty then
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
					module_Url = rc("Url") & "?SS_ID="&RS("ID")
				rc.close()
				Else
					module_Url = "#"
				end if
             %>
			 <% 
			Flag = 0
			For J = LBound(menus) To UBound(menus)
			  IF Trim(Rs("ID"))=Trim(menus(J)) Then
				 Flag = 1
			  End If
			Next	
 			%>		  
			 <%if flag=1 then%>
                <tr>
                  <td class='<%= menutype %>' id='b<%= RS("ID") %>'<%= onmouseup %> height="25"><% If Module_url<>"#" Then %><a href="<%= module_Url %>" target="main"><%= rs("Title")%></a><% Else %><%= rs("Title")%><% End If %> </td>
                </tr>
				<%end if%>
				
				
				
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
        </td>
      </tr>
      <tr>
        <td height="20" align="center">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<TABLE width=150 border="0" align=center cellPadding=0 cellSpacing=0>
  <TR>
    <td height="25" align="center" background="../sysImages/menudown.gif" class="menu_title" style="cursor:hand">系统时间</TR>
        <TR class="body">
          
    <TD colspan="2" align="center" class="td_line"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20">&nbsp;</td>
        </tr>
        <tr> 
          <td height="20" align="center">

		  <div id="show"></div>

		  
		  </td>
        </tr>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table> </TD>
</TR>
</TABLE>
</body>
</html>
