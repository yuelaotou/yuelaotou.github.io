<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
if Trim(Request.QueryString("Action"))="Del" then
	id=Trim(Request.QueryString("id"))
	Parent_Count=conn.execute("select count(*) from Type where Parent_id="&id)(0)
	if Parent_Count=0 then
	Call del("Type",id,"?")
	else
	Call OutScript("ϵͳ��ʾ:\n\n����ɾ����������")
	end if
end if



%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<script language="JavaScript" type="text/javascript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/Tree.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<link href="../CSS/Tree.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="10">
  <tr>
    <td><fieldset style="padding: 5" class="body">
        
      <legEnd><img src="../sysImages/Station.gIF" width="24" height="22" align="absmiddle"><strong>վ��ṹ</strong>&nbsp; 
      <a href="javascript:parent.location.reload();"> <img src="../sysImages/reload.gIF" alt="ˢ��Ƶ��" width="15" height="18" border="0" align="absmiddle">      </a><a href="Type_Insert.asp" target="Type_Main"> <img src="../sysImages/sizeplus.gIF" alt="���Ƶ��" width="20" height="20" border="0" align="absmiddle"> 
      </a> &nbsp;</legEnd>
	  <table width="100%" border="0" cellspacing="0" cellpadding="10">
        <tr>
          <td>
		  <%
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
					Module_Title="ģ��:"&rc("Title")
					rc.close()
				end if
             %>	
			  <tr>
                <td class='<%= menutype %>' ID='b<%= RS("ID") %>'<%= onmouseup %> height="25"><a href="#" title="���:<%=rs("id")%><br>·��:<%=rs("SS_Path")%><br><%= Module_Title %>"><%= rs("Title")%></a>
				 <a href="Type_Insert.asp?SS_Path=<%= rs("SS_Path") %>&Parent_ID=<%= rs("ID") %>" target="Type_Main"><img src="../sysImages/sizeplus.gif" border="0" align="absmiddle"></a>
	&nbsp;
				<a href="Type_Edit.asp?id=<%=rs("id")%>" target="Type_Main"><img src="../sysImages/Edit.gif" border="0" align="absmiddle"></a>&nbsp;<a onClick="{if(confirm('ȷ��Ҫɾ����?')){location='?Action=Del&id=<%=rs("id")%>';return true;};return false}"><img src="../sysImages/del.gif" align="absmiddle" border="0"></a>
				</td>
              </tr>
			  <% 
 			  IF ChildCount>0 then
  			  %> 
              <tr id='a<%=rs("id")%>'>
                <td class='<%= listtype %>'>
				<% tree(RS("ID")) %>
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
      </table>
    
	</fieldset>
    </td>
  </tr>
</table>
<%
Call RS_Close()
Call RC_Close()
Call DB_Close()
%>
</body>
</html>
