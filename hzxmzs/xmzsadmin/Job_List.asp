<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
'删除单个
If Trim(Request.QueryString("Action"))="Del" Then
	Call Check_Right(4)	'验证用户权限
	ID = Trim(Request.QueryString("ID"))
	Call Del("[Job]",ID,"?")
End if 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'锁定操作
If Trim(Request.QueryString("Action"))="Lock" Then
	Call Check_Right(5)	'验证用户权限
	ID = Trim(Request.QueryString("ID"))
	Flag = Trim(Request.QueryString("Flag"))
	ReCord = Trim(Request.QueryString("ReCord"))
	Call Lock_OK("[Job]",ID,ReCord,Flag,"?")
End if 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Rs,Sql
set Rs=server.createobject("adodb.recordset")
Sql="select * from [Job] order by ID desc"
Rs.open sql,conn,1,1
Rs.pagesize=24
If Len(Request.QueryString("Page")) = 0 Then 
	Page = 1 
Else 
	Page=Cint(Request.Querystring("Page")) 
End If 
Rs.AbsolutePage=Page

%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>

</head>

<body oncontextmenu="return false">

  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
    <tr>
      <td> 
        <fieldset style="padding: 5" class="body">
        <legend><strong>诚聘英才</strong>		</legend>
		
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
            频道类别：</td>
          <td> <strong><font color="#FF0000">诚聘英才</font></strong> </td>
        </tr>
      </table>
	  
        
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
        <tr bgcolor="#FFFFFF"> 
          <td width="50" height="22" align="center">编号</td>
          <td align="center">职位</td>
          <td width="80" align="center">工作地点</td>
          <td width="80" align="center">发布时间</td>
          <td width="80" align="center">结束日期</td>
          <td width="50" align="center">查看</td>
          <td width="30" align="center">修改</td>
          <td width="30" align="center">锁定</td>
          <td width="30" align="center">删除</td>
        </tr>
        <%
For I=1 to rs.pagesize
    IF rs.eof THEN EXIT FOR 
 %>
        <tr bgcolor="#FFFFFF" style="cursor:hand" OnMouseOver="this.bgColor='#F0F0F0'" OnMouseOut="this.bgColor='#FFFFFF'"> 
          <td height="22" align="center"><%= rs("id") %></td>
          <td>&nbsp;<%= rs("Title") %></td>
          <td align="center"><%= rs("Address") %></td>
          <td align="center"><%= rs("DateTime") %></td>
          <td align="center">
		  <%
	   DateTime=datediff("D",date(),Rs("EndDate"))
	   IF left(DateTime,1)="-" then
	    %> <img src="../sysImages/Ok.gif" alt="已到期" width="16" height="16" border="0" align="absmiddle" > 
                    <% Else %>
            <font color="#FF0000">剩余&nbsp;<%= DateTime %>&nbsp;天</font> <% End IF %> 
		  
		  </td>
          <td align="center"><a href="Work_list.asp?Jobid=<%= RS("ID") %>"><img src="../sysImages/Search.gif" width="20" height="20" border="0" alt="查看应聘者信息"></a></td>
          <td align="center"> <a href="Job_Edit.asp?ID=<%= rs("ID") %>"><img src="../sysImages/Edit.gif" width="12" height="12" border="0" alt="修改"></a> 
          </td>
          <td align="center"> <% If Rs("Lock") = 0 Then %> <a href="?Action=Lock&ID=<%= rs("ID") %>&ReCord=Lock&Flag=1"> 
            <img src="../sysImages/unlock.gif" width="14" height="13" border="0" alt="锁定"> 
            </a> <% Else %> <a href="?Action=Lock&ID=<%= rs("ID") %>&ReCord=Lock&Flag=0"> 
            <img src="../sysImages/lock.gif" width="11" height="12" border="0" alt="解锁"> 
            </a> <% End If %> </td>
          <td align="center"> <img src="../sysImages/del.gif" width="16" height="16" border="0" onClick="{if(confirm('确定删除吗?')){location='?Action=Del&ID=<%= rs("id") %>';return true;}return false;}"> 
          </td>
        </tr>
        <% 
rs.MoveNext
NEXT
%>
      </table>
<% Call RS_Empty()	'没有记录 %>
  
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%" height="30" align="center" valign="middle"> 
              <% call Image_Page("") %>
            </td>
            <td align="center"> 
			<input type="button" name="Submit_OK" value="添加(A)" accesskey="a" onClick="location='Job_Insert.asp'">
              &nbsp; <input type="button" name="reset" value="刷新(R)" accesskey="R" onClick="javascript:location.reload();"> 
		    </td>
          </tr>
        </table>
</fieldset>
</td>
  </tr>
</table>

</body>
</html>
<% 
call rs_close()
call rc_close()
call db_close()
 %>
