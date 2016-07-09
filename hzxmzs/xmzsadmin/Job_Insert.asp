<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<%
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'验证用户权限
	Rs,Sql
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from [Job]"
	Rs.open sql,conn,1,3
	Rs.Addnew
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("Explain") = Trim(Request.Form("Explain"))
	Rs("Address") = Trim(Request.Form("Address"))
	Rs("EndDate") = Trim(Request.Form("EndDate"))
	Rs("Number") = Trim(Request.Form("Job_Number"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("Rank") = Trim(Request.Form("Rank"))
	Rs("DateTime") = Date()
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('添加成功!');location='Job_List.asp';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/Date.JS"></script>
</head>

<body oncontextmenu="return false">
<form name="theForm" method="post" action="" onSubmit="return input_ok(this)">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td> <fieldset style="padding: 5" class="body">
        <legend><strong> 人才招聘</strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
              频道类别：</td>
            <td> <strong><font color="#FF0000">诚聘英才</font></strong> </td>
          </tr>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">岗位名称：</td>
            <td bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" size="30"> 
              &nbsp;&nbsp; </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">结束日期：</td>
            <td bgcolor="#FFFFFF">
			<script language="JavaScript">arrowtag("EndDate",'');</script></td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">工资待遇：</td>
            <td bgcolor="#FFFFFF"><input name="Explain" type="text" id="Explain" size="30"> 
            </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">工作地点：</td>
            <td bgcolor="#FFFFFF"><input name="Address" type="text" id="Address" size="30"> 
            </td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">招聘人数：</td>
            <td bgcolor="#FFFFFF"><input name="Job_Number" type="text" id="Job_Number" size="20">
              [正整数]</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">排序：</td>
            <td bgcolor="#FFFFFF"> <input name="Rank" type="text" id="Rank" size="4" maxlength="4">
              [正整数]</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">内容：</td>
            <td bgcolor="#FFFFFF"> <textarea name="Content" cols="60" rows="10"></textarea> 
              &nbsp;<img src="../sysImages/sizeminus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows-=2" alt="减小编辑区"> 
              <img src="../sysImages/sizeplus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows+=2" alt="增高编辑区"> 
            </td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"> <input type="submit" name="Submit_OK" value="添加(A)" accesskey="a"> 
              &nbsp; <input type="button" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
			  <input name="Action" type="hidden" id="Action" value="Insert">
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
<% 
call rs_close()
call rc_close()
call db_close()
 %>
