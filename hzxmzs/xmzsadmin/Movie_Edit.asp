<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->

<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'验证用户权限
	ID = Trim(Request.Form("ID"))
	SS_ID = Trim(Request.Form("SS_ID"))
	set Rs=server.createobject("adodb.recordset")
	Sql="select * from Movie where id="&id
	Rs.open sql,conn,1,3
	Rs("Title") = Trim(Request.Form("Title"))
	Rs("Author") = Trim(Request.Form("Author"))
	Rs("Img") = Trim(Request.Form("Img"))
	Rs("address") = Trim(Request.Form("address"))
	Rs("BeginTime") = Trim(Request.Form("BeginTime"))
	Rs("Movie") = Trim(Request.Form("Upfile"))
	Rs("Flag") = Trim(Request.Form("Flag"))
	Rs("Content") = Trim(Request.Form("Content"))
	Rs("Rank") = Trim(Request.Form("Rank"))
	Rs("SS_ID") = SS_ID
	Rs("DateTime") = Date()
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('修改成功!');location='Movie.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ID = Request.QueryString("ID")
Sql = "Select * from Movie where id="&id
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open Sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript" type="text/JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="../JS/NewImage.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../js/Date.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
</head>

<body oncontextmenu="return false">
<script language="JavaScript">
function input_ok(){
    var text="";

//验证用户名
    var M1 = theForm.Title.value; 
    if (M1 == ""){text=text+"[电影名称] 不能为空；\n"} 
	
//验证用户名
    var M2 = theForm.Author.value; 
    if (M2 == ""){text=text+"[主演者] 必须选择；\n"} 

//验证姓名
    var M3 = theForm.address.value; 
    if (M3 == ""){text=text+"[电影地区] 不能为空；\n"} 

//验证姓名
    var M4 = theForm.BeginTime.value; 
    if (M4 == ""){text=text+"[上映时间] 不能为空；\n"} 
	
//验证用户名
    var M5 = theForm.Upfile.value; 
    if (M5 == ""){text=text+"[视频] 必须上传；\n"} 

//验证电话
   var M6 = theForm.Content.value;
   if (M6 == ""){text=text+"[内容] 不能为空；\n"} 

//向页面反馈信息	
   if (text == "") {
	   //document.theForm.submit();
	   }
   else {
    alert("出错提示：\n\n" + text);
    return false;
    }           
}


</script>

<form name="theForm" method="post" action="" onSubmit="return input_ok()">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr> 
      <td> <fieldset style="padding: 5" class="body">
        <legend><strong>宣传视频</strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100" height="22" align="center"> <img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle"> 
              频道类别：</td>
            <td width="150"> <strong><font color="#FF0000">
			<% 
Dim SS_ID
SS_ID = Request.QueryString("SS_ID")
set Rc=server.createobject("adodb.recordset")
Sql="select * from Type Where ID="&SS_ID&""
Rc.open sql,conn,1,1
Response.Write(Rc("Title"))
Rc.close
SET Rc=nothing
%>

			
			
			</font></strong></td>
            <td align="center">&nbsp;</td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D7DEF8">
          <tr> 
            <td width="100" height="22" align="center" bgcolor="#FFFFFF">名称：</td>
            <td bgcolor="#FFFFFF"><input name="Title" type="text" id="Title" size="30" value="<%= RS("Title") %>"> 
              &nbsp;&nbsp; </td>
            <td width="200" rowspan="9" align="center" valign="middle" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF">视频：</td>
            <td bgcolor="#FFFFFF"> <input name="Upfile" type="text" id="Upfile" size="35"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF">内容：</td>
            <td bgcolor="#FFFFFF"> <textarea name="Content" cols="45" rows="5"><%= rs("Content") %></textarea> 
              &nbsp; <img src="../sysImages/sizeminus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows-=2" alt="减小编辑区"> 
              <img src="../sysImages/sizeplus.gif" width="20" height="20" style="CURSOR: hand" onClick="Content.rows+=2" alt="增高编辑区">            </td>
          </tr>
        </table>
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30" align="center"> <input type="submit" name="Submit_OK" value="添加(A)" accesskey="a"> 
              &nbsp; <input type="button" name="reset" value="返回(B)" accesskey="B" onClick="javascript:history.go(-1);"> 
			  <input name="Action" type="hidden" id="Action" value="Edit">
			 <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>">
			 <input name="ID" type="hidden" id="ID" value="<%= ID %>">
            </td>
          </tr>
        </table>
      </fieldset></td>
    </tr>
  </table>
</form>
</body>
</html>
