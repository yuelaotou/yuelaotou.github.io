<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Insert" Then
	Call Check_Right(2)	'��֤�û�Ȩ��
	SS_ID = Trim(Request.Form("SS_ID"))
	name = Trim(Request.Form("name"))
	Img = Trim(Request.Form("Img"))
	area = Trim(Request.Form("area"))
	price = Trim(Request.Form("price"))
	designer = Trim(Request.Form("designer"))
	intro =Trim(Request.Form("intro"))
	status = Trim(Request.Form("status")) 
	content = Trim(Request.Form("content"))
	set Rs = Server.CreateObject("adodb.recordset")
	Sql="select * from Site"
	Rs.open sql,conn,1,3
	Rs.Addnew
	Rs("SS_ID") = SS_ID
	Rs("name") = name
	Rs("Img") = Img
	Rs("area") = area
	Rs("price") = price
	Rs("designer") = designer
	Rs("intro") = intro
	RS("status") = status
	Rs("content") = content
	Rs("lock") = "0"
	Rs("date") = date()
	Rs.Update
	RS.close
	Set RS = nothing
	Response.Write "<script>alert('��ӳɹ�!');location='Site.asp?SS_ID="&SS_ID&"';</script>"
	Response.end
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ʩ���ֳ����</title>
<script language="JavaScript" type="text/javascript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/NewImage.JS"></script>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/Date.JS"></script>


</head>

<body oncontextmenu="return false">
<form name="theForm" method="post" action="">
  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td bgcolor="#F0E7CF"><fieldset style="padding: 5">
        <legend><strong>��Ϣչʾ</strong></legend>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> Ƶ�����</td>
            <td width="150"><strong><font color="#FF0000">
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
            <td width="15%" height="22" align="center" bgcolor="#FFFFFF">�������ƣ�</td>
            <td bgcolor="#FFFFFF" width="60%"><input name="name" type="text" id="name" style="width:500px"/>
              &nbsp;&nbsp; </td>
            <td width="170" rowspan="5" align="center" valign="middle" bgcolor="#FFFFFF"><img name="logo" width="120" height="120" border="0" id="logo" style="CURSOR: hand" title="��ʾ:�鿴ͼƬʵ�ʴ�С" onClick="newimg(this.src);" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">��ҳʩ��ͼƬ��</td>
            <td bgcolor="#FFFFFF"><input name="Img" type="text" id="Img" size="35" />
              <input name="Submit_Open" type="button" id="Submit_Open" value="�ϴ�" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���������</td>
            <td bgcolor="#FFFFFF"><input name="area" type="text" id="area" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">������ۣ�</td>
            <td bgcolor="#FFFFFF"><input name="price" type="text" id="price" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���ʦ��</td>
            <td bgcolor="#FFFFFF"><input name="designer" type="text" id="designer" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���˵����</td>
            <td bgcolor="#FFFFFF"><input name="intro" type="text" id="intro" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">����״̬��</td>
            <td bgcolor="#FFFFFF">
				<input name="status" type="radio" id="status1" value="1" checked/>ʩ��ǰ��
				<input name="status" type="radio" id="status2" value="2"/>ʩ������
				<input name="status" type="radio" id="status3" value="3"/>ʩ��ĩ��
				<input name="status" type="radio" id="status4" value="4"/>���̽���
			</td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">��ϸ����˵��<br>���ͼƬ<br><br><br><font color="red">ͼƬ��С�������700���ؿ�֮��</font></td>
            <td bgcolor="#FFFFFF">
				<textarea name="content" style="display:none"></textarea> 
        <iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width="650" HEIGHT="500"></iframe>
			</td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="���" />
              &nbsp;
              <input type="button" name="reset" value="����"onClick="javascript:history.go(-1);" />
              <input name="Action" type="hidden" id="Action" value="Insert" />
              <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>" />
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
