<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<%
If Trim(Request.Form("Action"))="Edit" Then
	Call Check_Right(3)	'��֤�û�Ȩ��
	ID = Request.Form("ID")
	SS_ID = Trim(Request.Form("SS_ID"))
	name = Trim(Request.Form("name"))
	Img = Trim(Request.Form("Img"))
	area = Trim(Request.Form("area"))
	price = Trim(Request.Form("price"))
	designer = Trim(Request.Form("designer"))
	intro =Trim(Request.Form("intro"))
	status = Trim(Request.Form("status")) 
	content = Trim(Request.Form("content"))
	Sql = "Update Site Set SS_ID="&SS_ID&",name='"&name&"',Img='"&Img&"',area='"&area&"',price='"&price&"',designer='"&designer&"',intro='"&intro&"',status='"&status&"',content='"&content&"' where ID="&ID
	conn.execute(sql)
	Response.Write "<script>alert('�޸ĳɹ�!');location='Site.asp?SS_ID="&SS_ID&"';</script>"
	Response.end()
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ID = Request.QueryString("ID")
Set RS = Server.CreateObject("ADODB.RecordSet")
Sql = "Select * from Site where ID="&ID
RS.OPEN Sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ʩ���ֳ�</title>
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
        <legend><strong> ��Ϣչʾ </strong></legend>
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
            <td bgcolor="#FFFFFF" width="60%"><input name="name" type="text" id="name" value="<%= rs("name") %>" style="width:500px"/>
              &nbsp;&nbsp; </td>
            <td width="170" rowspan="5" align="center" valign="middle" bgcolor="#FFFFFF"><img name="logo" width="120" src="/UploadFile/<%= replace(rs("Img"),".","a.") %>" height="120" border="0" id="logo" style="CURSOR: hand" title="��ʾ:�鿴ͼƬʵ�ʴ�С" onClick="newimg(this.src);" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">��ҳʩ��ͼƬ��</td>
            <td bgcolor="#FFFFFF"><input name="Img" type="text" id="Img" value="<%= rs("Img") %>" style="width:500px"/>
              <input name="Submit_Open" type="button" id="Submit_Open" value="�ϴ�" onClick="NewWindow('Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;" /></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���������</td>
            <td bgcolor="#FFFFFF"><input name="area" type="text" id="area" value="<%= rs("area") %>" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">������ۣ�</td>
            <td bgcolor="#FFFFFF"><input name="price" type="text" id="price" value="<%= rs("price") %>" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���ʦ��</td>
            <td bgcolor="#FFFFFF"><input name="designer" type="text" id="designer" value="<%= rs("designer") %>" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">���˵����</td>
            <td bgcolor="#FFFFFF"><input name="intro" type="text" id="intro" value="<%= rs("intro") %>" style="width:500px"/></td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">����״̬��</td>
            <td bgcolor="#FFFFFF">
				<input name="status" type="radio" id="status1" value="1" <%If rs("status") = 1 Then  response.write "checked"%>/>ʩ��ǰ��
				<input name="status" type="radio" id="status2" value="2" <%If rs("status") = 2 Then  response.write "checked"%>/>ʩ������
				<input name="status" type="radio" id="status3" value="3" <%If rs("status") = 3 Then  response.write "checked"%>/>ʩ��ĩ��
				<input name="status" type="radio" id="status3" value="4" <%If rs("status") = 4 Then  response.write "checked"%>/>���̽���
			</td>
          </tr>
          <tr>
            <td height="22" align="center" bgcolor="#FFFFFF">��ϸ����˵��<br>���ͼƬ<br><br><br><font color="red">ͼƬ��С�������700���ؿ�֮��</font></td>
            <td bgcolor="#FFFFFF">
				<textarea name="content" style="display:none"> <%= rs("content") %> </textarea> 
        <iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width="650" HEIGHT="500"></iframe>
			</td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30" align="center"><input type="submit" name="Submit_OK" value="�޸�" />
              &nbsp;
              <input type="button" name="reset" value="����"onClick="javascript:history.go(-1);" />
              <input name="Action" type="hidden" id="Action" value="Edit" />
              <input name="SS_ID" type="hidden" id="SS_ID" value="<%= SS_ID %>" />
			  <input name="ID" type="hidden" id="ID" value="<%= ID %>" />
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
