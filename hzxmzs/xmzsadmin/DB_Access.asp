<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
Dim Submit
Submit = Trim(Request.Form("Submit_OK"))
IF Submit <> Empty Then
	Call Check_Right(10)	'��֤�û�Ȩ��
	Select Case Submit
		Case "����(S)":call copyData()
		Case "�ָ�(S)":call RestoreData()
		Case "ѹ��(S)"
		Call RS_Close()'�رռ�¼��
		Call DB_Close()'�ر����ݿ�
		call compactdata()
	End Select
End IF
 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="post" action="">  
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
      <td align="center">
<fieldset style="padding: 5" class="body">
      <legend><strong>����ACCESS���ݿ�</strong></legend>
        <table width="50%" border="0" cellpadding="0" cellspacing="1">
          <tr> 
            <td height="22" colspan="2" align="left"><marquee scrollamount="3" width="350"><font color="#FF0000">���棺</font>��ҳ���漰��վ����ȫ��,�����Ҷ�,���������վ�޷�����,�Ų�����!</marquee></td>
          </tr>
          <tr> 
            <td width="100" height="22" align="center">����Ŀ¼��</td>
            <td><input name="path" type=text value="../database/" size="30"></td>
          </tr>
          <tr> 
            <td height="22" align="center">�������ƣ�</td>
            <td><input name="title" type=text value="DB_Access.Bak" size="30">            </td>
          </tr>
          <tr>
            <td height="22" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
    <td height="30" align="center"><input type="submit" name="Submit_OK" value="����(S)" accesskey="S">
              &nbsp; <input type="Reset" name="reset2" value="����(R)" accesskey="R"></td>
  </tr>
</table>
      </fieldset>
</td>
  </tr>
</table>
</form>



<form name="form1" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
      <td align="center">
<fieldset style="padding: 5" class="body">
      <legend><strong>�ָ�ACCESS���ݿ�</strong></legend>

        <table width="50%" border="0" cellpadding="0" cellspacing="1">
          <tr> 
            <td height="22" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td width="100" height="22" align="center">�������ݿ⣺</td>
            <td><input name="path" type=text id="path" value="../database/DB_Access.Bak" size="30"></td>
          </tr>
          <tr>
            <td height="22" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="30" align="center"><input type="submit" name="Submit_OK" value="�ָ�(S)" accesskey="S">
&nbsp;
      <input type="Reset" name="reset" value="����(R)" accesskey="R"></td>
    </tr>
  </table>
</fieldset>
</td>
  </tr>
</table>

</form>

<form name="form1" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
    <td><fieldset style="padding: 5" class="body">
        <legend><strong>ѹ��ACCESS���ݿ�</strong>&nbsp;���棺���ȱ������ݿ����ѹ�����ݿ�!</legend>

  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
            <td height="26" align="center"> 
              <input type="submit" name="Submit_OK" value="ѹ��(S)" accesskey="S">
&nbsp;
      <input type="Reset" name="reset" value="����(R)" accesskey="R"></td>
    </tr>
  </table>
  </fieldset>
</td>
  </tr>
</table>
</form>
<% 
sub copyData()
	dim path,title,fso,FoundErr
	path=trim(request.form("path"))
	title=trim(request.form("title"))
	if FoundErr=True then exit sub
	path=server.MapPath(path)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.FileExists(dbpath) then
		If fso.FolderExists(path)=false Then
			fso.CreateFolder(path)
		end if
		fso.copyfile dbpath,path & "\" & title
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n�������ݿ�ɹ�!');location='?';</script>"
		Response.end
	Else
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n�Ҳ���Դ���ݿ��ļ�!!');location='?';</script>"
		Response.end
	End if
end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub RestoreData()
	dim path,fso
	path=request.form("path")
	path=server.mappath(path)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(path) then  					
		fso.copyfile path,Dbpath
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n���ݿ�ָ��ɹ�!');location='?';</script>"
		Response.end
	else
		Response.Write "<script>alert('ϵͳ��ʾ:\n\n�Ҳ���ָ���ı����ļ�!!');location='?';</script>"
		Response.end
	end if
end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub compactdata()
datapath=server.MapPath("../database/#mydb.mdb")
strdbpath = left(datapath,instrrev(datapath,"\"))
set fso = server.createobject("scripting.filesystemobject")
IF fso.fileexists(datapath) then
	Set Engine = CreateObject("JRO.JetEngine")
	Engine.compactdatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & datapath," provider=microsoft.jet.oledb.4.0;data source=" & strdbpath & "temp.mdb"
	fso.copyfile strdbpath & "temp.mdb",datapath
	fso.deletefile(strdbpath & "temp.mdb")
	set fso = nothing
	set Engine = nothing
	Response.Write "<script>alert('ϵͳ��ʾ:\n\nѹ�����ݿ�ɹ���');location='?';</script>"
	Response.end
else
	Response.Write "<script>alert('ϵͳ��ʾ:\n\n���ݿ�û���ҵ���');location='?';</script>"
	Response.end
end IF
End Sub
 %>
</body>
</html>
<% 
call rs_close()
call db_close()
call rc_close()

 %>





