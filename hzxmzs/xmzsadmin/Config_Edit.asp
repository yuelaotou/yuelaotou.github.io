<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<% 
Dim RS,Sql
IF Trim(Request.Form("Action"))="Edit" then
	Call Check_Right(3)	'��֤�û�Ȩ��
	Domain = Request.Form("Domain")
		IF Left(Domain,7)="http://" Then
	Else
		Domain = "http://" & Domain	
	End IF
	set RS=Server.createObject("adodb.recordset")
	Sql="select Top 1 * from Config"
	RS.open sql,conn,1,3
		'��վ
		RS("Site_Title")=Request.Form("Site_Title")
		'RS("Logo")=Request.Form("img")
		'RS("qq")=Request.Form("qq")
		'RS("Site_EndDate")=Request.Form("Site_EndDate")
		'RS("Site_Lock")=Request.Form("Site_Lock")
		RS("Keywords")=Request.Form("Keywords")
		RS("ICP")=Request.Form("ICP")
		RS("CopyRight")=Request.Form("CopyRight")
		RS("Address")=Request.Form("Address")
		RS("service")=Request.Form("service")
		RS("flash_index")=Request.Form("flash_index")
		RS("flash_about")=Request.Form("flash_about")
		RS("flash_culture")=Request.Form("flash_culture")
		RS("flash_case")=Request.Form("flash_case")
		RS("flash_service")=Request.Form("flash_service")
		RS("flash_site")=Request.Form("flash_site")
		RS("flash_knowledge")=Request.Form("flash_knowledge")
		RS("flash_employee")=Request.Form("flash_employee")
		RS("flash_message")=Request.Form("flash_message")
		RS("flash_contact")=Request.Form("flash_contact")
		RS("flash_aboutUs")=Request.Form("flash_aboutUs")
		RS("flash_shinei")=Request.Form("flash_shinei")
		RS("flash_dianmian")=Request.Form("flash_dianmian")
		RS("flash_bangongshi")=Request.Form("flash_bangongshi")
		RS("flash_bieshu")=Request.Form("flash_bieshu")
		RS("flash_xiezilou")=Request.Form("flash_xiezilou")
		RS("flash_xfbp")=Request.Form("flash_xfbp")
		
		'�ʼ�
		'RS("Email")=Request.Form("Email")
		'RS("Mail_Send")=Request.Form("Mail_Send")
		'RS("Mail_UserName")=Request.Form("Mail_UserName")
		'RS("Mail_Password")=Request.Form("Mail_Password")
		'����
		'RS("Domain")=Domain
		'RS("Domain_EndDate")=Request.Form("Domain_EndDate")
		'����
		'RS("Bank_Mid")=Request.Form("Bank_Mid")
		'RS("Bank_Key")=Request.Form("Bank_Key")
		'RS("Bank_SCompany")=Request.Form("Bank_SCompany")
		'RS("Bank_Url")=Request.Form("Bank_Url")
		RS.update
		RS.requery
		RS.close
		set RS=nothing
		Response.Write "<script>alert('�޸ĳɹ�!');location='?';</script>"
		Response.End
End IF

'IF Request.QueryString("Action")="DelImage" then
	'Call Check_Right(210)	'��֤�û�Ȩ��
	'Call Del_File("Config",1,"Logo","?")
'End IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set RS=Server.createObject("adodb.recordset")
Sql="select Top 1 * from Config"
RS.open sql,conn,1,1
 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/Alert.JS"></script>
<script language="JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" src="../JS/NewWindow.JS"></script>
<script language="JavaScript" src="../JS/NewImage.JS"></script>
<script language="JavaScript" src="../JS/Date.JS"></script>
<script language="JavaScript" type="text/JavaScript" src="Check.JS"></script>
<script language="JavaScript" type="text/javascript" src="../JS/zoomtextarea.Js"></script>
</head>

<body oncontextmenu="return false">

<form name="theForm" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1">
  <tr>
      <td bgcolor="#F0E7CF"> 
        <fieldset style="padding: 5">
        <legend><strong>ϵͳ����</strong></legend>
      
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <!--<fieldset style="padding: 5">
              <legend><strong>��վ</strong></legend> -->
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="100" height="22" align="center">��վ���ƣ�</td>
                  <td colspan="2">
<input name="Site_Title" type="text" id="Site_Title" size="50" value="<%= Rs("Site_Title") %>"></td>
                </tr>
                <tr> 
                  <!--<td height="22" align="center">Logo��ַ��</td>
                  <td>
<input name="Img" type="text" id="Img" size="20" value="<%= Rs("Logo") %>" readonly="0"><input name="Submit_Open" type="button" id="Submit_Open" value="�ϴ�" onClick="NewWindow('../Upfile/Upfile_Form.asp?UpfileType=Img_logo','UpFile','300','120','');return false;">
                  </td>
                  <td width="200" rowspan="3" align="center"> 
                    <table width="100%" border="0" cellspacing="1" cellpadding="0">
                      <tr> 
                        <td align="center"> 
                          <% IF RS("Logo")<>Empty Then %>
                          <img src="../../UploadFile/<%= RS("Logo") %>" name="logo" border="1" id="logo" onClick="newimg(this.src);" style="CURSOR: hand" width="88" height="31" alt="<%= RS("Site_Title") %>/UploadFile/<%= RS("Logo") %>"> 
                          <% Else %>
                          <img src="../Images/Link/Microsoft.gif" name="logo" width="88" height="31" id="logo">	
                          <% End IF %>
                        </td>
                      </tr>
                      <tr> 
                        <td height="20" align="center">
                          <% IF RS("Logo")<>Empty Then %>
                          <a href="#" onClick="{if(confirm('ȷ��ɾ����?')){location='?Action=DelImage';;return true;}return false;}">��ɾ��ͼƬ��</a> 
                          <% End IF %>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="22" align="center">վ��QQ��</td>
                  <td>
<input name="qq" type="text" id="qq" size="20" value="<%= Rs("qq") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">����ʱ�䣺</td>
                  <td><script language="JavaScript">arrowtag("Site_EndDate",'<%= Rs("Site_EndDate") %>');</script>&nbsp;&nbsp;
				  <%
	   DateTime=datediff("D",date(),Rs("Site_EndDate"))
	   IF left(DateTime,1)="-" then
	    %>
            <img src="../Images/Buttom/Ok.gif" alt="�ѵ���" width="16" height="16" border="0" align="absmiddle" > 
            <% Else %> <font color="#FF0000">ʣ�ࣺ&nbsp;<%= DateTime %>&nbsp;��</font> <% End IF %></td>
                </tr>
                <tr> 
                  <td height="22" align="center">������վ��</td>
                  <td colspan="2">&nbsp;��&nbsp;
<input type="radio" name="Site_Lock" value="1" <% If Rs("Site_Lock")=1 Then %>checked<% End If %>>
                    &nbsp;��&nbsp;
<input name="Site_Lock" type="radio" value="0" <% If Rs("Site_Lock")=0 Then %>checked<% End If %>></td> -->
                </tr>
                <tr> 
                  <td height="22" align="center">��&nbsp;��&nbsp;�֣�</td>
                  <td colspan="2"><textarea name="Keywords" cols="50" rows="5" id="Keywords"><%= Rs("Keywords") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�ٶ�ͳ�ƣ�</td>
                  <td colspan="2"><textarea name="CopyRight" cols="50" rows="5" id="CopyRight"><%= Rs("CopyRight") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��������</td>
                  <td colspan="2"><textarea name="Address" cols="50" rows="5" id="Address"><%= Rs("Address") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��Ȩ��Ϣ��</td>
                  <td colspan="2">
<textarea name="ICP" cols="50" rows="5" id="ICP"><%= Rs("ICP") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">������Ŀflash��ַ��</td>
                  <td colspan="2">
					<textarea name="service" cols="50" rows="5" id="service"><%= Rs("service") %></textarea>
        <!--<iframe ID="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=service&style=s_coolblue" frameborder="0" scrolling="no" width="450" HEIGHT="300"></iframe>-->
				</td>
                </tr>
                <tr> 
                  <td height="22" align="center">��ҳflash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_index" cols="50" rows="5" id="flash_index"><%= Rs("flash_index") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��˾���flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_about" cols="50" rows="5" id="flash_about"><%= Rs("flash_about") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��ҵ�Ļ�flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_culture" cols="50" rows="5" id="flash_culture"><%= Rs("flash_culture") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">���̰���flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_case" cols="50" rows="5" id="flash_case"><%= Rs("flash_case") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�ͻ�����flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_service" cols="50" rows="5" id="flash_service"><%= Rs("flash_service") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">ʩ���ֳ�flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_site" cols="50" rows="5" id="flash_site"><%= Rs("flash_site") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">װ��֪ʶflash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_knowledge" cols="50" rows="5" id="flash_knowledge"><%= Rs("flash_knowledge") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">������ʿflash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_employee" cols="50" rows="5" id="flash_employee"><%= Rs("flash_employee") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">������ѯflash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_message" cols="50" rows="5" id="flash_message"><%= Rs("flash_message") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��ϵ����flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_contact" cols="50" rows="5" id="flash_contact"><%= Rs("flash_contact") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�������flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_aboutUs" cols="50" rows="5" id="flash_aboutUs"><%= Rs("flash_aboutUs") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">����flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_shinei" cols="50" rows="5" id="flash_shinei"><%= Rs("flash_shinei") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">����flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_dianmian" cols="50" rows="5" id="flash_dianmian"><%= Rs("flash_dianmian") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�칫��flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_bangongshi" cols="50" rows="5" id="flash_bangongshi"><%= Rs("flash_bangongshi") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">����flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_bieshu" cols="50" rows="5" id="flash_bieshu"><%= Rs("flash_bieshu") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">д��¥flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_xiezilou" cols="50" rows="5" id="flash_xiezilou"><%= Rs("flash_xiezilou") %></textarea></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��������flash��ַ��</td>
                  <td colspan="2">
<textarea name="flash_xfbp" cols="50" rows="5" id="flash_xfbp"><%= Rs("flash_xfbp") %></textarea></td>
                </tr>
              </table>
             <!-- </fieldset> --></td>
          </tr>
        </table>
        <!--<table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>�ʼ�</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">��&nbsp;&nbsp;����</td>
                  <td>
<input name="Email" type="text" id="Email" size="40" value="<%= Rs("Email") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��������</td>
                  <td>
<input name="Mail_Send" type="text" id="Mail_Send" size="40" value="<%= Rs("Mail_Send") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">�û�����</td>
                  <td>
<input name="Mail_UserName" type="text" id="Mail_UserName" size="20" value="<%= Rs("Mail_UserName") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">��&nbsp;&nbsp;�룺</td>
                  <td>
<input name="Mail_Password" type="password" id="Mail_Password" size="20" value="<%= Rs("Mail_Password") %>"></td>
                </tr>
              </table>
              </fieldset></td>
          </tr>
        </table> -->
<!--        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr> 
            <td> <fieldset style="padding: 5">
              <legend><strong>����</strong></legend>
			  
			  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="150" height="22" align="center">������</td>
                  <td> <input name="Domain" type="text" id="Domain" size="40" value="<%= Rs("Domain") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">���ڣ�</td>
                  <td><script language="JavaScript">arrowtag("Domain_EndDate",'<%= Rs("Domain_EndDate") %>');</script>
                    &nbsp;&nbsp; <%
	   DateTime=datediff("D",date(),Rs("Domain_EndDate"))
	   IF left(DateTime,1)="-" then
	    %> <img src="../Images/Buttom/Ok.gif" alt="�ѵ���" width="16" height="16" border="0" align="absmiddle" > 
                    <% Else %> <font color="#FF0000">ʣ�ࣺ&nbsp;<%= DateTime %>&nbsp;��</font> <% End IF %> </td>
                </tr>
              </table>
              </fieldset>
			  
            </td>
          </tr>
        </table> -->
		<!--<table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td>		
  <fieldset style="padding: 5">
              <legend><strong>����</strong></legend>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="22" align="center">��&nbsp;��&nbsp;�ţ�</td>
                  <td> <input name="Bank_Mid" type="text" id="Bank_Mid" size="20" value="<%= Rs("Bank_Mid") %>"></td>
                </tr>
                <tr> 
                  <td height="22" align="center">MD5˽Կ��</td>
                  <td> <input name="Bank_Key" type="password" id="Bank_Key" size="20" value="<%= Rs("Bank_Key") %>"></td>
                </tr>
                <tr> 
                  <td width="150" height="22" align="center">�������ԣ�</td>
                  <td> <input name="Bank_SCompany" type="text" id="Bank_SCompany" size="40" value="<%= Rs("Bank_SCompany") %>"></td>
                </tr>
                <tr>
                  <td height="22" align="center">֧�������</td>
                  <td><input name="Bank_Url" type="text" id="Bank_Url" size="40" value="<%= Rs("Bank_Url") %>"></td>
                </tr>
              </table>
</fieldset>
</td>
          </tr>
        </table> -->		
		
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
            <td height="30" align="center"> 
<input type="submit" name="Submit_OK" value="�޸�(E)" accesskey="E">
              &nbsp; <input type="Reset" name="reset" value="ˢ��(R)" accesskey="R" onClick="javascript:location.reload();"> 
              <input name="Action" type="hidden" id="Action" value="Edit">

		    </td>
  </tr>
</table>

</fieldset>
      </td>
  </tr>
</table>
</form>


</body>
</html>
<%
call rs_close()
call db_close()
%>
