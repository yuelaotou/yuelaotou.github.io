<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
<script language="javascript" type="text/javascript">
	function checkrestore(){
		var res = document.all.restore.value;
		if (res=="")
		{
			alert("�ظ����ݲ���Ϊ��");
			return false;
		}else{
			return true;
		}
	}
</script>
</head>
<%

if request("send")="restore" then
restore=request.Form("restore")
sid=request.Form("sid")

Set rs=server.CreateObject("ADODB.RecordSet")
    rs.open "select * from message where id = "&sid,conn,1,3
	rs("restore")=restore
	rs.update
	rs.close
	set rs=nothing  
response.Write "<script>alert('�ظ��ɹ�');location='zxyy.asp'</script>"
end if
%>
<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>
      <fieldset style="padding: 5" class="body">
      <legend><strong>�ͻ�����</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> �������</td>
          <td width="150"><strong><font color="#FF0000">��������</font></strong> </td>
          <td width="100" align="center">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      </fieldset>
    </td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="1" cellspacing="0" cellpadding="0">
        <form action="?send=restore" method="post" name="form1" id="form1">
          <%id=request("id")
		set rs=conn.execute("select * from message where id="&id)
		%>
		
          <tr>
            <td width="13%" height="25" align="center">�ա�������</td>
            <td width="87%" height="25" align="left">
				<input type="hidden" name="sid" value="<%=rs("id")%>"/>
				<%=rs("nickname")%>
			</td>
          </tr>
		  
          <tr>
            <td width="13%" height="25" align="center">�����ʼ���</td>
            <td width="87%" height="25" align="left">
				<%=rs("email")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">��ϵ�绰��</td>
            <td width="87%" height="25" align="left">
				<%=rs("tel")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center"> QQ   ��</td>
            <td width="87%" height="25" align="left">
				<%=rs("qq")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">��    �⣺</td>
            <td width="87%" height="25" align="left">
				<%=rs("title")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">��    �ݣ�</td>
            <td width="87%" height="25" align="left">
				<%=rs("content")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">����ʱ�䣺</td>
            <td width="87%" height="25" align="left">
				<%=rs("datetime")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">����ظ���</td>
            <td width="87%" height="25" align="left">
				<input type="text" name="restore" value="<%=rs("restore")%>" size="105" height="100%"/>
				<input type="button" value="��ݻظ�" name="fast" onClick="document.all.restore.value='���ã������Ѿ�������ز��Ŵ�������ȴ���'">
				<input type="button" value="лл" name="thanks" onClick="document.all.restore.value='лл��-��'">
			</td>
          </tr>
          <tr>
            <td colspan="2" height="25" align="center">
              <input type="submit" name="Submit" value="��   ��" onClick="return checkrestore();">
              <input type="button" name="but" value="��   ��" onClick="history.go(-1)">
            </td>
          </tr>
          <%
dim I
For I=1 to rs.pagesize
IF rs.eof THEN EXIT FOR 
rs.MoveNext
NEXT
%>
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
