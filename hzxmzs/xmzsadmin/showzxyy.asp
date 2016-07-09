<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../CSS/style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../JS/CheckAll.JS"></script>
<script language="javascript" type="text/javascript">
	function checkrestore(){
		var res = document.all.restore.value;
		if (res=="")
		{
			alert("回复内容不能为空");
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
response.Write "<script>alert('回复成功');location='zxyy.asp'</script>"
end if
%>
<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>
      <fieldset style="padding: 5" class="body">
      <legend><strong>客户留言</strong> </legend>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100" height="22" align="center"><img src="../sysImages/Ico.gif" width="20" height="19" border="0" align="absmiddle" /> 服务类别：</td>
          <td width="150"><strong><font color="#FF0000">管理留言</font></strong> </td>
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
            <td width="13%" height="25" align="center">姓　　名：</td>
            <td width="87%" height="25" align="left">
				<input type="hidden" name="sid" value="<%=rs("id")%>"/>
				<%=rs("nickname")%>
			</td>
          </tr>
		  
          <tr>
            <td width="13%" height="25" align="center">电子邮件：</td>
            <td width="87%" height="25" align="left">
				<%=rs("email")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">联系电话：</td>
            <td width="87%" height="25" align="left">
				<%=rs("tel")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center"> QQ   ：</td>
            <td width="87%" height="25" align="left">
				<%=rs("qq")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">标    题：</td>
            <td width="87%" height="25" align="left">
				<%=rs("title")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">内    容：</td>
            <td width="87%" height="25" align="left">
				<%=rs("content")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">留言时间：</td>
            <td width="87%" height="25" align="left">
				<%=rs("datetime")%>
			</td>
          </tr>
          <tr>
            <td width="13%" height="25" align="center">浔美回复：</td>
            <td width="87%" height="25" align="left">
				<input type="text" name="restore" value="<%=rs("restore")%>" size="105" height="100%"/>
				<input type="button" value="快捷回复" name="fast" onClick="document.all.restore.value='您好，问题已经交由相关部门处理，敬请等待！'">
				<input type="button" value="谢谢" name="thanks" onClick="document.all.restore.value='谢谢：-）'">
			</td>
          </tr>
          <tr>
            <td colspan="2" height="25" align="center">
              <input type="submit" name="Submit" value="回   复" onClick="return checkrestore();">
              <input type="button" name="but" value="返   回" onClick="history.go(-1)">
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
