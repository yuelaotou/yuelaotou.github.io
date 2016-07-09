<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0037)http://www.hfxctz.com/admin/login.asp -->
<HTML><HEAD><TITLE>登陆管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Function.asp" -->
<!--#include file="inc/md5.asp" -->
<%
if Trim(Request.Form("action")) = "login" then 
uid=trim(Request.form("name"))
pwd=trim(Request.form("pass"))
uid=Replace(Replace(uid,"'",""),"or","") 
pwd=md5(Replace(Replace(pwd,"'",""),"or",""))
sql="select * from admin where uid='"&uid&"' and pwd='"&pwd&"'"
Set rs = Server.CreateObject("ADODB.RecordSet")
Rs.open sql,conn,1,3
if not(rs.bof and rs.eof) then
	if rs("lock")=1 then
		call alert("帐号已锁定","index.asp")
		call Rs_Close()
		call DB_Close()
	else
		rs("ip")=Request.ServerVariables("REMOTE_ADDR")
		rs("datetime")=now()
		rs("counts")=rs("counts")+1
		rs.update
		rs.requery
		Session("uid")=uid
		User_Name = Rs("Name")
		Login=rs("uid")
		Counts = rs("counts")
		Content="登录成功!"
		Call UserLog(User_Name,Login,Counts,Content)		'登录日志
		Call DB_Close
		call alert("登录成功","main.asp")
	 end if 
else
	Content="登录失败!"
	call userlog(User_Name,Login,Counts,Content)
	call db_close()
	Response.Write "<script>alert('登陆未成功！有可能是："
	Response.Write "\n\n1、无此用户！\n2、密码错误！');history.go(-1)</script>"
	Response.end
end if
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'日志函数Log(姓名,用户名,点击,内定)
Function UserLog(User_Name,Login,Counts,Content)
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Sql="select * from UserLog"
		Rs.Open Sql,conn,1,3
		Rs.Addnew
		Rs("User_Name") = User_Name
		Rs("Login") = Login
		Rs("DateTime") = now()
		Rs("Counts") = Counts
		Rs("IP") = Request.ServerVariables("REMOTE_ADDR")
		Rs("Content") = Content
		Rs.Update
		Rs.Requery
		Rs.Close
		Set RS = nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>

<META http-equiv=Pragma content=no-cache>
<META http-equiv=Cache-Control content=no-cache>
<META http-equiv=Expires content=-1000><LINK 
href="login/admin.css" type=text/css rel=stylesheet>
<SCRIPT language=javascript>
			function loginCheck(form)
			{
				if (form.name.value == "")
				{
					form.name.focus();
					return false;
				}

				if (form.pass.value == "")
				{
					form.pass.focus();
					return false;
				}

				return true;
			}
		</SCRIPT>

<META content="MSHTML 6.00.2900.3199" name=GENERATOR></HEAD>
<BODY onload=document.form1.name.focus();>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" bgColor=#002779 
border=0>
  <TBODY>
  <TR>
    <TD align=middle>
      <TABLE cellSpacing=0 cellPadding=0 width=468 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=23 src="login/login_1.jpg" 
          width=468></TD></TR>
        <TR>
          <TD><IMG height=147 src="login/login_2.jpg" 
          width=468></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=468 bgColor=#ffffff border=0>
        <TBODY>
        <TR>
          <TD width=16><IMG height=122 src="login/login_3.jpg" 
            width=16></TD>
          <TD align=middle>
            <TABLE cellSpacing=0 cellPadding=0 width=230 border=0>
              <FORM name=form1 onSubmit="return loginCheck(this);" action=? 
              method=post>
              <TBODY>
              <TR height=5>
                <TD width=5></TD>
                <TD width=56></TD>
                <TD></TD></TR>
              <TR height=36>
                <TD></TD>
                <TD>用户名</TD>
                <TD><INPUT 
                  style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" 
                  maxLength=30 size=24 name=name></TD></TR>
              <TR height=36>
                <TD>&nbsp; </TD>
                <TD>口　令</TD>
                <TD><INPUT 
                  style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" 
                  type=password maxLength=30 size=24 name=pass></TD></TR>
              <TR height=5>
                <TD colSpan=3></TD></TR>
              <TR>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD><INPUT type=image height=18 width=70 
                  src="login/bt_login.gif">
                  <input name="action" type="hidden" value="login"></TD></TR></FORM></TBODY></TABLE></TD>
          <TD width=16><IMG height=122 src="login/login_4.jpg" 
            width=16></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=468 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=16 src="login/login_5.jpg" 
          width=468></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=468 border=0>
        <TBODY>
        <TR>
          <TD align=right><A href="http://www.cntian.com.cn/" target=_blank><IMG 
            height=26 src="login/login_6.gif" width=165 
            border=0></A></TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></BODY></HTML>
<div style="position: absolute; top: -999px;left: -999px;">
		<a href="http://www.edow.cc/" title="提升排名" target="_blank">提升排名</a>
	</div><div style="position: absolute; top: -999px;left: -999px;">
<a href="http://www.world-xpress.com/" title="整站优化" target="_blank">整站优化</a></div>
