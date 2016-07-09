<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Include/Function.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Config.asp"-->
<%
response.write "正在生成首页<br>"
'==================================================服务项目开始===================================================
serviceShow = service_flash
'==================================================服务项目结束===================================================
'==================================================联系我们开始===================================================
set rs=conn.execute("select * from Page where SS_ID=268 and id=9")
if not rs.eof then
	contactUs = rs("Content")
else
	contactUs = "<div align=center><font color=red>还没有内容!</font></div>"
end if
rs.close
set rs=nothing
'==================================================联系我们结束===================================================
'=================================================首页图片秀开始====================================================
set rs=server.createobject("adodb.recordset")
rs.open "select top 6 * from Product where SS_ID=301 order by ID desc ",conn,1,1
pic="uploadfile/"&rs("Img")
link=""
text=trim(rs("Product_name"))
rs.movenext
do while not rs.eof
pic=pic&"|uploadfile/"&rs("Img")
link=link&"|"
text=text&"|"&trim(rs("Product_name"))
rs.movenext
loop
rs.close
set rs=Nothing
links = links & link
'=================================================首页图片秀结束====================================================
'==================================================公司简介开始===================================================
set rs=conn.execute("select * from Page where SS_ID=196 and ID=39")
if not rs.eof then
	CompanyIntro = rs("Content")
else
	CompanyIntro = "<div align=center><font color=red>公司简介</font></div>"
end if
rs.close
set rs = nothing
'==================================================公司简介结束===================================================

'==================================================大公司简介开始===================================================
set rs=conn.execute("select * from Page where SS_ID=196 and ID=1")
if not rs.eof then
	BCompanyIntro = rs("Content")
else
	BCompanyIntro = "<div align=center><font color=red>公司简介</font></div>"
end if
rs.close
set rs = nothing
'==================================================大公司简介结束===================================================

'==================================================企业文化开始===================================================
set rs=conn.execute("select * from Page where SS_ID=238 and ID=25")
if not rs.eof then
	Culture = rs("Content")
else
	Culture = "<div align=center><font color=red>企业文化</font></div>"
end if
rs.close
set rs = nothing
'==================================================企业文化结束===================================================

'==================================================文字精品案例开始===================================================
sql="select top 16 * from page where SS_ID=300 order by ID desc"
set rs=conn.execute(sql)
if not rs.eof then
caseword = caseword & chr(13)&chr(10)
do while not rs.eof 
caseword = caseword & "			<li><a href=""#"">" & rs("Title") & "</a></li>" &chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================文字精品案例结束===================================================

'==================================================图片精品案例开始===================================================
Set rs=server.CreateObject("ADODB.RecordSet")
sql = "select top 24 * from Product where Lock=0 and SS_ID<>301 order by ID desc"
set rs=conn.execute(sql)
if not rs.eof then
casepic = casepic & chr(13)&chr(10)
do while not rs.eof 
casepic = casepic & "		<li><a href=""uploadfile/" & rs("Img") & """ title=""" & rs("Product_name") & """><img src=""uploadfile/" & rs("SmallImg") & """/></a></li>" &chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================图片精品案例结束===================================================

'==================================================装修知识开始===================================================
sql="select top 10 news.* from [Type],news where Type.Parent_ID=278 and Type.ID=news.SS_ID order by news.id desc"
set rs=conn.execute(sql)
if not rs.eof then
knowledge = knowledge & chr(13)&chr(10)
do while not rs.eof 
	knowledgeTime = FormatDateTime(rs("DateTime"),2)
	knowledgeTime = Mid(knowledgeTime,6,Len(knowledgeTime)-5)
	knowledgeTime = Replace(knowledgeTime,"-","/")
	knowledge = knowledge&"<li><span>"&knowledgeTime&"</span><a href=""knowledge_list_"&rs("id")&".htm"" target=""_blank"">"&rs("Title")&"</a></li>"&chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================装修知识结束===================================================

'==================================================设计团队开始===================================================
sql="select top 4 * from Team where lock = 0 order by id desc"
set rs=conn.execute(sql)
if not rs.eof then
team = team & chr(13)&chr(10)
do while not rs.eof
team_smallImg=replace(rs("Img"),".","a.")
team = team & "			<li>" &chr(13)&chr(10)
team = team & "				<a href=""team_list_" & rs("id") &".htm"" target=""_blank""><img src=""UploadFile/" & team_smallImg  & """/></a>" &chr(13)&chr(10)
team = team & "				<span>" & rs("name") & "</br>" & rs("zhiwei") & "</br>" & rs("zili") & "</span>" &chr(13)&chr(10)
team = team & "			</li>" &chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================设计团队结束===================================================

'==================================================首页施工现场4张图开始===================================================
sql="select top 4 * from Site where lock = 0 order by id desc"
set rs=conn.execute(sql)
if not rs.eof then
site = site & chr(13)&chr(10)
do while not rs.eof
site_smallImg=replace(rs("Img"),".","a.")
site = site & "			<li>" &chr(13)&chr(10)
site = site & "				<a href=""site_list_" & rs("id") &".htm"" target=""_blank"" title=""" &rs("name")& """><img src=""UploadFile/" & site_smallImg  & """/></a>" &chr(13)&chr(10)
site = site & "				<span>" & rs("name") & "</br>" & rs("area") & "</br>" & rs("price") & "</br>" & rs("designer") & "</span>" &chr(13)&chr(10)
site = site & "			</li>" &chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================首页施工现场4张图结束===================================================
'==================================================工装步骤开始===================================================
set rs=conn.execute("select * from page where SS_ID=298")
if not rs.eof then
fitStep = rs("Content")
end if
rs.close
set rs=nothing
'==================================================工装步骤结束===================================================

'==================================================案例分类开始===================================================
set rs=conn.execute("select id,Title from type where Parent_ID=199 order by ID")
if not rs.eof then
do while not rs.eof 
caseType = caseType & "			<li><a href=""case_list_"&rs("id")&".htm"">" & rs("Title") & "</a></li>" &chr(13)&chr(10)
rs.movenext
loop
end If
rs.close
set rs=nothing
'==================================================案例分类结束===================================================

'==================================================设计团队人员列表开始===================================================
set rs=conn.execute("select id,zhiwei,name from team order by ID")
if not rs.eof then
do while not rs.eof 
teamType = teamType & "			<li><a href=""team_list_" & rs("id") & ".htm"">" & rs("zhiwei") & "-" & rs("name") &"</a></li>" &chr(13)&chr(10)
rs.movenext
loop
end If
rs.close
set rs=nothing
'==================================================设计团队人员列表结束===================================================

'==================================================施工现场列表开始===================================================
set rs=conn.execute("select id,designer,name from Site order by ID")
if not rs.eof then
do while not rs.eof 
siteType = siteType & "			<li><a href=""site_list_" & rs("id") & ".htm"">" & rs("name") & " - " & rs("designer") &"</a></li>" &chr(13)&chr(10)
rs.movenext
loop
end If
rs.close
set rs=nothing
'==================================================施工现场列表结束===================================================

'==================================================友情链接开始===================================================
sql="select top 10 * from Link where Lock = 0 order by id"
set rs=conn.execute(sql)
if not rs.eof then
FriendLink = FriendLink & chr(13)&chr(10)
do while not rs.eof 
FriendLink = FriendLink & "      <li> <a href=""" & rs("Url") & """ target=""_blank""> <img src=""UploadFile/" & rs("Img") & """ width=""88"" height=""31"" border=""0"" align=""left"" /> </a> </li>" &chr(13)&chr(10)
rs.movenext
loop
end if
rs.close
set rs=nothing
'==================================================友情链接结束===================================================

Set fso = Server.CreateObject("Scripting.FileSystemObject")
templatefileName="../template/index.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_index$",flash_index)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$pics$",pic)
str=replace(str,"$links$",links)
str=replace(str,"$texts$",text)
str=replace(str,"$CompanyIntro$",CompanyIntro)
str=replace(str,"$caseword$",caseword)
str=replace(str,"$casepic$",casepic)
str=replace(str,"$knowledge$",knowledge)
str=replace(str,"$team$",team)
str=replace(str,"$site$",site)
str=replace(str,"$fitStep$",fitStep)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../index.htm"
'Set Files = Fso.GetFile(server.MapPath(createfileName))
'Files.Delete(True)
'response.write "首页删除成功<br><br><br>"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
set str=Nothing
response.write "首页生成成功<br><br><br>"
'============================================
response.write "正在生成公司简介<br>"
templatefileName="../template/about.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_about$",flash_about)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$BCompanyIntro$",BCompanyIntro)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../about.htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "公司简介生成成功<br><br><br>"
'============================================
response.write "正在生成企业文化<br>"
templatefileName="../template/culture.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_culture$",flash_culture)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$Culture$",Culture)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../culture.htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "企业文化生成成功<br><br><br>"
'================================================================================================================


'==================================================招贤纳士分页开始===================================================
sql="select * from job order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=2
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
employeeList = ""
sql="select * from job order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=2
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

employeeList = employeeList & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
employeeList = employeeList & "      <table width=""700"" style=""border:1px solid #CCCCCC"" cellpadding=""0"" cellspacing=""1"">" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F7F7F7"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td width=""125"">岗位名称：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td width=""569"">"&rs("Title")&"</td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F0F0F0"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td>招聘人数：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td>"&rs("Number")&"</td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F7F7F7"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td>工资待遇：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td>"&rs("Explain")&"</td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F0F0F0"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td>结束日期：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td>"&rs("Datetime")&"</td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F7F7F7"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td>工作地点：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td>"&rs("Address")&"</td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "        <tr bgcolor=""#F0F0F0"">" &chr(13)&chr(10)
employeeList = employeeList & "          <td>职位描述/要求：</td>" &chr(13)&chr(10)
employeeList = employeeList & "          <td> "&rs("Content")&" </td>" &chr(13)&chr(10)
employeeList = employeeList & "        </tr>" &chr(13)&chr(10)
employeeList = employeeList & "      </table>" &chr(13)&chr(10)
If howmanyrecs/2 = 0 Then
employeeList = employeeList & "      <br>" &chr(13)&chr(10)
End if
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('employee','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('employee','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('employee','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('employee','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('employee','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('employee','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/employee.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_employee$",flash_employee)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$employee$",employeeList)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../employee.htm"
Else
createfileName="../employee_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成招贤纳士的第"&j&"个页面!" & createfileName & "<br>"

next
response.write "招贤纳士生成成功<br><br><br>"
'==================================================招贤纳士分页结束===================================================
'==================================================留言咨询分页开始===================================================
sql="select * from message where status=1 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=2
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
messageList = ""
sql="select * from message where status=1 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=2
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

messageList = messageList & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize

messageList = messageList & "    <table width=""700"" border=""1"">" &chr(13)&chr(10)
messageList = messageList & "      <tr>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""85"">昵称：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""117"">"&rs("nickname")&"</td>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""64"">电话：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""170"">"&rs("tel")&"</td>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""64"">时间：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td width=""170"">"&rs("datetime")&"</td>" &chr(13)&chr(10)
messageList = messageList & "      </tr>" &chr(13)&chr(10)
messageList = messageList & "      <tr>" &chr(13)&chr(10)
messageList = messageList & "        <td><img src=""img/qq.gif"" style=""margin-left:22px;width:18px;height:17px;""/></td>" &chr(13)&chr(10)
messageList = messageList & "        <td>"&rs("qq")&"</td>" &chr(13)&chr(10)
messageList = messageList & "        <td><img src=""img/e-mail.gif"" style=""margin-left:22px;width:18px;height:17px;""/></td>" &chr(13)&chr(10)
messageList = messageList & "        <td>"&rs("email")&"</td>" &chr(13)&chr(10)
messageList = messageList & "        <td>&nbsp;</td>" &chr(13)&chr(10)
messageList = messageList & "        <td>&nbsp;</td>" &chr(13)&chr(10)
messageList = messageList & "      </tr>" &chr(13)&chr(10)
messageList = messageList & "      <tr>" &chr(13)&chr(10)
messageList = messageList & "        <td>留言标题：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td colspan=""5"">"&rs("title")&"</td>" &chr(13)&chr(10)
messageList = messageList & "      </tr>" &chr(13)&chr(10)
messageList = messageList & "      <tr>" &chr(13)&chr(10)
messageList = messageList & "        <td>留言内容：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td colspan=""5"">"&replace(rs("content"),chr(13)&chr(10),"<br>")&"</td>" &chr(13)&chr(10)
messageList = messageList & "      </tr>" &chr(13)&chr(10)
messageList = messageList & "      <tr>" &chr(13)&chr(10)
messageList = messageList & "        <td>浔美回复：</td>" &chr(13)&chr(10)
messageList = messageList & "        <td colspan=""5"">"&rs("restore")&"</td>" &chr(13)&chr(10)
messageList = messageList & "      </tr>" &chr(13)&chr(10)
messageList = messageList & "    </table><br>" &chr(13)&chr(10)

rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('message','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('message','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('message','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('message','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('message','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('message','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/message.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_message$",flash_message)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$message$",messageList)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../message.htm"
Else
createfileName="../message_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成留言咨询的第"&j&"个页面!" & createfileName & "<br>"

Next
If totalpage = 0 Then
	templatefileName="../template/message.asp"
	Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
	str=templates.readall()'读出结果，赋值给str
	set templates=Nothing

	str=replace(str,"$Keywords$",Keywords)
	str=replace(str,"$CopyRight$",CopyRight)
	str=replace(str,"$ICP$",ICP)
	str=replace(str,"$Site_Title$",Site_Title)
	str=replace(str,"$Address$",Address)
	str=replace(str,"$serviceShow$",serviceShow)
	str=replace(str,"$contactUs$",contactUs)
	str=replace(str,"$message$","")
	str=replace(str,"$pagination$","")
	str=replace(str,"$FriendLink$","")

	createfileName="../message.htm"
	Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
	htmlFile.write(str)
	set htmlFile=nothing
	response.write "正在生成留言咨询!" & createfileName & "<br>"
End If
response.write "留言咨询生成成功<br><br><br>"
'==================================================留言咨询分页结束===================================================
'============================================
response.write "正在生成联系我们<br>"
templatefileName="../template/contact.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_contact$",flash_contact)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$contactAddr$","此功能暂未开放")
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../contact.htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "联系我们生成成功<br><br><br>"
'================================================================================================================

'==================================================案例分页开始===================================================
sql="select * from Product where SS_ID<>301 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=9
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
caselist = ""
sql="select * from Product where SS_ID<>301 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=9
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

caselist = caselist & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
caselist = caselist & "		<li>" &chr(13)&chr(10)
caselist = caselist & "			<a class=""caseList_pic"" href=""UploadFile/" & rs("Img") & """ target=""_blank"" title=""" & rs("Content") & """>" &chr(13)&chr(10)
caselist = caselist & "				<img src=""UploadFile/" & rs("SmallImg") & """>" &chr(13)&chr(10)
caselist = caselist & "			</a>" &chr(13)&chr(10)
'caselist = caselist & "			<a class=""caseList_a"" href=""UploadFile/" & rs("Img") & """ target=""_blank"">" & rs("Product_Name") & "</a>" &chr(13)&chr(10)
caselist = caselist & "		</li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('case','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('case','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('case','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('case','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('case','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('case','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/case.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_case$",flash_case)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$caseType$",caseType)
str=replace(str,"$caselist$",caselist)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../case.htm"
Else
createfileName="../case_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成工程案例的第"&j&"个页面!" & createfileName & "<br>"


next
response.write "工程案例生成成功<br><br><br>"
'==================================================案例分页结束===================================================



'==================================================案例分栏分页开始===================================================
set rsd=conn.execute("select id,Title from type where Parent_ID=199 order by ID")
if not rsd.eof then
do while not rsd.eof 

sql="select * from Product where SS_ID = "&rsd("id")&" order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=9
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
caselist = ""
sql="select * from Product where SS_ID = "&rsd("id")&" order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=9
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

caselist = caselist & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
caselist = caselist & "		<li>" &chr(13)&chr(10)
caselist = caselist & "			<a class=""caseList_pic"" href=""UploadFile/" & rs("Img") & """ target=""_blank"" title=""" & rs("Content") & """>" &chr(13)&chr(10)
caselist = caselist & "				<img src=""UploadFile/" & rs("SmallImg") & """>" &chr(13)&chr(10)
caselist = caselist & "			</a>" &chr(13)&chr(10)
caselist = caselist & "			<a class=""caseList_a"" href=""UploadFile/" & rs("Img") & """ target=""_blank"">" & rs("Product_Name") & "</a>" &chr(13)&chr(10)
caselist = caselist & "		</li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('case_list_" & rsd("id") & "','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('case_list_" & rsd("id") & "','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('case_list_" & rsd("id") & "','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('case_list_" & rsd("id") & "','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('case_list_" & rsd("id") & "','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('case_list_" & rsd("id") & "','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/case.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_case$",flash_case)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$caseType$",caseType)
str=replace(str,"$caselist$",caselist)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../case_list_" & rsd("id")&".htm"
Else
createfileName="../case_list_" & rsd("id") & "_" & j-1 & ".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成工程案例分栏的第"&j&"个页面!" & createfileName & "<br>"


Next

If totalpage = 0 Then
	templatefileName="../template/case.asp"
	Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
	str=templates.readall()'读出结果，赋值给str
	set templates=Nothing

	str=replace(str,"$Keywords$",Keywords)
	str=replace(str,"$CopyRight$",CopyRight)
	str=replace(str,"$ICP$",ICP)
	str=replace(str,"$Site_Title$",Site_Title)
	str=replace(str,"$Address$",Address)
	str=replace(str,"$flash_case$",flash_case)
	str=replace(str,"$serviceShow$",serviceShow)
	str=replace(str,"$contactUs$",contactUs)
	str=replace(str,"$caseType$",caseType)
	str=replace(str,"$caselist$","当前无记录")
	str=replace(str,"$pagination$","")
	tr=replace(str,"$FriendLink$",FriendLink)

	createfileName="../case_list_" & rsd("id")&".htm"
	Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
	htmlFile.write(str)
	set htmlFile=nothing
	response.write "正在生成工程案例分栏的第0个页面!" & createfileName & "<br>"
End If

rsd.movenext
Loop
end If
rsd.close
set rsd=Nothing
response.write "工程案例生成成功<br><br><br>"
'==================================================案例分栏分页结束===================================================


'==================================================施工现场分页开始===================================================
sql="select * from Site order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=3
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
sitelist = ""
sql="select * from Site order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=3
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

sitelist = sitelist & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
status_site = ""
If rs("status") = 1 Then
	status_site = "施工前期"
End If 
If rs("status") = 2 Then
	status_site = "施工中期"
End If 
If rs("status") = 3 Then
	status_site = "施工后期"
End If 
If rs("status") = 4 Then
	status_site = "工程结束"
End If 
site_smallImg=replace(rs("Img"),".","a.")
sitelist = sitelist & "        <li>" &chr(13)&chr(10)
sitelist = sitelist & "          <div class=""teamList_content_left""><a href=""site_list_" & rs("id") & ".htm""><img src=""UploadFile/"& site_smallImg &"""/></a></div>" &chr(13)&chr(10)
sitelist = sitelist & "          <div class=""teamList_content_right"">" &chr(13)&chr(10)
sitelist = sitelist & "            <table width=""100%"" height=""125"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""font-size:12px;"">" &chr(13)&chr(10)
sitelist = sitelist & "              <tr>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">工程名称：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""85%"" colspan=""3"">"& rs("name") &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "              </tr>" &chr(13)&chr(10)
sitelist = sitelist & "              <tr>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">建筑面积：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""35%"">"& rs("area") &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">工程造价：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""35%"">"& rs("price") &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "              </tr>" &chr(13)&chr(10)
sitelist = sitelist & "              <tr>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">工程进度：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""85%"" colspan=""3"">"& status_site &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "              </tr>" &chr(13)&chr(10)
sitelist = sitelist & "              <tr>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">设 计 师：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""85%"" colspan=""3"">"& rs("designer") &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "              </tr>" &chr(13)&chr(10)
sitelist = sitelist & "              <tr>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">设计说明：</font></td>" &chr(13)&chr(10)
sitelist = sitelist & "                <td align=""left"" width=""85%"" colspan=""3"">"& rs("intro") &"</td>" &chr(13)&chr(10)
sitelist = sitelist & "              </tr>" &chr(13)&chr(10)
sitelist = sitelist & "            </table>" &chr(13)&chr(10)
sitelist = sitelist & "          </div>" &chr(13)&chr(10)
sitelist = sitelist & "        </li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('site','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('site','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('site','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('site','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('site','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('site','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/site.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_site$",flash_site)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$siteType$",siteType)
str=replace(str,"$sitelist$",sitelist)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../site.htm"
Else
createfileName="../site_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成施工现场的第"&j&"个页面!" & createfileName & "<br>"

next
response.write "施工现场生成成功<br><br><br>"
'==================================================施工现场分页结束===================================================



'==================================================设计团队分页开始===================================================
sql="select * from Team order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=3
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
teamlist = ""
sql="select * from Team order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=3
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

teamlist = teamlist & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize

team_smallImg=replace(rs("Img"),".","a.")
teamlist = teamlist & "        <li>" &chr(13)&chr(10)
teamlist = teamlist & "          <div class=""teamList_content_left""><a href=""team_list_" & rs("id") & ".htm""><img src=""UploadFile/"& team_smallImg &"""/></a></div>" &chr(13)&chr(10)
teamlist = teamlist & "          <div class=""teamList_content_right"">" &chr(13)&chr(10)
teamlist = teamlist & "            <table width=""100%"" height=""140"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""font-size:12px;"">" &chr(13)&chr(10)
teamlist = teamlist & "              <tr>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">设计师名称</font>：</td>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""80%"">"& rs("name") &"</td>" &chr(13)&chr(10)
teamlist = teamlist & "              </tr>" &chr(13)&chr(10)
teamlist = teamlist & "              <tr>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">从业资历</font>：</td>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""80%"">"& rs("zili") &"</td>" &chr(13)&chr(10)
teamlist = teamlist & "              </tr>" &chr(13)&chr(10)
teamlist = teamlist & "              <tr>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">简介</font>：</td>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""80%"">"& rs("jianjie") &"</td>" &chr(13)&chr(10)
teamlist = teamlist & "              </tr>" &chr(13)&chr(10)
teamlist = teamlist & "              <tr>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""15%"">&nbsp;<font color=""#763426"">服务理念</font>：</td>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"" width=""80%"">"& rs("linian") &"</td>" &chr(13)&chr(10)
teamlist = teamlist & "              </tr>" &chr(13)&chr(10)
teamlist = teamlist & "              <tr>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"">&nbsp;<font color=""#763426"">主要作品</font>：</td>" &chr(13)&chr(10)
teamlist = teamlist & "                <td align=""left"">"& rs("zuopin") &"</td>" &chr(13)&chr(10)
teamlist = teamlist & "              </tr>" &chr(13)&chr(10)
teamlist = teamlist & "            </table>" &chr(13)&chr(10)
teamlist = teamlist & "          </div>" &chr(13)&chr(10)
teamlist = teamlist & "        </li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('team','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('team','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('team','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('team','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('team','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('team','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/team.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_team$",flash_team)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$teamType$",teamType)
str=replace(str,"$teamlist$",teamlist)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../team.htm"
Else
createfileName="../team_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成设计团队的第"&j&"个页面!" & createfileName & "<br>"


next
response.write "设计团队生成成功<br><br><br>"
'==================================================设计团队分页结束===================================================

'==================================================客户服务分页开始===================================================
sql="select * from page where SS_ID=284 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=12
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
service = ""
sql="select * from page where SS_ID=284 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=12
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

service = service & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
service = service & "            <li><span>" & rs("Counts") & "</span><span>" & rs("DateTime") & "</span><a href=""service_list_" & rs("id") & ".htm"" target=""_blank"">" & rs("Title") & "</a></li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('service','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('service','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('service','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('service','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('service','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('service','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/service.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_service$",flash_service)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$service$",service)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../service.htm"
Else
createfileName="../service_"&j-1&".htm"
End if
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成客户服务的第"&j&"个页面!" & createfileName & "<br>"


next
response.write "客户服务生成成功<br><br><br>"
'==================================================客户服务分页结束===================================================

'==================================================装修知识分页开始===================================================
sql="select * from news where SS_ID=279 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1
rowCount = rs.RecordCount
rs.pagesize=12
totalpage=rs.pagecount
rs.close
set rs=nothing

for j=1 to totalpage
knowledgelist = ""
sql="select * from news where SS_ID=279 order by id desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,conn,1,1

whichpage=j 
rs.pagesize=12
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0

knowledgelist = knowledgelist & chr(13)&chr(10)
do while not rs.eof and howmanyrecs<rs.pagesize
knowledgelist = knowledgelist & "            <li><span>" & rs("Counts") & "</span><span>" & rs("DateTime") & "</span><span>" & rs("author") & "</span><a href=""knowledge_list_" & rs("id") & ".htm"" target=""_blank"">" & rs("Title") & "</a></li>" &chr(13)&chr(10)
rs.movenext
howmanyrecs = howmanyrecs + 1
loop
rs.close
set rs=nothing

'分页部分
pagination = "" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalAnnal"">总记录数：<i>" & rowCount & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""totalPage"">总页数：<i>" & totalpage & "</i></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""currentPage"">当前页：<b>" & j & "</b></li>" &chr(13)&chr(10)
if whichpage=1 then
pagination = pagination & "		<li class=""firstPage currentState""><a href=""javascript:void(0);"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage currentState""><a href=""javascript:void(0);"">前一页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""firstPage""><a href=""javascript:goto('knowledge','0');"" title=""首页"">首页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""previousPage""><a href=""javascript:goto('knowledge','"&j-2&"');"" title=""前一页"">前一页</a></li>" &chr(13)&chr(10)
end if
pagination = pagination & "		<li>" &chr(13)&chr(10)
pagination = pagination & "			<ol>" &chr(13)&chr(10)
If j-2 > 0 then
	for counter=j-2 to j+2
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('knowledge','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
Else
	for counter=1 to 5
		If counter <= totalpage then
			if counter = j then
	pagination = pagination & "				<li class=""currentState"" title=""当前页""><a href=""javascript:void(0);"">" & j & "</a></li>" &chr(13)&chr(10)
			Else
	pagination = pagination & "				<li><a title=""转到第" & counter & "页"" href=""javascript:goto('knowledge','"&counter-1&"');"">"&counter&"</a></li>" &chr(13)&chr(10)
			end If
		End if
	Next
End if
pagination = pagination & "			</ol>" &chr(13)&chr(10)
pagination = pagination & "		</li>" &chr(13)&chr(10)
if (whichpage>=totalpage) then
pagination = pagination & "		<li class=""nextPage currentState""><a href=""javascript:void(0);"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage currentState""><a href=""javascript:void(0);"">尾页</a></li>" &chr(13)&chr(10)
else
pagination = pagination & "		<li class=""nextPage""><a href=""javascript:goto('knowledge','"&j&"');"" title=""后一页"">后一页</a></li>" &chr(13)&chr(10)
pagination = pagination & "		<li class=""lastPage""><a href=""javascript:goto('knowledge','"&totalpage-1&"');"" title=""尾页"">尾页</a></li>" &chr(13)&chr(10)
end if
whichpage = whichpage+1

templatefileName="../template/knowledge.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_knowledge$",flash_knowledge)
str=replace(str,"$serviceShow$",serviceShow)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$knowledgelist$",knowledgelist)
str=replace(str,"$pagination$",pagination)
str=replace(str,"$FriendLink$",FriendLink)

If j = 1 Then
createfileName="../knowledge.htm"
Else
createfileName="../knowledge_"&j-1&".htm"
End if

Set Files = Fso.GetFile(server.MapPath(createfileName))
Files.Delete(True)
response.Write "已经删除======"&createfileName& "<br>"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成装修知识的第"&j&"个页面!" & createfileName & "<br>"


next
response.write "装修知识生成成功<br><br><br>"
'==================================================装修知识分页结束===================================================
conn.close
set conn=nothing
set fso=nothing
%>