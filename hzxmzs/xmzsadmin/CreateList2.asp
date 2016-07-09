<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Include/Function.asp" -->
<!--#include file="Conn.asp" -->
<!--#include file="../Include/Config.asp"-->
<%
response.write "正在生成分栏详细信息...<br>"
Set fso = Server.CreateObject("Scripting.FileSystemObject")
'==================================================服务项目开始===================================================
serviceShow = service_flash
'==================================================服务项目结束===================================================
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

'==================================================施工现场详细信息开始===================================================
set rs=conn.execute("select * from Site")
if not rs.eof Then
do while not rs.eof 
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
sitelist = "" & chr(13)&chr(10)
siteEmbed = "" & chr(13)&chr(10)
site_smallImg=replace(rs("Img"),".","a.")
sitelist = sitelist & "        <li>" &chr(13)&chr(10)
sitelist = sitelist & "          <div class=""teamList_content_left""><a href=""UploadFile/"& rs("Img") &"""><img src=""UploadFile/"& site_smallImg &"""/></a></div>" &chr(13)&chr(10)
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

siteEmbed = siteEmbed & rs("content") &chr(13)&chr(10)

templatefileName="../template/site_list.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_site$",flash_site)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$siteType$",siteType)
str=replace(str,"$sitelist$",sitelist)
str=replace(str,"$siteEmbed$",siteEmbed)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../site_list_" & rs("id") & ".htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成施工现场详细信息..."&createfileName&"<br>"

rs.movenext
loop
end If
rs.close
set rs=Nothing
response.write "施工现场详细信息生成成功<br>"
'==================================================施工现场详细信息结束===================================================

'==================================================设计团队人员详细信息开始===================================================
set rs=conn.execute("select * from Team")
if not rs.eof Then
do while not rs.eof 
teamlist = "" & chr(13)&chr(10)
teamEmbed = "" & chr(13)&chr(10)
team_smallImg=replace(rs("Img"),".","a.")
teamlist = teamlist & "        <li>" &chr(13)&chr(10)
teamlist = teamlist & "          <div class=""teamList_content_left""><a href=""UploadFile/"& rs("Img") &"""><img src=""UploadFile/"& team_smallImg &"""/></a></div>" &chr(13)&chr(10)
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

teamEmbed = teamEmbed & rs("embed") &chr(13)&chr(10)

templatefileName="../template/team_list.asp"
Set templates = fso.OpenTextFile(server.MapPath(templatefileName))
str=templates.readall()'读出结果，赋值给str
set templates=Nothing

str=replace(str,"$Keywords$",Keywords)
str=replace(str,"$CopyRight$",CopyRight)
str=replace(str,"$ICP$",ICP)
str=replace(str,"$Site_Title$",Site_Title)
str=replace(str,"$Address$",Address)
str=replace(str,"$flash_team$",flash_team)
str=replace(str,"$contactUs$",contactUs)
str=replace(str,"$teamType$",teamType)
str=replace(str,"$teamlist$",teamlist)
str=replace(str,"$teamEmbed$",teamEmbed)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../team_list_" & rs("id") & ".htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成设计团队..."&createfileName&"<br>"

rs.movenext
loop
end If
rs.close
set rs=Nothing
response.write "设计团队人员详细信息生成成功<br>"
'==================================================设计团队人员详细信息结束===================================================


'==================================================客户服务详细信息开始===================================================
set rs=conn.execute("select * from Page where SS_ID = 284")
if not rs.eof Then
do while not rs.eof 
servicelist = "" & chr(13)&chr(10)
servicelist = servicelist & "			<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
servicelist = servicelist & "			  <tr>" &chr(13)&chr(10)
servicelist = servicelist & "				<td>" &chr(13)&chr(10)
servicelist = servicelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
servicelist = servicelist & "					<tr>" &chr(13)&chr(10)
servicelist = servicelist & "					  <td height=""30"" align=""center"" valign=""top"" style=""font-size:14px; font-weight:bold"">" & rs("Title") & "</td>" &chr(13)&chr(10)
servicelist = servicelist & "					</tr>" &chr(13)&chr(10)
servicelist = servicelist & "				  </table>" &chr(13)&chr(10)
servicelist = servicelist & "				  <table width=""100%"" height=""34"" border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/news_11.gif"">" &chr(13)&chr(10)
servicelist = servicelist & "					<tr>" &chr(13)&chr(10)
servicelist = servicelist & "					  <td align=""center"">信息来源：浔美装饰　　发布时间：" & rs("DateTime") & "　　点击次数：<span class=""red"">" & rs("Counts") & "</span></td>" &chr(13)&chr(10)
servicelist = servicelist & "					</tr>" &chr(13)&chr(10)
servicelist = servicelist & "				  </table>" &chr(13)&chr(10)
servicelist = servicelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
servicelist = servicelist & "					<tr>" &chr(13)&chr(10)
servicelist = servicelist & "					  <td style=""padding:15px; line-height:24px"">" &chr(13)&chr(10)
servicelist = servicelist & "						<!--内容-->" &chr(13)&chr(10)
servicelist = servicelist & "						" & rs("Content") & " </td>" &chr(13)&chr(10)
servicelist = servicelist & "					</tr>" &chr(13)&chr(10)
servicelist = servicelist & "				  </table>" &chr(13)&chr(10)
servicelist = servicelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
servicelist = servicelist & "					<tr>" &chr(13)&chr(10)
servicelist = servicelist & "					  <td align=""center""></td>" &chr(13)&chr(10)
servicelist = servicelist & "					</tr>" &chr(13)&chr(10)
servicelist = servicelist & "					<tr>" &chr(13)&chr(10)
servicelist = servicelist & "					  <td height=""30"" align=""right""><a href=""javascript:window.close();"" class=""red01""><strong>[关闭本页]</strong></a>&nbsp;&nbsp;</td>" &chr(13)&chr(10)
servicelist = servicelist & "					</tr>" &chr(13)&chr(10)
servicelist = servicelist & "				  </table>" &chr(13)&chr(10)
servicelist = servicelist & "				</td>" &chr(13)&chr(10)
servicelist = servicelist & "			  </tr>" &chr(13)&chr(10)
servicelist = servicelist & "			</table>" &chr(13)&chr(10)


templatefileName="../template/service_list.asp"
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
str=replace(str,"$servicelist$",servicelist)
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../service_list_" & rs("id") & ".htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成客户服务..."&createfileName&"<br>"

rs.movenext
loop
end If
rs.close
set rs=Nothing
response.write "客户服务详细信息生成成功<br>"
'==================================================客户服务详细信息结束===================================================


'==================================================装修知识详细信息开始===================================================
set rs=conn.execute("select news.* from [Type],news where Type.Parent_ID=278 and Type.ID=news.SS_ID order by news.id desc")
if not rs.eof Then
do while not rs.eof 
knowledgelist = "" & chr(13)&chr(10)
knowledgelist = knowledgelist & "			<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "			  <tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				<td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					<tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					  <td height=""30"" align=""center"" valign=""top"" style=""font-size:14px; font-weight:bold"">" & rs("Title") & "</td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					</tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  </table>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  <table width=""100%"" height=""34"" border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/news_11.gif"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					<tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					  <td align=""center"">信息来源：浔美装饰　　发布时间：" & rs("DateTime") & "　　点击次数：<span class=""red"">" & rs("Counts") & "</span></td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					</tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  </table>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					<tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					  <td style=""padding:15px; line-height:24px"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "						<!--内容-->" &chr(13)&chr(10)
knowledgelist = knowledgelist & "						" & rs("Content") & " </td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					</tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  </table>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					<tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					  <td align=""center""></td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					</tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					<tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					  <td height=""30"" align=""right""><a href=""javascript:window.close();"" class=""red01""><strong>[关闭本页]</strong></a>&nbsp;&nbsp;</td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "					</tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				  </table>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "				</td>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "			  </tr>" &chr(13)&chr(10)
knowledgelist = knowledgelist & "			</table>" &chr(13)&chr(10)


templatefileName="../template/knowledge_list.asp"
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
str=replace(str,"$FriendLink$",FriendLink)

createfileName="../knowledge_list_" & rs("id") & ".htm"
Set htmlFile = fso.CreateTextFile(server.MapPath(createfileName))
htmlFile.write(str)
set htmlFile=nothing
response.write "正在生成装修知识..."&createfileName&"<br>"

rs.movenext
loop
end If
rs.close
set rs=Nothing
response.write "装修知识详细信息生成成功<br>"
'==================================================装修知识详细信息结束===================================================
conn.close
set conn=nothing
set fso=nothing
%>