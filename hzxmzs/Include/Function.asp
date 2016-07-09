<% 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'过滤非法的SQL字符
Function ReplaceChar(strChar)
	IF strChar = Empty Then
		ReplaceChar = Empty
	Else
		ReplaceChar = replace(strChar,"'","")
		ReplaceChar = replace(strChar,"*","")
		ReplaceChar = replace(strChar,"?","")
		ReplaceChar = replace(strChar,"(","")
		ReplaceChar = replace(strChar,")","")
		ReplaceChar = replace(strChar,"<","")
		ReplaceChar = replace(strChar,".","")
        ReplaceChar = replace(ReplaceChar," ","")
        ReplaceChar = replace(ReplaceChar,";","")
		ReplaceChar = LCase(ReplaceChar)
        ReplaceChar = replace(ReplaceChar,"or","")
        ReplaceChar = replace(ReplaceChar,"and","")
        ReplaceChar = replace(ReplaceChar,"not","")
        ReplaceChar = replace(ReplaceChar,"select","")
        ReplaceChar = replace(ReplaceChar,"drop","")
        ReplaceChar = replace(ReplaceChar,"delete","")
        ReplaceChar = replace(ReplaceChar,"update","")
        ReplaceChar = replace(ReplaceChar,"insert","")
        ReplaceChar = replace(ReplaceChar,"count","")
        ReplaceChar = replace(ReplaceChar,"exec","")
        ReplaceChar = replace(ReplaceChar,"truncate","")
        ReplaceChar = replace(ReplaceChar,"net","")
        ReplaceChar = replace(ReplaceChar,"asc","")
        ReplaceChar = replace(ReplaceChar,"char","")
        ReplaceChar = replace(ReplaceChar,"mID","")
	End IF
End Function
'删除一条
Sub del(table,id,url)
	sql = "delete from "&table&" where id="&id
	set rs=conn.execute(sql)
	call alert("删除成功",url)
End sub

Sub del_all(table,url)
	For Each I In Request.Form("ID")
		sql = "delete from "&table&" where id="&I
		set rs=conn.execute(sql)
	Next
	conn.close()
	call alert("删除成功",url)
End Sub

'返回
Sub OutScript(srt)
	Response.Write "<script language=javascript>alert('"&srt&"!');history.back()</script>"
End Sub

'提示
Sub alert(title,url)
	Response.Write "<script>"
	Response.Write "alert('"&title&"!');location='"&url&"';"
	Response.Write "</script>"
	Response.end
End sub

'锁定
Sub Lock_url(table,id,locks,url) 
	sql="update "&table&" set lock="&locks&" where id="&id
	conn.execute(sql)
	call alert("操作成功",url)
End sub
'审核
Sub Check_url(table,id,check,url) 
	sql="update "&table&" set status="&check&" where id="&id
	conn.execute(sql)
	call alert("操作成功",url)
End sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'审核记录
'Lock_OK(表,编号,标记,转到)
Sub Lock_OK(Table,ID,ReCord,Locks,Url)
	Dim RS,Sql
	set RS=server.createobject("adodb.recordset")
	Sql = "select * from "&Table&" where ID="&ID&""
	RS.open sql,conn,1,3
	RS(""&ReCord&"") = Locks
	RS.Update
	RS.Requery
	RS.close
	Set RS=nothing
	Response.write("<script>alert('操作成功');location='"&Url&"';</script>")
	Response.End
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除图片
'DelImage(表,编号,记录,转到)
Sub Del_File(Table,ID,ReCord,Url)
	set RS=Server.createObject("adodb.recordset")
	Sql="select * from "&Table&" where ID="&ID&""
	RS.open sql,conn,1,3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set Fso=server.createobject("Scripting.FileSystemObject")
		Upfile= server.MapPath("../UploadFile/"&rs(""&ReCord&""))
	IF ImgLogo<>Empty Then
		Fso.deletefile(Upfile)
		Fso.Close	
	End IF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	RS(""&ReCord&"")= Empty
	RS.update
	Response.Write "<script>alert('删除成功!');location='"&Url&"';</script>"
	Response.End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'分页
Sub Image_Page(url)
	If page<>1 Then
		Response.write "<a href='?page=1"&url&"&SS_ID="&request("SS_ID")&"'>首页</a>"
	Else
		Response.write "首页"
	End If
		Response.write "&nbsp;"
	If page>1 Then
		Response.write "<a href='?page="&page-1&""&url&"&SS_ID="&request("SS_ID")&"'>上一页</a>"
	Else
		Response.write "上一页"
	End if
		Response.write "&nbsp;"
	If page<rs.pagecount Then
		Response.Write "<a href='?page="&page+1&""&url&"&SS_ID="&request("SS_ID")&"'>下一页</a>"
	Else
		Response.Write "下一页"
	End If
		Response.Write("&nbsp;")
	If Rs.recordcount mod Rs.Pagesize = 0 Then
		sintPageMax = Rs.recordcount \ Rs.Pagesize
	Else
		sintPageMax = Rs.recordcount \ Rs.Pagesize + 1
	End If
	
	If Page<sintPageMax Then
		Response.Write("<a href='?page="&sintPageMax&""&url&"&SS_ID="&request("SS_ID")&"'>末页</a>")
	Else
		Response.Write("末页")
	End If
		Response.Write("&nbsp;合计"&Rs.recordcount&"项")
	If Rs.recordcount<>0 then
		Response.write ("&nbsp;第"&page&"")
	Else
		Response.write ("&nbsp;第"&sintPageMax&"")
	End If
		Response.write "/"
		Response.write (""&sintPageMax&"")
		Response.write "页&nbsp;"
	
	
	strTemp="&nbsp;转到：第<select name='page' onchange='window.location=(this.options[selectedIndex].value)'>"   
	for i = 1 to sintPageMax   
		strTemp=strTemp & "<option value='?page="&i&""&url&"&SS_ID="&request("SS_ID")&"'"
		if cint(page)=cint(i) then strTemp=strTemp & " selected "
		strTemp=strTemp & ">" & i & "</option>"   
	next
	strTemp=strTemp & "</select>页"
	
	response.write strTemp
End Sub


'显示类
Sub Type_Title(SS_ID)
	Type_SQL="select * from Type where id="&SS_ID
	Set rc=conn.execute(Type_SQL)
	response.write(rc("Title"))
	rc.close()
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'验证用户权限
Function Check_Right(ID)
  dim Flag,Sql,RS,I
  Flag = 0
  Sql="select * from Admin where uid='"&session("uid")&"'"
  set RS=Conn.Execute(Sql)
  IF RS("Rights") <> Empty Then 
    Return=Split(RS("Rights"),",")
    I=0
    For I = LBound(Return) To UBound(Return)
      IF trim(ID)=trim(Return(I)) Then
	    Flag = 1
	  End IF
    Next	
  End IF
  IF Flag = 0 Then 
	On Error Resume Next
	Rs.Close
	Set Rs = Nothing
	Conn.Close
	Set Conn = Nothing
	response.write "<script>alert('系统提示：\n\n此操作没有授权！');history.go(-1);</script>"
	response.End
  End IF
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'没有记录
Sub RS_Empty()
	IF RS.bof And RS.Eof Then
		Response.Write("<center><BR><BR>暂无信息<BR><BR></center>")
	End IF
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'审核记录
'Lock_Old(表,编号,标记,转到)
Sub Lock_old(Table,ID,ReCord,Locks,Url)
	Dim RS,Sql
	set RS=server.createobject("adodb.recordset")
	Sql = "select * from "&Table&" where ID="&ID&""
	RS.open sql,conn,1,3
	RS(""&ReCord&"") = Locks
	RS.Update
	RS.Requery
	RS.close
	Set RS=nothing
	Response.write("<script>alert('操作成功');location='"&Url&"';</script>")
	Response.End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除图片
'DelImage(表,编号,记录,转到)
Sub DelImage(Table,ID,ReCord,Url)
	set RS=Server.createObject("adodb.recordset")
	Sql="select * from "&Table&" where ID="&ID&""
	RS.open sql,conn,1,3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set Fso=server.createobject("Scripting.FileSystemObject")
		Image= server.MapPath("../../UploadFile/"&rs(""&ReCord&""))
	IF ImageLogo<>Empty Then
		Fso.deletefile(Image)
		Fso.Close	
	End IF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	RS(""&ReCord&"")= Empty
	RS.update
	Response.Write "<script>alert('[图片]删除成功!');location='"&Url&"';</script>"
	Response.End
End Sub

 %>
<%
' ============================================
' 格式化时间(显示)
' 参数：n_Flag
'   1:"yyyy-mm-dd hh:mm:ss"
'   2:"yyyy-mm-dd"
'   3:"hh:mm:ss"
'   4:"yyyy年mm月dd日"
'   5:"mm-dd"
' ============================================
Public Function Format_Time(s_Time, n_Flag)
    Dim y, m, d, h, mi, s
    Format_Time = ""
    If IsDate(s_Time) = False Then Exit Function
    y = CStr(Year(s_Time))
    m = CStr(Month(s_Time))
    If Len(m) = 1 Then m = "0" & m
    d = CStr(Day(s_Time))
    If Len(d) = 1 Then d = "0" & d
    h = CStr(Hour(s_Time))
    If Len(h) = 1 Then h = "0" & h
    mi = CStr(Minute(s_Time))
    If Len(mi) = 1 Then mi = "0" & mi
    s = CStr(Second(s_Time))
    If Len(s) = 1 Then s = "0" & s
    Select Case n_Flag
    Case 1
        ' yyyy-mm-dd hh:mm:ss
        Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
    Case 2
        ' yyyy-mm-dd
        Format_Time = y & "-" & m & "-" & d
    Case 3
        ' hh:mm:ss
        Format_Time = h & ":" & mi & ":" & s
    Case 4
        ' yyyy年mm月dd日
        Format_Time = y & "年" & m & "月" & d & "日"
    Case 5
        ' mm-dd
        Format_Time = m & "-" & d
    End Select
End Function 
sub CutStr(title,lenght)
title=title
if len(title)>=lenght then
title=left(title,lenght)&"..."
end if
response.Write title
end sub

Function newleft(str,l)
if strLength(str)>l then
   newleft=cutStr(str,l)
else
   newleft=str
end if
End Function
'**************************************************
'函数名：strLength
'作   用：求字符串长度。汉字算两个字符，英文算一个字符。
'参   数：str   ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
ON ERROR RESUME NEXT
dim WINNT_CHINESE
WINNT_CHINESE     = (len("网络")=2)
if WINNT_CHINESE then
         dim l,t,c
         dim i
         l=len(str)
         t=l
         for i=1 to l
          c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                 t=t+1
             end if
         next
         strLength=t
     else 
         strLength=len(str)
     end if
     if err.number<>0 then err.clear
end function
Function cutStr(str,strlen) 
dim l,t,c 
l=len(str) 
t=0 
for i=1 to l 
c=Abs(Asc(Mid(str,i,1))) 
if c>255 then 
t=t+2 
else 
t=t+1 
end if 
if t>=strlen then 
cutStr=left(str,i)&"..." 
exit for 
else 
cutStr=str 
end if 
next 
cutStr=replace(cutStr,chr(10),"") 
end Function

Function myReplace(myString)
 myString=Replace(myString,"&","&amp;")         '替换&为字符实体&amp;
 myString=Replace(myString,"&lt;","&lt;")          '替换&lt;
 myString=Replace(myString,"&gt;","&gt;")          '替换&gt;
 myString=Replace(myString,"","&lt;br&gt;")      '替换回车符  
 myString=Replace(myString,chr(32),"&nbsp;")    '替换空格符
 myString=Replace(myString,chr(9),"&nbsp; &nbsp; &nbsp; &nbsp;")     '替换Tab缩进符
 myString=Replace(myString,chr(39),"&acute;")   '替换单引号
 myString=Replace(myString,chr(34),"&quot;")    '替换双引号
 myReplace=myString                             '返回函数值
End Function



%>
