<% 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���˷Ƿ���SQL�ַ�
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
'ɾ��һ��
Sub del(table,id,url)
	sql = "delete from "&table&" where id="&id
	set rs=conn.execute(sql)
	call alert("ɾ���ɹ�",url)
End sub

Sub del_all(table,url)
	For Each I In Request.Form("ID")
		sql = "delete from "&table&" where id="&I
		set rs=conn.execute(sql)
	Next
	conn.close()
	call alert("ɾ���ɹ�",url)
End Sub

'����
Sub OutScript(srt)
	Response.Write "<script language=javascript>alert('"&srt&"!');history.back()</script>"
End Sub

'��ʾ
Sub alert(title,url)
	Response.Write "<script>"
	Response.Write "alert('"&title&"!');location='"&url&"';"
	Response.Write "</script>"
	Response.end
End sub

'����
Sub Lock_url(table,id,locks,url) 
	sql="update "&table&" set lock="&locks&" where id="&id
	conn.execute(sql)
	call alert("�����ɹ�",url)
End sub
'���
Sub Check_url(table,id,check,url) 
	sql="update "&table&" set status="&check&" where id="&id
	conn.execute(sql)
	call alert("�����ɹ�",url)
End sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��˼�¼
'Lock_OK(��,���,���,ת��)
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
	Response.write("<script>alert('�����ɹ�');location='"&Url&"';</script>")
	Response.End
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ɾ��ͼƬ
'DelImage(��,���,��¼,ת��)
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
	Response.Write "<script>alert('ɾ���ɹ�!');location='"&Url&"';</script>"
	Response.End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ҳ
Sub Image_Page(url)
	If page<>1 Then
		Response.write "<a href='?page=1"&url&"&SS_ID="&request("SS_ID")&"'>��ҳ</a>"
	Else
		Response.write "��ҳ"
	End If
		Response.write "&nbsp;"
	If page>1 Then
		Response.write "<a href='?page="&page-1&""&url&"&SS_ID="&request("SS_ID")&"'>��һҳ</a>"
	Else
		Response.write "��һҳ"
	End if
		Response.write "&nbsp;"
	If page<rs.pagecount Then
		Response.Write "<a href='?page="&page+1&""&url&"&SS_ID="&request("SS_ID")&"'>��һҳ</a>"
	Else
		Response.Write "��һҳ"
	End If
		Response.Write("&nbsp;")
	If Rs.recordcount mod Rs.Pagesize = 0 Then
		sintPageMax = Rs.recordcount \ Rs.Pagesize
	Else
		sintPageMax = Rs.recordcount \ Rs.Pagesize + 1
	End If
	
	If Page<sintPageMax Then
		Response.Write("<a href='?page="&sintPageMax&""&url&"&SS_ID="&request("SS_ID")&"'>ĩҳ</a>")
	Else
		Response.Write("ĩҳ")
	End If
		Response.Write("&nbsp;�ϼ�"&Rs.recordcount&"��")
	If Rs.recordcount<>0 then
		Response.write ("&nbsp;��"&page&"")
	Else
		Response.write ("&nbsp;��"&sintPageMax&"")
	End If
		Response.write "/"
		Response.write (""&sintPageMax&"")
		Response.write "ҳ&nbsp;"
	
	
	strTemp="&nbsp;ת������<select name='page' onchange='window.location=(this.options[selectedIndex].value)'>"   
	for i = 1 to sintPageMax   
		strTemp=strTemp & "<option value='?page="&i&""&url&"&SS_ID="&request("SS_ID")&"'"
		if cint(page)=cint(i) then strTemp=strTemp & " selected "
		strTemp=strTemp & ">" & i & "</option>"   
	next
	strTemp=strTemp & "</select>ҳ"
	
	response.write strTemp
End Sub


'��ʾ��
Sub Type_Title(SS_ID)
	Type_SQL="select * from Type where id="&SS_ID
	Set rc=conn.execute(Type_SQL)
	response.write(rc("Title"))
	rc.close()
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��֤�û�Ȩ��
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
	response.write "<script>alert('ϵͳ��ʾ��\n\n�˲���û����Ȩ��');history.go(-1);</script>"
	response.End
  End IF
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'û�м�¼
Sub RS_Empty()
	IF RS.bof And RS.Eof Then
		Response.Write("<center><BR><BR>������Ϣ<BR><BR></center>")
	End IF
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��˼�¼
'Lock_Old(��,���,���,ת��)
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
	Response.write("<script>alert('�����ɹ�');location='"&Url&"';</script>")
	Response.End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ɾ��ͼƬ
'DelImage(��,���,��¼,ת��)
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
	Response.Write "<script>alert('[ͼƬ]ɾ���ɹ�!');location='"&Url&"';</script>"
	Response.End
End Sub

 %>
<%
' ============================================
' ��ʽ��ʱ��(��ʾ)
' ������n_Flag
'   1:"yyyy-mm-dd hh:mm:ss"
'   2:"yyyy-mm-dd"
'   3:"hh:mm:ss"
'   4:"yyyy��mm��dd��"
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
        ' yyyy��mm��dd��
        Format_Time = y & "��" & m & "��" & d & "��"
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
'��������strLength
'��   �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��   ����str   ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
ON ERROR RESUME NEXT
dim WINNT_CHINESE
WINNT_CHINESE     = (len("����")=2)
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
 myString=Replace(myString,"&","&amp;")         '�滻&Ϊ�ַ�ʵ��&amp;
 myString=Replace(myString,"&lt;","&lt;")          '�滻&lt;
 myString=Replace(myString,"&gt;","&gt;")          '�滻&gt;
 myString=Replace(myString,"","&lt;br&gt;")      '�滻�س���  
 myString=Replace(myString,chr(32),"&nbsp;")    '�滻�ո��
 myString=Replace(myString,chr(9),"&nbsp; &nbsp; &nbsp; &nbsp;")     '�滻Tab������
 myString=Replace(myString,chr(39),"&acute;")   '�滻������
 myString=Replace(myString,chr(34),"&quot;")    '�滻˫����
 myReplace=myString                             '���غ���ֵ
End Function



%>
