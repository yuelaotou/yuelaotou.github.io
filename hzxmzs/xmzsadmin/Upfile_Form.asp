<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Session.asp" -->

<html>
<head>
<title>上传文件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/Style_body.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../JS/keydown.JS"></script>
<script language="JavaScript" src="../JS/ShowProcessBar.JS"></script>
<SCRIPT language=javascript>
function check(theForm) 
{
	var strFileName=document.theForm.file1.value;
	if (strFileName=="")
	{
    	alert("系统提示:\n\n请选择要上传的文件");
		document.theForm.file1.focus();
    	return false;
  	}
}
</SCRIPT>
<link href="../css/style_body.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#EEFFEE" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false;" onselectstart="return false;">

<table width="300" border="0" align="center" cellpadding="10" cellspacing="10">
  <tr>
            <td>		
  <fieldset style="padding: 5" class="body">
              <legend><strong>文件上传</strong></legend>
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td height="50" align="center"> 
            <form action="UpFile_Save.asp?UpfileType=<%= Request.QueryString("UpfileType") %>" method="post" enctype="multipart/form-data" name="form1" onSubmit="return check(this)">
              <input name="file1" type="file" value="" size="10">
              <input type="submit" name="Submit" value="上传" IsShowProcessBar="True">
            </form></td>
        </tr>
      </table>
      </fieldset>
   </td>
  </tr>
</table>



</body>
</html>
 