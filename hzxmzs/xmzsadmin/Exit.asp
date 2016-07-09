<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
Session.Contents.Remove("Uid")
Response.Write "<script>"
Response.Write "parent.location='../';"
Response.Write "</script>"
Response.end
 %>