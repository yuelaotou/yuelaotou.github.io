<%
Sql="select * from [Config] where id=1"
set Rs=server.createobject("adodb.recordset")
Rs.open sql,conn,1,1
Site_Title = rs("Site_Title")
Domain = rs("Domain")
Keywords = rs("Keywords") 
Address = rs("Address")
CopyRight = rs("CopyRight")
ICP = rs("Icp")
service_flash = RS("service")
flash_index = RS("flash_index")
flash_about = RS("flash_about")
flash_culture = RS("flash_culture")
flash_case = RS("flash_case")
flash_service = RS("flash_service")
flash_site = RS("flash_site")
flash_knowledge = RS("flash_knowledge")
flash_employee = RS("flash_employee")
flash_message = RS("flash_message")
flash_contact = RS("flash_contact")
flash_aboutUs = RS("flash_aboutUs")
flash_shinei = RS("flash_shinei")
flash_dianmian = RS("flash_dianmian")
flash_bangongshi = RS("flash_bangongshi")
flash_bieshu = RS("flash_bieshu")
flash_xiezilou = RS("flash_xiezilou")
flash_xfbp = RS("flash_xfbp")
%>