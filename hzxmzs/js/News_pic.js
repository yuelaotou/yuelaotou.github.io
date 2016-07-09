<!--
/*
<% 
int SS_ID=69;
string sql = "Select * from News where SS_ParentID="+SS_ID+" and showLinkImage=1 order by ID Desc";
SqlDataReader dr = DB.ExecuteReader(sql);
int i=1;
string pics=null,links=null,texts=null;
while(dr.Read())
{
	if(i>5)
		break;
	pics += "UpLoadFile/"+dr["LinkImage"].ToString()+"|";
	links += "house_NewsList.aspx?ID="+dr["ID"].ToString()+"|";
	texts += dr["Title"].ToString() +"|";
	i +=1;
}
dr.Close();
%>
<script language="JavaScript" type="text/javascript">
var focus_width=;
var focus_height=;
var text_height=;
var pics='<%= pics %>';
var links='<%= links %>';
var texts='<%= texts %>';
<script>
<script language="JavaScript" type="text/javascript" src="JS/News_pic.js?"><script>
*/
var swf_height = focus_height+text_height;
pics = pics.substring(0,pics.length-1);
links = links.substring(0,links.length-1);
texts = texts.substring(0,texts.length-1);
var str = '<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">';
str +='<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="Swf/pixviewer.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">';
str +='<param name="menu" value="false"><param name=wmode value="opaque">';
str +='<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">';
str +='<embed src="Swf/pixviewer.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#ffffff" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer"/></object>';
document.write(str);
-->
