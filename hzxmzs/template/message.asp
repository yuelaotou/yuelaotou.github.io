<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="$Keywords$">
<meta name="Keywords" content="$Keywords$">
<title>$Site_Title$</title>
<script src="js/common.js" type="text/javascript"></script>
<script src="js/marquee.js" type="text/javascript"></script>
<link href="css/common.css" rel="stylesheet" type="text/css" />
$CopyRight$
</head>
<body>
<div id="passport">
  <div class="main">
    <div class="date"> ����ʱ��:&nbsp;
      <script>
		var now = new Date();
		var mm = now.getMonth() + 1;
		var dd = now.getDate();
		var week;
		if(now.getDay()==0)
			week="������";
		if(now.getDay()==1)
			week="����һ";
		if(now.getDay()==2)
			week="���ڶ�";
		if(now.getDay()==3)
			week="������";
		if(now.getDay()==4)
			week="������";
		if(now.getDay()==5)
			week="������";
		if(now.getDay()==6)
			week="������";
		var day=now.getFullYear() + "��" + (mm < 10 ? "0" + mm : mm) + "��" + (dd < 10 ? "0" + dd : dd) + "�� ";
		document.write(day + " " + week );
	  </script>
    </div>
    <div class="welcome">
      <marquee onmouseover="this.stop()" onmouseout="this.start()"
              scrollamount="2" scrollDelay="20" width="500" height="20">
      <img src="img/icon1.gif"/>$Address$
      </marquee>
    </div>
    <div class="homepage"> <a class="i_setIndex" target="_self" href="http://www.hzxmzs.cn/" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.hzxmzs.cn/');return false;" title="�����װ����Ϊ��ҳ">��Ϊ��ҳ</a> <a href="javascript:window.external.AddFavorite('http://www.hzxmzs.cn','�������װ�ι������޹�˾')" class="i_reco">����ղ�</a> </div>
  </div>
</div>
<div class="blank4"></div>
<div id="container">
  <div id="banner">
	$flash_message$
  </div>
  <div class="blank4"></div>
  <div id="menu">
    <ul>
      <li><a href="/">��վ��ҳ</a></li>
      <li><a href="about.htm">��˾���</a></li>
      <li><a href="culture.htm">��ҵ�Ļ�</a></li>
      <li><a href="case.htm">���̰���</a></li>
      <li><a href="service.htm">�ͻ�����</a></li>
      <li><a href="site.htm">ʩ���ֳ�</a></li>
      <li><a href="knowledge.htm">װ��֪ʶ</a></li>
      <li><a href="employee.htm">������ʿ</a></li>
      <li><a href="message.htm" class="selected">������ѯ</a></li>
      <li style="margin-right:0px;"><a href="contact.htm">��ϵ����</a></li>
    </ul>
  </div>
  <div class="blank4"></div>
  <div id="contenttop">
    <div id="colLeft">
      <div class="colL colHeight219" style="border:0px;width:237px;height:220px;"><!--������Ŀ-->
		$serviceShow$
      </div>
      <div class="blank9"></div>
      <div class="colL colHeight200">
        <div class="titleBar ui_tit26">
          <div class="tit"><a href="contact.htm">��ϵ����</a></div>
        </div>
        <div class="content">
			$contactUs$
        </div>
      </div>
    </div>
    <div id="colAbout">
      <div class="colAbout contentshow font14">
        <div class="titleBar ui_tit26">
          <div class="tit">������ѯ</div>
        </div>
        <div class="content">
		  <ul class="pagination" title="��ҳ�б�">
$pagination$
		  </ul>
		  $message$
		  <table width="700" border="0" align="left" cellpadding="0" cellspacing="0" style="border:1px solid #ddd">
			<form action="template/messageadd.asp?send=add" name="msg" method="post">
			  <tr>
				<td valign="top" align="left" style="padding:15px">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td colspan="2" ><img src="img/message_11.gif" style="width:85px;height:33px;"/><a name="a02" id="a02"></a></td>
					</tr>
					<tr>
					  <td width="50%" valign="top">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td height="35">�ǡ����ƣ�
							  <input type="text" name="nickname" style="font-size:12px; color:#666;width:180px" />
							</td>
						  </tr>
						  <tr>
							<td height="35">�������䣺
							  <input type="text" name="email" style="font-size:12px; color:#666;width:180px" />
							</td>
						  </tr>
						  <tr>
							<td height="35">��ϵ�绰��
							  <input type="text" name="tel" style="font-size:12px; color:#666;width:180px" maxlength="13"/>
							</td>
						  </tr>
						  <tr>
							<td height="35">Q Q ���룺
							  <input type="text" name="qq" style="font-size:12px; color:#666;width:180px" />
							</td>
						  </tr>
						</table>
					  </td>
					  <td width="50%" valign="top">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td height="32">���Ա��⣺
							  <input type="text" name="title" style="font-size:12px; color:#666;width:180px" />
							</td>
						  </tr>
						  <tr>
							<td height="73">�������ݣ�
							  <textarea name="content" style="font-size:12px; color:#666;width:200px; height:60px"></textarea>
							</td>
						  </tr>
						  <tr>
							<td height="35" align="center">
							  <input type="submit" name="Submit22" value=" �� �� " onClick="return checkform();" />
							  <input type="reset" name="Submit22" value=" �� �� " />
							</td>
						  </tr>
						</table>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			</form>
		  </table>
        </div>
      </div>
    </div>
  </div>
  <div class="blank9"></div>
  <div id="FriendSite">
    <ul>
$FriendLink$
    </ul>
  </div>
</div>
<div class="blank9"></div>
<div id="about">
  <ul>
     <li> <a href="aboutUs.htm" target="_blank">&nbsp;�������&nbsp;</a>��-��<a href="about.htm"  target="_blank">&nbsp;��˾���&nbsp;</a>��-��<a href="service.htm"  target="_blank">&nbsp;�ͻ�����&nbsp;</a>��-��<a href="knowledge.htm"  target="_blank">&nbsp;װ����Ѷ&nbsp;</a>��-��<a href="sitemap.htm"  target="_blank">&nbsp;��վ��ͼ&nbsp;</a>��-��<a href="contact.htm" target="_blank">&nbsp;��ϵ����&nbsp;</a> </li>
  </ul>
</div>
<div id="footer">
	$ICP$
</div>
</body>
</html>