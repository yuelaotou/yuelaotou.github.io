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
    <div class="date"> 现在时刻:&nbsp;
      <script>
		var now = new Date();
		var mm = now.getMonth() + 1;
		var dd = now.getDate();
		var week;
		if(now.getDay()==0)
			week="星期日";
		if(now.getDay()==1)
			week="星期一";
		if(now.getDay()==2)
			week="星期二";
		if(now.getDay()==3)
			week="星期三";
		if(now.getDay()==4)
			week="星期四";
		if(now.getDay()==5)
			week="星期五";
		if(now.getDay()==6)
			week="星期六";
		var day=now.getFullYear() + "年" + (mm < 10 ? "0" + mm : mm) + "月" + (dd < 10 ? "0" + dd : dd) + "日 ";
		document.write(day + " " + week );
	  </script>
    </div>
    <div class="welcome">
      <marquee onmouseover="this.stop()" onmouseout="this.start()"
              scrollamount="2" scrollDelay="20" width="500" height="20">
      <img src="img/icon1.gif"/>$Address$
      </marquee>
    </div>
    <div class="homepage"> <a class="i_setIndex" target="_self" href="http://www.hzxmzs.cn/" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.hzxmzs.cn/');return false;" title="把浔美装饰设为首页">设为首页</a> <a href="javascript:window.external.AddFavorite('http://www.hzxmzs.cn','杭州浔美装饰工程有限公司')" class="i_reco">添加收藏</a> </div>
  </div>
</div>
<div class="blank4"></div>
<div id="container">
  <div id="banner">
	$flash_employee$
  </div>
  <div class="blank4"></div>
  <div id="menu">
    <ul>
      <li><a href="/">网站首页</a></li>
      <li><a href="about.htm">公司简介</a></li>
      <li><a href="culture.htm">企业文化</a></li>
      <li><a href="case.htm">工程案例</a></li>
      <li><a href="service.htm">客户服务</a></li>
      <li><a href="site.htm">施工现场</a></li>
      <li><a href="knowledge.htm">装修知识</a></li>
      <li><a href="employee.htm" class="selected">招贤纳士</a></li>
      <li><a href="message.htm">留言咨询</a></li>
      <li style="margin-right:0px;"><a href="contact.htm">联系我们</a></li>
    </ul>
  </div>
  <div class="blank4"></div>
  <div id="contenttop">
    <div id="colLeft">
      <div class="colL colHeight219" style="border:0px;width:237px;height:220px;"><!--服务项目-->
		$serviceShow$
      </div>
      <div class="blank9"></div>
      <div class="colL colHeight200">
        <div class="titleBar ui_tit26">
          <div class="tit"><a href="contact.htm">联系我们</a></div>
        </div>
        <div class="content">
			$contactUs$
        </div>
      </div>
    </div>
    <div id="colAbout">
      <div class="colAbout contentshow font14">
        <div class="titleBar ui_tit26">
          <div class="tit">招贤纳士</div>
        </div>
        <div class="content">
		  <ul class="pagination" title="分页列表">
$pagination$
		  </ul>
		  $employee$
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
     <li> <a href="aboutUs.htm" target="_blank">&nbsp;关于浔美&nbsp;</a>　-　<a href="about.htm"  target="_blank">&nbsp;公司简介&nbsp;</a>　-　<a href="service.htm"  target="_blank">&nbsp;客户服务&nbsp;</a>　-　<a href="knowledge.htm"  target="_blank">&nbsp;装修资讯&nbsp;</a>　-　<a href="sitemap.htm"  target="_blank">&nbsp;网站地图&nbsp;</a>　-　<a href="contact.htm" target="_blank">&nbsp;联系我们&nbsp;</a> </li>
  </ul>
</div>
<div id="footer">
	$ICP$
</div>
</body>
</html>