var timer = null;
// 显示数字图片
function displayClock(num){
  var dig = parseInt(num/10);
  var timetag="<img src='../sysImages/clock/" + dig + ".gif'>";
  dig = num%10;
  timetag +="<img src='../sysImages/clock/" + dig + ".gif'>";
  return timetag;
}

// 开始计时
function startClock(){
var time = new Date();
// 取得时间
var hours = displayClock(time.getHours()) + ":";
var minutes = displayClock(time.getMinutes()) + ":";
var seconds = displayClock(time.getSeconds());
// 显示时间
show.innerHTML = hours + minutes + seconds;
timer = setTimeout("startClock()",1000);
}
