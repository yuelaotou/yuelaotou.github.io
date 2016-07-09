<!--
function Login_ok(theForm){
    var text="";

	//验证用户名
    var M1 = theForm.uid.value; 
	if (M1 == "") {text=text+"[用户名] 不能为空；\n"}
	else {
		if (M1.indexOf("'") > -1 || M1.indexOf('"') > -1) {text=text+"[用户名] 不充许有非法字符；\n"}
	}
	
	//验证用户名
    var M2 = theForm.pwd.value; 
	var L2 = M2.length;
	if (M2 == "") {text=text+"[密 码] 不能为空；\n"}
	
//验证验证码
    var M3 = theForm.yzm.value; 
	var L3 = M3.length;
    if (M3 == "") {text=text+"[验证码] 不能为空；\n"}  
	else {
		if (M3.indexOf("'") > -1 || M3.indexOf('"') > -1) {text=text+"[验证码] 不充许有非法字符；\n"}
		else {
			if (isNaN(M3) || L3 < 4) {text=text+"[验证码] 必须为4位数字；\n"}
		}
	}

	//向页面反馈信息	
   if (text == ""){
   //document.Login.submit();
   return true;
   }
   else{
    alert(text);
    return false;
    }           
}
-->
