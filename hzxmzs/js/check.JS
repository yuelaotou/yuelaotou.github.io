<!--添加用户页面的验证代码__>
function input_ok(theForm){
    var text="";

//验证用户名
    var M1 = theForm.uid.value; 
    if (M1 == ""){text=text+"[帐号] 不能为空；\n"} 

//验证密码
    var M2 = theForm.pwd.value;
	var L2 = M2.length;
    if (M2 == ""){text=text+"[密码] 不能为空；\n"} 
	if (L2 < 6){text=text+"[密码长度]不能小于六位；\n"}

//验证姓名
    var M2 = theForm.Real_Name.value; 
    if (M2 == ""){text=text+"[真实姓名] 不能为空；\n"} 


//验证Email
    var M3 = theForm.Email.value; 
    if (M3 == "")
		{text=text+"[Email] 不能为空；\n"} 
	else
		{var p=M3.indexOf('@');
		if (p<1 || p==(M3.length-1)) {text=text+"[Email] 填写不规范；\n";}
		}



//验证QQ
    var M4 = theForm.qq.value; 
    if (M4 == "") {
		text=text+"[qq] 不能为空；\n"} 
	else {
		if (isNaN(M4)) {text=text+"[qq] 必须为数值；\n"}
	}

//验证电话
   var M5 = theForm.Tel.value;
   var L5=M5.Length
   if (M5 == ""){text=text+"[电话] 不能为空；\n"} 
	if(L5<7){text=text+"[电话长度]不能少于7位；\n"}

//验证地址
    var M6 = theForm.address.value; 
    if (M6 == ""){text=text+"[地址] 不能为空；\n"} 

//向页面反馈信息	
   if (text == "") {
	   //document.theForm.submit();
	   }
   else {
    alert("出错提示：\n\n" + text);
    return false;
    }           
}

