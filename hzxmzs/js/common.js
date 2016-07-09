function goto(type,n){
	if (n==0){
		window.location=type + ".htm";
	}else{
		window.location=type + "_" + n + ".htm";
	}
}

function checkform()
{
	if(document.msg.nickname.value=="")
	{
		alert("请填上姓名!");
		document.msg.nickname.focus();
		return false;
	}
	else if(document.msg.email.value=="")
	{
		alert("请填上电子邮箱!");
		document.msg.email.focus();
		return false;
	}
 	else if(document.msg.tel.value=="")
    {
		alert("请填上联系电话!");
		document.msg.tel.focus();
		return false;
	}
	else if(document.msg.title.value=="")
	{
		alert("请填上留言标题!");
		document.msg.title.focus();
		return false;
	}
	else if(document.msg.content.value=="")
	{
		alert("请填上留言内容!");
		document.msg.content.focus();
		return false;
	}
	return true;
}