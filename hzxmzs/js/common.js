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
		alert("����������!");
		document.msg.nickname.focus();
		return false;
	}
	else if(document.msg.email.value=="")
	{
		alert("�����ϵ�������!");
		document.msg.email.focus();
		return false;
	}
 	else if(document.msg.tel.value=="")
    {
		alert("��������ϵ�绰!");
		document.msg.tel.focus();
		return false;
	}
	else if(document.msg.title.value=="")
	{
		alert("���������Ա���!");
		document.msg.title.focus();
		return false;
	}
	else if(document.msg.content.value=="")
	{
		alert("��������������!");
		document.msg.content.focus();
		return false;
	}
	return true;
}