<!--
function Login_ok(theForm){
    var text="";

	//��֤�û���
    var M1 = theForm.uid.value; 
	if (M1 == "") {text=text+"[�û���] ����Ϊ�գ�\n"}
	else {
		if (M1.indexOf("'") > -1 || M1.indexOf('"') > -1) {text=text+"[�û���] �������зǷ��ַ���\n"}
	}
	
	//��֤�û���
    var M2 = theForm.pwd.value; 
	var L2 = M2.length;
	if (M2 == "") {text=text+"[�� ��] ����Ϊ�գ�\n"}
	
//��֤��֤��
    var M3 = theForm.yzm.value; 
	var L3 = M3.length;
    if (M3 == "") {text=text+"[��֤��] ����Ϊ�գ�\n"}  
	else {
		if (M3.indexOf("'") > -1 || M3.indexOf('"') > -1) {text=text+"[��֤��] �������зǷ��ַ���\n"}
		else {
			if (isNaN(M3) || L3 < 4) {text=text+"[��֤��] ����Ϊ4λ���֣�\n"}
		}
	}

	//��ҳ�淴����Ϣ	
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
