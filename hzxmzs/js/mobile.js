function checkMobile()
{
var mobile=document.theForm.MobileNum.value;
var reg0 = /^13\d{5,9}$/; //130--139������5λ�����7λ
var reg1 = /^153\d{4,8}$/; //��ͨ153������4λ�����8λ
var reg2 = /^159\d{4,8}$/; //�ƶ�159������4λ�����8λ
//var reg3 = /^0\d{10,11}$/; 
var my = false;
if (reg0.test(mobile))my=true;
if (reg1.test(mobile))my=true;
if (reg2.test(mobile))my=true;
//if (reg3.test(mobile))my=true;
if (!my)
{
//document.mobileform.MobileNum.value='';
alert('�Բ�����������ֻ������д���');
//document.mobileform.MobileNum.focus();
return false;
}

return true; 

} 