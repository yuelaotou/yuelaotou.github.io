function checkMobile()
{
var mobile=document.theForm.MobileNum.value;
var reg0 = /^13\d{5,9}$/; //130--139。至少5位，最多7位
var reg1 = /^153\d{4,8}$/; //联通153。至少4位，最多8位
var reg2 = /^159\d{4,8}$/; //移动159。至少4位，最多8位
//var reg3 = /^0\d{10,11}$/; 
var my = false;
if (reg0.test(mobile))my=true;
if (reg1.test(mobile))my=true;
if (reg2.test(mobile))my=true;
//if (reg3.test(mobile))my=true;
if (!my)
{
//document.mobileform.MobileNum.value='';
alert('对不起，您输入的手机号码有错误。');
//document.mobileform.MobileNum.focus();
return false;
}

return true; 

} 