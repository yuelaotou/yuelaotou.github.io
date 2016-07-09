
function input_ok(theForm){
    var text="";

//验证用户名
    var M1 = theForm.JpegPath.value; 
    if (M1 == ""){text=text+"[上传路径] 不能为空；\n"} 

//验证密码
    var M2 = theForm.JpegType.value; 
    if (M2 == ""){text=text+"[上传格式] 不能为空；\n"} 

//验证年龄
    var M3 = theForm.Jpegsize.value; 
    if (M3 == "") {
		text=text+"[上传大小] 不能为空；\n"} 
	else {
		if (isNaN(M3)) {text=text+"[上传大小] 必须为数值；\n"}
	}

//验证年龄
    var M4 = theForm.JpegWidth.value; 
    if (M4 == "") {
		text=text+"[缩略图宽度] 不能为空；\n"} 
	else {
		if (isNaN(M4)) {text=text+"[缩略图宽度] 必须为数值；\n"}
	}

//验证年龄
    var M5 = theForm.JpegHeight.value; 
    if (M5 == "") {
		text=text+"[缩略图高度] 不能为空；\n"} 
	else {
		if (isNaN(M5)) {text=text+"[缩略图高度] 必须为数值；\n"}
	}

//验证年龄
    var M6 = theForm.ImageWidth.value; 
    if (M6 == "") {
		text=text+"[水印在图片上的X坐标] 不能为空；\n"} 
	else {
		if (isNaN(M6)) {text=text+"[水印在图片上的X坐标] 必须为数值；\n"}
	}

//验证年龄
    var M7 = theForm.ImageHeight.value; 
    if (M7 == "") {
		text=text+"[水印在图片上的Y坐标] 不能为空；\n"} 
	else {
		if (isNaN(M7)) {text=text+"[水印在图片上的Y坐标] 必须为数值；\n"}
	}

//验证姓名
    var M8 = theForm.Img.value; 
    if (M8 == ""){text=text+"[水印图标] 不能为空；\n"} 

//验证地址
    var M9 = theForm.ImageAlpha.value; 
    if (M9 == ""){text=text+"[水印透明度] 不能为空；\n"} 
	
//验证
    var M10 = theForm.fontWidth.value; 
    if (M10 == "") {
		text=text+"[文字在图片上的X坐标] 不能为空；\n"} 
	else {
		if (isNaN(M10)) {text=text+"[文字在图片上的X坐标] 必须为数值；\n"}
	}
	
//验证
    var M11 = theForm.fontHeight.value; 
    if (M11 == "") {
		text=text+"[文字在图片上的Y坐标] 不能为空；\n"} 
	else {
		if (isNaN(M11)) {text=text+"[文字在图片上的Y坐标] 必须为数值；\n"}
	}

//验证电话
    var M12 = theForm.FontColor.value; 
    if (M12 == ""){text=text+"[文字颜色] 不能为空；\n"} 
	
//验证密码提示问题
    var M13= theForm.fontText.value; 
    if (M13 == ""){text=text+"[文字内容] 不能为空；\n"} 

	
//向页面反馈信息	
//   if (text == "") {document.Upfile.submit();}
   if (text == "") {}   
   else {
    alert("出错提示：\n\n" + text);
    return false;
    }           
}