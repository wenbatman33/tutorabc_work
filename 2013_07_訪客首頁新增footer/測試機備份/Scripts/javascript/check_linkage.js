

$(document).ready(function ()
{
	$("form").submit(function(){
		return checkForm();
	})
})

function checkForm()
{
	// 取出各節點
	//alert("DD");
    var str_cname = $("#name").val();		//姓名
	var str_cphone = $("#cphone").val();	//手機
	var str_gphone = $("#gphone").val();	//固定電話
	var str_gphone12 = $("#gphone12").val();//区号
	var str_email = $("#mail").val();		//email
	var str_age = $("#age_area").val();		//年齡	
	var str_sel_province = $("#sel_province").val();	//省份
	var str_sel_area = $("#sel_area").val(); //地區

    //檢驗中文姓名格式
	if (g_obj_verify.isNoneParticularCharacter(str_cname) == g_obj_error_handle.error_code.input_data_is_empty)
    {
        alert("Hi,为了尽速为您服务,请您提供姓名");
		$("#name").focus();
        return false;
    }
    if (g_obj_verify.isNoneParticularCharacter(str_cname) != g_obj_error_handle.success_code) 
    {
        alert("Hi,为了尽速为您服务,请您提供姓名");
		$("#name").focus();
        return false;
    }

	//檢驗手機格式
    if(g_obj_verify.isPhone(str_cphone,7,20) == g_obj_error_handle.error_code.input_data_is_empty)
    {
        alert("请您放心,您的信息受到绝对的保护,\n请提供最常使用的手机,顾问将尽速与您联系");
		$("#cphone").focus();
        return false;     
    }
	if(g_obj_verify.isPhone(str_cphone,7,20) != g_obj_error_handle.success_code )
    {
        alert("很抱歉,您所提供的手机格式不正确,\n请重新输入,顾问将尽速与您联系");
		$("#cphone").focus();
        return false;
    }
    // 20130124 阿捨新增 大陸手機 or 台灣手機 判斷
    if ("1" != str_cphone.substr(0, 1) && "0" != str_cphone.substr(0, 1)) {
			alert("很抱歉,您所提供的手机格式不正确,\n请重新输入,顾问将尽速与您联系");
			$("#cphone").focus();
            return false;
		}
	else if ("1" == str_cphone.substr(0, 1)) {
        if (g_obj_verify.isCellPhone(str_cphone, 1) != g_obj_error_handle.success_code) {
            alert("很抱歉,您所提供的手机格式不正确,\n请重新输入,顾问将尽速与您联系");
            $("#cphone").focus();
            return false;
        }
    }
    else if ("0" == str_cphone.substr(0, 1)) {
        if (g_obj_verify.isCellPhone(str_cphone, 2) != g_obj_error_handle.success_code) {
            alert("很抱歉,您所提供的手机格式不正确,\n请重新输入,顾问将尽速与您联系");
            $("#cphone").focus();
            return false;
        }
		
    } 
    //VM13041807檢驗固定電話	区号和电话号码 2013/4/22 paul
	if (g_obj_verify.isPhone(str_gphone12, 7, 20, true) != g_obj_error_handle.error_code.input_data_is_empty) {
        if (g_obj_verify.isPhone(str_gphone12, 3, 20, true) != g_obj_error_handle.success_code) {
            alert("请输入正确的区号！");
           $("#gphone12").focus();
            return false;
        }
    }
		if(str_gphone12 != "" && str_gphone == ""){
			alert("请输入正确的电话！");
	        $("#gphone").focus();
            return false;
	}
    if (g_obj_verify.isPhone(str_gphone, 7, 20, true) != g_obj_error_handle.error_code.input_data_is_empty) {
        if (g_obj_verify.isPhone(str_gphone, 7, 20, true) != g_obj_error_handle.success_code) {
            alert("请输入正确的电话！");
            $("#gphone").focus();
            return false;
        }
    }
	if(str_gphone12 == "" && str_gphone != ""){
			alert("请输入正确的区号！");
	        $("#gphone12").focus();
            return false;
	}
    //檢驗email格式
    if (g_obj_verify.isEmailAddressValid(str_email) == g_obj_error_handle.error_code.input_data_is_empty)
    {
        alert("请输入您最常用的邮箱,\n您的信息受到绝对的保护");
        $("#mail").focus();
        return false;
    }
    if (g_obj_verify.isEmailAddressValid(str_email) != g_obj_error_handle.success_code)
    {
        alert("很抱歉,您所提供的邮箱不正确,\n您将无法收取信件,请重新输入");
        $("#mail").focus();
        return false;
    }


    //驗證年齡/**/
    if (str_age == "")
    {
        alert("您尚未选择年龄");
        $("#age_area").focus();
        return false;
    }

    
    //驗證现居城市/**/
	if ( str_sel_province == ""|| str_sel_area == "")
	{
		alert("您尚未选择现居城市");
		if (str_sel_province == "") {		//沒有選省分
			$("#sel_province").focus();
			return false;
		} 
		if (str_sel_area == "") {	//沒有選地區
			$("#sel_area").focus();
			return false;
		} 
	}


    if (str_cname != "") {
        $("#name").val(escape(str_cname));
    }

    if (str_cphone != "") {
        $("#cphone").val(escape(str_cphone));
    }

    if (str_gphone != "") {
        $("#gphone").val(escape(str_gphone));
    }

    if (str_email != "") {
        $("#mail").val(escape(str_email));
    }

    if (str_age != "") {
        $("#age_area").val(escape(str_age));
    }

    if (str_sel_province != "") {
        $("#sel_province").val(escape(str_sel_province));
    }

    if (str_sel_area != "") {
        $("#sel_area").val(escape(str_sel_area));
    }
    return true;
}	