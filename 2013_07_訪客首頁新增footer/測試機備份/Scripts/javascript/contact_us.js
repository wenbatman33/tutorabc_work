//function checkForm()
//{
//    // 取出各節點
//    var str_cname = $.trim($("#txt_cname").val());
//    var str_email = $.trim($("#txt_email_addr").val());
//    var str_contact_time = $.trim($("#sel_contact_time").val());
//	var str_country_code = $.trim($("#hdn_phone_nationcode").val());
//    var str_cphone = $.trim($("#txt_cphone").val());
//    var str_gphone_code = $.trim($("#txt_gphone_code").val());
//    var str_gphone = $.trim($("#txt_gphone").val());
//    var str_gphone_ext = $.trim($("#txt_gphone_ext").val());
//    var str_sel_contact_us_qtype = $.trim($("#sel_contact_us_qtype").val());
//    var str_comment = $.trim($("#textarea_comment").val());
//    //檢驗姓名格式
//    if (g_obj_verify.isChineseNameValid(str_cname) == g_obj_error_handle.error_code.input_data_is_empty)
//    {
//        alert(CONTACT_US_NAMEPLEASE);
//		$("#txt_cname").focus();
//        return false;
//    }
//	if (g_obj_verify.isChineseNameValid(str_cname) != g_obj_error_handle.success_code ) 
//    {
//        alert(CONTACT_US_NAMEERROR);
//		$("#txt_cname").focus();
//        return false;
//    }
//    
//    //判斷選取的是手機還是市話
//    if (document.getElementById('rad_mobile').checked)
//	{ 
//	    //檢驗手機格式
//	    if (g_obj_verify.isPhone(str_cphone,1,20) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_country_code,1,10) == g_obj_error_handle.error_code.input_data_is_empty)
//	    {
//	        alert(CONTACT_US_MOBILEPLEASE);
//			$("#txt_cphone").focus();
//            return false;     
//	    }
//		if (g_obj_verify.isPhone(str_cphone,1,20) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_country_code,1,10) != g_obj_error_handle.success_code )
//	    {
//	        alert(CONTACT_US_MOBILEERROR);
//			$("#txt_cphone").focus();
//            return false;     
//	    }
//	    
//	}
//	else
//	{
//	    //檢驗市話格式
//	     if (g_obj_verify.isPhone(str_country_code,1,10) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_gphone_code,1,10) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_gphone,1,20) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_gphone_ext,1,10) == g_obj_error_handle.error_code.input_data_is_empty)
//	    {
//	        alert(GROUND_PHONE_PLZ);
//			$("#hdn_phone_nationcode").focus();
//            return false;     
//	    }
//		if (g_obj_verify.isPhone(str_country_code,1,10) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone_code,1,10) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone,1,20) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone_ext,1,10) != g_obj_error_handle.success_code)
//	    {
//	        alert(GROUND_PHONE_ERROR);
//			$("#hdn_phone_nationcode").focus();
//            return false;     
//	    }
//	   
//	}  
//	
//	//檢驗email格式    
//    if (g_obj_verify.isEmailAddressValid(str_email) == g_obj_error_handle.error_code.input_data_is_empty)
//    {
//        alert(CONTACT_US_EMAILPLEASE);
//		$("#txt_email_addr").focus();
//        return false;
//    }
//	if (g_obj_verify.isEmailAddressValid(str_email) != g_obj_error_handle.success_code)
//    {
//        alert(CONTACT_US_EMAILERROR);
//		$("#txt_email_addr").focus();
//        return false;
//    }	
//	
//	//檢驗問題類型
//	
//	if (str_sel_contact_us_qtype == "")
//	{
//	    alert(QUESTIONT_PYEPLZ);
//		$("#sel_contact_us_qtype").focus();
//	    return false;
//	}  
//	//檢驗建議內容
//	if (str_comment == "")
//	{
//	    alert(QUESTION_TYPE_EMPTY);
//		$("#sel_contact_us_qtype").focus();
//	    return false;
//	}
//	if (str_comment.length >200)
//	{
//	    alert(QUESTION_TYPE_TOO_LONG);
//		$("#sel_contact_us_qtype").focus();
//	    return false;
//	}
//	
//	return true;
//}
//當sel_contact_us_qtype改變時，紀錄選擇的欄位名稱
function change_select_value()
{   
    $("#q_type_value").val($("#sel_contact_us_qtype :selected").text()); 
}