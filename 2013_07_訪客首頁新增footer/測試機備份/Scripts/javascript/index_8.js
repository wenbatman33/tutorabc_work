//$(document).ready(function () {
//    $("form").submit(function () {
//        return checkForm();
//    })
//})

//檢核
function checkForm() {
    // 取出各節點
    var str_cname = $("#txt_cname").val();
    var str_nation_code = $("#sel_nation_code").val();
    var str_email = $("#txt_email_addr").val();
    var str_contact_time = $("#txt_get_sdate").val();
    var str_contact_clock = $("#sel_order_time").val();
    var str_cphone = $("#txt_cphone").val();
    var str_gphone = $("#txt_gphone").val();
    var str_know_abc = $("#rad_know_abc").val();
    //	var this_day = new Date();
    //    var current_hours = this_day.getHours();
    //    var select_demo_hours = str_contact_clock.substring(6,8);
    //檢驗姓名格式
    if (g_obj_verify.isChineseNameValid(str_cname) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert("请输入您的姓名");
        $("#txt_cname").focus();
        return false;
    }
    if (g_obj_verify.isChineseNameValid(str_cname) != g_obj_error_handle.success_code) {
        alert("您的姓名格式错误");
        $("#txt_cname").focus();
        return false;
    }
    if (str_cname == "请输入您的姓名") {
        alert("请输入您的姓名");
        $("#txt_cname").focus();
        return false;
    }
    //檢驗email格式    
    if (g_obj_verify.isEmailAddressValid(str_email) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert("请输入您的邮箱");
        $("#txt_email_addr").focus();
        return false;
    }
    if (g_obj_verify.isEmailAddressValid(str_email) != g_obj_error_handle.success_code) {
        alert("您输入的邮箱格式不正确");
        $("#txt_email_addr").focus();
        return false;
    }
    if (str_email == "请输入您的邮箱") {
        alert("请输入您的邮箱");
        $("#txt_email_addr").focus();
        return false;
    }
    //检验所在地
    if (str_nation_code == "0") {
        alert("请选择所在地");
        $("#sel_nation_code").focus();
        return false;
    }
    //檢驗手機格式
    if (g_obj_verify.isPhone(str_cphone, 10, 20) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert("请输入您的移动电话");
        $("#txt_cphone").focus();
        return false;
    }
    if (g_obj_verify.isPhone(str_cphone, 10, 20) != g_obj_error_handle.success_code) {
        alert("您输入的移动电话格式不正确");
        $("#txt_cphone").focus();
        return false;
    }
    if (str_cphone == "请输入移动电话") {
        alert("请输入移动电话");
        $("#txt_cphone").focus();
        return false;
    }
    // 20130124 阿捨新增 大陸手機 or 台灣手機 判斷
    if ("+86" == $("#hdn_phone_nationcode").val()) {
        if (g_obj_verify.isCellPhone(str_cphone, 1) != g_obj_error_handle.success_code) {
            alert("您输入的移动电话格式不正确");
            $("#txt_cphone").focus();
            return false;
        }
    }
    else if ("+886" == $("#hdn_phone_nationcode").val()) {
        if (g_obj_verify.isCellPhone(str_cphone, 2) != g_obj_error_handle.success_code) {
            alert("您输入的移动电话格式不正确");
            $("#txt_cphone").focus();
            return false;
        }
    }
    //驗證在線試讀的時間
    if (str_contact_time == "学习服务日期") {
        alert("请选择服务日期");
        $("#txt_get_sdate").focus();
        return false;
    }
    if (str_contact_clock == "0") {
        alert("请选择服务时间");
        $("#sel_order_time").focus();
        return false;
    }
    //如何得知VIPABC验证
    if (str_know_abc == "0") {
        alert("请选择从何处得知VIPABC");
        $("#rad_know_abc").focus();
        return false;
    }
    //年龄验证
    //	if(str_age == ""){
    //		alert(PLZ_INPUT_AGE);
    //		return false;
    //	}
    //    if( $("#rad_know_abc").val() == "" )
    //    {
    //        alert(KNOW_WEBSITE);
    //		$("#rad_know_abc").focus();
    //        return false;
    //    }
    return true;
}

