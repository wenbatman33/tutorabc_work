function checkForm() {
    // 取出各節點
    var str_cname = $("#txt_cname").val();
    var str_email = $("#txt_email_addr").val();
    var str_password = $("#pwd_passwd").val();
    var str_password_retry = $("#pwd_passwd_retry").val();
    var str_ename = $("#txt_ename").val();

    var str_age = $("#age_area").val();

    var str_sel_messenger = $("#sel_messenger").val();
    var str_birth = $("#txt_birth").val();
    var str_messenger_addr = $("#txt_messenger_addr").val();
    var str_contact_time = $("#sel_contact_time").val();
    var str_country_code = $("#hdn_phone_nationcode").val();
    var str_cphone = $("#txt_cphone").val();
    var str_gphone_code = $("#txt_gphone_code").val();
    var str_gphone = $("#txt_gphone").val();
    var str_gphone_ext = $("#txt_gphone_ext").val();
    var str_sel_contact_us_qtype = $("#sel_contact_us_qtype").val();
    var str_ver = $("#txt_verify_code").val();
    var str_session_ver = $("#ver_session").val();
    //var str_serial_numbers = $("#txt_promotional_serial_numbers").val();

    //檢驗email格式    
    if (g_obj_verify.isEmailAddressValid(str_email) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(CONTACT_US_EMAILPLEASE);
        $("#txt_email_addr").focus();
        return false;
    }
    if (g_obj_verify.isEmailAddressValid(str_email) != g_obj_error_handle.success_code) {
        alert(CONTACT_US_EMAILERROR);
        $("#txt_email_addr").focus();
        return false;
    }
    //驗證密碼
    if (g_obj_verify.isPasswordValid(str_password, 6, 20) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(PWS_INPUT);
        $("#pwd_passwd").focus();
        return false;
    }
    if (g_obj_verify.isPasswordValid(str_password, 6, 20) != g_obj_error_handle.success_code) {
        alert(PWS_INPUT_ERROR);
        $("#pwd_passwd").focus();
        return false;
    }
    //驗證密碼確認
    if (g_obj_verify.isPasswordConfirmValid(str_password, str_password_retry) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(PWS_CONFIRM_INPUT);
        $("#pwd_passwd_retry").focus();
        return false;
    }
    if (g_obj_verify.isPasswordConfirmValid(str_password, str_password_retry) != g_obj_error_handle.success_code) {
        alert(PWS_CONFIRM_FAIL);
        $("#pwd_passwd_retry").focus();
        return false;
    }
    //檢驗姓名格式
    if (g_obj_verify.isChineseNameValid(str_cname) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(CONTACT_US_NAMEPLEASE);
        $("#txt_cname").focus();
        return false;
    }
    if (g_obj_verify.isChineseNameValid(str_cname) != g_obj_error_handle.success_code) {
        alert(CONTACT_US_NAMEERROR);
        $("#txt_cname").focus();
        return false;
    }
    //檢驗性別

    if ($("input[name='chk_gender']:checked").length == 0) {
        alert(GENDER_INPUT);
        $("#chk_gender").focus();
        return false;
    }
    //驗證英文姓名
    if (g_obj_verify.isEnglishNameValid(str_ename) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(ENAME_INPUT);
        $("#txt_ename").focus();
        return false;
    }
    if (g_obj_verify.isEnglishNameValid(str_ename) != g_obj_error_handle.success_code) {
        alert(YOUR_ENAME_FORMAT_ERROR);
        $("#txt_ename").focus();
        return false;
    }
    //年龄验证
    if (str_age == "") {
        alert(PLZ_INPUT_AGE);
        return false;
    }

    //判斷選取的是手機還是市話
    if (document.getElementById('rad_mobile').checked) {
        //檢驗手機格式
        if (g_obj_verify.isPhone(str_cphone, 1, 20) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_country_code, 1, 10) == g_obj_error_handle.error_code.input_data_is_empty) {
            alert(CONTACT_US_MOBILEPLEASE);
            $("#txt_cphone").focus();
            return false;
        }
        if (g_obj_verify.isPhone(str_cphone, 1, 20) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_country_code, 1, 10) != g_obj_error_handle.success_code) {
            alert(CONTACT_US_MOBILEERROR);
            $("#txt_cphone").focus();
            return false;
        }
        // 20130124 阿捨新增 大陸手機 or 台灣手機 判斷
        if ("+86" == $("#hdn_phone_nationcode").val()) {
            if (g_obj_verify.isCellPhone(str_cphone, 1) != g_obj_error_handle.success_code) {
                alert(CONTACT_US_MOBILEERROR);
                $("#txt_cphone").focus();
                return false;
            }
        }
        else if ("+886" == $("#hdn_phone_nationcode").val()) {
            if (g_obj_verify.isCellPhone(str_cphone, 2) != g_obj_error_handle.success_code) {
                alert(CONTACT_US_MOBILEERROR);
                $("#txt_cphone").focus();
                return false;
            }
        }
    }
    else {
        //檢驗市話格式
        if (g_obj_verify.isPhone(str_country_code, 1, 10) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_gphone_code, 1, 10) == g_obj_error_handle.error_code.input_data_is_empty || g_obj_verify.isPhone(str_gphone, 1, 20) == g_obj_error_handle.error_code.input_data_is_empty) {
            alert(GROUND_PHONE_PLZ);
            $("#hdn_phone_nationcode").focus();
            return false;
        }
        if (g_obj_verify.isPhone(str_country_code, 1, 10) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone_code, 1, 10) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone, 1, 20) != g_obj_error_handle.success_code || g_obj_verify.isPhone(str_gphone_ext, 1, 10) != g_obj_error_handle.success_code) {
            alert(GROUND_PHONE_ERROR);
            $("#hdn_phone_nationcode").focus();
            return false;
        }
    }

    //驗證出生日期
    if (g_obj_verify.isBirthday(str_birth, true) == g_obj_error_handle.error_code.input_data_is_empty) {
        alert(BIRTHDAY_INPUT);
        $("#txt_birth").focus();
        return false;
    }
    if (g_obj_verify.isBirthday(str_birth, true) != g_obj_error_handle.success_code) {
        alert(BIRTHDAY_INPUT_FAIL);
        $("#txt_birth").focus();
        return false;
    }
    //驗證msn、qq帳號
    if ($("#sel_messenger").val() == "QQ") {
        if (g_obj_verify.isQQid(str_messenger_addr, true) == g_obj_error_handle.error_code.input_data_is_empty) {
            alert(QQ_INPUT);
            $("#txt_messenger_addr").focus();
            return false;
        }
        if (g_obj_verify.isQQid(str_messenger_addr, true) != g_obj_error_handle.success_code) {
            alert(QQ_INPUT_FAIL);
            $("#txt_messenger_addr").focus();
            return false;
        }
    }
    if ($("#sel_messenger").val() == "MSN") {
        if (g_obj_verify.isMsnValid(str_messenger_addr, true) == g_obj_error_handle.error_code.input_data_is_empty) {
            alert(MSN_INPUT);
            $("#txt_messenger_addr").focus();
            return false;
        }
        if (g_obj_verify.isMsnValid(str_messenger_addr, true) != g_obj_error_handle.success_code) {
            alert(MSN_INPUT_FAIL);
            $("#txt_messenger_addr").focus();
            return false;
        }
    }
    //檢驗性別

    if ($("input[name='rad_know_abc']:checked").length == 0) {
        alert(KNOW_WEBSITE);
        $("#chk_gender").focus();
        return false;
    }
    //驗證促销活动序号
    /*
    if(str_serial_numbers != "" && !checkPromotionalSerialNumbers(str_serial_numbers))
    {
    alert("<%=getWord("INPUT_PROMOTIONAL_CAMPAIGN_SERIAL_NUMBERS_FORMAT_ERROR")%>");
    $("#txt_promotional_serial_numbers").focus();
    return false;
    }
    */
    //驗證我已阅读并同意本站的使用
    if ($("input[name='check_agree_use']:checked").length == 0) {
        alert(CONFIRM_ARTICLE);
        $("#txt_email_addr").focus();
        return false;
    }
    //驗證驗證碼輸入正確與否
    checkCaptcha();
    return false;
    //驗證欄位
}
var haha;
//captcha檢核 驗證碼		
function checkCaptcha() {
    $.ajaxSetup({
        async: false
    });
    $.getJSON("http://www.vipabc.com/lib/functions/functions_captcha.asp?validateCaptchaCode=" + $("#captchacode").val() + "&format=json&jsoncallback=?", function (data) {
        if (data.status == "1") {
            //$("frm_regist").submit();
            haha = 1;
            return true;
        }
        else {
            reFreshImage("imgCaptcha");
            //alert(CAPCHRA_WRONG);
            //            $("#captchacode").focus();
            haha = 2;
            return false;
            //alert(captcha);
        }
    });
}
function checkPromotionalSerialNumbers(str_input_serial_numbers) {
    //先寫死
    var bol_flag_promotional_serial_numbers

    if (str_input_serial_numbers == "ILOVEVIPABC" || str_input_serial_numbers == "IAMVIP" || str_input_serial_numbers == "iamvip" || str_input_serial_numbers == "I am vip" || str_input_serial_numbers == "i am vip") {
        bol_flag_promotional_serial_numbers = true;
    }
    else {
        bol_flag_promotional_serial_numbers = false;
    }
    return bol_flag_promotional_serial_numbers;
}
function checkEmailIsMember(p_str_email) {

    var strPostData = "get_score=chkemail&get_email_addr=" + p_str_email + "&get_login_password=";
    aurl = "../member/ajax_member_login.asp"; //call 的頁面
    var str_return_value = -1;
    $.ajax({
        type: "POST",
        cache: false,
        url: aurl,
        data: strPostData,

        success: function (msg) {
            if (msg == "success") {
                //成功表示db有此筆資料
                document.getElementById('ismember').innerHTML = ACCOUNT_EXIST;
                str_return_value = 1;
            } else {
                document.getElementById('ismember').innerHTML = "";
                str_return_value = -1;
            }
        },
        error: function () {
            alert("Ajax Error!");
        }
    });
    return str_return_value;
}

$(document).ready(function () {
    //預設值：defaultDate:(new Date(new Date().getFullYear()-20+"/01/01")-new Date())/(1000*60*60*24 
    $("#txt_birth").datepicker({ showOtherMonths: true, maxDate: '0d', rangeSelect: false, changeFirstDay: false, dateFormat: 'yy/mm/dd', yearRange: '-100:0', defaultDate: '-30y' }).keydown(function () { return false; });
});