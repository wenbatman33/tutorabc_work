//验证邮箱格式
function isEmailCheck(mailID, errorID, textEmpty, type) {
    var email = document.getElementById(mailID).value;
    //檢驗email格式    
    if (g_obj_verify.isEmailAddressValid(email) == g_obj_error_handle.error_code.input_data_is_empty) {
        if (errorID == "") {
            alert(Tip(type, 1));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 1);
        }
        return false;
    }
    if (g_obj_verify.isEmailAddressValid(email) != g_obj_error_handle.success_code) {
        if (errorID == "") {
            alert(Tip(type, 2));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 2);
        }
        return false;
    }
    if (textEmpty != "" && email == textEmpty) {
        if (errorID == "") {
            alert(Tip(type, 1));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 1);
        }
        return false;
    }
    //正确的状况
    if (errorID != "") {
        document.getElementById(errorID).innerHTML = Tip(type, 3);
    }
    return true;
}
//验证邮箱格式
function isPasswordCheck(passwordID, errorID, type) {
    var pwd = document.getElementById(passwordID).value;
    //檢驗email格式    
    if (pwd == "") {
        if (errorID == "") {
            alert(Tip(type, 1));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 1);
        }
        return false;
    }
    //正确的状况
    if (errorID != "") {
        document.getElementById(errorID).innerHTML = Tip(type, 3);
    }
    return true;
}
//验证邮箱格式
function isPasswordCheck2(passwordID, passwordID2, errorID, type) {
    var pwd = document.getElementById(passwordID).value;
    var pwd2 = document.getElementById(passwordID2).value;
    //檢驗email格式    
    if (pwd2 == "") {
        if (errorID == "") {
            alert(Tip(type, 1));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 1);
        }
        return false;
    }
    if (pwd != pwd2) {
        if (errorID == "") {
            alert(Tip(type, 2));
        }
        else {
            document.getElementById(errorID).innerHTML = Tip(type, 2);
        }
        return false;
    }
    //正确的状况
    if (errorID != "") {
        document.getElementById(errorID).innerHTML = Tip(type, 3);
    }
    return true;
}
function Tip(type, error) {
    var tip = "";
    if (type == 1) {
        return "您输入的账号有误，请重新输入！";
    }
    if (type == 2) {
        if (error == 1) {
            return "<img src='/Content/images/error.png'/>";
        }
        if (error == 2) {
            return "<img src='/Content/images/error.png'/>";
        }
        if (error == 3) {
            return "<img src='/Content/images/ok.png'/>";
        }
    }

}