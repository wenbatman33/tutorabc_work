<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript" src="/lib/javascript/verify.js"></script>
<script type="text/javascript" src="/lib/javascript/error_handle.js"></script>
<script type="text/javascript">
	function checkForm()
	{
		// 取出各節點
		var str_password_retry = $("#pwd_passwd_retry").val();
		var str_passwd_second = $("#pwd_passwd_second").val();
		//驗證新密碼
		if (g_obj_verify.isPasswordValid(str_password_retry, 6, 20) == g_obj_error_handle.error_code.input_data_is_empty)
		{
			alert("<%=getWord("please_input_new_password")%>");
			return false;
		}
		if (g_obj_verify.isPasswordValid(str_password_retry, 6, 20) != g_obj_error_handle.success_code ) 
		{
			alert("<%=getWord("PWS_INPUT_ERROR")%>");
			return false;
		}
		//驗證密碼確認
		if (g_obj_verify.isPasswordConfirmValid(str_password_retry, str_passwd_second) == g_obj_error_handle.error_code.input_data_is_empty)
		{
			alert("<%=getWord("PWS_CONFIRM_INPUT")%>");
			return false;
		}
		if (g_obj_verify.isPasswordConfirmValid(str_password_retry, str_passwd_second) != g_obj_error_handle.success_code ) 
		{
			alert("<%=getWord("your_confirm_password_different")%>");
			return false;
		}	
		//驗證欄位
		return true;
	}
    //check login 是否存在
    function chkOrgPassword(p_str_password) {

		var strPostData ="get_email_addr=<%=session("client_email")%>&get_login_password="+p_str_password;
		aurl = "../member/ajax_member_login.asp?get_score=chkpassword" //call 的頁面	
        //alert(aurl);
        $.ajax({
            type: "POST",
            cache: false,
            url: aurl,
            data: strPostData,
            success: function(msg) {
                //alert("AJAX 成功"+unescape(msg));
                //alert(msg);
                if (msg == "success") {
                    document.getElementById('font_password_msg').innerHTML = "";
					//document.getElementById('font_password_msg').style.display = "none";
                } else {
					//$("font_password_msg").html("您提供的密碼不正確."); 
                    document.getElementById('font_password_msg').innerHTML = "<img src='../../images/images_cn/add/alerts_cion.jpg' width='14' height='16' hspace='2' vspace='2' border='0' align='absmiddle'/><%=getWord("your_password_error")%>";
                }
            },
            error: function() {
                alert("<%=getWord("AJAX_FAIL")%>");

            }
        });

        //return false;
    }
	
</script>



       <!--內容start-->
       <form name="frm_password" action="profile_password_exe.asp" method="post" onsubmit="return checkForm();">
        <div class="main_membox">

        <div class='page_title_5'><h2 class='page_title_h2'>修改密码</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 

        <!--標題start-->
        <div class="main_memtitle_line">
 
        <div class="main_memtitle_line_right">
        <div class="main_memtitle_line_right_w">
        
        <div class="main_memfiled_right">
        <!--修改密码请输入旧密码、确认新密码、请输入新密码-->
        <div class="main_memfiled_pw">
        <div class='profile_password'>
            <span class='profile_password_text'><%=getWord("please_old_password")%>:</span>
            <span class='profile_password_field'><input type="password" id="pwd_passwd" name="pwd_passwd" class="login_box3 width1" value="" onblur="chkOrgPassword(this.value)"></span>
        </div>
        <!--091130新增--><div id="font_password_msg" class="wrong"></div><!--091130新增-->
        
        <div class='profile_password'>
            <span class='profile_password_text'><%=getWord("please_new_password")%>:</span>
            <span class='profile_password_field'><input type="password" id="pwd_passwd_retry" name="pwd_passwd_retry" class="login_box3 width1 m-top5" value=""></span>
        </div>
        <!--091130新增--><div id="font_chkpassword_msg" class="wrong"></div><!--091130新增-->
        
        <div class='profile_password'>
            <span class='profile_password_text'><%=getWord("confirm_new_password")%>:</span>
            <span class='profile_password_field'><input type="password" id="pwd_passwd_second" name="pwd_passwd_second"  class="login_box3 width1 m-top5" value="" ></span>
        </div>
        </div>
        <div class="clear"></div>
        </div>
        
        </div>
        </div>
        <div class="clear"></div>
        </div>
        <!--標題end-->
        <div>
        
        
        <div class="profile_password_btns"><!--091130新增-->
        <input type="submit" name="button" id="button" value="+ <%=getWord("save")%>" class="btn_1"/>
        <input type="button" name="button" id="button_del" value="+ <%=getWord("CANCEL")%>" class="btn_1 m-left5" onclick="location.href='profile.asp'" /><!--091130新增--></div>
        </div>
        </div>
        
        <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
        </form>
        <!--內容end-->
        
