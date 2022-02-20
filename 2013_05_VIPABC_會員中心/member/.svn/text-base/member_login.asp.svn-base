<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/include/html/template/template_change_keyword.asp"-->
<%
    Dim str_tmp_login_account_value									'Cookies裡面記錄 account
    Dim str_tmp_login_password										'Cookies裡面記錄 password
    Dim str_tmp_remember_account_checked							'是否在Cookies記得 account
    Dim str_tmp_remember_password_checked							'是否在Cookies記得 password
	'response.write "test"&session("exporURL")
    if (Request.Cookies("remember_account")<>"") then
	    str_tmp_login_account_value = Request.Cookies("remember_account")
	    str_tmp_remember_account_checked = "checked"
    else
	    str_tmp_login_account_value = getWord("please_input_email")	'请输入您注册的E-Mail
    end if
    if (Request.Cookies("remember_password")<>"") then
	    str_tmp_login_password = Request.Cookies("remember_password")
	    str_tmp_remember_password_checked = "checked"
    else
	    str_tmp_login_password = ""
    end if
%>
<% 'VM12082820 好友推薦名單獲禮活動 - Constant接收該活動篇名，以利login後可直接導入活動頁 by Lily Lin 2012/10/08
	dim strConstant : strConstant = getRequest("Constant", CONST_DECODE_NO)
	'response.write strConstant
%>
<script type="text/javascript">
var please_input_email = '<%=getWord("please_input_email")%>';
var please_input_account = '<%=getWord("please_input_account")%>';
var PWS_INPUT = '<%=getWord("PWS_INPUT")%>';
var ajax_fail = '<%=getWord("ajax_fail")%>';
var password_email_sended = '<%=getWord("password_email_sended")%>';
var ACCOUNT_EXIST = '<%=getWord("ACCOUNT_EXIST")%>';
var ajax_error = '<%=getWord("ajax_error")%>!';
var QUICK_REGISTER_REMIND_MSG = "<%=getWord("QUICK_REGISTER_REMIND_MSG")%>";
<!--
	
//-->
</script>
<script type='text/javascript' language="javascript" src="/lib/javascript/framepage/js/member_login.js"></script>
<div id="send_mail"></div>
<!-- Menber Content Start -->
    <div>       
       <!--ADDED 內容start-->
		<div class="main_fun">
			<div class="loginbox" id="div_member_login" >
				<form name="frm_member_login" action="member_login_exe.asp?Constant=<%response.write strConstant%>&mgm=<%=request("mgm")%>&edm=<%=request("edm")%>" method="post" onsubmit="return checkLogin();">
				<div class="tte1 big color_1"><%=getWord("member_login")%></div><!--会员登录-->
				<div class="tte2"></div>
				
				<div class="inbox">
					<!--ajax錯誤訊息 start -->
					<div class="dbv" id="div_login_msg" style="display:none; margin-top:15px;">
						<div class="oopsbox">
							<div class="oops" id="div_login_msg_show"></div>
						</div>
					</div>
					 <!--ajax錯誤訊息 end -->  
									 
					<div>
					
						<div class="dbv">
							<span class="color_9">
								<%=getWord("ACCOUNT")%><!--账号-->
							</span>
							<input type="text" class="login_box" name="txt_login_account" id="txt_login_account" value="<%=str_tmp_login_account_value%>" /><br />
						
						</div>
						<div class="dbv">
							<%=getWord("REGISTER_MAIL_9")%><!--密码-->
							<input type="password" class="login_box" name="pwd_login_password" id="pwd_login_password" value="<%=str_tmp_login_password%>" style="margin-top:8px;" >
						</div>
						<div class="dbv2">
							<input type="checkbox" name="chk_remember_account" id="Checkbox1" value="Y" <%=str_tmp_remember_account_checked%> >
							<%=getWord("remember_my_account")%><!--记住我的Email账号-->
							<input type="checkbox" name="chk_remember_password" id="Checkbox2" value="Y" <%=str_tmp_remember_password_checked%> />
							<%=getWord("remember_my_password")%><!--记住我的密码-->
						</div>
						<div class="dbv">
							<span><a href="javascript:void(0);" onclick="showForgetPassword();">
								<%=getWord("forget_password")%>？<!--忘记密码-->
							</a></span>
							<!--登录-->
							<input type="submit" name="button" id="button" value="<%=getWord("login")%>" class="btn_1" >
						</div>
					</div>
					
				</div>
				<!--20100111 lucy搜尋那塊導向這個會員登入頁面, 並用這個hidden input紀錄將來要導到大會堂定課頁面或者是精選視頻頁面-->
				<input type="hidden" name="gobackurl" id="gobackurl" value="<%=getRequest("gobackurl", CONST_DECODE_NO)%>" />
				
				<div class="clear"></div>
				</form>
			</div>
        
		
			<div class="regbox" id="div_register_member" >
				<form name="frm_regist" action="/program/regist/regist_add.asp" method="post" onsubmit="return checkQuickMemberRegister();">
				<div class="tte">还不是会员？马上加入</div>
				<div class="dbv3">
					<%=getWord("REGISTER_ACCOUNT")%><!--账　　号-->
					<input type="text" class="login_box3 width2" id="txt_quick_register_email_addr" name="txt_email_addr" style="margin-top:7px;" value=""  onBlur="javascript:checkEmailIsMember(this.value);">
					<font id="ismember" color="#FF0000"></font>
				</div>
				<div class="dbv3">
					<%=getWord("REGISTER_PASSWORD")%><!--密　　码-->
					<input type="password" class="login_box3 width2" id="Text5" name="pwd_passwd" style="margin-top:7px;" value="">
				</div>
				<div class="dbv3">
					<%=getWord("REGISTER_CONFIRM_PASSWORD")%><!--确认密码-->
					<input type="password" class="login_box3 width2" id="Text6" name="pwd_passwd_retry" style="margin-top:7px;" value=""><br />
				</div>
				<div class="dbv3">
					<!--注册会员-->
					<input type="submit" name="button" id="Submit1" value="+ <%=getWord("reg_user")%>" class="btn_1"/>
				</div>
				<!--<div class="word">
					<span class="color_3">加入我们，您可以</span><br>
					<%=getWord("JOIN_US_AB")%>
				</div>-->
			
				<div class="clear"></div>
				</form>
			</div>
			
			<!--忘記密碼-->
			<span id="div_send_forget_password_mail" class="fun2_box" style="display:none;">
				<div class="bold color_1 big"><%=getWord("FORGET_PASSWORD")%><!--忘记密码--></div>
				<div class="reg_1">&nbsp;</div>
				<div class="reg_2">
					<div class="reg-forget">
						<%=getWord("PLEASE_INPUT_EMAIL_FOR_SEND_YOU_PASSWORD")%><!--请输入注册时的电子邮件地址，我们会将密码发送至您的电子邮箱--><br />
						<input type="text" class="login_box3 width1" name="txt_member_forget_password" id="txt_member_forget_password" />
						<input type="button" value="+ <%=getWord("send_password")%>" class="btn_1" onClick="return showErrorMsg('passwd','txt_member_forget_password','div_paw_msg');"/>
					</div>
				</div>
				<!--ajax錯誤訊息-->
				<div id="div_paw_msg" style="display:none">
					<div class="oopsbox">
						<div class="oops" id="div_paw_msg_show"></div>
					</div>
				</div>
				<!--ajax錯誤訊息-->
				<div class="clear"></div>
			</span>
			
			
			<!--重发认证信-->
			<span id="div_send_authorized_mail" class="fun2_box" style="display:none;">
				<div class="bold color_1 big"><%=getWord("AJAX_MEMBER_LOGIN_2")%><!--重发认证信--></div>
				<div class="reg_1"></div>
				<div class="reg_2">
					<div class="reg-forget">
						<%=getWord("PLEASE_INPUT_EMAIL_FOR_SEND_YOU_AUTHORIZED")%><!--请输入注册时的电子邮件地址，我们会重发认证信至您的电子邮箱--><br />
						<input type="text" class="login_box3 width1" name="txt_send_authenticate_mail" id="txt_send_authenticate_mail" />
						<!--重发认证信-->
						<input type="button" value="+ <%=getWord("AJAX_MEMBER_LOGIN_2")%>" class="btn_1" onClick="return showErrorMsg('authenticate','txt_send_authenticate_mail','div_authenticate_msg');"/>
					</div>
				</div>
				<!--ajax錯誤訊息-->
				<div id="div_authenticate_msg" style="display:none">
					<div class="oopsbox"><div class="oops" id="div_authenticate_msg_show"></div>
					</div>
				</div>
				<!--ajax錯誤訊息-->
				<div class="clear"></div>
			</span>
        </div>
        <!--ADDED 內容end-->
        <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
<!-- Main Content End -->
<script type="text/javascript">
    $(document).ready(function() {
        // 登入帳號
        $("#txt_login_account").Watermark("<%=str_tmp_login_account_value%>");
        // 快速註冊帳號
        // 请输入E-mail，此E-mail将成为您的会员登录帐号
        $("#txt_quick_register_email_addr").Watermark("<%=getWord("QUICK_REGISTER_REMIND_MSG")%>");	//请输入E-mail，此E-mail将成为您的会员登录帐号
    });
</script>
