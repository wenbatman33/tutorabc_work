<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<% 
dim bolDeBugMode : bolDeBugMode = false '除錯模式
%>
<script type="text/javascript">
var please_trim_chinese_name = "<%=getWord("please_trim_chinese_name")%>";
var CONTACT_US_NAMEERROR = "<%=getWord("CONTACT_US_NAMEERROR")%>";
var please_trim_ename = "<%=getWord("please_trim_ename")%>";
var your_ename_format_error = "<%=getWord("your_ename_format_error")%>";
var please_trim_fname = "<%=getWord("please_trim_fname")%>";
var your_fname_format_error = "<%=getWord("your_fname_format_error")%>";
var please_select_your_sex = "<%=getWord("please_select_your_sex")%>";
var BIRTHDAY_INPUT = "<%=getWord("BIRTHDAY_INPUT")%>";
var BIRTHDAY_INPUT_FAIL = "<%=getWord("BIRTHDAY_INPUT_FAIL")%>";
var FILL_IN_MAIL_ADDRESS_PLZ = "<%=getWord("FILL_IN_MAIL_ADDRESS_PLZ")%>";
var CONTACT_US_MOBILEPLEASE = "<%=getWord("CONTACT_US_MOBILEPLEASE")%>";
var CONTACT_US_MOBILEERROR = "<%=getWord("CONTACT_US_MOBILEERROR")%>";
var CONTACT_US_MOBIL_PHONE_format_error = "<%=getWord("CONTACT_US_MOBIL_PHONE")%><%=getWord("format_error")%>";
var GROUND_PHONE_PLZ = "<%=getWord("GROUND_PHONE_PLZ")%> ";
var GROUND_PHONE_ERROR = "<%=getWord("GROUND_PHONE_ERROR")%>";
var please_input_office_tel = "<%=getWord("please_input_office_tel")%>";
var OFFICE_PHONE_FORMAT_ERROR = "<%=getWord("OFFICE_PHONE")%><%=getWord("FORMAT_ERROR")%>";
<!--
//-->
</script>
<script type='text/javascript' src="/lib/javascript/framepage/js/profile_1.js"></script>
		<form name="frm_profile" action="profile_exe.asp" method="post" onsubmit="return checkForm();">
<%
'判斷從什麼地方導過來的：LearningCard
Dim strEvent : strEvent = getRequest("event",CONST_DECODE_NO)
Dim strEventData : strEventData = getRequest("event_data",CONST_DECODE_NO)
'塞資料的相關設定值
Dim str_sql
Dim var_arr '傳excuteSqlStatementRead的陣列
Dim arr_result '接回來的陣列值
Dim str_tablecol 'table欄位
Dim str_tablevalue 'table值 
		
'profile date
Dim str_email 'email
Dim str_cname '中文姓名
Dim str_fname '英文名
Dim str_lname	 '英文姓
Dim str_sex '性別 2:女 1:男
Dim str_birth '生日
Dim str_clb_province '省份
Dim str_clb_area	 '地區
'20120918 阿捨新增 地址第三層
dim strCRegion : strCRegion = "" '地址第三層
Dim str_addr '地址
Dim str_zip_code '郵遞區號
		
Dim str_cphone_cycode '第一組手機國碼
Dim str_cphone '手機號碼
Dim str_cphone_cycode_2 '第二組手機國碼
Dim str_cphone_2 '第二組手機號碼
Dim str_htel_code '家用區碼
Dim str_htel '家用電話
Dim str_htel_branch '家用電話分機
Dim str_otel_code '公司電話區碼
Dim str_otel '公司電話
Dim str_otel_branch '公司電話分機
		
str_sql = " select sn,email,cname,fname,lname,sex,birth,clb_province,clb_area,zip_code,client_address, "
str_sql = str_sql & " cphone_code,cphone,clb_cphone_cycode_2,clb_cphone_2, " '手機部份
str_sql = str_sql & " htel_code,htel,htel_branch,otel_code,otel,otel_branch, " '市話
str_sql = str_sql & " clb_cphone_cycode, clb_region "		
str_sql = str_sql & " from client_basic where sn = @sn "

var_arr = Array(g_var_client_sn)
arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
if ( isSelectQuerySuccess(arr_result) ) then	
    if ( bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 客戶資訊:" & g_str_sql_statement_for_debug & "<br>")
    end if	   
    if ( Ubound(arr_result) > 0 ) then  
		'有資料 
		str_email = arr_result(1,0)
		str_cname = arr_result(2,0)
		str_fname = arr_result(3,0)
		str_lname = arr_result(4,0)
				
		str_sex = arr_result(5,0)
		str_birth = arr_result(6,0)
		str_clb_province 	= arr_result(7,0)
		str_clb_area = arr_result(8,0)
        '20120918 阿捨新增 地址第三層
        strCRegion = arr_result(22,0)
		str_zip_code = arr_result(9,0)
		str_addr = arr_result(10,0)
		str_cphone = arr_result(11,0)&arr_result(12,0)
		str_cphone_cycode_2 = arr_result(13,0)
		str_cphone_2 = arr_result(14,0)
		str_cphone_cycode	= arr_result(21,0)
				
		str_htel_code 	= arr_result(15,0)
		str_htel = arr_result(16,0)
		str_htel_branch = arr_result(17,0)
		str_otel_code = arr_result(18,0)
		str_otel = arr_result(19,0)
		str_otel_branch = arr_result(20,0)
	else
		'無資料
	end if 
end if
		
'---客戶資料 COM 物件
Dim float_user_data_complete_rate
Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)

if (int_member_result_exe = CONST_FUNC_EXE_SUCCESS) then        
	float_user_data_complete_rate = obj_member_opt_exe.getData(CONST_DATA_COMPLETE_RATE)
else
	'錯誤訊息
	Response.Write "member prepareData error = " & obj_member_opt_exe.PROCESS_STATE_RESULT & "<br>"
end if
set int_member_result_exe = nothing
set obj_member_opt_exe = nothing
%>
        <input type="hidden" id="event" name="event" value="<%=strEvent%>" />
        <input type="hidden" id="event_data" name="event_data" value="<%=strEventData%>" />
        <!--內容start-->

        <div class='page_title_5'><h2 class='page_title_h2'>编辑个人资料</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 

        <div id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>" class="main_membox">
        <!--標題start-->
        <div class="main_memtitle_line">
        <div class="main_memtitle_line_left"><%=getWord("PROFILE_MANAGEMENT")%></div>
        <div class="main_memtitle_line_right">
        <div class="main_memtitle_line_right_w"><%=getWord("DATA_COMPLETE_RATE")%>：<%=float_user_data_complete_rate%>%</div>
        </div>
        <div class="clear"></div>
        </div>
        <!--標題end-->
        <div>
        <!--mail start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("EMAIL_ACCOUNT")%></div>
        <div class="main_memfiled_right"><%=str_email%>&nbsp;&nbsp;(<a href="profile_password.asp"><%=getWord("MODIFY_PASSWORD")%></a>)</div>
        <% If (Request("hdn_get_password") = "true") then %>
        <!--091130新增--><div><span class="m-left10 wrong"><font color="#FF0000"><%=getWord("modify_password_success")%></font></span><!--091130新增--></div>
        <% end if  %>
        <div class="clear"></div>
        </div>
        <!--mail end-->

        <!--pw start-->
        <!--hdn_get_password<div class="main_memfiled_all">
        <div class="main_memfiled_left">修改密码</div>
        <div class="main_memfiled_right">
        <div class="main_memfiled_pw">***************</div>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>-->
        <!--pw end-->

        <!--name start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("CONTRACT_35")%></div>
        <div class="main_memfiled_right"><input type="text" id="txt_cname" name="txt_cname" class="login_box3 width1" value="<%=str_cname%>"></div>
        <div class="clear"></div>
        </div>
        <!--name end-->

        <!--name2 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("ENAME")%></div>
        <div class="main_memfiled_right"><input type="text" id="txt_fname" name="txt_fname" class="login_box3 width1" value="<%=str_fname%>"></div>
        <div class="clear"></div>
        </div>
        <!--name2 end-->

        <!--name3 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("FNAME")%></div>
        <div class="main_memfiled_right"><input type="text" id="txt_lname" name="txt_lname" class="login_box3 width1" value="<%=str_lname%>"></div>
        <div class="clear"></div>
        </div>
        <!--name3 end-->
        <!--性別 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("sex")%></div>

        <div class="main_memfiled_right"><input type="radio" name="chk_gender" value="1" <%if (str_sex = 1) Then Response.write "checked" end if %>/><%=getWord("MALE")%><input type="radio" name="chk_gender" value="2" <%if (str_sex = 2) Then Response.write "checked" end if %> /><%=getWord("FEMALE")%></div>
        <div class="clear"></div>
        </div>
        <!--性別 end-->

        <!--birth start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("BIRTHDAY")%></div>
        <div class="main_memfiled_right"><input type="text" id="txt_birth"  name="txt_birth" class="login_box3 width1" value="<%=str_birth%>"></div>
        <div class="clear"></div>
        </div>
        <!--birth end-->

        <!--通讯地址 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("COMMUNICATION_ADDRESS")%></div>
        <div id="div_address" class="main_memfiled_right">
			<!-- 20090109 加入include省份的資料 -->
        </div>
        <div class="clear"></div>
        </div>
        <!--通讯地址 end-->

        <!--移动电话 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("CONTACT_US_MOBIL_PHONE")%>1</div>
        <div class="main_memfiled_right">
        	<!--移动电话 1 start-->
            <%
			'----------- 國碼資料從xml出來，轉成陣列 ------- start -------
			'todo: 語言 
			Dim obj_node, obj_xml, lang, obj_name, str_selected ,str_selstring , str_selvalue
			Dim arr_selstring , arr_selvalue
			Dim int_sel_i
			Set obj_xml = New XML
				if (obj_xml.load(Server.Mappath("/xml/nation_code.xml"))) then
					for each obj_node in obj_xml.getXmlObj.selectNodes("/Nations/Country")
						str_selstring = str_selstring & obj_node.getAttribute("zh_cn_name") & "@" 
						str_selvalue = str_selvalue & obj_node.getAttribute("value") & "@" 
					next
					
				end if 
			Set obj_xml = Nothing
			
			if ( not isEmptyOrNull(str_selstring) ) then 
				arr_selstring = split(str_selstring,"@")
				arr_selvalue = split(str_selvalue,"@")
			end if 
			'----------- 國碼資料從xml出來，轉成陣列 ------- end ------- 
			%>
            <input name="hdn_txt_cphone_cycode" id="hdn_txt_cphone_cycode" type="hidden" value="<%=str_cphone_cycode%>" />
            <select name="select" id="sel_nation_code1" class="login_box4 width9 m-left5 m-right5" style="width:130px" onchange="changeNationalCode(this,'txt_cphone_cycode');">
            <option value="0" ><%=getWord("SELECT")%></option>
            <%
               if (isSelectQuerySuccess(arr_selstring)) then
			   		for int_sel_i = 0 to ubound(arr_selstring) -1
						str_selected = ""
						if ( arr_selvalue(int_sel_i) = "+"&str_cphone_cycode) then 
							str_selected = " selected=""selected"""
						end if 
				   %>
					   <option value="<%=arr_selvalue(int_sel_i)%>" <%=str_selected%> ><%=arr_selstring(int_sel_i)%></option>
				   <%
			   		next
			   end if 
            %>
            </select>     
            <input type="text" id="txt_cphone_cycode" name="txt_cphone_cycode" class="login_box3 width12 m-right5" value="<%="+"&str_cphone_cycode%>" disabled="disabled">-
            <input type="text" id="txt_cphone" name="txt_cphone" class="login_box3 width10 m-left5" value="<%=str_cphone%>">
        </div>
        <div class="clear"></div>
        </div>
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("CONTACT_US_MOBIL_PHONE")%>2</div>
        <div class="main_memfiled_right">
	   
        	<!--移动电话 2 start-->
            <input name="hdn_txt_cphone_cycode2" id="hdn_txt_cphone_cycode2" type="hidden" value="<%=str_cphone_cycode_2%>" />
            <select name="select" id="sel_nation_code2" class="login_box4 width9 m-left5 m-right5" style="width:130px" onchange="changeNationalCode(this,'txt_cphone_cycode2');">
            <option value="0" ><%=getWord("SELECT")%></option>
            <%
                if (isSelectQuerySuccess(arr_selstring)) then
			   		for int_sel_i = 0 to ubound(arr_selstring) -1
						str_selected = ""
						if ( arr_selvalue(int_sel_i) = "+"&str_cphone_cycode_2) then 
							str_selected = " selected=""selected"""
						end if 
				   %>
					   <option value="<%=arr_selvalue(int_sel_i)%>" <%=str_selected%>><%=arr_selstring(int_sel_i)%></option>
				   <%
			   		next
			   end if 
            %>
            </select>     
            <input type="text" id="txt_cphone_cycode2" name="txt_cphone_cycode2" class="login_box3 width12 m-right5" value="<%="+"&str_cphone_cycode_2%>" disabled="disabled">-
            <input type="text" id="txt_cphone2" name="txt_cphone2" class="login_box3 width10 m-left5" value="<%=str_cphone_2%>">
            
        </div>
        <div class="clear"></div>
        </div>
        <!--移动电话 end-->

        <!--住家电话 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("CONTACT_US_GROUND_PHONE")%></div>
        <div class="main_memfiled_right">
        <input type="text" id="txt_gphone_code" name="txt_gphone_code" class="login_box3 width12 m-right5" value="<%=str_htel_code%>">
        <input type="text" id="txt_gphone" name="txt_gphone" class="login_box3 width185 m-left5 m-right5" value="<%=str_htel%>">
         <%=getWord("ext")%><input type="text" id="txt_gphone_ext" name="txt_gphone_ext" class="login_box3 width12 m-right5" value="<%=str_htel_branch%>"></div>
        <div class="clear"></div>
        </div>
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("OFFICE_PHONE")%><br />
        </div>
        <div class="main_memfiled_right">
        <input type="text" id="txt_ophone_code" name="txt_ophone_code" class="login_box3 width12 m-right5" value="<%=str_otel_code%>">
        <input type="text" id="txt_ophone" name="txt_ophone" class="login_box3 width185 m-left5 m-right5" value="<%=str_otel%>">
         <%=getWord("ext")%><input type="text" id="txt_ophone_ext" name="txt_ophone_ext" class="login_box3 width12 m-right5" value="<%=str_otel_branch%>"></div>
        <div class="clear"></div>
        </div>
        <!--住家电话 end-->

      
        <div class="t-left l_margin" style='margin-top:20px'><input type="submit" name="button" id="button" value="+ <%=getWord("SEND_CONFIRM")%>" class="btn_1"/></div>
        </div>
        </div>
        </form>
        <!--內容end-->
<script type="text/javascript">
var str_clb_province = '<%=str_clb_province%>';
var str_clb_area = '<%=str_clb_area%>';
var str_zip_code = '<%=str_zip_code%>';
var strCRegion = '<%=strCRegion%>';
var str_addr = '<%=str_addr%>';
</script>
<script type='text/javascript' src="/lib/javascript/framepage/js/profile.js"></script>

                
