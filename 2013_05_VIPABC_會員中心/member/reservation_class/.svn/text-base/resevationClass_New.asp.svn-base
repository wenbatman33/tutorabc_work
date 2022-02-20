<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class_order_class.asp"-->
<script type="text/javascript" src="javascript/resevationClass_New.js"></script>

<%
	'檢查客戶是否有登入
	'檢查是否在退費中 若在退費中且沒有堂數的話 不能訂課
	'by <!--#include file="include/include_prepare_and_check_data.inc"-->
	
	'使用合約開始/結束日判斷是否能訂課
	'參考 program\member\reservation_class\include\ajax_reservation_normal_and_vip_class.asp
	'str_service_start_date, str_service_end_date (include_prepare_and_check_data.inc.inc內的data)
	
	'設定 // 可選擇合約開始日起一個月內的時間 (使用在日曆(datepicker)上)
	
	' 可選擇合約開始日起一個月內的時間
		'if (dat_contract_start_date > dat_now_sys_datetime)如果現在是合約開始前
	' 或是 首次上課客戶且有預約First Session，預約的日期和時間
		'else if (g_obj_class_data.str_last_first_session_datetime != "-1" && 
		'		  g_obj_class_data.str_last_first_session_datetime != "-2")
	' 或   選擇今日起一個月內的時間
		'// 若系統日期>合約結束日 則 alert 合约于YYYY/MM/DD到期请，至 ｢学习纪录｣ 查询详细资讯。
		'// 合約已到期
        '    if (str_contract_end_date == "" || dat_now_sys_datetime > dat_contract_end_date)
		'		// 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
		'// 系統日期+30天 > 合約結束日
        '    else if (dat_contract_end_date < dat_now_sys_datetime_added)
		
	'取得可供客戶訂課的時間	ajax_class_info.asp
	'// 載入可供訂課的諮詢時間
		'$("#li_reservation_class_time").load(g_obj_reservation_class_ctrl.arr_page_url.ajax_class_info, {ndate: new Date(), act : 4, rcdate : str_class_date, rctype: str_class_type});
		
	'functions_reservation_class.asp裡面的getHtmlSelectClassTime會回傳HTML控制項<select> 可供客戶訂課的時間
	
	
	
	'=== aboochen CS09011902  開放合約起始日前7天,預訂合約開始當天全時段的課程 開始===
%>
<script type="text/javascript">
    $(function ()
    {
        g_obj_resevation_class_calender.refreshCalender("2010","6","14");
    });
</script>
<!--內容start-->
<div class="main_membox">
	<%
		'在這個訂畫面面記錄課程類別
		Dim str_reservation_class_type : str_reservation_class_type = Request("class_type")
	%>
	<input type="hidden" name="hd_reservation_class_type" id="hd_reservation_class_type" value="<%=str_reservation_class_type%>" />
	
	<!--說明start-->
	<div class="top_class6">
		<h1>
			预约 / 取消课程时请注意<br />
			<span class="org">1.新增预约课程</span>：课程开始的<span class="org">24</span>小时前完成，收取<span class="org">1</span>课时数费用<br />
			<span class="org">2.取消预约课程</span>：需在课程开始的<span class="org">4</span>小时前完成<br /><span class="org">3.临时预约24小时之内课程</span><br />
			(1)开课前24小时~开课前12小时预约：收取1.25课时数费用<br />
			(2)开课前12小时~开课前4小时预约：收取1.5课时数费用<br />
			(3)开课前4小时~开课前2小时预约：收取2课时数费用，且无法取消课程
		</h1>
		<h2><img src="/images/images_cn/b_ps.gif" width="15" height="16" />您目前预订的是<span class="ps">一般课程/小班(1-3人)</span></h2>
		<h3>截至目前为止，您尚可订位<span class="org"><%=getCustomerOneManyBookingCount(g_var_client_sn, CONST_VIPABC_RW_CONN)%></span>课时咨询</h3>
		<br class="clear">
	</div>   
	<!--說明end-->

	<!---->
	<div class="book">
		<div class="btop">
			<div id="div_calender_head">
            	<div align="center">
					<img id="img_waiting" src="/lib/javascript/JQuery/images/ajax-loader.gif" width="25" height="25" border="0" />
                </div>
			</div>
		</div>
		<div id= "div_calender_day">
		</div>
		
		<div class="t-right m-top15">
			<form method="post" name="res_class_form" action="reservation_class_exe.asp">
			<input type="hidden" name="hdn_total_reservation_class_dt" id="hdn_total_reservation_class_dt" /><!--客戶訂課資訊-->
			<input type="hidden" name="hdn_order_or_reorder_type" id="hdn_order_or_reorder_type" value="1" /><!--1:新預定的  2:重訂課程  3:取消課程-->
			<input type="button" name="button" id="button" value="确定订课" class="btn_1" onclick="javascript:g_obj_resevation_class_calender.reservateClass('<%=str_reservation_class_type%>'); return false;"/>
			</form>
		</div>
	</div>
	<!---->

	<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->