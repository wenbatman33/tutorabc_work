<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
'請假期數,未實作, 之後有此資料時再讀入
int_leave_session = 0
%>
<!--#include virtual="/lib/include/vipabc_csp_include.asp"-->
<%
function newMod(ByVal firstNum,ByVal secondNum)
	newMod = cdbl(firstNum) - cdbl(secondNum) * (cdbl(firstNum) \ cdbl(secondNum))
end function

if int_leave_session > 0 then
	intContractSessionWeek = dbl_contract_session_week - int_leave_session
else
	intContractSessionWeek = dbl_contract_session_week
end if

'20121022 Beson 因合約最後一期有歸還的話, 會回到紅利包, 所以要重新計算紅利包的總堂數 Start
reCountMaxBonusPoint = 0
Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
sql = "SELECT  CONVERT(DECIMAL(18, 2), MAX(csh_bonus_points) / 65.) reCountMaxBonusPoint "
sql = sql & "FROM    dbo.customer_session_points_log WITH ( NOLOCK ) "
sql = sql & "WHERE   csh_client_sn = @csh_client_sn "
sql = sql & "        AND csh_purchase_sn = @csh_purchase_sn "
arrParam = Array(g_var_client_sn, str_purchase_account_sn)
intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_RW_CONN)
if not rs.eof then
	reCountMaxBonusPoint = rs("reCountMaxBonusPoint")
end if
rs.close
Set rs = nothing
'20121022 Beson 因合約最後一期有歸還的話, 會回到紅利包, 所以要重新計算紅利包的總堂數 End

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493			
if ( Instr(CONST_DIANPING_GROUPBUY_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
	intContractSessionWeek = 1
end if

'20120621 Beson 取得客戶課時已使用完畢的週期
if (cdbl(flt_available_session)+cdbl(flt_remain_week_session)+cdbl(flt_bonus_session)) = 0 and not isEmptyOrNull(str_purchase_account_sn) then
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
	sql = "SELECT TOP 1 "
	sql = sql & "        csh_cycle_week "
	sql = sql & "FROM    customer_session_points_log WITH (NOLOCK) "
	sql = sql & "WHERE   csh_client_sn = @csh_client_sn "
	sql = sql & "        AND csh_purchase_sn = @csh_purchase_sn "
	sql = sql & "        AND csh_total_points = 0 "
	sql = sql & "        AND csh_bonus_points = 0 "
	sql = sql & "        AND csh_week_points = 0 "
	sql = sql & "ORDER BY csh_sn "
	arrParam = Array(g_var_client_sn,str_purchase_account_sn)
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_VIPABC_RW_CONN)
	if not rs.eof then
		SessionOverCycle = rs("csh_cycle_week")
	end if
	rs.close
	Set rs = nothing
'20120621 Beson 大錢包已用完,取得當期應使用堂數
elseif cdbl(flt_available_session) = 0 and datediff("d",date,dateadd("d",int_account_cycle_day*dbl_used_session-1,datContractServiceBegin))>=0 and not isEmptyOrNull(str_purchase_account_sn) then
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
	sql = "SELECT "
	sql = sql & "        csp_week_start_points/65. csp_week_start_points "
	sql = sql & "FROM    customer_session_points WITH (NOLOCK) "
	sql = sql & "WHERE   csp_client_sn = @csp_client_sn "
	sql = sql & "        AND csp_purchase_sn = @csp_purchase_sn "
	arrParam = Array(g_var_client_sn,str_purchase_account_sn)
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_VIPABC_RW_CONN)
	if not rs.eof then
		csp_week_start_points = rs("csp_week_start_points")
	end if
	rs.close
	Set rs = nothing	
elseif cdbl(flt_available_session) > 0 then
	'先取得預約課時
	bookOfSvalueTotal = 0
	Set objRs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT SUM(svalue) bookOfSvalueTotal FROM client_attend_list WITH (NOLOCK) WHERE client_sn = @client_sn AND valid = 1 AND attend_room IS NULL"
	arrParam = Array(g_var_client_sn)
	intQueryResult = objRs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_RW_CONN)
	if not objRs.eof then
		bookOfSvalueTotal = objRs("bookOfSvalueTotal")
	end if
	objRs.close
	Set objRs  = Nothing
	'用來計算是否有超用的變數, 剩餘課時+每週應用課時x已用週期
	tempAvailableSession = cdbl(flt_available_session) + cdbl(flt_available_week_session) * cdbl(dbl_used_session) + bookOfSvalueTotal
	for i=1 to intContractSessionWeek
		if tempAvailableSession < 0 then '若有小於0的值出現, 代表有超用超過一週期
			if isEmptyOrNull(SessionOverCycle) then SessionOverCycle = i-1 '取得客戶課時已使用完畢的週期
		elseif tempAvailableSession > 0 and tempAvailableSession < flt_available_week_session and i <>intContractSessionWeek then '若有出現大於0,小於應用課時,又不是在最後一期的話
			csp_week_start_points = tempAvailableSession '這就是該期的應用課時
		end if
		tempAvailableSession = tempAvailableSession - cdbl(flt_available_week_session) '每週期減去應用課時
	next
end if

%>
<script type='text/javascript' src="/lib/javascript/framepage/js/learning_leaen.js"></script>
<script language="javascript">
$(document).ready(function() {
	
	//When page loads...
	$(".session_content").hide(); //Hide all content
	$("ul.sessions li#<%=dbl_used_session%>").addClass("active").show(); //Activate first session
	//$(".session_content#week<%=dbl_used_session%>").show(); //Show first session content

	//On Click Event
	$("ul.sessions li").click(function() {

		$("ul.sessions li").removeClass("active"); //Remove any "active" class
		$(this).addClass("active"); //Add "active" class to selected session
		$(".session_content").hide(); //Hide all session content

		var activesession = $(this).find("a").attr("id"); //Find the href attribute value to identify the active session + content
		$(activesession).show(); //Fade in the active ID content
		return false;
	});
	$("#left").click(function() {
		var first_week = parseInt(get_first_show_week())-1;
		var last_week = parseInt(get_last_show_week());
		var int_min_week = 0 ;
		$("#"+first_week).show();	
		if (first_week != int_min_week) //到達最小值不隱藏
		{
			$("#"+last_week).hide();
		}		
		return false;
	});
	$("#right").click(function() {
		var first_week = get_first_show_week();
		var last_week = parseInt(get_last_show_week())+1;
		var int_max_week = parseInt(<%=contractCycleAndRefundCycle + dbl_bonus_session_week%>)+1;
		$("#"+last_week).show();
		if (last_week != int_max_week) //到達最大值不隱藏
		{	
			$("#"+first_week).hide();
		}
		return false;
	});

	$(".showeturn").hover(function(){   
		$("#Consultantlesson").css("left",$(this).offset().left+50);
		$("#Consultantlesson").css("top",$(this).offset().top+30);  
		$("#Consultantlesson").show(500);      
	});   
	   
	$("#closeConsultantlesson").click(function(){      
		$("#Consultantlesson").hide(500);      
	});
	
	$(".showegift").hover(function(){
		$("#gitftlesson").css("left",$(this).offset().left+50);
		$("#gitftlesson").css("top",$(this).offset().top+30);  
		$("#gitftlesson").show(500);      
	});
	
	$("#closegitftlesson").click(function(){      
		$("#gitftlesson").hide(500);      
	}); 

	$("#showmore").click(function(){      
		$("#moreinfo").show(500);   
	});
});
</script>
<link href="/lib/javascript/framepage/css/learning_learn.css" rel="stylesheet" type="text/css" />
<!--內容start-->
<div class="main_membox">

	<!--上版位課程start-->

		<!--#include virtual="/lib/include/html/include_learning.asp"--> 
<%
response.redirect "learning_learn_new.asp"
Response.end
%>

	<!--上版位課程end-->
    <!--课程购买明细 start-->
	<div class="arrow_title"><%= getWord("class_buy_detail")%></div>
	<div>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" class="history_bd1">
                <tr>
                  <th>合同期数</th>
                  <th>每期</th>
                  <th class="rightLine2">合同起讫日期</th>
                  <th>合同</th>
                  <th class="rightLine3">备注</th>
                </tr>
                <tr>
					<td><%=intContractSessionWeek%>期</td>

					<%
						'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493			
						if ( Instr(CONST_DIANPING_GROUPBUY_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
						%>
							<td>付费大会堂2堂，一对三真人授课3堂</td>						
						<%
						else
						%>
							<td><%=flt_available_week_session%>课時</td>
						<%	
						end if
					%>					
					
                    <td><%=datPurchaseServiceBegin%>-<%=datPurchaseServiceEnd%></td>
                    <td><a target="_blank" href="/program/member/contract.asp?contract_sn=<%=int_contract_sn%>&cagree=1">[合同预览]</a></td>
                    <td class="rightLine3">&nbsp;</td>
                </tr>
          </table>
		<div class="clear"></div>
	</div>
    <!--课程购买明细 end-->
    <!--总课週期资讯 START -->
	<div class="arrow_title">总课周期资讯</div>
	<div>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="history_bd1">
            <tr>
			  <th width="12%">合同期数</th>
			  <th width="2%" rowspan="2" class="dd">+</th>
			  <th width="12%">归还课时<br />延展期数</th>
			  <th width="2%" rowspan="2" class="dd">+</th>
			  <th width="12%">赠送期数</th>
			  <th width="3%" rowspan="2" class="dd">+</th>
			  <th width="12%">请假期数</th>
			  <th width="3%" rowspan="2" class="dd">=</th>
			  <th width="12%">可使用期数</th>
			  <th width="3%" rowspan="2" class="dd">-</th>
			  <th width="12%">已使用期数</th>
			  <th width="3%" rowspan="2" class="dd">=</th>
			  <th width="12%" >未使用期数</th>
            </tr>
			  <td>
				<%=intContractSessionWeek%>期							      
			  </td>  <%'合同期數%>
			  <td class="showeturn"><a href="#" id="refund"></a></td><%'归还课时延展期数%>
			  <td class="showegift"><a href="#"><%=dbl_bonus_session_week%>期</a></td><%'贈送期數%>
			  <td><%=int_leave_session%>期</td><%'請假期數%>
			  <td id="totalSession"></td><%'可使用期數%>
			  <td id="haveUseSession"><%=dbl_used_session%>期</td><%'已使用期數%>
			  <td id="unuseSession"></td><%'未使用期數%>
			</table>
        <div class="clear"></div>
	</div>
    <!--总课週期资讯 END -->
	<!--各期明细 START -->
	<div class="arrow_title">各期明细</div>
    <div class="clear"></div>
    <div class="history_bd2">
    <div class="htop">
     	<ul class="sessions">
        <table cellpadding="0" cellspacing="0">
        <tr nowrap="nowrap">
        	<td><div class="hleft"><a href="#" id="left"></a></div></td>
        
		<%
		dim intCount
		dim intSumCount : intSumCount = 0
		dim intShowTag : intShowTag = 0
		dim strTitle : strTitle= ""
		if dbl_used_session = 0 then
			dbl_used_session = 1
		end if
		dim int_show_count : int_show_count = 0
		'使用課时數歷程明細標籤(正常錢包)
		'Beson 若紅利包已開啟,用紅利包開啟日找出合約期數+歸還期數
		cycleFirstDate = datEveryCycleFirstDate '設定週期開始日
		lastCycle = 0
		if not isEmptyOrNull(dat_bonus_begin_datetime) then
			for intCount = 1 to cint(dbl_remain_session_week + dbl_used_session)
				if datediff("d",cycleFirstDate,dat_bonus_begin_datetime) <= 0 then
					exit for
				else
					lastCycle = lastCycle + 1
				end if
				cycleFirstDate = dateadd("d",int_account_cycle_day,cycleFirstDate)
			next
		end if

		if 	lastCycle > 0 then
			if dbl_contract_session_week => lastCycle then
				contractCycleAndRefundCycle = dbl_contract_session_week
			else
				contractCycleAndRefundCycle = lastCycle
			end if
		else
			if dbl_contract_session_week > (dbl_remain_session_week + dbl_used_session) then
				contractCycleAndRefundCycle = dbl_contract_session_week
			else
				contractCycleAndRefundCycle = dbl_remain_session_week + dbl_used_session
			end if
		end if

		cycleFirstDate = datEveryCycleFirstDate '重新設定週期開始日
		hiddenCycle = ""
		for intCount = 1 to cint(contractCycleAndRefundCycle)
			'計算初始顯示期數 start
			if ( int_show_count >= 6 ) then
				str_display = "display:none"
			else
				'response.Write(dbl_remain_session_week&"--"&dbl_used_session&"--"&intCount&"<br>")
'				response.end		
				if (dbl_remain_session_week - dbl_used_session < 5 and dbl_used_session - intCount < 6) then
					str_display = ""
					int_show_count = int_show_count + 1
				elseif (dbl_used_session < intCount and int_show_count <5) then
					str_display = ""
					int_show_count = int_show_count + 1
				else
					str_display = "display:none"
				end if
			end if
			intSumCount = intSumCount + 1
			if ( intSumcount <= dbl_contract_session_week ) then
				strTitle = "合同期数"
			else
				strTitle = "非当期还课时延展期数"
			end if
			'計算初始顯示期數 end
			
			if not isEmptyOrNull(dat_bonus_begin_datetime) then
				if datediff("d",cycleFirstDate,dat_bonus_begin_datetime) <= 0 then
					hiddenCycle = hiddenCycle & intSumCount & ","
					
				end if
			end if
			cycleFirstDate = dateadd("d",int_account_cycle_day,cycleFirstDate)
			if ( g_var_client_sn = 264398 ) then
				if (intSumCount < 12) then
					str_display = "display:none"
				else
					str_display = ""
				end if
			end if
		%>
            <td><li title="<%=strTitle%>" id="<%=intSumCount%>" style="<%=str_display%>"><a id="#week<%=intSumCount%>" href="#">第<%=intSumCount%>期</a></li></td>
        <%	
		next
		
		'使用課时數歷程明細標籤(紅利包)
		for intCount = 1 to cint(dbl_bonus_session_week) 
			'計算初始顯示期數 start
			if ( intSumCount >= 6 ) then
				str_display = "display:none"
			else
				str_display = ""
			end if
			intSumCount = intSumCount + 1
			'計算初始顯示期數 end
			if ( g_var_client_sn = 264398 ) then
				if (intSumCount < 12) then
					str_display = "display:none"
				else
					str_display = ""
				end if
			end if
			
		%>
        	<td><li title="贈送期数" id="<%=intSumCount%>" style="<%=str_display%>"><a id="#week<%=intSumCount%>" href="#">第<%=intSumCount%>期</a></li></td>
		<%
        next
        %>
        <td><div class="hright"><a href="#" id="right"></a></div></td>
        </tr>
        </table>
        </ul>
	</div>
        <div class="clear"></div>
		<div class="session_container">
<%	
if (intSumCount = cint(contractCycleAndRefundCycle) and not isEmptyOrNull(dat_bonus_begin_datetime) ) then
	datEveryCycleLastDate = dat_bonus_begin_datetime
	int_cycle_week_start_session = int_used_session
end if
intSumCount = 0
'各期明細資料(正常錢包)
'response.Write(dbl_remain_session_week )
intTotalSessionRefund = 0 '當期歸還總堂數
intTotalOverSession = 0 '超用總堂數
for intCount = 1 to cint(contractCycleAndRefundCycle)
	intSumCount = intSumCount + 1
	'最後一期需帶入紅利包啟用時間，顯示預約課時時間才會正確
	if (intCount = cint(contractCycleAndRefundCycle) and not isEmptyOrNull(dat_bonus_begin_datetime)) then
		datEveryCycleLastDate = datevalue(dat_bonus_begin_datetime)
	end if
	if not isEmptyOrNull(dat_bonus_begin_datetime) then
		if (datediff("d",datEveryCycleLastDate,dat_bonus_begin_datetime) > 0) then
		else
			datEveryCycleLastDate = datevalue(dat_bonus_begin_datetime)
		end if
	end if
	str_sql_week_list ="select isnull(used_session,0) as used_session , isnull(order_session,0) as order_session , isnull(refund_session,0) as refund_session , isnull(refund_session_all,0) as refund_session_all, isnull(used_video , 0) as used_video from " &_
				"(select sum(svalue) as used_session from client_attend_list " &_
				"where client_sn = '"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid= 1 and refund = 0 " &_
				"and session_sn is not null) as used_session, " &_
				"(select sum(svalue) as order_session from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid = 1 and session_sn is null ) as order_session," &_
				"(select sum(svalue) as refund_session from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and (valid = 0 " &_
				" or ( refund in (1,2) and valid = 1 )) ) as refund_session, " &_
				"(select sum(svalue) as refund_session_all from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid = 1 " &_
				"and refund in (1,2) ) as refund_session_all, "&_
				"(SELECT sum(view_cost) as used_video FROM sessionrecord_view_list "&_
				"left join sessionrecord_fileinfo on "&_
				"sessionrecord_view_list.view_session_recordfiles_sn = sessionrecord_fileinfo.sn "&_
				"where client_sn ='"&g_var_client_sn&"' "&_
				"and view_datetime >= '"&datEveryCycleFirstDate&"' "&_
				"and view_datetime < '"&datEveryCycleLastDate&"') as used_video " 
	arr_week_list = excuteSqlStatementRead(str_sql_week_list , array(), CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_week_list)) then
    	if (Ubound(arr_week_list) >= 0) then
			if (Not isEmptyOrNull(arr_week_list(0,0)) or Not isEmptyOrNull(arr_week_list(4,0))) then
				int_used_session = arr_week_list(0,0) '20120509 Beson 暫時先不加上錄影檔的堂數 + arr_week_list(4,0) / CONST_SESSION_TRANSFORM_POINT
			else
				int_used_session = 0
			end if
			
			if (Not isEmptyOrNull(arr_week_list(1,0))) then
				int_reserve_session = arr_week_list(1,0)
			else
				int_reserve_session = 0
			end if
			
			if (Not isEmptyOrNull(arr_week_list(2,0))) then
				int_refund_session = arr_week_list(2,0)
			else
				int_refund_session = 0
			end if
			if (Not isEmptyOrNull(arr_week_list(3,0))) then
				int_refund_session_all = arr_week_list(3,0)
			else
				int_refund_session_all = 0
			end if
			
		end if
	end if
	'計算每一期可使用的堂數，最後一期可能因為refund而有所不同
	if (intSumCount >= contractCycleAndRefundCycle ) then
		If (flt_available_week_session = 0 ) then 
			int_cycle_week_start_session = 0
		else
			'int_cycle_week_start_session = cdbl(flt_available_session) mod cdbl(flt_available_week_session)
			'20130222 Beson 因mod只取得整數，所以改寫以下方式
			int_cycle_week_start_session = newMod(flt_available_session,flt_available_week_session)
		end if
		
		if (int_cycle_week_start_session = 0 ) then 
			int_cycle_week_start_session = flt_available_week_session
		end if
	else
		int_cycle_week_start_session = flt_available_week_session
	end if
%>
        	<div id="week<%=intSumCount%>" class="session_content">
		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd">
        <tr>
        <th colspan="10" class="top cycleDate">
		<%=datEveryCycleFirstDate%> ~ 
		<%
        '有歸還而延展的最後一周期，如果紅利包開通，顯示日期從上一周期的起始至紅利包開通日，並計算當週期應使用課時
        if (intSumCount = cint(contractCycleAndRefundCycle) and not isEmptyOrNull(dat_bonus_begin_datetime) ) then
            'response.Write(datevalue(dat_bonus_begin_datetime))
            datEveryCycleLastDate = dat_bonus_begin_datetime
			if flt_available_week_session = 0 then
				int_cycle_week_start_session = 0
			else
				'應用使用課時 = 每期應用課時 - (abs(總當期歸還課時 - 總超用課時) mod 每期應用課時)
				if (intTotalSessionRefund - intTotalOverSession) < 0 then
					int_cycle_week_start_session = flt_available_week_session - (abs(intTotalSessionRefund - intTotalOverSession) mod int_cycle_week_start_session)
				else
					int_cycle_week_start_session = (intTotalSessionRefund - intTotalOverSession) mod int_cycle_week_start_session
					'20120507 Beson 如果客戶沒超用也沒當期歸還，則該期應為每期應用課時
					if int_cycle_week_start_session = 0 then
						int_cycle_week_start_session = flt_available_week_session
					end if
				end if
			end if
        elseif ( intSumCount = cint(contractCycleAndRefundCycle) )then
            'response.Write(datEveryCycleLastDate)
        else
			if not isEmptyOrNull(dat_bonus_begin_datetime) then
				if (datediff("d",datEveryCycleLastDate,dat_bonus_begin_datetime) = 0) then
					int_cycle_week_start_session = flt_available_week_session - (abs(intTotalSessionRefund - intTotalOverSession) mod int_cycle_week_start_session)
				else
					datEveryCycleLastDate = dateadd("d", -1, datEveryCycleLastDate)
				end if
			else
				datEveryCycleLastDate = dateadd("d", -1, datEveryCycleLastDate)
			end if
            'response.Write(dateadd("d", -1, datEveryCycleLastDate))					
        end if

		'20120621 Beson 如果客戶已使用完畢課時，又有超用超過一週期的課時
		if not isEmptyOrNull(SessionOverCycle) then
			if intCount = cint(SessionOverCycle) then
				if (intTotalSessionRefund - intTotalOverSession) < 0 then
					int_cycle_week_start_session = flt_available_week_session - (abs(intTotalSessionRefund - intTotalOverSession) mod int_cycle_week_start_session)
				else
					int_cycle_week_start_session = (intTotalSessionRefund - intTotalOverSession) mod int_cycle_week_start_session
				end if
				'20120821 Beson 客戶姚舜龙另外處理
				'if g_var_client_sn = "276123" and str_purchase_account_sn = "131072" then int_cycle_week_start_session = csp_week_start_points
			end if
		elseif not isEmptyOrNull(csp_week_start_points) then
			if intCount = dbl_used_session then
				int_cycle_week_start_session = cint(csp_week_start_points)
			elseif intCount > cint(dbl_used_session) then
				int_cycle_week_start_session = 0
			end if
		end if
				
		if not isEmptyOrNull(dat_bonus_begin_datetime) then
			if (datediff("d",datEveryCycleLastDate,dat_bonus_begin_datetime) > 0) then
			else
				datEveryCycleLastDate = dateadd("d", -1, dat_bonus_begin_datetime)
				isLastRefundCycle = true
			end if
		end if
		
		if not isEmptyOrNull(datEveryCycleLastDate) then
			response.Write datevalue(datEveryCycleLastDate)
		end if

		%>
        </th>
        </tr>
            <tr>
                <th class="tt">应使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">已使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">预约课时</th>
                <td rowspan="3" class="dd">=</td>
                <th class="tt">本期未用<br />(超用)课时</th>
            </tr>
            <tr>
				<%
				'20120621 Beson 如果課時已用完, 應使用課時為0
				if not isEmptyOrNull(SessionOverCycle) then
					 if intCount > cint(SessionOverCycle) then
						int_cycle_week_start_session = 0
					 end if
				end if
				
				int_unuse_session = int_cycle_week_start_session - int_used_session - int_reserve_session '本期未用(超用)课时
				if int_unuse_session < 0 then
					strUnuseSession = "("&int_unuse_session&")"
				else
					strUnuseSession = int_unuse_session
				end if
				if ( g_var_client_sn = 308416 ) then
					if (intSumCount = 13) then
						int_cycle_week_start_session = 4
						strUnuseSession = 0
					end if
				end if
				%>
                <td><%=int_cycle_week_start_session%></td>
                <td><%=int_used_session%></td>
                <td><%=int_reserve_session%></td>
                <td><%=strUnuseSession%></td>
            </tr>
        </table>

    <!--各期明细 END -->
	<!--使用课时数历程资讯 START-->

		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd2">
        <tr>
        <th colspan="6" class="top">使用课时数历程资讯</th>
        </tr>
        <tr>
          <th >日期</th>
          <th >时间</th>
          <th class="tt2">详细说明</th>
          <th >订课资讯</th>
          <th>使用<br>课时数</th>
          <th class="tt">详细资讯</th>
        </tr>
        <%
	if not isEmptyOrNull(datEveryCycleFirstDate) then
		strDatEveryCycleFirstDate = datevalue(datEveryCycleFirstDate)&" 00:00:00"
	else
		strDatEveryCycleFirstDate = datEveryCycleFirstDate
	end if
	
	if not isEmptyOrNull(datEveryCycleLastDate) then
		strDatEveryCycleLastDate = datevalue(datEveryCycleLastDate)&" 23:59:59"
	else
		strDatEveryCycleLastDate = datEveryCycleLastDate
	end if

	arr_cycle = array(g_var_client_sn , strDatEveryCycleFirstDate , strDatEveryCycleLastDate)
	arr_result = excuteSqlStatementRead(str_sql , arr_cycle, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
    	if (Ubound(arr_result) >= 0) then
			'int_unuse_session為負值時代表有超用
			if int_unuse_session < 0 then '當期有超用情況
				if false = isLastRefundCycle then '如果是最後一期
					intTotalOverSession = intTotalOverSession + abs(int_unuse_session)
				end if
				intRefundInSession = 0 '有超用, 則所有歸還都屬超用歸還
			else '當期沒有超用情況
				if int_unuse_session < int_refund_session_all then
					intRefundInSession = int_unuse_session
				else
					intRefundInSession = int_refund_session_all
				end if
			end if
			intTotalSessionRefund = intTotalSessionRefund + intRefundInSession
			'記算是否已是超用的課時
			intOverUseSession = int_cycle_week_start_session
			for intCycleCount = 0 to Ubound(arr_result , 2) 
			%>
			<tr>
                <td width="60"><%=arr_result(0 , intCycleCount)%>&nbsp;</td>
                <td><%=arr_result(1 , intCycleCount)&":30"%>&nbsp;</td>
              <td width="320" style="text-align:left">
				<a href="learning_record_detail.asp?session_sn=<%=arr_result(7 , intCycleCount)%>&read_only=y"><font color="blue"><%=arr_result(2 , intCycleCount)%></font></a>&nbsp;</td>
                <td>
				<%
				if (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_NORMAL and  not isEmptyOrNull(arr_result(4 , intCycleCount)) ) then
					response.Write("大会堂")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_NORMAL ) then
					response.Write("一对三")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_ONE_ON_ONE ) then
					response.Write("一对一")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_ONE_ON_THREE ) then
					response.Write("一对三")
				end if
				%>&nbsp;
                </td>
              <td><%=arr_result(5 , intCycleCount)%>&nbsp;</td>
               <td class="tt"><%
				if (arr_result(6 , intCycleCount)="1") then
					if intRefundInSession > 0 then
						response.Write("<span style=""color:blue;"">当期归还</span>")
						intRefundInSession = intRefundInSession - arr_result(5 , intCycleCount)
					else
						response.Write("<span style=""color:red;"">超用归还</span>")
					end if
				elseif (arr_result(8 , intCycleCount)="0") then
					response.Write("取消课程")
				else
					intOverUseSession = intOverUseSession - arr_result(5 , intCycleCount)
					if intOverUseSession < 0 then
						response.Write "<span style=""color:red;"">超用</span>"
					else
						response.Write("&nbsp;")
					end if
				end if
				%>&nbsp;
                </td>
            </tr>
			<%
			next
		else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
      	end if
	else
		%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
		<%
   	end if   
	%>
            <tr align="center">
                <th colspan="2">觀看時間</th>
                <th width="320">详细說明</th>
                <th>订课资讯</th>
                <th>使用<br>课时数</th>
                <th class="tt">详细资讯</th>
            </tr>
	<%
	'錄影檔紀錄
	arr_video = array(g_var_client_sn , strDatEveryCycleFirstDate , strDatEveryCycleLastDate)
	arr_result = excuteSqlStatementRead(str_sql_video , arr_video, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
    	if (Ubound(arr_result) >= 0) then
			for intCycleCount = 0 to Ubound(arr_result , 2) 
			%>
			<tr>
                <td colspan="2" width="150"><%=arr_result(3 , intCycleCount)%>&nbsp;</td>
                <td width="280"><%=arr_result(6 , intCycleCount)%>&nbsp;</td>
                <td>
				<%
				if (isEmptyOrNull(arr_result(12, intCycleCount))) then
					response.Write("他人錄影檔")
				else
					response.Write("大會堂錄影檔")
				end if
				%></td>
              <td><%=arr_result(2 , intCycleCount)/CONST_SESSION_TRANSFORM_POINT%>&nbsp;</td>
              <td class="tt">&nbsp;</td>
            </tr>
			<%
			next
		else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
      	end if
	else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
   	end if   
	'每周期的第一天
	datEveryCycleFirstDate = dateadd ("d" , int_account_cycle_day , datEveryCycleFirstDate )
	'每周期的最後一天
	datEveryCycleLastDate = dateadd ("d" , int_account_cycle_day , datEveryCycleFirstDate )
		%>
        
        </table>
        </div>
<%
next
'紅利包開通後週期日期須重新計算
if (Not isEmptyOrNull(dat_bonus_begin_datetime)) then
	datEveryCycleFirstDate = datevalue(dat_bonus_begin_datetime)
	'datEveryCycleLastDate = dateadd("d" , int_account_cycle_day , datEveryCycleFirstDate)
	'2012/08/07 Beson 修正第一期紅利包結束日
	datEveryCycleLastDate = dateadd("d", (int_account_cycle_day * (intSumCount + 1)), datContractServiceBegin)
end if
'各期明細資料只顯示紅利包開通後的資料(紅利包)
for intCount = 1 to cint(dbl_bonus_session_week)
	
	intSumCount = intSumCount + 1
	if ( intSumCount = cint(contractCycleAndRefundCycle + dbl_bonus_session_week) )then
		datEveryCycleLastDate = dateadd("d",1,datContractServiceEnd)
	end if
	'20130221 Beson 針對客戶钱晓纹暫時解
	if g_var_client_sn = "308945" and intCount = 2 and str_purchase_account_sn = "130707" then
		datEveryCycleLastDate = "2013/02/08"
	end if
	if g_var_client_sn = "308945" and intCount = 3 and str_purchase_account_sn = "130707" then
		datEveryCycleFirstDate = "2013/02/08"
	end if
	str_sql_week_list ="select isnull(used_session,0) as used_session , isnull(order_session,0) as order_session , isnull(refund_session,0) as refund_session , isnull(refund_session_all,0) as refund_session_all, isnull(used_video , 0) as used_video from " &_
				"(select sum(svalue) as used_session from client_attend_list " &_
				"where client_sn = '"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid= 1 and refund = 0 " &_
				"and session_sn is not null) as used_session, " &_
				"(select sum(svalue) as order_session from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid = 1 and session_sn is null ) as order_session," &_
				"(select sum(svalue) as refund_session from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and (valid = 0 " &_
				" or ( refund in (1,2) and valid = 1 )) ) as refund_session, " &_
				"(select sum(svalue) as refund_session_all from client_attend_list " &_
				"where client_sn ='"&g_var_client_sn&"' " &_
				"and attend_date >= '"&datEveryCycleFirstDate&"' and attend_date < '"&datEveryCycleLastDate&"' " &_
				"and valid = 1 " &_
				"and refund in (1,2) ) as refund_session_all, "&_
				"(SELECT sum(view_cost) as used_video FROM sessionrecord_view_list "&_
				"left join sessionrecord_fileinfo on "&_
				"sessionrecord_view_list.view_session_recordfiles_sn = sessionrecord_fileinfo.sn "&_
				"where client_sn ='"&g_var_client_sn&"' "&_
				"and view_datetime >= '"&datEveryCycleFirstDate&"' "&_
				"and view_datetime < '"&datEveryCycleLastDate&"') as used_video " 
	arr_week_list = excuteSqlStatementRead(str_sql_week_list , array(), CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_week_list)) then
    	if (Ubound(arr_week_list) >= 0) then
			if (Not isEmptyOrNull(arr_week_list(0,0)) or Not isEmptyOrNull(arr_week_list(4,0))) then
				int_used_session = arr_week_list(0,0) + arr_week_list(4,0) / CONST_SESSION_TRANSFORM_POINT
			else
				int_used_session = 0
			end if
			
			if (Not isEmptyOrNull(arr_week_list(1,0))) then
				int_reserve_session = arr_week_list(1,0)
			else
				int_reserve_session = 0
			end if
			
			if (Not isEmptyOrNull(arr_week_list(2,0))) then
				int_refund_session = arr_week_list(2,0)
			else
				int_refund_session = 0
			end if
			if (Not isEmptyOrNull(arr_week_list(3,0))) then
				int_refund_session_all = arr_week_list(3,0)
			else
				int_refund_session_all = 0
			end if
			
		end if
	end if
	'計算每一期可使用的堂數，最後一期為紅利包
	if (intCount = cint(dbl_bonus_session_week)) then
	'當已是最後一期紅利包了
		'應用課時 = 贈送課時 - 總紅利包使用課時
		'20121022 Beson 因合約最後一期有歸還的話, 會回到紅利包, 所以要重新計算紅利包的總堂數
		if cint(reCountMaxBonusPoint) > cint(flt_total_bonus_session) then
			int_cycle_week_start_session = cint(reCountMaxBonusPoint) - cint(intTotalUseBonusSession)
		else
			int_cycle_week_start_session = flt_total_bonus_session - intTotalUseBonusSession
		end if
		'int_cycle_week_start_session = cdbl(flt_bonus_session) mod cdbl(flt_available_week_session)
		if (int_cycle_week_start_session = 0 ) then 
			int_cycle_week_start_session = flt_available_week_session
		end if
		'20121005 因工單VC12081514 由業務主管與財務主管確認加贈三期時數
		if g_var_client_sn = "251085" and int_contract_sn = "50000273" then int_cycle_week_start_session = 10
	else
		int_cycle_week_start_session = flt_available_week_session
		'應用課時 > 已使用課時 , 代表當期沒有用完課時,所以該期使用課時為應用課時
		if int_cycle_week_start_session > int_used_session then
			'總紅利包使用課時 = 總紅利包使用課時 + 應用課時
			intTotalUseBonusSession = intTotalUseBonusSession + int_cycle_week_start_session
		else
		'應用課時 <= 已使用課時 , 代表當期有用完課時或超用,所以該期使用課時為已使用課時
			'總紅利包使用課時 = 總紅利包使用課時 + 已使用課時
			intTotalUseBonusSession = intTotalUseBonusSession + int_used_session
		end if
	end if
%>
	<div id="week<%=intSumCount%>" class="session_content">
		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd">
        <tr>
        <th colspan="10" class="top cycleDate"><%=datEveryCycleFirstDate%> ~ 
		<%
			if ( intSumCount = cint(contractCycleAndRefundCycle + dbl_bonus_session_week) )then
				'response.Write(datEveryCycleLastDate)
				response.write datContractServiceEnd
			else
				response.Write(dateadd("d", -1, datEveryCycleLastDate))					
			end if
			%>
		</th>
        </tr>
            <tr>
                <th class="tt">应使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">已使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">预约课时</th>
                <td rowspan="3" class="dd">=</td>
                <th class="tt">本期未用<br />(超用)课时</th>
            </tr>
            <tr>
			<%
			'20121209 Beson 重大抱怨客戶 - 赵骏 特別處理
			if g_var_client_sn = "250326" and str_purchase_account_sn = "124461" and intSumCount = 25 then
				int_cycle_week_start_session = 16
			end if
			
			'20130224 Beson 針對客戶武威暫時解
			if g_var_client_sn = "322391" and intCount = 3 and str_purchase_account_sn = "132252" then
				int_cycle_week_start_session = 13
			end if
			
			int_unuse_session = int_cycle_week_start_session - int_used_session - int_reserve_session
			if int_unuse_session < 0 then
				strUnuseSession = "("&int_unuse_session&")"
			else
				strUnuseSession = int_unuse_session
			end if
			%>
               	<td><%=int_cycle_week_start_session%></td>
                <td><%=int_used_session%></td>
                <td><%=int_reserve_session%></td>
                <td><%=strUnuseSession%></td>
            </tr>
        </table>

    <!--各期明细 END -->
	<!--使用课时数历程资讯 START-->

		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd2">
        <tr>
        <th colspan="6" class="top">使用课时数历程资讯</th>
        </tr>
        <tr>
          <th >日期</th>
          <th >时间</th>
          <th class="tt2">详细说明</th>
          <th >订课资讯</th>
          <th>使用<br />课时数</th>
          <th class="tt">详细资讯</th>
        </tr>
        <%
	arr_cycle = array(g_var_client_sn , datEveryCycleFirstDate , datEveryCycleLastDate)
	arr_result = excuteSqlStatementRead(str_sql , arr_cycle, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
    	if (Ubound(arr_result) >= 0) then
			for intCycleCount = 0 to Ubound(arr_result , 2) 
			%>
			<tr>
                <td width="60"><%=arr_result(0 , intCycleCount)%>&nbsp;</td>
                <td><%=arr_result(1 , intCycleCount)&":30"%>&nbsp;</td>
                <td width="320">
				<a href="learning_record_detail.asp?session_sn=<%=arr_result(7 , intCycleCount)%>&read_only=y"><font color="blue"><%=arr_result(2 , intCycleCount)%></font></a>&nbsp;</td>
                <td>
				<%
				if (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_NORMAL and  not isEmptyOrNull(arr_result(4 , intCycleCount)) ) then
					response.Write("大会堂")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_NORMAL ) then
					response.Write("一对三")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_ONE_ON_ONE ) then
					response.Write("一对一")
				elseif (arr_result(3 , intCycleCount) = CONST_CLASS_TYPE_ONE_ON_THREE ) then
					response.Write("一对三")
				end if
				%>&nbsp;
                </td>
                <td><%=arr_result(5 , intCycleCount)%>&nbsp;</td>
                <td class="tt"><%
				if (arr_result(6 , intCycleCount)="1") then
					response.Write("归还")
				elseif (arr_result(8 , intCycleCount)="0") then
					response.Write("取消课程")
				else
					response.Write("&nbsp;")
				end if
				%>&nbsp;
                </td>
            </tr>
			<%
			next
		else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
      	end if
	else
		%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
		<%
   	end if   
	%>
            <tr align="center">
                <th colspan="2">觀看時間</th>
                <th width="320">详细說明</th>
                <th>订课资讯</th>
                <th>使用<br>课时数</th>
                <th class="tt">详细资讯</th>
            </tr>
	<%
	'錄影檔紀錄
	arr_video = array(g_var_client_sn , datEveryCycleFirstDate , datEveryCycleLastDate)
	arr_result = excuteSqlStatementRead(str_sql_video , arr_video, CONST_VIPABC_RW_CONN)
	response.Write(g_str_run_time_err_msg)
	if (isSelectQuerySuccess(arr_result)) then
    	if (Ubound(arr_result) >= 0) then
			for intCycleCount = 0 to Ubound(arr_result , 2) 
			%>
			<tr>
                <td colspan="2" width="150"><%=arr_result(3 , intCycleCount)%>&nbsp;</td>
                <td width="280"><%=arr_result(6 , intCycleCount)%>&nbsp;</td>
                <td>
				<%
				if (isEmptyOrNull(arr_result(12, intCycleCount))) then
					response.Write("他人錄影檔")
				else
					response.Write("大會堂錄影檔")
				end if
				%></td>
                <td><%=arr_result(2 , intCycleCount)/CONST_SESSION_TRANSFORM_POINT%>&nbsp;</td>
                <td class="tt">&nbsp;</td>
            </tr>
			<%
			next
		else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
      	end if
	else
			%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
			<%
   	end if   
	'每周期的第一天
	'2012/08/07 Beson 修正第一期紅利包結束日
	datEveryCycleFirstDate = datEveryCycleLastDate
	'每周期的最後一天
	datEveryCycleLastDate = dateadd ("d" , int_account_cycle_day , datEveryCycleFirstDate )
		%>
        
        </table>
        </div>
<%
next
%>
	</div>
    </div>
    <!--使用课时数历程资讯 END-->
	
	
		<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
	</div>
<!--內容end-->

<%
Function Change_Money_To_Int(ByVal InputStr)
	IF InputStr="" Then
		Change_Money_To_Int = ""
		Exit Function
	End IF

	On Error Resume Next 
	Dim str_result
	Dim i
	str_result = ""
	IF ( (Len(Cstr(InputStr))\3 >= 1) and  (Len(Cstr(InputStr))<>3) ) Then
	   For i= 1 To Len(Cstr(InputStr))\3
		   str_result = "," & Right(InputStr,3) & str_result
		   InputStr = Left(Inputstr,Len(Inputstr)-3)
	   Next
	   str_result = InputStr & str_result
	Else
	   str_result = InputStr
	End IF
	IF (Left(str_result,1)=",") Then
	   str_result = Mid(str_result,2,Len(str_result))
	End IF

	Change_Money_To_Int = str_result
End Function

intExendSession = intTotalSessionRefund - intTotalOverSession
if flt_available_week_session = 0 then
	intRefundStep = 0
else
	intRefundStep = CInt(intExendSession)\CInt(flt_available_week_session)
	if (CInt(intExendSession) mod CInt(flt_available_week_session)) > 0 then
		intRefundStep = intRefundStep + 1
	end if
end if

if intExendSession <= 0 then
	strStyle = "style=""color:red;font-weight:bold;"""
	intRefundStep = 0
end if

%>

<div id="Consultantlesson" class="lesson-pop-desc local-lesson-con" style="display:none; position:absolute">
	<div class="pop-desc-top" style="padding-right:500px;"><img src="/images/icon-feedback-desc.gif" width="11" height="9"></div>
	<div class="pop-desc-body" style="height: 115px;">
		<div class="pop-lesson-title" style="text-align:left;"><img src="/images/closebox.png" id="closeConsultantlesson" class="closebox"/>归还课时[ <span style="color:#ff5400;"><strong><%=intTotalSessionRefund%></strong></span> ]课时, 共计延展 [ <span style="color:#ff5400;"><strong id="refundShow"><%=intRefundStep%></strong></span> ]期</div>
		<table width="90%" border="0" cellspacing="0" cellpadding="0" class="history_bd1" style="text-align:left;">
			<tr>
				<th width="18%" >当期归还总课时</th>
				<th width="3%" rowspan="2" class="dd">-</th>
				<th width="17%" >超用课时</th>
				<th width="3%" rowspan="2" class="dd">=</th>
				<th width="18%" >延展期数课时</th>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF"><%=intTotalSessionRefund%></td>
				<td bgcolor="#FFFFFF"><%=intTotalOverSession%></td>
				<td bgcolor="#FFFFFF" <%=strStyle%>><%=intExendSession%></td>
			</tr>
		</table>
		<div class="clear"></div>
	</div>
</div>

<div id="gitftlesson" class="lesson-pop-desc local-lesson-con"  style=" display:none;position:absolute">
	<div class="pop-desc-top" style="padding-right:500px;"><img src="/images/icon-feedback-desc.gif" ></div>
	<div class="pop-desc-body" style="width:550px">
		<div class="pop-lesson-title" style="text-align:left;"><img src="/images/closebox.png" id="closegitftlesson" class="closebox"/>赠送[ <span style="color:#ff5400;"><strong><%=flt_total_bonus_session%></strong></span> ]课时, 共计[ <span style="color:#ff5400;"><strong><%=dbl_bonus_session_week%></strong></span> ]期 
			<img src="/images/icon-ask.gif"  id="showmore" style="cursor:pointer"/>
		</div>
		<div id="moreinfo"  style="display:none;text-align:left;">
			<span class="redfont">赠送期数会接续在合同期数使用完毕之最后一天启用。</span><br />
			(举例)若合同期数最后一期为2012/3/1~2012/3/30，赠送期数为2012/3/31~2012/4/29，当于2012/3/25用毕合同期数最后一课时，即赠送期数会提前由2012/3/25开启
		</div>
		<div class="clear"></div>
	</div>
</div>
<script type="text/javascript">
		
	$(function() {
		$("#refund").html($("#refundShow").html()+"期");
		<%
		int_total_session = CInt(intContractSessionWeek)+CInt(intRefundStep)+CInt(dbl_bonus_session_week)+CInt(int_leave_session)
		unuseSession = int_total_session-dbl_used_session

		if unuseSession < 0 then unuseSession = 0
		%>
		$("#totalSession").html(<%=int_total_session%>+"期");<%'合同期數+归还课时延展期数+贈送期數+請假期數%>
		$("#unuseSession").html(<%=unuseSession%>+"期");
	
<%
	'當紅利包已開始使用, 因可能提早使用, 而目前已使用期數會錯誤, 所以用現在日期算出目前所在期數,再jQuery取代 已使用期數 與 未使用期數
	if not isEmptyOrNull(dat_bonus_begin_datetime) then
		currentDate = datevalue(now)
%>
		<%'從最後一期去比對, 速度較快%>
		for (i=<%=(contractCycleAndRefundCycle+dbl_bonus_session_week)%>;i>=1;i=i-1)
		{
			<%'該期的開始日 < 目前日期%>
			//alert((Date.parse($.trim($($("th.cycleDate")[i-1]).html().split("~")[0]))).valueOf() +"<+" + (Date.parse("<%=currentDate%>")).valueOf() + "\n" + $($("th.cycleDate")[i-1]).html().split("~")[0] + "<<%=currentDate%>"+ "\n" + i);
			if (((Date.parse($.trim($($("th.cycleDate")[i-1]).html().split("~")[0]))).valueOf()) < (Date.parse("<%=currentDate%>")).valueOf())
			{
				$("ul.sessions li").removeClass("active"); <%'移除所有的active Class%>
				$("ul.sessions li#"+i).addClass("active").show(); <%'加入active Class在目前的期數%>
				$("td#haveUseSession").html(i+"期"); <%'修改 已使用期數%>
				$("td#unuseSession").html((<%=contractCycleAndRefundCycle+dbl_bonus_session_week%>-i)+"期"); <%'修改 未使用期數%>
				break;
			}
		}
<%
	end if
	
	'若有合約期數有超用,則用 "超用" 取代該期資訊
	if not isEmptyOrNull(hiddenCycle) then
		arrHiddenCycle = Split(hiddenCycle,",")
		for i= 0 to Ubound(arrHiddenCycle)-1
%>
		$("div#week<%=arrHiddenCycle(i)%>").html("<table width=\"720\" cellpadding=\"0\" cellspacing=\"0\" class=\"sessiontable2\"><tr><td style=\"font-size:20px;\">超用</td></tr></table>");
<%
		next
	end if
%>
	});
</script>