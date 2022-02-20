<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
'請假期數
int_leave_session = 0
%>
<!--#include virtual="/lib/include/vipabc_csp_include_new.asp"-->
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
		//TODO 這還要加延展期數
		var int_max_week = parseInt(<%=dbl_used_session + dbl_bonus_session_week%>)+1;
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
					<td><%=dbl_contract_session_week%>期</td><!--合同期数-->

					<%
						'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493			
						if ( Instr(CONST_DIANPING_GROUPBUY_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
						%>
							<td>付费大会堂2堂，一对三真人授课3堂</td>						
						<%
						else
						%>
							<td><%=flt_available_week_session%>课時</td><!--每期-->
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
			  <th width="3%" rowspan="2" class="dd">=</th>
			  <th width="12%">可使用期数</th>
			  <th width="3%" rowspan="2" class="dd">-</th>
			  <th width="12%">已使用期数</th>
			  <th width="3%" rowspan="2" class="dd">=</th>
			  <th width="12%" >未使用期数</th>
            </tr>
			  <td><%=dbl_contract_session_week%>期</td><!--合同期數-->
			  <td class="showeturn"><a href="#" id="refund"></a></td><!--归还课时延展期数-->
			  <td class="showegift"><a href="#"><%=dbl_bonus_session_week%>期</a></td><!--贈送期數-->
			  <td id="totalSession"></td><!--可使用期數-->
			  <td id="haveUseSession"><%=dbl_used_session%>期</td><!--已使用期數-->
			  <td id="unuseSession"></td><!--未使用期數-->
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
'期數標籤 開始----------------------------------------------------------------------------

'合約期數標籤 開始--------------------------------------------------------------------
Dim intCycleNum					'期數
Dim datCycleStartDate			'每週期開始日
Dim datCycleEndDate				'每週期結束日
Dim intRemainSession			'剩餘課時
Dim intCycleShouldUseSession	'每週應用課時
Dim intCycleIsUseSession		'每週期已用課時 (含歸還課時)
Dim intCycleRefundSession		'每週期歸還課時 (含當期歸還與超用歸還)
Dim intCycleUnUseSession		'每週期未用課時
Dim intCycleOverUseSession		'每週期超用課時
Dim intTotalOverUserSession		'超用合約課時 [用來計算最後一期是否有超用]

'設定開始值
intCycleNum = 1											'期數
datCycleStartDate = datContractServiceBegin				'每週期開始日
intRemainSession = flt_contract_total_session			'剩餘課時
intTotalOverUserSession = -(flt_contract_total_session)	'超用合約課時 [用來計算最後一期是否有超用]
' response.write "intTotalOverUserSession:" & intTotalOverUserSession & "<br/>"
' response.write "datCycleStartDate:" & datCycleStartDate & "<br/>"						'每週期開始日
' response.write "intRemainSession:" & intRemainSession & "<br/>"							'剩餘課時
' response.write "int_account_cycle_day:" & int_account_cycle_day & "<br/>"				'週期天數
' response.write "flt_available_week_session:" & flt_available_week_session & "<br/>"		'每週期固定使用課時

'合約每週期 [剩餘課時 < 0 和 期數 > 合同期數]
do until (intRemainSession <= 0 and intCycleNum > dbl_contract_session_week)

	'應用課時 [剩餘課時 > 每週應用課時]
	if cdbl(intRemainSession) > cdbl(flt_available_week_session) then
		'應用課時 = 每週應用課時
		intCycleShouldUseSession = flt_available_week_session
	else
		'應用課時 = 剩餘課時
		intCycleShouldUseSession = intRemainSession
	end if
	
	'週期結束日
	datCycleEndDate = getFormatDateTime(dateadd("d", int_account_cycle_day - 1, datCycleStartDate), 5)
	
	'週期課時使用 [已用課時,歸還課時]
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT  ISNULL(SUM(svalue),0) AS totalSvalue , " &_
	"        ISNULL(SUM(CASE WHEN refund = 1 THEN svalue " &_
	"                 ELSE 0 " &_
	"            END),0) AS totalRefund " &_
	"FROM    dbo.client_attend_list WITH ( NOLOCK ) " &_
	"WHERE   client_sn = @client_sn " &_
	"        AND attend_date BETWEEN @startdate + ' 00:00:00' " &_
	"                        AND     @enddate + ' 23:59:00' " &_
	"        AND valid = 1 /*取有訂課的*/ "
	arrParam = Array(g_var_client_sn,getFormatDateTime(datCycleStartDate, 5), getFormatDateTime(datCycleEndDate, 5))
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_R_CONN)
	if not rs.eof then
		intCycleIsUseSession	= rs("totalSvalue")	'已用課時
		intCycleRefundSession	= rs("totalRefund")	'歸還課時
	else
		intCycleIsUseSession	= 0
		intCycleRefundSession	= 0
	end if
	rs.close
	Set rs = nothing
	
	'未用課時 [已用課時 > 應用課時]
	if cdbl(intCycleIsUseSession) > cdbl(intCycleShouldUseSession) then
		'未用課時 = 0
		intCycleUnUseSession	= 0
	else
		'未用課時 = 應用課時 - 已用課時
		intCycleUnUseSession	= cdbl(intCycleShouldUseSession) - cdbl(intCycleIsUseSession)
	end if
	
	'超用課時 [(已用課時 - 歸還課時) > 應用課時]
	if (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) > cdbl(intCycleShouldUseSession) then
		'超用課時 = (已用課時 - 歸還課時) - 應用課時
		intCycleOverUseSession	= (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) - cdbl(intCycleShouldUseSession)
	else
		'超用課時 = 0
		intCycleOverUseSession	= 0
	end if
	
	' response.write "期數:" & intCycleNum & "剩餘課時:" & intRemainSession & "應用課時:" & intCycleShouldUseSession
	' response.write "已用課時:" & intCycleIsUseSession & "未用課時:" & intCycleUnUseSession
	' response.write "歸還課時:" & intCycleRefundSession & "超用課時:" & intCycleOverUseSession
	' response.write "<br/>"

	'剩下課時 [(剩餘課時 - (已用課時 + 未用課時) + 歸還課時) > 0]
	if cdbl(intRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) + cdbl(intCycleRefundSession) > 0 then
		'剩下課時 = (剩餘課時 - (已用課時 + 未用課時) + 歸還課時)
		intRemainSession	= cdbl(intRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) + cdbl(intCycleRefundSession)
	else
		'剩下課時 = 0
		intRemainSession	= 0
	end if
	
	'顯示標籤
	if intCycleNum <= dbl_contract_session_week then
		strTitle = "合同期数"
	else
		strTitle = "非当期还课时延展期数"
	end if
	
	'於最後一期是否有超用
	'計算是否有超用合約堂數
	'超用合約堂數 = 超用合約堂數 + 已用課時 + 未用課時 - 歸還課時
	intTotalOverUserSession = intTotalOverUserSession + intCycleIsUseSession + intCycleUnUseSession - intCycleRefundSession
	
	if intCycleNum >= dbl_used_session and intCycleNum <= dbl_used_session + 3 then
		str_display = ""
	else
		str_display = "display:none"
	end if
%>
			<td><li title="<%=strTitle%>" id="<%=intCycleNum%>" style="<%=str_display%>"><a id="#week<%=intCycleNum%>" href="#">第<%=intCycleNum%>期</a></li></td>
<%
	'下週期開始日
	datCycleStartDate = getFormatDateTime(dateadd("d", int_account_cycle_day, datCycleStartDate), 5)
	
	'期數加1
	intCycleNum = intCycleNum + 1
loop
'合約期數標籤 結束--------------------------------------------------------------------

'贈送期數標籤 開始--------------------------------------------------------------------
Dim intBonusCycleNum		'贈送期數

'有贈送堂數才有贈送期
if cint(flt_total_bonus_session) > 0 then
	'贈送每週期
	for intBonusCycleNum = intCycleNum to (intCycleNum + cint(dbl_bonus_session_week) - 1)
		if intBonusCycleNum >= dbl_used_session and intBonusCycleNum <= dbl_used_session + 3 then
			str_display = ""
		else
			str_display = "display:none"
		end if
%>
			<td><li title="贈送期数" id="<%=intBonusCycleNum%>" style="<%=str_display%>"><a id="#week<%=intBonusCycleNum%>" href="#">第<%=intBonusCycleNum%>期</a></li></td>
<%
	next
end if
'贈送期數標籤 結束--------------------------------------------------------------------

'期數標籤 結束----------------------------------------------------------------------------
%>
			<td><div class="hright"><a href="#" id="right"></a></div></td>
        </tr>
        </table>
        </ul>
	</div>
        <div class="clear"></div>
		<div class="session_container">
<%
'期數明細 開始----------------------------------------------------------------------------

'合約期數明細 開始--------------------------------------------------------------------
Dim intCycleNowRefundSession	'當期歸還課時
Dim intCycleOverRefundSession	'超用歸還課時
Dim intTotalNowRefundSession	'總當期歸還課時
Dim intTotalOverUseSession		'總超用課時

'重新設定開始值
intCycleNum = 1											'期數
datCycleStartDate = datContractServiceBegin				'每週期開始日
intRemainSession = flt_contract_total_session			'剩餘課時
intTotalOverUserSession = -(flt_contract_total_session)
intTotalNowRefundSession = 0
intTotalOverUseSession = 0

'合約每週期 [剩餘課時 < 0 和 期數 > 合同期數]
do until (intRemainSession <= 0 and intCycleNum > dbl_contract_session_week)

	'應用課時 [剩餘課時 > 每週應用課時]
	if cdbl(intRemainSession) > cdbl(flt_available_week_session) then
		'應用課時 = 每週應用課時
		intCycleShouldUseSession = flt_available_week_session
	else
		'應用課時 = 剩餘課時
		intCycleShouldUseSession = intRemainSession
	end if
	
	'週期結束日
	datCycleEndDate = getFormatDateTime(dateadd("d", int_account_cycle_day - 1, datCycleStartDate), 5)
	
	'針對最後一期日期作處理以附舊合約的結束日
	'應用課時 = 剩餘課時 and 贈送堂數 <= 0 and 合約結束日 > 週期結束日
	if (intCycleShouldUseSession = intRemainSession) and _
		flt_total_bonus_session <= 0 and _
		datediff("d",datCycleEndDate,datPurchaseServiceEnd) > 0 then
		datCycleEndDate = datPurchaseServiceEnd
	end if
	
	'週期課時使用 [已用課時,歸還課時]
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT  ISNULL(SUM(svalue),0) AS totalSvalue , " &_
	"        ISNULL(SUM(CASE WHEN refund = 1 THEN svalue " &_
	"                 ELSE 0 " &_
	"            END),0) AS totalRefund " &_
	"FROM    dbo.client_attend_list WITH ( NOLOCK ) " &_
	"WHERE   client_sn = @client_sn " &_
	"        AND attend_date BETWEEN @startdate + ' 00:00:00' " &_
	"                        AND     @enddate + ' 23:59:00' " &_
	"        AND valid = 1 /*取有訂課的*/ "
	arrParam = Array(g_var_client_sn,getFormatDateTime(datCycleStartDate, 5), getFormatDateTime(datCycleEndDate, 5))
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_R_CONN)
	if not rs.eof then
		intCycleIsUseSession	= rs("totalSvalue")	'已用課時
		intCycleRefundSession	= rs("totalRefund")	'歸還課時
	else
		intCycleIsUseSession	= 0
		intCycleRefundSession	= 0
	end if
	rs.close
	Set rs = nothing
	
	'未用課時 [已用課時 > 應用課時]
	if cdbl(intCycleIsUseSession) > cdbl(intCycleShouldUseSession) then
		'未用課時 = 0
		intCycleUnUseSession	= 0
	else
		'未用課時 = 應用課時 - 已用課時
		intCycleUnUseSession	= cdbl(intCycleShouldUseSession) - cdbl(intCycleIsUseSession)
	end if
	
	'超用課時 [(已用課時 - 歸還課時) > 應用課時]
	if (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) > cdbl(intCycleShouldUseSession) then
		'超用課時 = (已用課時 - 歸還課時) - 應用課時
		intCycleOverUseSession	= (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) - cdbl(intCycleShouldUseSession)
	else
		'超用課時 = 0
		intCycleOverUseSession	= 0
	end if
	
	'計算總超用課時
	intTotalOverUseSession = intTotalOverUseSession + cdbl(intCycleOverUseSession)
	
	'當期歸還 [應用課時 > 已用課時]
	if cdbl(intCycleShouldUseSession) > cdbl(intCycleIsUseSession) then
		'當期歸還 = 歸還課時
		intCycleNowRefundSession = intCycleRefundSession
	else
		'歸還課時 > 0
		if cdbl(intCycleRefundSession) > 0 then
			'應用課時 > (已用課時 - 歸還課時)
			if cdbl(intCycleShouldUseSession) > (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) then
				'當期歸還 = 應用課時 - (已用課時 - 歸還課時)
				intCycleNowRefundSession = cdbl(intCycleShouldUseSession) - (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession))
			else
				'當期歸還 = 0
				intCycleNowRefundSession = 0
			end if
		else
			'當期歸還 = 0
			intCycleNowRefundSession = 0
		end if
	end if
	
	'計算總當期歸還課時
	intTotalNowRefundSession = intTotalNowRefundSession + cdbl(intCycleNowRefundSession)
	
	'超用歸還 [歸還課時 - 當期歸還]
	intCycleOverRefundSession = cdbl(intCycleRefundSession) - cdbl(intCycleNowRefundSession)
	
	' response.write "期數:" & intCycleNum & "剩餘課時:" & intRemainSession & "應用課時:" & intCycleShouldUseSession
	' response.write "已用課時:" & intCycleIsUseSession & "未用課時:" & intCycleUnUseSession
	' response.write "歸還課時:" & intCycleRefundSession & "超用課時:" & intCycleOverUseSession
	' response.write "<br/>"

	'剩下課時 [(剩餘課時 - (已用課時 + 未用課時) + 歸還課時) > 0]
	if cdbl(intRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) + cdbl(intCycleRefundSession) > 0 then
		'剩下課時 = (剩餘課時 - (已用課時 + 未用課時) + 歸還課時)
		intRemainSession	= cdbl(intRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) + cdbl(intCycleRefundSession)
	else
		'剩下課時 = 0
		intRemainSession	= 0
	end if
	
	'於最後一期是否有超用
	'計算是否有超用合約堂數
	'超用合約堂數 = 超用合約堂數 + 已用課時 + 未用課時 - 歸還課時
	intTotalOverUserSession = intTotalOverUserSession + intCycleIsUseSession + intCycleUnUseSession - intCycleRefundSession
%>
        	<div id="week<%=intCycleNum%>" class="session_content">
		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd">
        <tr>
        <th colspan="10" class="top cycleDate">
			<%=datCycleStartDate%> ~ <%=datCycleEndDate%>
        </th>
        </tr>
            <tr>
                <th class="tt">应使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">已使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">未使用课时</th>
                <td rowspan="3" class="dd">=</td>
                <th class="tt">超用课时</th>
            </tr>
            <tr>
                <td><%=intCycleShouldUseSession%></td><!--应使用课时-->
                <td><%=intCycleIsUseSession%></td><!--已使用课时-->
                <td><%=intCycleUnUseSession%></td><!--未使用课时-->
                <td><%=intCycleOverUseSession%></td><!--超用课时-->
            </tr>
        </table>

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
	'週期課時使用明細
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT  CONVERT(VARCHAR(10),a.attend_date,120) AS attend_date, " &_
	"		a.attend_sestime, " &_
	"		ISNULL(b.ltitle,'') AS material_title, " &_
	"		CASE " &_
	"			WHEN a.attend_livesession_types = 4 AND a.special_sn > 0 THEN N'大会堂' " &_
	"			WHEN a.attend_livesession_types = 4 THEN N'一对三' " &_
	"			WHEN a.attend_livesession_types = 1 OR a.attend_livesession_types = 7 THEN N'一对一' " &_
	"			ELSE '' " &_
	"		END AS class_type, " &_
	"		a.svalue, " &_
	"		a.refund, " &_
	"		a.valid, " &_
	"		a.session_sn, " &_
	"		a.sn AS ClientAttendListSn " &_
	"FROM    dbo.client_attend_list AS a WITH ( NOLOCK ) " &_
	"LEFT JOIN dbo.material AS b WITH (NOLOCK) ON a.attend_mtl_1 = b.course " &_
	"LEFT JOIN dbo.ClientAttendListDetail WITH (NOLOCK) ON a.sn = ClientAttendListDetail.ParentSn " &_
	"WHERE   client_sn = @client_sn " &_
	"        AND attend_date BETWEEN @startdate + ' 00:00:00' " &_
	"                        AND     @enddate + ' 23:59:00' " &_
	"ORDER BY a.attend_date,a.attend_sestime,ISNULL(ClientAttendListDetail.Minutes,30) "
	arrParam = Array(g_var_client_sn,getFormatDateTime(datCycleStartDate, 5), getFormatDateTime(datCycleEndDate, 5))
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_R_CONN)
	tempNowRefund = 0
	if not rs.eof then
		numCount = rs.RecordCount
		for i = 0 to numCount-1
%>
			<tr>
                <td width="60"><%=rs("attend_date")%>&nbsp;</td><!--日期-->
                <td><%=getSessionDateTime(rs("ClientAttendListSn"), rs("attend_date"), rs("attend_sestime"), rs("session_sn"), 6, 0, CONST_VIPABC_RW_CONN)%>&nbsp;</td><!--时間-->
				<td width="320" style="text-align:left"><!--详细說明-->
					<a href="learning_record_detail.asp?session_sn=<%=rs("session_sn")%>&read_only=y"><font color="blue"><%=rs("material_title")%></font></a>&nbsp;
				</td>
                <td><%=rs("class_type")%>&nbsp;</td><!--订课资讯-->
				<td><%=rs("svalue")%>&nbsp;</td><!--使用课时数-->
				<td class="tt"><!--详细资讯-->
<%
				if rs("valid") = "0" then
					response.write "取消課程"
				elseif rs("refund") = "1" then
					if cdbl(intCycleNowRefundSession) > 0 and tempNowRefund < intCycleNowRefundSession then
						tempNowRefund = tempNowRefund + cdbl(rs("svalue"))
						response.write "<span style=""color:blue;"">當期歸還</span>"
					else
						response.write "<span style=""color:red;"">超用歸還</span>"
					end if
				else
					intCycleShouldUseSession = intCycleShouldUseSession - cdbl(rs("svalue"))
					if intCycleShouldUseSession < 0 then
						response.Write "<span style=""color:red;"">超用</span>"
					end if
				end if
%>
				&nbsp;</td>
            </tr>
<%
			rs.MoveNext
		next
	else
%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
<%
	end if
	rs.close
	Set rs = nothing
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
	arr_video = array(g_var_client_sn , datCycleStartDate , datCycleEndDate)
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
%>
				</td>
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
%>
        </table>
        </div>
<%
	'下週期開始日
	datCycleStartDate = getFormatDateTime(dateadd("d", int_account_cycle_day, datCycleStartDate), 5)
	
	'期數加1
	intCycleNum = intCycleNum + 1
loop
'合約期數明細 結束--------------------------------------------------------------------

'贈送期數明細 開始--------------------------------------------------------------------
Dim intBonusRemainSession	'剩餘贈送課時

'設定開始值
'如果有超用, 就用贈送課時扣除
if cint(intTotalOverUserSession) > 0 then
	'剩餘贈送課時 = 贈送總課時 - 超用合約課時
	intBonusRemainSession = flt_total_bonus_session - intTotalOverUserSession
	'總超用課時 = 總超用課時 - 超用合約課時
	intTotalOverUseSession = intTotalOverUseSession - intTotalOverUserSession
else
	intBonusRemainSession = flt_total_bonus_session
end if

'有贈送堂數才有贈送期
if cint(flt_total_bonus_session) > 0 then
	'贈送每週期
	for intBonusCycleNum = intCycleNum to (intCycleNum + cint(dbl_bonus_session_week) - 1)
		'應用課時 [剩餘課時 > 每週應用課時]
		if cdbl(intBonusRemainSession) > cdbl(flt_available_week_session) then
			'應用課時 = 每週應用課時
			intCycleShouldUseSession = flt_available_week_session
		else
			'應用課時 = 剩餘課時
			intCycleShouldUseSession = intBonusRemainSession
		end if
		
		'週期結束日
		datCycleEndDate = getFormatDateTime(dateadd("d", int_account_cycle_day - 1, datCycleStartDate), 5)
		
		'針對最後一期日期作處理以附合合約算錯問題
		'應用課時 = 剩餘課時
		if (intCycleShouldUseSession = intBonusRemainSession) then
			datCycleEndDate = getFormatDateTime(dateadd("d",1,datCycleEndDate), 5)
		end if
		
		'週期課時使用 [已用課時,歸還課時]
		Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql = "SELECT  ISNULL(SUM(svalue),0) AS totalSvalue , " &_
		"        ISNULL(SUM(CASE WHEN refund = 1 THEN svalue " &_
		"                 ELSE 0 " &_
		"            END),0) AS totalRefund " &_
		"FROM    dbo.client_attend_list WITH ( NOLOCK ) " &_
		"WHERE   client_sn = @client_sn " &_
		"        AND attend_date BETWEEN @startdate + ' 00:00:00' " &_
		"                        AND     @enddate + ' 23:59:00' " &_
		"        AND valid = 1 /*取有訂課的*/ "
		arrParam = Array(g_var_client_sn,getFormatDateTime(datCycleStartDate, 5), getFormatDateTime(datCycleEndDate, 5))
		intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_R_CONN)
		if not rs.eof then
			intCycleIsUseSession	= rs("totalSvalue")	'已用課時
			intCycleRefundSession	= rs("totalRefund")	'歸還課時
		else
			intCycleIsUseSession	= 0
			intCycleRefundSession	= 0
		end if
		rs.close
		Set rs = nothing
		
		'未用課時 [已用課時 > 應用課時]
		if cdbl(intCycleIsUseSession) > cdbl(intCycleShouldUseSession) then
			'未用課時 = 0
			intCycleUnUseSession	= 0
		else
			'未用課時 = 應用課時 - 已用課時
			intCycleUnUseSession	= cdbl(intCycleShouldUseSession) - cdbl(intCycleIsUseSession)
		end if
		
		'超用課時 [(已用課時 - 歸還課時) > 應用課時]
		if (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) > cdbl(intCycleShouldUseSession) then
			'超用課時 = (已用課時 - 歸還課時) - 應用課時
			intCycleOverUseSession	= (cdbl(intCycleIsUseSession) - cdbl(intCycleRefundSession)) - cdbl(intCycleShouldUseSession)
		else
			'超用課時 = 0
			intCycleOverUseSession	= 0
		end if
		
		' response.write "贈送期數:" & intBonusCycleNum & "剩餘課時:" & intBonusRemainSession & "應用課時:" & intCycleShouldUseSession
		' response.write "已用課時:" & intCycleIsUseSession & "未用課時:" & intCycleUnUseSession
		' response.write "歸還課時:" & intCycleRefundSession & "超用課時:" & intCycleOverUseSession
		' response.write "<br/>"

		'檢查是否有超用下期課時又有歸還的
		intBonusOverUseRefuntSession = 0
		'已用課時 > 應用課時
		if cdbl(intCycleIsUseSession) > cdbl(intCycleShouldUseSession) then
			'(已用課時 - 應用課時) > 歸還課時
			if (cdbl(intCycleIsUseSession) - cdbl(intCycleShouldUseSession)) > cdbl(intCycleRefundSession) then
				'已用課時 > 剩下課時
				if cdbl(intCycleIsUseSession) > cdbl(intBonusRemainSession) then
					'超用歸還課時 = 歸還課時 - (已用課時 - 剩餘課時)
					intBonusOverUseRefuntSession = cdbl(intCycleRefundSession) - (cdbl(intCycleIsUseSession) - cdbl(intBonusRemainSession))
				else
					'超用歸還課時 = 歸還課時
					intBonusOverUseRefuntSession = cdbl(intCycleRefundSession)
				end if
			else
				'超用歸還課時 = 已用課時 - 歸還課時
				intBonusOverUseRefuntSession = cdbl(intCycleIsUseSession) - cdbl(intCycleShouldUseSession)
			end if
		else
			'超用歸還課時 = 0
			intBonusOverUseRefuntSession = 0
		end if
		
		'剩下課時 [(剩餘課時 - (已用課時 + 未用課時)) > 0]
		'不計算歸還課時,贈送期的歸還需當期使用完畢
		if cdbl(intBonusRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) > 0 then
			'剩下課時 = (剩餘課時 - (已用課時 + 未用課時)) + 超用歸還課時
			intBonusRemainSession	= cdbl(intBonusRemainSession) - (cdbl(intCycleIsUseSession) + cdbl(intCycleUnUseSession)) + cdbl(intBonusOverUseRefuntSession)
		else
			'剩下課時 = 0 + 超用歸還課時
			intBonusRemainSession	= 0 + cdbl(intBonusOverUseRefuntSession)
		end if
%>
	<div id="week<%=intBonusCycleNum%>" class="session_content">
		<table width="100%" cellpadding="0" cellspacing="0" class="hisbd">
        <tr>
        <th colspan="10" class="top cycleDate">
			<%=datCycleStartDate%> ~ <%=datCycleEndDate%>
		</th>
        </tr>
            <tr>
                <th class="tt">应使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">已使用课时</th>
                <td rowspan="3" class="dd">-</td>
                <th class="tt">未使用课时</th>
                <td rowspan="3" class="dd">=</td>
                <th class="tt">超用课时</th>
            </tr>
            <tr>
               	<td><%=intCycleShouldUseSession%></td><!--应使用课时-->
                <td><%=intCycleIsUseSession%></td><!--已使用课时-->
                <td><%=intCycleUnUseSession%></td><!--未使用课时-->
                <td><%=intCycleOverUseSession%></td><!--超用课时-->
            </tr>
        </table>

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
	'週期課時使用明細
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT  CONVERT(VARCHAR(10),a.attend_date,120) AS attend_date, " &_
	"		a.attend_sestime, " &_
	"		ISNULL(b.ltitle,'') AS material_title, " &_
	"		CASE " &_
	"			WHEN a.attend_livesession_types = 4 AND a.special_sn > 0 THEN N'大会堂' " &_
	"			WHEN a.attend_livesession_types = 4 THEN N'一对三' " &_
	"			WHEN a.attend_livesession_types = 1 OR a.attend_livesession_types = 7 THEN N'一对一' " &_
	"			ELSE '' " &_
	"		END AS class_type, " &_
	"		a.svalue, " &_
	"		a.refund, " &_
	"		a.valid, " &_
	"		a.session_sn ," &_
	"		a.sn AS ClientAttendListSn " &_
	"FROM    dbo.client_attend_list AS a WITH ( NOLOCK ) " &_
	"LEFT JOIN dbo.material AS b WITH (NOLOCK) ON a.attend_mtl_1 = b.course " &_
	"LEFT JOIN dbo.ClientAttendListDetail WITH (NOLOCK) ON a.sn = ClientAttendListDetail.ParentSn " &_
	"WHERE   client_sn = @client_sn " &_
	"        AND attend_date BETWEEN @startdate + ' 00:00:00' " &_
	"                        AND     @enddate + ' 23:59:00' " &_
	"ORDER BY a.attend_date,a.attend_sestime, ISNULL(ClientAttendListDetail.Minutes,30)"
	arrParam = Array(g_var_client_sn,getFormatDateTime(datCycleStartDate, 5), getFormatDateTime(datCycleEndDate, 5))
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_R_CONN)
	tempNowRefund = 0
	if not rs.eof then
		numCount = rs.RecordCount
		for i = 0 to numCount-1
%>
			<tr>
                <td width="60"><%=rs("attend_date")%>&nbsp;</td><!--日期-->
                <td><%=getSessionDateTime(rs("ClientAttendListSn"), rs("attend_date"), rs("attend_sestime"), rs("session_sn"), 6, 0, CONST_VIPABC_RW_CONN)%>&nbsp;</td><!--时間-->
                <td width="320"><!--详细說明-->
					<a href="learning_record_detail.asp?session_sn=<%=rs("session_sn")%>&read_only=y"><font color="blue"><%=rs("material_title")%></font></a>&nbsp;
				</td>
                <td><%=rs("class_type")%>&nbsp;</td><!--订课资讯-->
                <td><%=rs("svalue")%>&nbsp;</td><!--使用课时数-->
                <td class="tt"><!--详细资讯-->
<%
				if rs("valid") = "0" then
					response.write "取消課程"
				elseif rs("refund") = "1" then
					response.write "歸還"
				else
					intCycleShouldUseSession = intCycleShouldUseSession - cdbl(rs("svalue"))
					if intCycleShouldUseSession < 0 then
						response.Write "<span style=""color:red;"">超用</span>"
					end if
				end if
%>
				&nbsp;</td>
            </tr>
<%
			rs.MoveNext
		next
	else
%>
			<tr><td colspan="6" class="tt">查无资料</td></tr>
<%
	end if
	rs.close
	Set rs = nothing
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
	arr_video = array(g_var_client_sn , datCycleStartDate , datCycleEndDate)
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
%>
				</td>
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
%>
        </table>
        </div>
<%
		'下週期開始日
		datCycleStartDate = getFormatDateTime(dateadd("d", int_account_cycle_day, datCycleStartDate), 5)
	next
end if
'贈送期數明細 結束--------------------------------------------------------------------

'期數明細 結束----------------------------------------------------------------------------
%>
	</div>
    </div>
	
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

' response.write dbl_contract_session_week & ","
' response.write dbl_bonus_session_week & ","
' response.write intBonusCycleNum & ","
if (intCycleNum - 1) - dbl_contract_session_week > 0 then
	intRerundWeek = (intCycleNum - 1) - dbl_contract_session_week
else
	intRerundWeek = 0
end if

%>
<script type="text/javascript">
$(function(){
	//歸還期數
	$("#refund").html($("#refundShow").html()+"期");
	//可使用期數 = 合同期數 + 歸還期數 + 贈送期數
	$("#totalSession").html(<%=(cint(dbl_contract_session_week) + cint(intRerundWeek) + cint(dbl_bonus_session_week))%>+"期");
	//未使用期數 = 可使用期數 - 已使用期數
	$("td#unuseSession").html((<%=(cint(dbl_contract_session_week) + cint(intRerundWeek) + cint(dbl_bonus_session_week) - cint(dbl_used_session))%>)+"期"); <!--修改 未使用期數-->
});
</script>
<div id="Consultantlesson" class="lesson-pop-desc local-lesson-con" style="display:none; position:absolute">
	<div class="pop-desc-top" style="padding-right:500px;"><img src="/images/icon-feedback-desc.gif" width="11" height="9"></div>
	<div class="pop-desc-body" style="height: 115px;">
		<div class="pop-lesson-title" style="text-align:left;"><img src="/images/closebox.png" id="closeConsultantlesson" class="closebox"/>归还课时[ <span style="color:#ff5400;"><strong><%=intTotalNowRefundSession%></strong></span> ]课时, 共计延展 [ <span style="color:#ff5400;"><strong id="refundShow"><%=intRerundWeek%></strong></span> ]期</div>
		<table width="90%" border="0" cellspacing="0" cellpadding="0" class="history_bd1" style="text-align:left;">
			<tr>
				<th width="18%" >当期归还总课时</th>
				<th width="3%" rowspan="2" class="dd">-</th>
				<th width="17%" >超用课时</th>
				<th width="3%" rowspan="2" class="dd">=</th>
				<th width="18%" >延展期数课时</th>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF"><%=intTotalNowRefundSession%></td><!--当期归还总课时-->
				<td bgcolor="#FFFFFF"><%=intTotalOverUseSession%></td><!--超用课时累积-->
				<td bgcolor="#FFFFFF"><%=(intTotalNowRefundSession - intTotalOverUseSession)%></td><!--延展期数课时-->
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