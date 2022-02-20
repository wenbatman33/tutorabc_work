<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<% 
dim str_contract_sn : str_contract_sn = trim( request("hdn_contract_sn") ) '合約編號
dim str_dsn_sn : str_dsn_sn = trim( request("hdn_dsn_sn") ) 'dsn

'cookie補強session機制
if ( isEmptyOrNull(g_var_client_sn) ) then
    session("client_sn") = Request.Cookies("client_sn")
    g_var_client_sn = session("client_sn")
end if

str_sql = " SELECT client, cname, porderno, damount "
str_sql = str_sql & " FROM contract_pay "
str_sql = str_sql & " WHERE client_sn = @client_sn AND sn = @contract_sn AND dsn = @dsn "
var_arr = Array(g_var_client_sn, str_contract_sn, str_dsn_sn)

arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
'response.write g_str_sql_statement_for_debug
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		str_client_name = arr_result(0,0) 
		str_product_name = arr_result(1,0) 
		str_damount = arr_result(3,0) 
		str_proderno = arr_result(2,0) 
	else

	end if 
end if 
%>
<div class="temp_contant">
    <!--內容start-->
    <div class="main_membox">
        <div class="buy">
            <div class="w">
                <%=getWord("BUY_access_1")%>：<span class="bold color_1"><%=getWord("BUY_buy_fail_1")%></span></div>
            <!--明細start-->
            <div>
                <div class="box">
                    <!--订单编号 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_1_7")%></div>
                        <div class="show">
                            <%=str_proderno%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--订单编号 end-->
                    <!--产品名称 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_1_17")%></div>
                        <div class="show">
                            <%=str_product_name%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--产品名称 end-->
                    <!--客户姓名 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_3")%></div>
                        <div class="show">
                            <%=str_client_name%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--客户姓名 end-->
                    <!--信用卡号码 start-->
                    <!--
            <div class="con">
                <div class="name"><%=getWord("BUY_access_4")%></div>
                <div class="show">xxxxx-xxxxx-78522 </div>
                <div class="clear"></div>
            </div>
            -->
                    <!--信用卡号码 end-->
                    <!--刷卡金额 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_5")%></div>
                        <div class="show">
                            <%=str_damount%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--刷卡金额end-->
                    <!--返回结果 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_6")%></div>
                        <div class="show color_1 bold">
                            <%=getWord("BUY_buy_fail_1")%><span class="normal color_5">（<%=getWord("BUY_buy_fail_2")%>）</span></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--返回结果 end-->
                </div>
                <div class="w2">
                    <input type="submit" name="contact_button" id="contact_button" value="+ <%=getWord("BUY_access_8")%>"
                        class="btn_1" onclick="location.href='../contact/contact_us.asp'" />
                    <input type="submit" name="next_step_button" id="Submit1" value="+ <%="下一步"%>" class="btn_1 m-left10"
                        onclick="location.href='../member/buy_1.asp'" />
                </div>
        </div>
        <!--明細end-->
    </div>
</div>
<!--內容end-->
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font></div> 