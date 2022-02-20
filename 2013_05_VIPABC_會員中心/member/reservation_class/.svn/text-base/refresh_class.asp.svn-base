<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% dim bolDebugMode : bolDebugMode = false %>
<!--#include virtual="/program/member/reservation_class/include/getComboProductData.asp"-->
<span class="qty red" ><%
'會有六個session 五個變數
'bolComboProduct, bolComboHaving1On1, bolComboHaving1On3, bolComboHavingLobby, bolComboHavingRecord
if ( true = bolDebugMode ) then
    response.Write "bolComboProduct: " & bolComboProduct & "<br />"
    response.Write "bolComboHaving1On1: " & bolComboHaving1On1 & "<br />"
    response.Write "bolComboHaving1On3: " & bolComboHaving1On3 & "<br />"
    response.Write "bolComboHavingLobby: " & bolComboHavingLobby & "<br />"
    response.Write "bolComboHavingRecord: " & bolComboHavingRecord & "<br />"
end if

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493     
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

'20121110 阿捨新增 套餐產品設計 這隻程式只for大會堂剩餘堂數顯示寫死99
if ( true = bolComboProduct ) then
    dim intHavingSession : intHavingSession = getComboProductSession(int_member_client_sn, str_now_use_contract_sn, "99", 1, 0, CONST_VIPABC_R_CONN)
    '20130102 阿捨新增 套餐產品判斷 for 贈送堂數
    dim intComboGiftSession : intComboGiftSession = getComboProductGift(str_now_use_contract_sn, "", "99", "", 0, CONST_VIPABC_R_CONN)
    response.Write (sCDbl(intHavingSession) + sCDbl(intComboGiftSession))
else
    response.Write flt_client_total_session + flt_client_week_session + flt_client_bonus_session
end if
%></span>
