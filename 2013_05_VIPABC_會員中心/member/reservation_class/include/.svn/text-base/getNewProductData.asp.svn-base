<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<!--#include virtual="/lib/functions/functions_contract.asp" -->
<% 

'所有判斷變數均讀DB 避免重復讀取改為session
dim bol1On3Session : bol1On3Session = Session("Having1On3Session")
dim bol1On4Session : bol1On4Session = Session("Having1On4Session")

'session存在 就讀session
if ( Not isEmptyOrNull(bol1On3Session) AND Not isEmptyOrNull(bol1On4Session) ) then
else
    bol1On3Session = false
    bol1On4Session = false
    '否則才使用function讀取db建立session
    'Set objProductSessionData = getProductSessionType(str_now_use_product_sn, 0, CONST_VIPABC_RW_CONN) '課程類型空值表讀出全部
    'while not objProductSessionData.eof
    '    intSessionType = objProductSessionData("sst_number")
    '    if ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = sCInt(intSessionType) ) then
    '        bol1On3Session = true
    '    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) = sCInt(intSessionType) ) then
    '        bol1On4Session = true
    '    end if
    '    objProductSessionData.movenext
    'wend
   
   '201214 阿捨新增 改寫為function
    if ( true = isNewProduct(str_now_use_product_sn, 1, 0, CONST_VIPABC_RW_CONN) ) then
        bol1On4Session = true
    else
        bol1On3Session = true
    end if

    Session("Having1On3Session") = bol1On3Session
    Session("Having1On4Session") = bol1On4Session
end if
%>