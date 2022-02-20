<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    '解密後的安全碼
    Dim str_safe_code : str_safe_code = 1
    Dim str_act, int_client_sn, str_order_fault_msg
    Dim str_class_type, str_reservation_class_date, str_reservation_class_time, str_reservation_class_datetime

    str_act = Trim(Request("act"))            '要執行的動作
    str_class_type = Trim(Request("rctype"))  '訂課種類 一對一 OR 一對多 OR 大會堂    
    str_order_fault_msg = ""                  '訂課失敗的錯誤訊息

    '取得客戶編號
    if (Session("client_sn") <> "") then
        int_client_sn = Trim(Session("client_sn"))
    end if
    
    Response.Clear

    '檢查客戶選到的訂課時間 取得預計扣除課時數
    if (str_act = "3") then
        Dim strLobbySessionSn : strLobbySessionSn = getRequest("lobbysn", CONST_DECODE_NO)
        str_reservation_class_datetime = Trim(Unescape(Request("rcdatetime")))  '客戶選擇的訂課日期和時間
        Response.Clear
        Response.Write getSubtractSession(Array(str_reservation_class_datetime, str_class_type, str_now_use_product_sn, strLobbySessionSn), _
                                          CONST_VIPABC_RW_CONN)


    '回傳HTML控制項<select> 可供客戶訂課的時間
    elseif (str_act = "4") then
        str_reservation_class_date = Trim(Unescape(Request("rcdate")))          '客戶選擇的訂課日期
        Response.Clear
        Response.Write getHtmlSelectClassTime(str_msg_no_sel_class_time, "", str_reservation_class_date, str_class_type, int_client_sn, str_last_first_session_datetime, CONST_VIPABC_RW_CONN)  '请选择咨询时段


    '檢查客戶是否已經預定了此時段的課程
    elseif (str_act = "chkorder") then
        str_reservation_class_date = Trim(Unescape(Request("rcdate")))          '客戶選擇的訂課日期
        str_reservation_class_time = Trim(Unescape(Request("rctime")))          '客戶選擇的訂課時間
        str_reservation_class_sn = Trim(Unescape(Request("rcsn")))          '客戶選擇的lobby sn
        Response.Clear
        
        '已經預定了
        if (isAlreadyLobbyReservation(str_reservation_class_date, str_reservation_class_time, str_reservation_class_sn)) then
            Response.Write "have"
        '沒有預定
        else
            Response.Write "none"
        end if
    end if

%>
