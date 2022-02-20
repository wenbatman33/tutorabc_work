<!--#include virtual="/lib/include/global.inc" -->
<%
dim bolTestCase : bolTestCase = false '測試狀態(for QA)
dim bolDebugMode : bolDebugMode = false '除錯模式

'Function addRecordViewList(ByVal p_intClientSn, ByVal p_intCostSession, ByVal p_intRecordSn, ByVal p_strClientIP, _
 '                                              ByVal p_intMaterialSn, ByVal p_intVideoType, ByVal p_intRecordType, ByVal p_intDebugMode, ByVal p_intConnectType)
 '/****************************************************************************************
'描述： 新增錄影檔觀看紀錄
'傳入參數：
'[p_intClientSn : integer] 客戶編號
'[p_intCostSession : integer] 扣堂堂數
'[p_intRecordSn : integer] 錄影檔編號
'[p_strClientIP : string] 使用者IP
'[p_intMaterialSn : integer] 教材編號
'[p_intVideoType : integer] 錄影檔類型 1: MyRecordType, 2: OtherRecordType, 3: LobbyRecordType
'[p_intRecordType : integer] 錄影檔類型 1: 原始, 2: VIPABC贈送
'[p_intDebugMode : integer] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType : Integer] 連線的主機型態
'回傳值：true or false
'示例：
'牽涉數據表：1: sessionrecord_view_list 2: sessionrecord_view_used_free_video_point_list
'備註：
'歷程(作者/日期/原因)：ArthurWu 2012/12/18 Created
'*****************************************************************************************/ 
'Function isRecordDaysDiff(ByVal p_intClientSn, ByVal p_intRecordSn, ByVal p_intDiffDays, ByVal p_intRecordType, ByVal p_intDebugMode, ByVal p_intConnectType)
 '/****************************************************************************************
'描述： 判斷觀看錄影檔時間與今日差距幾天內是否成立 (注意為小於)
'傳入參數：
'[p_intClientSn : integer] 客戶編號
'[p_intRecordSn : integer] 錄影檔編號
'[p_intDiffDays : integer] 差距幾天
'[p_intRecordType : integer] 錄影檔類型 1: 原始, 2: VIPABC贈送
'[p_intDebugMode : integer] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType : Integer] 連線的主機型態
'回傳值：true or false
'示例：
'牽涉數據表：1: sessionrecord_view_list 2:
'備註：
'歷程(作者/日期/原因)：ArthurWu 2012/12/18 Created
'*****************************************************************************************/ 
'Function isHavingFreeRecord(ByVal p_intClientSn, ByVal p_intTypeSn, ByVal p_intDebugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述：取得客戶是否擁有贈送錄影檔 for VIPABC客戶
'傳入參數：
'[p_intClientSn:Integer] 會員編號
'[p_intTypeSn:Integer] 1: 回傳 true or false 2: 回傳剩餘點數
'[p_intDebugMode:Integer] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： ture or false
'牽涉數據表：sessionrecord_view_used_free_video_point_list, client_free_video_point
'備註 ：
'歷程(作者/日期/原因) ：[ArTHurWu] 2013/03/20 Created
'*****************************************************************************************/
dim intClientSn : intClientSn = getRequest("client_sn", CONST_DECODE_NO)
dim intCostSession : intCostSession = getRequest("cost_session", CONST_DECODE_NO)
dim intRecordSn : intRecordSn = getRequest("record_sn", CONST_DECODE_NO)
dim strClientIP : strClientIP = getIpAddress()
dim intMaterialSn : intMaterialSn = getRequest("material_sn", CONST_DECODE_NO)
dim intVideoType : intVideoType = getRequest("video_type", CONST_DECODE_NO)
dim intMethodType : intMethodType = getRequest("method_type", CONST_DECODE_NO)
dim intDiffDays : intDiffDays = CONST_IN_DAYS_FOR_FREE_RECORD
dim bolAddSucc : bolAddSucc = false
dim intRecordType : intRecordType = ""

if ( true = isHavingFreeRecord(intClientSn, 1, 0, CONST_VIPABC_RW_CONN) ) then
	intRecordType = 2 '贈送錄影檔
else
    '再檢查一次贈送錄影檔七日免費觀看
    if ( true = isRecordDaysDiff(intClientSn, intRecordSn, intDiffDays, 2, 0, CONST_VIPABC_RW_CONN) ) then
        intRecordType = 2 '贈送錄影檔
    else
        intRecordType = 1 '原始錄影檔
    end if
end if
	
if ( 1 = sCInt(intMethodType) ) then '新增錄影檔觀看紀錄
	bolAddSucc = addRecordViewList(intClientSn, intCostSession, intRecordSn, strClientIP, intMaterialSn, intVideoType, intRecordType, 0, CONST_VIPABC_RW_CONN)
elseif ( 2 = sCInt(intMethodType) ) then '判斷是否為時間內重復觀看
    if ( true = isRecordDaysDiff(intClientSn, intRecordSn, intDiffDays, intRecordType, 0, CONST_VIPABC_RW_CONN) ) then
        response.Write "free"
    else
        response.Write (intCostSession / CONST_SESSION_TRANSFORM_POINT)
    end if
end if
%>
