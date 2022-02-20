<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/alipayto/Alipay_md5.asp"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<%
Dim str_hitrust_sn : str_hitrust_sn = ""   ' hitrust_record_all的sn
Dim str_score : str_score = "" '成功與否
Dim str_sql 'sql
Dim var_arr '傳入sql的陣列資訊
Dim arr_result 'sql執行的結果
Dim str_client_sn : str_client_sn = "" '客戶的client_sn
Dim str_contract_sn : str_contract_sn = "" '合約編號
Dim key : key = "" '支付宝安全教研码
Dim partner : partner = "" '支付宝合作
Dim total_fee : total_fee = "" '支付的总价格
Dim receive_name : receive_name = "" '收货人姓名
Dim receive_address : receive_address = "" '收货人地址
Dim receive_zip : receive_zip = "" '收货人邮编
Dim receive_phone : receive_phone = "" '收货人电话
Dim receive_mobile : receive_mobile = "" '收货人手机
Dim alipayNotifyURL : alipayNotifyURL = "" '支付寶網址
Dim Retrieval '物件
Dim ResponseTxt '回傳值
Dim i,j
Dim str_tablecol
Dim str_change_flag '資料是否改變成功
Dim str_email_flag '寄信是否成功
Dim varItem
Dim mark
Dim value
Dim str_dsn_sn : str_dsn_sn = "" 'detail_sn
Dim str_status : str_status = "" '狀態
Dim returnTxt : returnTxt = "success"

str_hitrust_sn = DelStr(Request("subject")) '获取定单号
'total_fee = DelStr(Request("total_fee")) '获取支付的总价格
'receive_name = DelStr(Request("receive_name")) '获取收货人姓名
'receive_address = DelStr(Request("receive_address")) '获取收货人地址
'receive_zip = DelStr(Request("receive_zip")) '获取收货人邮编
'receive_phone = DelStr(Request("receive_phone")) '获取收货人电话
'receive_mobile = DelStr(Request("receive_mobile")) '获取收货人手机

'if ( g_var_client_sn = "223543" ) then
    intContractSn = str_contract_sn
	intHitrustSn = str_hitrust_sn
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<%
'end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_ALIPAY_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_ALIPAY_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_ALIPAY_NAME
end if

key = CONST_CREDIT_CARD_ALIPAY_KEY ' 支付宝安全教研码
partner = CONST_CREDIT_CARD_ALIPAY_CODE '支付宝合作id

'******************************************判断消息是不是支付宝发出
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
Retrieval.setOption 2, 13056 
Retrieval.open "GET", alipayNotifyURL, False, "", "" 
Retrieval.send()
ResponseTxt = Retrieval.ResponseText
Set Retrieval = Nothing
'*****************************************
Dim mystr, Count, minmax, minmaxSlot, temp, md5str, mysign
'获取支付宝GET过来通知消息,判断消息是不是被修改过

For Each varItem in Request.Form
    mystr = varItem & "=" & Request(varItem) & "^" & mystr
Next 

If ( mystr <> "" ) then 
    mystr = Left(mystr,Len(mystr)-1)
End If 

mystr = SPLIT(mystr, "^")
Count = ubound(mystr)

'对参数排序
For i = Count TO 0 Step -1
    minmax = mystr( 0 )
    minmaxSlot = 0
    For j = 1 To i
        mark = (mystr( j ) > minmax)
        If mark then 
            minmax = mystr( j )
            minmaxSlot = j
        End If 
    Next
    
    If minmaxSlot <> i then 
        temp = mystr( minmaxSlot )
        mystr( minmaxSlot ) = mystr( i )
        mystr( i ) = temp
    End If
Next

'构造md5摘要字符串
For j = 0 To Count Step 1
    value = SPLIT(mystr( j ), "=")
    If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  then
        If j=Count then
        md5str = md5str&mystr( j )
        Else 
        md5str = md5str&mystr( j )&"&"
        End If 
    End If 
Next

md5str = md5str & key
mysign = md5(md5str)
'********************************************************

'**************如果付款成功****************start**************
If ( mysign = request.Form("sign") AND "true" = ResponseTxt ) then 		
	'更新狀態為刷卡成功
	str_status = CONST_CREDIT_CARD_ALIPAY_PAY_SUCCEED_CODE
	'如果成功寫回hitrust_record_all update 欄位
	str_tablecol = "retcode=@retcode, out_date=getdate(), ReturnUrl=@ReturnUrl, type=@type"
	var_arr = Array("0@" & str_status, NULL, strPlamName, str_hitrust_sn)
	arr_result = excuteSqlStatementWrite(" UPDATE hitrust_record_all SET " & str_tablecol & " WHERE sn=@sn", var_arr, CONST_VIPABC_RW_CONN)
	'選出該筆刷卡紀錄，取出dsn
	str_sql = " SELECT TOP 1 payment_detail_sn, account_sn FROM hitrust_record_all WHERE sn = @sn "
	var_arr = Array(str_hitrust_sn)
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
	if ( isSelectQuerySuccess(arr_result) ) then		
		'如果有查尋hitrust_record_all資料
		if ( Ubound(arr_result) >= 0 ) then 
			'回傳detail_sn
			str_dsn_sn = arr_result(0,0)
			str_contract_sn = arr_result(1,0)
			'根據detail_sn查出client_sn
			str_sql = " SELECT TOP 1 client_sn FROM contract_pay WHERE dsn = @dsn AND isbuy = @isbuy "	
			var_arr = Array(str_dsn_sn,"0")
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
			'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
			if ( isSelectQuerySuccess(arr_result) ) then
				'有資料
				if ( Ubound(arr_result) >= 0 ) then 
					str_client_sn = arr_result(0,0)
                    '---------- 跟新相關資訊 client_payment_detail,client_temporal_contract------ start --------
	                if ( str_contract_sn <> "" OR str_dsn_sn <> "" ) then	
		                str_change_flag = changeCardPaymentDate(str_contract_sn, str_dsn_sn)
		                'response.write str_change_flag
		                '----------------寄給業務  &  主管 & 財務部	    
		                '--------------- 刷卡成功寄信 --------------- start ------------
		                'str_email_flag = SendCardPaymentMail(str_client_sn,str_dsn_sn,str_status, strPlamName)
                        str_email_flag = SendCardPaymentMail_test(str_client_sn, str_dsn_sn, str_hitrust_sn, str_status, strPlamName)            	
		                'response.write str_email_flag
		                '--------------- 刷卡成功寄信 --------------- end ------------
	                else
		                response.write "fail"
	                end if
                    '返回數據給支付寶
                    Response.write "success"			
				else
                    '錯誤訊息
					'response.write "ccc........." &  arr_result(0,0)& "<br>"
				end if 
			end if 			
		else
			'錯誤訊息
			'response.write "ccc........." &  arr_result(0,0)& "<br>"
		end if 
	end if		
Else
    Response.end '这里可以指定你需要显示的内容
End If

'取代特殊字元
Function DelStr(Str)
    If IsNull(Str) Or IsEmpty(Str) then
	    Str = ""
    End If
    DelStr = Replace(Str,";","")
    DelStr = Replace(DelStr,"'","")
    DelStr = Replace(DelStr,"&","")
    DelStr = Replace(DelStr," ","")
    DelStr = Replace(DelStr,"　","")
    DelStr = Replace(DelStr,"%20","")
    DelStr = Replace(DelStr,"--","")
    DelStr = Replace(DelStr,"==","")
    DelStr = Replace(DelStr,"<","")
    DelStr = Replace(DelStr,">","")
    DelStr = Replace(DelStr,"%","")
End Function
%>