<%       
'20110614 阿捨新增 特殊客戶關懷
intClientSn = getSession("client_sn", CONST_DECODE_NO)
dim strMsg : strMsg = ""
strEmailbody = ""
strEmailSubject = ""
dim bolSpecialVIPCare : bolSpecialVIPCare = false
dim bolSpecialSpyCare : bolSpecialSpyCare = false
dim bolTestAccount : bolTestAccount = false

dim objLeaderCellPhone : objLeaderCellPhone = null
dim objLeaderCellPhoneNode : objLeaderCellPhoneNode = null
dim arrSpecialCareClientSn : arrSpecialCareClientSn = null
dim arrSpecialCareClientCName : arrSpecialCareClientCName = null

dim strSMSKey : strSMSKey = ""

if ( true = bolDeBugMode ) then
    response.Write "intClientSn: " & intClientSn & "<br />"
end if

if ( true = bolTestCase ) then
    intClientSn = 223543
end if

if ( true = bolDeBugMode ) then
    response.Write "intClientSn: " & intClientSn & "<br />"
end if

arrSpecialCareClientSn = Array("263739", "283768", "283532", "281827", "304043", "304046", "304049", "305795", "223543")
arrSpecialCareClientCName = Array("彭斯光", "李伟", "金焱", "钱婷", "Jason Zhang", "Doris Pan", "Margaret Feng", "Boyu Hu", "測試信用卡")
'if ( intClientSn = "263739" OR intClientSn = "357409" OR intClientSn = "283768" OR intClientSn = "283532" OR intClientSn = "281827" OR intClientSn = "304043" OR intClientSn = "304046" OR intClientSn = "304049" OR intClientSn = "223543" OR intClientSn = "305795" ) then
if ( isArray(arrSpecialCareClientSn) ) then
    intTotal = Ubound(arrSpecialCareClientSn)
    if ( true = bolDeBugMode ) then
        response.Write "intTotal: " & intTotal & "<br />"
    end if
    for intI = 0 to intTotal
        intSpecialCareClientSn = arrSpecialCareClientSn(intI)
        if ( true = bolDeBugMode ) then
            response.Write "intSpecialCareClientSn: " & intSpecialCareClientSn & "<br />"
            response.Write "intClientSn: " & intClientSn & "<br />"
        end if
        if ( sCLng(intSpecialCareClientSn) = sCLng(intClientSn) ) then
            strSendCName = sCStr(arrSpecialCareClientCName(intI))
            strEmailToAddr = CONST_MAIL_TO_CS_SPECIAL_CARE_CUSTOMER_ACCOUNT '預設為CS重大抱怨客戶
            
            '先判斷客戶通報身分
            select case sCLng(intClientSn)
                case 263739
                    bolSpecialSpyCare = true
                case 304043
                    bolSpecialVIPCare = true
                case 304046
                    bolSpecialVIPCare = true  
                case 304049
                    bolSpecialVIPCare = true
                case 305795
                    bolSpecialVIPCare = true
                case 223543
                    bolTestAccount = false
                    strEmailToAddr = "arthurwu@vipabc.com"
            end select

            '根據身分選擇不同群組寄發
            if ( true = bolSpecialSpyCare ) then
                strEmailToAddr = CONST_MAIL_TO_SPECIAL_CARE_CUSTOMER_ACCOUNT 'SPY客戶
            elseif ( true = bolSpecialVIPCare ) then
                strEmailToAddr = CONST_MAIL_TO_CS_GIS_SPECIAL_CARE_CUSTOMER_ACCOUNT 'VIP客戶
            end if
           
            if ( true = bolDeBugMode ) then
                response.Write "bolTestAccount: " & bolTestAccount & "<br />"
            end if
           
           '判斷通知類型
            'if ( 1 = intNoifiyType ) then '登入通知
            strEmailSubject = "VIPABC-特殊客戶關懷通知信"
            select case (intNoifiyType)
                case 1  '登入通知
                    strMsg = strSendCName & "於" & now() & "登入VIPABC官网,惟请特别留意"
                    strEmailbody = "Dear All:<br />"
                    strEmailbody = strEmailbody & strMsg & "！<br />"
                    '登入通知for VIP客戶不同 沒有GM
                    if ( true = bolSpecialVIPCare ) then
                        strEmailToAddr = CONST_MAIL_TO_CS_GIS_NO_GM_SPECIAL_CARE_CUSTOMER_ACCOUNT '登入通知noGM
                    end if
                case 2 '訂課/取消通知
                    strMsg = strSendCName & "剛" 
                    if ( bolOrder ) then
                        strMsg = strMsg & "預定" & intOrderClassNum & "堂" & strClassType & strOrderClassInfo
                    end if
                    if ( bolCancel ) then
                        strMsg = strMsg & "取消" & intCancelClassNum & "堂 " & strClassType & strCancelClassInfo
                    end if
                    strEmailbody = "Dear All:<br />"
                    strEmailbody = strEmailbody & strMsg & "，惟請特別留意！<br />"
                    strSMSKey = "vipcareclass"
                case 3 '課後評鑑通知
					strMsg = strSendCName & "送出" & strSessionDate & "課後評鑑,顧問:" & intConsultantScore & ",教材:" & intMaterialsScore & ",通訊:" & intOverallScore
                    if ( Not isEmptyOrNull(strSuggestions) ) then
                        strMsg = strMsg & ",建議:" & strSuggestions
                    end if
					if ( Not isEmptyOrNull(strCompliment) ) then
                        strMsg = strMsg & ",讚美:" & strCompliment
                    end if
                    strEmailbody = "Dear All:<br />"
                    strEmailbody = strEmailbody & strMsg & "<br />"
                    strSMSKey = "vipcarefeedback"
            end select
            strMsg = unescape(strMsg)
            'end if

            if ( Not isEmptyOrNull(strEmailbody) ) then '發郵件
                strEmailFormAddr = CONST_MAIL_ADDRESS_SUPPORTS_ABC
                arrEmailFromParas = Array(strEmailFormAddr, "VIPABC", strEmailSubject, strEmailbody, "", "", "")
                arrEmailToParas = Array(strEmailToAddr, "系統判定", "", "")
                arrSetParas = Array(CONST_MAIL_SERVER_STAFF, "utf-8", "", "", "")
                arrOptParas = Array(true, false, false, "", "", "", "")

                g_int_err_code = sendEmail(arrEmailFromParas, arrEmailToParas, arrSetParas, arrOptParas, CONST_TUTORABC_WEBSITE)
                if ( g_int_err_code = CONST_FUNC_EXE_SUCCESS ) then
                else
                end if
            end if 'if ( Not isEmptyOrNull(strEmailbody) ) then

            if ( Not isEmptyOrNull(strMsg) ) then '發簡訊
                Set objLeaderCellPhone = Server.CreateObject("MSXML2.DOMDocument")
                if ( true = objLeaderCellPhone.load(Server.Mappath("/xml/LeaderCellPhone.xml")) ) then                                    
                    Set root = objLeaderCellPhone.documentElement 
                    for each objLeaderCellPhoneNode in root.selectnodes("/Leader/CellPhone")    
                        intStaffStatus = objLeaderCellPhoneNode.getAttribute("status")
		                intStaffStatus = sCInt(intStaffStatus)
                        strCellPhone = objLeaderCellPhoneNode.getAttribute("value")
                        intTestCase = objLeaderCellPhoneNode.getAttribute("test")
		                intTestCase = sCInt(intTestCase)
                        if ( Not isEmptyOrNull(strSMSKey) ) then
                            intSendSMS = objLeaderCellPhoneNode.getAttribute(strSMSKey)
                            intSendSMS = sCInt(intSendSMS)
                        end if
                        intCNPhone = objLeaderCellPhoneNode.getAttribute("cnphone")
                        intCNPhone = sCInt(intCNPhone)
                        
                        if ( true = bolDeBugMode ) then
                            response.Write "intTestCase: " & intTestCase & "<br />"
                            response.Write "strSMSKey: " & strSMSKey & "<br />"
                            response.Write "intSendSMS: " & intSendSMS & "<br />"
                        end if
                        
                        if ( true = bolSpecialVIPCare AND 1 = intSendSMS ) then
                            if ( 1 = intStaffStatus ) then '必須在職
                                '內地手機
                                if ( 1 = intCNPhone ) then 
                                    Call SmsDxtonSend(strCellPhone, strMsg, "", "")
                                else
                                    Call SmsChtSend(strCellPhone, strMsg, "", "")
                                end if
                            end if 'if ( 1 = intStaffStatus ) then
                        end if 'if ( true = bolSpecialVIPCare AND 1 = intVIPCareLogin ) then

                        if ( true =  bolTestAccount AND 1 = intTestCase ) then
                            if ( 1 = intStaffStatus ) then '必須在職
                                '內地手機
                                if ( true = bolDeBugMode ) then
                                    response.Write "intStaffStatus: " & intStaffStatus & "<br />"
                                end if
                                if ( 1 = intCNPhone ) then 
                                    Call SmsDxtonSend(strCellPhone, strMsg, "", "")
                                else
                                    Call SmsChtSend(strCellPhone, strMsg, "", "")
                                end if
                            end if 'if ( 1 = intStaffStatus ) then
                        end if 'if ( true = bolSpecialVIPCare AND 1 = intVIPCareLogin ) then

                    next 'for each objLeaderCellPhoneNode in root.selectnodes("/Leader/CellPhone")
                end if 'if ( true = objLeaderCellPhone.load(Server.Mappath("/xml/LeaderCellPhone.xml")) ) then 
            end if 'if ( Not isEmptyOrNull(strMsg) ) then

        end if 'if ( isArray(arrSpecialCareClientSn) ) then
    next 'for intI = 0 to intTotal
end if 'if ( isArray(arrSpecialCareClientSn) ) then
 %>