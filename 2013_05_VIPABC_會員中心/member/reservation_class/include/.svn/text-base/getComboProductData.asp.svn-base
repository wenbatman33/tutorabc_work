<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<% 

'所有判斷變數均讀DB 避免重復讀取改為session
'是否為套餐產品
dim bolComboProduct : bolComboProduct = Session("ComboProduct")
'session存在 就讀session
if ( Not isEmptyOrNull(bolComboProduct) ) then
'否則才使用function讀取db建立session
else
    bolComboProduct = isComboProduct(str_now_use_product_sn, 0, CONST_VIPABC_R_CONN)
    if ( true = bolDebugMode ) then
        response.Write "bolComboProduct: " & bolComboProduct & "<br />"
    end if  
    Session("ComboProduct") = bolComboProduct
end if

'如果是套餐 判斷每種課程類型的權限 
if ( true = bolComboProduct ) then
    'session存在 就讀session
    intTotalNum = 0
    dim objComboSessionType : objComboSessionType = null
    if  ( true = IsObject(Session("ComboSessionType")) ) then
        Set objComboSessionType = Session("ComboSessionType") '是字典檔
        intTotalNum = objComboSessionType.Count
    end if
    
    if ( 0 < intTotalNum ) then
    '否則才使用function讀取db建立session
    else
        '20121109 阿捨新增 改讀課程類型資料庫
        Set objProductTypeData = getProductTypeData("", "C", 0, CONST_VIPABC_R_CONN) '課程類型空值表讀出全部
        Set objComboSessionType = CreateObject("Scripting.Dictionary")
        while not objProductTypeData.eof
            intDataClassType = objProductTypeData("sst_number")
            strDataClassEName = objProductTypeData("sst_description_en")
            objComboSessionType(intDataClassType) = strDataClassEName
            if ( true = bolDebugMode ) then
                response.Write "intDataClassType: " & intDataClassType & "<br />"
                response.Write "strDataClassEName: " & strDataClassEName & "<br />"
            end if  
        objProductTypeData.movenext
        wend
        set Session("ComboSessionType") = objComboSessionType
    end if
    intTotalNum = objComboSessionType.Count

    '字典檔存在時
    if ( 0 < intTotalNum ) then
        dim strKeyWord : strKeyWord = objComboSessionType.Keys '字典檔的key 也是課程類型編號
        dim strComboHavingSessionType : strComboHavingSessionType = ""
        dim strComboSessionTypeEName : strComboSessionTypeEName = ""
        dim intComboSessionType : intComboSessionType = ""
        dim strDimCode : strDimCode = "" '動態變數字串
        for each strKeyWord in objComboSessionType.Keys
            strComboSessionTypeEName = objComboSessionType.Item(strKeyWord) '字典檔的value 也是課程類型的EN_word for 動態變數用
            intComboSessionType = strKeyWord
             if ( true = bolDebugMode ) then
                response.Write "strComboSessionTypeEName: " & strComboSessionTypeEName & "<br />"
                response.Write "intComboSessionType: " & intComboSessionType & "<br />"
            end if  
            if ( Not isEmptyOrNull(intComboSessionType) AND Not isEmptyOrNull(strComboSessionTypeEName) ) then
                strComboHavingSessionType = Session("ComboHaving" & strComboSessionTypeEName)
                strDimCode = ""
                'session存在 就讀session
                if ( Not isEmptyOrNull(strComboHavingSessionType) ) then
                    strDimCode = "dim bolComboHaving" & strComboSessionTypeEName & " : bolComboHaving" & strComboSessionTypeEName & " = " & strComboHavingSessionType
                else
                    '否則才使用function讀取db建立session
                    intMaxSession = getComboProductMaxSession(str_now_use_product_sn, intComboSessionType, 0, CONST_VIPABC_R_CONN)
                    if ( 0 < sCInt(intMaxSession) ) then
                        'strComboHavingSessionType 反正不存在, 就拿來用
                        strComboHavingSessionType = "true"
                    else
                        strComboHavingSessionType = "false"
                    end if 
                    strDimCode = "dim bolComboHaving" & strComboSessionTypeEName & " : bolComboHaving" & strComboSessionTypeEName & " = " & strComboHavingSessionType
                    Session("ComboHaving" & strComboSessionTypeEName) = strComboHavingSessionType
                end if
                if ( true = bolDebugMode ) then
                    response.Write "strDimCode: " & strDimCode & "<br />"
                end if
                '最後執行動態變數字串 宣告相關權限變數
                Execute(strDimCode) 
            end if
        next
    end if 'if ( 0 < intTotalNum ) then

end if
%>