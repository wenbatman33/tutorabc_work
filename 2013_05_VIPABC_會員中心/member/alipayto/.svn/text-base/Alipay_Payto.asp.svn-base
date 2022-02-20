<!--#include virtual="/program/member/alipayto/alipay_Config.asp"-->
<!--#include virtual="/program/member/alipayto/Alipay_md5.asp"-->
<%
'支付寶網址
'dim strTransBeijingDomain : strTransBeijingDomain = CONST_TRANS_BEIJING_DOMAIN
Dim INTERFACE_URL

'if ( "1" = strTransBeijingDomain ) then
    'INTERFACE_URL = "http://www.vipabc.com/program/member/ajax_write_payment_record.asp"
'else
    INTERFACE_URL = "https://www.alipay.com/cooperate/gateway.do?"
'end if

Class creatAlipayItemURL
    Public Function creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
        Dim mystr,Count
        Dim i,j
        Dim minmax,minmaxSlot
        Dim mark,temp,value,md5str,sign,itemURL
        Dim str_service_type : str_service_type = "" 
        ' trade_create_by_buyer or create_direct_pay_by_user   
        ' service="trade_create_by_buyer",标准双接口(又被称为担保交易接口)
        ' service="create_direct_pay_by_user"即时到账接口
        str_service_type = "service=create_direct_pay_by_user"
        mystr = Array(str_service_type,"partner="&partner,"subject="&subject,"body="&body,"out_trade_no="&out_trade_no,"price="&price,"discount="&discount,"show_url="&show_url,"quantity="&quantity,"payment_type=1","seller_email="&seller_email,"notify_url="&notify_url,"return_url="&return_url)
        Count = ubound(mystr)
        For i = Count TO 0 Step -1
            minmax = mystr( 0 )
            minmaxSlot = 0
            For j = 1 To i
                mark = (mystr( j ) > minmax)
                If mark Then 
                    minmax = mystr( j )
                    minmaxSlot = j
                End If
            Next
            If minmaxSlot <> i Then 
                temp = mystr( minmaxSlot )
                mystr( minmaxSlot ) = mystr( i )
                mystr( i ) = temp
            End If
        Next

        For j = 0 To Count Step 1
            value = SPLIT(mystr( j ), "=")
            If  value(1)<>"" then
                If j = Count Then
                    md5str = md5str & mystr( j )
                Else 
                    md5str = md5str & mystr( j ) & "&"
                End If 
            End If 
        Next

        md5str = md5str & key
        sign = md5(md5str)
        itemURL = itemURL & INTERFACE_URL

        For j = 0 To Count Step 1
            value = SPLIT(mystr( j ), "=")
            If  value(1)<>"" then
                itemURL = itemURL&mystr( j )&"&"
            End If 	     
        Next

        itemURL = itemURL & "sign=" & sign & "&sign_type=" & "MD5"
        creatAlipayItemURL= itemURL
    End Function
End Class
%>