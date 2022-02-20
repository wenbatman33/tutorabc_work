<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<body>
<%
dim clientsn:clientsn=request("clientsn")
dim paymentdetailsn:paymentdetailsn=request("paymentdetailsn")
dim ordernumber:ordernumber=request("ordernumber")
dim number:number=request("number")
dim pay_platform:pay_platform=request("pay_platform")
dim result:result=request("result")
dim dataexist:dataexist=request("dataexist")'0表示可以寄信，1表示不可以寄信
'if (pay_platform <> "") then
    'pay_platform = Replace(pay_platform, "(测试)", "")
'end if
 
'寄送email start**************************************
if (clientsn <> "") and (paymentdetailsn <> "") and (ordernumber <> "") and (number <> "") and (pay_platform <> "") then
    if (result="0") then
	    if (dataexist<>"1") then
            Call SendCardPaymentMail_test(clientsn, paymentdetailsn, ordernumber, "1001", pay_platform)
        end if
	    response.write "<script>alert('刷卡成功');window.location.href = 'http://www.vipabc.com/program/member/buy_1.asp';</script>"
    else
        response.write "<script>alert('刷卡失败，请重试！');</script>"
    end if
else
    response.write "<script>alert('参数错误');</script>"
end if
'寄送email end****************************************
%>
</body>
</html>
