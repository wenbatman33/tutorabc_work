﻿<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/VIPABC_ContractHtml/functions_VIPABC_ContractHtml.asp" -->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<!--#include virtual="/lib/functions/functions_class.asp" -->
<% 
dim bolDebugMode : bolDebugMode = false

dim objContractHtml : objContractHtml = null
dim intContractSn : intContractSn = Session("contract_sn")
if ( isEmptyOrNull(intContractSn) ) then 
    intContractSn = getRequest("contract_sn", CONST_DECODE_NO)
end if
dim strWholeHtmlContent : strWholeHtmlContent = ""
Set objContractHtml = New ContractHtml
strWholeHtmlContent = objContractHtml.getContractContent(intContractSn, 0, CONST_VIPABC_RW_CONN)
%>
<%=strWholeHtmlContent%>