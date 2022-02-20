<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<html>
<table width="100%" align="center"><img id="img_ajax_loader" src="/lib/javascript/JQuery/images/ajax-loader.gif" border="0" /></table>

<%
' We graf all the session variable names/values and stick them in a form
' and then we submit the form to our receiving ASP.NET page (ASPNETPage_Session.aspx)...

Response.Write "<form name=""t"" id=""t"" action=""./VideoShow"&strVIPABCJR&"/saveAspSession.aspx"" method=""post"">"
'For each Item in Session.Contents
    'if( Cstr(Item) = "client_sn" ) then
        'Response.Write("<input type='hidden' name='" & Cstr(Item) &"' value='" & Session.Contents(item) & "' />")
    'end if
	
	'Response.Write( " value='" & Session.Contents(item) & "' / >")
    'Response.Write( item) )
    'Response.Write( Session.Contents(item) )
'next
Response.Write "<input type=""hidden"" name=""client_sn"" value=""" & Session("client_sn") & """ />"
'Response.Write "<input type='submit' value='送出' style='display:none'>"
Response.Write "</form>" 
Response.Write "<script type=""text/javascript"">t.submit();</script>"
%>
</html>