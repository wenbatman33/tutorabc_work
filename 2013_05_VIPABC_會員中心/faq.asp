<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/include/html/template/template_change_keyword.asp"-->
<%
'定義變數
Dim str_link_Hash_Point  '接傳送過來的錨點(hash point)
Dim arr_link_Hash_Point

str_link_Hash_Point = getRequest("linkHashPoint", CONST_DECODE_NO)





'從會員右側欄點來的,請出現這一行的function
function getGoBackButtonBySidebar()

Dim str_http_referer_by_faq 
str_http_referer_by_faq = Request.serverVariables("HTTP_REFERER")		'接收來源頁面

	if (str_http_referer_by_faq<>"") then 
		response.Write "<input type=""button"" name=""btn"" value=""回上一頁"" class=""btn_2"" onClick=""javascript:location.href='"&str_http_referer_by_faq&"'""/>"
	end if	

end function





%>
<script type="text/javascript">
$(document).ready(function(){
	//  這裡就是載入HTML控制項完後立即會執行的 可以在這裡面寫 function 或 其他有的沒的
			
	<% 
		if (str_link_Hash_Point<>"") then 
			arr_link_Hash_Point = Split(str_link_Hash_Point,"_") 
	%>
		var divName = "<%=arr_link_Hash_Point(0)&"_"&arr_link_Hash_Point(1)%>" ;
		$("#<%=arr_link_Hash_Point(0)&"_"&arr_link_Hash_Point(1)%>_answer_div").css({ display: "" }); 
		location.hash = '<%=str_link_Hash_Point%>';
		$("#"+divName+"_atag").text("关闭");
		$("#"+divName+"_question_div").removeClass();
		$("#"+divName+"_question_div").addClass("f-right faq-btn-close");
	<%
		end if
	%>
});
</script>
<script type='text/javascript' language="javascript" src="/lib/javascript/framepage/js/fag.js"></script>


<div id="temp_contant">
       <!--內容start-->
        <div class='page_title_2'><h2 class='page_title_h2'>上课记录&评分</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 

        <div class="main_faq">
            <!--1類start-->
            <div class="main_faqbox">
            <div>
            <div class="cate color_1"><%=getWord("faq_1")%></div>
            <div class="f-right faq-btn-open" id="class_kind_question_div">
            <a id="class_kind_atag" href="javascript:openAndCloseDiv('class_kind');" title="<%=getWord("faq_2")%>" ><%=getWord("faq_2")%></a></div>
            <div class="clear"></div>
            </div>
             <!--quest start-->
             <div class="question">
             <div class="top1"></div>
             <!--1-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('class_kind','q1');"><%=getWord("faq_3")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('class_kind','q2');"><%=getWord("faq_4")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--1-->
             <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('class_kind','q3');"><%=getWord("faq_5")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('class_kind','q4');"><%=getWord("faq_6")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
             </div>
             <div class="line"></div>
             <!--quest end-->
             <div class="answer" id="class_kind_answer_div" style="display:none;">
             <div class="box">
             <!--a1-->
             <div class="say">
             <p><a name="class_kind_q1"><%=getWord("faq_3")%></a></p>
             <%=getWord("faq_7")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a1-->
             <!--a2-->
             <div class="say">
             <p><a name="class_kind_q2"><%=getWord("faq_4")%></a></p>
             <%=getWord("faq_8")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a2-->
             <!--a3-->
             <div class="say">
             <p><a name="class_kind_q3"><%=getWord("faq_5")%></a></p>
              <%=getWord("faq_9")%><br><%=getWord("faq_10")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a3-->
             <!--a4-->
             <div class="say">
             <p><a name="class_kind_q4"><%=getWord("faq_6")%></a></p>
             <%=getWord("faq_11")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a4-->
             </div>
             <div class="ps"> <%if (isPayMember()) then %><%=getWord("faq_12")%><a href="/program/contact/contact_us.asp"><%=getWord("faq_13")%></a>
			<!--<%=getWord("faq_14")%><a href="#"><%=getWord("faq_15")%></a>--><%end if%>
             </div>        
             </div>
             
             </div>
             <!--1類end-->
             <!--2類start-->
              <div class="main_faqbox">
            <div>
            <div class="cate color_1"><%=getWord("faq_16")%></div>
            <div class="f-right faq-btn-open" id="web_kind_question_div">
            <a id="web_kind_atag" href="javascript:openAndCloseDiv('web_kind');" title="<%=getWord("faq_2")%>" ><%=getWord("faq_2")%></a></div>
            <div class="clear"></div>
            </div>
             <!--quest start-->
             <div class="question">
             <div class="top1"></div>
             <!--1-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q1');"><%=getWord("faq_17")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q2');"><%=getWord("faq_18")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--1-->
             <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q3');"><%=getWord("faq_19")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q4');"><%=getWord("faq_20")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
             <!--3-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q5');"><%=getWord("faq_21")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q6');"><%=getWord("faq_22")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--3-->
             <!--4-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q7');"><%=getWord("faq_23")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('web_kind','q8');"><%=getWord("faq_24")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--4-->
             </div>
             <div class="line"></div>
             <!--quest end-->
             
             <div class="answer" id="web_kind_answer_div" style="display:none;">
             <div class="box">
             <!--a1-->
             <div class="say">
             <p><a name="web_kind_q1"><%=getWord("faq_17")%></a></p>
             <%=getWord("faq_25")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a1-->
             <!--a2-->
             <div class="say">
             <p><a name="web_kind_q2"><%=getWord("faq_18")%></a></p>
             <%=getWord("faq_26")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a2-->
             <!--a3-->
             <div class="say">
             <p><a name="web_kind_q3"><%=getWord("faq_19")%></a></p>
             <%=getWord("faq_27")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a3-->
             <!--a4-->
             <div class="say">
             <p><a name="web_kind_q4"><%=getWord("faq_20")%></a></p>
             <%=getWord("faq_28")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a4-->
             <!--a5-->
             <div class="say">
             <p><a name="web_kind_q5"><%=getWord("faq_21")%></a></p>
             <%=getWord("faq_29")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a5-->
             <!--a6-->
             <div class="say">
             <p><a name="web_kind_q6"><%=getWord("faq_22")%></a></p>
             <%=getWord("faq_30")%><br />
             <%=getWord("faq_34")%><br />
             <%=getWord("faq_35")%><br />
             <%=getWord("faq_36")%><br />
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a6-->
             <!--a7-->
             <div class="say">
             <p><a name="web_kind_q7"><%=getWord("faq_23")%></a></p>
             <%=getWord("faq_31")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a7-->
             <!--a8-->
             <div class="say">
             <p><a name="web_kind_q8"><%=getWord("faq_24")%></a></p>
             <%=getWord("faq_32")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a8-->
             </div>
             <div class="ps"> <%if (isPayMember()) then %><%=getWord("faq_12")%><a href="/program/contact/contact_us.asp"><%=getWord("faq_13")%></a>
			<!--<%=getWord("faq_14")%><a href="#"><%=getWord("faq_15")%></a>--><%end if%>
             </div>        
             </div>
             
             </div>
             <!--2類end-->
             
             <!--3類start-->
             <%
			 'deney說先mark起來 modify by aboo at 20100225
			 Dim mark : mark = true
			 if (mark) then 
			 %>
			<div class="main_faqbox">
            <div>
            <div class="cate color_1"><%=getWord("faq_cuatomer_32")%></div>
            <div class="f-right faq-btn-open" id="member_kind_question_div">
            <a id="member_kind_atag" href="javascript:openAndCloseDiv('member_kind');" title="打开" ><%=getWord("faq_2")%></a></div>
            <div class="clear"></div>
            </div>
             <!--quest start-->
             <div class="question">
             <div class="top1"></div>
             <!--1-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q1');"><%=getWord("faq_cuatomer_33")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q2');"><%=getWord("faq_cuatomer_34") & "？"%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--1-->
             <!--2-->
             <div class="list">
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q3');"><%=getWord("faq_cuatomer_35")%></a></div>
             </div>
			 -->
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q4');"><%=getWord("faq_cuatomer_36")%></a></div>
             </div>
			 -->
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q5');"><%=getWord("faq_cuatomer_37")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q6');"><%=getWord("faq_cuatomer_38")%></a></div>
             </div>
			 -->
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q7');"><%=getWord("CLASS_VOCABULARY_13")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q8');"><%=getWord("faq_cuatomer_39")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q9');"><%=getWord("faq_cuatomer_40")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q10');"><%=getWord("faq_cuatomer_41")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q11');"><%=getWord("faq_cuatomer_42")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q12');"><%=getWord("faq_cuatomer_43")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q13');"><%=getWord("faq_cuatomer_44")%> </a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q14');"><%=getWord("faq_cuatomer_45")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q15');"><%=getWord("faq_cuatomer_46")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q16');"><%=getWord("faq_cuatomer_47")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q17');"><%=getWord("faq_cuatomer_48")%></a></div>
             </div>
			 
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q27');"><%=getWord("faq_cuatomer_57")%></a></div>
             </div>
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q18');"><%=getWord("faq_cuatomer_49")%></a></div>
             </div>
			 -->
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">			 
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q29');"><%=getWord("faq_cuatomer_59")%></a></div>	
			 </div>	
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q19');"><%=getWord("faq_cuatomer_50")%> </a></div>
             </div>
			 -->
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q20');"><%=getWord("faq_cuatomer_51")%> </a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q21');"><%=getWord("faq_cuatomer_52")%> </a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q22');"><%=getWord("FAQ_5")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q23');"><%=getWord("faq_cuatomer_53")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q24');"><%=getWord("faq_cuatomer_54")%></a></div>
             </div>
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
			 		 
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q28');"><%=getWord("faq_cuatomer_58")%></a></div>
             </div>
			 <!--
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q25');"><%=getWord("faq_cuatomer_55")%></a></div>
             </div>
			  -->
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q26');"><%=getWord("faq_cuatomer_56")%></a></div>
             </div>
			 -->
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             <div class="clear"></div>
             </div>
             <!--2-->
			 <!--2-->
             <div class="list">
             </div>
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q30');"><%=getWord("faq_cuatomer_60")%></a></div>
             </div>
             <div class="clear"></div>
             </div>-->
             <!--2-->
			 <!--2-->
             <div class="list">
			 <!--todo 暫時拿掉
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q31');"><%=getWord("faq_cuatomer_61")%></a></div>
             </div>
             <div class="q">
             <div class="w"><a href="javascript:gotoHashPoint('member_kind','q32');"><%=getWord("faq_cuatomer_62")%></a></div>
             </div>
			 -->
             <div class="clear"></div>
             </div>
             <!--2-->
			 <div class="line"></div>
             </div>
             <!--quest end-->
             
             
             <div class="answer" id="member_kind_answer_div" style="display:none;">
             <div class="box">
             <!--a1-->
             <div class="say">
             <p><a name="member_kind_q1"><%=getWord("FAQ_CUATOMER_33")%></a></p>
             <%=getWord("FAQ_CUATOMER_63")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a1-->
             <!--a2-->
             <div class="say">
             <p><a name="member_kind_q2"><%=getWord("FAQ_CUATOMER_34")%></a></p>
             <%=getWord("FAQ_CUATOMER_64")%><br /><%=getWord("FAQ_CUATOMER_65")%><br />
			 <%=getWord("FAQ_CUATOMER_66")%><br /><br />
			 <%=getWord("FAQ_CUATOMER_67")%><br /><%=getWord("FAQ_CUATOMER_68")%><br /><%=getWord("FAQ_CUATOMER_69")%><br />
			 
			 <!--todo 暫時拿除
			 <%=getWord("FAQ_CUATOMER_70")%><br /><%=getWord("FAQ_CUATOMER_71")%><br /><%=getWord("FAQ_CUATOMER_72")%><br /><%=getWord("FAQ_CUATOMER_73")%>
			 -->
			 <br /><br />
			 <%=getWord("FAQ_CUATOMER_74")%><br /><%=getWord("FAQ_CUATOMER_75")%><br /><%=getWord("FAQ_CUATOMER_66")%><br /><br />
			 <%=getWord("FAQ_CUATOMER_76")%><br /><%=getWord("FAQ_CUATOMER_77")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a2-->
             <!--a3-->
			 <!--todo 暫時拿除
             <div class="say">
             <p><a name="member_kind_q3"><%=getWord("FAQ_CUATOMER_35")%></a></p>
             <%=getWord("FAQ_CUATOMER_78")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a3-->
             <!--a4-->
			 <!--todo 暫時拿除
             <div class="say">
             <p><a name="member_kind_q4"><%=getWord("FAQ_CUATOMER_36")%></a></p>
             <%=getWord("FAQ_CUATOMER_79")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a4-->
			 <!--a5-->
			 <!--todo 暫時拿除
             <div class="say">
             <p><a name="member_kind_q5"><%=getWord("FAQ_CUATOMER_37")%></a></p>
             <%=getWord("FAQ_CUATOMER_80")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a5-->
			 <!--a6-->
			 <!--todo 暫時拿除
             <div class="say">
             <p><a name="member_kind_q6"><%=getWord("FAQ_CUATOMER_38")%></a></p>
             <%=getWord("FAQ_CUATOMER_81")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a6-->
			 <!--a7-->
             <div class="say">
             <p><a name="member_kind_q7"><%=getWord("CLASS_VOCABULARY_13")%></a></p>
             <%=getWord("FAQ_CUATOMER_82")%><br><%=getWord("FAQ_CUATOMER_83")%> <br><%=getWord("FAQ_CUATOMER_84")%> <br><%=getWord("FAQ_CUATOMER_85")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a7-->
			 <!--a8-->
             <div class="say">
             <p><a name="member_kind_q8"><%=getWord("FAQ_CUATOMER_39")%> </a></p>
             <%=getWord("FAQ_CUATOMER_86")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a8-->
			 <!--a9-->
             <div class="say">
             <p><a name="member_kind_q9"><%=getWord("FAQ_CUATOMER_40")%> </a></p>
             <%=getWord("FAQ_CUATOMER_87")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a9-->
			 <!--a10-->
			 <!--todo 暫時拿除 
             <div class="say">
             <p><a name="member_kind_q10"><%=getWord("FAQ_CUATOMER_41")%> </a></p>
             <%=getWord("FAQ_CUATOMER_88")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a10-->
			 <!--a11-->
             <div class="say">
             <p><a name="member_kind_q11"><%=getWord("FAQ_CUATOMER_42")%> </a></p>
             <%=getWord("FAQ_CUATOMER_89")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a11-->
			 <!--a12-->
             <div class="say">
             <p><a name="member_kind_q12"><%=getWord("FAQ_CUATOMER_43")%> </a></p>
             <%=getWord("FAQ_CUATOMER_90")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a12-->
			 <!--a13-->
             <div class="say">
             <p><a name="member_kind_q13"><%=getWord("FAQ_CUATOMER_44")%></a></p>
              <%=getWord("FAQ_CUATOMER_91")%> 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a13-->
			 <!--a14-->
             <div class="say">
             <p><a name="member_kind_q14"><%=getWord("FAQ_CUATOMER_45")%></a></p>
             <%=getWord("FAQ_CUATOMER_92")%> <br><%=getWord("FAQ_CUATOMER_93")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a14-->
			 <!--a15-->
             <div class="say">
             <p><a name="member_kind_q15"><%=getWord("FAQ_CUATOMER_46")%></a></p>
             <%=getWord("FAQ_CUATOMER_94")%><br><%=getWord("FAQ_CUATOMER_95")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a15-->
			 <!--a16-->
             <div class="say">
             <p><a name="member_kind_q16"><%=getWord("FAQ_CUATOMER_47")%></a></p>
             <%=getWord("FAQ_CUATOMER_96")%><br><%=getWord("FAQ_CUATOMER_95")%><br><%=getWord("FAQ_CUATOMER_98")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a16-->
			 <!--a17-->
             <div class="say">
             <p><a name="member_kind_q17"><%=getWord("FAQ_CUATOMER_48")%></a></p>
             <%=getWord("FAQ_CUATOMER_99")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a17-->
			 <!--a18-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q18"><%=getWord("FAQ_CUATOMER_49")%></a></p>
              <%=getWord("FAQ_CUATOMER_100")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a18-->
			 <!--a19-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q19"><%=getWord("FAQ_CUATOMER_50")%></a></p>
             <%=getWord("FAQ_CUATOMER_101")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a19-->
			 <!--a20-->
             <div class="say">
             <p><a name="member_kind_q20"><%=getWord("FAQ_CUATOMER_51")%></a></p>
             <%=getWord("FAQ_CUATOMER_102")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a20-->
			 <!--a21-->
             <div class="say">
             <p><a name="member_kind_q21"><%=getWord("FAQ_CUATOMER_52")%></a></p>
             <%=getWord("FAQ_CUATOMER_103")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a21-->
			 <!--a22-->
             <div class="say">
             <p><a name="member_kind_q22"><%=getWord("FAQ_5")%></a></p>
             <%=getWord("FAQ_CUATOMER_104")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a22-->
			 <!--a23-->
             <div class="say">
             <p><a name="member_kind_q23"><%=getWord("FAQ_CUATOMER_53")%></a></p>
             <%=getWord("FAQ_CUATOMER_105")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a23-->
			 <!--a24-->
             <div class="say">
             <p><a name="member_kind_q24"><%=getWord("FAQ_CUATOMER_54")%></a></p>
             <!--修改開始 start-->
              <%=getWord("FAQ_CUATOMER_54_1")%>
             <div style="margin-left:10px;"><span class="color_4"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_1")%></span><br />
             <span class="bold color_2" style="text-decoration:underline;"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_2")%></span><br />
             <span class="color_2"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_3")%></span><br />
             <img src="/images/images_cn/vip_faq_01.gif" width="473" height="340" /><br /><br />
             </div>
			 <%=getWord("FAQ_CUATOMER_54_2")%>
             <div style="margin-left:10px;"><span class="bold color_2" style="text-decoration:underline;"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_4")%></span><br />
             <span class="color_2"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_5")%><br />
             <%=getWord("FAQ_CUATOMER_54_SUB_FAQ_6")%></span><br />
             <img src="/images/images_cn/vip_faq_02.gif" width="473" height="283" /><br /><br />
             <span class="color_2"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_7")%></span><br />
             <img src="/images/images_cn/vip_faq_03.gif" width="410" height="479" /><br /><br />
             <span class="color_2"><%=getWord("FAQ_CUATOMER_54_SUB_FAQ_8")%><br />
             <%=getWord("FAQ_CUATOMER_54_SUB_FAQ_9")%></span><br />
             <img src="/images/images_cn/vip_faq_04.gif" width="413" height="440" /><br /><br />
            
             </div>
			 <%=getWord("FAQ_CUATOMER_54_SUB_FAQ_10")%><br>
			 <%=getWord("FAQ_CUATOMER_54_SUB_FAQ_11")%><br />
             <%=getWord("FAQ_CUATOMER_54_SUB_FAQ_12")%>
             <!--修改結束 end-->
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a24-->
			 <!--a25-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q25"><%=getWord("FAQ_CUATOMER_55")%></a></p>
             <%=getWord("FAQ_CUATOMER_110")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a25-->
			 <!--a26-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q26"><%=getWord("FAQ_CUATOMER_56")%></a></p>
             <%=getWord("FAQ_CUATOMER_111")%><br>
			 <%=getWord("FAQ_CUATOMER_112")%><br><%=getWord("FAQ_CUATOMER_113")%><br>
			 <%=getWord("FAQ_CUATOMER_114")%><br><%=getWord("FAQ_CUATOMER_115")%><br>
			 <%=getWord("FAQ_CUATOMER_116")%><br><%=getWord("FAQ_CUATOMER_117")%><br>
			 <%=getWord("FAQ_CUATOMER_118")%><br><%=getWord("FAQ_CUATOMER_119")%><br>
			 <%=getWord("FAQ_CUATOMER_120")%><br><%=getWord("FAQ_CUATOMER_121")%>
			 
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a26-->
			 <!--a27-->
             <div class="say">
             <p><a name="member_kind_q27"><%=getWord("FAQ_CUATOMER_57")%></a></p>
             <%=getWord("FAQ_CUATOMER_122")%><br><%=getWord("FAQ_CUATOMER_123")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a27-->
			 <!--a28-->
             <div class="say">
             <p><a name="member_kind_q28"><%=getWord("FAQ_CUATOMER_58")%></a></p>
             <%=getWord("FAQ_CUATOMER_124")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a28-->
			 <!--a29-->
             <div class="say">
             <p><a name="member_kind_q29"><%=getWord("FAQ_CUATOMER_59")%></a></p>
             <%=getWord("FAQ_CUATOMER_125")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
             <!--a29-->
			 <!--a30-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q30"><%=getWord("FAQ_CUATOMER_60")%></a></p>
             <%=getWord("FAQ_CUATOMER_126")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>             
            </div>
			-->
             <!--a30-->
			 <!--a31-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q31"><%=getWord("FAQ_CUATOMER_61")%></a></p>
             <%=getWord("FAQ_CUATOMER_127")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>		
            </div>
			-->
             <!--a31-->
			 <!--a32-->
			 <!--todo 暫時拿掉
             <div class="say">
             <p><a name="member_kind_q32"><%=getWord("FAQ_CUATOMER_62")%></a></p>
             <%=getWord("FAQ_CUATOMER_128")%>
			<%if (str_link_Hash_Point<>"") then %>
            <div class="t-right"><%=getGoBackButtonBySidebar()%></div>
            <%end if%>
             </div>
			 -->
             <!--a32-->
			 
             </div>
             <div class="ps"> <%if (isPayMember()) then %><%=getWord("faq_12")%><a href="/program/contact/contact_us.asp"><%=getWord("faq_13")%></a>
			<!--<%=getWord("faq_14")%><a href="#"><%=getWord("faq_15")%></a>--><%end if%>
             </div>        
 
             </div>
             
             </div>
             <%
			 end if
			 %>
			 <!--3類end-->             
        </div>
        <!--內容end-->
        
  <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
