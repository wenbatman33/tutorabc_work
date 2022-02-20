<!--#include virtual="/lib/include/html/template/template_change.asp"-->


<!-- Main Content Start -->
<div class="main">
    <div class="main_con">
       <!--內容start-->
        <div class="main_membox">
        <!--上版位課程start-->
		<!--#include virtual="/lib/include/html/include_learning.asp"--> 
        <div class="top_block"></div>
        <!--上版位課程end-->
        <!--上課紀錄評分-->
        <ul class="word_list">
		<%
		Dim obj_mtl_voc_opt
			Dim arr_voc_front_name_result
			Dim arr_voc_result
			Dim str_voc_front_name
			Dim int_col
				str_voc_front_name = request("voc_front_name")
			if(str_voc_front_name  = "") then
				str_voc_front_name = "*"
			end if
			Set obj_mtl_voc_opt = Server.CreateObject("TutorLib.MaterialOpt.MaterialVocabularyOpt")
		%>
       <%
			arr_voc_front_name_result = obj_mtl_voc_opt.getVocabularyBankFrontChar(g_var_client_sn)
			if (obj_mtl_voc_opt.PROCESS_STATE_ID = 1) then 
			Response.write("<li class=""Eword""><a href=""learning_word.asp?type=*&language=zh-cn"" class=""ch"">全部</a></li>")		
				for int_col = 0 to Ubound(arr_voc_front_name_result)
				Response.write("<li><a href=""learning_word.asp?voc_front_name="&arr_voc_front_name_result(int_col)&"&language=zh-cn"" class=""Eword"">"&arr_voc_front_name_result(int_col)&"</a></li>")				
				next
			else
			end if
		 %>
        </ul>
        <ul class="class_list4">
        <li>
        <h3>字汇</h3>
        <h5>发音</h5>
        <h5>释义</h5>
        </li>
       
        <li>
        <h3>compatriot<br />
        <a href="#">删除</a></h3>
        <h5><input name="" type="button" class="btn_play" /></h5>
        <h5><span class="bold">字汇解释 : </span><br />
        v. 废止, 革除, 取消<br /><br /><span class="bold">英文释义 </span>: <br />
        废除 / do away with; &quot;Slavery was bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbabolished in the mid-19th century in America and in Russia&quot;<br />
        <span class="bold">例句 : bad customs and laws ought to be ajjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjbolished.</span></h5>
        </li>
		<%
		 Dim str_voc
		 arr_voc_result = obj_mtl_voc_opt.getVocabularyBank(g_var_client_sn, str_voc_front_name,0,0)
		 for int_col = 0 to Ubound(arr_voc_result)
		 str_voc = arr_voc_result(int_col)
		 if (int_col mod 2)= 0 then
		 %>
        <li class="C4_BgGreen">
		<%else%>
		<li>
		<%end if%>
        <h3><%=lcase(str_voc)%><br/>
        <%
		
		Dim str_voc_opt : str_voc_opt = "add"
		Dim str_voc_des : str_voc_des = getWord("ROLL_BACK")
		if (obj_mtl_voc_opt.checkIsInVocabularyBank(g_var_client_sn, str_voc)) then
			str_voc_opt = "del"
			str_voc_des = getWord("delete")
		%>
		<%end if%>
        <a id ="voc_opt_<%=str_voc%>" href="javascript:void(0);" onclick="excuteVocOpt('<%=str_voc%>');" title="<%=str_voc_opt%>"><%=str_voc_des%></a>
        </h3>
		<h5><a  id ="voc_sound_<%=str_voc%>" href="javascript:void(0);" onclick="OpenVocWin2('<%=str_voc%>');" alt="<%=str_voc%>" ><img src="/images/images_cn/arrow_icon.jpg" /></a></h5>
        <h5><%=obj_mtl_voc_opt.getVocabularyHTMLDescription(str_voc, "zh-cn")%></h5>
        </li>
    
		<%next%>
        </ul>
        <div class="page"><div class="floatL"><span class="color_1">在此期间您共存了 12 个单词，共计 9 页，您目前在第 1 页</span></div><a href="#">上一頁</a><a href="#">1</a><a href="#">2</a><a href="#">3</a><a href="#">4</a><a href="#">5</a><a href="#">6</a><a href="#">7</a><a href="#">8</a><a href="#">下一頁</a></div>
        <!--上課紀錄評分end-->
        </div>
        <!--內容end-->
        </div>

<div class="clear"></div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!-- Main Content End -->
