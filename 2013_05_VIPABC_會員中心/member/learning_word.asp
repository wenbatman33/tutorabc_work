<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script language="JavaScript">
var g_var_client_sn = "&data=<%=g_var_client_sn%>";
var roll_back = "<%=getWord("roll_back")%>";
var getWorddelete = "<%=getWord("delete")%>";
</script>
<script type='text/javascript' src="/lib/javascript/framepage/js/learning_word.js"></script>
<!--內容start-->
<div id="main_con">
    <div class="main_membox">
        <div class='page_title_2'><h2 class='page_title_h2'>我的单词银行</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 
        <!--上版位課程start-->
        <!--#include virtual="/lib/include/html/include_learning.asp"-->
        <!--上版位課程end-->
        <!--上課紀錄評分-->
        <%
			Dim obj_mtl_voc_opt
			Dim arr_voc_front_name_result
			Dim arr_voc_result
			Dim str_voc_front_name
			Dim int_col
				str_voc_front_name = request("voc_front_name")
			Set obj_mtl_voc_opt = Server.CreateObject("TutorLib.MaterialOpt.MaterialVocabularyOpt")
        %>
        <ul class="word_list">
            <%
			arr_voc_front_name_result = obj_mtl_voc_opt.getVocabularyBankFrontChar(g_var_client_sn)
			if (obj_mtl_voc_opt.PROCESS_STATE_ID = 1) then 
			'Response.write("<li class=""Eword""><a href=""learning_word.asp?type=*&language=zh-cn"" class=""ch"">全部</a></li>")		
				for int_col = 0 to Ubound(arr_voc_front_name_result)
                if(str_voc_front_name  = "") then
				    str_voc_front_name = arr_voc_front_name_result(0)
			    end if
				Response.write("<li><a href=""learning_word.asp?voc_front_name="&arr_voc_front_name_result(int_col)&"&language=zh-cn"" class=""Eword"">"&arr_voc_front_name_result(int_col)&"</a></li>")				
				next
			else
			end if
            %>
        </ul>
        <ul class="class_list4">
            <li>
                <h3>
                    <%=getWord("CLASS_VOCABULARY_7")%></h3>
                <h5>
                    <%=getWord("CLASS_VOCABULARY_8")%></h5>
                <h5>
                    <%=getWord("CLASS_VOCABULARY_9")%></h5>
            </li>
            <%
		 Dim str_voc
         dim vocabulary
		 arr_voc_result = obj_mtl_voc_opt.getVocabularyBank(g_var_client_sn, str_voc_front_name,20,0)
		 for int_col = 0 to Ubound(arr_voc_result)
		 str_voc = arr_voc_result(int_col)
         vocabulary = replace(arr_voc_result(int_col)," ","_")
		 if (int_col mod 2)= 0 then
            %>
            <li class="C4_BgGreen">
                <%else%>
                <li>
                    <%end if%>
                    <h3>
                        <%=lcase(str_voc)%><br />
                        <%
		
		Dim str_voc_opt : str_voc_opt = "add"
		Dim str_voc_des : str_voc_des = getWord("ROLL_BACK")
		if (obj_mtl_voc_opt.checkIsInVocabularyBank(g_var_client_sn, str_voc)) then
			str_voc_opt = "del"
			str_voc_des = getWord("delete")
                        %>
                        <%end if%>
                        <br />
                        <a href="javascript:void(0)" onclick="window.open('http://dict.baidu.com/s?wd=<%=lcase(str_voc)%>&tn=dict')">
                            <img src="../../images/vip_voc_01.gif" border="0" alt="" /></a>
                        <br />
                        <br />
                        <a id="voc_opt_<%=Replace(str_voc, " ", "_")%>" href="javascript:void(0);" onclick="excuteVocOpt('<%=str_voc%>');"
                            title="<%=str_voc_opt%>">
                            <%=str_voc_des%></a>
                    </h3>
                    <!--<h5><a id ="voc_sound_<%=str_voc%>" href="javascript:void(0);" onclick="OpenVocWin2('<%=str_voc%>');" alt="<%=str_voc%>" ><img src="/images/images_cn/arrow_icon.jpg" /></a></h5>-->
                    <h5>
                        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" id="PlayMp3" width="31"
                            height="25" codebase="http://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab">
                            <param name="movie" value="PlayMp3.swf" />
                            <param name="quality" value="high" />
                            <param name="bgcolor" value="#869ca7" />
                            <param name="allowScriptAccess" value="sameDomain" />
                            <param name="FlashVars" value="vocName=<%=str_voc%>" />
                            <embed src="PlayMp3.swf" flashvars="vocName=<%=str_voc%>" quality="high" bgcolor="#869ca7"
                                width="31" height="25" name="PlayMp3" align="middle" play="true" loop="false"
                                quality="high" allowscriptaccess="sameDomain" type="application/x-shockwave-flash"
                                pluginspage="http://www.adobe.com/go/getflashplayer">
			</embed>
                        </object>
                    </h5>
                    <h5>
                        <%
        dim htmlDescription
        htmlDescription = obj_mtl_voc_opt.getVocabularyHTMLDescription(vocabulary, "zh-cn")
        if right(htmlDescription,37) = "<span class=""bold"">英文释义 :</span><br/>" then
            'Anfernee 2013/3/30 這邊沒有教材編號，所以只能用單字當成where條件去查第一筆 (c_dictionary_state、e_dictionary_state都設為1的優先查出來)
            dim sqlParameters : sqlParameters = Array(str_voc)
            dim sqlStatement : sqlStatement = "SELECT [definition] FROM dbo.material_vocabulary WHERE vocabulary=@v AND valid=1 AND [definition] IS NOT NULL AND is_checked = 1 ORDER BY c_dictionary_state DESC,e_dictionary_state DESC"
            dim queryResult : queryResult = excuteSqlStatementScalar(sqlStatement, sqlParameters, CONST_VIPABC_RW_CONN)
            if (isSelectQuerySuccess(queryResult)) then
	            if (Ubound(queryResult) >= 0) then
		            htmlDescription = htmlDescription & queryResult(0)
	            end if
            end if
        end if
        response.Write htmlDescription%></h5>
                </li>
                <%next%>
        </ul>
        <div class="page">
            <div class="floatL">
                <span class="color_1">
                    <%=getWord("until_save")%><%=Ubound(arr_voc_result)+1%><%=getWord("CLASS_VOCABULARY_11")%></span></div>
        </div>
        <!--上課紀錄評分end-->
        <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
</div>
<!--內容end-->