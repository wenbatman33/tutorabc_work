function gotoHashPoint(p_divName,p_question_id){
		$("#"+p_divName+"_answer_div").css({ display: "" }); 
		//location.hash = "#"+p_divName+"_"+p_question_id;
		//var  x=$(obj).parents().parents().offset();
		var obj= p_divName+"_"+p_question_id;
		var num=$('p [name='+obj+']').offset().top;
		//var top=num.top()
		var top=num-150;
		console.log(top);
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
		$body.stop().animate({	scrollTop: top}, 1000 , "easeOutExpo");
		$("#"+p_divName+"_atag").text("关闭");
		$("#"+p_divName+"_question_div").removeClass();
		$("#"+p_divName+"_question_div").addClass("f-right faq-btn-close");
}


function openAndCloseDiv(p_divName){
	if ($("#"+p_divName+"_answer_div").css("display") != "none")
	{	
		$("#"+p_divName+"_answer_div").slideUp("slow");
		$("#"+p_divName+"_atag").text("打开");
		$("#"+p_divName+"_question_div").removeClass();
		$("#"+p_divName+"_question_div").addClass("f-right faq-btn-open");
	}
	else
	{
	
	//20101013 會議增加 打開一個要關閉另一個
	$(".answer").hide();	
	
		$("#"+p_divName+"_answer_div").slideDown("slow");
		$("#"+p_divName+"_atag").text("关闭");
		$("#"+p_divName+"_question_div").removeClass();
		$("#"+p_divName+"_question_div").addClass("f-right faq-btn-close");
	}

}