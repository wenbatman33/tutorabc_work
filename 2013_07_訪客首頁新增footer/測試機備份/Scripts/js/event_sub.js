$(function() {
	
		$('.fixed_box').css({'position': 'fixed', 'top':160 });		
		$('#top_news').css({'position': 'fixed' ,'z-index':3 });
		$('#top_nav').css({'position': 'fixed' ,'top':48});
		
		
		
		$('.news_list_wrap').hover(function(){

		$(this).find('.news_list_on').stop().animate({
					top: 0
				}, 200);
			}, function(){
				$(this).find('.news_list_on').animate({
					top: 112
				}, 200);
	});

	var scroll_bottom = $(document).height() -600- $('.registration_form').height() ;


	var sticky_navigation = function(){
	var scroll_top = $(window).scrollTop();
	
	if (scroll_top > scroll_bottom) { 
		$('.fixed_box').fadeOut(300)
	} else {
		$('.fixed_box').fadeIn(300);
	}   
	};
	
	sticky_navigation();

	$(window).scroll(function(){
		////固定nav bar
		sticky_navigation();
	});	
});


function turnMore(turnType) {
		if (turnType == "more") {
			//切換為纖細信息
						$("#j-album-less").removeClass("closeview").addClass("openview");
			$(".shouzhanOpen").removeClass("closeview").addClass("openview");
			$("#j-album-more").removeClass("openview").addClass("closeview");
		} else if (turnType == "less") {
			//切換為縮略信息
			
			
				$("#j-album-less").removeClass("openview").addClass("closeview");
			$(".shouzhanOpen").removeClass("openview").addClass("closeview");
			$("#j-album-more").removeClass("closeview").addClass("openview");

		}
	}