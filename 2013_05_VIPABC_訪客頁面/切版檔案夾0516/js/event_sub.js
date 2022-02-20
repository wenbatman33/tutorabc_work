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

		
		
});