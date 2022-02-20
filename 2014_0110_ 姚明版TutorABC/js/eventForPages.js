$(function() {
	$('.topNavFixed').css({'position': 'fixed', 'top':0, 'left':0 });
	$(".language").click(function() {
	$(".dropdown-menu").toggle();
	});
	
	$(".fancybox").fancybox();

	$(".various").fancybox({
		padding: 3,
		maxWidth	: 800,
		maxHeight	: 600,
		fitToView	: false,
		width		: '70%',
		height		: '70%',
		autoSize	: false,
		closeClick	: false,
		openEffect	: 'none',
		closeEffect	: 'none',
		beforeShow: function(){$(".fancybox-skin").css("backgroundColor","#000");}
	});
	////固定nav bar
	var sticky_navigation_offset_top = 380;
	var sticky_navigation = function(){
	var scroll_top = $(window).scrollTop();
	/*
		if (scroll_top > sticky_navigation_offset_top) { 
			$('.regFixedWrap').css({'position': 'fixed', 'top':120 });
			$('.regCustomerService').css({'position': 'fixed', 'top':150 });
			
			
		} else {
			$('.regFixedWrap').css({'position': 'relative' ,'top':700}); 
			$('.regCustomerService').css({'position': 'absolute', 'top':760 });
		}
		*/
		if (scroll_top > sticky_navigation_offset_top) { 
			$('.regFixedWrap').css({'position': 'fixed', 'top':120 });
			$('.regForm').css({'right': '0px', 'margin-right':'0px' });
			$('.regCustomerServiceIcon').css({'position': 'fixed', 'top':120 });
			
		} else {
			$('.regFixedWrap').css({'position': 'relative' ,'top':400}); 
			$('.regForm').css({'right': '0px', 'margin-right':'0px' });
			$('.regCustomerServiceIcon').css({'position': 'absolute', 'top':0 });
			
		}
		var scroll_bottom = $(document).height() -700 - $('.footer').height() ;
		if (scroll_top > scroll_bottom) { 
			$('.regFixedWrap').fadeOut(300);
		} else {
			$('.regFixedWrap').fadeIn(300);
		}   
	};
	sticky_navigation();
	$(window).scroll(function(){
		////固定nav bar
		sticky_navigation();
		////固定nav bar
		// 捲動的數字
		var scroH = Number($(this).scrollTop());
	});
});	