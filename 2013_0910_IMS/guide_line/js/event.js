$(function() {
	////固定nav bar
	var nav_offset_top = 160;
	var sticky_navigation = function(){
		var scroll_top = $(window).scrollTop(); 
		if (scroll_top > nav_offset_top) { 
		
			$('.wrapNav').css({'position': 'fixed', 'top':32,'margin':'0 auto'  });
		} else {
			$('.wrapNav').css({'position': 'relative' ,'top':0}); 
		}
	};
	sticky_navigation();
	$(window).scroll(function(){
			////固定wrapNav
			sticky_navigation();
			////固定wrapNav
	});
});
