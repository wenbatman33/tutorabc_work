$(function() {
	$('.topNavFixed').css({'position': 'fixed', 'top':0, 'left':0 });
	$(".language").click(function() {
	$(".dropdown-menu").toggle();
	});
	
	$(".fancybox").fancybox();

	$(".various").fancybox({
		padding: 10,
		maxWidth	 : 800,
		maxHeight	 : 600,
		fitToView	 : false,
		width		 : '70%',
		height		 : '70%',
		autoSize	 : false,
		closeClick	 : false,
		openEffect	 : 'none',
		closeEffect : 'none',
		beforeShow  : function(){$(".fancybox-skin").css("backgroundColor","#000");},
		helpers     : {overlay: {locked: false}}
	});
	
	
	$(".fancyboxIframe").fancybox({
			padding     : 3,
			width		 : 692,
			height		 : 431,
			fitToView	 : false,
			autoSize	 : false,
			closeClick	 : false,
			openEffect	 : 'none',
			closeEffect : 'none',
			type        :'iframe',
			closeClick  : false,
			beforeLoad  : function(){var url= $(this.element).data("href");this.href = url;	},  
			afterLoad   : function(){jQuery('.fancybox-skin').css('background-color','#000');},
			helpers     : {overlay: {locked: false}}

	});
	
	

	//固定 landing from
	var sticky_navigation_offset_top = 580;
	var sticky_navigation = function(){
	var scroll_top = $(window).scrollTop();
	var scroll_bottom = $(document).height() -$('.landingForm').height()-120 - $('.footer').height() ;
		console.log($('.landingForm').height())
		if (scroll_top > scroll_bottom) { 
			$('.landing').css({'position': 'relative' ,'top':(scroll_bottom+110)}); 
		}else if (scroll_top > sticky_navigation_offset_top) { 
			$('.landing').css({'position': 'fixed', 'top':120 });
		}else {
			$('.landing').css({'position': 'relative' ,'top':700}); 
		}
		
		
	};
	sticky_navigation();
	$(window).scroll(function(){
		sticky_navigation();
	});
	
	


});	