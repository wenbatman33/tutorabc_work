$(function() {
	////固定nav bar
	var sticky_navigation_offset_top = $('#top_nav').offset().top;
	var sticky_navigation = function(){
		var scroll_top = $(window).scrollTop(); 
		if (scroll_top > sticky_navigation_offset_top) { 
			$('#top_nav').css({'position': 'fixed', 'top':0, 'left':0 });
		} else {
			$('#top_nav').css({'position': 'relative' }); 
		}   
	};
	

	sticky_navigation();
	


	////SmoothScroll 項目 還需要再調整
	
	var part_01_top = $("#part_01").offset().top;
	var part_02_top = $("#part_02").offset().top;
	var part_03_top = $("#part_03").offset().top;
	var part_04_top = $("#part_04").offset().top;
	var part_05_top = $("#part_05").offset().top;
	var part_06_top = $("#part_06").offset().top;

	$(window).scroll(function(){
		
		 ////固定nav bar
		 sticky_navigation();
		 ////固定nav bar
		
		
		
		var scroH = $(this).scrollTop()+150;
		console.log($(this).scrollTop());

		 if(scroH >= part_06_top){
			set_cur(".part_06");
		}else if(scroH>= part_05_top){
			set_cur(".part_05");
		}else if(scroH>= part_04_top){
			set_cur(".part_04");
		}else if(scroH>= part_03_top){
			set_cur(".part_03");
		}else if(scroH>= part_02_top){
			set_cur(".part_02");
		}else if(scroH>= part_01_top){
			set_cur(".part_01");
		}
	});
	$(".nav_bar li a").click(function() {
		var el = "#"+$(this).attr('class');
		
		var num =$(el).offset().top;
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.animate({
				scrollTop: num-82
			}, 1200);
 	});
});




function pageScroll(obj){
	if(!obj){
		$.scrollTo({ top:0, left:0 }, scrollSpeed, {easing:"easeInOutQuart"});
	} else {
		$.scrollTo(obj, scrollSpeed, {easing:"easeInOutQuart"});
	}
}

function set_cur(n){
	
	
	
	if($(".nav_bar a").hasClass("cur")){
		$(".nav_bar a").removeClass("cur");
	}
	$(".nav_bar a"+n).addClass("cur");
}

