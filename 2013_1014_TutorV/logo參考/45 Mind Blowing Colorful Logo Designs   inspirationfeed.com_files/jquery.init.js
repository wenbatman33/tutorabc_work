jQuery(function($){
	$(document).ready(function(){
		
			
		// Menu	
		 $("ul.sf-menu").superfish({ 
			delay: 600 // 0.6 second delay on mouseout 
		});
		
		//Toggle Sidebar Menu
		$("#sidebar .menu li").menu({
  			itemLink: 0
  		});
		
		
		//homepage slides
		$('#home-slides').slides({
			effect: 'fade',
			crossfade: true,
			fadeSpeed: 600,
			play: 5000,
			pause: 10,
			autoHeight: false,
			preload: false,
			hoverPause: true,
			generateNextPrev: false,
			generatePagination: true
		});
		
		
		//portfolio slides
		$("#portfolio-slides").slides({
			effect: 'fade',
			crossfade: true,
			fadeSpeed: 400,
			pause: 10,
			pagination: true,
			generatePagination: true,
			generateNextPrev: true,
			play: 7000,
			autoHeight: true
		});
		
		//add left margin to slider pagination to center it
		if($("ul.pagination").length) {
			var $paginationWidth = $("ul.pagination").outerWidth() ;
			var $paginationHalfWidth = $paginationWidth / 2;
			$(".pagination").css('marginLeft', -$paginationHalfWidth);
		}
		
		
		// Back to top	
		$(window).scroll(function () {
			if ($(this).scrollTop() > 400) {
				$('a[href=#top]').fadeIn();
			} else {
				$('a[href=#top]').fadeOut();
			}
		});
		$('a[href=#top]').click(function(){
			$('html, body').animate({scrollTop:0}, 'slow');
			return false;
		});	
		
		// Add button class to portfolio category links
		$("#portfolio-cats li.current-cat a").addClass('active');
		$("#portfolio-cats a").first().addClass("primary");
		$("#portfolio-cats a").addClass('gh-button');
		
		
		
		// Add Class To Main Element On Sidebar Pages
		if($("#sidebar").length) {
			$("#main").addClass("main-bg");
		}
		
		
		
		// PrettyPhoto Without gallery
		$(".prettyphoto-link").prettyPhoto({
			theme: 'pp_default', // light_rounded / dark_rounded / light_square / dark_square / facebook */
			animation_speed:'normal',
			allow_resize: true,
			keyboard_shortcuts: true,
			show_title: false,
			social_tools: false,
			autoplay_slideshow: false
		});
			
			
		//PrettyPhoto With Gallery
		$("a[rel^='prettyPhoto']").prettyPhoto({
			theme: 'pp_default', // light_rounded / dark_rounded / light_square / dark_square / facebook */
			animation_speed:'normal',
			allow_resize: true,
			keyboard_shortcuts: true,
			show_title: false,
			slideshow:3000,
			social_tools: false,
			autoplay_slideshow: false,
			overlay_gallery: false
		});
			
			
			
		// Toggle	
		$(".toggle_container").hide(); 
		$("h3.trigger").click(function(){
			$(this).toggleClass("active").next().slideToggle("normal");
			return false; //Prevent the browser jump to the link anchor
		});
		
		
		
		// Tabvs & Accordion			
		$( ".tabs" ).tabs({ fx: { height: 'toggle', opacity: 'toggle' } });
		$( ".accordion" ).accordion({
			autoHeight: false
		});
		
		
		
		// Videos		
		$(".fitvids").fitVids();



		// portfolio		
		$('.portfolio-item').hover(function(){
			$(this).find("img").animate({opacity:'0.7'}, 300);
			$(this).find("h2").stop(true, true).slideDown("normal"); }
			,function(){
			$(this).find("img").animate({opacity:'1'}, 300);
			$(this).find("h2").stop(true, true).slideUp("normal", "jswing");
		});
		
		
		// animate thumbnails		
		$('.single-gallery-thumb, .loop-entry-thumbnail, .post-thumbnail, .portfolio-single-img, .gallery-item, .portfolio-gallery-item').hover(function(){
			$(this).find("img").animate({opacity:'0.80'}, 300);}
			,function(){
			$(this).find("img").animate({opacity:'1'}, 300);
		});
		
		
		//equal heights
		function equalHeight(group) {
			var tallest = 0;
			group.each(function() {
				var thisHeight = $(this).height();
				if(thisHeight > tallest) {
					tallest = thisHeight;
				}
			});
			group.height(tallest);
		}
		equalHeight($(".service-item"))
		
		
	
	}); // END doc.ready	
}); // END function
