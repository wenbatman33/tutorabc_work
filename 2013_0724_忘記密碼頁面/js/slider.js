
		var theInt = null;
		var $crosslink, $navthumb;
		var curclicked = 0;
		
		theInterval = function(cur){
			clearInterval(theInt);
			
			if( typeof cur != 'undefined' )
				curclicked = cur;
			
			$crosslink.removeClass("active-thumb");
			$navthumb.eq(curclicked).parent().addClass("active-thumb");
			
				$(".stripNav ul li a").eq(curclicked).trigger('click');
						
			
			
			theInt = setInterval(function(){
				$crosslink.removeClass("active-thumb");
				
				
				$navthumb.eq(curclicked).parent().addClass("active-thumb");
				
				
				
				$(".stripNav ul li a").eq(curclicked).trigger('click');
				curclicked++;
				
				if( 6 == curclicked )
					curclicked = 0;
				
			}, 3000);
		};
		
		$(function(){
			
			$("#main-photo-slider").codaSlider();
			
			$navthumb = $(".nav-thumb");
			$crosslink = $(".cross-link");
			
			
			$navthumb.click(function() {
				
				var $this = $(this);
				
				$('.nav-thumb').css('border-color','#f99');
				
				$(this).css('border-color','#fff');

				
				theInterval($this.parent().attr('href').slice(1) - 1);
				return false;
			});
			
			theInterval();
		});

