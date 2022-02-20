$(function(){
		var $marqueeUl = $('div#user_marquee ul'),
			$marqueeli = $marqueeUl.children(),
			_height = 40 ,
			scrollSpeed = 600,
			timer,
			speed = 1000 + scrollSpeed,
			total= $marqueeli.length,
			_now =0;

			
			
			$marqueeUl.css('top', 0);
			$marqueeli.hover(function(){
				clearTimeout(timer);
			}, function(){
				timer = setTimeout(showad, speed);
			});

		function showad(){
			_now+=1;
			var moveNum= 0-( _now * _height);
			$marqueeUl.animate({
				top:moveNum
			}, scrollSpeed, function(){

				if(_now >= total+1){
					_now=0;
					$marqueeUl.css('top', 40);
					$marqueeUl.animate({top:0}, scrollSpeed);
			
				}
			});

			timer = setTimeout(showad, speed);
		}

		timer = setTimeout(showad, speed);
		$('a').focus(function(){
			this.blur();
		});
	});