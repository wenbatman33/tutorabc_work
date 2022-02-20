$(function () {

var $block = $('#album'),
			$ad = $block.find('.ad'),
			currentIndex = 0,
			oldIndex=0,
			clickIndex=0,
			fadeOutSpeed = 500,
			fadeInSpeed = 500,
			defaultZ = 100,
			isHover = false,
			totalImg=$ad.length,
			speed = 8000;
			var time;
			timer = setTimeout(rollNumber, speed);
			$ad.css({opacity: 0,zIndex: defaultZ - 1	}).eq(currentIndex).css({	opacity: 1,	zIndex: defaultZ});
			/*-------------------------   點點   -------------------------*/
			var str = '';
			for(var i=0;i<$ad.length;i++){	str += '<a href="#" data-index="'+ i +'"></a>';}
			var $controlA = $('#album').append($('<div class="control">' + str + '</div>').css('zIndex', defaultZ + 1)).find('.control a');
			$('.control a').eq(0).addClass('on');
			/*-------------------------   點點   -------------------------*/
			
	
		function rollNumber(){
			rollImages()
		}	
			
		function rollImages(){
			
		
				$ad.eq(oldIndex).stop().fadeTo(fadeOutSpeed, 0).css({zIndex: -1});
				$('.control a').eq(oldIndex).removeClass('on');

				oldIndex=currentIndex;

				$ad.eq(currentIndex).stop().fadeTo(fadeInSpeed, 1).css({zIndex: defaultZ});
				$('.control a').eq(currentIndex).addClass('on');
				timer = setTimeout(rollNumber, speed);
				
				//alert(currentIndex);
				currentIndex+=1;
				if(currentIndex>totalImg-1){
					currentIndex=0
				}
		
		}

		
		$controlA.click(function(){
			clearTimeout(timer);
			currentIndex = Number($(this).attr("data-index"));
			rollImages();
		});
		
});