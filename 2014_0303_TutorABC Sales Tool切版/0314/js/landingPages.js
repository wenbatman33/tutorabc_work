var IMG_WIDTH = 770;
var currentImg=0;
var maxImages=5;
var speed=500;
var imgs;
var imgsWrap;
var swipeOptions={
	triggerOnTouchEnd : true,	
	swipeStatus : swipeStatus,
	allowPageScroll:"vertical",
	threshold:200		
}

$(function(){				
	imgs = $("#content");
	imgsWrap = $("#imgs");
	imgs.swipe( swipeOptions );
});
function swipeStatus(event, phase, direction, distance){
	

	
	if( phase=="end" && direction==null){
		var par = $(event.target).parent();
        var  urlLink = par.data('link')
		//console.log(urlLink);
		window.open(urlLink, '_blank');
	}	else if( phase=="move" && (direction=="left" || direction=="right") ){
		var duration=0;
		if (direction == "left")
			scrollImages((IMG_WIDTH * currentImg) + distance, duration);
		else if (direction == "right")
			scrollImages((IMG_WIDTH * currentImg) - distance, duration);
	}
	else if ( phase == "cancel"){
		scrollImages(IMG_WIDTH * currentImg, speed);
	}else if ( phase =="end" ){
		if (direction == "right")
			previousImage()
		else if (direction == "left")			
			nextImage()
	}
}
function previousImage(){
	currentImg = Math.max(currentImg-1, 0);
	scrollImages( IMG_WIDTH * currentImg, speed);
}

function nextImage(){
	currentImg = Math.min(currentImg+1, maxImages-1);
	scrollImages( IMG_WIDTH * currentImg, speed);
}
function scrollImages(distance, duration)
{
	imgsWrap.css("-webkit-transition-duration", (duration/1000).toFixed(1) + "s");
	var value = (distance<0 ? "" : "-") + Math.abs(distance).toString();
	imgsWrap.css("-webkit-transform", "translate3d("+value +"px,0px,0px)");
}