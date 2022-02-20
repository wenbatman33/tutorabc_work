$( document ).ready( function(){
		$(".popup").fancybox({
			fitToView	: false,
			type       :'iframe',
			width		: 800,
			height		: 600,
			autoSize	: false,
			closeClick	: false,
			openEffect	: 'none',
			closeEffect	: 'none',
			beforeShow: function () {
			this.width = $(this.element).data("width") ? $(this.element).data("width") : null;
			this.height = $(this.element).data("height") ? $(this.element).data("height") : null;
			}
		});
});