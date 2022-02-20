for(var i = 0; i < 4; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
u3.tabIndex = 0;

u3.style.cursor = 'pointer';
$axure.eventManager.click('u3', u3Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u3LinksClick'></div>")
var u3LinksClick = document.getElementById('u3LinksClick');
function u3Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u3LinksClick');
}

InsertBeforeEnd(u3LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u3Clicku6880e725abd6411282bd7b0d41e06218(event)'>未登入</div>");
function u3Clicku6880e725abd6411282bd7b0d41e06218(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_未登_.html');

	ToggleLinks(e, 'u3LinksClick');
}

InsertBeforeEnd(u3LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u3Clicku7db16a0cc3064102ad341a1c19fe4522(event)'>已登入</div>");
function u3Clicku7db16a0cc3064102ad341a1c19fe4522(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_已登_.html');

	ToggleLinks(e, 'u3LinksClick');
}
gv_vAlignTable['u1'] = 'center';u2.tabIndex = 0;

u2.style.cursor = 'pointer';
$axure.eventManager.click('u2', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('本周精選.html');

}
});
