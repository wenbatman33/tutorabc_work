for(var i = 0; i < 5; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'top';gv_vAlignTable['u4'] = 'top';gv_vAlignTable['u1'] = 'center';u2.tabIndex = 0;

u2.style.cursor = 'pointer';
$axure.eventManager.click('u2', u2Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u2LinksClick'></div>")
var u2LinksClick = document.getElementById('u2LinksClick');
function u2Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u2LinksClick');
}

InsertBeforeEnd(u2LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u2Clickufe4a06a643b145d4be39bc775e5a6d4a(event)'>未登入</div>");
function u2Clickufe4a06a643b145d4be39bc775e5a6d4a(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_未登_.html');

	ToggleLinks(e, 'u2LinksClick');
}

InsertBeforeEnd(u2LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u2Clicku95824d20693a4ca2bfb5079659230799(event)'>已登入</div>");
function u2Clicku95824d20693a4ca2bfb5079659230799(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_已登_.html');

	ToggleLinks(e, 'u2LinksClick');
}
