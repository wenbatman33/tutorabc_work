for(var i = 0; i < 4; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'center';gv_vAlignTable['u1'] = 'center';document.getElementById('u2_img').tabIndex = 0;

u2.style.cursor = 'pointer';
$axure.eventManager.click('u2', u2Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u2LinksClick'></div>")
var u2LinksClick = document.getElementById('u2LinksClick');
function u2Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u2LinksClick');
}

InsertBeforeEnd(u2LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u2Clicku8abd084350a34200b942b5b85654037d(event)'>已有預約訂課</div>");
function u2Clicku8abd084350a34200b942b5b85654037d(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('已有預約訂課.html');

	ToggleLinks(e, 'u2LinksClick');
}

InsertBeforeEnd(u2LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u2Clickub8036769c28942f5a71c5b218e819c04(event)'>尚未預約訂課</div>");
function u2Clickub8036769c28942f5a71c5b218e819c04(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('尚未預約訂課.html');

	ToggleLinks(e, 'u2LinksClick');
}

InsertBeforeEnd(u2LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u2Clicku7d7948187a12400d99251f3926bf7c05(event)'>進入時的時段未預約</div>");
function u2Clicku7d7948187a12400d99251f3926bf7c05(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('進入時的時段未預約.html');

	ToggleLinks(e, 'u2LinksClick');
}
