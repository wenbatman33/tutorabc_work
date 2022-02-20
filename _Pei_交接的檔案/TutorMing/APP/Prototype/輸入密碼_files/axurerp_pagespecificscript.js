for(var i = 0; i < 26; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'center';u16.tabIndex = 0;

u16.style.cursor = 'pointer';
$axure.eventManager.click('u16', u16Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u16LinksClick'></div>")
var u16LinksClick = document.getElementById('u16LinksClick');
function u16Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u16LinksClick');
}

InsertBeforeEnd(u16LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u16Clicku6680b734db4045a2aa26012dbaa47f3e(event)'>登入成功</div>");
function u16Clicku6680b734db4045a2aa26012dbaa47f3e(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('六宮格.html');

	ToggleLinks(e, 'u16LinksClick');
}

InsertBeforeEnd(u16LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u16Clickudf2167a7c6a44f9fa0a13fb29fa4bdb4(event)'>登入失敗</div>");
function u16Clickudf2167a7c6a44f9fa0a13fb29fa4bdb4(e)
{

	SetPanelVisibility('u17','','none',500);

	ToggleLinks(e, 'u16LinksClick');
}
gv_vAlignTable['u12'] = 'center';gv_vAlignTable['u8'] = 'top';gv_vAlignTable['u19'] = 'center';u10.tabIndex = 0;

u10.style.cursor = 'pointer';
$axure.eventManager.click('u10', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_未登_.html');

}
});
gv_vAlignTable['u22'] = 'center';u25.tabIndex = 0;

u25.style.cursor = 'pointer';
$axure.eventManager.click('u25', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('About_TutorMing_1.html');

}
});
gv_vAlignTable['u1'] = 'center';gv_vAlignTable['u9'] = 'top';gv_vAlignTable['u14'] = 'center';gv_vAlignTable['u20'] = 'top';u23.tabIndex = 0;

u23.style.cursor = 'pointer';
$axure.eventManager.click('u23', function(e) {

if (true) {

	SetPanelVisibility('u17','hidden','none',500);

}
});
gv_vAlignTable['u24'] = 'top';gv_vAlignTable['u7'] = 'center';