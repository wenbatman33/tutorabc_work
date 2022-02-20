for(var i = 0; i < 16; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'center';u12.tabIndex = 0;

u12.style.cursor = 'pointer';
$axure.eventManager.click('u12', u12Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u12LinksClick'></div>")
var u12LinksClick = document.getElementById('u12LinksClick');
function u12Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u12LinksClick');
}

InsertBeforeEnd(u12LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u12Clicku8e78bf7dad55436a8e91d518563bb874(event)'>未登入</div>");
function u12Clicku8e78bf7dad55436a8e91d518563bb874(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_未登_.html');

	ToggleLinks(e, 'u12LinksClick');
}

InsertBeforeEnd(u12LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u12Clicku5b54937545d040d2a780ee956c93d9bd(event)'>已登入</div>");
function u12Clicku5b54937545d040d2a780ee956c93d9bd(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_已登_.html');

	ToggleLinks(e, 'u12LinksClick');
}
u15.tabIndex = 0;

u15.style.cursor = 'pointer';
$axure.eventManager.click('u15', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('Record_Video_1.html');

}
});
gv_vAlignTable['u5'] = 'center';gv_vAlignTable['u1'] = 'center';gv_vAlignTable['u9'] = 'center';gv_vAlignTable['u14'] = 'center';gv_vAlignTable['u11'] = 'center';gv_vAlignTable['u7'] = 'center';