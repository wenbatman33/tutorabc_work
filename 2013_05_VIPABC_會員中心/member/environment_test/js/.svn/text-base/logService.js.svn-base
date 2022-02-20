jQuery.fn.exists = function(){return ($(this).length > 0);}



$(document).ready(function() {
		$(".FTPStep0").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=FTPClickStep0') );
		$(".FTPStep1").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=FTPClickStep1') );
		$(".FTPStep2").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=FTPClickStep2') );
		$(".FTPStep3").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=FTPClickStep3') );
		$(".FTPStep4").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=FTPClickStep4') );
		
		$(".QSB").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=QSB') );
		$(".QRB").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=QRB') );
		
		$(".javaDownload").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=javaDownload') );
		$(".flashDownload").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=flashDownload') );
		$(".joinNetDownload").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=joinNetDownload') );
		
		$(".InternetStep_out").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=internetStep_out') );
		
		$(".orderTest").click( eventCallback ('EnvironmentTestLogAjaxService.asp','testLogKind=orderTest') );
});



function eventCallback ( url , sendData){
	return function() {
		ExeAjaxFunction ( url , sendData);
	}
}

function ExeAjaxFunction ( url , sendData){
	//alert('xxx');
	//console.log("ExeAjaxFunction");
	$.ajax({
		type: "POST",
		cache: false,
		url: url,
		data: sendData,
		async: false,
		success: function (data)
		{
		},
		error: function (xhr, ajaxOptions, thrownError)
		{
		}
	});
}