<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<!--#include virtual="/lib/include/html/template/template_innerpage_full_vipEnvNetSpeedTest.asp"--><%'特殊情況 include vip網測特用template%>
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/lib/include/testLog.inc"-->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--

%>
<%
call envirService.getListOfFqdn( objPCInfo )
''''''''''''''''''''''''''''''''''''''''''''
'POST
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	Dim sql2
	Set objPCInfo2 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql2 = " select * from NetworkTestingRecord with(nolock) where clientSn = @clientSn and valid = 1 "
	intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
	'將舊的資料庫valid變 0
		Dim myValidArray()
		redim myValidArray(objPCInfo2.RecordCount)
		while not objPCInfo2.EOF 
			myValidArray( objPCInfo2.index ) = objPCInfo2("sn")
			objPCInfo2.moveNext()
		wend
		For Each item In myValidArray
			sql2 = " update NetworkTestingRecordDetail set valid=0  where parentSn = @parentSn and valid = 1 "
			intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( item ) , CONST_TUTORABC_RW_CONN)
		Next
		sql2 = " update NetworkTestingRecord set valid=0  where clientSn = @clientSn and valid = 1 "
		intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
	'將舊的資料庫valid變 0 end
	'新增測試的速度
		sql2 = " INSERT INTO dbo.NetworkTestingRecord ( ClientSn, Remark, Valid )  VALUES  ( @ClientSn , N'首測重測', 1 ) "
		intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
		
		sql2 = " select * from NetworkTestingRecord with(nolock) where clientSn = @clientSn and valid = 1 "
		intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
		Dim insertParentSn:insertParentSn = 0
		if not objPCInfo2.EOF then
			insertParentSn = objPCInfo2("sn")
		end if
		
		
	
	Dim con_all_good : con_all_good = 0
	Dim con_all_bad : con_all_bad = 0
	Dim con_all_state_good : con_all_state_good = 0 '若所有網路測速品質皆良好or都不佳?
	Dim serverLocation : serverLocation =""  '依據設定喜愛教室。設定
	Dim priority : priority = 999
	Dim priority_tmp : priority_tmp = 0
	Dim netDate
	if not insertParentSn =0 then	
		Dim RecordCount : RecordCount =0
		RecordCount = objPCInfo.RecordCount
		ReDim UploadRate_arr(RecordCount)
		ReDim DownloadRate_arr(RecordCount)
		ReDim serverLocation_arr(RecordCount)
		while not objPCInfo.EOF
			UploadRate_ = getRequest("UploadRate_"+objPCInfo("Sn"), 1)
			DownloadRate_ = getRequest("DownloadRate_"+objPCInfo("Sn"), 1)
			UploadRate_arr(objPCInfo.Index) = UploadRate_
			DownloadRate_arr(objPCInfo.Index) = DownloadRate_
			serverLocation_arr(objPCInfo.Index) = objPCInfo("serverLocation")
			
			priority_tmp = objPCInfo("priority")
			
			
			'若所有網路測速品質皆良好or都不佳?
			'品質都很好
			writeTestLog ( UploadRate_ &">"&CFG_MIN_UPLOAD_SPEED&" and "&DownloadRate_&">"&CFG_MIN_DOWNLOAD_SPEED&" - "&   (UploadRate_*1 > CFG_MIN_UPLOAD_SPEED and  DownloadRate_*1 > CFG_MIN_DOWNLOAD_SPEED)   )
			if UploadRate_*1 > CFG_MIN_UPLOAD_SPEED and  DownloadRate_*1 > CFG_MIN_DOWNLOAD_SPEED then
				con_all_good = con_all_good+1
				writeTestLog ("priority_tmp:"&priority_tmp)
				writeTestLog ("priority:"&priority)
				writeTestLog ("priority_tmp <= priority :"&   priority_tmp*1 <= priority  )
				if priority_tmp*1 <= priority then
					priority = priority_tmp*1
					writeTestLog (  objPCInfo("serverLocation") )
					serverLocation = objPCInfo("serverLocation")
				end if
			'品質都不好
			'elseif UploadRate_*1 < CFG_MIN_UPLOAD_SPEED and  DownloadRate_*1 < CFG_MIN_DOWNLOAD_SPEED then
				con_all_bad = con_all_bad+1
			end if			
			writeTestLog ("con_all_state_good:"&con_all_state_good &"/con_all_good:"&con_all_good&"/con_all_bad:"&con_all_bad)
			
			'若所有網路測速品質皆良好or都不佳? END
			
			sql2 = " INSERT INTO dbo.NetworkTestingRecordDetail ( ParentSn, UploadRate, DownloadRate , FqdnSn , Valid )  VALUES  ( @ParentSn , @UploadRate , @DownloadRate , @FqdnSn , 1 ) "
			intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( insertParentSn ,UploadRate_ , DownloadRate_ , objPCInfo("sn")  ) , CONST_TUTORABC_RW_CONN)
			
			objPCInfo.moveNext()
		wend
		if con_all_bad=RecordCount*1 or con_all_bad=RecordCount*1  then
			con_all_state_good = 1
		end if
		
		countResult = logService.CountNetworkStatus (objPCInfo2 , serverLocation_arr,UploadRate_arr,DownloadRate_arr)
		set logService = nothing
		netDate =  envirService.saveEnvirNetwork( objPCInfo2 , insertParentSn )  '資訊寫入 環境測試資料表
		serverLocation = netDate(2) '從envirService.saveEnvirNetwork 取得的變數
	end if
	'若有一處以上(含)網路品質不佳，(則依dbo.NetworkTestingFqdn的Priority欄位為依據設定喜愛教室。)<--規則改變  改為 取另外兩個伺服器下載數值最高為喜愛位置(規則變更)。
	writeTestLog ("serverLocation:"&serverLocation)
	writeTestLog ("con_all_state_good:"&con_all_state_good)
	if con_all_state_good = 0 and serverLocation<>"" then
		'規則改變  改為 取另外兩個伺服器下載數值最高為喜愛位置(規則變更)。
		
		if CFG_SERVER_LOCATION_DIC.Exists(serverLocation) then
			sql2 = " INSERT INTO dcgs_client_dcgs_tag_bas (client_sn , cdt_category_sn,cdt_specific_category_sn) VALUES (@clientSn,53,"&CFG_SERVER_LOCATION_DIC(serverLocation)&") "
			intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
		end if
	end if
	'新增測試的速度 end	
	'依據GIS建議數值進行判斷，若所有網路測速品質皆良好or都不佳，則刪除客戶的喜愛教室位置：
	if con_all_state_good = 1 then
		sql2 = " DELETE FROM dbo.dcgs_client_dcgs_tag_bas WHERE cdt_category_sn=53 AND client_sn = @clientSn "
		intResult = objPCInfo2.excuteSqlStatementEach( sql2 , Array( session( "client_sn") ) , CONST_TUTORABC_RW_CONN)
	end if
	
	
	objPCInfo2.close                              '關閉物件
	Set objPCInfo2 = nothing                       '設定物件為空

	response.redirect("guide_EnvirVolumn.asp?CountNetworkStatus=<%=countResult%>")
end if
%>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.timers.js"></script>
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<link type="text/css" href="/lib/javascript/JQuery/css/ui-lightness_white/jquery-ui-1.8.7.custom.css" rel="stylesheet" />


<script>

<%'取得flash 回傳回來的各網址的值
%>
var SpeedDataInfoArr = new Array();
var SpeedDataInfoArr_i = 0;
function SpeedDataInfo( arr ){
	SpeedDataInfoArr[SpeedDataInfoArr_i] = arr;
	SpeedDataInfoArr_i++;
	show_SpeedDataInfo();
}

<%
' 取得flash 網速的值
' 1.SN
' 2.下載檔案的大小
' 3.下載檔案的時間
' 4.下載檔案的速度
' 5.上傳檔案的時間
' 6.上傳檔案的速度
%>
var fqdn_count = 0;
var max_speed_sum = 0 ;
var max_speed_upload = 0;
var max_speed_download = 0 ;
var max_SpeedDataInfoArr_i ;
function show_SpeedDataInfo(){
	msg_tmp ="";
	for(i=0 ; i<SpeedDataInfoArr.length ; i++){
		$("#DownloadRate_"+SpeedDataInfoArr[i][0]).val(  parseInt(SpeedDataInfoArr[i][3]) );
		$("#UploadRate_"+SpeedDataInfoArr[i][0]).val(  parseInt(SpeedDataInfoArr[i][5]) );
		
		if (   (SpeedDataInfoArr[i][3]+SpeedDataInfoArr[i][5]) > max_speed_sum ){
			max_speed_sum = SpeedDataInfoArr[i][3]+SpeedDataInfoArr[i][5];
			max_speed_download = parseInt(SpeedDataInfoArr[i][3]);
			max_speed_upload = parseInt(SpeedDataInfoArr[i][5]);
			max_SpeedDataInfoArr_i = SpeedDataInfoArr[i][0];
		}
	
		//for(j=0 ; j<SpeedDataInfoArr[i].length ; j++){
		//	msg_tmp += "\n"+SpeedDataInfoArr[i][j] ;
		//}
		msg_tmp += "\n"+SpeedDataInfoArr[i][6] ;
	}
	fqdn_count++;
	
	$.ajax({
		type: "POST",
		cache: false,
		url: "guide_EnvirNetwork_Exe_ajax.asp",
		data: "netState=" + msg_tmp,
		async: false,
		success: function (data)
		{
			//alert('success');
		},
		error: function ()
		{
			//alert('error');
		}
	});
	
	//全部資料讀取完畢
	if ( fqdn_count == <%=objPCInfo.RecordCount%> ){
		$("#guide_result").show();
		$("#flash_div").hide();
		//alert(msg_tmp);
		$("#netState").val( msg_tmp );
		
		
		drawSpeedView();
	}
	
}
//顯示 上下載速度資料
function drawSpeedView(){
	$("#upload_speed_text").html( max_speed_upload/1000 );
	$("#download_speed_text").html( max_speed_download/1000 );
	if (  max_speed_download < <%=CFG_MIN_DOWNLOAD_SPEED%>  ||  max_speed_upload < <%=CFG_MIN_UPLOAD_SPEED%>){
		$("#result3").show();
	}else if(  (max_speed_download >= <%=CFG_MIN_DOWNLOAD_SPEED%> &&  max_speed_download < <%=CFG_BEST_DOWNLOAD_SPEED%>)  ||  (  max_speed_upload >= <%=CFG_MIN_UPLOAD_SPEED%> && max_speed_upload < <%=CFG_BEST_UPLOAD_SPEED%>  ) ){
		$("#result2").show();
	}else{
		$("#result1").show();
	}
}


<%'產生flash並 傳值給flash值
%>
function getFlashMovieObject(movieName){
	if (window.document[movieName]){
		return window.document[movieName];
	}else if (navigator.appName.indexOf("Microsoft Internet")==-1){
	if (document.embeds && document.embeds[movieName])
		return document.embeds[movieName];
	}else{
		return document.getElementById(movieName);
	}
}

$(document).ready(function() {
	var UrlData = new Array( <%= objPCInfo.RecordCount %> ); 
	<%
	'傳送設定傳給flash  
	' 1．下載網址
	' 2．上傳網址
	' 3．fqdn sn
	Dim i :i=0 
	objPCInfo.moveFirst()
	while not objPCInfo.EOF
	%>
		//UrlData[<%=i%>] = new Array('http://<%=objPCInfo("Fqdn")%>/TestSite/Penguins.jpg','http://<%=objPCInfo("Fqdn")%>/TestSite/TestUpload2.asp','<%=objPCInfo("Sn")%>');
		UrlData[<%=i%>] = new Array('http://<%=objPCInfo("Fqdn")%>/images/f1265612690837.jpg','http://<%=objPCInfo("Fqdn")%>/TestUpload2.asp','<%=objPCInfo("Sn")%>');
		//UrlData[<%=i%>] = new Array('http://<%=objPCInfo("Fqdn")%>/env_test/f1265612690837.jpg','http://<%=objPCInfo("Fqdn")%>/env_test/TestUpload2.asp','<%=objPCInfo("Sn")%>');
	<%
		i = i+1
		objPCInfo.moveNext()
	wend%>
	//new Array( new Array('http://192.168.23.174/images/f1265612690837.jpg','','1') ,new Array("http://192.168.23.180/TestSite/Penguins.jpg",'http://192.168.23.180/TestSite/TestUpload2.asp','2') );	
	//myContent.SendFromJavascriptUrlData(UrlData);
	//swfobject.getObjectById("myContent").SendFromJavascriptUrlData(UrlData);
	//var obj2 = getFlashMovieObject("myContent_1");
	<%
	'判斷瀏覽器來 使javascript呼叫flash 可以正確抓到object (因為ff跟ie抓法不同)
	%>
	var flashObject = "myContent_ie";
	var isIE = navigator.userAgent.search("MSIE") > -1; 
	var isIE7 = navigator.userAgent.search("MSIE 7") > -1; 
	var isFirefox = navigator.userAgent.search("Firefox") > -1; 
	var isOpera = navigator.userAgent.search("Opera") > -1; 
	var isSafari = navigator.userAgent.search("Safari") > -1;//Google瀏覽器是用這核心 
	if ( !isIE ){ flashObject = "myContent_ff"; }
	
	$(this).oneTime(2000, function() {
			getFlashMovieObject(flashObject).SendFromJavascriptUrlData(UrlData);
	});
	
	
	$("#con_quality_bad_msg").dialog({
		autoOpen: false,
		title: '資料',
		bgiframe: true,
		width: 700,
		height: 350,
		modal: true,
		draggable: true,
		resizable: false,
		overlay:{opacity: 0.7, background: "#FF8899" }, 
		buttons: {
			'關閉': function() {
				$(this).dialog('close');
			}
		}
	});

	$( ".show_pop" ).click(function() {
		$("#con_quality_bad_msg").dialog("open");
	});
		
});

</script>
<%
'objPCInfo.close                               '關閉物件
'Set objPCInfo = nothing                       '設定物件為空
%>