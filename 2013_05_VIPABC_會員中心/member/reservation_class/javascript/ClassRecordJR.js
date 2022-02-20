function class_record_VIPJR(ProfileId, ProfileLearnLevel, ContractId, ClassStartDate, ClassStartTime, ClassCostPoints, ClassTypeId, ClassIntervalId, LobbySessionId, ClassStatusId, BrandId, StaffId, client_sn, contract_sn)
{
    //alert("123");
    $.ajax({
        url: "/program/aspx/JR/ClassRecord/ClassRecordViews",
        cache: false,
        async: false,
        type : "get",
		//dataType: 'json',
		//data : MyJson,
		data:{
		    ProfileId: ProfileId,
			ProfileLearnLevel: ProfileLearnLevel,
			ContractId: ContractId,
			ClassStartDate: ClassStartDate,
			ClassStartTime: ClassStartTime,
			ClassCostPoints: ClassCostPoints,
			ClassTypeId: ClassTypeId,
			ClassIntervalId: ClassIntervalId,
            LobbySessionId : LobbySessionId,
			ClassStatusId: ClassStatusId,
			BrandId: BrandId,
			StaffId: StaffId,
			client_sn: client_sn,  //old
            contract_sn : contract_sn //old contract_sn
            //attend_list_sn: attend_list_sn  //old
		},
        success: function (result) {
            //alert("true");
		},
        error: function (result) {
            //alert("error:" + result);
        }
    });
	//return result;
}