var entitydata;
var table;
var starttime;
var endtime;
var ssdd;
var sszd;
var search;
var type=2;
var ssddtext;
var tablezd;
var seltype;
var pagecount;
var todaytotaldata = [];
var searchtext;
var deviceselect = 2;
$(function () {

    $("#header").load('../Top.aspx', function () {
        $("#header ul li:eq(3)").addClass("active");
    });
});


function Divselect(number,mythis)
{
    //window.event ? window.event.cancelBubble = true : e.stopPropagation();
    //resizetbwidth(number);
    deviceselect = number;
    type = number;
    
    $(".cell-button").removeClass("activeButton");

    $(mythis).addClass("activeButton");


    $("#L_Type").html($(mythis).html());


    

    if ($('.end_form_datetime').html() < $('.start_form_datetime').html()) {
        $("#alertmodal").modal("show");
        return;
    };


    if (!table) {
        createDataTable();
       
    } else {


        $("#search-result-table").DataTable().ajax.reload();
 
    }

}



function resizetbwidth(number) {
    switch (number) {
        case 1:   //车载视频
        case 2:
        case 3:
        case 7:
           
            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '100px');
            $('#search-result-table tr:eq(0) th:eq(2)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(3)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(4)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(5)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(6)').css('width', '80px');
            break;
        case 4:   //警务通
        case 6://辅警通
           
            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '60px');
            $('#search-result-table tr:eq(0) th:eq(2)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(3)').css('width', '40px');
            $('#search-result-table tr:eq(0) th:eq(4)').css('width', '70px');
            $('#search-result-table tr:eq(0) th:eq(5)').css('width', '60px');
            $('#search-result-table tr:eq(0) th:eq(6)').css('width', '40px');
            $('#search-result-table tr:eq(0) th:eq(7)').css('width', '85px');
            $('#search-result-table tr:eq(0) th:eq(8)').css('width', '105px');
            $('#search-result-table tr:eq(0) th:eq(9)').css('width', '120px');
            break;
        case 5: //执法记录仪
  
            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '80px');
            $('#search-result-table tr:eq(0) th:eq(2)').css('width', '75px');
            $('#search-result-table tr:eq(0) th:eq(3)').css('width', '85px');
            $('#search-result-table tr:eq(0) th:eq(4)').css('width', '95px');
            $('#search-result-table tr:eq(0) th:eq(5)').css('width', '95px');
            $('#search-result-table tr:eq(0) th:eq(6)').css('width', '70px');
            $('#search-result-table tr:eq(0) th:eq(7)').css('width', '90px');
            $('#search-result-table tr:eq(0) th:eq(8)').css('width', '65px');

            break;
    }
}



//$(document).on('click.bs.carousel.data-api', '#requestbtn', function (e) {
//    if ($('.end_form_datetime').val() < $('.start_form_datetime').val()) {
//        $("#alertmodal").modal("show");
//        return;
//    };
//    $(".tablediv label:eq(0)").text("| " + $("#deviceselect").find("option:selected").text() + "报表")

//    if (!table) {
//        createDataTable();
//    } else {

//        $("#search-result-table").DataTable().ajax.reload();
//    }
//});


switch (true) {
    case window.screen.height > 910:
        pagecount = 10;
        break;
    case window.screen.height > 850:
        pagecount = 8;
        break;
     default:
         pagecount = 5;
        break;

}
function transferDate(date) {
    // 年  
    var year = date.getFullYear();
    // 月  
    var month = date.getMonth() + 1;
    // 日  
    var day = date.getDate();
    if (month >= 1 && month <= 9) {
        month = "0" + month;
    }
    if (day >= 0 && day <= 9) {
        day = "0" + day;
    }
    var dateString = year + '/' + month + '/' + day;
    return dateString;
}


function hbdatetime(date) {
    var curDate = new Date(date);
    return transferDate(new Date(curDate.getTime() - 7 * 24 * 60 * 60 * 1000));
}

//startdatetimedefalute();

function Next_Last_Week(last,next)
{
    

    var curDate = new Date();
  
    var day = curDate.getDay(); //从周日 到周六 为一个星期  （0-6）

   


    if (day != 0) {//计算时间从周一到周日
            var preDate = new Date(curDate.getTime() - 24 * (day) * 60 * 60 * 1000); //前day天  结束时间

            var beforepreDate = new Date(curDate.getTime() - 24 * (day + 6) * 60 * 60 * 1000); //前7天+day 天 开始时间   
    }


    else {
        var preDate = new Date(curDate.getTime()); // 结束时间

        var beforepreDate = new Date(curDate.getTime() - 24 * (6) * 60 * 60 * 1000); // 开始时间

      }


        if (last) //上一周
        {
            var lastDate = new Date($('.start_form_datetime').html());

            var preDate = new Date(lastDate.getTime() - 24 * (1) * 60 * 60 * 1000); //前day天  结束时间

            var beforepreDate = new Date(lastDate.getTime() - 24 * (7) * 60 * 60 * 1000); //前7天+day 天 开始时间
        }

        else if (next)//下一周
        {
            var nextDate = new Date($('.end_form_datetime').html());

            var preDate = new Date(nextDate.getTime() + 24 * (7) * 60 * 60 * 1000); //前day天  结束时间

            var beforepreDate = new Date(nextDate.getTime() + 24 * (1) * 60 * 60 * 1000); //前7天+day 天 开始时间
        }

    $('.end_form_datetime').html(transferDate(preDate));//结束时间


    $('.start_form_datetime').html(transferDate(beforepreDate)); //开始时间

    if (last || next) {

        if (!table) {

            createDataTable();
        } else {

            $("#search-result-table").DataTable().ajax.reload(null, false);
        }
    }


}


//$.ajax({
//    type: "POST",
//    url: "../Handle/GetEntitys.ashx",
//    data: "",
//    dataType: "json",
//    success: function (data) {
//        if (data.title != "331000000000") {
//            $("#brigadeselect").attr("disabled", "disabled");
//        }
//        entitydata = data.data; //保存单位数据
//        for (var i = 0; i < entitydata.length; i++) {

//            if (entitydata[i].SJBM == "331000000000") {
//                if (data.title == entitydata[i].BMDM) {
//                    $("#brigadeselect").append("<option value='" + entitydata[i].BMDM + "' selected>" + entitydata[i].BMJC + "</option>");
//                    changesquadronselect(entitydata[i].BMDM);
//                }
//                else {
//                    $("#brigadeselect").append("<option value='" + entitydata[i].BMDM + "'>" + entitydata[i].BMJC + "</option>");
//                }
//            }
//        }
//        switch (data.title) {
//            case "331000000000":
//            case "331001000000":
//            case "331002000000":
//            case "331003000000":
//            case "331004000000":
//            case "33100000000x":
//                break;
//            default:
//                $("#squadronselect").attr("disabled", "disabled");
//                changeentitysel(data.title)
//                break;
//        }

//    },
//    error: function (msg) {
//        console.debug("错误:ajax");
//    }
//});

$.ajax({
    type: "POST",
    url: "../Handle/dataManagementConfig.ashx",
    data: { 'requesttype': 'Request' },
    dataType: "json",
    success: function (data) {
        if (data.r == "0") {
            var val = data.result[0].val.split("|");
            $("#configmodal input:eq(0)").val(val[0]);
            $("#configmodal input:eq(1)").val(val[1]);
        }
    },
    error: function (msg) {
        console.debug("错误:ajax");
    }
});

$(document).on('click.bs.carousel.data-api', '#configmodal .btn-save', function (e) {
    $(".alainfo").hide();
   var _val1 = $("#configmodal input:eq(0)").val();
   var _val2 = $("#configmodal input:eq(1)").val();
    if (_val1 == "") {
        createUserAlarm($("#configmodal .date:eq(0)"), "设备使用标准不能为空");
        return;
    }
    if (_val2 == "") {
        createUserAlarm($("#configmodal .date:eq(1)"), "设备在线标准不能为空");
        return;
    }

    $.ajax({
        type: "POST",
        url: "../Handle/dataManagementConfig.ashx",
        data: { 'requesttype': 'update', 'val': _val1 +"|"+_val2 },
        dataType: "json",
        success: function (data) {
            
            $("#configmodal").modal("hide");

        },
        error: function (msg) {
            console.debug("错误:ajax");
            $("#configmodal").modal("hide");
        }
    });


});


function createUserAlarm($ele, txt) {
    var $doc = $ele;
    $doc.find("label").show();
}


function changesquadronselect(brigadeselectvalue) {
    $("#squadronselect").empty();
    $("#squadronselect").append("<option value='all'>全部</option>");
    $("#squadronselect").removeAttr("disabled");
    for (var i = 0; i < entitydata.length; i++) {
        if (entitydata[i].SJBM == brigadeselectvalue) {
            $("#squadronselect").append("<option value='" + entitydata[i].BMDM + "'>" + entitydata[i].BMJC + "</option>");
        }
    }
}

function changeentitysel(BMDM) {
    $("#squadronselect").empty();
    $("#brigadeselect").empty();
    var SJBM;
    for (var i = 0; i < entitydata.length; i++) {
        if (BMDM == entitydata[i].BMDM) {
            $("#squadronselect").append("<option value='" + entitydata[i].BMDM + "' selected>" + entitydata[i].BMJC + "</option>");
            SJBM = entitydata[i].SJBM;
        }
    }
    for (var i = 0; i < entitydata.length; i++) {
        if (SJBM == entitydata[i].BMDM) {
            $("#brigadeselect").append("<option value='" + entitydata[i].BMDM + "' selected>" + entitydata[i].BMJC + "</option>");
        }
    }
    if ($("#brigadeselect").size() == 0) {
        $("#brigadeselect").append("<option value='0' selected>其它部门</option>");
    }
    if ($("#squadronselect").size() == 0) {
        $("#squadronselect").append("<option value='" + BMDM + "' selected>" + BMDM + "</option>");
    }
}




//更换大队选择
$(document).on('change.bs.carousel.data-api', '#brigadeselect', function (e) {
    //所属中队逻辑
    $("#squadronselect").empty();
    $("#squadronselect").append("<option value='all'>全部</option>");
    $("#squadronselect").removeAttr("disabled");
    if (e.target.value == "all") {
        $("#squadronselect").attr("disabled", "disabled");
        return;
    }
    for (var i = 0; i < entitydata.length; i++) {
        if (entitydata[i].SJBM == e.target.value) {
            $("#squadronselect").append("<option value='" + entitydata[i].BMDM + "'>" + entitydata[i].BMJC + "</option>");
        }
    }
});

function Load() {
    $(".cell-button").removeClass("activeButton");

    $(".cell-button:eq(0)").addClass("activeButton");


    $("#L_Type").html("对讲机");
    Next_Last_Week(false, false);

    //window.event ? window.event.cancelBubble = true : e.stopPropagation();
    //resizetbwidth(2);

    createDataTable();
}

Load();//加载


//参数配置
$(document).on('click.bs.carousel.data-api', '#configbtn', function (e) {
    $("#configmodal").modal("show");

});
$('.datadetail').on('hidden.bs.modal', function () {

    $("#search-result-table").find(".trselect").removeClass("trselect"); //移除选择
});
function loadTatolData() {
    var data =
   {
       search: $(".search input").val(),
       type: deviceselect,
       //ssdd: $("#brigadeselect").val(),
       sszd: $("#L_select").html(),
       begintime: $(".start_form_datetime").html(),
       endtime: $(".end_form_datetime").html(),
       hbbegintime: hbdatetime($(".start_form_datetime").html()),
       hbendtime: hbdatetime($(".end_form_datetime").html()),
       dates: datecompare($(".end_form_datetime").html(), $(".start_form_datetime").html()),
       requesttype: "查询汇总"
   }
    $.ajax({
        type: "POST",
        url: "../Handle/dataManagement.ashx",
        data: data,
        dataType: "json",
        success: function (data) {
            if (data.r=="0"){
                joinData(data);
            }
   
        },
        error: function (msg) {
            console.debug("错误:ajax");
        }
    });

}

function joinData(data) {
    var  creadata = {
        result:[]
    }
    var val1 = { Value: {} };
    var val2 = { Value: {} };
    var val3 = { Value: {} };
    var val4 = { Value: {} };
    var val5 = { Value: {} };
    var val6 = { Value: {} };
    var val7 = { Value: {} };
    var val8 = { Value: {} };
    val1.Value = todaytotaldata[0];
    creadata.result.push(val1)
    val2.Value = parseFloat(todaytotaldata[1]) * 3600;
    creadata.result.push(val2)
    val3.Value = todaytotaldata[2];
    creadata.result.push(val3)
    val4.Value = todaytotaldata[3];
    creadata.result.push(val4)
    val5.Value = data.result[0].Value;
    creadata.result.push(val5)
    val6.Value = data.result[1].Value;
    creadata.result.push(val6)
    val7.Value = data.result[2].Value;
    creadata.result.push(val7)
    val8.Value = data.result[3].Value;
    creadata.result.push(val8)
    createTatolRS(creadata);
}

function createTatolRS(data) {
    //配发量
    if (data.result[0].Value != "0" && data.result[0].Value != "" && data.result[4].Value != "" && data.result[4].Value != "0") {
        var intpf1 = parseInt(data.result[0].Value);
        var intpf2 = parseInt(data.result[4].Value);
        var tbpf = (intpf1 - intpf2) * 100 / intpf2;
        $("#ulsbpf li:eq(0)").text(intpf1);
        if (tbpf < 0) {
            $("#ulsbpf li:eq(2)").html("同比上周减少" + formatFloat(tbpf,1) + "%<i class='fa fa-arrow-down' aria-hidden='true'>");
        }
        else
        {
            $("#ulsbpf li:eq(2)").html("同比上周增加" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-up' aria-hidden='true'>");
        }
    }
    else
    {
        $("#ulsbpf li:eq(0)").text("0");
        $("#ulsbpf li:eq(2)").html("同比上周减少 --%");

    }
    //在线时长
    if (data.result[1].Value != "0" && data.result[1].Value != "" && data.result[5].Value != "" && data.result[5].Value != "0") {
        var intpf1 = parseInt(data.result[1].Value);
        var intpf2 = parseInt(data.result[5].Value);
        var tbpf = (intpf1 - intpf2) * 100 / intpf2;
        $("#ulsysc li:eq(0)").text(formatFloat(intpf1 / 3600,1)+"h");
        if (tbpf < 0) {
            $("#ulsysc li:eq(2)").html("同比上周减少" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-down' aria-hidden='true'>");
        }
        else {
            $("#ulsysc li:eq(2)").html("同比上周增加" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-up' aria-hidden='true'>");
        }
    }
    else {
        $("#ulsysc li:eq(0)").text("0");
        $("#ulsysc li:eq(2)").html("同比上周 --%");

    }
    //设备使用数量
    if (data.result[2].Value != "0" && data.result[2].Value != "" && data.result[6].Value != "" && data.result[6].Value != "0") {
        var intpf1 = parseInt(data.result[2].Value);
        var intpf2 = parseInt(data.result[6].Value);
        var tbpf = (intpf1 - intpf2) * 100 / intpf2;
        $("#ulsysl li:eq(0)").text(intpf1);
        if (tbpf < 0) {
            $("#ulsysl li:eq(2)").html("同比上周减少" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-down' aria-hidden='true'>");
        }
        else {
            $("#ulsysl li:eq(2)").html("同比上周增加" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-up' aria-hidden='true'>");
        }
    }
    else {
        $("#ulsysl li:eq(0)").text("0");
        $("#ulsysl li:eq(2)").html("同比上周 --%");

    }
    //设备在线数
    if (data.result[3].Value != "0" && data.result[3].Value != "" && data.result[7].Value != "" && data.result[7].Value != "0") {
        var intpf1 = parseInt(data.result[3].Value);
        var intpf2 = parseInt(data.result[7].Value);
        var tbpf = (intpf1 - intpf2) * 100 / intpf2;
        $("#ulzxsb li:eq(0)").text(intpf1);
        if (tbpf < 0) {
            $("#ulzxsb li:eq(2)").html("同比上周减少" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-down' aria-hidden='true'>");
        }
        else {
            $("#ulzxsb li:eq(2)").html("同比上周增加" + formatFloat(tbpf, 1) + "%<i class='fa fa-arrow-up' aria-hidden='true'>");
        }
    }
    else {
        $("#ulzxsb li:eq(0)").text("0");
        $("#ulzxsb li:eq(2)").html("同比上周 --%");

    }



}
function datecompare(end, start) {
    start = new Date(start).getTime();
    end = new Date(end).getTime();
    var time = 0
    time = end - start;
    return Math.floor(time / 86400000) + 1;
};
function formatFloat(value, y) {
    var result = Math.floor((value) * Math.pow(10, y)) / Math.pow(10, y);
    return result;
};

function eachbrigadeselect() {
    var entitys="";
    $("#brigadeselect option").each(function (index, el) {
        if (index > 0) {
            entitys += (index > 1) ? ","+($(this).val()) : $(this).val()
        }
    });
    return entitys;
}


function createDataTable() {

    var columns = [
  
                          { "data": "cloum1","orderable": false  },
                          { "data": "cloum2" },
                          { "data": "cloum3" },
                          { "data": "cloum4" },
                          { "data": "cloum5" },
                          { "data": "cloum6" },
                          { "data": "cloum7", "visible": false },
                          { "data": "cloum8", "visible": false },
                          { "data": "cloum9", "visible": false },
                          { "data": "cloum10", "visible": false }
                          //{ "data": "cloum11", "visible": false }
                          //{ "data": null, "orderable": false }
        ];


        table = $('#search-result-table')
           .on('error.dt', function (e, settings, techNote, message) {
  

           })
            //.on('preXhr.dt', function (e, settings, data) {
            //    $('.progresshz').show()
            //    //$('.btnsjx').attr("disabled", "disabled");
            //    $('#search-result-table').hide();
            //    alert("123");
            //})

             .on('xhr.dt', function (e, settings, json, xhr) {
              
                 if (json.data.length > 0) {
                     todaytotaldata.length = 0;
                     todaytotaldata.push(json.data[0]["cloum3"]);
                     todaytotaldata.push(json.data[0]["cloum4"]);
                     todaytotaldata.push(json.data[0]["cloum5"]);
                     todaytotaldata.push(json.data[0]["cloum14"]);
                 }
           
                 $('.progresshz').hide();
                 $('#search-result-table').show();
                 //$('.btnsjx').removeAttr("disabled");

                 seltype = deviceselect;
                 table.column(1).visible(true);
                 table.column(2).visible(true);
                 table.column(3).visible(true);
                 table.column(4).visible(true);
                 table.column(5).visible(true);
                 table.column(6).visible(true);
                 table.column(7).visible(true);

                 switch (deviceselect + "") {
                     case "1"://车载视频
                     case "2": //对讲机
                     case "3":   //拦截仪
                     case "7":
                 
                         //table.column(8).visible(false);
                         table.column(7).visible(false);
                         table.column(8).visible(false);
                         table.column(9).visible(false);
             
                         $('#search-result-table tr:eq(0) th:eq(0)').text("序号");
                         $('#search-result-table tr:eq(0) th:eq(1)').text("部门");
                         $('#search-result-table tr:eq(0) th:eq(2)').text("设备配发数(台)");
                         $('#search-result-table tr:eq(0) th:eq(3)').text("设备使用数量");
                         $('#search-result-table tr:eq(0) th:eq(4)').text("在线时长总和（h）");
                         $('#search-result-table tr:eq(0) th:eq(5)').text("设备使用率（%）");
                         $('#search-result-table tr:eq(0) th:eq(6)').text("使用率排名");
                     
                         break;
                     case "5":   //执法记录仪
           
                         table.column(7).visible(true);
                         table.column(8).visible(true);
                         table.column(9).visible(false);
                         $('#search-result-table tr:eq(0) th:eq(0)').text("序号");
                         $('#search-result-table tr:eq(0) th:eq(1)').text("部门");
                         $('#search-result-table tr:eq(0) th:eq(2)').text("设备配发数(台)");
                         $('#search-result-table tr:eq(0) th:eq(3)').text("设备使用数量(台)");
                         $('#search-result-table tr:eq(0) th:eq(4)').text("设备未使用数量(台)");
                         $('#search-result-table tr:eq(0) th:eq(5)').text("视频时长总和（h）");
                         $('#search-result-table tr:eq(0) th:eq(6)').text("视频大小(GB)");
                         $('#search-result-table tr:eq(0) th:eq(7)').text("设备使用率（%）");
                         $('#search-result-table tr:eq(0) th:eq(8)').text("使用率排名");
                         break;

                     case "4"://警务通
                         table.column(7).visible(true);
                         table.column(8).visible(true);
                         table.column(9).visible(true);
                         $('#search-result-table tr:eq(0) th:eq(0)').text("序号");
                         $('#search-result-table tr:eq(0) th:eq(1)').text("部门");
                         $('#search-result-table tr:eq(0) th:eq(2)').text("设备配发数(台)");
                         $('#search-result-table tr:eq(0) th:eq(3)').text("警员数");
                         $('#search-result-table tr:eq(0) th:eq(4)').text("警务通处罚数");
                         $('#search-result-table tr:eq(0) th:eq(5)').text("人均处罚量");
                         $('#search-result-table tr:eq(0) th:eq(6)').text("查询量");
                         $('#search-result-table tr:eq(0) th:eq(7)').text("设备平均处罚量");
                         $('#search-result-table tr:eq(0) th:eq(8)').text("设备平均处罚量排名");

                         $('#search-result-table tr:eq(0) th:eq(9)').text("无采集违停设备(台)")


                         break;

                     case "6"://辅警通

                         table.column(7).visible(true);
                         table.column(8).visible(true);
                         table.column(9).visible(true);
                         $('#search-result-table tr:eq(0) th:eq(0)').text("序号");
                         $('#search-result-table tr:eq(0) th:eq(1)').text("部门");
                         $('#search-result-table tr:eq(0) th:eq(2)').text("设备配发数(台)");
                         $('#search-result-table tr:eq(0) th:eq(3)').text("辅警数");
                         $('#search-result-table tr:eq(0) th:eq(4)').text("违停采集(例)");
                
                         $('#search-result-table tr:eq(0) th:eq(5)').text("人均处罚量");
                         $('#search-result-table tr:eq(0) th:eq(6)').text("查询量");
                         $('#search-result-table tr:eq(0) th:eq(7)').text("设备平均处罚量");
                         $('#search-result-table tr:eq(0) th:eq(8)').text("设备平均处罚量排名");
                        
                         $('#search-result-table tr:eq(0) th:eq(9)').text("本月无采集违停设备(台)")

           
                         break;
                     default:
                         break;
                 }
                 $('.infodiv').html("<span>共 " + json.data.length + " 条记录</span>");
                 //$('.daochu').html("<a class='buttons-excel'  href='../Handle/upload/" + json.title + "'><span>导 出</span></a>");
                 resizetbwidth(seltype);
                 //rebuildsjx();
                 //loadTatolData();

             })
            .on('draw.dt', function () {
                //给第一列编号
     
            })

            .DataTable({
                ajax: {
                    url: "../Handle/Window1.ashx",
                    type: "POST",
                    data: function () {
                        starttime = $(".start_form_datetime").html();
                        endtime = $(".end_form_datetime").html()
                        //ssdd = $("#brigadeselect").val();
                        sszd = $("#L_select").html();
                        search = $(".seach-box input").val();
                        type = deviceselect;
                        //searchtext = $(".search input").val();
                       
                        return data = {
                            //search: $(".search input").val(),
                            type: deviceselect,
                            //ssdd: $("#brigadeselect").val(),
                            sszd: $("#L_select").html(),
                            ssdd1: eachbrigadeselect(),
                            begintime: $(".start_form_datetime").html(),
                            endtime: $(".end_form_datetime").html(),
                            dates: datecompare($(".end_form_datetime").html(), $(".start_form_datetime").html()),
                            ssddtext: $("#brigadeselect").find("option:selected").text(),
                            sszdtext:$("#squadronselect").find("option:selected").text(),
                            requesttype: "查询报表",
                            URL:$("#L_Url").html(),
                            //onlinevalue: $("#configmodal input:eq(1)").val(),
                            //usedvalue:$("#configmodal input:eq(0)").val()
                        }
                    }

                },
                Paginate: true,
                pageLength: pagecount,
                Processing: true, //DataTables载入数据时，是否显示‘进度’提示  
                serverSide: false,   //服务器处理
                responsive: true,
                paging: true,
                autoWidth: true,
                "order": [],
                columns: columns,
                //columnDefs: [
                //         {
                //             targets:11,
                //             render: function (a, b, c, d) { var html = "<a  class=\'btn btn-sm btn-primary txzs-btn\' id='addedit' entityid='" + c.cloum12 + "'  >查看详情</a>"; return html; }
                //         }
                //],
                buttons: [ ],
                "language": {
                    "lengthMenu": "_MENU_每页",
                    "zeroRecords": "没有找到记录",
                    "info": "第 _PAGE_ 页 ( 总共 _PAGES_ 页 )",
                    "infoEmpty": "无记录",
                    "infoFiltered": "(从 _MAX_ 条记录过滤)",
                    "search": "查找设备:",
                    "paginate": {
                        "previous": "上一页",
                        "next": "下一页"
                    }
                },

                dom: "" + "t" + "<'row'p><'row infodiv'>",

                initComplete: function () {}
            });
}

$(document).on('click.bs.carousel.data-api', '.btnsjx', function (e) {
    if ($(this).attr("disabled") == "disabled") {
        return;
    }
    var $doc = $(this).text();
    if ($doc == "显示数据项") {
        $('#shujuxiang').css("display", "inline-block");
        $(this).text("隐藏数据项");
        rebuildsjx();
    }
    else
    {
        $('#shujuxiang').css("display", "none");
        $(this).text("显示数据项");
    }
});

