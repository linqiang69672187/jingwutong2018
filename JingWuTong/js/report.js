var entitydata;
var table;
var starttime;
var endtime;
var ssdd;
var sszd;
var search;
var type = 2;
var ssddtext;
var tablezd;
var seltype;
var pagecount;
var todaytotaldata = [];
var searchtext;
var deviceselect = 2;
var selEntityID;
var time1;
var table1;
//var columns = [];

//设备类型
$.ajax({
    type: "POST",
    url: "../Handle/getDevices.ashx",
    data: { 'requesttype': 'huizong' },
    dataType: "json",
    success: function (data) {
        var data = data.data;
        for (var i = 0; i < data.length; i++) {
            $("#deviceselect").append("<option value='" + data[i].ID + "' >" + data[i].TypeName + "</option>");

        }
    },
    error: function (msg) {
        console.debug("错误:ajax");
    }
});



$.ajax({
    type: "POST",
    url: "../Handle/ReportTimes.ashx",
    data: "meth=load",
    dataType: "text",
    success: function (msg) {
        time1 = msg;
    },
    error: function (msg) {

        console.debug("错误:ajax");
    }
});



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
//时间
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
$('.start_form_datetime,.end_form_datetime').datetimepicker({
    format: 'yyyy/mm/dd',
    autoclose: true,
    todayBtn: true,
    minView: 2
});

function startdatetimedefalute() {
    var curDate = new Date();
    var preDate = new Date(curDate.getTime() - 24 * 60 * 60 * 1000); //前一天
    var beforepreDate = new Date(curDate.getTime() - 48 * 60 * 60 * 1000); //前一天
    $('.start_form_datetime').val(transferDate(beforepreDate));
    $('.end_form_datetime').val(transferDate(preDate));
}

function hbdatetime(date) {
    var curDate = new Date(date);
    return transferDate(new Date(curDate.getTime() - 7 * 24 * 60 * 60 * 1000));
}
//时间
startdatetimedefalute();

//部门
$.ajax({
    type: "POST",
    url: "../Handle/GetEntitys.ashx",
    data: "",
    dataType: "json",
    success: function (data) {
        if (data.title != "331000000000") {
            $("#brigadeselect").attr("disabled", "disabled");
        }
        entitydata = data.data; //保存单位数据
        for (var i = 0; i < entitydata.length; i++) {

            if (entitydata[i].SJBM == "331000000000") {
                if (data.title == entitydata[i].BMDM) {
                    $("#brigadeselect").append("<option value='" + entitydata[i].BMDM + "' selected>" + entitydata[i].BMJC + "</option>");
                    changesquadronselect(entitydata[i].BMDM);
                }
                else {
                    $("#brigadeselect").append("<option value='" + entitydata[i].BMDM + "'>" + entitydata[i].BMJC + "</option>");
                }
            }
        }
        switch (data.title) {
            case "331000000000":
            case "331001000000":
            case "331002000000":
            case "331003000000":
            case "331004000000":
            case "33100000000x":
                break;
            default:
                $("#squadronselect").attr("disabled", "disabled");
                changeentitysel(data.title)
                break;
        }

    },
    error: function (msg) {
        console.debug("错误:ajax");
    }
});



function createUserAlarm($ele, txt) {
    var $doc = $ele;
    $doc.find("label").show();
}

//部门
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
//部门
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




//更换大队选择 部门
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


//重置按钮
$(document).on('click.bs.carousel.data-api', '#resetbtn', function (e) {
    $("#deviceselect").val("1");
    $("#brigadeselect").val("all");
    $("#squadronselect").val("all");
    $("#squadronselect").attr("disabled", "disabled");
    startdatetimedefalute();
    $(".search input").val("");
});
//查询
$(document).on('click.bs.carousel.data-api', '#requestbtn', function (e) {
    if ($('.end_form_datetime').val() < $('.start_form_datetime').val()) {
        $("#alertmodal").modal("show");
        return;
    };
    //$(".tablediv label:eq(0)").text("| " + $("#deviceselect").find("option:selected").text() + "报表")
    //$('#search-result-table').show();
    $('#table2').hide();
    $('#table1').show();
    //HeadcreateDataTable();
    //$('#search-result-table').dataTable().fnDestroy();

    if (!table) {
     
    
        createDataTable();
    } else {
    
        $("#search-result-table").DataTable().ajax.reload();
    }
});



//时间查询
$(document).on('click.bs.carousel.data-api', '#requestbtnTime', function (e) {
    if ($('.end_form_datetime').val() < $('.start_form_datetime').val()) {
        $("#alertmodal").modal("show");
        return;
    };
    //$(".tablediv label:eq(0)").text("| " + $("#deviceselect").find("option:selected").text() + "报表")
    $('#table1').hide();
    $('#table2').show();
    //$('#search-result-table2').show();

    if (!table1) {

        createDataTableTime();
    } else {
  
        $("#search-result-table2").DataTable().ajax.reload();
    }
});



function datecompare(end, start) {
    start = new Date(start).getTime();
    end = new Date(end).getTime();
    var time = 0
    time = end - start;
    return Math.floor(time / 86400000) + 1;
};
//function formatFloat(value, y) {
//    var result = Math.floor((value) * Math.pow(10, y)) / Math.pow(10, y);
//    return result;
//};

function eachbrigadeselect() {
    var entitys = "";
    $("#brigadeselect option").each(function (index, el) {
        if (index > 0) {
            entitys += (index > 1) ? "," + ($(this).val()) : $(this).val()
        }
    });
    return entitys;
}


//function resizetbwidth() {
//    switch ($("#deviceselect").val()) {
//        case "1":   //车载视频
//        case "2":
//        case "3":
//        case "7":
//            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
//            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '200px');
//            $('#search-result-table tr:eq(0) th:eq(2)').css('width', '100px');
//            $('#search-result-table tr:eq(0) th:eq(3)').css('width', '100px');
//            $('#search-result-table tr:eq(0) th:eq(4)').css('width', '100px');
//            $('#search-result-table tr:eq(0) th:eq(5)').css('width', '100px');
//            $('#search-result-table tr:eq(0) th:eq(6)').css('width', '100px');
//            break;
//        case "4":   //车载视频
//        case "6":
//            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
//            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '200px');
//            $('#search-result-table tr:eq(0) th:eq(2)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(3)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(4)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(5)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(6)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(7)').css('width', '80px');
//            $('#search-result-table tr:eq(0) th:eq(8)').css('width', '60px');
//            $('#search-result-table tr:eq(0) th:eq(9)').css('width', '110px');
//            break;
//        case "5":
//            $('#search-result-table tr:eq(0) th:eq(0)').css('width', '40px');
//            $('#search-result-table tr:eq(0) th:eq(1)').css('width', '200px');

//            break;
//    }
//}



//时间统计查询
function createDataTableTime() {


    var arry = time1.split(',');

    var chansu = arry[5].split('|');


    var columns = [
                      { "data": "cloum1", "orderable": false },
                      { "data": "cloum2" },
                      { "data": "cloum3" },
                      { "data": "cloum4" },
                      { "data": "cloum5" },

                      { "data": "cloum6" },
                      { "data": "cloum7" },
                      { "data": "cloum8" },
                      { "data": "cloum9" },
                      { "data": "cloum10" },
                      { "data": "cloum11" },
                      { "data": "cloum12" },
                      { "data": "cloum13" },
                      { "data": "cloum14" },
                      { "data": "cloum15" },
                      { "data": "cloum16" },
                      { "data": "cloum17" },
                      { "data": "cloum18" },
                      { "data": "cloum19" },
                      { "data": "cloum20" },
                      { "data": "cloum21" },
                      { "data": "cloum22" },
                      { "data": "cloum23" },
                      { "data": "cloum24" },
                      { "data": "cloum25" },
                      { "data": "cloum26" },
                      { "data": "cloum27" },
                      { "data": "cloum28" },
                      { "data": "cloum29" },
                      { "data": "cloum30" },
                      { "data": "cloum31" },
                      { "data": "cloum32" }

    ];

    table1 = $('#search-result-table2')
       .on('error.dt', function (e, settings, techNote, message) {

       })
     .on('preXhr.dt', function (e, settings, data) {
                         $('.progresshz').show()
                 
                         $('#search-result-table').hide();
                     })

         .on('xhr.dt', function (e, settings, json, xhr) {


             $('.progresshz').hide();
             $('#search-result-table2').show();
             var brigade = "";
             if ($("#brigadeselect").val() != "all") {
                 brigade = $("#brigadeselect").find("option:selected").text();

             }


             var squadron = "";
             if ($("#squadronselect").val() != "all") {
                 squadron = $("#squadronselect").find("option:selected").text();

             }

             let gongshi;

             switch ($("#deviceselect").val()) {
                 case "1"://车载视频
                 case "3":   //拦截仪
                 case "2": //对讲机
                   
                   //大队中队

                         //table.column(8).visible(true);
                         //table.column(9).visible(true);
                         //table.column(10).visible(true);
                         //table.column(11).visible(true);
                         //table.column(12).visible(true);
                         //table.column(13).visible(true);
                         //table.column(14).visible(true);
                         //table.column(15).visible(true);
                         //table.column(16).visible(true);

                         var html = "<tr>" +
  "<th colspan='17' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                       "</tr>" +
                        "<tr>" +
                        "<th rowspan='2' style='text-align:center'>日期</th>" +
                      "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                       "<th colspan='3' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='3' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='3' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='3' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='3' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                         "<tr>" +
                         //---------------------------------------------------------
                       "<th>设备使用数量</th>" +
                       "<th>在线时长总和(小时)</th>" +
                       "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                   "<th>设备使用数量</th>" +
                       "<th>在线时长总和(小时)</th>" +
                       "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                   "<th>设备使用数量</th>" +
                       "<th>在线时长总和(小时)</th>" +
                       "<th>设备使用率</th>" +
                         //---------------------------------------------------------;
                        "<th>设备使用数量</th>" +
                       "<th>在线时长总和(小时)</th>" +
                       "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                        "<th>设备使用数量</th>" +
                       "<th>在线时长总和(小时)</th>" +
                       "<th>设备使用率</th>" +
                         "</tr>"

                         $("#head2").html(html);
                    gongshi = "设备使用率=设备使用数÷设备配发数量";

                     break;

                 case "4"://警务通

                         //table.column(8).visible(true);
                         //table.column(9).visible(true);
                         //table.column(10).visible(true);
                         //table.column(11).visible(true);
                         //table.column(12).visible(true);
                         //table.column(13).visible(true);
                         //table.column(14).visible(true);
                         //table.column(15).visible(true);
                         //table.column(16).visible(true);

                         var html = "<tr>" +
  "<th colspan='28' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                       "</tr>" +
                        "<tr>" +
                        "<th rowspan='2' style='text-align:center'>日期</th>" +
                      "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                       "<th rowspan='2' style='text-align:center'>警员数</th>" +
                       "<th colspan='5' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='5' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='5' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='5' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='5' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                         "<tr>" +
                         //---------------------------------------------------------
                       "<th>警务通处罚数</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无处罚数的警务通(台)</th>" +
                         //---------------------------------------------------------
                             "<th>警务通处罚数</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无处罚数的警务通(台)</th>" +
                         //---------------------------------------------------------
                             "<th>警务通处罚数</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无处罚数的警务通(台)</th>" +
                         //---------------------------------------------------------;
                           "<th>警务通处罚数</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无处罚数的警务通(台)</th>" +
                         //---------------------------------------------------------
                       "<th>警务通处罚数</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无处罚数的警务通(台)</th>" +
                         "</tr>"

                         $("#head2").html(html);
                         gongshi = "人均处罚量=警务通处罚数÷警员数，设备平均处罚量=警务通处罚数÷配发数";


                     break;
                 case "5"://执法记录仪

                         //table.column(8).visible(true);
                         //table.column(9).visible(true);
                         //table.column(10).visible(true);
                         //table.column(11).visible(true);
                         //table.column(12).visible(true);
                         //table.column(13).visible(true);
                         //table.column(14).visible(true);
                         //table.column(15).visible(true);
                         //table.column(16).visible(true);

                         var html = "<tr>" +
  "<th colspan='32' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                       "</tr>" +
                        "<tr>" +
                        "<th rowspan='2' style='text-align:center'>日期</th>" +
                      "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +

                       "<th colspan='6' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='6' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='6' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='6' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='6' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                         "<tr>" +
                         //---------------------------------------------------------
                       "<th>设备使用数量(台)</th>" +
                       "<th>设备未使用数量(台)</th>" +
                       "<th>视频时长总和(小时)</th>" +
                        "<th>视频大小(GB)</th>" +
                        "<th>48小时规范上传率</th>" +
                           "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                     "<th>设备使用数量(台)</th>" +
                       "<th>设备未使用数量(台)</th>" +
                       "<th>视频时长总和(小时)</th>" +
                        "<th>视频大小(GB)</th>" +
                        "<th>48小时规范上传率</th>" +
                           "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                             "<th>设备使用数量(台)</th>" +
                       "<th>设备未使用数量(台)</th>" +
                       "<th>视频时长总和(小时)</th>" +
                        "<th>视频大小(GB)</th>" +
                        "<th>48小时规范上传率</th>" +
                           "<th>设备使用率</th>" +
                         //---------------------------------------------------------;
                     "<th>设备使用数量(台)</th>" +
                       "<th>设备未使用数量(台)</th>" +
                       "<th>视频时长总和(小时)</th>" +
                        "<th>视频大小(GB)</th>" +
                        "<th>48小时规范上传率</th>" +
                           "<th>设备使用率</th>" +
                         //---------------------------------------------------------
                              "<th>设备使用数量(台)</th>" +
                       "<th>设备未使用数量(台)</th>" +
                       "<th>视频时长总和(小时)</th>" +
                        "<th>视频大小(GB)</th>" +
                        "<th>48小时规范上传率</th>" +
                           "<th>设备使用率</th>" +
                         "</tr>"

                         $("#head2").html(html);
                         gongshi = "设备使用率=设备使用数÷设备配发数量";

                     break;
                 case "6"://辅警通
                   
                    //大队和中队

                         //table.column(8).visible(true);
                         //table.column(9).visible(true);
                         //table.column(10).visible(true);
                         //table.column(11).visible(true);
                         //table.column(12).visible(true);
                         //table.column(13).visible(true);
                         //table.column(14).visible(true);
                         //table.column(15).visible(true);
                         //table.column(16).visible(true);

                         var html = "<tr>" +
  "<th colspan='28' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                       "</tr>" +
                        "<tr>" +
                        "<th rowspan='2' style='text-align:center'>日期</th>" +
                      "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                       "<th rowspan='2' style='text-align:center'>辅警数</th>" +
                       "<th colspan='5' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='5' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='5' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='5' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='5' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                         "<tr>" +
                         //---------------------------------------------------------
                     "<th>违停采集(列)</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无违停设备</th>" +
                         //---------------------------------------------------------
                        "<th>违停采集(列)</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无违停设备</th>" +
                         //---------------------------------------------------------
                        "<th>违停采集(列)</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无违停设备</th>" +
                         //---------------------------------------------------------;
                     "<th>违停采集(列)</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无违停设备</th>" +
                         //---------------------------------------------------------
                       "<th>违停采集(列)</th>" +
                       "<th>人均处罚量</th>" +
                       "<th>查询量</th>" +
                        "<th>设备平均处罚量</th>" +
                        "<th>无违停设备</th>" +
                         "</tr>"

                         $("#head2").html(html);
                         gongshi = "人均处罚量=违停采集(例)÷辅警数，设备平均处罚量=违停采集(例)÷配发数";
      
                     break;
                 default:
                     break;
             }
             var td_count = $("#head2 tr:eq(2) th").length;
             switch (true) {
                 case td_count < 6:
                     $("#search-result-table2").width('100%');
                     $(".table2 .row").width('100%');
                     break;
                 case td_count < 16:
                     $("#search-result-table2").width(1800);
                     $(".table2 .row").width(1800);
                     break;
                 default:
                     $("#search-result-table2").width(3000);
                     $(".table2 .row").width(3000);
                     break;

             }
             $('.infodiv').html("<span>共 " + json.data.length + " 条记录</span><span>" + gongshi + "</span>");
             $('.daochu').html("<a class='buttons-excel'  href='../Handle/upload/" + json.title + "'><span>导 出</span></a>");

             //loadTatolData();
         })
        .on('draw.dt', function () {
            //给第一列编号
            //setTimeout(resizetbwidth, 50);

        })

        .DataTable({

            ajax: {
                url: "../Handle/Timesharing_ReportsTime.ashx",
                type: "POST",
                data: function () {
                    starttime = $(".start_form_datetime").val();
                    endtime = $(".end_form_datetime").val()
                    ssdd = $("#brigadeselect").val();
                    sszd = $("#squadronselect").val();
                    //search = $(".seach-box input").val();
                    type = $("#deviceselect").val();
                    searchtext = $(".search input").val();

                    return data = {
                        search: $(".search input").val(),
                        type: $("#deviceselect").val(),
                        ssdd: $("#brigadeselect").val(),
                        sszd: $("#squadronselect").val(),
                        //ssdd1: eachbrigadeselect(),
                        begintime: $(".start_form_datetime").val(),
                        endtime: $(".end_form_datetime").val(),
                        dates: datecompare($(".end_form_datetime").val(), $(".start_form_datetime").val()),
                        ssddtext: $("#brigadeselect").find("option:selected").text(),
                        sszdtext: $("#squadronselect").find("option:selected").text(),
                        //requesttype: "查询报表",
                        onlinevalue: chansu[0],
                        usedvalue: chansu[1]
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
            //             targets: 13,
            //            render: function (a, b, c, d) { var html = "<a  class=\'btn btn-sm btn-primary txzs-btn\' id='addedit' entityid='" + c.cloum12 + "'  >查看详情</a>"; return html; }
            //         }
            //],
            buttons: [],
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

            initComplete: function () { }
        });
}


function HeadcreateDataTable()
{



    //switch ($("#deviceselect").val()) {
    //    case "1"://车载视频
    //    case "3":   //拦截仪
    //        columns = [
    //               { "data": "cloum1", "orderable": false },
    //               { "data": "cloum2" },
    //               { "data": "cloum3" },
    //               { "data": "cloum4" },
    //               { "data": "cloum5" },

    //               { "data": "cloum6" },
    //               { "data": "cloum7" },
    //               { "data": "cloum8" },
    //               { "data": "cloum9" },
    //               { "data": "cloum10" },
    //               { "data": "cloum11" },
    //               { "data": "cloum12" },
    //               { "data": "cloum13" },
    //               { "data": "cloum14" },
    //               { "data": "cloum15" },
    //               { "data": "cloum16" },
    //               { "data": "cloum17" }
    //        ];


    //        break;
    //    case "5"://执法记录仪
    //        columns = [
    //                 { "data": "cloum1", "orderable": false },
    //                 { "data": "cloum2" },
    //                 { "data": "cloum3" },
    //                 { "data": "cloum4" },
    //                 { "data": "cloum5" },

    //                 { "data": "cloum6" },
    //                 { "data": "cloum7" },
    //                 { "data": "cloum8" },
    //                 { "data": "cloum9" },
    //                 { "data": "cloum10" },
    //                 { "data": "cloum11" },
    //                 { "data": "cloum12" },
    //                 { "data": "cloum13" },
    //                 { "data": "cloum14" },
    //                 { "data": "cloum15" },
    //                 { "data": "cloum16" },
    //                 { "data": "cloum17" },
    //                 { "data": "cloum18" },
    //                 { "data": "cloum19" },
    //                 { "data": "cloum20" },
    //                 { "data": "cloum21" },
    //                 { "data": "cloum22" },
    //                 { "data": "cloum23" },
    //                 { "data": "cloum24" },
    //                 { "data": "cloum25" },
    //                 { "data": "cloum26" },
    //                 { "data": "cloum27" },
    //                 { "data": "cloum28" },
    //                 { "data": "cloum29" },
    //                 { "data": "cloum30" },
    //                 { "data": "cloum31" },
    //                 { "data": "cloum32" }

    //        ];
    //        break;
    //}



}



$(document).on('click.bs.carousel.data-api', '.daochuall,.daochuall_time', function (e) {
    if ($('.end_form_datetime').val() < $('.start_form_datetime').val()) {
        $("#alertmodal").modal("show");
        return;
    };
    $("#loadingModal").modal("show");
    $('.creating').show();
    $('.createok').hide();
    var arry = time1.split(',');
    var chansu = arry[5].split('|');
    var data =
   {
       search: $(".search input").val(),
       type: $("#deviceselect").val(),
       ssdd: $("#brigadeselect").val(),
       sszd: $("#squadronselect").val(),
       ssdd1: eachbrigadeselect(),
       begintime: $(".start_form_datetime").val(),
       endtime: $(".end_form_datetime").val(),
       dates: datecompare($(".end_form_datetime").val(), $(".start_form_datetime").val()),
       ssddtext: $("#brigadeselect").find("option:selected").text(),
       sszdtext: $("#squadronselect").find("option:selected").text(),
       requesttype: this.innerText,
       onlinevalue: chansu[0],
       usedvalue: chansu[1]

   }
    $.ajax({
        type: "POST",
        url: (this.innerText == "一键导出")?"../Handle/exportAll_Timesharing_Reports.ashx":"../Handle/exportAll_Timesharing_ReportsTime.ashx",
        data: data,
        dataType: "json",
        success: function (data) {
            if (data.data) {
                $('.createok a').attr('href', '../Handle/upload/' + data.data);
                $('.createok').show();
                $('.creating').hide();
            }

        },
        error: function (msg) {
            console.debug("错误:ajax");
        }
    });

});
$(document).on('click.bs.carousel.data-api', '.createok a', function (e) {
    $("#loadingModal").modal("hide");
});


function createDataTable() {

    try
    {


        var arry = time1.split(',');

        var chansu = arry[5].split('|');

     
       var columns = [
                 { "data": "cloum1", "orderable": false },
                 { "data": "cloum2" },
                 { "data": "cloum3" },
                 { "data": "cloum4" },
                 { "data": "cloum5" },

                 { "data": "cloum6" },
                 { "data": "cloum7" },
                 { "data": "cloum8" },
                 { "data": "cloum9" },
                 { "data": "cloum10" },
                 { "data": "cloum11" },
                 { "data": "cloum12" },
                 { "data": "cloum13" },
                 { "data": "cloum14"},
                 { "data": "cloum15" },
                 { "data": "cloum16"},
                 { "data": "cloum17"},
                 { "data": "cloum18" },
                 { "data": "cloum19" },
                 { "data": "cloum20" },
                 { "data": "cloum21" },
                 { "data": "cloum22"},
                 { "data": "cloum23"},
                 { "data": "cloum24"},
                 { "data": "cloum25" },
                 { "data": "cloum26"},
                 { "data": "cloum27"},
                 { "data": "cloum28" },
                 { "data": "cloum29"},
                 { "data": "cloum30" },
                 { "data": "cloum31"},
                 { "data": "cloum32"}
                 //{ "data": null, "orderable": false }



       ];

        table = $('#search-result-table')
           .on('error.dt', function (e, settings, techNote, message) {

           })

                 .on('preXhr.dt', function (e, settings, data) {
                     $('.progresshz').show()
          
                     $('#search-result-table').hide();
                 })


             .on('xhr.dt', function (e, settings, json, xhr) {
  
         
                 $('.progresshz').hide();
                 $('#search-result-table').show();

                 var brigade="";
                 if ($("#brigadeselect").val() != "all") {
                     brigade = $("#brigadeselect").find("option:selected").text();

                 }


                 var squadron = "";
                 if ($("#squadronselect").val() != "all") {
                     squadron = $("#squadronselect").find("option:selected").text();

                 }
          
                 let gongshi;
                 switch ($("#deviceselect").val()) {
                     case "1"://车载视频
                     case "3":   //拦截仪
                     case "2":   //对讲机
                         //个人
                      var typename=  ($("#deviceselect").val()=="2")?"呼号":"设备编号" 
                         if ($("#brigadeselect").val() != "all" && $("#squadronselect").val() != "all") {
                            
      

                             //table.column(0).visible(true);
                             //table.column(1).visible(true);
                             //table.column(2).visible(true);
                             //table.column(3).visible(true);
                             //table.column(4).visible(true);
                             //table.column(5).visible(true);
                             //table.column(6).visible(true);
                             //table.column(7).visible(true);

                             //table.column(8).hide();
                             //table.column(9).visible(false);
                             //table.column(10).visible(false);
                             //table.column(11).visible(false);
                             //table.column(12).visible(false);
                             //table.column(13).visible(false);
                             //table.column(14).visible(false);
                             //table.column(15).visible(false);
                             //table.column(16).visible(false);

                             //table.column(17).visible(false);


                             //table.column(18).visible(false);
                             //table.column(19).visible(false);
                             //table.column(20).visible(false);
                             //table.column(21).visible(false);
                             //table.column(22).visible(false);
                             //table.column(23).visible(false);
                             //table.column(24).visible(false);
                             //table.column(25).visible(false);
                             //table.column(26).visible(false);

                             //table.column(27).visible(false);
                             //table.column(28).visible(false);
                             //table.column(29).visible(false);
                             //table.column(30).visible(false);
                             //table.column(31).visible(false);
       

                           
                             
                                         var html = "<tr>" +
                                                  "<th colspan='8' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() +""+ brigade +""+ squadron +""+ $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                             "</tr>" +
                              "<tr>" +
                                    "<th rowspan='2' style='text-align:center'>警员姓名</th>" +
                                  "<th rowspan='2' style='text-align:center'>警员编号</th>" +
                                  "<th rowspan='2' style='text-align:center'>" + typename + "</th>" +
                                   "<th style='text-align:center'>" + arry[0] + "</th>" +
                                   "<th  style='text-align:center'>" + arry[1] + "</th>" +
                                    "<th  style='text-align:center'>" + arry[2] + "</th>" +
                                    "<th  style='text-align:center'>" + arry[3] + "</th>" +
                                     "<th  style='text-align:center'>" + arry[4] + "</th>" +
                                     "</tr>" +
                                "<tr>" +
                                  "<th>使用时长(小时)</th>" +
                                   "<th>使用时长(小时)</th>" +
                                   "<th>使用时长(小时)</th>" +
                                   "<th>使用时长(小时)</th>" +
                                   "<th>使用时长(小时)</th>" +
                                   "</tr>"
                                         $("#head").html(html);
                                         

                         }

                         else {
           
                             try
                             {



                                 //table.column(0).visible(true);
                                 //table.column(1).visible(true);
                                 //table.column(2).visible(true);
                                 //table.column(3).visible(true);
                                 //table.column(4).visible(true);
                                 //table.column(5).visible(true);
                                 //table.column(6).visible(true);
                                 //table.column(7).visible(true);

                                 //table.column(8).show();
                                 //table.column(9).show();
                                 //table.column(10).show();
                                 //table.column(11).show();
                                 //table.column(12).show();
                                 //table.column(13).show();
                                 //table.column(14).show();
                                 //table.column(15).show();
                                 //table.column(16).show();
                                 //table.column(17).show();


                                 //table.column(17).visible(false);
                                 //table.column(18).visible(false);
                                 //table.column(19).visible(false);
                                 //table.column(20).visible(false);
                                 //table.column(21).visible(false);
                                 //table.column(22).visible(false);
                                 //table.column(23).visible(false);
                                 //table.column(24).visible(false);
                                 //table.column(25).visible(false);
                                 //table.column(26).visible(false);

                                 //table.column(27).visible(false);
                                 //table.column(28).visible(false);
                                 //table.column(29).visible(false);
                                 //table.column(30).visible(false);
                                 //table.column(31).visible(false);
        

                                 var html = "<tr>" +
  "<th colspan='17' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                               "</tr>" +
                                "<tr>" +
                                "<th rowspan='2' style='text-align:center'>部门</th>" +
                              "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                               "<th colspan='3' style='text-align:center'>" + arry[0] + "</th>" +
                               "<th colspan='3' style='text-align:center'>" + arry[1] + "</th>" +
                                "<th colspan='3' style='text-align:center'>" + arry[2] + "</th>" +
                                "<th colspan='3' style='text-align:center'>" + arry[3] + "</th>" +
                                 "<th colspan='3' style='text-align:center'>" + arry[4] + "</th>" +
                                 "</tr>" +
                                 "<tr>" +
                                 //---------------------------------------------------------
                               "<th>设备使用数量</th>" +
                               "<th>在线时长总和</th>" +
                               "<th>设备使用率</th>" +
                                 //---------------------------------------------------------
                                 "<th>设备使用数量</th>" +
                              "<th>在线时长总和</th>" +
                                 "<th>设备使用率</th>" +
                                 //---------------------------------------------------------
                                 "<th>设备使用数量</th>" +
                                 "<th>在线时长总和</th>" +
                                "<th>设备使用率</th>" +
                                 //---------------------------------------------------------;
                              "<th>设备使用数量</th>" +
                                 "<th>在线时长总和</th>" +
                                 "<th>设备使用率</th>" +
                                 //---------------------------------------------------------
                                "<th>设备使用数量</th>" +
                                 "<th>在线时长总和</th>" +
                                 "<th>设备使用率</th>" +
                                 "</tr>"

                                 $("#head").html(html);
                             }
                             catch (ex)
                             {
                                 alert(ex.message);

                             }

                         }
                         gongshi = "设备使用率=设备使用数÷设备配发数量";

                         break;

                

                     case "4"://警务通
                         //个人
                         if ($("#brigadeselect").val() != "all" && $("#squadronselect").val() != "all") {


                             //table.column(8).visible(false);
                             //table.column(9).visible(false);
                             //table.column(10).visible(false);
                             //table.column(11).visible(false);
                             //table.column(12).visible(false);
                             //table.column(13).visible(false);
                             //table.column(14).visible(false);
                             //table.column(15).visible(false);
                             //table.column(16).visible(false);

                             var html = "<tr>" +
  "<th colspan='13' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                 "</tr>" +
                  "<tr>" +
                        "<th rowspan='2' style='text-align:center'>警员姓名</th>" +
                        "<th rowspan='2' style='text-align:center'>警员编号</th>" +
                      "<th rowspan='2' style='text-align:center'>设备编号</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='2' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                    "<tr>" +
                      "<th>警务通处罚量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "<th>警务通处罚量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                      "<th>警务通处罚量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "<th>警务通处罚量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                      "<th>警务通处罚量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "</tr>"
                             $("#head").html(html);
                         }
                         else {//大队和中队
                             try
                             {
                    

                              


                        


                                 var html = "<tr>" +
  "<th colspan='28' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                               "</tr>" +
                                "<tr>" +
                                "<th rowspan='2' style='text-align:center'>部门</th>" +
                              "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                               "<th rowspan='2' style='text-align:center'>警员数</th>" +
                               "<th colspan='5' style='text-align:center'>" + arry[0] + "</th>" +
                               "<th colspan='5' style='text-align:center'>" + arry[1] + "</th>" +
                                "<th colspan='5' style='text-align:center'>" + arry[2] + "</th>" +
                                "<th colspan='5' style='text-align:center'>" + arry[3] + "</th>" +
                                 "<th colspan='5' style='text-align:center'>" + arry[4] + "</th>" +
                                 "</tr>" +
                                 "<tr>" +
                                 //---------------------------------------------------------
                               "<th>警务通处罚数</th>" +
                               "<th>人均处罚量</th>" +
                               "<th>查询量</th>" +
                                "<th>设备平均处罚量</th>" +
                                "<th>无处罚数的警务通(台)</th>" +
                                 //---------------------------------------------------------
                                     "<th>警务通处罚数</th>" +
                               "<th>人均处罚量</th>" +
                               "<th>查询量</th>" +
                                "<th>设备平均处罚量</th>" +
                                "<th>无处罚数的警务通(台)</th>" +
                                 //---------------------------------------------------------
                                     "<th>警务通处罚数</th>" +
                               "<th>人均处罚量</th>" +
                               "<th>查询量</th>" +
                                "<th>设备平均处罚量</th>" +
                                "<th>无处罚数的警务通(台)</th>" +
                                 //---------------------------------------------------------;
                                   "<th>警务通处罚数</th>" +
                               "<th>人均处罚量</th>" +
                               "<th>查询量</th>" +
                                "<th>设备平均处罚量</th>" +
                                "<th>无处罚数的警务通(台)</th>" +
                                 //---------------------------------------------------------
                               "<th>警务通处罚数</th>" +
                               "<th>人均处罚量</th>" +
                               "<th>查询量</th>" +
                                "<th>设备平均处罚量</th>" +
                                "<th>无处罚数的警务通(台)</th>" +
                                 "</tr>"

                                 $("#head").html(html);

                             }
                             catch (e) {
                                 alert(e.message);

                             }

                         }
               
                         gongshi = "人均处罚量=警务通处罚数÷警员数，设备平均处罚量=警务通处罚数÷配发数";

                         break;
                     case "5"://执法记录仪
            
                         //个人
                         if ($("#brigadeselect").val() != "all" && $("#squadronselect").val() != "all") {


                             //table.column(8).visible(false);
                             //table.column(9).visible(false);
                             //table.column(10).visible(false);
                             //table.column(11).visible(false);
                             //table.column(12).visible(false);
                             //table.column(13).visible(false);
                             //table.column(14).visible(false);
                             //table.column(15).visible(false);
                             //table.column(16).visible(false);

                             var html = "<tr>" +
  "<th colspan='13' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                 "</tr>" +
                  "<tr>" +
                        "<th rowspan='2' style='text-align:center'>警员姓名</th>" +
                        "<th rowspan='2' style='text-align:center'>警员编号</th>" +
                      "<th rowspan='2' style='text-align:center'>设备编号</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='2' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                    "<tr>" +
                      "<th>视频时长总和(小时)</th>" +
                       "<th>视频大小(GB)</th>" +
                    //-------------------------
                       "<th>视频时长总和(小时)</th>" +
                       "<th>视频大小(GB)</th>" +
                    //-------------------------
                       "<th>视频时长总和(小时)</th>" +
                       "<th>视频大小(GB)</th>" +
                    //-------------------------
                      "<th>视频时长总和(小时)</th>" +
                       "<th>视频大小(GB)</th>" +
                    //-------------------------
                       "<th>视频时长总和(小时)</th>" +
                       "<th>视频大小(GB)</th>" +
                    //-------------------------
                       "</tr>"
                             $("#head").html(html);
                         }
                         else {//大队和中队

                             //table.column(8).visible(true);
                             //table.column(9).visible(true);
                             //table.column(10).visible(true);
                             //table.column(11).visible(true);
                             //table.column(12).visible(true);
                             //table.column(13).visible(true);
                             //table.column(14).visible(true);
                             //table.column(15).visible(true);
                             //table.column(16).visible(true);

                             var html = "<tr>" +
  "<th colspan='32' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                           "</tr>" +
                            "<tr>" +
                            "<th rowspan='2' style='text-align:center'>部门</th>" +
                          "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
         
                           "<th colspan='6' style='text-align:center'>" + arry[0] + "</th>" +
                           "<th colspan='6' style='text-align:center'>" + arry[1] + "</th>" +
                            "<th colspan='6' style='text-align:center'>" + arry[2] + "</th>" +
                            "<th colspan='6' style='text-align:center'>" + arry[3] + "</th>" +
                             "<th colspan='6' style='text-align:center'>" + arry[4] + "</th>" +
                             "</tr>" +
                             "<tr>" +
                             //---------------------------------------------------------
                           "<th>设备使用数量(台)</th>" +
                           "<th>设备未使用数量(台)</th>" +
                           "<th>视频时长总和(小时)</th>" +
                            "<th>视频大小(GB)</th>" +
                            "<th>48小时规范上传率</th>" +
                               "<th>设备使用率</th>" +
                             //---------------------------------------------------------
                         "<th>设备使用数量(台)</th>" +
                           "<th>设备未使用数量(台)</th>" +
                           "<th>视频时长总和(小时)</th>" +
                            "<th>视频大小(GB)</th>" +
                            "<th>48小时规范上传率</th>" +
                               "<th>设备使用率</th>" +
                             //---------------------------------------------------------
                                 "<th>设备使用数量(台)</th>" +
                           "<th>设备未使用数量(台)</th>" +
                           "<th>视频时长总和(小时)</th>" +
                            "<th>视频大小(GB)</th>" +
                            "<th>48小时规范上传率</th>" +
                               "<th>设备使用率</th>" +
                             //---------------------------------------------------------;
                         "<th>设备使用数量(台)</th>" +
                           "<th>设备未使用数量(台)</th>" +
                           "<th>视频时长总和(小时)</th>" +
                            "<th>视频大小(GB)</th>" +
                            "<th>48小时规范上传率</th>" +
                               "<th>设备使用率</th>" +
                             //---------------------------------------------------------
                                  "<th>设备使用数量(台)</th>" +
                           "<th>设备未使用数量(台)</th>" +
                           "<th>视频时长总和(小时)</th>" +
                            "<th>视频大小(GB)</th>" +
                            "<th>48小时规范上传率</th>" +
                               "<th>设备使用率</th>" +
                             "</tr>"

                             $("#head").html(html);

                         }

                         gongshi = "设备使用率=设备使用数÷设备配发数量";

                         break;
    
                     case "6"://辅警通
                         //个人
                         if ($("#brigadeselect").val() != "all" && $("#squadronselect").val() != "all") {


                             //table.column(8).visible(false);
                             //table.column(9).visible(false);
                             //table.column(10).visible(false);
                             //table.column(11).visible(false);
                             //table.column(12).visible(false);
                             //table.column(13).visible(false);
                             //table.column(14).visible(false);
                             //table.column(15).visible(false);
                             //table.column(16).visible(false);

                             var html = "<tr>" +
  "<th colspan='13' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                 "</tr>" +
                  "<tr>" +
                        "<th rowspan='2' style='text-align:center'>警员姓名</th>" +
                        "<th rowspan='2' style='text-align:center'>警员编号</th>" +
                      "<th rowspan='2' style='text-align:center'>PDAID</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[0] + "</th>" +
                       "<th colspan='2' style='text-align:center'>" + arry[1] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[2] + "</th>" +
                        "<th colspan='2' style='text-align:center'>" + arry[3] + "</th>" +
                         "<th colspan='2' style='text-align:center'>" + arry[4] + "</th>" +
                         "</tr>" +
                    "<tr>" +
                      "<th>违停采集量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "<th>违停采集量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                      "<th>违停采集量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "<th>违停采集量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                      "<th>违停采集量</th>" +
                       "<th>查询量</th>" +
                    //-------------------------
                       "</tr>"
                             $("#head").html(html);
                         }
                         else {//大队和中队

                             //table.column(8).visible(true);
                             //table.column(9).visible(true);
                             //table.column(10).visible(true);
                             //table.column(11).visible(true);
                             //table.column(12).visible(true);
                             //table.column(13).visible(true);
                             //table.column(14).visible(true);
                             //table.column(15).visible(true);
                             //table.column(16).visible(true);

                             var html = "<tr>" +
  "<th colspan='28' style='text-align:center'>" + $(".start_form_datetime").val() + "-" + $(".end_form_datetime").val() + "" + brigade + "" + squadron + "" + $("#deviceselect").find("option:selected").text() + "分时段统计汇总</th>" +
                           "</tr>" +
                            "<tr>" +
                            "<th rowspan='2' style='text-align:center'>部门</th>" +
                          "<th rowspan='2' style='text-align:center'>设备配发数(台)</th>" +
                           "<th rowspan='2' style='text-align:center'>辅警数</th>" +
                           "<th colspan='5' style='text-align:center'>" + arry[0] + "</th>" +
                           "<th colspan='5' style='text-align:center'>" + arry[1] + "</th>" +
                            "<th colspan='5' style='text-align:center'>" + arry[2] + "</th>" +
                            "<th colspan='5' style='text-align:center'>" + arry[3] + "</th>" +
                             "<th colspan='5' style='text-align:center'>" + arry[4] + "</th>" +
                             "</tr>" +
                             "<tr>" +
                             //---------------------------------------------------------
                           "<th>违停采集(列)</th>" +
                           "<th>人均处罚量</th>" +
                           "<th>查询量</th>" +
                            "<th>设备平均处罚量</th>" +
                            "<th>无违停设备</th>" +
                             //---------------------------------------------------------
                           "<th>违停采集(列)</th>" +
                           "<th>人均处罚量</th>" +
                           "<th>查询量</th>" +
                            "<th>设备平均处罚量</th>" +
                            "<th>无违停设备</th>" +
                             //---------------------------------------------------------
                           "<th>违停采集(列)</th>" +
                           "<th>人均处罚量</th>" +
                           "<th>查询量</th>" +
                            "<th>设备平均处罚量</th>" +
                            "<th>无违停设备</th>" +
                             //---------------------------------------------------------;
                           "<th>违停采集(列)</th>" +
                           "<th>人均处罚量</th>" +
                           "<th>查询量</th>" +
                            "<th>设备平均处罚量</th>" +
                            "<th>无违停设备</th>" +
                             //---------------------------------------------------------
                           "<th>违停采集(列)</th>" +
                           "<th>人均处罚量</th>" +
                           "<th>查询量</th>" +
                            "<th>设备平均处罚量</th>" +
                            "<th>无违停设备</th>" +
                             "</tr>"

                             $("#head").html(html);

                         }

                         gongshi = "人均处罚量=违停采集(例)÷辅警数，设备平均处罚量=违停采集(例)÷配发数";

                         break;
                     default:
                         break;
                 }
                 var td_count = $("#head tr:eq(2) th").length;
                 switch (true)
                 {
                     case td_count<6:
                         $("#search-result-table").width('100%');
                         $("#table1 .row").width('100%');
                         break;
                     case td_count<16:
                         $("#search-result-table").width(1800);
                         $("#table1 .row").width(1800);
                         break;
                     default:
                         $("#search-result-table").width(3000);
                         $("#table1 .row").width(3000);
                         break;

                 }
                 //alert(json.title);
                 $('.infodiv').html("<span>共 " + json.data.length + " 条记录</span><span>" + gongshi + "</span>");
                 $('.daochu').html("<a class='buttons-excel'  href='../Handle/upload/" + json.title + "'><span>导 出</span></a>");
      
                 //loadTatolData();
             })
            .on('draw.dt', function () {
                //给第一列编号
                //setTimeout(resizetbwidth, 50);

            })

            .DataTable({

                ajax: {
                    url: "../Handle/Timesharing_Reports.ashx",
                    type: "POST",
                    data: function () {
                        starttime = $(".start_form_datetime").val();
                        endtime = $(".end_form_datetime").val()
                        ssdd = $("#brigadeselect").val();
                        sszd = $("#squadronselect").val();
                        //search = $(".seach-box input").val();
                        type = $("#deviceselect").val();
                        searchtext = $(".search input").val();
              
                        return data = {
                            search: $(".search input").val(),
                            type: $("#deviceselect").val(),
                            ssdd: $("#brigadeselect").val(),
                            sszd: $("#squadronselect").val(),
                            //ssdd1: eachbrigadeselect(),
                            begintime: $(".start_form_datetime").val(),
                            endtime: $(".end_form_datetime").val(),
                            dates: datecompare($(".end_form_datetime").val(), $(".start_form_datetime").val()),
                            ssddtext: $("#brigadeselect").find("option:selected").text(),
                            sszdtext: $("#squadronselect").find("option:selected").text(),
                            //requesttype: "查询报表",
                            onlinevalue: chansu[0],
                            usedvalue: chansu[1]
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
                //             targets: 13,
                //            render: function (a, b, c, d) { var html = "<a  class=\'btn btn-sm btn-primary txzs-btn\' id='addedit' entityid='" + c.cloum12 + "'  >查看详情</a>"; return html; }
                //         }
                //],
                buttons: [],
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

                initComplete: function () { }
            });

    }
    catch (ex)
    {
      alert(ex.message);
    }
}
