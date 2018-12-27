


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
    format: 'yyyy/mm/dd hh:ii:ss ',
    autoclose: true,
    todayBtn: true,
    minView: 0
});

//function startdatetimedefalute() {
//    var curDate = new Date();
//    var preDate = new Date(curDate.getTime() - 24 * 60 * 60 * 1000); //前一天
//    var beforepreDate = new Date(curDate.getTime() - 48 * 60 * 60 * 1000); //前一天
//    $('.start_form_datetime').val(transferDate(beforepreDate));
//    $('.end_form_datetime').val(transferDate(preDate));
//}


//startdatetimedefalute();