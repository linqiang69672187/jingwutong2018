/// <reference path="../Top.aspx" />

//$(function () {

//    $("#header").load('../Top.aspx', function () {
//        $("#header ul li:eq(1)").addClass("active");
//    });

//})


function headload(id) {


    $("#header").load('../Top.aspx', function () {
        $("#header ul li:eq("+id+")").addClass("active");
    });

}



$(document).on('mouseover.bs.carousel.data-api', '.leftbanner ul li', function (e) {
    var $doc = $(this);
    $doc.addClass("leftbannerover");
});
$(document).on('mouseout.bs.carousel.data-api', '.leftbanner ul li', function (e) {
    var $doc = $(this);
    $doc.removeClass("leftbannerover");
});
//$(document).on('click.bs.carousel.data-api', '.leftbanner ul li', function (e) {
//    var $doc = $(this);
//    $(".leftbanner ul li").removeClass("leftbanneractive");
//    $doc.addClass("leftbanneractive");
//    $(".rightbody").load('configs/entitymanage.html')
//});

function Load(url,mythis) {
    var $doc = $(mythis);
    $(".leftbanner ul li").removeClass("leftbanneractive");
    $doc.addClass("leftbanneractive");
    //$(".rightbody").load(url);


}

//$(document).on('click.bs.carousel.data-api', '#exprotIn', function (e) {
//    $("#daochumodal").modal("show");
//});
