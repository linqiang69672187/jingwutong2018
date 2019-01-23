var app = new Vue({
    el: '#setbody',
    data: {
        rolename: '管理员',
        creater:'林强',
        remark:'备注信息'
    }
});

$(function(){
    $('tr.parent').click(function(){ // 获取所谓的父行
      $(this).toggleClass("selected") .siblings('.child_'+this.id).toggle(); // 隐藏/显示所谓的子行
    });
});
