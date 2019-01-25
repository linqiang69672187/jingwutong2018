var app = new Vue({
    el: '#setbody',
    data: {
        role: {},
        pages: []
    },
    mounted:
        function(){
            $.ajax({
                type: "POST",
                url: "../Handle/permissions_set.ashx",
                data: { 'requesttype': 'add', 'roleid': GetQueryString('roleid') },
                dataType: "json",
                success: function (data) {
                    app.role = data.role;
                    app.pages = data.pages;
                },
                error: function (msg) {
                    console.debug("错误:ajax");
                }
            });
        }
    ,
    methods: {
        createid:function(head,id){
            return head+id;
        },
        createclass:function(head,id){
            return head+id+" childrow";
        },
        selectchild: function (event) {
            $(event.currentTarget).toggleClass("selected").siblings('.child_' + event.currentTarget.id).toggle(); // 隐藏/显示所谓的子行
        },
        save: function (event) {
            var _this =this;
                $.ajax({
                    type: "POST",
                    url: "../Handle/permissions_set.ashx",
                    data: { 'requesttype': 'save', 'data': JSON.stringify(_this.pages), 'role': JSON.stringify(_this.role), 'roleid': GetQueryString('roleid'), 'roleid':'' },
                    dataType: "json",
                    success: function (data) {
                        alert(data)
                    },
                    error: function (msg) {
                        console.debug("错误:ajax");
                    }
                });
        }
    },
    
      
});
function GetQueryString(name) {
    var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)");
    var r = window.location.search.substr(1).match(reg);
    if (r != null) return unescape(r[2]); return null;
}
//$('tr.parent td').click(function () { // 获取所谓的父行
//    $(this).toggleClass("selected").siblings('.child_' + this.id).toggle(); // 隐藏/显示所谓的子行
//});
//    { name: '首页',checked:true },
//    { name: '设备查看',checked:false},
//    { name: '实时状况',checked:false,buttons:[
//        {name:'查看',checked:true},
//        {name:'历史轨迹',checked:false},
//    ] },
//    { name: '数据统计',checked:false,child_page:[
//        {name:'报表统计',checked:true,buttons:[
//            {name:'查询',checked:true},
//            {name:'重置',checked:false},
//            {name:'导出',checked:false},
//            {name:'一键导出',checked:true},
//            {name:'参数配置',checked:false},
//            {name:'数据项刷选',checked:false},  
//        ]},
//        {name:'分时段报表统计',buttons:[
//            {name:'查询',checked:true},
//            {name:'时间分类',checked:true},
//            {name:'重置',checked:true},
//            {name:'导出',checked:true},
//            {name:'一键导出',checked:false},
//            {name:'一键导出时间统计',checked:false},                         
//        ]}
//    ] },
//    { name: '设备管理',checked:false ,child_page:[
//        {name:'设备登记'},{name:'维修统计'},{name:'回收统计'},{name:'设备提醒'}
//    ]},
//    { name: '人员管理',checked:false,child_page:[{name:'警员管理'},{name:'角色管理'}] },
//    { name: '系统设置' ,checked:true,child_page:[
//      {name:'首页参数设置'},{name:'部门管理'},{name:'系统日志'},{name:'预警设置'}
//    ]
//},