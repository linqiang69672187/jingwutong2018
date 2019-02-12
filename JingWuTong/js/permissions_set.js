var app = new Vue({
    el: '#configmodal',
    data: {
        role: {},
        pages: [],
        checkvalue: { rolename_wrong: false },
        modelname: '新增角色',
        selctedall: false
    },
    watch: {
        pages: {
            handler(oldpages, newpages) {
                this.selctedall = true;
                for (var page in this.pages) {
                    if (this.pages[page].ischecked == false) { this.selctedall = false; return; }
                    for (var btn in this.pages[page].buttons) {
                        if (this.pages[page].buttons[btn].ischecked == false) { this.selctedall = false; return; }
                    }
                    for (var childpage in this.pages[page].child_page) {
                        if (this.pages[page].child_page[childpage].ischecked == false) { this.selctedall = false; return; }
                        for (var childbtn in this.pages[page].child_page[childpage].buttons) {
                            if (this.pages[page].child_page[childpage].buttons[childbtn].ischecked == false) { this.selctedall = false; return; }
                        }
                    }

                }
            },
            deep:true
        }
    },
    mounted:
        function(){
          
        }
    ,
    methods: {
        createid:function(head,id){
            return head+id;
        },
        faclass: function (index) {
            // 隐藏/显示所谓的子行
            var iclass;
            if (this.pages[index].child_page.length == 0) return;
            return (this.pages[index].child_page[0].isshow) ? "fa fa-caret-down" : "fa fa-caret-up";
        },
        selectchild: function (index) {       
            for (var i = 0; i < this.pages[index].child_page.length; i++) {
                this.pages[index].child_page[i].isshow = !this.pages[index].child_page[i].isshow
            }
        },
        save: function (event) {
            var _this = this;

            if (_this.role.name.trim() == "") {
                _this.checkvalue.rolename_wrong = true;

               var timef= setTimeout(function () {
                    _this.checkvalue.rolename_wrong = false;
                    clearTimeout(timef);
                    timef = null;
                }, 3000);
                return;
            }
                $.ajax({
                    type: "POST",
                    url: "../Handle/permissions_set.ashx",
                    data: { 'requesttype': 'save', 'data': JSON.stringify(_this.pages), 'role': JSON.stringify(_this.role), 'roleid': GetQueryString('roleid'), 'roleid':'' },
                    dataType: "json",
                    success: function (data) {
                        window.location = 'Role_Management.aspx';
                    },
                    error: function (msg) {
                        console.debug("错误:ajax");
                    }
                });
        },
        selctedallfunc: function () {
            this.selctedall = !this.selctedall;
            for (var page in this.pages) {
                this.pages[page].ischecked = this.selctedall;
                for (var btn in this.pages[page].buttons) {
                    this.pages[page].buttons[btn].ischecked = this.selctedall;
                }
                for (var childpage in this.pages[page].child_page) {
                    this.pages[page].child_page[childpage].ischecked = this.selctedall;
                    for (var childbtn in this.pages[page].child_page[childpage].buttons) {
                        this.pages[page].child_page[childpage].buttons[childbtn].ischecked = this.selctedall;
                    }
                }

            }
        }
    },
    
      
});
function GetQueryString(name) {
    var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)");
    var r = window.location.search.substr(1).match(reg);
    if (r != null) return unescape(r[2]); return null;
}
function open_win() {
    $.ajax({
        type: "POST",
        url: "../Handle/permissions_set.ashx",
        data: { 'requesttype': 'add', 'roleid':null },
        dataType: "json",
        success: function (data) {
            app.role = data.role;
            app.pages = data.pages;
            app.modelname = "新增角色";
            $("#configmodal").modal("show");
        },
        error: function (msg) {
            console.debug("错误:ajax");
        }
    });

}

function open_win_eide(ID) {
    $.ajax({
        type: "POST",
        url: "../Handle/permissions_set.ashx",
        data: { 'requesttype': 'add', 'roleid': ID },
        dataType: "json",
        success: function (data) {
            app.role = data.role;
            app.pages = data.pages;
            app.modelname = "编辑角色";
            $("#configmodal").modal("show");
        },
        error: function (msg) {
            console.debug("错误:ajax");
        }
    });
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