<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Top.aspx.cs" Inherits="JingWuTong.Top" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>单警科技装备管理系统</title>
    <link href="../css/top.css" rel="stylesheet" />

    <link href="../Static_Seed_Project/font-awesome/css/font-awesome.min.css" rel="stylesheet" />

    <%--<script src="js/jquery/jquery-3.3.1.min.js"></script>--%>
</head>
<body>

 

    <form id="form1" runat="server">
      
      <div class="row">
   
        <div>
  
            <div> <img src="../PImage/bannar-logo.png" /></div>
            <div>
                <ul>
                <li vspglabel="首页"><a href="../index.html"><span><i class="fa fa-home" aria-hidden="true"></i>首页</span></a></li>
                <li vspglabel="设备查看"><a href="../Equipment/Equipment_Examine.aspx"><span><i class="fa fa-desktop" aria-hidden="true"></i>设备查看</span></a></li>
                <li vspglabel="实时状况"><a href="../map.html"><span><i class="fa fa-map-o" aria-hidden="true"></i>实时状况</span></a></li>
                <li vspglabel="数据统计"><a href="../Statistics/Statistics_Management.aspx"><span><i class="fa fa-bar-chart" aria-hidden="true"></i>数据统计</span></a></li>
                <li vspglabel="设备管理"><a href="../Equipment/Equipment_Management.aspx"><span><i class="fa fa-laptop" aria-hidden="true"></i>设备管理</span></a></li>
                <li vspglabel="设备管理"><a href="../Personnel/Personnel_Manageme.aspx"><span><i class="fa fa-user" aria-hidden="true"></i>人员管理</span></a></li>
                <li vspglabel="系统设置"><a  href="../SystemSetup/SystemSetup.aspx"><span><i class="fa fa-cogs" aria-hidden="true"></i>系统设置</span></a></li>
               
                </ul>
            </div>
        </div>
        <div style="padding-top: 10px;"><i class="fa fa-user-circle"></i><span id="LoginName" runat="server"> Admin</span> <i class="fa fa-dot-circle-o"></i><span id="out" style="cursor: pointer;" > 退出</span><span class="bannertime">2018 04-09 18:27</span> </div>
    </div>
    <!-- Mainly scripts -->
<%--  <asp:Button ID="B_Out" runat="server" Text="Button" OnClick="B_Out_Click" Width="0px" Height="0px" BackColor="Transparent"  BorderColor="Transparent" />--%>

    </form>
</body>
</html>




<script type="text/javascript">


    $(function () {
        

        $("#out").click(function () {
      

            $.ajax({
                type: "POST",
                url: "../Handle/OperationLog.ashx",
                data: "",
                dataType: "json",
                success: function (data) {
                    if (data.r == "0") {
                  
                    }

                },
                error: function (msg) {
                    console.debug("错误:ajax");
                }
            });
        
            window.location.href = "../login.aspx";

            //$("#B_Out").click();
        });


    });


    setInterval(function () {
        var date = new Date();
        var year = date.getFullYear();
        var month = Appendzero(date.getMonth() + 1);
        var day = Appendzero(date.getDate());
        var hour = Appendzero(date.getHours());
        var min = Appendzero(date.getMinutes());
        var sencond = Appendzero(date.getSeconds());

        $(".bannertime").text(year + "-" + month + "-" + day + " " + hour + ":" + min + ":" + sencond);
    }, 50);

    function Appendzero(obj) {
        if (obj < 10) return "0" + "" + obj;
        else return obj;
    }

    $.ajax({
        type: "POST",
        url: "../Handle/permissions_load.ashx",
        data: { 'type': 'page' },
        dataType: "json",
        success: function (data) {
            var data = data.data;
            for (var i = 0; i < data.length; ++i) {
                $("[vspglabel='" + data[i]["name"] + "']").hide();
            }
        },
        error: function (msg) {
            console.debug("错误:ajax");
        }
    });


</script>