<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="JingWuTong.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
     <meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>单警科技装备管理系统</title>

      <style type="text/css"> 

         html,body{margin:0px; height:100%;width:100%}

          #Dr_DeviceNmae {


    /*Chrome和Firefox里面的边框是不一样的，重新覆盖一下*/
        border: solid 1px #000;
        /*很关键：将默认的select选择框样式清除*/
        appearance:none;
        -moz-appearance:none;
        -webkit-appearance:none;
        /*在选择框的最右侧中间显示下拉箭头图片*/
         background: url("image/DropDownList.png")  right center no-repeat;
        /*为下拉小箭头留出一点位置，避免被文字覆盖*/
        padding-right: 20px;

          }

      select::-ms-expand { display: none; }



     </style> 
 
     <link href="css/bootstrap.css" rel="stylesheet" />


   


</head>

<body style="background-image:url(../image/设备查看背景.jpg); background-repeat:no-repeat;background-size:100% 100%;margin:0px; background-attachment:fixed; " >
    <form id="form1" runat="server">
    <div>
     
    <select name="Dr_DeviceNmae" id="Dr_DeviceNmae" class="dropdown-toggle" style="color:White;background-color:rgb(6,18,55);">
	<option selected="selected" value="1">车载视频</option>
	<option value="2">对讲机</option>
	<option value="3">拦截仪</option>
	<option value="4">警务通</option>
	<option value="5">执法记录议</option>
	<option value="6">辅警通</option>
	<option value="7">测速仪</option>
	<option value="8">酒精测试仪</option>

</select>

    </div>

  
  




    </form>
</body>
</html>
