<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Equipment_Management.aspx.cs" Inherits="JingWuTong.Equipment_Management" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>单警科技装备管理系统</title>
     <style type="text/css"> 


     </style> 

 


    <link href="../css/bootstrap.css" rel="stylesheet" />
    <link href="../PCSS/config.css" rel="stylesheet" />
     <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
     <script type="text/javascript" src="../js/bootstrap.js"></script>
    <script type="text/javascript" src="../js/config1.js"></script>


</head>
<body id="body1" onload="headload(4)">
        <div id="header"></div>
        <form id="form1" runat="server">
    <div class="contentbody">
        <div class="leftbanner">
                <ul >
             
                    <li class="leftbanneractive" onclick="Load('Equipment_Register.aspx',this)">
                        <a href="Equipment_Register.aspx" target="frame">设备登记</a>
                    </li>
                   
                    <li onclick="Load('MainTain_DeviceLog.aspx',this)">
                  <a href="MainTain_DeviceLog.aspx" target="frame">  维修统计</a>  
                
                    </li>
          
                    <li onclick="Load('Recycle_Device.aspx',this)">
                      <a href="Recycle_Device.aspx" target="frame"> 回收统计</a>

                         

                      
                    </li>
                    <li onclick="Load('Remind_Device.aspx',this)">
                       <a href="Remind_Device.aspx" target="frame">   设备提醒</a>
                        
             
                    </li>
                 
                </ul>
            </div>

    <%--        <div style="height:600px;width:80%;float:left; border-style:solid;border-color:rgb(2,34,150);border-width:medium;margin-top:1em">
            <iframe    id="iframe"  name="frame"   scrolling="no"  frameborder="0"  width="100%" height="100%" ></iframe>  
            </div>--%>
             <div class="rightbody">
      <iframe src="Equipment_Register.aspx"  id="iframe"  name="frame"   scrolling="auto"  frameborder="0"  width="100%" height="100%" ></iframe> 
             </div>



             </div>
    </form>
</body>
</html>

<script type="text/javascript">
    function lq_changeifr(name) {


        var ifr = document.getElementById("iframe")
        if (ifr) {
            ifr.src = name + ".aspx";
        }
    }




</script>