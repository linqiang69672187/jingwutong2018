<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SystemSetup.aspx.cs" Inherits="JingWuTong.SystemSetup.SystemSetup" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
     <style type="text/css"> 

     </style> 

 
     <link href="../css/bootstrap.css" rel="stylesheet" />
        <link href="../PCSS/config.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <script type="text/javascript" src="../js/bootstrap.js"></script>
    <script type="text/javascript" src="../js/config1.js"></script>

</head>
<body id="body1" onload="headload(6)">
    <form id="form1" runat="server">
               <div id="header"></div>

    <div class="contentbody">
        <div class="leftbanner">

                <ul>
                   
             
                    <li class="leftbanneractive" onclick="Load('#',this)">
                         <a href="../configs/indexconfig.html" target="frame"> 首页参数设置</a>
                          
                 
                    </li>

                    <li onclick="Load('Department_Management.aspx',this)">
                    <a href="Department_Management.aspx" target="frame"> 部门管理</a>
                       
                    </li>
            

                   <li onclick="Load('Equipment_Log.aspx',this)">
                   <a href="Equipment_Log.aspx" target="frame"> 设备日志</a>
                        
               
                    </li>

                   <li onclick="Load('System_Log.aspx',this)">
                     <a href="System_Log.aspx" target="frame"> 系统日志</a>
                      
                    </li>

                    <li onclick="Load('AlarmGate.aspx',this)">
                     <a href="AlarmGate.aspx" target="frame"> 预警设置</a>
                      
                    </li>


                 
                </ul>
            </div>

                 <div class="rightbody">
      <iframe src="../configs/indexconfig.html" id="iframe"  name="frame"   scrolling="auto"  frameborder="0"  width="100%" height="100%" ></iframe> 
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