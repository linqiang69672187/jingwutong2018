<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Statistics_Management.aspx.cs" Inherits="JingWuTong.Statistics.Statistics_Management" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
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
<body id="body1" onload="headload(3)">
        <div id="header"></div>
        <form id="form1" runat="server">
    <div class="contentbody">
        <div class="leftbanner">
                <ul >
                
                    <li class="leftbanneractive" onclick="Load('dataManagement.html',this)">
                        <a href="../dataManagement.html" target="frame">报表统计</a><script src="../js/entitymanage.js"></script>
                    </li>
          
                    <li onclick="Load('Timesharing_Reports.html',this)">
                  <a href="Timesharing_Reports.html" target="frame">  分时段报表统计</a>  
                
                    </li>
          
                </ul>
            </div>

    <%--        <div style="height:600px;width:80%;float:left; border-style:solid;border-color:rgb(2,34,150);border-width:medium;margin-top:1em">
            <iframe    id="iframe"  name="frame"   scrolling="no"  frameborder="0"  width="100%" height="100%" ></iframe>  
            </div>--%>
             <div class="rightbody">
      <iframe src="../dataManagement.html"  id="iframe"  name="frame"   scrolling="auto"  frameborder="0"  width="100%" height="100%" ></iframe> 
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