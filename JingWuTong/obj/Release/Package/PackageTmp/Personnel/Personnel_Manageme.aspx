<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Personnel_Manageme.aspx.cs" Inherits="JingWuTong.Personnel_Manageme.Personnel_Manageme" %>

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
<body id="body1" onload="headload(5)">
    <form id="form1" runat="server">

        <div id="header"></div>

    <div class="contentbody">
        <div class="leftbanner">

                <ul>
             
                    <li class="leftbanneractive" onclick="Load('Police_Management.aspx',this)">
                             <a href="Police_Management.aspx" target="frame"> 警员管理</a>
                   

                    </li>
          
         
                    <li onclick="Load('Role_Management.aspx',this)">
                   <a href="Role_Management.aspx" target="frame"> 角色管理</a>
                           
                   
                    </li>
          
                
                 
                </ul>
            </div>

              <div class="rightbody">
      <iframe    id="iframe" src="Police_Management.aspx"  name="frame"   scrolling="auto"  frameborder="0"  width="100%" height="100%" ></iframe> 
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