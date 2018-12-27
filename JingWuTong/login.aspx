<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="login.aspx.cs" Inherits="JingWuTong.login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
      <style type="text/css">  
        body, html, form
        {
            height: 100%;
            margin: 0px; 
            color: White;
            font-size: 12px;
        }
        #div1
        {
        position:absolute;
        z-index:0;
        height: 100%;
        width: 100%;
              top: 1px;
              left: 0px;
          }
        #tblogin
        {
            width: 451px;
            height: 206px;
         /*background: url('Images/login/loginin.png') no-repeat;*/

        }
        #Button1
        {
            cursor: pointer;
            background-color: transparent;
            background: url('Images/login/loginbt_bg.png') -138px 0px;
            width: 67px;
            height: 22px;
            color: #ffffff;
            border: 0;
        }
        #Button2
        {
            cursor: pointer;
            background-color: transparent;
            background: url('Images/login/loginbt_bg.png');
            border: 0;
            width: 67px;
            height: 22px;
            color: #ffffff;
        }
        #Button2:hover
        {
            cursor: pointer;
            background-position: -69px 0px;
        }
        #Button1:hover
        {
            cursor: pointer;
            background-position: -207px 0px;
        }
        #title1
        {
            width: 451px;
            height: 70px;
        }
        #Button4
        {
            background: url('Images/login/config.png');
            border: 0px;
            margin-right: 10px;
            margin-top: 7px;
            width: 40px;
            height: 23px;
        }
        #TextBox1, #TextBox2
        {
            border: 1px solid #2489ae;
        }
        .style1
        {
            width: 54px;
        }
    </style>


   <script type="text/javascript" src="js/jquery/jquery-3.3.1.js"></script>
    <link href="css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="js/bootstrap.js"></script>
</head>
<body onload="aload()">
<form id="form1" runat="server">
    <div id="div1"><img alt="" src="image/登录界面.jpg"  style="width:100%;height:100%;" /></div>
   <div  id="divlogin" style="position:absolute;z-index:1;background-color:rgb(3,14,61);opacity:0.8;">
    <table >
        <tr>
                  <td align="center">

               <table >
                    <tr>
                        <td id="title1" align="center"> <img src="image/登录界面头.png" /></td>
                    </tr>
                </table>




           <table id="tblogin">
                    <tr>
                        <td>
                            <table style="height: 202px; width: 100%;" border="0" cellspacing="0">
                                <tr>
                                 <%--   <td style="width: 170px; height: 100%">
                                    </td>--%>
                                    <td>
                                        <table width="100%" height="202px" border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                                <td>
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table id="myloginit" border="0" cellspacing="0" cellpadding="0"  style="width:400px;height:87px;">
                                                        
                                                        <tr>
                                                         
                                                            <td align="center">
                                                              
                                                                     <div class="col-sm-10 col-sm-push-2" >
                                                            <input type="text" class="form-control" id="firstname" runat="server" placeholder="请输入警号">
                                                                   </div>
                                                            </td>
                                                        </tr>
                                                    <tr>
                                                    <td>
                                                    &nbsp;
                                                    </td>
                                                    </tr>
                                                        <tr>
                                                   
                                                            <td align="center">
                                                                 <div class="col-sm-10 col-sm-push-2" >
                                                            <input type="password" class="form-control" id="firstpassword" runat="server" placeholder="请输入密码">
                                                               </div>
                                                            </td>
                                                        </tr>
                                             
                                                        <tr>
                                                        
                                                            <td align="left" >
                                                                 &nbsp; &nbsp; &nbsp;   &nbsp; &nbsp; &nbsp;
                                                                    &nbsp; &nbsp; &nbsp;   &nbsp; &nbsp; &nbsp;
                                                                  &nbsp; &nbsp; &nbsp;   &nbsp; &nbsp; &nbsp;
                                                             
                                                          
                                                                  <div class="col-sm-10 col-sm-push-2" >
                                                              <input type="checkbox" id="maintain" value="option1" runat="server" > 保持登录状态
                                                                 &nbsp; &nbsp; &nbsp;

                                                            <%--     <input type="checkbox" id="certificate" value="option1"  runat="server"> 数字证书登录--%>
                                                                      </div>
                                                            </td>
                                                        

                                                        </tr>
                                                       <tr>
                                                    <td>
                                                    &nbsp;
                                                    </td>
                                                    </tr>
                                                        <tr>
                                                            <td colspan="2" height="35px"  align="center">
                                                              <div class="col-sm-10 col-sm-push-2" >
                                                         
                                                                <button type="button" runat="server" class="btn btn-primary btn-block" onserverclick="Button2_Click" >登录</button>
                                                                  </div>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="42px" align="right" valign="bottom" >
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>










            </td>
        </tr>
            
    </table>
</div>
    </form>
</body>
</html>

<script type="text/javascript">



    $(function () {
        
            $.ajax({
                type: "POST",
                url: "../Handle/Login.ashx",
                data: { strWidth: document.body.clientWidth, strHeight: document.body.clientHeight },
                dataType: "json",
                success: function (data) {
                    if (data.r == "0") {
                  
                    }

                },
                error: function (msg) {
                    console.debug("错误:ajax");
                }
            });
        
        
        });



    var aload = function () {
        var divx = document.getElementById("divlogin");
        divx.style.left = (parseFloat(document.body.clientWidth) -500) / 2 + "px";
        divx.style.top = (parseFloat(document.body.clientHeight) -330) / 2+"px";
     
    }

    window.onresize = function () {
        var divx = document.getElementById("divlogin");
        divx.style.left = (parseFloat(document.body.clientWidth) - 500) / 2 + "px";
        divx.style.top = (parseFloat(document.body.clientHeight) - 330) / 2 + "px";
    };

</script>