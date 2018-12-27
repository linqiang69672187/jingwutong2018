<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Equipment_Examine.aspx.cs" Inherits="JingWuTong.Equipment_Examine" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

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
            width:100%;
            margin: 0px; 
            color: White;
            font-size: 12px;
        }

             /*body {background-image: url(image/设备查看背景.jpg);}*/

          #bn {

      position:absolute;
       z-index:2;
      /*right: 730px;
      top:500px*/
     
          }

          #trian {
      /*position:absolute;

      /*right: 900px;*/
       left: 50%;
            /*top: 50%;*/


          }

          /*图例样式*/
          .dl-horiz dt {
       float: left;
       width: 30px;
      overflow: hidden;
      clear: left;
      text-align: right;
      text-overflow: ellipsis;
       white-space: nowrap;
     /*padding: 10px 5px 10px 8px;*/
     margin : 10px 0px 0px 5px;


          }

          .dl-horiz dd {
            margin : 10px 0px 0px 5px;
                  width: 100px;


          }


              /*#div1
        {
        position:absolute;
        z-index:0;
        height: 100%;
        width: 100%;
              
          }*/

          #divlogin {
              position:absolute;
              z-index:1;

                 /*height: 100%;
        width: 100%;*/

          }

          .ctable {
             /*// background-color:rgb(3,14,61);
              //opacity:0.8;*/

          }

          .div2  {
               background-color:rgb(3,14,61);
              opacity:0.6;
              padding: 10px 10px 10px 10px;
              margin : 10px 40px 10px 40px;
              cursor:pointer;
             
          }
         



            .div2:hover{
    background-color:rgb(14,72,167);
}


          #triangle {
              width: 0;
              height: 0;
              border-top: 40px solid white;
              border-left:30px solid transparent;
              border-right:30px solid transparent;
              position: relative;
              z-index: 1;
              /*display:none;*/
                left: 50%;
           
          }

          #trianglebox {
        position: relative;
        width: 600px;
        height: 90px;
        border: 2px solid white;
        border-radius: 8px;
        background-color:white;
              display:none;

                left: 50%;


          }


           #trianglebox:before {
        position: absolute;
        content: "";
        top: -30px;
        left: 200px;
        border-left: 30px solid transparent;
        border-right: 30px solid transparent;
        border-bottom: 30px solid white;
    }


          #t_eq td {
              color:black;
               width:200px;
        height:auto;
          }


          #t_eq {


          }

      #hidebg { position:absolute;left:0px;top:0px; 
      background-color:#000; 
      width:100%;  /*宽度设置为100%，这样才能使隐藏背景层覆盖原页面*/ 
      filter:alpha(opacity=60);  /*设置透明度为60%*/ 
      opacity:0.6;  /*非IE浏览器下设置透明度为60%*/ 
      display:none; /* http://www.jb51.net */ 
      z-Index:2;}

     </style> 

        <script type="text/javascript">
            function show()  //显示隐藏层和弹出层 
            {
                var hideobj = document.getElementById("hidebg");
                hidebg.style.display = "block";  //显示隐藏层 
                hidebg.style.height = document.body.clientHeight + "px";  //设置隐藏层的高度为当前页面高度 

                document.getElementById("btn_all").disabled = "disabled";

            }

       

            function C_div(mythis) {

                $("#"+mythis).css("backgroundColor", "rgb(14,72,167)");
                //document.getElementById(mythis).style.backgroundColor = "red";

            }


            //C_div("one");

        </script>
        <link href="../css/bootstrap.css" rel="stylesheet" />
        <link href="../PCSS/config.css" rel="stylesheet" />
      <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <script type="text/javascript" src="../js/bootstrap.js"></script>

       <script type="text/javascript" src="../js/config1.js"></script>

</head>

<body onload="headload(1);aload()">


      <div id="header"></div>

    <form id="form1" runat="server">


      <asp:TextBox ID="HiddenState" runat="server" Width="0px" Height="0px" BorderColor="transparent" ></asp:TextBox>
      <asp:TextBox ID="HiddenNState" runat="server" Width="0px" Height="0px" BorderColor="transparent"  ></asp:TextBox>


        <asp:Button ID="first" runat="server" Text="Button" OnClick="Button1_Click" style="display: none;" />
         <asp:Button ID="second" runat="server" Text="Button" OnClick="Button2_Click" style="display: none;" />
         <asp:Button ID="third" runat="server" Text="Button" OnClick="Button3_Click" style="display: none;" />
         <asp:Button ID="fourth" runat="server" Text="Button" OnClick="Button4_Click" style="display: none;" />


        <div id="hidebg"></div><%--不可使用--%>

            <div class="contentbody">

       <div id="headline" style="height:60px;width:100%; text-align: center; font-size:50px;font-family:'微软雅黑';margin-bottom:70px">
             <img alt="" src="../image/线条.png"  style="visibility:hidden"  / > <img alt="" src="../image/设备查看标题4.png" style="display:none"  />
                   <span>台州市公安局交通警察局</span>
           <img alt="" src="../image/线条.png"  style="visibility:hidden"  />

     
       </div>
          <div id="trian" style="display:none">
                       <div id="triangle"> </div>

                 <div id="trianglebox" runat="server"></div>
           </div>
      
<div>

    <div style="width:18%;float:left">

        <div class="div2" id="one"  onclick="document.getElementById('first').click();" runat="server" style="visibility:hidden">
           <table class="ctable" align="center"  style=" width:100px  ">

              <tr>  <td align="center" style="font-size:150%;font-weight:bold">一大队</td> </tr>
              <tr>  <td align="center" runat="server" id="T_first"></td>  </tr>

          </table>
        </div>

          <div class="div2" id="three" onclick="document.getElementById('third').click();" runat="server"  style="visibility:hidden">
           <table class="ctable" align="center"  style=" width:100px  ">
         <tr>  <td align="center" style="font-size:150%;font-weight:bold">三大队</td>  </tr>
        <tr>  <td align="center" runat="server" id="T_Three"></td>  </tr>

         </table>
        </div>

    </div>


     
          <div style="width:60%;float:left">


            <%--        <div id="bn">
                          <button type="button" id="btn_all"  class="btn btn-primary btn-xs" runat="server" onserverclick="BShow_Click">全部显示</button>
                              
                           </div>--%>

        <asp:Chart ID="Chart1" runat="server"  BackColor="Transparent" OnClick="Chart1_Click" OnCustomizeLegend="Chart1_CustomizeLegend" CssClass="chartWidth">
                                  <Series>

                                  </Series>

                                  <ChartAreas>
               <asp:ChartArea Name="ChartArea1" BorderColor="Transparent" BackColor="Transparent">
                                  
                </asp:ChartArea>
                                  </ChartAreas>

                 <Legends>
             
        <asp:Legend BackColor="Transparent" BorderColor="Transparent" Name="Legend1" ForeColor="White" TitleForeColor="Transparent" Font="Microsoft Sans Serif, 10pt" IsTextAutoFit="False" ItemColumnSpacing="50" BorderWidth="0" LegendStyle="Column">

            <CustomItems>
                <asp:LegendItem BorderColor="Transparent" BorderWidth="0" PostBackValue="true" >
                


                    <Cells>
                        <asp:LegendCell Name="Cell1" BackColor="17,107, 242"  Font="8pt" Text="全部" CellType="Text" >

                            
                            <Margins  Bottom="15" Top="15"></Margins>
                        </asp:LegendCell>

                    </Cells>

                </asp:LegendItem>

            </CustomItems>


        </asp:Legend>


                                  </Legends>


                              </asp:Chart>


              </div>  


        
    <div style="width:18%;float:right">
        <div class="div2" id="two" onclick="document.getElementById('second').click();" runat="server"  style="visibility:hidden">
            <table class="ctable" align="center" style=" width:100px ">
                           <tr> <td align="center" style="font-size:150%;font-weight:bold">二大队</td> 
                            </tr>
                           <tr> <td align="center" runat="server" id="T_Second"></td>  </tr>
                          </table>
        </div>


           <div class="div2" id="four" onclick="document.getElementById('fourth').click();" runat="server"  style="visibility:hidden">
           <table class="ctable" align="center" style=" width:100px ">
                           <tr> <td align="center" style="font-size:150%;font-weight:bold">四大队</td> </tr>
                           <tr> <td align="center" runat="server" id="T_Four"></td>  </tr>
                          </table>
        </div>

    </div>

</div>
  
</div>


    </form>
</body>
</html>

<script type="text/javascript">
    




   

    $(function () {


      
        $("#triangle").click(function () {


            $("#triangle").hide();
            $("#trianglebox").show();

        })


        $("#trianglebox").click(function () {

            $("#triangle").show();
            $("#trianglebox").hide();


        })



    



    })






    document.getElementById("HiddenState").value = document.body.clientWidth;

    document.getElementById("HiddenNState").value = document.body.clientHeight;





    var aload = function () {
        //var divx = document.getElementById("bn");

        //divx.style.left = (parseFloat(document.body.clientWidth)) *12/16 + "px";
        //divx.style.top = (parseFloat(document.body.clientHeight)) *2/ 6 + "px";

<%--       var a="<%=Getstrdx()%>";--%>

        //var Chart = document.getElementById("Chart1");
        //Chart.style.width = (parseFloat(document.body.clientWidth)) + "px";

        //Chart.style.height = (parseFloat(document.body.clientHeight)) + "px";

        var d_header = document.getElementsByTagName("img");

        d_header[0].style.width=(parseFloat(document.body.clientWidth))/4+"px"
        d_header[1].style.width = (parseFloat(document.body.clientWidth)) / 3 + "px"
        d_header[2].style.width = (parseFloat(document.body.clientWidth))/4 + "px"

    }

    window.onresize = function () {
        //var divx = document.getElementById("bn");
        //divx.style.left = (parseFloat(document.body.clientWidth)) * 10/13 + "px";
        //divx.style.top = (parseFloat(document.body.clientHeight)) * 2 /5 + "px";
     <%--   var a = "<%=Getstrdx()%>";--%>

        var Chart = document.getElementById("Chart1");


        Chart.style.width = (parseFloat(document.body.clientWidth)) * 2/3 + "px";
        Chart.style.height = (parseFloat(document.body.clientHeight))/2 + "px";


        var d_header = document.getElementsByTagName("img");

        d_header[0].style.width = (parseFloat(document.body.clientWidth)) / 4 + "px"
        d_header[1].style.width = (parseFloat(document.body.clientWidth)) / 3 + "px"
        d_header[2].style.width = (parseFloat(document.body.clientWidth)) / 4 + "px"

        //alert($("#headline").outerWidth(true));

        //alert(document.body.clientWidth);



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


    };










</script>