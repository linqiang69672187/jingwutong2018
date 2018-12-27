<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Timesharing_Reports.aspx.cs" Inherits="JingWuTong.Statistics.Timesharing_Reports" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
  <style type="text/css"> 

         html,body{margin:0px; height:100%;width:100%}

         select{
    /*Chrome和Firefox里面的边框是不一样的，重新覆盖一下*/
        border: solid 1px White;
        /*很关键：将默认的select选择框样式清除*/
        appearance:none;
        -moz-appearance:none;
        -webkit-appearance:none;
        /*在选择框的最右侧中间显示下拉箭头图片*/
         background: url("../image/DropDownList.png")  right center no-repeat;
        /*为下拉小箭头留出一点位置，避免被文字覆盖*/
        padding-right: 20px;

        color:White;
        background-color:rgb(6,18,55);
        border-radius:4px;

          }

      select::-ms-expand { display: none; }

          input {
          background-color:transparent;
        border: solid 1px White;
        border-radius:4px;
          }



          #FileUpload1 {
   background-color:transparent;
        border: solid 1px White;
        border-radius:4px;

          }


     </style> 
  

     <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <link href="../css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/bootstrap.js"></script>

    <script type="text/javascript" src="../js/My97DatePicker/WdatePicker.js"></script>

  
    <link href="../css/bootstrap-datetimepicker.min.css" rel="stylesheet" />

    <link href="../PCSS/Date.css" rel="stylesheet" />

    <script src="../comjs/bootstrap-datetimepicker.js"></script>

  <link href="../Static_Seed_Project/font-awesome/css/font-awesome.min.css" rel="stylesheet" />
   
    <link href="../css/report.css" rel="stylesheet" />
    <link href="../css/jquery.dataTables.css" rel="stylesheet" />

</head>
<body style="background-color:transparent">
    <form id="form1" runat="server">
           <div  class="tablediv">

     <div class="row"><label>| 报表</label></div>

            <div class="row">

                 <label>  设备类型:</label>
               
                       <asp:DropDownList ID="Dr_DeviceNmae" runat="server"  >
                
                      </asp:DropDownList>

               

              <label>  单位:</label>
                 
                  <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>
              

      
            </div>


                 <div class="row">
                 <label>统计时间</label>
                  
                      <div class="date" >  <asp:TextBox ID="T_strat" class=" start_form_datetime"  runat="server" Height="22px" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>
         <label>—</label>
         <div class="date">    <asp:TextBox ID="T_now" class=" end_form_datetime" runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

           

                <asp:TextBox ID="T_search" runat="server" placeholder="输入警员编号进行搜索"></asp:TextBox>
                
            
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" runat="server">查询</button>
                      &nbsp;
                   <button id="Button2" type="button" class="btn btn-primary btn-xs"  runat="server">日期汇总</button>
                      &nbsp;
                  <button id="Button3" type="button" class="btn btn-primary btn-xs"  runat="server">重置</button>


          </div>
         <div class="row">
         <div class="progress progress-striped active progresshz" style="width:80%;margin:0 auto;"><div class="floot-zhbar progress-bar progress-bar-success" role="progressbar" aria-valuenow="60" aria-valuemin="0" aria-valuemax="100" style="width: 100%;"></div> </div>
                <table id="search-result-table" style="width:100%;" border="1">
                    <thead id="head" runat="server"   >
                   

                      </thead>

                   <tbody>


                   </tbody>

                        </table>
           
 
             </div>
               </div>



    </form>
</body>
</html>
<script src="../js/report.js"></script>
<script src="../comjs/jquery.dataTables.js"></script>

