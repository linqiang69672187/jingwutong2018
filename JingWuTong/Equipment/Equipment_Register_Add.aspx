<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Equipment_Register_Add.aspx.cs" Inherits="JingWuTong.Equipment.Equipment_Register_Add" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>

<style type="text/css"> 

    html,body{margin:0px; height:100%;width:100%}

    font {
        color:red;
        font-size:large;
        }

      input {
        background-color:transparent;
        border: solid 1px White;
        border-radius:4px;
          }

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


 </style>

   <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <link href="../css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/bootstrap.js"></script>
    <script type="text/javascript" src="../js/My97DatePicker/WdatePicker.js"></script>


     
    <link href="../css/bootstrap-datetimepicker.min.css" rel="stylesheet" />

    <link href="../PCSS/Date.css" rel="stylesheet" />

    <script src="../comjs/bootstrap-datetimepicker.js"></script>

  <link href="../Static_Seed_Project/font-awesome/css/font-awesome.min.css" rel="stylesheet" />


</head>
        

<body style="background:url('../Image/datamanagermentbg.jpg')">
    <form id="form1" runat="server">


       <div style="width:100%;height:100%;padding-left:2em">

    <div style="padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;设备登记   </label>

    </div>

           <div style="color:white;">

               <table style="border-collapse:separate; border-spacing:6px 10px;" >
              <caption>设备信息</caption>
                   <tr> 
                       <td align="right"><font>*</font>设备名称:</td>
                       <td align="left"><asp:DropDownList ID="Dr_DeviceNmae" runat="server"  AutoPostBack="True" OnSelectedIndexChanged="Dr_DeviceNmae_SelectedIndexChanged">
                
                      </asp:DropDownList></td>
                       <td align="right"><font>*</font>厂家:</td>
                       <td align="left">
                       <asp:TextBox ID="T_Manufacturer" runat="server" placeholder="请输入厂家名称"></asp:TextBox></td>
                   </tr>

                    <tr> 
                       <td align="right"><font>*</font>设备编号:</td>
                       <td align="left"><asp:TextBox ID="T_DevId" runat="server" placeholder="请输入设备编号"></asp:TextBox>
 <asp:RequiredFieldValidator ID="R_T_DevId" runat="server" ErrorMessage="设备编号不能为空" ControlToValidate="T_DevId" ForeColor="Red"></asp:RequiredFieldValidator>
                       </td>
                       <td align="right"><font>*</font>设备型号:</td>
                       <td align="left"><asp:DropDownList ID="Dr_SBXH" runat="server" AutoPostBack="True" OnSelectedIndexChanged="Dr_SBXH_SelectedIndexChanged"  >
                      </asp:DropDownList>
                         </td>
                   </tr>

                   <tr>
                       <td align="right"><font>*</font>规格:</td>
                    <td align="left"><asp:TextBox ID="T_SBGG" runat="server" placeholder="请输入设备规格"></asp:TextBox></td>
                       <td align="right"><font>*</font>项目名称:</td>
                       <td align="left"><asp:TextBox ID="T_ProjName" runat="server" placeholder="请输入项目名称" AutoPostBack="True"  OnTextChanged="T_ProjName_TextChanged"></asp:TextBox>

                            <asp:RequiredFieldValidator ID="R_T_ProjName" runat="server" ErrorMessage="项目名称不能为空" ControlToValidate="T_ProjName" ForeColor="Red"></asp:RequiredFieldValidator>

                       </td>
                   </tr>

                   <tr>
                       <td align="right"><font>*</font>项目编号:</td>
                       <td align="left"><asp:TextBox ID="T_ProjNum" runat="server" placeholder="请输入项目编号"></asp:TextBox></td>
                       <td align="right"><font>*</font>单价:</td>
                       <td align="left"><asp:TextBox ID="T_Price" runat="server" placeholder="请输入单价"></asp:TextBox></td>

                   </tr>
                   <tr>
                       <td align="right"><font>*</font>项目负责人:</td>
                       <td align="left"><asp:TextBox ID="T_XMFZR" runat="server" placeholder="请输入负责人姓名"></asp:TextBox></td>
                       <td align="right"><font>*</font>项目负责人电话:</td>
                       <td align="left"><asp:TextBox ID="T_XMFZRDH" runat="server" placeholder="请输入负责人电话"></asp:TextBox></td>
                   </tr>
                   <tr>
                       <td align="right"><font>*</font>采购时间:</td>
                       <td align="left">

        <div class="date" >  <asp:TextBox ID="T_CGSJ" runat="server" class="start_form_datetime" ></asp:TextBox><i class="fa fa-calendar"></i></div>
                       </td>
                       <td align="right"><font>*</font>保修期:</td>
                       <td align="left">

        <div class="date" >  <asp:TextBox ID="T_BXQ" runat="server" class="start_form_datetime" ></asp:TextBox><i class="fa fa-calendar"></i></div>


                       </td>
                   </tr>

                     <tr>
                       <td align="right"><font>*</font>报废期限:</td>
                       <td align="left">

       <div class="date" >  <asp:TextBox ID="T_BFQX" runat="server" class="start_form_datetime" ></asp:TextBox><i class="fa fa-calendar"></i></div>

                           
<asp:RequiredFieldValidator ID="R_T_BFQX" runat="server" ErrorMessage="报废期限不能为空" ControlToValidate="T_BFQX" ForeColor="Red"></asp:RequiredFieldValidator>

                       </td>
                       <td align="right"><font>*</font>设备状态:</td>
                       <td align="left"><asp:DropDownList ID="Dr_State" runat="server"></asp:DropDownList></td>
                   </tr>

                      <tr id="checktime" runat="server" visible="false">
                       <td align="right">下次检验时间:</td>
                       <td align="left">
       <div class="date" >  <asp:TextBox ID="T_XCJYSJ" runat="server" class="start_form_datetime" ></asp:TextBox><i class="fa fa-calendar"></i></div>

                       </td>
                       <td align="right">本次检验时间:</td>
                       <td align="left">
       <div class="date" >  <asp:TextBox ID="T_BCJYSJ" runat="server" class="start_form_datetime" ></asp:TextBox><i class="fa fa-calendar"></i></div>
                       </td>
                   </tr>

                   <tr>
                   <td align="right">配发状态</td>
                  <td colspan="3" align="left"><asp:DropDownList ID="Dr_AllocateState" runat="server" AutoPostBack="True"></asp:DropDownList></td>

                   </tr>
                   </table>


               <table style="border-collapse:separate; border-spacing:6px 10px;">

                    <caption>设备信息</caption>
                    <tr >
                       <td align="right">使用人警号:</td>
                       <td align="left"><asp:TextBox ID="T_JYBH" runat="server" placeholder="请输入使用警号" AutoPostBack="True" OnTextChanged="T_JYBH_TextChanged"></asp:TextBox></td>
                       <td align="right">使用人电话:</td>
                       <td align="left"><asp:TextBox ID="T_SJ" runat="server" placeholder="请输入使用人电话"></asp:TextBox></td>
           
                   </tr>

                <tr>
                       <td align="right">使用人姓名:</td>
                       <td align="left" colspan="3"><asp:TextBox ID="T_XM" runat="server" placeholder="请输入使用人姓名"></asp:TextBox></td>


                 
                   </tr>
                   

                   <tr>
                             <td align="right"><font>*</font>使用人单位:</td>
                       <td align="left" colspan="2"> <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged" ></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>



                       </td>

                    <td  align="center">
                             <button id="Button3" type="button" class="btn btn-primary btn-xs" onserverclick="Button3_Click" runat="server" visible="false">解除绑定</button>

                       </td>

                   </tr>
              

         <%--          <tr>

                       <td colspan="4" align="center">
                             <button id="Button3" type="button" class="btn btn-primary btn-xs" onserverclick="Button3_Click" runat="server" visible="false">解除绑定</button>

                       </td>


                   </tr>--%>




                   <tr> 
                       <td colspan="4" align="center">
                            <button id="Button2" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">保存</button>

                       </td>

                   </tr>
               </table>

           </div>
    
    </div>





    </form>
</body>
</html>

<script src="../js/Date.js"></script>
<script type="text/javascript">
    



</script>