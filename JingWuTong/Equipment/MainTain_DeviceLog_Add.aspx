<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MainTain_DeviceLog_Add.aspx.cs" Inherits="JingWuTong.Maintain.MainTain_DeviceLog_Add" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
     <title>单警科技装备管理系统</title>

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

    textarea {

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


    .auto-style1 {
        height: 27px;
    }


 </style>

   <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <link href="../css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/bootstrap.js"></script>
    <script type="text/javascript" src="../js/My97DatePicker/WdatePicker.js"></script>


</head>
<body style="background:url('../Image/datamanagermentbg.jpg')">
    <form id="form1" runat="server">
       <div style="width:100%;height:100%;padding-left:2em">

    <div style="padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;新增维修设备   </label>

    </div>

           <div style="color:white;">

               <table style="border-collapse:separate; border-spacing:6px 10px;" >
      
                   <tr> 
                       <td align="right"><font>*</font>设备名称:</td>
                       <td align="left""><asp:DropDownList ID="Dr_DeviceNmae" runat="server"  >
                
                      </asp:DropDownList></td>
                       <td align="right">设备编号:</td>
                       <td align="left">
                     <asp:TextBox ID="T_DevId" runat="server" placeholder="请输入设备编号"></asp:TextBox></td>
                   </tr>

                    <tr> 
                       <td align="right">设备状态:</td>
                       <td align="left"><asp:DropDownList ID="Dr_State" runat="server"></asp:DropDownList></td>

                       <td align="right">警号:</td>
                     <td align="left"><asp:TextBox ID="T_JYBH" runat="server" placeholder="请输入警号" OnTextChanged="T_JYBH_TextChanged"></asp:TextBox></td>
                   </tr>

                   <tr>
                       <td align="right">联系人:</td>
                    <td align="left"> <asp:TextBox ID="T_XM" runat="server" placeholder="请输入设备编号"></asp:TextBox></td>
                       <td align="right">联系电话:</td>
                      <td align="left"><asp:TextBox ID="T_SJ" runat="server" placeholder="请输入联系电话"></asp:TextBox></td>
                   </tr>

                   <tr>
                       <td align="right">备注:</td>
                       <td align="left"><asp:TextBox ID="T_BZ" runat="server" TextMode="MultiLine" placeholder="请输入备注信息"></asp:TextBox></td>
                       

                   </tr>
                  

                       <tr>
                             <td align="right">所属部门:</td>
                       <td align="left" colspan="3"> <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList></td>

                   </tr>


                   
                   <tr> 
                       <td colspan="4" align="center">
                            <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">保存</button>

                       </td>

                   </tr>

                   </table>


             

           </div>
    
    </div>
    </form>
</body>
</html>

<script type="text/javascript">
    



</script>