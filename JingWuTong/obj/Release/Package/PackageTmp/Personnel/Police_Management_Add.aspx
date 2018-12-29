<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Police_Management_Add.aspx.cs" Inherits="JingWuTong.Personnel.Police_Management_Add" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
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
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)" runat="server" id="L_head">    </label>

    </div>

           <div style="color:white;">

               <table style="border-collapse:separate; border-spacing:6px 10px;" >
      
                   <tr> 
                       <td align="right">警员姓名:</td>
                       <td align="left""><asp:TextBox ID="T_XM" runat="server" placeholder="请输入警员姓名"></asp:TextBox></td>
                       <td align="right">联系方式:</td>
                       <td align="left">
                     <asp:TextBox ID="T_SJ" runat="server" placeholder="请输入警员联系方式"></asp:TextBox>

                           <asp:RegularExpressionValidator ID="R_T_SJ" runat="server" ErrorMessage="警员电话格式不正确" ControlToValidate="T_SJ" ForeColor="Red" ValidationExpression="^1[3|4|5|7|8][0-9]{9}$"></asp:RegularExpressionValidator>


 

                       </td>
                   </tr>

                    <tr> 
                       <td align="right"><font>*</font>警号:</td>
                       <td align="left">   <asp:TextBox ID="T_JYBH" runat="server" placeholder="请输入警号"></asp:TextBox>

                           <asp:RequiredFieldValidator ID="R_T_JYBH" runat="server" ErrorMessage="警员编号不能为空" ControlToValidate="T_JYBH" ForeColor="Red"></asp:RequiredFieldValidator>

                       </td>

                       <td align="right">身份证号:</td>
                     <td align="left"><asp:TextBox ID="T_SFZMHM" runat="server" placeholder="请输入身份证" ></asp:TextBox>
                           <asp:RegularExpressionValidator ID="R_T_SFZMHM" runat="server" ErrorMessage="身份证格式不正确" ControlToValidate="T_SFZMHM" ForeColor="Red" ValidationExpression="\d{17}[\d|X]|\d{15}"></asp:RegularExpressionValidator>

                     </td>
                   </tr>

              

                   <tr>
                       <td align="right">角色类型:</td>
                       <td align="left" ><asp:DropDownList ID="Dr_JSID" runat="server"></asp:DropDownList></td>
                       <td align="right">警员类型:</td>
                      <td align="left">  <asp:DropDownList ID="Dr_JYLX" runat="server" AutoPostBack="True"></asp:DropDownList></td>

                   </tr>
                  

                   <tr>
                  <td align="right">警员职务</td>
                  <td align="left" colspan="3"> <asp:DropDownList ID="Dr_LDJB" runat="server" ></asp:DropDownList></td>

                   </tr>

                        <tr>
                       <td align="right">单位:</td>
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