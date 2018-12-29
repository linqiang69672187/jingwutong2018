<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Role_Management_Add.aspx.cs" Inherits="JingWuTong.Personnel.Role_Management_Add" %>

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


    <script type="text/javascript">
         

        function loadCheck() {

            var checkArr = $("#HiddenPowerLoad").val().split('|');

            for (var i = 0; i < checkArr.length; i++) {
                if (checkArr[i] == 1) {

                    $(".Item:eq(" + i + ")").prop('checked', true);
                }
                else {

                    $(".Item:eq(" + i + ")").prop('checked', false);

                }

            }


        }



    </script>


   <script type="text/javascript" src="../js/jquery/jquery-3.3.1.js"></script>
    <link href="../css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/bootstrap.js"></script>
    <script type="text/javascript" src="../js/My97DatePicker/WdatePicker.js"></script>


</head>
<body style="background:url('../Image/datamanagermentbg.jpg')">
    <form id="form1" runat="server">
           <asp:HiddenField ID="HiddenPower" runat="server"/>
        <asp:HiddenField ID="HiddenPowerLoad" runat="server" />
        <%--  --%>

        <asp:Button ID="Button2" runat="server" Text="Button" Width="0px" Height="0px" OnClick="Button2_Click1" />
       <div style="width:100%;height:100%;padding-left:2em">

    <div style="padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)" runat="server" id="L_head">  </label>

    </div>

           <div style="color:white;">

               <table style="border-collapse:separate; border-spacing:6px 10px;" >
      
                   <tr> 
                       <td align="right"><font>*</font>角色名称:</td>
                       <td align="left""><asp:TextBox ID="T_RoleName" runat="server" placeholder="请输入角色名称"></asp:TextBox></td>
                     
                   </tr>

                    <tr> 
                       <td align="right">创建人:</td>
                       <td align="left"> <label runat="server" id="L_Crateator"> </label></td>

                  
                   </tr>

              

                   <tr>
                       <td align="right">备注:</td>
                       <td align="left" > <asp:TextBox ID="T_Bz"  runat="server" TextMode="MultiLine" placeholder="请输入备注信息"></asp:TextBox></td>
   

                   </tr>
                  

                   <tr>
                       <td colspan="2">

                  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:0px 10px;">
                             <thead>
                                 <tr>
                                     <td><input type="checkbox" id="checkall"/> 导航名称</td>
                                      <td>权限分配</td>
                                    <td>全选</td>
                                 </tr>

                             </thead>

                            <tbody>
                                <tr>
                                     <td><input type="checkbox" class="Item" id="check3"/>首页</td>
                                      <td style="visibility:hidden"><input type="checkbox" class="Item" id="check4" />操作</td>
                                      <td><input type="checkbox" id="home"/></td>
                                 </tr>

                                    <tr>
                                     <td><input type="checkbox" class="Item" id="check5"/>设备查看</td>
                                      <td ><input type="checkbox" class="Item" id="check6"/>操作</td>
                                      <td><input type="checkbox" id="check"/></td>
                                 </tr>
                                    <tr>
                                     <td><input type="checkbox" class="Item" id="check7"/>实时状况</td>
                                      <td style="visibility:hidden"><input type="checkbox" class="Item" id="check8"/>搜索</td>
                                      <td><input type="checkbox" id="status" /></td>
                                 </tr>

                                    <tr>
                                     <td><input type="checkbox" class="Item" id="check9"/>数据统计</td>
                                      <td style="visibility:hidden"> <input type="checkbox" class="Item" id="check10"/>搜索&nbsp;
                                          <input type="checkbox" class="Item" id="check11"/>导出&nbsp;
                                          <input type="checkbox" class="Item" id="check12"/>查看详情&nbsp;

                                      </td>
                                      <td><input type="checkbox" id="statistics"/></td>
                                 </tr>

                                    <tr>
                                     <td><input type="checkbox" class="Item"  id="check13"/>设备管理</td>
                                      <td><input type="checkbox" class="Item" id="check14"/>新增&nbsp;
                                          <input type="checkbox" class="Item" id="check15"/>删除&nbsp;
                                          <input type="checkbox" class="Item" id="check16"/>编辑&nbsp;
                                          <input type="checkbox" class="Item" id="check17"/>导出&nbsp;
                                          <input type="checkbox" class="Item" id="check18"/>导入&nbsp;
                                          <input type="checkbox" class="Item" id="check19"/>打印&nbsp;

                                      </td>
                                      <td><input type="checkbox" id="management"/></td>
                                 </tr>


                                   <tr>
                                     <td><input type="checkbox" class="Item" id="check20"/>人员管理</td>
                                      <td><input type="checkbox" class="Item" id="check21"/>新增&nbsp;
                                          <input type="checkbox" class="Item" id="check22"/>删除&nbsp;
                                          <input type="checkbox" class="Item" id="check23"/>编辑&nbsp;
                                          <input type="checkbox" class="Item" id="check24"/>搜索&nbsp;
                                          <input type="checkbox" class="Item" id="check25"/>导入&nbsp;

                                      </td>
                                      <td><input type="checkbox" id="person"/></td>
                                 </tr>


                                 <tr>
                                     <td><input type="checkbox" class="Item" id="check26"/>系统设置</td>
                                      <td><input type="checkbox" class="Item" id="check27"/>新增&nbsp;
                                          <input type="checkbox" class="Item"  id="check28"/>删除&nbsp;
                                          <input type="checkbox" class="Item" id="check29"/>编辑&nbsp;
                                          <input type="checkbox" class="Item" id="check30"/>搜索&nbsp;
                                          <input type="checkbox" class="Item" id="check31"/>导入&nbsp;
                                            <input type="checkbox" class="Item" id="check32"/>同步&nbsp;
                                      </td>
                                      <td><input type="checkbox" id="system"/></td>
                                 </tr>



                            </tbody>


                         </table>



                       </td>


                   </tr>


                   
                   <tr> 
                       <td colspan="4" align="center">
                            <button id="Button1" type="button" class="btn btn-primary btn-xs" onclick="load()"  runat="server">保存</button>

                       </td>

                   </tr>

                   </table>


             

           </div>
    
    </div>
    </form>
</body>
</html>

<script type="text/javascript">
   
  

    $(function () {

        //首页
        $("#home").click(function () {

            if ($(this).is(":checked")) {

                $("#check3").prop('checked', true);//选中
                $("#check4").prop('checked', true);//选中
            }

            else {

                $("#check3").prop('checked', false);//选中
                $("#check4").prop('checked', false);//选中
            }


        })
      





        //设备查看
        $("#check").click(function () {

            if ($(this).is(":checked")) {

                $("#check5").prop('checked', true);//选中
                $("#check6").prop('checked', true);//选中
            }

            else {

                $("#check5").prop('checked', false);//选中
                $("#check6").prop('checked', false);//选中
            }


        })


        //实时状况
        $("#status").click(function () {

            if ($(this).is(":checked")) {

                $("#check7").prop('checked', true);//选中
                $("#check8").prop('checked', true);//选中
            }

            else {

                $("#check7").prop('checked', false);//选中
                $("#check8").prop('checked', false);//选中
            }


        })



        //数据统计
        $("#statistics").click(function () {

            if ($(this).is(":checked")) {

                $("#check9").prop('checked', true);//选中
                $("#check10").prop('checked', true);//选中
                $("#check11").prop('checked', true);//选中
                $("#check12").prop('checked', true);//选中
            }

            else {

                $("#check9").prop('checked', false);//选中
                $("#check10").prop('checked', false);//选中
                $("#check11").prop('checked', false);//选中
                $("#check12").prop('checked', false);//选中
            }


        })




        //设备管理
        $("#management").click(function () {

            if ($(this).is(":checked")) {

                $("#check13").prop('checked', true);//选中
                $("#check14").prop('checked', true);//选中
                $("#check15").prop('checked', true);//选中
                $("#check16").prop('checked', true);//选中
                $("#check17").prop('checked', true);//选中
                $("#check18").prop('checked', true);//选中
                $("#check19").prop('checked', true);//选中
            }

            else {


                $("#check13").prop('checked', false);//选中
                $("#check14").prop('checked', false);//选中
                $("#check15").prop('checked', false);//选中
                $("#check16").prop('checked', false);//选中
                $("#check17").prop('checked', false);//选中
                $("#check18").prop('checked', false);//选中
                $("#check19").prop('checked', false);//选中
            }

        })



            //人员管理
            $("#person").click(function () {

                if ($(this).is(":checked")) {

                    $("#check20").prop('checked', true);//选中
                    $("#check21").prop('checked', true);//选中
                    $("#check22").prop('checked', true);//选中
                    $("#check23").prop('checked', true);//选中
                    $("#check24").prop('checked', true);//选中
                    $("#check25").prop('checked', true);//选中
        
                }

                else {

                    $("#check20").prop('checked', false);//选中
                    $("#check21").prop('checked', false);//选中
                    $("#check22").prop('checked', false);//选中
                    $("#check23").prop('checked', false);//选中
                    $("#check24").prop('checked', false);//选中
                    $("#check25").prop('checked', false);//选中
                }



            })



            //系统设置
            $("#system").click(function () {

                if ($(this).is(":checked")) {

                    $("#check26").prop('checked', true);//选中
                    $("#check27").prop('checked', true);//选中
                    $("#check28").prop('checked', true);//选中
                    $("#check29").prop('checked', true);//选中
                    $("#check30").prop('checked', true);//选中
                    $("#check31").prop('checked', true);//选中
                    $("#check32").prop('checked', true);//选中

        
                }

                else {

                    $("#check26").prop('checked', false);//选中
                    $("#check27").prop('checked', false);//选中
                    $("#check28").prop('checked', false);//选中
                    $("#check29").prop('checked', false);//选中
                    $("#check30").prop('checked', false);//选中
                    $("#check31").prop('checked', false);//选中
                    $("#check32").prop('checked', false);//选中
                }



            })


            $("#checkall").click(function () { //全选

                if ($("#checkall").is(":checked")) {

                    $("input[type=checkbox]").each(function () {

                        $(this).prop('checked', true);//选中


                    })

                }

                else {

                    $("input[type=checkbox]").each(function () {

                        $(this).prop('checked',false);//选中


                    })

                }


            })


    
    })



    function load()
    {

        var count = "";

        $(".Item").each(function () {

            if ($(this).is(":checked")) {
                count = count + "1|";

            }

            else {

                count = count + "0|";
            }

        })


        $("#HiddenPower").val(count);

        $("#Button2").click();

    }









</script>