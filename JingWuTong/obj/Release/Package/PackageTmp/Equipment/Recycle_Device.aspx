<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Recycle_Device.aspx.cs" Inherits="JingWuTong.Recycle.Recycle_Device" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
   <title>单警科技装备管理系统</title>
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
<body style="background-color:transparent">

    <form id="form1" runat="server">
           <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;回收统计  </label>

    </div>

          <table border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:10px 10px;">
              <tr>
                  <td>所属部门:</td>
                  <td>   <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>

                  </td>
                  <td>申请时间:</td>
                  <td>

  <div class="date" >  <asp:TextBox ID="T_strat" class=" start_form_datetime"  runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

         <label>—</label>

 <div class="date">    <asp:TextBox ID="T_now" class=" end_form_datetime" runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

              
                  </td>
                </tr>
       
              <tr>
                  <td>回收设备:</td>
                  <td colspan="3"><asp:DropDownList ID="Dr_State" runat="server"></asp:DropDownList>
                   <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>

                      </td>
              </tr>
              <tr>
                  <td colspan="4">
        
                 
              <%--         <button type="button" class="btn btn-primary btn-xs" onclick="open_win()" >新增</button>--%>
                      
                         <button type="button" class="btn btn-primary btn-xs" onclick="printpage()" runat="server" id="B_Print">打印</button>
                             &nbsp;
                        <button id="B_Out" type="button" class="btn btn-primary btn-xs" runat="server"   onserverclick="btnOutExcel_Click">导出</button>
                                             &nbsp;
                        <button id="B_SelectDate" type="button" class="btn btn-primary btn-xs" onclick="SelectDate(this)">选择数据项目</button>

                  </td>


              </tr>
           
        </table>


                      <table id="DateItem" border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:10px 10px;">
           <tr>

                      <td><input type="checkbox" class="Item" id="Checkbox1" onclick="FGridView(1, this)"/>所属部门</td>
                     <td><input type="checkbox" class="Item" id="Checkbox2" onclick="FGridView(2, this)"/>联系人</td>
                     <td><input type="checkbox" class="Item" id="Checkbox3" onclick="FGridView(3, this)"/>联系电话</td>
                     <td><input type="checkbox" class="Item" id="Checkbox4" onclick="FGridView(4, this)"/>设备状态</td>
                     <td><input type="checkbox" class="Item" id="Checkbox5" onclick="FGridView(5, this)"/>警号</td>
                     <td><input type="checkbox" class="Item" id="Checkbox6" onclick="FGridView(6, this)"/>回收名称</td>
                     <td><input type="checkbox" class="Item" id="Checkbox7" onclick="FGridView(7, this)"/>设备编号</td>
                     <td><input type="checkbox" class="Item" id="Checkbox8" onclick="FGridView(8, this)"/>回收时间</td>
                     <td><input type="checkbox" class="Item" id="Checkbox9" onclick="FGridView(9, this)"/>备注</td>
                 

              </tr>

            </table>




            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td>
                          <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                        DataSourceID="ObjectDataSource1" BorderWidth="0px" CellSpacing="0" BackColor="Transparent"  
                                      BorderStyle="Ridge" AllowPaging="True" AllowSorting="True"
                                        CellPadding="10" GridLines="None"  EmptyDataText="没有查询到相关数据！" 
                                        DataKeyNames="ID" PageSize="10" OnRowDataBound="GridView1_RowDataBound" OnRowCommand="GridView1_RowCommand" ForeColor="White"  HeaderStyle-Height="30px" >
                                  
                                        <Columns>
                             
         <asp:TemplateField>

                                          <headertemplate>
                                     <asp:CheckBox ID="CheckAll" runat="server" onclick="javascript: SelectAllCheckboxes(this);"  ToolTip="按一次全选，再按一次取消全选" />
                                         </headertemplate>


                                                <ItemTemplate>
                                                    <asp:CheckBox ID="CheckBox1" runat="server" />
                                                </ItemTemplate>
                                           <ItemStyle Width="40px" />
                                            </asp:TemplateField>
                                            
                                       <asp:TemplateField Visible="false">
                                        <ItemTemplate>
                                            <asp:Label ID="L_ID" runat="server" Text='<%#Eval("ID") %>'></asp:Label>
                                        </ItemTemplate>
                                      </asp:TemplateField>


                                     <asp:BoundField DataField="BMMC" HeaderText="所属部门" SortExpression="BMMC">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="BXR" HeaderText="联系人" SortExpression="BXR">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="SJ" HeaderText="联系电话" SortExpression="SJ">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                             <asp:BoundField DataField="StateName" HeaderText="设备状态" SortExpression="StateName">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                             <asp:BoundField DataField="JYBH" HeaderText="警号" SortExpression="JYBH">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TypeName" HeaderText="回收名称" SortExpression="TypeName">
                                                <ItemStyle HorizontalAlign="Left"   />
                                            </asp:BoundField>

                                          
                                            <asp:BoundField DataField="DevId"  HeaderText="设备编号">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>
                                   

                                               <asp:BoundField DataField="LogTime" HeaderText="回收时间" >
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                            
                                            <asp:BoundField DataField="BZ" HeaderText="备注" >
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                            <asp:TemplateField HeaderText="更新" Visible="false">
                                                <ItemTemplate>
                                                    <font onclick="open_win_eide('<%#Eval("ID") %>')"                       
                                                        style="cursor: pointer;">
                                                   <img id="img_modify" runat="server" src="../image/修改.png" /><%--<span  class="Langtxt"   id="Modify" ></span>--%></font>
                                                </ItemTemplate>
                                                <ItemStyle HorizontalAlign="Center" Width="50px" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="删除">
                                                <ItemTemplate>
                                            <asp:LinkButton ID="ImageButton1" OnClientClick="javascript:return confirm('删除该设备后，设备将恢复正常状态是否删除?')" CommandName="MyDel" CommandArgument='<%# Eval("ID")+","+Eval("DevId") %>'
                                                    runat="server">
                                                        <img id="img_del" runat="server" src="../image/删除.png" /><%--<font color="black"><span  class="Langtxt"   id="Delete" ></span></font>--%>
                                                    </asp:LinkButton>
                                                </ItemTemplate>
                                                <ItemStyle HorizontalAlign="Center" Width="50px" />
                                            </asp:TemplateField>
                                        </Columns>

                                      <HeaderStyle CssClass="gridheadcss" Font-Bold="True" BackColor="#116BF2" />
                                        <RowStyle  Height="30px" ForeColor="White" />
                                        <PagerStyle HorizontalAlign="Justify" />




          <PagerTemplate>
          <br />
         <asp:Label ID="lblPage" runat="server" Text='<%# "第" + (((GridView)Container.NamingContainer).PageIndex + 1)  + "页/共" + (((GridView)Container.NamingContainer).PageCount) + "页" %> '></asp:Label>


        <asp:LinkButton ID="lbnFirst" runat="Server" Text="首页"  Enabled='<%# ((GridView)Container.NamingContainer).PageIndex != 0 %>' CommandName="Page" CommandArgument="First"  ></asp:LinkButton>
         <asp:LinkButton ID="lbnPrev" runat="server" Text="上一页" Enabled='<%# ((GridView)Container.NamingContainer).PageIndex != 0 %>' CommandName="Page" CommandArgument="Prev"  ></asp:LinkButton>
        <asp:LinkButton ID="lbnNext" runat="Server" Text="下一页" Enabled='<%# ((GridView)Container.NamingContainer).PageIndex != (((GridView)Container.NamingContainer).PageCount - 1) %>' CommandName="Page" CommandArgument="Next" ></asp:LinkButton>
         <asp:LinkButton ID="lbnLast" runat="Server" Text="尾页"   Enabled='<%# ((GridView)Container.NamingContainer).PageIndex != (((GridView)Container.NamingContainer).PageCount - 1) %>' CommandName="Page" CommandArgument="Last" ></asp:LinkButton>
        <br />
     </PagerTemplate>
                             
                                    </asp:GridView>


                    </td>

                </tr>
            </table>

              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcountRecycle" SelectMethod="getcountpagingRecycle" TypeName="BLL.B_DeviceLog">

                                           <SelectParameters>

                                                      <asp:ControlParameter ControlID="dr_three" Name="three" PropertyName="SelectedValue"/><%-- 所属部门--%>  
                                                        <asp:ControlParameter ControlID="dr_second" Name="second" PropertyName="SelectedValue"/><%-- 使用人单位--%>   
                                                       <asp:ControlParameter ControlID="T_strat" Name="strat" PropertyName="Text" /><%-- 采购开始时间--%>
                                                        <asp:ControlParameter ControlID="T_now" Name="now" PropertyName="Text" /><%-- 采购结束时间--%>
                                                        <asp:ControlParameter ControlID="Dr_State" Name="State" PropertyName="SelectedValue" /> <%--设备状态--%>

                                            
                                         
                                              <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" /><%--搜索--%>

                                             </SelectParameters>

                                    </asp:ObjectDataSource>
   

          </div>
    </form>
  
</body>
</html>

<script src="../js/Date.js"></script>

<script type="text/javascript">



    function SelectAllCheckboxes(spanChk) {
        elm = document.forms[0];
        for (i = 0; i <= elm.length - 1; i++) {
            if (elm[i].type == "checkbox" && elm[i].id != spanChk.id && elm.elements[i].className!= "Item") {
                if (elm.elements[i].checked != spanChk.checked)
                    elm.elements[i].click();
            }
        }
    }


    function FGridView(number, mythis) {


        var rows = document.getElementById("GridView1").rows;



        for (i = 0; i < rows.length; i++) {

            if ($(mythis).is(":checked")) {
                rows[i].cells[number].style.display = "";
            }

            else {



                rows[i].cells[number].style.display = "none";

            }

        }


    }




    function loadCheck() {


        $(".Item").each(function () {

            $(this).attr("checked", true)


        });

        $("#DateItem").css("display", "none");


        if (document.getElementById("GridView1").rows.length == 1) {

            $("#B_SelectDate").attr('disabled', true);;
        }


    }



    loadCheck();


    function SelectDate(mythis) {


        if ($(mythis).html() == "选择数据项目") {

            $("#DateItem").css("display", "");
            $(mythis).html("隐藏数据项目");
        }
        else {


            $(mythis).html("选择数据项目");
            $("#DateItem").css("display", "none");

        }


    }







    function open_win() {
        var iWidth = 800; //弹出窗口的宽度;
        var iHeight = 600; //弹出窗口的高度;

        var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

        var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

        window.open("Equipment_Register_Add.aspx", "_blank", "top=" + iTop + ", left=" + iLeft + ",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=" + iWidth + ", height=" + iHeight + "");

    }

    function open_win_eide(ID) {
        var iWidth = 800; //弹出窗口的宽度;
        var iHeight = 600; //弹出窗口的高度;

        var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

        var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

        window.open("Equipment_Register_Add.aspx?ID=" + ID, "_blank", "top=" + iTop + ", left=" + iLeft + ",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=" + iWidth + ", height=" + iHeight + "");

    }

    function printpage() {
        window.print();
    }


</script>