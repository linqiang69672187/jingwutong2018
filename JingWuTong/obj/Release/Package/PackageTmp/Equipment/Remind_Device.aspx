<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Remind_Device.aspx.cs" Inherits="JingWuTong.Remind.Remind_Device" %>

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

        <asp:TextBox ID="HiddenState" runat="server" Width="0px" Height="0px" BorderColor="transparent" ></asp:TextBox>
       <asp:TextBox ID="HiddenNState" runat="server" Width="0px" Height="0px" BorderColor="transparent"  ></asp:TextBox>

     <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;设备提醒  </label>

    </div>

          <table border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:10px 10px;">
               <tr>
                  <td>设备名称:</td>
                  <td>
                   <asp:DropDownList ID="Dr_DeviceNmae" runat="server"  >
                
                     </asp:DropDownList>

                  </td>
                  <td>采购时间:</td>
                  <td> 
      <div class="date" >  <asp:TextBox ID="T_strat" class=" start_form_datetime"  runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

         <label>—</label>

 <div class="date">    <asp:TextBox ID="T_now" class=" end_form_datetime" runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

                  </td>
                    </tr>
               <tr>

               <td>所属部门:</td>
                  <td>  
                   <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" 
OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>

                  </td>



                  <td >提醒类型:</td>
                  <td>
                    <asp:DropDownList ID="Dr_Remind" runat="server">
                        <asp:ListItem Selected="True" Value="-1">全部</asp:ListItem>
                        <asp:ListItem Value="1">检定提醒</asp:ListItem>
                        <asp:ListItem Value="2">报废提醒</asp:ListItem>
                        <asp:ListItem Value="3">保修提醒</asp:ListItem>
                      </asp:DropDownList>
                    <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                   </td>  
              </tr>
              <tr>
                  <td colspan="4">
             
                        <button type="button" class="btn btn-primary btn-xs" runat="server"  onserverclick="btnOutExcel_Click" id="B_Out">导出</button>
                      <%--  &nbsp;
                       <button type="button" class="btn btn-primary btn-xs" onclick="open_win()" >新增</button>--%>
                        &nbsp;
                       <button type="button" class="btn btn-primary btn-xs" onclick="printpage()" runat="server" id="B_Print">打印</button>
                        &nbsp;
                      <button type="button" class="btn btn-primary btn-xs" onclick="OneKey('scrap',40)">一键报废</button>
                   <%--  &nbsp;
                        <button type="button" class="btn btn-primary btn-xs">选择数据项目</button>--%>

                  </td>


              </tr>
           
        </table>

            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td>
                          <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                        DataSourceID="ObjectDataSource1" BorderWidth="0px" CellSpacing="0" BackColor="Transparent"  
                                      BorderStyle="Ridge" AllowPaging="True" AllowSorting="True"
                                        CellPadding="10" GridLines="None"  EmptyDataText="没有查询到相关数据！" 
                                        DataKeyNames="ID" PageSize="10" OnRowDataBound="GridView1_RowDataBound" ForeColor="White" 
                           
                  OnRowEditing="GridView1_RowEditing" OnRowUpdating="GridView1_RowUpdating" OnRowCancelingEdit="GridView1_RowCancelingEdit"  HeaderStyle-Height="30px" >
                                  
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



                 <asp:TemplateField HeaderText="设备名称">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_TypeName" runat="server" Text='<%# Bind("TypeName") %>' Width="60px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_TypeName" runat="server" Text='<%# Bind("TypeName") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>




                 <asp:TemplateField HeaderText="厂家">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_Manufacturer" runat="server" Text='<%# Bind("Manufacturer") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_Manufacturer" runat="server" Text='<%# Bind("Manufacturer") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>



                   <asp:TemplateField HeaderText="型号">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_SBXH" runat="server" Text='<%# Bind("SBXH") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_SBXH" runat="server" Text='<%# Bind("SBXH") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

   


                    <asp:TemplateField HeaderText="项目名称">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_ProjName" runat="server" Text='<%# Bind("ProjName") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_ProjName" runat="server" Text='<%# Bind("ProjName") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

      


                   <asp:TemplateField HeaderText="采购时间">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_CGSJ" runat="server" Text='<%# Bind("CGSJ") %>' Width="60px" onclick="WdatePicker()" class="Wdate"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_CGSJ" runat="server" Text='<%# Bind("CGSJ") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                    
                <asp:TemplateField HeaderText="到期时间">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_BFQX" runat="server" Text='<%# Bind("BFQX") %>' Width="60px" onclick="WdatePicker()" class="Wdate"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_BFQX" runat="server" Text='<%# Bind("BFQX") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                          
                                        
                           <asp:TemplateField HeaderText="单价">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_Price" runat="server" Text='<%# Bind("Price") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_Price" runat="server" Text='<%# Bind("Price") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                          

            <asp:TemplateField HeaderText="提醒类型">
                  <%--  <EditItemTemplate>
                        <asp:TextBox ID="T_Price" runat="server" Text='<%# Bind("Remind") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>--%>
                    <ItemTemplate>
                        <asp:Label ID="L_Remind" runat="server" Text='<%# Bind("Remind") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
         



<%--          <asp:BoundField DataField="Remind" HeaderText="提醒类型" SortExpression="Remind" >
         <ItemStyle HorizontalAlign="Center" />
           </asp:BoundField>--%>




                                  
                             <asp:CommandField HeaderText="编辑" ShowEditButton="True" />
                       
                                  
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

              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcountRemind" SelectMethod="getcountpagingRemind" TypeName="BLL.Device">

                                           <SelectParameters>
                                                        <asp:ControlParameter ControlID="Dr_DeviceNmae" Name="DeviceNmae" PropertyName="SelectedValue" /> <%--设备名称--%>
                                                        <asp:ControlParameter ControlID="T_strat" Name="strat" PropertyName="Text" /><%-- 采购开始时间--%>
                                                         <asp:ControlParameter ControlID="T_now" Name="now" PropertyName="Text" /><%-- 采购结束时间--%>

                                                        <asp:ControlParameter ControlID="Dr_Remind" Name="Remind" PropertyName="SelectedValue" /> <%--设备提醒--%>

                                  
                                         <asp:ControlParameter ControlID="dr_three" Name="three" PropertyName="SelectedValue"/><%-- 所属部门--%>  
                                           <asp:ControlParameter ControlID="dr_second" Name="second" PropertyName="SelectedValue"/><%-- 使用人单位--%>   
                                                 <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" /><%--搜索--%>

                                             </SelectParameters>
                   
                        

                                    </asp:ObjectDataSource>
   

          </div>





         <div class="modal fade"  id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog" >
        <div class="modal-content" style="background-color:rgb(0,15,70)">
            <div class="modal-header">

                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">
					&times;
				</button>
                 <label style="color:white;border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;一键报废   </label>
           
            </div>
            <div class="modal-body" >
                      <table style="border-collapse:separate; border-spacing:6px 10px;" >
        
                   <tr> 
                       <td align="right">报废设备台数:</td>
                       <td align="left"> <label runat="server" id="TS"></label></td>
                    </tr>

                     <tr>
                       <td align="right">操作人:</td>
                       <td align="left"> <label runat="server" id="L_BXR"></label> </td>
                     </tr>

                     <tr>
                       <td align="right">操作时间:</td>
                       <td align="left"> <label runat="server" id="L_LogTime"></label></td>
                     </tr>
                     <tr>
                       <td align="right">备注:</td>
                       <td align="left"><asp:TextBox ID="T_BZ" runat="server" TextMode="MultiLine" placeholder="请输入备注信息"></asp:TextBox></td>
                     </tr>

                          <tr>

                              <td align="center" colspan="2"><button id="Button2" type="button" class="btn btn-primary" onserverclick="btnScrap_Click" runat="server">确认</button></td>

                          </tr>
                          </table>
             



            </div>
       
                
       
        </div>
    </div>
</div>



    </form>

  
</body>
</html>
<script src="../js/Date.js"></script>

<script type="text/javascript">

    function SelectAllCheckboxes(spanChk) {
        elm = document.forms[0];
        for (i = 0; i <= elm.length - 1; i++) {
            if (elm[i].type == "checkbox" && elm[i].id != spanChk.id) {
                if (elm.elements[i].checked != spanChk.checked)
                    elm.elements[i].click();
            }
        }
    }



    function OneKey(State, NState) {

        //获取报修设备台数
        var n = 0
        for (i = 1; i < document.getElementById("GridView1").rows.length; i++) {

            var cb = document.getElementById("GridView1").rows(1).cells(0).children(0);

            if (cb.checked) {
                n++;

            }
        }

        document.getElementById("TS").innerText = n;

        document.getElementById("HiddenState").value = State;

        document.getElementById("HiddenNState").value = NState;

    }



    function printpage() {
        window.print();
    }

    </script>
