<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AlarmGate.aspx.cs" Inherits="JingWuTong.SystemSetup.AlarmGate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
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

</head>
<body style="background-color:transparent">

    <form id="form1" runat="server">


    <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;预警设置  </label>

    </div>

          <table border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:0px 10px;">
     <tr>
    

                <td>设备名称:</td>
                      
                  <td> 
                       <asp:DropDownList ID="Dr_DeviceNmae" runat="server"  >
                
                      </asp:DropDownList>

                  </td>

                  <td>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

              </tr>
         
              <tr>
                  <td colspan="4">

         <%--            <button type="button" class="btn btn-primary btn-xs"  >同步</button>
                        &nbsp;--%>
                 <%--   <button type="button" class="btn btn-primary btn-xs" data-toggle="modal" data-target="#Div_Into" runat="server"  id="B_Into">导入</button>--%>
                     <%--   &nbsp;--%>
                    <%--   <button type="button" class="btn btn-primary btn-xs" >保存</button>
                        &nbsp;--%>

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
                                     <asp:CheckBox ID="CheckAll" runat="server" onclick="javascript: SelectAllCheckboxes(this);"  ToolTip="按一次全选，再按

一次取消全选" />
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
                  <%--  <EditItemTemplate>
                        <asp:TextBox ID="T_BM" runat="server" Text='<%# Bind("TypeName") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>--%>
                    <ItemTemplate>
                        <asp:Label ID="L_TypeName" runat="server" Text='<%# Bind("TypeName") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>


                    <asp:TemplateField HeaderText="告警类型">
           
                    <ItemTemplate>
                        <asp:Label ID="L_Name" runat="server" Text='<%# Bind("Name") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>




                 <asp:TemplateField HeaderText="普通告警门限">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_CommonAlarmGate" runat="server" Text='<%# Bind("CommonAlarmGate") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_CommonAlarmGate" runat="server" Text='<%# Bind("CommonAlarmGate") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>



                   <asp:TemplateField HeaderText="严重告警门限">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_UrgencyAlarmGate" runat="server" Text='<%# Bind("UrgencyAlarmGate") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_UrgencyAlarmGate" runat="server" Text='<%# Bind("UrgencyAlarmGate") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

                                  
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

              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcount" SelectMethod="getcountpaging" TypeName="BLL.B_AlarmGate">

                           
                                  <SelectParameters>

                 <asp:ControlParameter ControlID="Dr_DeviceNmae" Name="DeviceNmae" PropertyName="SelectedValue" /> <%--设备名称--%>

                                      </SelectParameters>
                   
                        

                                    </asp:ObjectDataSource>
   

          </div>










    </form>

  
</body>
</html>
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



   


    function printpage() {
        window.print();
    }

    </script>
