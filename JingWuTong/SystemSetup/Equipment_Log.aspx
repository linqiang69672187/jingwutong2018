<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Equipment_Log.aspx.cs" Inherits="JingWuTong.SystemSetup.Equipment_Log" %>

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

</head>
<body style="background-color:transparent">

    <form id="form1" runat="server">



          <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;设备日志  </label>

    </div>

          <table border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:0px 10px;">
     
       
              <tr>

                  <td >使用人单位:</td>
                  <td colspan="4">
                  <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>
              
                  </td>
              </tr>
         

               <tr>
                  <td>设备名称:</td>
                  <td><asp:DropDownList ID="Dr_DeviceNmae" runat="server"  >
                
                      </asp:DropDownList>

                  </td>

                  <td>设备状态:</td>
                  <td><asp:DropDownList ID="Dr_State" runat="server"></asp:DropDownList></td>
        
                   <td >
                  <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

                </tr>

       <%--       <tr>
                  <td colspan="4">
                  <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

              </tr>--%>

           
        </table>


              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcountLog" SelectMethod="getcountpagingLog" TypeName="BLL.B_DeviceLog">

                                           <SelectParameters>
                                              <asp:ControlParameter ControlID="Dr_DeviceNmae" Name="DeviceNmae" PropertyName="SelectedValue" /> <%--设备名称--%>
                                                  
                                               <asp:ControlParameter ControlID="Dr_State" Name="State" PropertyName="SelectedValue" /> <%--设备状态--%>
   
                                               <asp:ControlParameter ControlID="dr_three" Name="three" PropertyName="SelectedValue"/><%-- 使用人单位--%> 
                                               
                                                <asp:ControlParameter ControlID="dr_second" Name="second" PropertyName="SelectedValue"/><%-- 使用人单位--%>  
                                                 
                                     <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" /><%--搜索DevId 设备编号--%>

                                             </SelectParameters>

                                    </asp:ObjectDataSource>

            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td>
                          <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                        DataSourceID="ObjectDataSource1" BorderWidth="0px" CellSpacing="0" BackColor="Transparent"  
                                      BorderStyle="Ridge" AllowPaging="True" AllowSorting="True"
                                        CellPadding="10" GridLines="None"  EmptyDataText="没有查询到相关数据！" 
                                        DataKeyNames="ID" PageSize="10" OnRowDataBound="GridView1_RowDataBound"  ForeColor="White"  HeaderStyle-Height="30px" >
                                  
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

                                            <asp:BoundField DataField="TypeName" HeaderText="设备名称" SortExpression="TypeName">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                                   <asp:BoundField DataField="DevId" HeaderText="设备编号" SortExpression="DevId">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>


                                                <asp:BoundField DataField="StateName" HeaderText="设备状态" SortExpression="StateName">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="UserName" HeaderText="使用人" SortExpression="UserName" >
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                        
                                             <asp:BoundField DataField="BMJC" HeaderText="单位" SortExpression="BMJC">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Tel" HeaderText="联系方式" SortExpression="Tel">
                                                <ItemStyle HorizontalAlign="Left"   />
                                            </asp:BoundField>

                                             <asp:BoundField DataField="LogTime" HeaderText="时间" SortExpression="LogTime">
                                                <ItemStyle HorizontalAlign="Left"   />
                                            </asp:BoundField>
                                           

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
