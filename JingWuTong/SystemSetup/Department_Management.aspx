<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Department_Management.aspx.cs" Inherits="JingWuTong.SystemSetup.Department_Management" %>

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
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;部门管理  </label>

    </div>

          <table border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:0px 10px;">
     <tr>
                  <td >所属单位:</td>
                  <td>
                  <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" 

OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>
                  </td>

                  <td>
                 <asp:TextBox ID="T_search" runat="server" placeholder="输入部门全称进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

              </tr>
         
              <tr>
                  <td colspan="4">

                     <button type="button" class="btn btn-primary btn-xs" onserverclick="Btn_Click" runat="server" >同步</button>
                        &nbsp;
                    <button type="button" class="btn btn-primary btn-xs" data-toggle="modal" data-target="#Div_Into" runat="server"  id="B_Into">导入</button>
                        &nbsp;
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
                           
                  OnRowEditing="GridView1_RowEditing" OnRowUpdating="GridView1_RowUpdating" OnRowCancelingEdit="GridView1_RowCancelingEdit"  HeaderStyle-Height="30px">
                                  
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



                 <asp:TemplateField HeaderText="部门名称" SortExpression="BM">
           <%--         <EditItemTemplate>
                        <asp:TextBox ID="T_BM" runat="server" Text='<%# Bind("BM") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>--%>
                    <ItemTemplate>
                        <asp:Label ID="L_BM" runat="server" Text='<%# Bind("BM") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>


                    <asp:TemplateField HeaderText="部门简称" SortExpression="BMJC">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_BMJC" runat="server" Text='<%# Bind("BMJC") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_BMJC" runat="server" Text='<%# Bind("BMJC") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>




                 <asp:TemplateField HeaderText="上级单位">
             <%--       <EditItemTemplate>
                        <asp:TextBox ID="T_SJ" runat="server" Text='<%# Bind("SJ") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>--%>
                    <ItemTemplate>
                        <asp:Label ID="L_SJ" runat="server" Text='<%# Bind("SJ") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>



                   <asp:TemplateField HeaderText="地址">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_LXDZ" runat="server" Text='<%# Bind("LXDZ") %>' Width="120px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_LXDZ" runat="server" Text='<%# Bind("LXDZ") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

   


                    <asp:TemplateField HeaderText="坐标" >
                    <EditItemTemplate>
                        <asp:TextBox ID="T_Lo" runat="server" Text='<%# Bind("Lo") %>'  Width="60px"></asp:TextBox>
                          <asp:TextBox ID="T_La" runat="server" Text='<%# Bind("La") %>'  Width="60px"></asp:TextBox>
                 
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_Lo" runat="server" Text='<%# Bind("Lo") %>' Width="60px"></asp:Label>
                        <asp:Label ID="L_La" runat="server" Text='<%# Bind("La") %>' Width="60px"></asp:Label>
                                 
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>

      


                   <asp:TemplateField HeaderText="机构编码" SortExpression="BMDM">
           <%--         <EditItemTemplate>
                        <asp:TextBox ID="T_BMDM" runat="server" Text='<%# Bind("BMDM") %>' Width="60px" ></asp:TextBox>
                    </EditItemTemplate>--%>
                    <ItemTemplate>
                        <asp:Label ID="L_BMDM" runat="server" Text='<%# Bind("BMDM") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                    
                <asp:TemplateField HeaderText="负责人" SortExpression="FZR">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_FZR" runat="server" Text='<%# Bind("FZR") %>' Width="60px" ></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_FZR" runat="server" Text='<%# Bind("FZR") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                          
                                        
                           <asp:TemplateField HeaderText="联系电话" SortExpression="LXDH">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_LXDH" runat="server" Text='<%# Bind("LXDH") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_LXDH" runat="server" Text='<%# Bind("LXDH") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                                          

            <asp:TemplateField HeaderText="排序">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_Sort" runat="server" Text='<%# Bind("Sort") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_Sort" runat="server" Text='<%# Bind("Sort") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
         

         <asp:TemplateField HeaderText="法院">
                    <EditItemTemplate>
                        <asp:TextBox ID="T_FY" runat="server" Text='<%# Bind("FY") %>' Width="40px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="L_FY" runat="server" Text='<%# Bind("FY") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>


     <asp:TemplateField HeaderText="是否显示">
                    <EditItemTemplate>
<%--                        <asp:TextBox ID="T_IsDel" runat="server" Text='<%# Bind("IsDel") %>' Width="40px"></asp:TextBox>--%>

                        <asp:DropDownList ID="D_IsDel" runat="server" Text='<%# Bind("IsDel") %>'>
                        <asp:ListItem Value="0">是</asp:ListItem>
                        <asp:ListItem Value="1">否</asp:ListItem>
        

                        </asp:DropDownList>

                    </EditItemTemplate>
                    <ItemTemplate>
            <asp:Label ID="D_IsDel" runat="server" Text='<%# Bind("IsDel1") %>'></asp:Label>
        
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

              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcount" SelectMethod="getcountpaging" TypeName="BLL.Entity">

                           
                                  <SelectParameters>

                                  <asp:ControlParameter ControlID="dr_three" Name="three" PropertyName="SelectedValue" /><%-- 所属单位--%>   
                                 <asp:ControlParameter ControlID="dr_second" Name="second" PropertyName="SelectedValue"/><%-- 使用人单位--%>  
                                 <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" /><%--搜索--%>

                                      </SelectParameters>
                   
                        

                                    </asp:ObjectDataSource>
   

          </div>






            <div class="modal fade"  id="Div_Into" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content" style="background-color:rgb(0,15,70)">
            <div class="modal-header">

                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">
					&times;</button>
                 <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;导入   </label>
           
            </div>
            <div class="modal-body" >
                      <table style="color:white;border-collapse:separate; border-spacing:6px 10px;" >
        
                   <tr> 
                       <td align="left"><button type="button" class="btn btn-primary" runat="server" onserverclick="btnDownload_Click">下载</button></td>
                       <td align="left"> <label id="Label1" runat="server">下载设备模版，批量导入设备信息</label></td>
                    </tr>

                         <tr>
                     <td align="right">  <asp:FileUpload ID="FileUpload1" runat="server" /></td>
                       <td align="left"> <label id="Label3" runat="server">选择需要导入的文件</label></td>


                         </tr>

                     <tr>
                       <td align="left"> <button type="button" class="btn btn-primary" runat="server"  onserverclick="btnIntExcel_Click" >上传</button></td>
                       <td align="left"> <label id="Label2" runat="server">上传填写或者修改好的设备信息</label> </td>
                     </tr>

                          </table>
             



            </div>

        </div>
    </div>
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

        document.getElementById("HiddenState").innerText = State;

        document.getElementById("HiddenNState").innerText = NState;

    }



    function printpage() {
        window.print();
    }

    </script>
