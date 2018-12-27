<%@ Page Language="C#" EnableEventValidation = "false" AutoEventWireup="true" CodeBehind="Equipment_Register.aspx.cs" Inherits="JingWuTong.Equipment.Equipment_Register" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
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

          textarea {

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
   
 

</head>
<body style="background-color:transparent">

    <form id="form1" runat="server">
 


        <asp:TextBox ID="HiddenState" runat="server" Width="0px" Height="0px" BorderColor="transparent" ></asp:TextBox>
       <asp:TextBox ID="HiddenNState" runat="server" Width="0px" Height="0px" BorderColor="transparent"  ></asp:TextBox>
          


        <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;设备登记  </label>

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
     <div class="date" >  <asp:TextBox ID="T_strat" class=" start_form_datetime"  runat="server" Height="22px" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>
         <label>—</label>
         <div class="date">    <asp:TextBox ID="T_now" class=" end_form_datetime" runat="server" placeholder="请输入搜索时间" ></asp:TextBox><i class="fa fa-calendar"></i></div>

                  </td>
                </tr>
       
              <tr>
                  <td>设备状态:</td>
                  <td>  <asp:DropDownList ID="Dr_State" runat="server"></asp:DropDownList> </td>

                  <td >使用人单位:</td>
                  <td>
                  <asp:DropDownList ID="dr_first" runat="server" AutoPostBack="True" Visible="false"></asp:DropDownList>
                  <asp:DropDownList ID="dr_second" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dr_second_SelectedIndexChanged"></asp:DropDownList>
                  <asp:DropDownList ID="dr_three" runat="server"></asp:DropDownList>
                  
                 
                      </td>
             
              </tr>

              <tr>
                       <td>配发状态</td>
                  <td><asp:DropDownList ID="Dr_AllocateState" runat="server"></asp:DropDownList></td>
                  <td colspan="2"> <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号，警员编号或者警员姓名" Width="260px"></asp:TextBox> 
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button></td>
              </tr>
         
<%--              <tr>
                  <td colspan="4">
                  <asp:TextBox ID="T_search" runat="server" placeholder="输入设备编号进行搜索"></asp:TextBox> 
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

              </tr>--%>
              <tr>
                  <td colspan="4">
                    <button type="button" class="btn btn-primary btn-xs" data-toggle="modal" data-target="#Div_Into"  runat="server" id="B_Into">导入</button>
                       &nbsp;
                    <button type="button" class="btn btn-primary btn-xs" runat="server"  onserverclick="btnOutExcel_Click" id="B_Out">导出</button>
                        &nbsp;
                       <button type="button" class="btn btn-primary btn-xs" onclick="open_win()"  runat="server" id="B_Add">新增</button>
                        <%--&nbsp;--%>
                       <button type="button" class="btn btn-primary btn-xs" onclick="printpage()" runat="server" id="B_Print">打印</button>
                        &nbsp;
                       <button type="button" data-toggle="modal" data-target="#myModal" class="btn btn-primary btn-xs" onclick="OneKey('repairs',20,'报修')">报修</button>
                        &nbsp;
                        <button type="button"  data-toggle="modal" data-target="#myModal"  class="btn btn-primary btn-xs" onclick="OneKey('recycle',1,'回收')">回收</button>
                        &nbsp;
                       <%-- <button type="button" class="btn btn-primary btn-xs">一键年审</button>
                        &nbsp;--%>
                        <button type="button" data-toggle="modal" data-target="#myModal" class="btn btn-primary btn-xs" onclick="OneKey('scrap',40,'报废')">报废</button>
                        &nbsp
                       <button type="button" data-toggle="modal" data-target="#myModal" class="btn btn-primary btn-xs" onclick="OneKey('maintain',30,'同意維修')">同意維修</button>
                        &nbsp;
                      <button type="button" data-toggle="modal" data-target="#myModal" class="btn btn-primary btn-xs" onclick="OneKey('normal',10,'确认修复')">确认修复</button>
                            &nbsp;
                        <button id="B_SelectDate" type="button" class="btn btn-primary btn-xs" onclick="SelectDate(this)">选择数据项目</button>

                  </td>


              </tr>
         
         
        </table>

          <table id="DateItem" border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:10px 10px;">
           <tr>

                      <td><input type="checkbox" class="Item" id="Checkbox1" onclick="FGridView(1,this)"/>设备名称</td>
                     <td><input type="checkbox" class="Item" id="Checkbox2" onclick="FGridView(2, this)"/>设备编号</td>
                     <td><input type="checkbox" class="Item" id="Checkbox3" onclick="FGridView(3, this)"/>使用人姓名</td>
                     <td><input type="checkbox" class="Item" id="Checkbox4" onclick="FGridView(4, this)"/>厂家</td>
                     <td><input type="checkbox" class="Item" id="Checkbox5" onclick="FGridView(5, this)"/>型号</td>
                     <td><input type="checkbox" class="Item" id="Checkbox6" onclick="FGridView(6, this)"/>规格</td>
                     <td><input type="checkbox" class="Item" id="Checkbox7" onclick="FGridView(7, this)"/>采购时间</td>
                     <td><input type="checkbox" class="Item" id="Checkbox8" onclick="FGridView(8, this)"/>项目编号</td>
                     <td><input type="checkbox" class="Item" id="Checkbox9" onclick="FGridView(9, this)"/>项目负责人</td>
                     <td><input type="checkbox" class="Item" id="Checkbox10" onclick="FGridView(10, this)"/>联系电话</td>

              </tr>

            </table>

        

              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcount" SelectMethod="getcountpaging" TypeName="BLL.Device">

                                           <SelectParameters>
                                                        <asp:ControlParameter ControlID="Dr_DeviceNmae" Name="DeviceNmae" PropertyName="SelectedValue" /> <%--设备名称--%>
                                                        <asp:ControlParameter ControlID="T_strat" Name="strat" PropertyName="Text" /><%-- 采购开始时间--%>
                                                         <asp:ControlParameter ControlID="T_now" Name="now" PropertyName="Text" /><%-- 采购结束时间--%>

                                                        <asp:ControlParameter ControlID="Dr_State" Name="State" PropertyName="SelectedValue" /> <%--设备状态--%>
   
                                                    <asp:ControlParameter ControlID="dr_three" Name="three" PropertyName="SelectedValue"/><%-- 使用人单位--%>   

                                                 <asp:ControlParameter ControlID="dr_second" Name="second" PropertyName="SelectedValue"/><%-- 使用人单位--%>   

                                     <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" /><%--搜索DevId 设备编号--%>

                                                 <asp:ControlParameter ControlID="Dr_AllocateState" Name="AllocateState" PropertyName="SelectedValue" /> <%--配发状态--%>

                                             </SelectParameters>

                                    </asp:ObjectDataSource>

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

                     <%--                       
                                            <asp:BoundField DataField="ID" HeaderText="ID" SortExpression="ID">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>--%>

                                       <asp:TemplateField Visible="false">
                                        <ItemTemplate>
                                            <asp:Label ID="L_ID" runat="server" Text='<%#Eval("ID") %>'></asp:Label>
                                             <asp:Label ID="L_DevType" runat="server" Text='<%#Eval("DevType") %>'></asp:Label>
                                             <asp:Label ID="L_DevId" runat="server" Text='<%#Eval("DevId") %>'></asp:Label>
                                             <asp:Label ID="L_BMDM" runat="server" Text='<%#Eval("BMDM") %>'></asp:Label>
                                             <asp:Label ID="L_WorkState" runat="server" Text='<%#Eval("WorkState") %>'></asp:Label>
                                             <asp:Label ID="L_AllocateState" runat="server" Text='<%#Eval("AllocateState") %>'></asp:Label>

                                        </ItemTemplate>
                                      </asp:TemplateField>

                                            <asp:BoundField DataField="TypeName" HeaderText="设备名称" SortExpression="TypeName">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                                  <asp:BoundField DataField="DevId" HeaderText="设备编号" SortExpression="DevId">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                                  <asp:BoundField DataField="XM" HeaderText="使用人姓名" SortExpression="XM">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>



                                                <asp:BoundField DataField="Manufacturer" HeaderText="厂家" SortExpression="Manufacturer">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="SBXH" HeaderText="型号" SortExpression="SBXH" >
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                        
                                             <asp:BoundField DataField="SBGG" HeaderText="规格" SortExpression="SBGG">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CGSJ" HeaderText="采购时间" SortExpression="CGSJ">
                                                <ItemStyle HorizontalAlign="Left"   />
                                            </asp:BoundField>

                                             <asp:BoundField DataField="ProjNum" HeaderText="项目编号" SortExpression="ProjNum">
                                                <ItemStyle HorizontalAlign="Left"   />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="XMFZR"  HeaderText="项目负责人">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>
                                   

                                               <asp:BoundField DataField="XMFZRDH" HeaderText="联系电话" >
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>

                                            <asp:TemplateField HeaderText="更新">
                                                <ItemTemplate>
                                                    <font onclick="open_win_eide('<%#Eval("ID") %>')"                       
                                                        style="cursor: pointer;">
                                                   <img id="img_modify" runat="server" src="../image/修改.png" /><%--<span  class="Langtxt"   id="Modify" ></span>--%></font>
                                                </ItemTemplate>
                                                <ItemStyle HorizontalAlign="Center" Width="50px" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="删除">
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="ImageButton1" OnClientClick="javascript:return confirm('是否删除改设备?')" CommandName="MyDel" CommandArgument='<%# Eval("ID")+","+Eval("JYBH") %>'
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


          </div>

    <div class="modal fade"  id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog" >
        <div class="modal-content" style="background-color:rgb(0,15,70)">
            <div class="modal-header">

                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">
					&times;</button>
                 <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)" id="L_Key">&nbsp; &nbsp; </label>
           
            </div>
            <div class="modal-body" >
                      <table style="color:white;border-collapse:separate; border-spacing:6px 10px;" >
        
                   <tr> 
                       <td align="right">设备台数:</td>
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

                              <td align="center" colspan="2"><button type="button" class="btn btn-primary" onserverclick="btnScrap_Click" runat="server">确认</button></td>

                          </tr>
                          </table>
             



            </div>
       
                
       
        </div>
    </div>
</div>







            <div class="modal fade"  id="Div_Into" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content" style="background-color:rgb(0,15,70)">
            <div class="modal-header">

                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">
					&times;
				</button>
                 <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;导入   </label>
           
            </div>
            <div class="modal-body" >
                      <table style="color:white;border-collapse:separate; border-spacing:6px 10px;" >
        
                   <tr> 
                       <td align="left"><button type="button" class="btn btn-primary" runat="server" onserverclick="btnDownload_Click">下载</button></td>
                       <td align="left"> <label id="Label1" runat="server">下载设备模版，批量导入设备信息</label></td>
                    </tr>

                         <tr>
                     <td align="right">  <asp:FileUpload ID="FileUpload1" runat="server"  BackColor="Transparent" SkinID="Transparent"/></td>
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

<%--   <script src="../js/dataManagement.js"></script>--%>
<script src="../js/Date.js"></script>

<script type="text/javascript">

    function SelectAllCheckboxes(spanChk) {
        elm = document.forms[0];
        for (i = 0; i <= elm.length - 1; i++) {
            if (elm[i].type == "checkbox" && elm[i].id != spanChk.id&& elm.elements[i].className!= "Item") {
                if (elm.elements[i].checked != spanChk.checked)
                    elm.elements[i].click();
            }
        }
    }


    function FGridView(number,mythis)
    {
   

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


        if (document.getElementById("GridView1").rows.length ==1)
        {

            $("#B_SelectDate").attr('disabled', true);;
        }


    }



    loadCheck();


    function SelectDate(mythis)
    {


        if ($(mythis).html() == "选择数据项目") {

            $("#DateItem").css("display", "");
            $(mythis).html("隐藏数据项目");
        }
        else {
           

            $(mythis).html("选择数据项目");
            $("#DateItem").css("display", "none");

        }


    }





    function open_win()
    {
        var iWidth = 800; //弹出窗口的宽度;
        var iHeight = 600; //弹出窗口的高度;

        var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

        var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

        window.open("Equipment_Register_Add.aspx", "_blank", "top="+iTop+", left="+iLeft+",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width="+iWidth+", height="+iHeight+"");
       
    }

    function open_win_eide(ID)
    {
        var iWidth = 800; //弹出窗口的宽度;
        var iHeight = 600; //弹出窗口的高度;

        var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

        var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

        window.open("Equipment_Register_Add.aspx?ID="+ID, "_blank", "top=" + iTop + ", left=" + iLeft + ",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=" + iWidth + ", height=" + iHeight + "");

    }


   







    function OneKey(State,NState,s_Key)
    {
      

        try
        {
      
     
            //获取报修设备台数
            var n = 0

            var a = document.getElementById("GridView1").rows.length;

            for (i = 1; i < document.getElementById("GridView1").rows.length; i++) {

                var cb = document.getElementById("GridView1").rows[i].cells[0].children[0];

                if (cb.checked)
                {
                    n++;

                }
            }

            document.getElementById("L_Key").innerText =s_Key
            document.getElementById("TS").innerText = n;

            document.getElementById("HiddenState").value = State;

            document.getElementById("HiddenNState").value = NState;
       

        }
        catch (e)
        {
            alert(e);
        }

    }


    function printpage() {
        window.print();
    }


</script>
