<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Role_Management.aspx.cs" Inherits="JingWuTong.Personnel.Role_Management" %>

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
         <link href="../css/permissions_set.css" rel="stylesheet" />
        <link href="../Static_Seed_Project/font-awesome/css/font-awesome.min.css" rel="stylesheet" />

<body style="background-color:transparent">

    <form id="form1" runat="server">



          <div style="width:100%;height:100%;padding-left:1em;padding-right:1em">

    <div style="border-bottom-style:dashed;padding-bottom:1em">
        <label style="border-left-color:rgb(72,181,212);border-left-style:solid; color:rgb(72,181,212)"> &nbsp; &nbsp;角色管理  </label>

    </div>

          <table  border="0" cellspacing="0" cellpadding="0" style="color:white;margin:1em 1em 1em 1em;border-collapse:separate; border-spacing:0px 10px;">
             
       
              <tr>
                
                  <td >
                 <asp:TextBox ID="T_search" runat="server" placeholder="输入角色名称进行搜索"></asp:TextBox>
                  <button id="Button1" type="button" class="btn btn-primary btn-xs" onserverclick="Button2_Click" runat="server">搜索</button>
                  </td>

              </tr>
         
              <tr>
                  <td colspan="4">
                       <button type="button" class="btn btn-primary btn-xs" onclick="open_win()" runat="server"  id="B_Add"> 新增</button>
                        &nbsp;
                       <button type="button" class="btn btn-primary btn-xs" onclick="printpage()">打印</button>
                        &nbsp;

                  </td>


              </tr>
           
        </table>


              <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" EnablePaging="true"  SelectCountMethod="getcount" SelectMethod="getcountpaging" TypeName="BLL.B_Role">

                                           <SelectParameters>

                                               
                                         <asp:ControlParameter ControlID="T_search" Name="search" PropertyName="Text" />

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

               

                                       <asp:TemplateField Visible="false">
                                        <ItemTemplate>
                                            <asp:Label ID="i_ID" runat="server" Text='<%#Eval("ID") %>'></asp:Label>
                                        </ItemTemplate>
                                      </asp:TemplateField>

                                            <asp:BoundField DataField="RoleName" HeaderText="角色名称" SortExpression="RoleName">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="XM" HeaderText="创建人" SortExpression="XM" >
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="CreateDate" HeaderText="创建时间" SortExpression="CreateDate">
                                                <ItemStyle HorizontalAlign="Left"  />
                                            </asp:BoundField>
                                             <asp:BoundField DataField="Bz" HeaderText="备注" SortExpression="Bz">
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                           

                                            <asp:TemplateField HeaderText="编辑">
                                                <ItemTemplate>
                                                    <font onclick="open_win_eide('<%#Eval("ID") %>')"                       
                                                        style="cursor: pointer;">
                                                   <img id="img_modify" runat="server" src="../image/修改.png" /><%--<span  class="Langtxt"   id="Modify" ></span>--%></font>
                                                </ItemTemplate>
                                                <ItemStyle HorizontalAlign="Center" Width="50px" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="删除">
                                                <ItemTemplate>
         <asp:LinkButton ID="ImageButton1" OnClientClick="javascript:return confirm('是否删除改角色?')" CommandName="MyDel" CommandArgument='<%# Eval("ID")%>'
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






    </form>
   <div class="modal fade custom-modal in" id="configmodal" tabindex="-1" role="dialog" aria-labelledby="mydeleModeNo"   v-cloak>
        <div class="modal-backdrop fade in" style="height: 955px;z-index: -1;"></div>
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                    <h4 class="modal-title">{{modelname}}</h4>
                </div>
                <div class="modal-body">
            <div class="row"><div>角色名称：</div><div><input v-model="role.name" placeholder="请输入角色名称" type="text" /><span class="err"  v-show="checkvalue.rolename_wrong">*角色名称不能为空</span></div></div>
            <div class="row"><div>创建人：</div><div>{{role.creater}}</div></div>
            <div class="row"><div>备注：</div><div><textarea v-model="role.remark"></textarea></div></div>
            <div class="row"><div>设备状态：</div></div>
            <div class="row">
                <table>
                    <thead>
                      <tr>
                        <th> <input v-model="selctedall"  @click="selctedallfunc()" type="checkbox" />导航名称</th>
                        <th>权限</th>
                      </tr>
                    </thead>
                    <tbody>
                       <template v-for="(item, index) in pages" >
                          <tr class="parent">
                              <td ><input v-model="pages[index].ischecked" type="checkbox" /><span  @click="selectchild(index)">{{pages[index].name}}</span><i :class="faclass(index)" v-show="pages[index].child_page.length > 0"  @click="selectchild(index)"></i></td>
                              <td ><div v-for="(item, index) in pages[index].buttons" >
                                  <input v-model="item.ischecked" type="checkbox" />{{item.name}}
                                 </div> </td>
                            </tr>
                            <tr class="childrow" v-for="(item, n) in pages[index].child_page"  v-show="item.isshow"  >
                              <td><input v-model="item.ischecked" type="checkbox" />{{item.name}}</td>
                              <td>
                                  <div v-for="(item, index) in item.buttons">
                                      <input v-model="item.ischecked" type="checkbox" />{{item.name}}
                                  </div>
                              </td>
                            </tr>
                       </template>
                    </tbody>
                  </table>

            </div>


                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-primary btn-save"  @click="save($event)" >保存</button>
                    <button type="button" class="btn btn-primary btn-qx" data-dismiss="modal">取消</button>
                </div>
            </div>
        </div>
    </div>

      <script src="../comjs/Vue.js"></script>
    <script src="../js/permissions_set.js"></script>
</body>
</html>


<script type="text/javascript">
    //function open_win() {
    //    var iWidth = 800; //弹出窗口的宽度;
    //    var iHeight = 600; //弹出窗口的高度;

    //    var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

    //    var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

    //    window.open("Role_Management_Add.aspx", "_blank", "top=" + iTop + ", left=" + iLeft + ",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=" + iWidth + ", height=" + iHeight + "");
      
    
    //}

    //function open_win_eide(ID) {
    //    var iWidth = 800; //弹出窗口的宽度;
    //    var iHeight = 600; //弹出窗口的高度;

    //    var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;

    //    var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;

    //    window.open("Role_Management_Add.aspx?ID=" + ID, "_blank", "top=" + iTop + ", left=" + iLeft + ",titlebar=no,scrollbars=no,toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=" + iWidth + ", height=" + iHeight + "");
      
      
    //}




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
