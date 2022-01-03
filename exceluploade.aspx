<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="excelupload.aspx.cs" Inherits="cooperativesocietysoftware.excelupload" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="ajaxToolkit" %>
<link href="Content/bootstrap.min.css" rel="stylesheet" type="text/css" />
<link href="Content/bootstrap-responsive.min.css" rel="stylesheet" type="text/css" />
<style type="text/css">
    .nav
    {
        width:100%;
        padding-left:15px;
    }
 
</style> 
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    
    <style type="text/css">
        .auto-style1 {
            width: 100%;
            border-style: solid;
            border-width: 1px;
            background-color: #FF9900;
        }
        .auto-style2 {
            width: 100%;
            height: 90px;
        }
        .auto-style3 {
            width: 375px;
        }
        .auto-style4 {
            margin-left: 0px;
        }
        .auto-style8 {
            width: 100%;
            height: 32px;
            background-color: #C0C0C0;
        }
        .auto-style9 {
            width: 868px;
        }
        .auto-style16 {
            width: 234px
        }
        .auto-style22 {
            width: 100%;
            background-color: #FF9933;
        }
        .auto-style23 {
            height: 99px;
        }
        .auto-style29 {
            height: 99px;
            width: 866px;
        }
        .auto-style32 {
            margin-left: 2px;
            margin-top: 0px;
        }
        .auto-style35 {
            height: 33px;
        }
        .auto-style36 {
            height: 32px;
        }
        .auto-style39 {
            height: 99px;
            width: 260px;
        }
        .auto-style40 {
            height: 99px;
            width: 498px;
        }
        .auto-style41 {
            margin-top: 0px;
        }
    </style>
    
</head>
<body>
    <form id="form1" runat="server">
        <div class="auto-style35">
            <table class="auto-style1">
                <tr>
                    <td class="auto-style36">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label3" runat="server" Text="Welcome" ForeColor="White"></asp:Label>
                    &nbsp;&nbsp;&nbsp;
                        <asp:Label ID="loginname" runat="server" Text="Label"></asp:Label>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:Label ID="loginusertype" runat="server" Text="Label"></asp:Label>
                        </td>
                    <td class="auto-style36"></td>
                </tr>
                </table>
        </div>
    <table class="auto-style2">
        <tr>
            <td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<asp:Image ID="Image1" runat="server" CssClass="auto-style4" Height="48px" ImageUrl="~/images/logo.png" Width="166px" />
            </td>
            <td>
                <asp:Panel ID="Panel1" runat="server" BackColor="#6699FF" Height="84px">
                </asp:Panel>
            </td>
        </tr>
        </table>
    <table class="auto-style8">
        <tr>
            <td class="auto-style16">&nbsp;</td>
            <td class="auto-style9">
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="Medium" Text="THT STAFF (SOMOLU) COOPERATIVE  MULTIPURPOSE SOCIETY"></asp:Label>
            </td>
            <td>&nbsp;</td>
        </tr>
    </table>
        <table class="auto-style22">
            <tr>
                <td class="auto-style40">
                    <asp:Panel ID="Panel4" runat="server" BackColor="#CCCCCC" CssClass="auto-style41" Height="810px">
                        <asp:Menu ID="NavigationMenu0" runat="server"  EnableViewState="false" 
          IncludeStyleBlock="false" Orientation="vertical"
          CssClass="navbar navbar-fixed-middle"
          StaticMenuStyle-CssClass= "nav"
          StaticSelectedStyle-CssClass="active"
          DynamicMenuStyle-CssClass="dropdown-menu" BackColor="Black" ForeColor="Black"
           >
                            <DynamicMenuStyle CssClass="dropdown-menu"></DynamicMenuStyle>
                            <Items>
                                <asp:MenuItem Text="Home" ToolTip="Home" NavigateUrl="~/landmembers.aspx"></asp:MenuItem>
                                <asp:MenuItem NavigateUrl="~/Adminstartup.aspx" Text="Home Admin" Value="Home Admin"></asp:MenuItem>
                                <asp:MenuItem Text="Control Panel" Value="Control Panel">
                                    <asp:MenuItem NavigateUrl="~/Changeusertype.aspx" Text="Change User Types" Value="Change User Types"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Membership" Value="Membership">
                                    <asp:MenuItem Text="Members Awaiting Verification" Value="Members Awaiting Verification" NavigateUrl="~/awaitingnewmembersadmin.aspx"></asp:MenuItem>
                                    <asp:MenuItem Text="Approved Members" Value="Approved Members" NavigateUrl="~/Approvemembersadmin.aspx"></asp:MenuItem>
                                    <asp:MenuItem Text="Members Ledger" Value="Members Ledger"></asp:MenuItem>
                                    <asp:MenuItem Text="Delist a member" Value="Delist a member"></asp:MenuItem>
                                    <asp:MenuItem Text="Members Retirement" Value="Members Retirement">
                                        <asp:MenuItem Text="Retirement List" Value="Retirement List"></asp:MenuItem>
                                        <asp:MenuItem Text="Retirement Verification/Approval" Value="Retirement Verification/Approval"></asp:MenuItem>
                                    </asp:MenuItem>
                                    <asp:MenuItem Text="Reports" Value="Reports"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Savings" Value="Savings"></asp:MenuItem>
                                <asp:MenuItem Text="Loan" Value="Loan">
                                    <asp:MenuItem Text="Loan" Value="Loan">
                                        <asp:MenuItem Text="Loan Verification" Value="Loan Verification"></asp:MenuItem>
                                        <asp:MenuItem Text="Awaiting Verification/Approval" Value="Awaiting Verification/Approval" NavigateUrl="~/awaitingnewloansadmin.aspx"></asp:MenuItem>
                                        <asp:MenuItem Text="Loan Approval" Value="Loan Approval"></asp:MenuItem>
                                        <asp:MenuItem Text="Undisbursed Loans" Value="Undisbursed Loans"></asp:MenuItem>
                                        <asp:MenuItem Text="Loan Approval List" Value="Loan Approval List" NavigateUrl="~/approvedloans.aspx"></asp:MenuItem>
                                        <asp:MenuItem Text="Disapproved Loan List" Value="Disapproved Loan List"></asp:MenuItem>
                                        <asp:MenuItem NavigateUrl="~/adminloansetup.aspx" Text="Loan Types" Value="Loan Types"></asp:MenuItem>
                                    </asp:MenuItem>
                                    <asp:MenuItem Text="Loan Top Up" Value="Loan Top Up">
                                        <asp:MenuItem Text="Verification" Value="Verification"></asp:MenuItem>
                                        <asp:MenuItem Text="Approval" Value="Approval"></asp:MenuItem>
                                        <asp:MenuItem Text="Undisbursed Top Up" Value="Undisbursed Top Up"></asp:MenuItem>
                                        <asp:MenuItem Text="Approved Top Up" Value="Approved Top Up"></asp:MenuItem>
                                    </asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Commodity Program" Value="Commodity Program">
                                    <asp:MenuItem NavigateUrl="~/commoditysetup.aspx" Text="Create Commodities" Value="Create Commodities"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Payments" Value="Payments">
                                    <asp:MenuItem Text="Partial Liquidation" Value="Partial Liquidation" NavigateUrl="~/Paymentpartliquidation.aspx"></asp:MenuItem>
                                    <asp:MenuItem Text="Payments to Members" Value="Payments to Members" NavigateUrl="~/Paymentsmembers.aspx"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Receipts" Value="Receipts">
                                    <asp:MenuItem Text="Monthly Savings Upload" Value="Monthly Savings Upload" NavigateUrl="~/excelupload.aspx"></asp:MenuItem>
                                    <asp:MenuItem Text="Receipts " Value="Receipts from Members" NavigateUrl="~/Receiptadmin.aspx"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Financial Accounts" Value="Financial Accounts"></asp:MenuItem>
                                <asp:MenuItem Text="Application Forms" ToolTip="Click here to fill application forms" Value="Application Forms">
                                    <asp:MenuItem Text="Loan Application" ToolTip="Apply for loan" Value="Loan Application"></asp:MenuItem>
                                    <asp:MenuItem Text="Change Monthly Contribution" ToolTip="Change your monthly contribution" Value="Change Monthly Contribution"></asp:MenuItem>
                                    <asp:MenuItem Text="Flexi Save Application" ToolTip="Flexi Save Application" Value="Flexi Save Application"></asp:MenuItem>
                                    <asp:MenuItem Text="Loan Topup" ToolTip="Loan Topup" Value="Loan Topup"></asp:MenuItem>
                                    <asp:MenuItem Text="Other Schemes" ToolTip="Other Schemes" Value="Other Schemes"></asp:MenuItem>
                                    <asp:MenuItem Text="Closure" ToolTip="Click here if you want to exit the society" Value="Closure"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem NavigateUrl="~/Adminreportsdashboard.aspx" Text="Reports Dashboard" Value="Reports Dashboard"></asp:MenuItem>
                                <asp:MenuItem Text="Resources" ToolTip="Resources" Value="Resources">
                                    <asp:MenuItem Text="Loan Calculator" Value="Loan Calculator"></asp:MenuItem>
                                    <asp:MenuItem Text="Ammortization Schedule" Value="Ammortization Schedule"></asp:MenuItem>
                                    <asp:MenuItem Text="Track My Application" Value="Track My Application"></asp:MenuItem>
                                </asp:MenuItem>
                                <asp:MenuItem Text="Log Out" Value="Log Out" NavigateUrl="~/Login.aspx"></asp:MenuItem>
                            </Items>
                            <StaticMenuStyle CssClass="nav"></StaticMenuStyle>
                            <StaticSelectedStyle CssClass="active"></StaticSelectedStyle>
                        </asp:Menu>
                    </asp:Panel>
                </td>
                <td class="auto-style23">
                    &nbsp;</td>
                <td class="auto-style29">
                    <asp:Panel ID="Panel3" runat="server" BackColor="#FFFFCC" CssClass="auto-style32" Height="810px">
                        &nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label4" runat="server" Text="Date"></asp:Label>
                        &nbsp;&nbsp;&nbsp;
                        <asp:TextBox ID="datetxt" runat="server" TextMode="Date"></asp:TextBox>
                        &nbsp; Enter a date in the month you want to upload savings<br /> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label5" runat="server" Text="Memo"></asp:Label>
&nbsp;
                        <asp:DropDownList ID="DropDownList1" runat="server">
                            <asp:ListItem>Regular Savings</asp:ListItem>
                            <asp:ListItem>Loan</asp:ListItem>
                            <asp:ListItem>shares</asp:ListItem>
                            <asp:ListItem>Trade debtors </asp:ListItem>
                        </asp:DropDownList>
                        &nbsp; Enter the description, e.g December Savings.<br /> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:FileUpload ID="Upload" runat="server" />
                        <br />
&nbsp;
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Monthly contribution upload" />
                        &nbsp; After loading the excel file click Monthly Contribution<br /> &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" PageSize="30">
                            <Columns>
                                <asp:BoundField DataField="memberid" HeaderText="memberid" SortExpression="memberid" />
                                <asp:BoundField DataField="TotalAmount" HeaderText="TotalAmount" SortExpression="TotalAmount" />
                                <asp:BoundField DataField="date" HeaderText="date" SortExpression="date" />
                                <asp:BoundField DataField="details" HeaderText="details" SortExpression="details" />
                            </Columns>
                        </asp:GridView>
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="Post to Members Ledger" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Click Post to members ledger to credit accounts&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="Button" Visible="False" />
                        <br />
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </asp:Panel>
                </td>
                <td class="auto-style39">&nbsp;&nbsp;</td>
                <td class="auto-style23">
                    <br />
                    <br />
                    <br />
                    <br />
                </td>
            </tr>
            </table>
    <p>
                <asp:TextBox ID="passcheck" runat="server" Visible="False"></asp:TextBox>
                    </p>
        <p>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" DeleteCommand="DELETE  FROM [excelimport] " InsertCommand="INSERT INTO excelimport(phone, name, membersid, TotalAmount, email, date, details, transactioncategory) VALUES (@phone, @name, @membersid, @TotalAmount, @email, @date, @details, @transactioncategory)" SelectCommand="SELECT memberid, TotalAmount,date, details  FROM excelimport" UpdateCommand="UPDATE [excelimport] SET  [date] = @date, [details] = @details ">
                    <InsertParameters>
                        <asp:Parameter Name="phone" Type="String" />
                        <asp:Parameter Name="name" Type="String" />
                        <asp:Parameter Name="membersid" />
                        <asp:Parameter Name="TotalAmount" Type="Double" />
                        <asp:Parameter Name="email" Type="String" />
                        <asp:Parameter DbType="Date" Name="date" />
                        <asp:Parameter Name="details" Type="String" />
                        <asp:Parameter Name="transactioncategory" />
                    </InsertParameters>
                    <UpdateParameters>
                        <asp:ControlParameter ControlID="datetxt" DbType="Date" Name="date" PropertyName="Text" />
                        <asp:ControlParameter ControlID="DropDownList1" Name="details" PropertyName="SelectedValue" Type="String" />
                    </UpdateParameters>
                </asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourcepost" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, dateposted, email, transactiontype, description) SELECT name, TotalAmount, date, email, transactioncategory, details FROM excelimport" SelectCommand="SELECT name, TotalAmount, date, email, transactioncategory, details FROM excelimport"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourcepostordinarysavings" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, datetransacted, email, transactiontype) SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport" SelectCommand="SELECT memberid, totalamount, date, transactioncategory, details FROM excelimport"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourcepostloanrepayment" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, datetransacted, email, transactiontype) SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport" SelectCommand="SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourcepostshares" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, datetransacted, email, transactiontype) SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport" SelectCommand="SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourcepostproduct" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, datetransacted, email, transactiontype) SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport" SelectCommand="SELECT memberid, totalamount, date, transactioncategory,details FROM excelimport

"></asp:SqlDataSource>
                    <asp:SqlDataSource ID="SqlDataSourceemailupdate" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" SelectCommand="SELECT FROM members INNER JOIN membersledger ON members.id = membersledger.id" UpdateCommand="update      membersledger
set         email  = members.email
from        membersledger
inner join  members
on         membersledger.memberid= members.memberid"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSourceemailupdate0" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" SelectCommand="SELECT FROM members INNER JOIN membersledger ON members.id = membersledger.id" UpdateCommand="update      membersledger
set         email  = members.email
from        membersledger
inner join  members
on         membersledger.memberid= members.memberid"></asp:SqlDataSource>
                    </p>
                <asp:SqlDataSource ID="SqlDataSourcepostspecialsavings" runat="server" ConnectionString="<%$ ConnectionStrings:memberConnectionString %>" InsertCommand="INSERT INTO membersledger(memberid, creditamount, dateposted, email, transactioncategory) SELECT name, TotalAmount, date, email, details FROM excelimport" SelectCommand="SELECT name, TotalAmount, date, email, transactioncategory, details FROM excelimport"></asp:SqlDataSource>
    </form>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
</body>
</html>
