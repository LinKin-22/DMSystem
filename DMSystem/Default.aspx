<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="DMSystem._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h1>Dormitory Management System</h1>
    <h3>Import file form Excel Sheet</h3>
    
    <div>
        <table>
            <tr>
                <td>Select File : </td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
                <td>
                    <asp:Button ID="btnImport" runat="server" Text="Import Data" OnClick="btnImport_Click" />
                </td>
            </tr>
        </table>
        
        <div>
            <br />
            <asp:Label ID="lblMessage" runat="server" Font-Bold="true"/>
            <br />
            <asp:GridView ID="gvData" runat="server" AutoGenerateColumns="False">
                <EmptyDataTemplate>
                    <div style="padding: 10px">
                        Data not found!
                    </div>
                </EmptyDataTemplate>
                <Columns>
                    <asp:BoundField HeaderText="Student ID" DataField="StudentID"/>
                    <asp:BoundField HeaderText="Student Roll" DataField="StudentRoll"/>
                    <asp:BoundField HeaderText="Student Name" DataField="StudentName"/>
                    <asp:BoundField HeaderText="Student Department" DataField="DepartmentName"/>
                    <asp:BoundField HeaderText="Student Address" DataField="StudentAddress"/>
                    <asp:BoundField HeaderText="Student Contact" DataField="ContactTitle"/>
                    <asp:BoundField HeaderText="Student Email" DataField="StudentEmail"/>
                </Columns>
            </asp:GridView>
        </div>
    </div>
    
</asp:Content>
