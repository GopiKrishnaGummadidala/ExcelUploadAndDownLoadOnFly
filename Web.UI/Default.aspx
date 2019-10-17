<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Web.UI._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>ASP.NET</h1>
        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS, and JavaScript.</p>
        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
    </div>
    <asp:FileUpload ID="flExcel" runat="server" />
    <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click"  Text="Upload" />
    <asp:Button ID="btnDownLoad" runat="server" OnClick="btnDownLoad_Click"   Text="DownLoad" />
    <asp:Label ID="lblStatus" runat="server" ></asp:Label>
    <asp:GridView ID="GridView1" runat="server"></asp:GridView>
</asp:Content>
