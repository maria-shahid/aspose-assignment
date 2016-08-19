<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Aspose.Web._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
       <h2>Employee salary increments</h2>

       <p id="paraAttach" runat="server"> Use this page to upload employees excel file and then generate and send them increment letters automatically
           <br />
            <br />
           Download the excel sample file from <asp:Hyperlink ID="lnkSampleFile" runat="server" Text="here" NavigateUrl="~/Samples/Employees.xlsx"></asp:Hyperlink>
           <br /> 
            <br /> 
           <asp:FileUpload ID="employeeFile" runat="server" />
            <asp:RequiredFieldValidator ID="rfvFile" runat="server" ControlToValidate="employeeFile" ErrorMessage="Please attach a file." ForeColor="Red" Display="Dynamic" ValidationGroup="Upload">
                            </asp:RequiredFieldValidator>
               <br /> 
        <asp:Button ID="btnFileUpload" runat="server" CssClass="btn btn-primary btn-lg" Text="Upload" ValidationGroup="Upload" OnClick="btnFileUpload_Click" />
       </p>
       
   
    </div>

    <div class="row">
        <div class="col-md-8">
              <p id="lblMessage" runat="server"></p>
            <p id="paraVerify" runat="server" visible="false">
                Please verify the list of employess before sending them increment letters.
            
        <asp:GridView ID="gvEmployees" runat="server" CssClass="table table-bordered">

        </asp:GridView>

            <asp:Button ID="btnSend" runat="server" Text="Send Increment Letters" CssClass="btn btn-primary" OnClick="btnSend_Click" />
                </p>
        </div>
      
    </div>

</asp:Content>
