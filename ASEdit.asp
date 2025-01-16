<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ASs where  ASID = " & Request.QueryString("ASID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

set RecGetBank2 = Server.CreateObject("ADODB.Recordset")
RecGetBank2.ActiveConnection = MM_Site_STRING
RecGetBank2.Source = "Select * FROM BankTypes Order By BankLabel Asc"
RecGetBank2.CursorType = 0
RecGetBank2.CursorLocation = 2
RecGetBank2.LockType = 3
RecGetBank2.Open()
RecGetBank2_numRows = 0

%>
<script>
  function TicketsQuantity() {
<%While Not RecGetBank2.EOF%>
    if (document.getElementById("RaffleDollars").value == <%=(RecGetBank2.Fields.Item("BankID").Value)%>) {
      document.getElementById("RaffleTickets").value = "<%=(RecGetBank2.Fields.Item("BranchCode").Value)%>";
    }
<%
RecGetBank2.MoveNext
Wend
%>
  }
</script>
<!-- header -->
   <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
            </div>
            <div class="nine columns">
	<%If Request.QueryString("ASUpdated") = "True" Then%><div class="alert-box success"><%=SupervisorLabel%> <strong><%=Request.QueryString("ASName")%> (<%=Request.QueryString("ASEmpCode")%>)</strong> Updated In The System.</div><%End If%>
                <div class="content panel">
                    <div class="eight columns"><h1>Update <%=SupervisorLabel%>: <%=RecEdit.Fields.Item("ASEmpCode").Value%></h1></div>
		    <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
	    	

		<form name="AddUser" action="ASEdit2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">

                                <label for="agentEmail">First Name *</label>
                                <input type="text" name="ASFirstName" class="input-text" required Value="<%=RecEdit.Fields.Item("ASFirstName").Value%>" />
                                <label for="agentEmail">Last Name *</label>
                                <input type="text" name="ASLastName" class="input-text" required Value="<%=RecEdit.Fields.Item("ASLastName").Value%>" />
    
                                <label for="agentCell">Email *</label>
                                <input type="text" name="ASEmail" class="input-text" required Value="<%=RecEdit.Fields.Item("ASEmail").Value%>" />
    
                                <label for="agentEmail">Mobile * (e.g. 0831234567)</label>
                                <input type="text" name="ASCell" class="input-text" required Value="<%=RecEdit.Fields.Item("ASCell").Value%>" />
    
                                <label for="agentCell">ID Number *</label>
                                <input type="text" name="IDNumber" class="input-text" required Value="<%=RecEdit.Fields.Item("IDNumber").Value%>" />

<%
set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0

set RecGetGender = Server.CreateObject("ADODB.Recordset")
RecGetGender.ActiveConnection = MM_Site_STRING
RecGetGender.Source = "Select * FROM GenderTypes Order By GenderType Asc"
RecGetGender.CursorType = 0
RecGetGender.CursorLocation = 2
RecGetGender.LockType = 3
RecGetGender.Open()
RecGetGender_numRows = 0

set RecGetRace = Server.CreateObject("ADODB.Recordset")
RecGetRace.ActiveConnection = MM_Site_STRING
RecGetRace.Source = "Select * FROM RaceTypes Order By RaceLabel Asc"
RecGetRace.CursorType = 0
RecGetRace.CursorLocation = 2
RecGetRace.LockType = 3
RecGetRace.Open()
RecGetRace_numRows = 0

set RecGetBank = Server.CreateObject("ADODB.Recordset")
RecGetBank.ActiveConnection = MM_Site_STRING
RecGetBank.Source = "Select * FROM BankTypes Order By BankLabel Asc"
RecGetBank.CursorType = 0
RecGetBank.CursorLocation = 2
RecGetBank.LockType = 3
RecGetBank.Open()
RecGetBank_numRows = 0

set RecGetACCType = Server.CreateObject("ADODB.Recordset")
RecGetACCType.ActiveConnection = MM_Site_STRING
RecGetACCType.Source = "Select * FROM AccountTypes Order By AccountLabel Asc"
RecGetACCType.CursorType = 0
RecGetACCType.CursorLocation = 2
RecGetACCType.LockType = 3
RecGetACCType.Open()
RecGetACCType_numRows = 0


%>
  				<label for="agentEmail">Region</label>
                                <select name="RID" class="dropdown">
<%While Not RecRegion.EOF%>
                                    <option value="<%=RecRegion.Fields.Item("RID").Value%>" <%If RecEdit.Fields.Item("RID").Value = RecRegion.Fields.Item("RID").Value Then%>Selected<%End If%>><%=RecRegion.Fields.Item("RegionName").Value%></option>
<%
RecRegion.Movenext
Wend
%>
                                </select>
<link rel="stylesheet" href="assets/css/pikaday.css">
    <style>

    a { color: #2996cc; }
    a:hover { text-decoration: none; }

    p { line-height: 1.5em; }
    .small { color: #666; font-size: 10px; }
    .large { font-size: 12px; }

    label {
        font-weight: bold;
    }

    </style> 
<%
StartDate = Day(RecEdit.Fields.Item("ASStartDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("ASStartDate").Value)) & " " & Year(RecEdit.Fields.Item("ASStartDate").Value)

%>
	<label for="MachineName">Start Date *</label><input type="text" id="datepicker1" Name="ASStartDate" class="input-text" Value="<%=StartDate%>">

                                <label for="agentEmail">Gender</label>
				<select Name="GenderID">
<%
While Not RecGetGender.EOF
%>
				<option Value="<%=RecGetGender.Fields.Item("GenID").Value%>" <%If RecEdit.Fields.Item("GenderID").Value = RecGetGender.Fields.Item("GenID").Value Then%>Selected<%End If%>><%=RecGetGender.Fields.Item("GenderType").Value%></option>
<%
RecGetGender.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Race</label>
				<select Name="RaceID">
<%
While Not RecGetRace.EOF
%>
				<option Value="<%=RecGetRace.Fields.Item("RaceID").Value%>" <%If RecEdit.Fields.Item("RaceID").Value = RecGetRace.Fields.Item("RaceID").Value Then%>Selected<%End If%>><%=RecGetRace.Fields.Item("RaceLabel").Value%></option>
<%
RecGetRace.MoveNext
Wend
%>
				</select>


                                <label for="agentEmail">Tax Office *</label>
				<input type="text" name="TaxOffice" class="input-text" Value="<%=RecEdit.Fields.Item("TaxOffice").Value%>" />

                                <label for="agentEmail">Tax No. *</label>
				<input type="text" name="TaxNumber" class="input-text" Value="<%=RecEdit.Fields.Item("TaxNumber").Value%>" />

<%
SystemItem = "235"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
%>
                                <label for="agentEmail">Employee Code *</label>
				<input type="text" name="EmployeeCode" class="input-text" Value="<%=RecEdit.Fields.Item("ASEmpCode").Value%>" />
<%
Else
%><input type="Hidden" Name="EmployeeCode" Value="<%=RecEdit.Fields.Item("ASEmpCode").Value%>">
<%
End If
%>
                                <p>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Update <%=SupervisorLabel%>">
                                </p>
                            </div>
                            <div class="five columns">

                               <label for="agentEmail">Residential Address 1</label>
				<input type="text" name="ResidentialAddress1" class="input-text" Value="<%=RecEdit.Fields.Item("ResidentialAddress1").Value%>" />
                               <label for="agentEmail">Residential Address 2</label>
				<input type="text" name="ResidentialAddress2" class="input-text" Value="<%=RecEdit.Fields.Item("ResidentialAddress2").Value%>" />
                               <label for="agentEmail">Residential Address 3</label>
				<input type="text" name="ResidentialAddress3" class="input-text" Value="<%=RecEdit.Fields.Item("ResidentialAddress3").Value%>" />
                               <label for="agentEmail">Residential Address Code</label>
				<input type="text" name="ResidentialCode" class="input-text" Value="<%=RecEdit.Fields.Item("ResidentialCode").Value%>" />
				<hr><h3>Banking details</h3>
                                <label for="agentEmail">Bank</label>
				<select Name="BankID" id="RaffleDollars" onChange="TicketsQuantity();">
<%
While Not RecGetBank.EOF
%>
				<option Value="<%=RecGetBank.Fields.Item("BankID").Value%>" <%IF RecGetBank.Fields.Item("BankID").Value = RecEdit.Fields.Item("BankID").Value Then%>Selected<%End If%>><%=RecGetBank.Fields.Item("BankLabel").Value%></option>
<%
RecGetBank.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Branch Code *</label>
				<input type="text" name="BranchCode" class="input-text" required  Value="<%=RecEdit.Fields.Item("BranchCode").Value%>"  id="RaffleTickets"/>
                                <label for="agentEmail">Account Type</label>
				<select Name="AccountType">
<%
While Not RecGetACCType.EOF
%>
				<option Value="<%=RecGetACCType.Fields.Item("AccountID").Value%>" <%If RecGetACCType.Fields.Item("AccountID").Value = RecEdit.Fields.Item("AccountType").Value Then%>Selected<%End If%>><%=RecGetACCType.Fields.Item("AccountLabel").Value%></option>
<%
RecGetACCType.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Account Number *</label>
				<input type="text" name="AccNo" class="input-text" required Value="<%=RecEdit.Fields.Item("AccNo").Value%>" />
				<hr><h3>Mobi Site Access</h3>
                               <label for="agentEmail">Mobi Site Password *</label>
				<input type="text" name="ASPassword" class="input-text" required Value="<%=RecEdit.Fields.Item("ASPassword").Value%>" />

     				</div>
                        </fieldset>
<input type="Hidden" Name="ItemID" Value="<%=Request.QueryString("ItemID")%>">
<input type="Hidden" Name="AppCat" Value="<%=Request.QueryString("AppCat")%>">
<input type="Hidden" Name="AppSubCatID" Value="<%=Request.QueryString("AppSubCatID")%>">
<input type="Hidden" Name="ASID" Value="<%=Request.QueryString("ASID")%>">
<input type="Hidden" Name="ASActive" Value="True">
                    </form>
		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

