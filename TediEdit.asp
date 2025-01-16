<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT Top(1)* FROM ViewTediDetail where  TID = " & Request.QueryString("TID")
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
                <div class="content panel">
                    <div class="row heading"><h1>Update Agent: <%=RecEdit.Fields.Item("TediEmpCode").Value%></h1>
			</div>

		<form name="AddUser" action="TediEdit2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">

                                <label for="agentEmail">First Name *</label>
                                <input type="text" name="TediFirstName" class="input-text" required Value="<%=RecEdit.Fields.Item("TediFirstName").Value%>" />
                                <label for="agentEmail">Last Name *</label>
                                <input type="text" name="TediLastName" class="input-text" required Value="<%=RecEdit.Fields.Item("TediLastName").Value%>" />
    
                                <label for="agentCell">Email *</label>
                                <input type="text" name="TediEmail" class="input-text" Value="<%=RecEdit.Fields.Item("TediEmail").Value%>" />
    
                                <label for="agentEmail">Primary Mobile * (e.g. 0831234567)</label>
                                <input type="text" name="TediCell" class="input-text" required Value="<%=RecEdit.Fields.Item("TediCell").Value%>" />

                                <label for="agentEmail">Secondary Mobile * (e.g. 0831234567)</label>
                                <input type="text" name="TediCell2" class="input-text" Value="<%=RecEdit.Fields.Item("TediCell2").Value%>" />

                                <label for="agentEmail">Tertiary Mobile Number * (e.g. 0831234567)</label>
                                <input type="text" name="TertiaryMobileNumber" class="input-text" Value="<%=RecEdit.Fields.Item("TertiaryMobileNumber").Value%>" />

                                <label for="agentCell">ID Number *</label>
                                <input type="text" name="IDNumber" class="input-text" required Value="<%=RecEdit.Fields.Item("IDNumber").Value%>" />

<%
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

set RecGetClusters = Server.CreateObject("ADODB.Recordset")
RecGetClusters.ActiveConnection = MM_Site_STRING
RecGetClusters.Source = "Select * FROM Clusters Where ClusterActive = 'True' Order By ClusterName Asc"
RecGetClusters.CursorType = 0
RecGetClusters.CursorLocation = 2
RecGetClusters.LockType = 3
RecGetClusters.Open()
RecGetClusters_numRows = 0

set RecGetDD = Server.CreateObject("ADODB.Recordset")
RecGetDD.ActiveConnection = MM_Site_STRING
RecGetDD.Source = "Select * FROM DDStatusList where DDStatusActive = 'True' Order By DDStatusOrder Asc"
RecGetDD.CursorType = 0
RecGetDD.CursorLocation = 2
RecGetDD.LockType = 3
RecGetDD.Open()
RecGetDD_numRows = 0

If RecEdit.Fields.Item("TediParent").Value = "0" Then
%>
  				<label for="agentEmail">Sub Region / <%=SupervisorLabel%></label>
                                <select name="SRID" class="dropdown">
<%

set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT * FROM ViewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName,SubRegionName Asc"
'Response.write(RecRegion.Source)
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0
While Not RecRegion.EOF

set RecASs = Server.CreateObject("ADODB.Recordset")
RecASs.ActiveConnection = MM_Site_STRING
RecASs.Source = "Select * FROM ASs where ASActive = 'True' and RID = " & RecRegion.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecASs.CursorType = 0
RecASs.CursorLocation = 2
RecASs.LockType = 3
RecASs.Open()
RecASs_numRows = 0
While Not RecASs.EOF
Selected = ""
If RecRegion.Fields.Item("SRID").Value = RecEdit.Fields.Item("SRID").Value Then
If RecASs.Fields.Item("ASID").Value = RecEdit.Fields.Item("ASID").Value Then
Selected = "Selected"
End If
End If
%>
<option value="<%=RecRegion.Fields.Item("SRID").Value%>-<%=RecASs.Fields.Item("ASID").Value%>" <%=Selected%>><%=RecRegion.Fields.Item("RegionName").Value%> - <%=RecRegion.Fields.Item("SubRegionName").Value%> - <%=RecASs.Fields.Item("ASFirstName").Value%>&nbsp;<%=RecASs.Fields.Item("ASLastName").Value%></option>
<%
RecASs.MoveNext
Wend
RecRegion.Movenext
Wend
%>
                                </select>
<%
Else
%>
<input type="hidden" Name="SRID" Value="<%=RecEdit.Fields.Item("SRID").Value%>-<%=RecEdit.Fields.Item("ASID").Value%>">
<%End If%>
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
StartDate = Day(RecEdit.Fields.Item("TediStartDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("TediStartDate").Value)) & " " & Year(RecEdit.Fields.Item("TediStartDate").Value)
BEESigDate = Day(RecEdit.Fields.Item("TaxOffice").Value) & " " & MonthName(Month(RecEdit.Fields.Item("TaxOffice").Value)) & " " & Year(RecEdit.Fields.Item("TaxOffice").Value)
WorkPermitDate = Day(Now()) & " " & MonthName(Month(Now())) & " " & Year(Now())
If IsDate(RecEdit.Fields.Item("WorkPermitExpiryDate").Value) = "True" then
WorkPermitDate = Day(RecEdit.Fields.Item("WorkPermitExpiryDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("WorkPermitExpiryDate").Value)) & " " & Year(RecEdit.Fields.Item("WorkPermitExpiryDate").Value)
End If
%>
	<label for="MachineName">Start Date *</label><input type="text" id="datepicker" Name="TediStartDate" class="input-text" Value="<%=StartDate%>">

                                <label for="agentEmail">Gender</label>
				<select Name="GenderID">
<%
While Not RecGetGender.EOF
%>
				<option Value="<%=RecGetGender.Fields.Item("GenID").Value%>" <%If RecGetGender.Fields.Item("GenID").Value = RecEdit.Fields.Item("GenderID").Value Then%>Selected<%End If%>><%=RecGetGender.Fields.Item("GenderType").Value%></option>
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
				<option Value="<%=RecGetRace.Fields.Item("RaceID").Value%>" <%If RecGetRace.Fields.Item("RaceID").Value = RecEdit.Fields.Item("RaceID").Value Then%>Selected<%End If%>><%=RecGetRace.Fields.Item("RaceLabel").Value%></option>
<%
RecGetRace.MoveNext
Wend
%>
				</select>


                                <label for="agentEmail">BEE Signature Date *</label>
				<input type="text" id="datepicker1" Name="TaxOffice" class="input-text" Value="<%=BEESigDate%>">

                                <label for="agentEmail">Cluster *</label>
				<select Name="TaxNumber" class="input-text" />
<%
While Not RecGetClusters.EOF
%>				<Option Value="<%=RecGetClusters.Fields.Item("ClusterID").Value%>" <% If RecEdit.Fields.Item("TaxNumber").Value = RecGetClusters.Fields.Item("ClusterID").Value Then%>Selected<%End if%>><%=RecGetClusters.Fields.Item("ClusterName").Value%></Option>
<%
RecGetClusters.MoveNext
Wend
%>
				</select>

                                <label for="agentCell">Trading Spot *</label>
                                <input type="text" name="TradingSpot" class="input-text" Value="<%=RecEdit.Fields.Item("TradingSpot").Value%>" />

                                <label for="agentEmail">Exclude From M-Charge Bulk File Generation</label>
				<Select name="ExcludeFromMchargeBulkFile">
				<option Value="False" <%If RecEdit.Fields.Item("ExcludeFromMchargeBulkFile").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Airtime Purse Limit *<br>(Just the numeric value e.g. 300)</label>
				<input type="text" name="PurseLimit" class="input-text" required Value="<%=RecEdit.Fields.Item("PurseLimit").Value%>" />

                                <label for="agentEmail">Mobile Money Purse Limit *<br>(Just the numeric value e.g. 300)</label>
				<input type="text" name="PurseLimitMM" class="input-text" required value="300" />

                                <label for="agentEmail">Opt In For Real Time Airtime Commission</label>
				<Select name="RealTimeCommOptIn">
				<option Value="False" <%If RecEdit.Fields.Item("RealTimeCommOptIn").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("RealTimeCommOptIn").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>
				<hr>
                                <label for="agentEmail">Consent Form Submitted</label>
				<Select name="DDConsentForm">
				<option Value="False" <%If RecEdit.Fields.Item("DDConsentForm").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDConsentForm").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Crim Check</label>
				<Select name="DDCrimCheck">
				<option Value="False" <%If RecEdit.Fields.Item("DDCrimCheck").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDCrimCheck").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Crim Record</label>
				<Select name="DDCrimRecord">
				<option Value="False" <%If RecEdit.Fields.Item("DDCrimRecord").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDCrimRecord").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">AML Trained</label>
				<Select name="DDAMLTrained">
				<option Value="False" <%If RecEdit.Fields.Item("DDAMLTrained").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDAMLTrained").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">AML Passed</label>
				<Select name="DDAMLPassed">
				<option Value="False" <%If RecEdit.Fields.Item("DDAMLPassed").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDAMLPassed").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Phone Allocated</label>
				<Select name="DDPhoneAllocated">
				<option Value="False" <%If RecEdit.Fields.Item("DDPhoneAllocated").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDPhoneAllocated").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">MSISDN Allocated</label>
				<Select name="DDMSISDNAllocated">
				<option Value="False" <%If RecEdit.Fields.Item("DDMSISDNAllocated").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDMSISDNAllocated").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">TDR Onboarded</label>
				<Select name="DDTDROboarded">
				<option Value="False" <%If RecEdit.Fields.Item("DDTDROboarded").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDTDROboarded").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Validated</label>
				<Select name="DDValidated">
				<option Value="False" <%If RecEdit.Fields.Item("DDValidated").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDValidated").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Agent Status</label>
				<Select name="DDStatusID">
<%
While Not RecGetDD.EOF
%>
				<option Value="<%=RecGetDD.Fields.Item("DDStatusID").Value%>" <%If RecGetDD.Fields.Item("DDStatusID").Value = RecEdit.Fields.Item("DDStatusID").Value Then%>Selected<%End If%>><%=RecGetDD.Fields.Item("DDStatusLabel").Value%></option>
<%
RecGetDD.MoveNext
Wend
%>				</select>
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
				<br><br><br><br><br><hr><h3>Banking details</h3>
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
                                <label for="agentEmail">Branch Code</label>
				<input type="text" name="BranchCode" class="input-text" Value="<%=RecEdit.Fields.Item("BranchCode").Value%>" id="RaffleTickets" />
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
                                <label for="agentEmail">Account Number</label>
				<input type="text" name="AccNo" class="input-text" Value="<%=RecEdit.Fields.Item("AccNo").Value%>" />

                                <label for="agentEmail">Mobile Money Account Number</label>
				<input type="text" name="MoMoAccNo" class="input-text" Value="<%=RecEdit.Fields.Item("MoMoAccNo").Value%>" />

<label for="agentEmail">Airtime Allocation Type</label>
				<select Name="AirtimeTypeID">
<%
set RecGetAirtimeType = Server.CreateObject("ADODB.Recordset")
RecGetAirtimeType.ActiveConnection = MM_Site_STRING
RecGetAirtimeType.Source = "Select * FROM AirtimeAllocationTypes where AirtimeAlloActive = 'True' Order By AirtimeAlloLabel Asc"
RecGetAirtimeType.CursorType = 0
RecGetAirtimeType.CursorLocation = 2
RecGetAirtimeType.LockType = 3
RecGetAirtimeType.Open()
RecGetAirtimeType_numRows = 0
While Not RecGetAirtimeType.EOF
%>
				<option Value="<%=RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value%>" <%If RecEdit.Fields.Item("AirtimeTypeID").Value = RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value Then%>Selected<%End If%>><%=RecGetAirtimeType.Fields.Item("AirtimeAlloLabel").Value%></option>
<%
RecGetAirtimeType.MoveNext
Wend
%>
				</select>
				<hr><h3>Mobi Site Access</h3>
                               <label for="agentEmail">Mobi Site Password *</label>
				<input type="text" name="TediPassword" class="input-text" required Value="<%=RecEdit.Fields.Item("TediPassword").Value%>"/>
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
				<input type="text" name="EmployeeCode" class="input-text"  Value="<%=RecEdit.Fields.Item("TediEmpCode").Value%>" />
<%
Else
%><input type="Hidden" Name="EmployeeCode" Value="<%=RecEdit.Fields.Item("TediEmpCode").Value%>">
<%
End If
%>
				<hr><h3>Agent Products</h3>
                                <label for="agentEmail">Airtime Agent</label>
				<select Name="MChargeTedi">
				<option Value="False" <%If RecEdit.Fields.Item("MChargeTedi").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("MChargeTedi").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Mobile Money Agent</label>
				<select Name="MobileMoneyTedi">
				<option Value="False" <%If RecEdit.Fields.Item("MobileMoneyTedi").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("MobileMoneyTedi").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>


                                <label for="agentEmail">Made for Skhokho GSM</label>
				<select Name="DDSkhokhoGSM">
				<option Value="False" <%If RecEdit.Fields.Item("DDSkhokhoGSM").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDSkhokhoGSM").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

                                <label for="agentEmail">Made for Skhokho Dedicated</label>
				<select Name="DDSkhokhoDedicated">
				<option Value="False" <%If RecEdit.Fields.Item("DDSkhokhoDedicated").Value = "False" Then%>Selected<%End If%>>No</option>
				<option Value="True" <%If RecEdit.Fields.Item("DDSkhokhoDedicated").Value = "True" Then%>Selected<%End If%>>Yes</option>
				</select>

				<hr><h3>Work Permit</h3>

                               <label for="agentEmail">Work Permit Expiry Date *</label>
				<input type="text" id="datepicker2" Name="WorkPermitExpiryDate" class="input-text" Value="<%=WorkPermitDate%>">
                                <p><br><br>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Update Agent">
                                </p>
     				</div>
                        </fieldset>
<%
UpdateReason = "Agent Updated"
If Request.QueryString("UpdateType") = "Reactivate" Then
UpdateReason = "Agent Re-Activated"
End If
%>
<input type="Hidden" Name="UpdateReason" Value="<%=UpdateReason%>">
<input type="Hidden" Name="TID" Value="<%=Request.QueryString("TID")%>">
<input type="Hidden" Name="AgentActive" Value="True">
                    </form>

		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

