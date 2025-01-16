<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

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
<!-- #include file="includes/RandomPasswordGen.inc" -->
   <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
            </div>
            <div class="nine columns">
                <div class="content panel">
                    <div class="row heading"><h1>Add A New Agent:</h1>
			</div>
<%If Request.QueryString("TediAdded") = "True" Then%><div class="alert-box success">Agent <strong><%=Request.QueryString("TediName")%> (<%=Request.QueryString("TediEmpCode")%>)</strong> Added To The System.</div><%End If%>
    <form name="form1" method="get">
Add A:   
<select name="menu2" onChange="MM_jumpMenu2('parent',this,0)" Class="text3_frm">
<Option Value="TediAdd.asp">Select Option</Option>

                <option value="TediAdd.asp?Type=New" <%If Request.QueryString("Type") = "New" Then%>Selected<%End If%>>New Agent</option>
                <option value="TediAdd.asp?Type=Sub" <%If Request.QueryString("Type") = "Sub" Then%>Selected<%End If%>>Sub Agent</option>
              </select>
            
        </form>            
<%If Request.QueryString("Type") = "New" Then
set RecMainRegions = Server.CreateObject("ADODB.Recordset")
RecMainRegions.ActiveConnection = MM_Site_STRING
RecMainRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
'Response.write(RecMainRegions.Source)
RecMainRegions.CursorType = 0
RecMainRegions.CursorLocation = 2
RecMainRegions.LockType = 3
RecMainRegions.Open()
RecMainRegions_numRows = 0
%>
<form name="form2" method="get">
Select <%=SupervisorLabel%>:   
<select name="menu2" onChange="MM_jumpMenu2('parent',this,0)" Class="text3_frm">
<Option Value="TediAdd.asp?Type=New">Select Option</Option>
<%
While Not RecMainRegions.EOF
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RecMainRegions.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
While Not RecCurrent.EOF
%>
                <option value="TediAdd.asp?Type=New&ASID=<%=(RecCurrent.Fields.Item("ASID").Value)%>" <%If Request.QueryString("ASID") = Cstr(RecCurrent.Fields.Item("ASID").Value) Then%>Selected<%End If%>><%=(RecMainRegions.Fields.Item("RegionName").Value)%> - <%=(RecCurrent.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("ASLastName").Value)%></option>
<%
RecCurrent.MoveNext
Wend

RecMainRegions.MoveNext
Wend
%>
              </select>
            
        </form>
<%If Request.QueryString("ASID") <> "" Then%>
		<form name="AddUser" action="TediAdd2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">

                                <label for="agentEmail">First Name *</label>
                                <input type="text" name="TediFirstName" class="input-text" required />
                                <label for="agentEmail">Last Name *</label>
                                <input type="text" name="TediLastName" class="input-text" required />
    
                                <label for="agentCell">Email *</label>
                                <input type="text" name="TediEmail" class="input-text" />
    
                                <label for="agentEmail">Primary Mobile * (e.g. 0831234567)</label>
                                <input type="text" name="TediCell" class="input-text" required />

                                <label for="agentEmail">Secondary Mobile * (e.g. 0831234567)</label>
                                <input type="text" name="TediCell2" class="input-text" />

                                <label for="agentEmail">Tertiary Mobile Number * (e.g. 0831234567)</label>
                                <input type="text" name="TertiaryMobileNumber" class="input-text" />
    
                                <label for="agentCell">ID Number *</label>
                                <input type="text" name="IDNumber" class="input-text" required />

<%
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ASs where  ASID = " & Request.QueryString("ASID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT * FROM ViewUserSubRegions where SubRegionActive = 'True' and UserID = " & Session("UNID") & " and RID = " & RecEdit.Fields.Item("RID").Value & " Order By SubRegionName Asc"
'Response.write(RecRegion.Source)
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
%>
  				<label for="agentEmail">Sub Region</label>
                                <select name="SRID" class="dropdown">
<%
While Not RecRegion.EOF
%>
                                    <option value="<%=RecRegion.Fields.Item("SRID").Value%>"><%=RecRegion.Fields.Item("SubRegionName").Value%></option>
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
StartDate = Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)
BEESigDate =  Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)
%>
	<label for="MachineName">Start Date *</label><input type="text" id="datepicker" Name="TediStartDate" class="input-text" Value="<%=StartDate%>">

                                <label for="agentEmail">Gender</label>
				<select Name="GenderID">
<%
While Not RecGetGender.EOF
%>
				<option Value="<%=RecGetGender.Fields.Item("GenID").Value%>"><%=RecGetGender.Fields.Item("GenderType").Value%></option>
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
				<option Value="<%=RecGetRace.Fields.Item("RaceID").Value%>"><%=RecGetRace.Fields.Item("RaceLabel").Value%></option>
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
%>				<Option Value="<%=RecGetClusters.Fields.Item("ClusterID").Value%>"><%=RecGetClusters.Fields.Item("ClusterName").Value%></Option>
<%
RecGetClusters.MoveNext
Wend
%>
				</select>

                                <label for="agentCell">Trading Spot *</label>
                                <input type="text" name="TradingSpot" class="input-text" />


                                <label for="agentEmail">Exclude From M-Charge Bulk File Generation</label>
				<Select name="ExcludeFromMchargeBulkFile">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Purse Limit *<br>(Just the numeric value e.g. 300)</label>
				<input type="text" Value="300" name="PurseLimit" class="input-text" required />

                                <label for="agentEmail">Mobile Money Purse Limit *<br>(Just the numeric value e.g. 500)</label>
				<input type="text" name="PurseLimitMM" class="input-text" required value="100" />

                                <label for="agentEmail">Opt In For Real Time Airtime Commission</label>
				<Select name="RealTimeCommOptIn">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>
<hr>
                                <label for="agentEmail">Consent Form Submitted</label>
				<Select name="DDConsentForm">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Crim Check</label>
				<Select name="DDCrimCheck">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Crim Record</label>
				<Select name="DDCrimRecord">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">AML Trained</label>
				<Select name="DDAMLTrained">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">AML Passed</label>
				<Select name="DDAMLPassed">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Phone Allocated</label>
				<Select name="DDPhoneAllocated">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">MSISDN Allocated</label>
				<Select name="DDMSISDNAllocated">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">TDR Onboarded</label>
				<Select name="DDTDROboarded">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Validated</label>
				<Select name="DDValidated">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Agent Status</label>
				<Select name="DDStatusID">
<%
While Not RecGetDD.EOF
%>
				<option Value="<%=RecGetDD.Fields.Item("DDStatusID").Value%>"><%=RecGetDD.Fields.Item("DDStatusLabel").Value%></option>
<%
RecGetDD.MoveNext
Wend
%>				</select>
                            </div>
                            <div class="five columns">

                               <label for="agentEmail">Residential Address 1</label>
				<input type="text" name="ResidentialAddress1" class="input-text" />
                               <label for="agentEmail">Residential Address 2</label>
				<input type="text" name="ResidentialAddress2" class="input-text" />
                               <label for="agentEmail">Residential Address 3</label>
				<input type="text" name="ResidentialAddress3" class="input-text" />
                               <label for="agentEmail">Residential Address Code</label>
				<input type="text" name="ResidentialCode" class="input-text" />
				<hr><h3>Banking details</h3>
                                <label for="agentEmail">Bank</label>
				<select Name="BankID" id="RaffleDollars" onChange="TicketsQuantity();">
<%
While Not RecGetBank.EOF
%>
				<option Value="<%=RecGetBank.Fields.Item("BankID").Value%>" <%IF RecGetBank.Fields.Item("BankID").Value = "1" Then%>Selected<%End If%>><%=RecGetBank.Fields.Item("BankLabel").Value%></option>
<%
RecGetBank.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Branch Code</label>
				<input type="text" name="BranchCode" class="input-text" id="RaffleTickets" />
                                <label for="agentEmail">Account Type</label>
				<select Name="AccountType">
<%
While Not RecGetACCType.EOF
%>
				<option Value="<%=RecGetACCType.Fields.Item("AccountID").Value%>" <%If RecGetACCType.Fields.Item("AccountID").Value = "1" Then%>Selected<%End If%>><%=RecGetACCType.Fields.Item("AccountLabel").Value%></option>
<%
RecGetACCType.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Account Number</label>
				<input type="text" name="AccNo" class="input-text" />



                                <label for="agentEmail">Mobile Money Account Number</label>
				<input type="text" name="MoMoAccNo" class="input-text" />

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
				<option Value="<%=RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value%>" <%If RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value = "1" Then%>Selected<%End If%>><%=RecGetAirtimeType.Fields.Item("AirtimeAlloLabel").Value%></option>
<%
RecGetAirtimeType.MoveNext
Wend
%>
				</select>

				<hr><h3>Mobi Site Access</h3>
<%
RN = RandomString()
%>
                               <label for="agentEmail">Mobi Site Password *</label>
				<input type="text" name="TediPassword" class="input-text" Value="<%=RN%>" required/>
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
				<input type="text" name="EmployeeCode" class="input-text" />
<%
Else
%><input type="Hidden" Name="EmployeeCode" Value="Generate">
<%
End If
%>
				<hr><h3>Agent Products</h3>
                                <label for="agentEmail">Airtime Agent</label>
				<select Name="MChargeTedi">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Mobile Money Agent</label>
				<select Name="MobileMoneyTedi">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Made for Skhokho GSM</label>
				<select Name="DDSkhokhoGSM">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Made for Skhokho Dedicated</label>
				<select Name="DDSkhokhoDedicated">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

				<hr><h3>Work Permit</h3>

                               <label for="agentEmail">Work Permit Expiry Date *</label>
				<input type="text" id="datepicker2" Name="WorkPermitExpiryDate" class="input-text" Value="<%=BEESigDate%>">
                                <p><br><br>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Add Agent">
                                </p>
     				</div>
                        </fieldset>
<input type="Hidden" Name="ASID" Value="<%=Request.QueryString("ASID")%>">
<input type="Hidden" Name="TediParent" Value="0">
                    </form>
<%
End If
End If
%>
<%If Request.QueryString("Type") = "Sub" Then%>
		<form name="AddUser" action="TediAdd2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">

                                <label for="agentEmail">First Name *</label>
                                <input type="text" name="TediFirstName" class="input-text" required />
                                <label for="agentEmail">Last Name *</label>
                                <input type="text" name="TediLastName" class="input-text" required />
    
                                <label for="agentCell">Email *</label>
                                <input type="text" name="TediEmail" class="input-text" required />
    
                                <label for="agentEmail">Mobile *</label>
                                <input type="text" name="TediCell" class="input-text" required />
    
                                <label for="agentCell">ID Number *</label>
                                <input type="text" name="IDNumber" class="input-text" required />

<%
set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0



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
  				<label for="agentEmail">Parent Agent</label>
                                <select name="TediParent" class="dropdown">
<%
While Not RecEdit.EOF
set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT * FROM ViewTediDetail where TediActive = 'True' and RID = " & RecEdit.Fields.Item("RID").Value & " and TediParent = '0' Order By TediFirstName Asc"
'Response.write(RecRegion.Source)
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0
While Not RecRegion.EOF
%>
                                    <option value="<%=RecRegion.Fields.Item("TID").Value%>"><%=RecRegion.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecRegion.Fields.Item("TediLastName").Value%>&nbsp;(<%=RecRegion.Fields.Item("TediEmpCode").Value%>)</option>
<%
RecRegion.Movenext
Wend
RecEdit.MoveNext
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
StartDate = Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)

%>
	<label for="MachineName">Start Date *</label><input type="text" id="datepicker1" Name="TediStartDate" class="input-text" Value="<%=StartDate%>">

                                <label for="agentEmail">Gender</label>
				<select Name="GenderID">
<%
While Not RecGetGender.EOF
%>
				<option Value="<%=RecGetGender.Fields.Item("GenID").Value%>"><%=RecGetGender.Fields.Item("GenderType").Value%></option>
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
				<option Value="<%=RecGetRace.Fields.Item("RaceID").Value%>"><%=RecGetRace.Fields.Item("RaceLabel").Value%></option>
<%
RecGetRace.MoveNext
Wend
%>
				</select>


                                <label for="agentEmail">Tax Office *</label>
				<input type="text" name="TaxOffice" class="input-text" required />

                                <label for="agentEmail">Tax No. *</label>
				<input type="text" name="TaxNumber" class="input-text" required />

                                <label for="agentEmail">Exclude From Airtime or Mobile Money Bulk File Generation</label>
				<Select name="ExcludeFromMchargeBulkFile">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Airtime Purse Limit *<br>(Just the numeric value e.g. 300)</label>
				<input type="text" name="PurseLimit" class="input-text" required value="300" />

                                <label for="agentEmail">Mobile Money Purse Limit *<br>(Just the numeric value e.g. 300)</label>
				<input type="text" name="PurseLimitMM" class="input-text" required value="300" />
                            </div>
                            <div class="five columns">

                               <label for="agentEmail">Residential Address 1</label>
				<input type="text" name="ResidentialAddress1" class="input-text" />
                               <label for="agentEmail">Residential Address 2</label>
				<input type="text" name="ResidentialAddress2" class="input-text" />
                               <label for="agentEmail">Residential Address 3</label>
				<input type="text" name="ResidentialAddress3" class="input-text" />
                               <label for="agentEmail">Residential Address Code</label>
				<input type="text" name="ResidentialCode" class="input-text" />
				<hr><h3>Banking details</h3>
                                <label for="agentEmail">Bank</label>
				<select Name="BankID">
<%
While Not RecGetBank.EOF
%>
				<option Value="<%=RecGetBank.Fields.Item("BankID").Value%>" <%IF RecGetBank.Fields.Item("BankID").Value = "1" Then%>Selected<%End If%>><%=RecGetBank.Fields.Item("BankLabel").Value%></option>
<%
RecGetBank.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Branch Code *</label>
				<input type="text" name="BranchCode" class="input-text" required />
                                <label for="agentEmail">Account Type</label>
				<select Name="AccountType">
<%
While Not RecGetACCType.EOF
%>
				<option Value="<%=RecGetACCType.Fields.Item("AccountID").Value%>" <%If RecGetACCType.Fields.Item("AccountID").Value = "1" Then%>Selected<%End If%>><%=RecGetACCType.Fields.Item("AccountLabel").Value%></option>
<%
RecGetACCType.MoveNext
Wend
%>
				</select>
                                <label for="agentEmail">Account Number *</label>
				<input type="text" name="AccNo" class="input-text" required />

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
				<option Value="<%=RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value%>" <%If RecGetAirtimeType.Fields.Item("AirtimeTypeID").Value = "1" Then%>Selected<%End If%>><%=RecGetAirtimeType.Fields.Item("AirtimeAlloLabel").Value%></option>
<%
RecGetAirtimeType.MoveNext
Wend
%>
				</select>
				<hr><h3>Mobi Site Access</h3>
                               <label for="agentEmail">Mobi Site Password *</label>

<%
RN = RandomString()
%>
				<input type="text" name="TediPassword" class="input-text" value="<%=RN%>" required/>
				<hr><h3>Agent Products</h3>
                                <label for="agentEmail">Airtime Agent</label>
				<select Name="MChargeTedi">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>

                                <label for="agentEmail">Mobile Money Agent</label>
				<select Name="MobileMoneyTedi">
				<option Value="False">No</option>
				<option Value="True">Yes</option>
				</select>
				<hr><h3>Work Permit</h3>

                               <label for="agentEmail">Work Permit Expiry Date *</label>
				<input type="text" id="datepicker2" Name="WorkPermitExpiryDate" class="input-text" Value="<%=BEESigDate%>">


                                <p><br><br>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Add Agent">
                                </p>
     				</div>
<input Type="Hidden" name="CompanyID" Value="<%=Session("CompanyID")%>">
                        </fieldset>

                    </form>
<%
End If%>
		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

