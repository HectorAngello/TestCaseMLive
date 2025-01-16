<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetail where  TID = " & Request.QueryString("TID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

set RecTerms = Server.CreateObject("ADODB.Recordset")
RecTerms.ActiveConnection = MM_Site_STRING
RecTerms.Source = "SELECT * FROM TerminationReasons where TermActive = 'True' Order by TermReason Asc"
RecTerms.CursorType = 0
RecTerms.CursorLocation = 2
RecTerms.LockType = 3
RecTerms.Open()
RecTerms_numRows = 0

CanDelete = "Yes"
set RecCheck = Server.CreateObject("ADODB.Recordset")
RecCheck.ActiveConnection = MM_Site_STRING
RecCheck.Source = "SELECT * FROM Tedis where TediActive = 'True' and TediParent = " & Request.QueryString("TID") & " Order By TediFirstName Asc"
RecCheck.CursorType = 0
RecCheck.CursorLocation = 2
RecCheck.LockType = 3
RecCheck.Open()
RecCheck_numRows = 0
If Not RecCheck.EOF and Not RecCheck.BOF Then
CanDelete = "No"
End If
%>
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
                    <div class="row heading"><h1>Delete An Agent: <%=RecEdit.Fields.Item("TediEmpCode").Value%></h1>
			</div>

		<form name="AddUser" action="TediDel2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">

                                First Name<label for="agentEmail"><%=RecEdit.Fields.Item("TediFirstName").Value%></label>
                                Last Name<label for="agentEmail"><%=RecEdit.Fields.Item("TediLastName").Value%></label>
    				Email<%=RecEdit.Fields.Item("TediEmail").Value%></label>
    				Mobile<label for="agentEmail"><%=RecEdit.Fields.Item("TediCell").Value%></label>
    				ID Number<%=RecEdit.Fields.Item("IDNumber").Value%></label>

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


%>

<%
StartDate = Day(RecEdit.Fields.Item("TediStartDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("TediStartDate").Value)) & " " & Year(RecEdit.Fields.Item("TediStartDate").Value)

%>
	Start Date<label for="agentEmail"><%=StartDate%></label>

                                Gender<label for="agentEmail"><%=RecEdit.Fields.Item("GenderType").Value%></label>
                                Race<label for="agentEmail"><%=RecEdit.Fields.Item("RaceLabel").Value%></label>

                                Tax Office<label for="agentEmail"><%=RecEdit.Fields.Item("TaxOffice").Value%></label>

                                Tax No.<label for="agentEmail"><%=RecEdit.Fields.Item("TaxNumber").Value%></label>

                                Purse Limit<label for="agentEmail"><%=RecEdit.Fields.Item("PurseLimit").Value%></label>
<%
If CanDelete = "Yes" Then
%>
				<label for="agentEmail">Termination Reason:</label>
				<select Name="TermReason">
<%
While Not RecTerms.Eof
%>				<option Value="<%=RecTerms.Fields.Item("TermID").Value%>"><%=RecTerms.Fields.Item("TermReason").Value%></option>
<%
RecTerms.MoveNext
Wend
%>
				</select><br>
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
TermDate = Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now)

%>
	<label for="MachineName">Termination Date</label><input type="text" id="datepicker1" Name="ASTermDate" class="input-text" Value="<%=TermDate%>">
<%Else%>
Unable to delete this Agent, the following Sub Agents are allocated to this Agent, Please update these Agents First and then try again.
<%
While Not RecCheck.EOF
%>
				<br><a href="TediEdit.asp?TID=<%=RecCheck.Fields.Item("TID").Value%>"><%=RecCheck.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecCheck.Fields.Item("TediLastName").Value%>&nbsp;(<%=RecCheck.Fields.Item("TediEmpCode").Value%>)</a>
<%
RecCheck.MoveNext
Wend
%>	
<%End If%>
                            </div>
                            <div class="five columns">


                               Residential Address 1<label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress1").Value%></label>
                               Residential Address 2<label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress2").Value%></label>
                               Residential Address 3<label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress3").Value%></label>
                               Residential Address Code<label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialCode").Value%></label>
				<hr><h3>Banking details</h3>
                                Bank<label for="agentEmail"><%=RecEdit.Fields.Item("BankLabel").Value%></label>
                                Branch Code<label for="agentEmail"><%=RecEdit.Fields.Item("BranchCode").Value%></label>
                                Account Type<label for="agentEmail"><%=RecEdit.Fields.Item("AccountLabel").Value%></label>
                                Account Number<label for="agentEmail"><%=RecEdit.Fields.Item("AccNo").Value%></label>
				<hr><h3>Mobi Site Access</h3>
                               Mobi Site Password<label for="agentEmail"><%=RecEdit.Fields.Item("TediPassword").Value%></label>
<%
If CanDelete = "Yes" Then
%>
                                <p><br><br><br>
                                    <input type="Submit" class="orange nice button radius" value="Confirm Agent Deletion">
                                </p>
<%Else%>

<%End If%>
     				</div>
                        </fieldset>
<input type="Hidden" Name="TID" Value="<%=Request.QueryString("TID")%>">
<input type="Hidden" Name="TermedBy" Value="<%=Session("UNID")%>">
                    </form>

		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

