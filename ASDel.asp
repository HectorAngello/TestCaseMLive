<!-- #include file="includes/header.asp" -->

<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
DashboardItemCount = 0

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewASDetail where ASActive = 'True' and ASID = " & Request.QueryString("ASID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

ZZZ = 0
set RecTediCount = Server.CreateObject("ADODB.Recordset")
RecTediCount.ActiveConnection = MM_Site_STRING
RecTediCount.Source = "SELECT * FROM Tedis where  ASID = " & Request.QueryString("ASID") & " Order by TediFirstName Asc"
RecTediCount.CursorType = 0
RecTediCount.CursorLocation = 2
RecTediCount.LockType = 3
RecTediCount.Open()
RecTediCount_numRows = 0
While Not RecTediCount.EOF
ZZZ = ZZZ + 1
RecTediCount.MoveNext
Wend

set RecTerms = Server.CreateObject("ADODB.Recordset")
RecTerms.ActiveConnection = MM_Site_STRING
RecTerms.Source = "SELECT * FROM TerminationReasons where TermActive = 'True' Order by TermReason Asc"
RecTerms.CursorType = 0
RecTerms.CursorLocation = 2
RecTerms.LockType = 3
RecTerms.Open()
RecTerms_numRows = 0
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
                    <div class="row heading"><h1>Delete A <%=SupervisorLabel%>: <%=RecEdit.Fields.Item("ASEmpCode").Value%></h1>
			
 
		

                            <div class="seven columns">

                                First Name :<label for="agentEmail"><%=RecEdit.Fields.Item("ASFirstName").Value%></label>
                                <br>Last Name: <label for="agentEmail"><%=RecEdit.Fields.Item("ASLastName").Value%></label>
    
                                <br>Email: <label for="agentCell"><%=RecEdit.Fields.Item("ASEmail").Value%></label>
    
                                <br>Mobile: <label for="agentCell"><%=RecEdit.Fields.Item("ASCell").Value%></label>
    
                                <br>ID Number: <label for="agentCell"><%=RecEdit.Fields.Item("IDNumber").Value%></label>


  				<br>Region: <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%></label>
 
<%
StartDate = Day(RecEdit.Fields.Item("ASStartDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("ASStartDate").Value)) & " " & Year(RecEdit.Fields.Item("ASStartDate").Value)

%>
	<br>Start Date: <label for="agentEmail"><%=StartDate%></label>

                                <br>Gender: <label for="agentEmail"><%=RecEdit.Fields.Item("GenderType").Value%></label>
                                <br>Race: <label for="agentEmail"><%=RecEdit.Fields.Item("RaceLabel").Value%></label>

                                <br>Tax Office: <label for="agentEmail"><%=RecEdit.Fields.Item("TaxOffice").Value%></label>

                                <br>Tax No.:<label for="agentEmail"><%=RecEdit.Fields.Item("TaxNumber").Value%></label>


                                <p>
<%
If ZZZ > 0 Then
%>				Unable to Delete this <%=SupervisorLabel%>, the following Agents are allocated to this <%=SupervisorLabel%>, please first re-allocate them to a different <%=SupervisorLabel%> and then try again.
<form action="BulkASTediReallocate.asp" Method="Post" Name="Reallocate">
<h3>Bulk Reallocate Agents:</h3>
<table>
<thead>
<tr>
	<th>Agent</th>
	<th>EmpCode</th>
	<th>Status</th>
	<th><%If Request.Querystring("All") = "" Then%><a href="ASDel.asp?ASID=<%=Request.QueryString("ASID")%>&All=True">All</a><%End If%><%If Request.Querystring("All") = "True" Then%><a href="ASDel.asp?ASID=<%=Request.QueryString("ASID")%>">None</a><%End If%></th>
</tr>
</thead>
<tbody>
<%
RecTediCount.MoveFirst
LC = 0
While Not RecTediCount.EOF
LC = LC + 1

CheckedMe = ""
If Request.Querystring("All") = "True" Then
CheckedMe = "Checked"
End If

TediStatus = "Active"
If RecTediCount.Fields.Item("TediActive").Value = "False" Then
TediStatus = "In-Active"
End If

%>
<tr>
<td><%=LC%>. <a href="TediView.asp?TID=<%=RecTediCount.Fields.Item("TID").Value%>"><%=RecTediCount.Fields.Item("TediFirstName").Value%>&nbsp;<%=RecTediCount.Fields.Item("TediLastName").Value%></a></td>
<td><%=RecTediCount.Fields.Item("TediEmpCode").Value%></td>
<td><%=TediStatus%></td>
<td><input type="checkbox" Name="Tedi<%=(RecTediCount.Fields.Item("TID").Value)%>" Value="Yes" <%=CheckedMe%>></td></tr>
<%
RecTediCount.MoveNext
Wend
%>	
</tbody>
</table>
Select New <%=SupervisorLabel%>:	
<select Name="NewASID">
<%
RID = RecEdit.Fields.Item("RID").Value
set RecNewAS = Server.CreateObject("ADODB.Recordset")
RecNewAS.ActiveConnection = MM_Site_STRING
RecNewAS.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RID & " Order By ASFirstName Asc"
RecNewAS.CursorType = 0
RecNewAS.CursorLocation = 2
RecNewAS.LockType = 3
RecNewAS.Open()
RecNewAS_numRows = 0
While Not RecNewAS.EOF
%><option Value="<%=RecNewAS.Fields.Item("ASID").Value%>"><%=RecNewAS.Fields.Item("ASFirstName").Value%>&nbsp;<%=RecNewAS.Fields.Item("ASLastName").Value%></option>
<%
RecNewAS.MoveNext
Wend
%>
</select>
<input type="Hidden" Name="ASID" Value="<%=Request.Querystring("ASID")%>">
<input type="Hidden" Name="UNID" Value="<%=Session("UNID")%>">
<input type="Submit" Value="Reallocate Agents" class="orange nice button radius">
</form>	
<%Else%>
<form name="ASDel" action="ASDel2.asp" method="post"  class="nice">
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
                                    <input type="Submit" class="orange nice button radius" value="Confirm <%=SupervisorLabel%> Deletion">
<input type="Hidden" Name="ASID" Value="<%=Request.QueryString("ASID")%>">
<input type="Hidden" Name="TermedBy" Value="<%=Session("UNID")%>">
                    </form>
<%End If%>
                                </p>
                            </div>
                            <div class="four columns">

                               Residential Address 1: <label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress1").Value%></label>
                               <br>Residential Address 2: <label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress2").Value%></label>
                               <br>Residential Address 3: <label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialAddress3").Value%></label>
                               <br>Residential Address Code: <label for="agentEmail"><%=RecEdit.Fields.Item("ResidentialCode").Value%></label>
				<hr><h3>Banking details</h3>
                                Bank: <label for="agentEmail"><%=RecEdit.Fields.Item("BankLabel").Value%></label>
                                <br>Branch Code: <label for="agentEmail"><%=RecEdit.Fields.Item("BranchCode").Value%></label>
                                <br>Account Type: <label for="agentEmail"><%=RecEdit.Fields.Item("AccountLabel").Value%></label>
                                <br>Account Number: <label for="agentEmail"><%=RecEdit.Fields.Item("AccNo").Value%></label>
				</div>



		</div>
                  


</div>
			
                        
                    
                    
<!-- #include file="includes/footer.asp" -->

