<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

SimMonth = Request.QueryString("M")
SimYear = Request.QueryString("Y")
ASID = Request.QueryString("ASID")

set RecMainRegions = Server.CreateObject("ADODB.Recordset")
RecMainRegions.ActiveConnection = MM_Site_STRING
RecMainRegions.Source = "SELECT * FROM ASs where ASActive = 'True' and ASID = " & ASID
RecMainRegions.CursorType = 0
RecMainRegions.CursorLocation = 2
RecMainRegions.LockType = 3
RecMainRegions.Open()
RecMainRegions_numRows = 0
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
                    <div class="row heading">
                        <div class="eight columns"><h1>System</h1></div>
                    </div>
			<div class="row">
                        <div class="ten columns"><h2>Months Sims Breakdown For Mentor:<br><%=(RecMainRegions.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecMainRegions.Fields.Item("ASLastName").Value)%> - <%=MonthName(SimMonth,True) & " " & SimYear%></h2></div>
                        <div class="two columns buttons"><a href="javascript:history.back(1)" class="nice white radius button">Back</a></div>
                    </div>
<br><br>
                    <table>

<%
LC = 0

set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM Tedis where TediActive = 'True' and ASID = " & ASID & " Order By TediFirstName Asc"
'response.write(RecCurrent.Source)
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0

%>

            
                            <tr>
                                <td><strong><%=AgentLabel%></strong></td>
                                <td><strong>Emp Code</strong></td>
				<td><strong>Sims Allocated</strong></td>
				<td>&nbsp;</td>
                            </tr>
	<tbody>
<%
ZC = 0
While Not RecCurrent.EOF
ZC = ZC + 1

SimsAllocated = 0
SimsNotAllocated = 0

set RecTotalSimsAllocatedToAgents = Server.CreateObject("ADODB.Recordset")
RecTotalSimsAllocatedToAgents.ActiveConnection = MM_Site_STRING
RecTotalSimsAllocatedToAgents.Source = "Select Count(SimID) As TotalSims From ViewSimsAllocationDetails where AllocatedTo <> 0 and Month(AllocatedDate) = " & SimMonth & " and Year(AllocatedDate) = " & SimYear & " and AllocatedTo = " & RecCurrent.Fields.Item("TID").Value
RecTotalSimsAllocatedToAgents.CursorType = 0
RecTotalSimsAllocatedToAgents.CursorLocation = 2
RecTotalSimsAllocatedToAgents.LockType = 3
RecTotalSimsAllocatedToAgents.Open()
RecTotalSimsAllocatedToAgents_numRows = 0
If IsNull(RecTotalSimsAllocatedToAgents.Fields.Item("TotalSims").Value) = "False" Then
SimsAllocated = RecTotalSimsAllocatedToAgents.Fields.Item("TotalSims").Value
End If


%>
                            <tr>
                                <td><%=ZC%>. <%=(RecCurrent.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("TediLastName").Value)%></td>
                                <td><%=(RecCurrent.Fields.Item("TediEmpCode").Value)%></td>
				<td><%=SimsAllocated%></td>
				<td class="action-td"><a href="TediView.asp?TID=<%=(RecCurrent.Fields.Item("TID").Value)%>&Item=14" class="view-button"></a></td>
                            </tr>            
<%
Response.flush
RecCurrent.MoveNext
Wend

%>
		

			</tbody>

			</table>
<!-- #include file="includes/footer.asp" -->

