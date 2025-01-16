<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

SimMonth = Request.QueryString("M")
SimYear = Request.QueryString("Y")

set RecMainRegions = Server.CreateObject("ADODB.Recordset")
RecMainRegions.ActiveConnection = MM_Site_STRING
RecMainRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and UserID = " & Session("UNID") & " Order By RegionName Asc"
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
                        <div class="ten columns"><h2>Months Sims Breakdown Per Mentor: <%=MonthName(SimMonth,True) & " " & SimYear%></h2></div>
                        <div class="two columns buttons"><a href="javascript:history.back(1)" class="nice white radius button">Back</a></div>
                    </div>
<br><br>
                    <table>

<%
LC = 0
While Not RecMainRegions.EOF
set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM ASs where ASActive = 'True' and RID = " & RecMainRegions.Fields.Item("RID").Value & " Order By ASFirstName Asc"
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0
If Not RecCurrent.EOF and Not RecCurrent.BOF then
LC = LC + 1
%>
                            <tr>
                                <td colspan="5"><h4><%=LC%>. <%=(RecMainRegions.Fields.Item("RegionName").Value)%>&nbsp;<%=SupervisorLabel%>s</h4></td>
                            </tr>
            
                            <tr>
                                <td><strong><%=SupervisorLabel%></strong></td>
                                <td><strong><%=SupervisorLabel%> Code</strong></td>
				<td><strong>Sims Allocated</strong></td>
                                <td><strong>Sims Un-Allocated</strong></td>
                                <td><strong>Total Sims Allocated</strong></td>
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
RecTotalSimsAllocatedToAgents.Source = "Select Count(SimID) As TotalSims From Sims where AllocatedTo <> 0 and Month(ImportDate) = " & SimMonth & " and Year(ImportDate) = " & SimYear & " and ASID = " & RecCurrent.Fields.Item("ASID").Value
RecTotalSimsAllocatedToAgents.CursorType = 0
RecTotalSimsAllocatedToAgents.CursorLocation = 2
RecTotalSimsAllocatedToAgents.LockType = 3
RecTotalSimsAllocatedToAgents.Open()
RecTotalSimsAllocatedToAgents_numRows = 0
If IsNull(RecTotalSimsAllocatedToAgents.Fields.Item("TotalSims").Value) = "False" Then
SimsAllocated = RecTotalSimsAllocatedToAgents.Fields.Item("TotalSims").Value
End If

set RecTotalSimsNotAllocatedToAgents = Server.CreateObject("ADODB.Recordset")
RecTotalSimsNotAllocatedToAgents.ActiveConnection = MM_Site_STRING
RecTotalSimsNotAllocatedToAgents.Source = "Select Count(SimID) As TotalSims From Sims where AllocatedTo = 0 and Month(ImportDate) = " & SimMonth & " and Year(ImportDate) = " & SimYear & " and ASID = " & RecCurrent.Fields.Item("ASID").Value
RecTotalSimsNotAllocatedToAgents.CursorType = 0
RecTotalSimsNotAllocatedToAgents.CursorLocation = 2
RecTotalSimsNotAllocatedToAgents.LockType = 3
RecTotalSimsNotAllocatedToAgents.Open()
RecTotalSimsNotAllocatedToAgents_numRows = 0
If IsNull(RecTotalSimsNotAllocatedToAgents.Fields.Item("TotalSims").Value) = "False" Then
SimsNotAllocated = RecTotalSimsNotAllocatedToAgents.Fields.Item("TotalSims").Value
End If
%>
                            <tr>
                                <td><%=ZC%>. <%=(RecCurrent.Fields.Item("ASFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("ASLastName").Value)%></td>
                                <td><%=(RecCurrent.Fields.Item("ASEmpCode").Value)%></td>
				<td><%=SimsAllocated%></td>
                                <td><%=SimsNotAllocated%></td>
                                <td><%=SimsAllocated + SimsNotAllocated%></td>
				<td class="action-td"><a href="ViewMonthsSimsAgents.asp?M=<%=SimMonth%>&Y=<%=SimYear%>&ASID=<%=(RecCurrent.Fields.Item("ASID").Value)%>" class="view-button"></a></td>
                            </tr>            
<%
Response.flush
RecCurrent.MoveNext
Wend

End If
%>
		

			</tbody>
<%
RecMainRegions.MoveNext
Wend
%>
			</table>
<!-- #include file="includes/footer.asp" -->

