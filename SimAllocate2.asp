<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT * FROM ViewTediDetailWithTotals where CompanyID = " & Session("CompanyID") & " and TID = " & Request.QueryString("TID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

UType = 1
UserID = Request.QueryString("TID")

TediType = Session("AgentLabel")
If RecEdit.Fields.Item("TediParent").Value <> 0 Then


set RecParent = Server.CreateObject("ADODB.Recordset")
RecParent.ActiveConnection = MM_Site_STRING
RecParent.Source = "SELECT * FROM ViewTediDetailWithTotals where  TID = " & RecEdit.Fields.Item("TediParent").Value
RecParent.CursorType = 0
RecParent.CursorLocation = 2
RecParent.LockType = 3
RecParent.Open()
RecParent_numRows = 0
TediType = "Sub-" & Session("AgentLabel") & " - <a href='TediView.asp?TID=" & RecEdit.Fields.Item("TediParent").Value & "'>" & RecParent.Fields.Item("TediFirstName").Value & " " & RecParent.Fields.Item("TediLastName").Value & "</a>"
End If


BrickList= ""
If Request.QueryString("BrickCode1") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode1") & " - " 
End If
If Request.QueryString("BrickCode2") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode2") & " - " 
End If
If Request.QueryString("BrickCode3") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode3") & " - " 
End If
If Request.QueryString("BrickCode4") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode4") & " - " 
End If
If Request.QueryString("BrickCode5") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode5") & " - " 
End If
If Request.QueryString("BrickCode6") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode6") & " - " 
End If
If Request.QueryString("BrickCode7") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode7") & " - " 
End If
If Request.QueryString("BrickCode8") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode8") & " - " 
End If
If Request.QueryString("BrickCode9") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode9") & " - " 
End If
If Request.QueryString("BrickCode10") <> "" Then
BrickList = BrickList & Request.QueryString("BrickCode10") & " - " 
End If

%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
		<!-- #include file="Includes/Tedisidebar.asp" -->
            </div>
            <div class="nine columns">
<%If Request.QueryString("TediUpdated") = "True" Then%><div class="alert-box success">Agent Updated In The System.</div><%End If%>
                <div class="content panel">

                        <div class="eight columns"><h1>Agent: <%=RecEdit.Fields.Item("TediEmpCode").Value%></h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br>


                                

<div class="row">
<div class="six columns">
        <span class="spacer-120">First Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("TediFirstName").Value%></label><br>
        <span class="spacer-120">Last Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("TediLastName").Value%></label><br>
        <span class="spacer-120">Email:</span> <label for="agentCell"><%=RecEdit.Fields.Item("TediEmail").Value%></label><br>
        <span class="spacer-120">Mobile:</span> <label for="agentCell"><%=RecEdit.Fields.Item("TediCell").Value%></label><br>
        <span class="spacer-120">Region:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%> - <%=RecEdit.Fields.Item("SubRegionName").Value%></label><br>
</div>
<div class="six columns">
<%
AgentStatus = "Active"
If RecEdit.Fields.Item("TediActive").Value = "False" Then
AgentStatus = "In-Active"
End If
AgentOnWatchList = "No"
If RecEdit.Fields.Item("OnwatchList").Value = "True" Then
AgentOnWatchList = "Yes"
End If

AgentMChargeExclude = "No"
If RecEdit.Fields.Item("ExcludeFromMchargeBulkFile").Value = "True" Then
AgentMChargeExclude = "Yes"
End If
%>
Agent Status: <label for="agentEmail"><%=AgentStatus%></label>
<%If AgentStatus = "Active" then%>
<br>Agent on watchlist: <label for="agentEmail"><%=AgentOnWatchList%></label>
<%End If%>
<br>Agent excluded from Airtime file generation: <label for="agentEmail"><%=AgentMChargeExclude%></label>

</div>
</div>

<%


TotalFNBDeposits = 0
TotalMchargeAllocations = 0
MChargeBalance = 0

If IsNull(RecEdit.Fields.Item("TediTotalBanked").Value) = false then
TotalFNBDeposits = RecEdit.Fields.Item("TediTotalBanked").Value
End If
If IsNull(RecEdit.Fields.Item("TediTotalAllocated").Value) = false then
TotalMchargeAllocations = RecEdit.Fields.Item("TediTotalAllocated").Value
End If
MChargeBalance = TotalMchargeAllocations - TotalFNBDeposits



LastBankedDate = "N/A"
If RecEdit.Fields.Item("LastBankedDate").Value <> "" Then
LastBankedDate = Day(RecEdit.Fields.Item("LastBankedDate").Value) & " " & MonthName(Month(RecEdit.Fields.Item("LastBankedDate").Value)) & " " & Year(RecEdit.Fields.Item("LastBankedDate").Value)
End If
%>


<hr>
			<h2>Capture Sims Against <%=Session("AgentLabel")%> Profile:</h2>
<%
set RecBrick = Server.CreateObject("ADODB.Recordset")
RecBrick.ActiveConnection = MM_Site_STRING
RecBrick.Source = "SELECT * FROM Sims Where ((BoxNumber = '" & Request.QueryString("BrickCode1") & "' or BoxNumber = '" & Request.QueryString("BrickCode2") & "' or BoxNumber = '" & Request.QueryString("BrickCode3") & "' or BoxNumber = '" & Request.QueryString("BrickCode4") & "' or BoxNumber = '" & Request.QueryString("BrickCode5") & "' or BoxNumber = '" & Request.QueryString("BrickCode6") & "' or BoxNumber = '" & Request.QueryString("BrickCode7") & "' or BoxNumber = '" & Request.QueryString("BrickCode8") & "' or BoxNumber = '" & Request.QueryString("BrickCode9") & "' or BoxNumber = '" & Request.QueryString("BrickCode10") & "' ) or (SerialNo = '" & Request.QueryString("BrickCode1") & "' or  SerialNo = '" & Request.QueryString("BrickCode2") & "' or  SerialNo = '" & Request.QueryString("BrickCode3") & "' or  SerialNo = '" & Request.QueryString("BrickCode4") & "' or  SerialNo = '" & Request.QueryString("BrickCode5") & "' or  SerialNo = '" & Request.QueryString("BrickCode6") & "' or  SerialNo = '" & Request.QueryString("BrickCode7") & "' or  SerialNo = '" & Request.QueryString("BrickCode8") & "' or  SerialNo = '" & Request.QueryString("BrickCode9") & "' or SerialNo = '" & Request.QueryString("BrickCode10") & "' ) or (BrickNumber = '" & Request.QueryString("BrickCode1") & "' or BrickNumber = '" & Request.QueryString("BrickCode2") & "' or BrickNumber = '" & Request.QueryString("BrickCode3") & "' or BrickNumber = '" & Request.QueryString("BrickCode4") & "' or BrickNumber = '" & Request.QueryString("BrickCode5") & "' or BrickNumber = '" & Request.QueryString("BrickCode6") & "' or BrickNumber = '" & Request.QueryString("BrickCode7") & "' or BrickNumber = '" & Request.QueryString("BrickCode8") & "' or BrickNumber = '" & Request.QueryString("BrickCode9") & "' or BrickNumber = '" & Request.QueryString("BrickCode10") & "' )) and ASID = " & RecEdit.Fields.Item("ASID").Value & " and AllocatedTo = 0 Order By BrickNumber, BoxNumber, SerialNo Asc"

'RecBrick.Source = "SELECT * FROM CartonNos Where ((CartonNo = '" & Request.QueryString("BrickCode1") & "' or CartonNo = '" & Request.QueryString("BrickCode2") & "' or CartonNo = '" & Request.QueryString("BrickCode3") & "' or CartonNo = '" & Request.QueryString("BrickCode4") & "' or CartonNo = '" & Request.QueryString("BrickCode5") & "' or CartonNo = '" & Request.QueryString("BrickCode6") & "' or CartonNo = '" & Request.QueryString("BrickCode7") & "' or CartonNo = '" & Request.QueryString("BrickCode8") & "' or CartonNo = '" & Request.QueryString("BrickCode9") & "' or CartonNo = '" & Request.QueryString("BrickCode10") & "' ) or (PalletNo = '" & Request.QueryString("BrickCode1") & "' or PalletNo = '" & Request.QueryString("BrickCode2") & "' or PalletNo = '" & Request.QueryString("BrickCode3") & "' or PalletNo = '" & Request.QueryString("BrickCode4") & "' or PalletNo = '" & Request.QueryString("BrickCode5") & "' or PalletNo = '" & Request.QueryString("BrickCode6") & "' or PalletNo = '" & Request.QueryString("BrickCode7") & "' or PalletNo = '" & Request.QueryString("BrickCode8") & "' or PalletNo = '" & Request.QueryString("BrickCode9") & "' or PalletNo = '" & Request.QueryString("BrickCode10") & "' )) and ZID = '0'  Order By CartonNo Asc"


'Response.write(RecBrick.Source)
RecBrick.CursorType = 0
RecBrick.CursorLocation = 2
RecBrick.LockType = 3
RecBrick.Open()
RecBrick_numRows = 0

BrickListT = Len(BrickList)
BrickList = Left(BrickList, BrickListT - 3)
%>
<%If Not RecBrick.EOF and Not RecBrick.BOF Then%>
SIM Numbers Found For <b><%=BrickList%></b><br><br>
<form action="SimAllocate3.asp" method="post">
<table border="0" cellspacing="2" cellpadding="2">
<% VC = 0
While Not RecBrick.EOF
VC = VC + 1
%>
<tr>
<td Class="quote"><%=VC%>. SerialNo:</td><td><%=RecBrick.Fields.Item("SerialNo").Value%></td>
<td Class="quote">Sim Brick:</td><td><%=RecBrick.Fields.Item("BrickNumber").Value%></td>
<td Class="quote">Sim Box:</td><td><%=RecBrick.Fields.Item("BoxNumber").Value%></td>
</tr>
<%
RecBrick.MoveNext
Wend
%>


<tr>
            <td colspan="6" align="center"><label><input type="Hidden" Name="Token" Value="<%=Session("DMGToken")%>">
              <input name="button2" type="submit" class="orange nice button radius" id="button2" value="Capture Sims">
            </label></td>
          </tr>
  </table>
<input type="hidden" Name="BrickCode1" Value="<%=Request.QueryString("BrickCode1")%>">
<input type="hidden" Name="BrickCode2" Value="<%=Request.QueryString("BrickCode2")%>">
<input type="hidden" Name="BrickCode3" Value="<%=Request.QueryString("BrickCode3")%>">
<input type="hidden" Name="BrickCode4" Value="<%=Request.QueryString("BrickCode4")%>">
<input type="hidden" Name="BrickCode5" Value="<%=Request.QueryString("BrickCode5")%>">

<input type="hidden" Name="BrickCode6" Value="<%=Request.QueryString("BrickCode6")%>">
<input type="hidden" Name="BrickCode7" Value="<%=Request.QueryString("BrickCode7")%>">
<input type="hidden" Name="BrickCode8" Value="<%=Request.QueryString("BrickCode8")%>">
<input type="hidden" Name="BrickCode9" Value="<%=Request.QueryString("BrickCode9")%>">
<input type="hidden" Name="BrickCode10" Value="<%=Request.QueryString("BrickCode10")%>">
<input type="Hidden" Name="UNID" Value="<%=Session("UNID")%>">
<input type="Hidden" Name="TID" Value="<%=Request.QueryString("TID")%>">
<input type="Hidden" Name="ZonerEmail" Value="<%=RecEdit.Fields.Item("TediEmail").Value%>">
<input type="Hidden" Name="ASID" Value="<%=RecEdit.Fields.Item("ASID").Value%>">
<input type="Hidden" Name="ZonerName" Value="<%=RecEdit.Fields.Item("TediFirstName").Value%>">
<input type="Hidden" Name="ZonerCell" Value="<%=RecEdit.Fields.Item("TediCell").Value%>">
</form>
<%Else
%>
<p>Unable to find any matching brick codes for <%=BrickList%>.
<br><br>Please ensure the sims are allocated to this <%=AgentLabel%>s <%=SupervisorLabel%>s Profile first. Only sims allocated to the <%=SupervisorLabel%>s profile can be allocated against the <%=AgentLabel%>s profile.</p>
<%End If%>



                    </div>
<!-- #include file="includes/footer.asp" -->

