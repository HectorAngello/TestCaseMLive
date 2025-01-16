<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
%>
<%
set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT * FROM viewTediDetail Where TID = " & Request.QueryString("TID") 
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0
ZonerPurseLimit = RecZoner.Fields.Item("PurseLimit").Value

set RecZ = Server.CreateObject("ADODB.Recordset")
RecZ.ActiveConnection = MM_Site_STRING
RecZ.Source = "SELECT * FROM viewTediDetail Where TID = " & Request.QueryString("TID") 
RecZ.CursorType = 0
RecZ.CursorLocation = 2
RecZ.LockType = 3
RecZ.Open()
RecZ_numRows = 0


set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM viewTediTransactions Where TID = " & Request.QueryString("TID")
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 3
RecCurrent.Open()
RecCurrent_numRows = 0


set RecTrans = Server.CreateObject("ADODB.Recordset")
RecTrans.ActiveConnection = MM_Site_STRING
RecTrans.Source = "SELECT * FROM viewTediTransactions Where TID = " & Request.QueryString("TID") & " Order by CID Desc"
'Response.write(RecTrans.Source)
RecTrans.CursorType = 0
RecTrans.CursorLocation = 2
RecTrans.LockType = 3
RecTrans.Open()
RecTrans_numRows = 0


CanUnallocate = "No"
SystemItem = "232"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
CanUnallocate = "Yes"
End If

CanDeleteTransAction = "No"
SystemItem = "233"
set RecHasPermission = Server.CreateObject("ADODB.Recordset")
RecHasPermission.ActiveConnection = MM_Site_STRING
RecHasPermission.Source = "Select * FROM ViewUserPermissions where ItemID = " & SystemItem & " and UserID = " & Session("UNID")
RecHasPermission.CursorType = 0
RecHasPermission.CursorLocation = 2
RecHasPermission.LockType = 3
RecHasPermission.Open()
RecHasPermissionr_numRows = 0
If Not RecHasPermission.EOF and Not RecHasPermission.BOF Then
CanDeleteTransAction = "Yes"
End If


Dim RecTrans__MMColParam
RecTrans__MMColParam = "1"
If (Request.QueryString("TID") <> "") Then 
  RecTrans__MMColParam = Request.QueryString("TID")
End If
%>
<%
WhichTrans = Request.QueryString("TransType")
Dim RecTrans
Dim RecTrans_cmd
Dim RecTrans_numRows

Set RecTrans_cmd = Server.CreateObject ("ADODB.Command")
RecTrans_cmd.ActiveConnection = MM_Site_STRING
RecTrans_cmd.CommandText = "SELECT * FROM viewTediTransactions WHERE TID = ? ORDER BY CDate DESC" 
RecTrans_cmd.Prepared = true
RecTrans_cmd.Parameters.Append RecTrans_cmd.CreateParameter("param1", 5, 1, -1, RecTrans__MMColParam) ' adDouble

%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
If Request.QueryString("Excel") <> "Yes" Then
Repeat1__numRows = 100
Else
Repeat1__numRows = 1000000
End If
Repeat1__index = 0
RecTrans_numRows = RecTrans_numRows + Repeat1__numRows
%>

<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim RecTrans_total
Dim RecTrans_first
Dim RecTrans_last

' set the record count
RecTrans_total = RecTrans.RecordCount

' set the number of rows displayed on this page
If (RecTrans_numRows < 0) Then
  RecTrans_numRows = RecTrans_total
Elseif (RecTrans_numRows = 0) Then
  RecTrans_numRows = 1
End If

' set the first and last displayed record
RecTrans_first = 1
RecTrans_last  = RecTrans_first + RecTrans_numRows - 1

' if we have the correct record count, check the other stats
If (RecTrans_total <> -1) Then
  If (RecTrans_first > RecTrans_total) Then
    RecTrans_first = RecTrans_total
  End If
  If (RecTrans_last > RecTrans_total) Then
    RecTrans_last = RecTrans_total
  End If
  If (RecTrans_numRows > RecTrans_total) Then
    RecTrans_numRows = RecTrans_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (RecTrans_total = -1) Then

  ' count the total records by iterating through the recordset
  RecTrans_total=0
  While (Not RecTrans.EOF)
    RecTrans_total = RecTrans_total + 1
    RecTrans.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (RecTrans.CursorType > 0) Then
    RecTrans.MoveFirst
  Else
    RecTrans.Requery
  End If

  ' set the number of rows displayed on this page
  If (RecTrans_numRows < 0 Or RecTrans_numRows > RecTrans_total) Then
    RecTrans_numRows = RecTrans_total
  End If

  ' set the first and last displayed record
  RecTrans_first = 1
  RecTrans_last = RecTrans_first + RecTrans_numRows - 1
  
  If (RecTrans_first > RecTrans_total) Then
    RecTrans_first = RecTrans_total
  End If
  If (RecTrans_last > RecTrans_total) Then
    RecTrans_last = RecTrans_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = RecTrans
MM_rsCount   = RecTrans_total
MM_size      = RecTrans_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
RecTrans_first = MM_offset + 1
RecTrans_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (RecTrans_first > MM_rsCount) Then
    RecTrans_first = MM_rsCount
  End If
  If (RecTrans_last > MM_rsCount) Then
    RecTrans_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<%If Request.QueryString("Excel") = "Yes" Then

SavePath = AppPath & "Reports/"
SaveFileName = RecZoner.Fields.Item("TediEmpCode").Value & "-Agent_MCharge_History-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Date, Transaction Type, Amount, Description"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)

End If%>
<%If Request.QueryString("Excel") <> "Yes" Then%>
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

                        <div class="eight columns"><h1>Agent M-Charge History: <%=(RecZoner.Fields.Item("TediEmpCode").Value)%></h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<%End If%>
<% If Not RecTrans.EOF Or Not RecTrans.BOF Then %>
<%If Request.QueryString("Excel") <> "Yes" Then%>        <p>&nbsp;
<br><br><br><b><%=TransHeading%></b> Transaction <strong><%=(RecTrans_first)%></strong> to <strong><%=(RecTrans_last)%></strong> of <strong><%=(RecTrans_total)%></strong> </p>

<%End If%>
        <table >
<%If Request.QueryString("Excel") <> "Yes" Then%> 
<thead>
          <tr>
            <th colspan="2">Date</th>
            <th>Type</th>
            <th>Amount</th>
            <th>Description</th>
	    <th>&nbsp;</th>
	    <th>&nbsp;</th>
          </tr>
</thead>
<%End If%>
          <% TCount = RecTrans_first
While ((Repeat1__numRows <> 0) AND (NOT RecTrans.EOF)) 
If RecTrans.Fields.Item("TID").Value = "2" Then

			  TV = "-" & RecTrans.Fields.Item("CAmount").Value
			  Else
			  TV = FormatNumber(RecTrans.Fields.Item("CAmount").Value,,,,0)
			  End If
TV = Replace(TV, "--", "")
If RecTrans.Fields.Item("CComments").Value <> "" Then
TC = RecTrans.Fields.Item("CComments").Value
Else
TC = "N/A"
End If
RowLineCol = "offtabGreen"
If TV > 0 Then
RowLineCol = "offtab"
End If
IsARefund = instr (1,TC, "Manual Refund", 1) 
if IsARefund > 0 then
RowLineCol = "offtabYellow"
End If

IsRetroCapture = instr (1,TC, "Retrospectively Captured", 1) 
if IsRetroCapture > 0 then
RowLineCol = "offtabblue"
End If
%><%If Request.QueryString("Excel") <> "Yes" Then%>
            <tr>
              <td><%=TCount%>.</td>
              <td><%=Day(RecTrans.Fields.Item("CDate").Value)%>&nbsp;<%=MonthName(Month(RecTrans.Fields.Item("CDate").Value),true)%>&nbsp;<%=Year(RecTrans.Fields.Item("CDate").Value)%></td>
              <td><%=(RecTrans.Fields.Item("TName").Value)%></td>
              <td>R <%=TV%></td>
              <td><%=TC%></td>
<%
If CanUnallocate = "Yes" then
%>
                                <td class="action-td"><a href="UnallocateTrans.asp?CID=<%=RecTrans.Fields.Item("CID").Value%>&TID=<%=Request.QueryString("TID")%>" class="clone-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>
<%
If CanDeleteTransAction = "Yes" then
%>
                                <td class="action-td"><a href="DeleteFNBTrans.asp?CID=<%=RecTrans.Fields.Item("CID").Value%>&TID=<%=Request.QueryString("TID")%>" class="delete-button"></a></td>
<%Else%><td>&nbsp;</td>
<%End If
%>
		
</tr>
<%Else
TV = Replace(TV, ",", ".")
TheFile.Writeline(Day(RecTrans.Fields.Item("CDate").Value) & " " & MonthName(Month(RecTrans.Fields.Item("CDate").Value),true) & " " & Year(RecTrans.Fields.Item("CDate").Value) & "," & RecTrans.Fields.Item("TName").Value & "," & TV & "," & TC)
%>
<%End If%>

            
            <% TCount = TCount + 1
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  RecTrans.MoveNext()
Wend
%>
        </table>
        <br>
<%If Request.QueryString("Excel") <> "Yes" Then%>
        <table border="0">
          <tr>
            <% If MM_offset <> 0 Then %><td width="50" class="offtab"><a href="<%=MM_moveFirst%>">First</a></td><% End If ' end MM_offset <> 0 %>                               
            <% If MM_offset <> 0 Then %><td width="50" class="offtab"><a href="<%=MM_movePrev%>">Previous</a></td><% End If ' end MM_offset <> 0 %>
            <% If Not MM_atTotal Then %><td width="50" class="offtab"><a href="<%=MM_moveNext%>">Next</a></td><% End If ' end Not MM_atTotal %>    
            <% If Not MM_atTotal Then %><td width="50" class="offtab"><a href="<%=MM_moveLast%>">Last</a></td><% End If ' end Not MM_atTotal %>
          </tr>
        </table>

<center><a href="TediMChargeHistoryFull.asp?TID=<%=Request.QueryString("TID")%>&Transtype=<%=Request.QueryString("Transtype")%>&Excel=Yes" class="nice red radius button">Export To Excel</a></center>
<%End If%>
        <% End If ' end Not RecTrans.EOF Or NOT RecTrans.BOF %>
<% If RecTrans.EOF And RecTrans.BOF Then %>
          <p><strong>No Transaction For This Agent</strong></p>
          <% End If ' end RecTrans.EOF And RecTrans.BOF %>
<%If Request.QueryString("Excel") <> "Yes" Then%> 
                    </div>
  
<!-- #include file="includes/footer.asp" -->
<%
Else
response.redirect("Reports/" & SaveFileName)
End If
%>

