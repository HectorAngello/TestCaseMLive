<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If


set RecTrans = Server.CreateObject("ADODB.Recordset")
RecTrans.ActiveConnection = MM_Site_STRING
RecTrans.Source = "SELECT * FROM ViewBulkTransactionHistoryMM Where BulkID = " & Request.QueryString("BulkID") & " Order by BulkDate DESC"
RecTrans.CursorType = 0
RecTrans.CursorLocation = 2
RecTrans.LockType = 3
RecTrans.Open()
RecTrans_numRows = 0

set RecTrans2 = Server.CreateObject("ADODB.Recordset")
RecTrans2.ActiveConnection = MM_Site_STRING
RecTrans2.Source = "SELECT * FROM BulkMChargeMM Where BulkID = " & Request.QueryString("BulkID")
RecTrans2.CursorType = 0
RecTrans2.CursorLocation = 2
RecTrans2.LockType = 3
RecTrans2.Open()
RecTrans2_numRows = 0

%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 100
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

                        <div class="eight columns"><h1>Viewing M-Charge Bulk File</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
  <p><br><br><br><b>MCharge Bulk Export for : <a href="MChargeBulkFiles/<%=(RecTrans2.Fields.Item("FileName").Value)%>.txt"><%=(RecTrans2.Fields.Item("FileName").Value)%>.txt</a></b></p>
<% If Not RecTrans.EOF Or Not RecTrans.BOF Then %>
        Transaction <strong><%=(RecTrans_first)%></strong> to <strong><%=(RecTrans_last)%></strong> of <strong><%=(RecTrans_total)%></strong> </p>
        <table border="0" cellspacing="2" cellpadding="2">
	<thead>
          <tr>
            <th colspan="2">Agent</th>
            <th>Agent MSISDN</th>
            <th>Credit Before</th>
            <th>Credit After</th>
            <th>Credit Amount</th>
          </tr>
	</thead>
          <% TCount = RecTrans_first
TotMCharge = 0
While ((Repeat1__numRows <> 0) AND (NOT RecTrans.EOF)) 
%>
            <tr>
              <td colspan="2"><%=TCount%>. <%=(RecTrans.Fields.Item("TediFirstName").Value)%>&nbsp;<%=(RecTrans.Fields.Item("TediLastName").Value)%> (<%=(RecTrans.Fields.Item("TediEmpCode").Value)%>)</td>
              <td><%=(RecTrans.Fields.Item("TediCell").Value)%></td>
              <td>R <%=(RecTrans.Fields.Item("ValBefore").Value)%></td>
              <td>R <%=(RecTrans.Fields.Item("ValAfter").Value)%></td>
              <td>R <%=(RecTrans.Fields.Item("MChargeAmount").Value)%></td>
</tr>
            <% ToTMCharge = TotMCharge + RecTrans.Fields.Item("MChargeAmount").Value
 TCount = TCount + 1
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  RecTrans.MoveNext()
Wend
%>
        </table>
        
        <table border="0">
          <tr>
            <% If MM_offset <> 0 Then %><td width="50" class="offtab"><a href="<%=MM_moveFirst%>">First</a></td><% End If ' end MM_offset <> 0 %>                               
            <% If MM_offset <> 0 Then %><td width="50" class="offtab"><a href="<%=MM_movePrev%>">Previous</a></td><% End If ' end MM_offset <> 0 %>
            <% If Not MM_atTotal Then %><td width="50" class="offtab"><a href="<%=MM_moveNext%>">Next</a></td><% End If ' end Not MM_atTotal %>    
            <% If Not MM_atTotal Then %><td width="50" class="offtab"><a href="<%=MM_moveLast%>">Last</a></td><% End If ' end Not MM_atTotal %>
          </tr>
        </table>
        <% End If ' end Not RecTrans.EOF Or NOT RecTrans.BOF %>
<% If RecTrans.EOF And RecTrans.BOF Then %>
          <p><strong>No Tedis Were Selected During The Export Process</strong></p>
          <% End If ' end RecTrans.EOF And RecTrans.BOF %>

                    </div>
         
<!-- #include file="includes/footer.asp" -->

