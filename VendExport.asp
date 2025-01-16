<!-- #include file="Connections/Site.asp" -->
<%
TID = Request.QueryString("TID")
ReportMonth = Request.QueryString("M")
ReportYear = Request.QueryString("Y")
EmpCode = Request.QueryString("EmpCode")

SavePath = AppPath & "Exports/"
SaveFileName = "Agent_VendingHistory-" & EmpCode & "-" & MonthName(ReportMonth) & "_" & ReportYear & "-GenDate-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Date, Amount, Destination MSISDN, Vend type"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)

set RecVendList = Server.CreateObject("ADODB.Recordset")
RecVendList.ActiveConnection = MM_Site_STRING
RecVendList.Source = "SELECT * From ViewVendingDetails WHERE CalMonth = " & ReportMonth & " and CalYear = " & ReportYear & " and TID = " & TID & " order by Venddate"
'Response.write(RecVendList.Source)
RecVendList.CursorType = 0
RecVendList.CursorLocation = 2
RecVendList.LockType = 3
RecVendList.Open()
RecVendList_numRows = 0
While Not RecVendList.EOF
TheFile.Writeline(Day(RecVendList.Fields.Item("VendDate").Value) & " " & MonthName(Month(RecVendList.Fields.Item("VendDate").Value),true) & " " & Year(RecVendList.Fields.Item("VendDate").Value) & "," & RecVendList.Fields.Item("VendAmount").Value & "," & RecVendList.Fields.Item("DestMSISDN").Value & "," & RecVendList.Fields.Item("VendNameType").Value)
RecVendList.MoveNext
Wend

response.redirect("Exports/" & SaveFileName)

%>