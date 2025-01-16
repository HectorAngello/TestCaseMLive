<!-- #include file="Connections/Site.asp" -->
<%
UnknownFileName = "UnknownVendsMSISDNs_" & Session("UNID") & "_" & Day(Now) & MonthName(Month(Now),true) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now)
TheFilePath=(AppPath & "Exports\" & UnknownFileName & ".csv")
'response.write(TheFilePath)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
'************ beginning of the file body ***********
set RecUnknownMSISDNs = Server.CreateObject("ADODB.Recordset")
RecUnknownMSISDNs.ActiveConnection = MM_Site_STRING
RecUnknownMSISDNs.Source = "SELECT DISTINCT OriginalMSISDN FROM Vends WHERE (TID = 0) Order By OriginalMSISDN Asc"
RecUnknownMSISDNs.CursorType = 0
RecUnknownMSISDNs.CursorLocation = 2
RecUnknownMSISDNs.LockType = 3
RecUnknownMSISDNs.Open()
RecUnknownMSISDNs_numRows = 0
While Not RecUnknownMSISDNs.EOF
TheFile.Writeline(RecUnknownMSISDNs.Fields.Item("OriginalMSISDN").Value)
RecUnknownMSISDNs.MoveNext
Wend
'************ end of the file body ***********
TheFile.close
Set FSO = nothing
Response.redirect("Exports/" & UnknownFileName & ".csv")
%>