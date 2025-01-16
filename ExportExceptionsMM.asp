<!-- #include file="Connections/Site.asp" -->
<%
OB = "FNBDate"
SavePath = AppPath & "Exports/"
SaveFileName = "MobileMoney_Exceptions-" & Day(Now) & Month(Now) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".csv"
TableHead = "Date,Description,Amount,Account Balance"
TheFilePath=(SavePath & SaveFileName)
Set FSO = Server.CreateObject("scripting.FileSystemObject")
Set TheFile = FSO.CreateTextFile(TheFilePath, True)
TheFile.Writeline(TableHead)

set RecTrans = Server.CreateObject("ADODB.Recordset")
RecTrans.ActiveConnection = MM_Site_STRING
RecTrans.Source = "SELECT * FROM View_Opto_UnallocatedFNBTrans_SimpleMM Order by " & OB & " DESC"
RecTrans.CursorType = 0
RecTrans.CursorLocation = 2
RecTrans.LockType = 3
RecTrans.Open()
RecTrans_numRows = 0
While Not RecTrans.EOF
TheFile.Writeline(Day(RecTrans.Fields.Item("FNBDate").Value) & " " & MonthName(Month(RecTrans.Fields.Item("FNBDate").Value),true) & " " & Year(RecTrans.Fields.Item("FNBDate").Value) & "," & RecTrans.Fields.Item("TransDescription").Value & "," & Replace(RecTrans.Fields.Item("TransAmount").Value, ",", ".") & "," & Replace(RecTrans.Fields.Item("AccountBalance").Value, ",", "."))
RecTrans.MoveNext
Wend

response.redirect("Exports/" & SaveFileName)
%>