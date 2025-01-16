<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
UpPath =  AppPath & "MChargeFiles\"
SID = Session("UNID")
LineItemsImported = 0
Dim X_Path, Y, File_Name, P, S, T, U, CompanyName
Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False 
Upload.ProgressID = Request.QueryString("PID")
Upload.OverwriteFiles = False
Count = Upload.Save(UpPath)
For Each File in Upload.Files
Uploaded_Full_Path =  File.Path 
S = InStrRev(Uploaded_Full_Path,"\") + 1
File_Uploaded = mid(Uploaded_Full_Path,S)
Dim a, b
a = File.Size 
b = a / 1024
If Right(LCASE(File_Uploaded),3) = "csv" Then
NewName = "MMFNBUpload-" & SID & "-" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & "--" & Hour(Now) & "-" & Minute(Now) & ".csv"

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1) * FROM ManualFNBFileUploads", MM_Site_STRINGWrite, 1, 2
rstSecond.AddNew
rstSecond("SID") = SID
rstSecond("OriginalFileName") = File_Uploaded
rstSecond("NewFileName") = NewName
rstSecond("FileSize") = b
rstSecond("Uploadprogress") = "File Uploaded"
rstSecond("LineItemsImported") = "0"
rstSecond("ImportType") = "Unknown"
rstSecond("ImportDateTime") = Now()
rstSecond("ImportOutCome") = "Incomplete"
rstSecond.Update
rstSecond.Close
set rstSecond = nothing	

OrgPath = UpPath & File_Uploaded
NewPath = UpPath & NewName

If File_Uploaded <> NewName Then
set fs=Server.CreateObject("Scripting.FileSystemObject")
fs.MoveFile OrgPath,NewPath
set fs=nothing
End If 

set RecNewestID = Server.CreateObject("ADODB.Recordset")
RecNewestID.ActiveConnection = MM_Site_STRING
RecNewestID.Source = "SELECT Top(1) * FROM ManualFNBFileUploads Where SID = " & SID & " Order By ID Desc"
RecNewestID.CursorType = 0
RecNewestID.CursorLocation = 2
RecNewestID.LockType = 3
RecNewestID.Open()
RecNewestID_numRows = 0
NewestID = RecNewestID.Fields.Item("ID").Value
Else
DelFilePath = UpPath & File_Uploaded
Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists(DelFilePath) then
  fs.DeleteFile(DelFilePath)
end if
set fs=nothing

%>
	<script language="JavaScript" type="text/JavaScript">
	<!--
	  alert("Error - Uploaded File Is Not A .csv File !!!");
	  history.go(-1);
	//-->
	</script>
<%
Response.end
End If
Next


ImportType = "Mobile Money"

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ManualFNBFileUploads Where ID = "& NewestID , MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Uploadprogress") = "File Identified as: " & ImportType
rstSecond("ImportType") = ImportType
rstSecond.Update
rstSecond.Close
set rstSecond = nothing


' OK . . . So we know the type, so lets pull the data in


'Response.end

strFileName = NewPath
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
fsoForReading = 1
Set objTextStream = objFSO.OpenTextFile(strFileName, fsoForReading)
Do while not objTextStream.AtEndOfStream

LCount = LCount + 1
RFILE = objTextStream.ReadLine
RFILE = Replace(RFILE, Chr(34), "")
RFILE = Replace(RFILE, Chr(39), "`")
IsInLine = instr (1,RFILE, "ACCOUNT TRANSACTION HISTORY", 1) 
if IsInLine > 0 then
Else
IsInLine2 = instr (1,RFILE, "FOR ACCOUNT NUMBER", 1) 
if IsInLine2 > 0 then
Else
IsInLine3 = instr (1,RFILE, "AMOUNT,REFERENCE,CHEQUE", 1) 
if IsInLine3 > 0 then
Else
IsInLine4 = instr (1,RFILE, "CHEQUE NUMBER,", 1) 
if IsInLine4 > 0 then
Else
IsInLine5 = instr (1,RFILE, ",,,", 1) 
if IsInLine5 > 0 then
Else
%>
<%=LCount%>. Original Line: <%=RFILE%><br>
<%

Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
EDateTemp = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)

if Right(EDateTemp,4)= cstr(Year(Now)) then
EdateYear = Right(EDateTemp,4)
EDateTemp = Replace(EDateTemp, "/" & EdateYear , "")
EdateMonth = Right(EDateTemp, 2)
EDateTemp = Replace(EDateTemp, "/" & EdateMonth , "")
EDateDay = EDateTemp
EDate = EDateDay & " " & MonthName(EdateMonth) & " " & EdateYear
Else
EDate = EDateTemp
End If
%>
EDate: <%=EDate%><br>
<%
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
ServiceFee = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
'ServiceFee = replace(ServiceFee, ".", ",")
%>
ServiceFee: <%=ServiceFee%><br>
<%
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
Amount = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
Amount2 = Amount
'Amount = replace(Amount, ".", ",")
'Amount = FormatNumber(Amount,,,,0)
'Amount = replace(Amount, ".", ",")
%>
Amount: <%=Amount%><br>
<%
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
Desc = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
If Left(Desc, 1) = " " Then
SerLen = Len(Desc)
Desc = Right(Desc, SerLen - 1)
End If
%>
Desc: <%=Desc%><br>
<%
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
ChequeNo = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
If Left(ChequeNo, 1) = " " Then
SerLen = Len(ChequeNo)
ChequeNo = Right(ChequeNo, SerLen - 1)
End If
%>
ChequeNo: <%=ChequeNo%><br>
<%
Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), Chr(44))
'Bal = mid(RFILE, 1, (Comma1 - 1))
Bal = RFILE
'RFILE = mid(RFILE, (Comma1 + 1), Curlength)

Bal = replace(Bal, ",", "")
Bal2 = Bal
Bal = replace(Bal, ".", ",")
%>
Bal: <%=Bal%><br>

<hr>
<%
IsChecqueDeposit = instr (1,Desc, "CHEQUE", 1) 
if IsChecqueDeposit > 0 then
Else
DoInsert = "Yes"
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_Site_STRINGWrite

Recordset1.Source = "SELECT Top(1)* FROM MChargeFNBTransMM where Day(FNBDate) = '" & Day(EDate) & "' and Month(FNBDate) = '" & Month(EDate) & "' and Year(FNBDate) = '" &Year(EDate) & "' and TransAmount = '" & Amount2 & "' and AccountBalance = '" + Replace(Bal, "'", "''") + "' and TransDescription = '" & Desc & "'"
'Recordset1.Source = "SELECT Top(1)* FROM MChargeFNBTransMM where FNBDate = '" & EDate & "' and TransAmount = '" + Replace(Amount, "'", "''") + "' and AccountBalance = '" + Replace(Bal, "'", "''") + "' and TransDescription = '" & Desc & "'"

'Response.Write(Recordset1.Source)
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
While Not Recordset1.EOF
DoInsert = "No"
Recordset1.MoveNext
Wend


If DoInsert = "Yes" Then
If Bal <> "0,00" Then
		Set conMain = Server.CreateObject ( "ADODB.Connection" )
		Set RecInsert = Server.CreateObject ( "ADODB.Recordset" )
		RecInsert.Open "SELECT Top(1) * FROM MChargeFNBTransMM", MM_Site_STRINGWrite, 1, 2
		RecInsert.AddNew
		RecInsert("FNBDate") = EDate
		RecInsert("ServiceFee") = ServiceFee
		RecInsert("TransAmount") = Amount
		RecInsert("TransDescription") = Desc
		RecInsert("TransChequeNo") = ChequeNo
		RecInsert("AccountBalance") = Bal
		RecInsert("Allocated") = "False"
		RecInsert("TediID") = "0"
		RecInsert.Update
		RecInsert.Close
LineItemsImported = LineItemsImported + 1
End If
'End If
End If
End If
End If
End If
End If
End If
X = X + 1
End If
Response.flush
Loop
objTextStream.Close

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstSecond = Server.CreateObject ( "ADODB.Recordset" )
rstSecond.Open "SELECT Top(1)* FROM ManualFNBFileUploads Where ID = " & NewestID , MM_Site_STRINGWrite, 1, 2
rstSecond.Update
rstSecond("Uploadprogress") = "File Imported Into MTN Live"
rstSecond("ImportOutCome") = "Successful"
rstSecond("LineItemsImported") = LineItemsImported
rstSecond.Update
rstSecond.Close
set rstSecond = nothing

RunMe = "Yes"

If RunMe = "Yes" Then
%>
<script type="text/javascript">
<!--
function delayer(){
	
	window.location = "TryMM.asp"

}
//-->
</script>
<body onLoad="setTimeout('delayer()', 100)">
Done</body>
<%
End If
%>