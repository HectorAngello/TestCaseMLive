<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
SMSCount = 0
UpPath =  AppPath & "UploadedFiles/"
SID = Request.QueryString("UNID")
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
NewName = "SMSSendUpload-" & SID & "-" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & "--" & Hour(Now) & "-" & Minute(Now) & ".csv"

OrgPath = UpPath & File_Uploaded
NewPath = UpPath & NewName

If File_Uploaded <> NewName Then
set fs=Server.CreateObject("Scripting.FileSystemObject")
fs.MoveFile OrgPath,NewPath
set fs=nothing
End If 

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

strFileName = NewPath
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
fsoForReading = 1
Set objTextStream = objFSO.OpenTextFile(strFileName, fsoForReading)
Do while not objTextStream.AtEndOfStream

RFILE = objTextStream.ReadLine
RFILE = Replace(RFILE, Chr(34), "")
RFILE = Replace(RFILE, Chr(39), "`")

Curlength = len(RFILE)
Comma1 = Instr(1, CStr(RFILE), ",,")
MSISDN = mid(RFILE, 1, (Comma1 - 1))
RFILE = mid(RFILE, (Comma1 + 1), Curlength)
MSISDN = Replace(MSISDN, " ", "")
MSISDN = Replace(MSISDN, "_", "")

SMSMSG = RFILE
SMSMSGT = Len(SMSMSG)
SMSMSG = Right(SMSMSG, SMSMSGT - 1)

TID = 0

Set conMain = Server.CreateObject ( "ADODB.Connection" )
set RecZoner = Server.CreateObject("ADODB.Recordset")
RecZoner.ActiveConnection = MM_Site_STRING
RecZoner.Source = "SELECT Top(1) * FROM Tedis Where Right(TediCell,9) = '" & right(MSISDN,9) & "'"
response.write(RecZoner.Source)
RecZoner.CursorType = 0
RecZoner.CursorLocation = 2
RecZoner.LockType = 3
RecZoner.Open()
RecZoner_numRows = 0
If Not RecZoner.EOF and Not RecZoner.BOF Then
TID = RecZoner.Fields.Item("TID").Value
End If

Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecADDSMS = Server.CreateObject ( "ADODB.Recordset" )
RecADDSMS.Open "SELECT Top(1) * FROM SMSCommunications", MM_Site_STRINGWrite, 1, 2
RecADDSMS.AddNew
RecADDSMS("UserType") = "1"
RecADDSMS("AlloID") = TID
RecADDSMS("SMSMSG") = SMSMSG
RecADDSMS("MobileNo") = MSISDN
RecADDSMS("SMSDate") = Now()
RecADDSMS("IsSent") = "False"
RecADDSMS.Update
RecADDSMS.Close

SMSCount = SMSCount + 1

Loop
objTextStream.Close

Response.redirect("Updated.asp?AppCat=16&AppSubCatID=32&SMSCount=" & SMSCount)
%>