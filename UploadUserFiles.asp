<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Site.asp" -->
<%
UserID = Request.QueryString("UserID")
UserType = Request.QueryString("UserType")

UpPath =  AppPath & "UploadedFiles\"
UpCount = 0
Dim X_Path, Y, File_Name, P, S, T, U, CompanyName
Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False 
Upload.ProgressID = Request.QueryString("PID")
Upload.OverwriteFiles = False
Count = Upload.Save(UpPath)
For Each File in Upload.Files
UpCount = UpCount + 1
Uploaded_Full_Path =  File.Path 
S = InStrRev(Uploaded_Full_Path,"\") + 1
File_Uploaded = mid(Uploaded_Full_Path,S)
Dim a, b
a = File.Size 
b = a / 1024

NewFName = File_Uploaded

Curlength = len(NewFName)
Comma1 = Instr(1, CStr(NewFName), ".")
TempName = mid(NewFName, 1, (Comma1 - 1))
RFILE = mid(NewFName, (Comma1 + 1), Curlength)

NewFName = TempName
NewFName = Replace(NewFName, " ", "_")
NewFName = Replace(NewFName, "&", "")
NewFName = Replace(NewFName, "%20", "_")
NewFName = Replace(NewFName, "%", "")

NewFName = NewFName & "-" & Day(Now) & MonthName(Month(Now),true) & Year(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & "." & RFILE

For Each Item in Upload.Form
GGCat = "CatID" & UpCount
If Item.Name = GGCat Then
CatIDXX = Item.Value
End If
Next


Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set RecUpdateZTTable = Server.CreateObject ( "ADODB.Recordset" )
RecUpdateZTTable.Open "SELECT Top (2) * FROM UserUploadedFiles", MM_Site_STRINGWrite, 1, 2
RecUpdateZTTable.AddNew
RecUpdateZTTable("UserID") = UserID
RecUpdateZTTable("UserType") = UserType
RecUpdateZTTable("FileName") = NewFName
RecUpdateZTTable("FileUploadCat") = CatIDXX
RecUpdateZTTable("FileActive") = "True"
RecUpdateZTTable("AddedDate") = Now()
RecUpdateZTTable("AddedBy") = Session("UNID")
RecUpdateZTTable("FileSize") = b
RecUpdateZTTable.Update
RecUpdateZTTable.Close



If File_Uploaded <> NewFName Then
set fs=Server.CreateObject("Scripting.FileSystemObject")
fs.MoveFile (UpPath & File_Uploaded), (UpPath & NewFName)
Set fs=nothing
End iF

Next

If UserType = "1" Then
Response.Redirect("TediView.asp?Item=4&TID=" & UserID)
End If
If UserType = "2" Then
Response.Redirect("ASView.asp?Item=4&ASID=" & UserID)
End If


%>