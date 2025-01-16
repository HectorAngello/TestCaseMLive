<!-- #include file="includes/header.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_Site_STRINGWrite
  MM_editTable = "Users"
  MM_editColumn = "UserID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "Dashboard.asp"
  MM_fieldsStr  = "Email|value|Password|value|MobileNumber|value|UserFirstName|value|UserLastName|value"
  MM_columnsStr = "[UEmail]|',none,''|[Password]|',none,''|[CellNo]|',none,''|UserFirstName|',none,''|UserLastName|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

' ----------------- Logging Starts ----------------- 
UD = "New Users Name: " & Request.Form("Username") & " | New Password: " & Request.Form("Password") & " | Rec Change Id: " & Request.Form("MM_recordId")
Set conMain = Server.CreateObject ( "ADODB.Connection" )
Set rstChangeUpdate = Server.CreateObject ( "ADODB.Recordset" )
rstChangeUpdate.Open "SELECT * FROM ChangeLog", MM_Site_STRINGWrite, 1, 2
rstChangeUpdate.AddNew
rstChangeUpdate("ChangeType") = "User Changed Their Log In Credentials"
rstChangeUpdate("ChangeBy") = Session("userUN") & " (SID=" & Session("UNID") & ")"
rstChangeUpdate("Changes") = UD
rstChangeUpdate.Update
rstChangeUpdate.Close
set rstChangeUpdate = nothing	
' ----------------- Logging Ends ----------------- 

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_Site_STRING
Recordset1.Source = "SELECT * FROM Users Where UserID = " & Session("UNID")
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
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
                    <div class="row heading">
                        <div class="eight columns"><h1>Update User Log In:</h1></div>
                        
                    </div>
                    
                        
<form method="POST" action="<%=MM_editAction%>" name="form1"  onSubmit="MM_validateForm('Email','','RisEmail','Password','','R','FullName','','R','IDNumber','','RisNum');return document.MM_returnValue" >
                       <table> <tbody>
                          
                  <tr> 
            <td>Username:</td>
            <td><%=(Recordset1.Fields.Item("Username").Value)%></td>
	          </tr>
        	  <tr> 
            <td class="offtab">Password:</td>
            <td><input type="text" name="Password" value="<%=(Recordset1.Fields.Item("Password").Value)%>" size="32"></td>
          	  </tr>
        	  <tr> 
            <td class="offtab">EmailAddress:</td>
            <td><input type="text" name="Email" value="<%=(Recordset1.Fields.Item("UEmail").Value)%>" size="32"></td>
          	  </tr>
          	        	  <tr> 
            <td class="offtab">Cell Number:</td>
            <td><input type="text" name="MobileNumber" value="<%=(Recordset1.Fields.Item("CellNo").Value)%>" size="32"></td>
          	  </tr>
        	  <tr> 
            <td class="offtab">First Name:</td>
            <td><input type="text" name="UserFirstName" value="<%=(Recordset1.Fields.Item("UserFirstName").Value)%>" size="32"></td>
          	  </tr>
        	  <tr> 
            <td class="offtab">Last Name:</td>
            <td><input type="text" name="UserLastName" value="<%=(Recordset1.Fields.Item("UserLastName").Value)%>" size="32"></td>
          	  </tr>

        </table>
        <p><input type="Submit" class="red nice button radius" value="Update User Account"></p>
        	
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("UserID").Value %>">

                    
      </form>
<!-- #include file="includes/footer.asp" -->

