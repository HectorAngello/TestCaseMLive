<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If
%>
<%
Dim RecCurrent
Dim RecCurrent_numRows

Set RecCurrent = Server.CreateObject("ADODB.Recordset")
RecCurrent.ActiveConnection = MM_Site_STRING
RecCurrent.Source = "SELECT * FROM Users WHERE UserID = " & Request.QueryString("UserID")
RecCurrent.CursorType = 0
RecCurrent.CursorLocation = 2
RecCurrent.LockType = 1
RecCurrent.Open()

RecCurrent_numRows = 0
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

                        <div class="eight columns"><h1>Security</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br><h4>Bulk Allocate Regions To User: <%=(RecCurrent.Fields.Item("UserFirstName").Value)%>&nbsp;<%=(RecCurrent.Fields.Item("UserLastName").Value)%></h4>
<%
set RecQuickJump = Server.CreateObject("ADODB.Recordset")
RecQuickJump.ActiveConnection = MM_Site_STRING
RecQuickJump.Source = "SELECT Distinct RegionName, RID FROM Regions where  Active = 'Yes' and CompanyID = " & Session("CompanyID") & " Order By RegionName Asc"
RecQuickJump.CursorType = 0
RecQuickJump.CursorLocation = 2
RecQuickJump.LockType = 3
RecQuickJump.Open()
RecQuickJump_numRows = 0
%>
<form name="form1" method="get" action="Telephone2.asp">
Select Region:<br>
             
<select name="menu2" onChange="MM_jumpMenu2('parent',this,0)" Class="text3_frm">
<Option selected>Select Region</Option>
                <%
While (NOT RecQuickJump.EOF)
%>
                <option value="BulkRegionAllo.asp?UserID=<%=Request.QueryString("UserID")%>&RID=<%=(RecQuickJump.Fields.Item("RID").Value)%>" <%If Cstr(RecQuickJump.Fields.Item("RID").Value) = Request.QueryString("RID") Then%>Selected<%End If%>><%=(RecQuickJump.Fields.Item("RegionName").Value)%></option>
                <%
  RecQuickJump.MoveNext()
Wend
If (RecQuickJump.CursorType > 0) Then
  RecQuickJump.MoveFirst
Else
  RecQuickJump.Requery
End If
%>
              </select>

        </form>
<%
RecQuickJump.Close()
If Request.QueryString("RID") <> "" Then
set RecMainRegion = Server.CreateObject("ADODB.Recordset")
RecMainRegion.ActiveConnection = MM_Site_STRING
RecMainRegion.Source = "SELECT * FROM Regions where Active = 'Yes' and RID = " & Request.QueryString("RID")
RecMainRegion.CursorType = 0
RecMainRegion.CursorLocation = 2
RecMainRegion.LockType = 3
RecMainRegion.Open()
RecMainRegion_numRows = 0
%><h4><%=(RecMainRegion.Fields.Item("RegionName").Value)%></h4>
<p>Selecting None, will still keep the existing allocations selected.</p>
<form Name="UpdateRegions" Method="Post" Action="BulkRegionAllo2.asp">
                    <table>
                        <thead>
                            <tr>
                                <th>Region</th>
                                <th>Sub Region</th>
				<th>Mentors With Agents In Sub Region</th>
                                <th><%If Request.QueryString("All") = "" Then%><a href="BulkRegionallo.asp?UserID=<%=Request.QueryString("UserID")%>&RID=<%=Request.QueryString("RID")%>&All=True">All</a><%End If%><%If Request.QueryString("All") = "True" Then%><a href="BulkRegionallo.asp?UserID=<%=Request.QueryString("UserID")%>&RID=<%=Request.QueryString("RID")%>">None</a><%End If%></th>
                            </tr>
                        </thead>
                        <tbody>
<%
set RecSubRegions = Server.CreateObject("ADODB.Recordset")
RecSubRegions.ActiveConnection = MM_Site_STRING
RecSubRegions.Source = "SELECT * FROM SubRegions where SubRegionActive = 'True' and RID = " & Request.QueryString("RID") & "   Order By SubRegionName Asc"
RecSubRegions.CursorType = 0
RecSubRegions.CursorLocation = 2
RecSubRegions.LockType = 3
RecSubRegions.Open()
RecSubRegions_numRows = 0
While Not RecSubRegions.EOF
ISChecked = ""

set RecIsChecked = Server.CreateObject("ADODB.Recordset")
RecIsChecked.ActiveConnection = MM_Site_STRING
RecIsChecked.Source = "SELECT * FROM UserRegion where SRID = " & RecSubRegions.Fields.Item("SRID").Value & " and UserID = " & Request.QueryString("UserID")
RecIsChecked.CursorType = 0
RecIsChecked.CursorLocation = 2
RecIsChecked.LockType = 3
RecIsChecked.Open()
RecIsChecked_numRows = 0
If Not RecIsChecked.EOF and Not RecIsChecked.BOF Then
ISChecked = "Checked"
End If

If Request.QueryString("All") = "True" Then
ISChecked = "Checked"
End If

MentorList = ""
SRID = RecSubRegions.Fields.Item("SRID").Value
set RecMentors = Server.CreateObject("ADODB.Recordset")
RecMentors.ActiveConnection = MM_Site_STRING
RecMentors.Source = "SELECT DISTINCT ASFirstName, ASLastName FROM ViewTediDetail WHERE (TediActive = 'True') AND (SRID = " & SRID & ") ORDER BY ASFirstName Asc"
RecMentors.CursorType = 0
RecMentors.CursorLocation = 2
RecMentors.LockType = 3
RecMentors.Open()
RecMentors_numRows = 0
While Not RecMentors.EOF
MentorList = MentorList  & RecMentors.Fields.Item("ASFirstName").Value & " " & RecMentors.Fields.Item("ASLastName").Value & ", "
RecMentors.MoveNext
Wend

If MentorList <> "" Then
MentorListT = Len(MentorList)
MentorList = Left(MentorList, MentorListT - 2)
End If

%>
                            <tr>
                                <td nowrap><%=(RecMainRegion.Fields.Item("RegionName").Value)%></td>
                                <td nowra><%=(RecSubRegions.Fields.Item("SubRegionName").Value)%></td>
				<td><%=MentorList%></td>
                                <td><input Name="SRID<%=(RecSubRegions.Fields.Item("SRID").Value)%>" type="checkbox" Value="Yes" <%=ISChecked%>></td>
                            </tr>
<%
Response.flush
RecSubRegions.MoveNext
Wend
%>
                        </tbody>
                    </table>
<input type="Submit" class="orange nice button radius" value="Allocate">
<input type="Hidden" Name="UserID" Value="<%=(RecCurrent.Fields.Item("UserID").Value)%>">
<input type="Hidden" Name="RID" Value="<%=(Request.QueryString("RID"))%>">
<%End If%>
                    </div>
<!-- #include file="includes/footer.asp" -->

