<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecEdit = Server.CreateObject("ADODB.Recordset")
RecEdit.ActiveConnection = MM_Site_STRING
RecEdit.Source = "SELECT Top(1)* FROM ViewASDetail where CompanyID = " & Session("CompanyID") & " and ASID = " & Request.QueryString("ASID")
RecEdit.CursorType = 0
RecEdit.CursorLocation = 2
RecEdit.LockType = 3
RecEdit.Open()
RecEdit_numRows = 0

UType = 1
UserID = Request.QueryString("ASID")

MentorType = SupervisorLabel

%>
<!-- header -->
    <!-- #include file="includes/topheader.inc" -->
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            <div class="three columns">
                <!-- #include file="Includes/sidebar.asp" -->
		<!-- #include file="Includes/EDIsidebar.asp" -->
            </div>
            <div class="nine columns">
<%If Request.QueryString("TediUpdated") = "True" Then%><div class="alert-box success">Agent Updated In The System.</div><%End If%>
                <div class="content panel">

                        <div class="eight columns"><h1>Mentor: <%=RecEdit.Fields.Item("ASEmpCode").Value%></h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br>


                                

<div class="row">
<div class="six columns">
        <span class="spacer-120">First Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("ASFirstName").Value%></label><br>
        <span class="spacer-120">Last Name:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("ASLastName").Value%></label><br>
        <span class="spacer-120">Email:</span> <label for="agentCell"><%=RecEdit.Fields.Item("ASEmail").Value%></label><br>
        <span class="spacer-120">Mobile:</span> <label for="agentCell"><%=RecEdit.Fields.Item("ASCell").Value%></label><br>
        <span class="spacer-120">Region:</span> <label for="agentEmail"><%=RecEdit.Fields.Item("RegionName").Value%></label><br>

</div>
<div class="six columns">
<%

%>


</div>
</div>


<hr>
			<h2>Capture Sims Against <%=Session("AgentLabel")%> Profile:</h2>
<form action="SimAllocateAS2.asp" method="get">

<table border="0" cellspacing="2" cellpadding="2">
<tr>
<td Class="quote">Brick / Box Code 1</td><td><input type="text" Name="BrickCode1"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 2</td><td><input type="text" Name="BrickCode2"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 3</td><td><input type="text" Name="BrickCode3"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 4</td><td><input type="text" Name="BrickCode4"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 5</td><td><input type="text" Name="BrickCode5"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 6</td><td><input type="text" Name="BrickCode6"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 7</td><td><input type="text" Name="BrickCode7"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 8</td><td><input type="text" Name="BrickCode8"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 9</td><td><input type="text" Name="BrickCode9"></td>
</tr>
<tr>
<td Class="quote">Brick / Box Code 10</td><td><input type="text" Name="BrickCode10"></td>
</tr>
<tr>
            <td colspan="2" align="center"><label>
              <input name="button2" type="submit" class="orange nice button radius" id="button2" value="Capture">
            </label></td>
          </tr>
  </table>
<input type="Hidden" Name="ASID" Value="<%=Request.QueryString("ASID")%>"></form>


                    </div>
<!-- #include file="includes/footer.asp" -->

