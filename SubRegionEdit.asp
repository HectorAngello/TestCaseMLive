<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecRegion = Server.CreateObject("ADODB.Recordset")
RecRegion.ActiveConnection = MM_Site_STRING
RecRegion.Source = "SELECT * FROM ViewRegionsDetail Where RID = " & Request.QueryString("RID")
RecRegion.CursorType = 0
RecRegion.CursorLocation = 2
RecRegion.LockType = 3
RecRegion.Open()
RecRegion_numRows = 0

set RecSubRegion = Server.CreateObject("ADODB.Recordset")
RecSubRegion.ActiveConnection = MM_Site_STRING
RecSubRegion.Source = "SELECT * FROM SubRegions Where SRID = " & Request.QueryString("SRID")
RecSubRegion.CursorType = 0
RecSubRegion.CursorLocation = 2
RecSubRegion.LockType = 3
RecSubRegion.Open()
RecSubRegion_numRows = 0
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
                        <div class="eight columns"><h1><%=RecRegion.Fields.Item("RegionName").Value%> Sub Regions</h1></div>
                        <div class="four columns buttons"><a href="javascript:history.back(1)" class="nice white radius button"><p class="new-button">Back</p></a></div>
                    </div>
                <h3>Update Sub Region</h3>
<form name="AddRegion" action="SubRegionEdit2.asp" method="post"  class="nice">
                        <fieldset>
                            <div class="five columns">
   
                                <label for="agencyCode">Region Name *</label>
                                <input type="text" name="RegionName" class="input-text" value="<%=RecSubRegion.Fields.Item("SubRegionName").Value%>" Required />
    
                               <label for="agencyName">Region Code *</label>
                                <input type="text" name="RegionCode" class="input-text" value="<%=RecSubRegion.Fields.Item("SubRegionCode").Value%>" Required />
    
                                <label for="agentEmail">Head Count Target *</label>
                                <input type="text" name="HeadCountTarget" class="input-text" value="<%=RecSubRegion.Fields.Item("HeadCountTarget").Value%>" Required />
                             

                                <p>* Required Fields<br>
                                    <input type="Submit" class="orange nice button radius" value="Update Sub Region">
                                </p>
                            </div>
                            
                        </fieldset>
<input type="Hidden" Name="SRID" Value="<%=Request.QueryString("SRID")%>">
<input type="Hidden" Name="RID" Value="<%=Request.QueryString("RID")%>">
                    </form>
<!-- #include file="includes/footer.asp" -->

