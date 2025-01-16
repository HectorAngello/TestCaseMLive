<!-- #include file="includes/header.asp" -->
<%
If Session("UNID") = "" Then
   Response.Redirect "Default.asp?Error=Expired" 
End If

set RecBulk = Server.CreateObject("ADODB.Recordset")
RecBulk.ActiveConnection = MM_Site_STRING
RecBulk.Source = "SELECT * FROM BulkMChargeMM Where BulkID = " & Request.QueryString("BulkID")
RecBulk.CursorType = 0
RecBulk.CursorLocation = 2
RecBulk.LockType = 3
RecBulk.Open()
RecBulk_numRows = 0
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

                        <div class="eight columns"><h1>Mobile Money Management</h1></div>
                        <div class="four columns buttons"><a href="Display.asp?AppCat=18&AppSubCatID=1045" class="nice white radius button"><p class="new-button">Back</p></a></div>
<br><br><br>Export File Has Been Scheduled For Generation<br><br>
Your File Name is: <%=(RecBulk.Fields.Item("FileName").Value)%>.txt
<br><br><a href="Display.asp?AppCat=18&AppSubCatID=1045">Back To M-Charge Bulk File Management</a>
<script language="javascript">      
       //Create an iframe and turn on the design mode for it 
       document.write ('<iframe src="CompileMobileMoneyExport.asp?UN=<%=Session("UNID")%>" id="Abstract" frameborder="0" width="1%" height="1"></iframe>')
frames.Abstract.document.designMode = "off";               
 </script>
                    </div>
                
<!-- #include file="includes/footer.asp" -->

