<!-- #include file="includes/header.asp" -->

<%
set RecReconRegions = Server.CreateObject("ADODB.Recordset")
RecReconRegions.ActiveConnection = MM_Site_STRING
RecReconRegions.Source = "SELECT Distinct RID, RegionName FROM viewUserRegion where Active = 'Yes' and CompanyID = " & Session("CompanyID") & " and UserID = " & Session("UNID") & " Order By RegionName Asc"
RecReconRegions.CursorType = 0
RecReconRegions.CursorLocation = 2
RecReconRegions.LockType = 3
RecReconRegions.Open()
RecReconRegions_numRows = 0


%><table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td Class="ontab"><h3>Master Agent File Export</h3></td>
  </tr>
  <tr>
    <td align="left" valign="top"><form action="Report-MasterFile2.asp" name="ZRCS" method="Post" Target="_parent">
<table>
<tr>
<td class="quote">Region:</td>
<td><select name="Region">

<Option selected value="0">Select Region</Option>
<option value="0">All</option>
                <% 
While (NOT RecReconRegions.EOF)
%>
                <option value="<%=(RecReconRegions.Fields.Item("RID").Value)%>" ><%=(RecReconRegions.Fields.Item("RegionName").Value)%></option>
                <%
  RecReconRegions.MoveNext()
Wend
%>
              </select></td></tr>

<tr><td class="quote">Status:</td><td><select Name="Type">
<option Value="0">All Agents</option>
<option Value="1">Terminated Agents</option>
<option Value="2">Active Agents</option>
</select></td></tr>
<tr><td class="quote">Export Fields:</td><td>
<table>
<tr>
<td>Name:</td>
<td><input type="checkbox" Name="AgentName" checked Value="Yes"></td>
<td>Region:</td>
<td><input type="checkbox" Name="AgentRegion" checked Value="Yes"></td>
</tr>
<tr>
<td>Employee Code:</td>
<td><input type="checkbox" Name="AgentEmpCode" checked Value="Yes"></td>
<td>Sub Region:</td>
<td><input type="checkbox" Name="AgentSubRegion" checked Value="Yes"></td>
</tr>
<tr>
<td>Start Date:</td>
<td><input type="checkbox" Name="StartDate" checked Value="Yes"></td>
<td>Primary Mobile Number:</td>
<td><input type="checkbox" Name="AgentCell" checked Value="Yes"></td>
</tr>
<tr>
<td>ID No:</td>
<td><input type="checkbox" Name="IDNo" checked Value="Yes"></td>
<td>Secondary Mobile Number:</td>
<td><input type="checkbox" Name="AgentCell2" checked Value="Yes"></td>
</tr>
<tr>
<td>Tertiary Mobile Number:</td>
<td><input type="checkbox" Name="TertiaryMobileNumber" checked Value="Yes"></td>
<td>Trading Spot</td>
<td><input type="checkbox" Name="TradingSpot" checked Value="Yes"></td>
</tr>
<tr>
<td>Gender:</td>
<td><input type="checkbox" Name="Gender" Value="Yes"></td>
<td>Race:</td>
<td><input type="checkbox" Name="Race" Value="Yes"></td>
</tr>
<tr>
<td>BEE Signature Date:</td>
<td><input type="checkbox" Name="TaxOffice" Value="Yes"></td>
<td>Cluster:</td>
<td><input type="checkbox" Name="SARS" Value="Yes"></td>
</tr>
<tr>
<td>Bank:</td>
<td><input type="checkbox" Name="Bank" checked Value="Yes"></td>
<td>Branch:</td>
<td><input type="checkbox" Name="Branch" checked Value="Yes"></td>
</tr>
<tr>
<td>Account Type:</td>
<td><input type="checkbox" Name="AccType" checked Value="Yes"></td>
<td>Account No:</td>
<td><input type="checkbox" Name="AccNo" checked Value="Yes"></td>
</tr>
<tr>
<td>Residential Address:</td>
<td><input type="checkbox" Name="ResidentialAddress" Value="Yes"></td>
<td>Opt In For Realtime Commission</td>
<td><input type="checkbox" Name="RealTimeCommOptIn" Value="Yes"></td>
</tr>
<tr>
<td>Termination Date:</td>
<td><input type="checkbox" Name="TermDate" checked Value="Yes"></td>
<td>Termination Reason:</td>
<td><input type="checkbox" Name="TermReason" Value="Yes"></td>
</tr>
<tr>
<td><%=SupervisorLabel%> Name:</td>
<td><input type="checkbox" Name="BDOName" checked Value="Yes"></td>
<td><%=SupervisorLabel%> Code:</td>
<td><input type="checkbox" Name="BDOEmpCode" Value="Yes"></td>
</tr>


<tr>
<td>M-Charge Balance:</td>
<td><input type="checkbox" Name="MChargeBal" checked Value="Yes"></td>
<td>Purse Limit:</td>
<td><input type="checkbox" Name="PurseLimit" checked Value="Yes"></td>
</tr>

<tr>
<td>Exclude From Bulk File Generation:</td>
<td><input type="checkbox" Name="ExcludeFromMchargeBulkFile" Value="Yes"></td>
<td>Mobi Site Password</td>
<td><input type="checkbox" Name="ShowMobiSitePass" value="Yes"></td>
</tr>

<tr>
<td>Agent Status:</td>
<td><input type="checkbox" Name="ShowAgentStatus" checked Value="Yes"></td>
<td>Regional Manager:</td>
<td><input type="checkbox" Name="ShowRegionAlManager" checked Value="Yes"></td>
</tr>

<tr>
<td>Last Banked Amount:</td>
<td><input type="checkbox" Name="ShowLastRefund" Value="Yes" checked></td>
<td>Last Banked Date:</td>
<td><input type="checkbox" Name="LastBankedDate" Value="Yes" checked></td>
</tr>

<tr>
<td>Last Airtime Amount:</td>
<td><input type="checkbox" Name="ShowLastAirtimeAmount" Value="Yes" checked></td>
<td>Last Airtime Date:</td>
<td><input type="checkbox" Name="LastAirtimeDate" Value="Yes" checked></td>
</tr>
<tr>
<td>Is A Mcharge Agent:</td>
<td><input type="checkbox" Name="ShowIfMchargeAgent" Value="Yes"></td>
<td>Is A Mobile Money Agent:</td>
<td><input type="checkbox" Name="ShowIfMobileMoneyAgent" Value="Yes"></td>
</tr>

<tr>
<td>Mobile Money Purse Limit:</td>
<td><input type="checkbox" Name="MMPurseLimit" Value="Yes" ></td>
<td>Mobile Money Balance:</td>
<td><input type="checkbox" Name="MobileMoneyBalance" Value="Yes" ></td>
</tr>

<tr>
<td>Mobile Money Last Banked Date:</td>
<td><input type="checkbox" Name="MMLastBankedDate" Value="Yes" ></td>
<td>Work Permit Expiry Date:</td>
<td><input type="checkbox" Name="WPED" Value="Yes" ></td>
</tr>

<tr>
<td>New Agent Summary:</td>
<td><input type="checkbox" Name="DDList" Value="Yes" ></td>
<td>Mobile Money Acc Number</td>
<td><input type="checkbox" Name="MoMoAccNo" Value="Yes" ></td>
</tr>

<tr>
<td>Made for Skhokho GSM:</td>
<td><input type="checkbox" Name="SkhokhoGSM" Value="Yes" ></td>
<td>Made for Skhokho Dedicated:</td>
<td><input type="checkbox" Name="SkhokhoDedicated" Value="Yes" ></td>
</tr>

<tr>
<td>Work Permit Expiry Date:</td>
<td><input type="checkbox" Name="WPExpiryDate" Value="Yes" ></td>
<td>Sub Region Code</td>
<td><input type="checkbox" Name="SubRegionCode" Value="Yes" ></td>
</tr>

</table>
</td></tr>
<tr><td class="quote">Order By:</td><td><select Name="OrderBy">
<option Value="1">Agent First Name</option>
<option Value="2">Agent Last Name</option>
<option Value="3">Agent Code</option>
</select></td></tr>
<tr><td class="quote">Output Format:</td><td><select Name="OutFormat">
<option Value="B">Browser</option>
<option Value="E">Excel</option>
</select></td></tr>

<tr><td Align="Center" Colspan="2"><input name="button2" type="submit" class="nice red radius button" id="button2" value="Run Report"></td></tr>
</table>
</form></td>
    
  </tr>

</table>