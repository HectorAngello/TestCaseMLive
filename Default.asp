<!-- #include file="includes/header.asp" -->
<%
' create Token
Part1 = Replace(Date, "/", "")
	Function RandomNumber(intHighestNumber)
		Randomize
		RandomNumber = Int(Rnd * intHighestNumber) + 1
	End Function

Part2 = Int(Right(5000 + RandomNumber(3000000000000),5))

If Len(Part2) < 8 Then
Part2 = Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2 & Part2
Part2 = Left(Part2,8)
End If
Part2 = "-" & Part2

Part3 = Int(Right(5000 + RandomNumber(9000000000000),5))
If Len(Part3) < 10 Then
Part3 = Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3 & Part3
Part3 = Left(Part3,10)
End If

Part3 = "-" & Part3

Part4 = "-" & DatePart("h",Now()) & "-" & DatePart("n",Now())
Token = Part1 & Part2 & Part3 & Part4
' End Create Token


%>
<!-- header -->
    <div class="header">
        <div class="row">
            <div class="nine columns">
                <!-- #include file="includes/Logo.inc" -->
            </div>
            <div class="three columns">
                <div class="logged-in">
                    <!-- #include file="Includes/User.inc" -->
                </div>
            </div>
        </div>
    </div>
    
	<!-- container -->
	<div class="container">
        <div id="main-menu" class="row">
            
            <div class="twelve columns">
                <div class="content">
                    <div class="row heading"><%If Request.QueryString("Error") = "UnknownUser" Then%><div class="alert-box error">Error, Incorrect username / Password, please try again, or use the 'Forgot Password' feature below.</div><%End If%>
<%If Request.QueryString("Login") = "Fail" Then%><div class="alert-box error">Error, Incorrect username / Password, please try again, or use the 'Forgot Password' feature below.</div><%End If%>
<%If Request.QueryString("Error") = "Expired" Then%><div class="alert-box error">Your user session has been inactive for too long and the system has automatically logged you out, please log in again.</div><%End If%>
<%If Request.QueryString("ForgotPass") = "True" Then%><div class="alert-box success">Username / Password sent to <%=Request.QueryString("Email")%></div><%End If%>
<%If Request.QueryString("ForgotPass") = "False" Then%><div class="alert-box warning">Username / Password not found for <%=Request.QueryString("Email")%>, please speak to your system administrator.</div><%End If%>
                        <div id="login-box" class="eight columns centered">
                <div class="panel">
                    <div class="row">
                        <div class="six columns">
                            <span class="span-heading">Login</span>
                            <form action="LoginNoOTP.asp" method="post" name="form1" onSubmit="MM_validateForm('UserName','','R','Password','','R');return document.MM_returnValue" class="login-form nice">
                                <label for="username">Username:</label>
                                <input type="text" name="UserName" class="input-text-small" />
                                
                                <label for="password">Password:</label>
                                <input type="password" name="Password" class="input-text-small" />
                                
                                <br><input type="submit" name="login" value="Login" class="nice red radius button" />
                            </form>
                        </div>
                        <div class="six columns">
                            <span class="span-heading">Forgot Password</span>
                            <form action="ForgotPass.asp" method="post" name="form2" onSubmit="MM_validateForm('EmailAddress','','RisEmail');return document.MM_returnValue" class="forgot-password-form nice">
                                <label for="email">Email Address:</label>
                                <input type="text" name="EmailAddress" class="input-text-small" />
                                
                                <input type="submit" name="resend" value="Send Password" class="nice red radius button" />
                            </form>
                        </div>
                    </div>
                </div>
            </div>
                    </div>
                    
<!-- #include file="includes/footer.asp" -->

