<%
'****************************************************
'* Password Protect File							*
'****************************************************
'* This file is used for protecting the database	*
'* editor, so only authorized users can use it		*
'****************************************************

'****************************************************************************************
'** Set the user and password to protect your database editor							*
'****************************************************************************************
Dim usePwdProtect, userUsername, userPassword

'****************************************************************************************
'** Use password protection?															*
'** 1. Yes																				*
'** 2. No																				*
usePwdProtect = 1															   '*
'****************************************************************************************

'****************************************************************************************
'** Enter the username here:															*
userUsername = "admin"														   '*	
'****************************************************************************************

'****************************************************************************************
'** Enter the password here:																*
userPassword = "1234"														   '*
'****************************************************************************************

'****************************************************************************************
'** The form might have been submitted so we enter										*
'** the values into the session variables, even if										*
'** they're empty; it doesn't matter.													*
'****************************************************************************************
If Request.Form("submit") = "Go" Then
	Session("userUsername") = Request.Form("username")
	Session("userPassword") = Request.Form("password")
	Response.Redirect "../index.asp"
End If
If usePwdProtect = 1 And (Session("userUsername") <> userUsername _
Or Session("userPassword") <> userPassword) Then
%>
	<!-- Powered by aspAccessEditor -->
	<!-- Developed by Dennis Pallett -->
	<!-- www.aspit.net -->
	<!-- Do Not remove These Headers -->
	<html>
		<head>
			<title>Login - aspAccessEditor</title>
			
			
		<link rel="stylesheet" type="text/css" href="extern/style.css">
		<script src="extern/jscript.js" type="text/javascript"></script>
		
		<script language= "JavaScript">
		<!--Break out of frames
			if (top.frames.length!=0)
			top.location=self.document.location;
			//-->
		</script>
		</head>
	
		<body class="bodyAdmin">
		<br><br><br>
		<table id="table1" cellpadding="5" cellspacing="0" border="0" align="center" width="450">
			<tr>
				<td>
					<table id="tdrow2" cellpadding="4" cellspacing="1" border="0" width="100%">
						<tr>
							<td id="tdrow1" colspan="1">
								<font size="1"><b>Please fill in a valid username and password:</b></font>
							</td>
						</tr>
						<tr id="detail">
							<td align="center" nowrap>
								<p>Make sure you use a valid username and password, or else you will be denied access.</p>
								<form action="includes/pwdprotect.asp" method="post" id="pwd" name="pwd">
								<table width="50%" id="tdrow2" cellpadding="0" cellspacing="1" border="0">
									<tr>
										<td width="10%" align="left">
										<b>Username</b>
										</td>
										<td align="left">
										<input id="textinput" type="text" name="username">
										<br><br>
										</td>
																	
									</tr>
									<tr>
										<td width="10%" align="left">
										<b>Password</b>
										</td>
										<td align="left">
										<input type="password" id="textinput" name="password">
										<br><br>
										</td>
									</tr>
									<tr>
									<tr>
										<td colspan="2" align="center">
										<input id="button" type="submit" name="submit" value="Go">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td></form>
			</tr>
		</table>
		
		</body>
	</html>
<%
Response.End
End If
%>
