<!-- #include file="includetop.asp" -->
<%
'****************************************************
'* Database Loader File								*
'****************************************************
'* This file is used to load the database in session*
'****************************************************

'****************************************************
'* The form might have been submitted so we enter	*
'* the values into the session variables, even if	*
'* they're empty; it doesn't matter.				*
'****************************************************
Session("dbPath") = Request.Form("server") & Request.Form("dbPath")
Session("dbUser") = Request.Form("dbUser")
Session("dbPassword") = Request.Form("dbPassword")

'****************************************************
'* If the Database Path is empty or blank, then 	*
'* we ask the user to fill it in, else we redirect	*
'* the user to the main page, to edit the database	*
'****************************************************
If IsBlank(Session("dbPath")) = False Then
	Response.Redirect "index.asp"
Else
%>
	<!-- Powered by aspAccessEditor -->
	<!-- Developed by Dennis Pallett -->
	<!-- www.aspit.net -->
	<!-- Do Not remove These Headers -->
	<html>
		<head>
			<title>Database Loader - aspAccessEditor</title>
			
			
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
								<font size="1"><b>Please fill in the database path:</b></font>
							</td>
						</tr>
						<tr id="detail">
							<td align="center" nowrap>
								<p>Make sure you use a valid database path, or you will get an error.</p>
								<form action="loaddb.asp" method="post" id=form1 name=form1>
								<table width="80%" id="tdrow2" cellpadding="0" cellspacing="1" border="0">
									<tr>
										<td align="left" colspan="2">
										<b>Database Path</b> - 
										<a href="javascript:openwindow('selectdb.asp', '250', '400', '250', '25');">
										Database List
										</a>
										</td>
									<tr>
									<tr>
										<td align="left" colspan="2">
										<input type="checkbox" name="server" value="<%= Server.MapPath("\") %>\">
										<i><%= Server.MapPath("\") %>\</i><input id="textinput" type="text" name="dbPath">
										<br><br>
										</td>
									</tr>
									<tr>
										<td align="left" colspan="2">
										<b>Database User</b> (Can be left blank)
										</td>
									</tr>
									<tr>
										<td align="left" colspan="2">
										<input type="text" id="textinput" name="dbUser">
										<br><br>
										</td>
									</tr>
									<tr>
										<td align="left" colspan="2">
										<b>Database Password</b> (Can be left blank)
										</td>
									</tr>
									<tr>
										<td align="left" colspan="2">
										<input type="password" id="textinput" name="dbPassword">
										</td>
									</tr>
									<tr>
										<td align="center">
										<input id="button" type="submit" name="submit" value="Load Database">
										</td>
										<td align="center">
										<input id="button" type="button" name="createdb" value="Create New Database" OnClick="javascript:document.location = 'createdb.asp'">
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
End If
%>

		
