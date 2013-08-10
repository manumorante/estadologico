<%
'****************************************************
'* Error File										*
'****************************************************
'* This file only contains one procedure, but a		*
'* very important one. It's supposed to catch all	*
'* the errors that happen, and produce a neat error	*
'* message.											*
'****************************************************
Sub SQLError(Byval errNumber, errDescription, errSource, errQuery)
	%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" dir="ltr">

	<head>
		<title>aspAccessEditor</title>

		<link rel="stylesheet" type="text/css" href="extern/style.css">
		<script src="extern/jscript.js" type="text/javascript"></script>


		<script type="text/javascript" language="javascript">
			<!--
			// Updates the title of the frameset if possible (ns4 does not allow this)
		   changetitle('SQL error on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
		   -->
		</script>
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large">SQL error on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>

	<font size="2">
		<p><b>Error</b></p>
		<p>
			SQL-query&nbsp;:&nbsp;
		<pre>
<%= errQuery %>
		</pre>
		</p>
		<p>
		    <%= errSource %> said: <br />
		<pre>
<%= errDescription %>
		</pre>
		</p>

		<a href="javascript:history.back(1);">Back</a>

	</font>
	
	<br />
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>
	</body>

	</html>
	<%	
End Sub


Sub ErrorMessage(Byval Title, Message)
	Response.Clear
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
		</head>
	
		<body class="bodyAdmin">
		<br><br><br>
		<table id="table1" cellpadding="5" cellspacing="0" border="0" align="center" width="450">
			<tr>
				<td>
					<table id="tdrow2" cellpadding="4" cellspacing="1" border="0" width="100%">
						<tr>
							<td id="tdrow1" colspan="1">
								<font size="1"><b>Error Message: <%= Title %></b></font>
							</td>
						</tr>
						<tr id="detail">
							<td align="left">
								<p><%= Message %></p>
							</td>
						</tr>
					</table>
				</td></form>
			</tr>
		</table>
		
		</body>
	</html>	<script src="jscript.js" type="text/javascript"></script>
	<%
End Sub
%>