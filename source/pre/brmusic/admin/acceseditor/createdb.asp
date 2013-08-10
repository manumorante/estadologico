<!-- #include file="includetop.asp" -->
<%
'****************************************************
'* Create Database File								*
'****************************************************
'* This file is used to create a new database and	*
'* optionally load it.								*
'****************************************************

'****************************************************
'* Check if the form has been submitted or not		*
'****************************************************
If IsBlank(Request.Form("submit")) = False Then
	CreateDB
Else
	ShowForm
End If

'****************************************************
'* Procedure for creating a new database			*
'****************************************************
Sub CreateDB
	If IsBlank(Request.Form("name")) = False Then
		'****************************************************************************************
		'* ADOX object has to created, because the connection procedure hasn't been used		*
		'****************************************************************************************
		Dim objADOX, path
		
		Set objADOX = Server.CreateObject("ADOX.Catalog")
	
		'****************************************************************************************
		'* Create a new database																*
		'****************************************************************************************
		
		'****************************************************************************************
		'* Create a valid database path & name													*
		'****************************************************************************************
		If Right(Request.Form("path"), 1) = "\" Or Right(Request.Form("path"), 1) = "/" Then
			path = Request.Form("path") & Request.Form("name") & ".mdb"
		Else
			path = Request.Form("path") & "\" & Request.Form("name") & ".mdb"
		End If
		
		'// Check if another database with this name, in
		'// this path doesn't already exist:
		
		'// Create FileSystemObject
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		'// Check if it already exists
		If objFSO.FileExists(path) = True Then '// Display error
			strError = "Another database with this name already exists in this directory. " & _
			"Please go back and choose another name (or directory) by <a href=""createdb.asp"">clicking here</a> or waiting a few seconds for a redirect."
			ErrorMessage "Database already exists", strError
	
			'// Destroy FSO object
			Set objFSO = Nothing
	
			'Finish off tasks
			IncludeBottom
	
			'Redirect to other page
			JSRedirect "createdb.asp", 7
	
			'Finish of processing this page
			Response.End
		End If	
		
		'// Destroy FSO object
		Set objFSO = Nothing
		
		'// End Check
				
		If Request.Form("version") = "97" Then
			'****************************************************************************************
			'* MS Access 97																			*
			'****************************************************************************************
			objADOX.Create "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & path & "; Jet OLEDB:Engine Type=4;"
		Else
			'****************************************************************************************
			'* MS Access 2000																		*
			'****************************************************************************************
			objADOX.Create "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & path & "; Jet OLEDB:Engine Type=5;"
		End If
		
		'****************************************************************************************
		'* Destroy ADOX object																	*
		'****************************************************************************************
		Set objADOX = Nothing
		
		'****************************************************************************************
		'* Check to see if the new database should also be loaded								*
		'****************************************************************************************
		If Request.Form("load") = "yes" Then
			Session("dbPath") = path
		End If
		
		Response.Redirect "index.asp"
	Else '// Display "fill in name" error
		strError = "Please fill in all the required fields. " & _
		"Please go back and try again by <a href=""createdb.asp"">clicking here</a> or waiting a few seconds for a redirect."
		ErrorMessage "Blank fields detected", strError
	
		'Finish off tasks
		IncludeBottom
	
		'Redirect to other page
		JSRedirect "createdb.asp", 5
	
		'Finish of processing this page
		Response.End
	End If
End Sub

'****************************************************
'* Procedure for showing the form to create a new	*
'* database.										*
'****************************************************
Sub ShowForm
	%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" dir="ltr">

	<head>
		<title><%= name %></title>
		<link rel="stylesheet" type="text/css" href="extern/style.css">
		<script src="extern/jscript.js" type="text/javascript"></script>

		<script type="text/javascript" language="javascript">
		<!--
			// Updates the title of the frameset if possible (ns4 does not allow this)
			changetitle('Creating new database on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
		//-->
		</script>
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large">Creating new database on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>

	<!-- create a new database form -->
	<form action="createdb.asp?action=createdb" method="POST" name="createdb">
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1" align="left">
				<th>Name (<i>Without .mdb!</i>)</th>
				<th>Path (<i>Server.MapPath: <%= Server.MapPath("\") %>\</i>)
				<th>Version</th>
				<th>Load Database</th>
			</tr>
    
			<tr id="tdrow2">
				<td align="left"><input type="text" id="textinput" name="name"></td>
				<td align="left"><input type="text" id="textinput" name="path" size="40"></td>
				<td align="left">
					<input type="radio" name="version" value="97">Access 97<br>
					<input type="radio" name="version" value="2000" checked>Access 2000
				</td>
				<td align="left">
					<input type="radio" name="load" value="yes" checked>Yes<br>
					<input type="radio" name="load" value="no">No
				</td>		
			</tr>
    
			<tr id="tdrow2">
				<td colspan="3" align="center"><input type="submit" name="submit" id="button" value="  Create!  "></td>
			</tr>

		</table>
	</form>

	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>
	</body>

	</html>
	<%
End Sub

'****************************************************
'* Call ending tasks procedure						*
'****************************************************
IncludeBottom
%>
