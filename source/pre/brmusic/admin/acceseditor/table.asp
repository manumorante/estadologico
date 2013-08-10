<!-- #include file="includetop.asp" -->
<%
'****************************************************
'* Table File										*
'****************************************************
'* This file contains everything that has to do		*
'* with tables and fields.							*
'****************************************************

'****************************************************
'* Check if a valid table has been passed			*
'****************************************************
table = Request.QueryString("table")

'// Strip away any pending ;
If Right(table, 1) = ";" Then table = Replace(table, ";", Empty)

If TableExists(table) = False Or IsNumeric(table) = True Then
	'Invalid Table
	strError = "An invalid table name has been passed to this page. " & _
	"Please go back and try again. If you clicked a valid link, please notify the system administrator."

	'Display error
	ErrorMessage "Invalid Table", strError
	
	'Finish off ending tasks
	IncludeBottom
	
	'Redirect to index
	JSRedirect "index.asp", 5
	
	'End this page
	Response.End
End If

'****************************************************
'* Get action and determine what to do				*
'****************************************************
action = Request.QueryString("action")

SELECT CASE action
	CASE "empty"
		EmptyTable
	CASE "drop"
		DropTable
	CASE "editfield"
		EditField
	CASE "dropfield"
		DropField
	CASE "updatefield"
		UpdateField
	CASE "addindex"
		AddIndex
	CASE "dropindex"
		DropIndex
	CASE "dofields"
		DoFields
	CASE "dosql"
		db.Query(Request.Form("sql"))
		Response.Redirect "table.asp?table=" & table
	CASE "rename"
		RenameTable
	CASE "addfield"
		AddField
	CASE "insertfield"
		InsertField
	CASE Else
		DisplayTable
END SELECT



'****************************************************
'* Procedure for displaying the table screen		*
'****************************************************
Sub DisplayTable
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
			changetitle('<%= db.objADOX.Tables(table).Name %> running on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
			//-->
		</script>

    
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large">table <i><%= db.objADOX.Tables(table).Name %></i> running on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>



	<!-- first browse links -->
	<p>
	[ 
	<%
    '****************************************************
    '* Determine if should link the browse text		*
    '****************************************************
	If db.objADOX.Tables(table).Columns.Count > 0 Then    
	    strQuery = "SELECT COUNT(*) FROM [" & table & "]"
		db.Query(strQuery)
		If db.objRS(0) > 0 Then
			%>
			<a href="record.asp?sql=SELECT * FROM [<%= table %>]">
			<% 
		End If
		db.QueryClose
	End If
    %>
     <b>Browse</b></a> ]&nbsp;&nbsp;&nbsp;
    [ <a href="record.asp?action=addrecord&table=<%= table %>&sql=bypassfilter">
        <b>Insert</b></a> ]&nbsp;&nbsp;&nbsp;
    [ <a href="javascript:ConfirmLink('table.asp?action=empty&table=<%= table %>', 'Are you sure you want to empty this table?')">
         <b>Empty</b></a> ]&nbsp;&nbsp;&nbsp;
    [ <a href="javascript:ConfirmLink('table.asp?action=drop&table=<%= table %>', 'Are you sure you want to drop this table?')">
         <b>Drop</b></a> ]
	</p>
    
	<!-- TABLE INFORMATIONS -->

	<form action="table.asp?action=dofields&table=<%= table %>" method="post" id="editfields" name="editfields">
		
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1">
				<td></td>
				<th>&nbsp;Field&nbsp;</th>
				<th>Type</th>
				<th>Null</th>
				<th>Default</th>
				<th>Size</th>
				<th>Extra</th>
				<th colspan="4">Action</th>
			</tr>
	
			<%
			'****************************************************
			'* Loop through the table's column (field) collectio*
			'* and display it and the column'a properties.		*
			'* Also, display links to edit, drop, or make an	*
			'* index out of the column.							*
			'****************************************************
			For Each field in db.objADOX.Tables(table).Columns
			%>
				<tr>
					<td align="center" id="tdrow2">
						<input type="checkbox" name="selected_<%= field.Name %>" value="<%= field.Name %>" id="checkbox_row_<%= table %>" />
					</td>
					<td id="tdrow2" nowrap="nowrap">&nbsp;<label for="checkbox_row_<%= table %>"><u><%= field.Name %></u></label>&nbsp;</td>
					<td id="tdrow2" nowrap="nowrap">
				<%
					'// Determine which field type to write
					Response.Write column.ColumnName(field.Type)
					%>
						<bdo dir="ltr"></bdo></td>
					<td id="tdrow2">&nbsp;
					<%
					'****************************************************
					'** Determine if zero lengths are allowed												*
					'****************************************************
					If field.Properties("Jet OLEDB:Allow Zero Length") = True Then
						Response.Write "Yes"
					Else
						Response.Write "No"
					End If
					%>
					</td>
					<td id="tdrow2" nowrap="nowrap">&nbsp;<%= field.Properties("Default") %></td>
					<td id="tdrow2" nowrap="nowrap">
					<%
					If field.DefinedSize > 0 Then
						Response.Write field.DefinedSize
					End If
					%>    
					&nbsp;</td>			
					<td id="tdrow2" nowrap="nowrap">
					<%
					If field.Properties("AutoIncrement") = True Then
						Response.Write "autoincrement"
					End If
					%>    
					&nbsp;</td>
					<td id="tdrow2">
						<a href="table.asp?action=editfield&field=<%= field.Name %>&table=<%= table %>">
						Change</a>
					</td>
					<td id="tdrow2">
						<a href="javascript:ConfirmLink('table.asp?action=dropfield&field=<%= field.Name %>&table=<%= table %>', 'Are you sure you want to drop this field?\n\nPlease note that fields with an index cannot be dropped. Drop the index before dropping the field.')">
						Drop</a>
					</td>
					<td id="tdrow2">
						<a href="table.asp?action=addindex&subaction=index&table=<%= table %>&field=<%= field.Name %>">
						Index</a>
					</td>
					<td id="tdrow2">
						<a href="javascript:ConfirmLink('table.asp?action=addindex&subaction=primary&table=<%= table %>&field=<%= field.Name %>', 'Please note that this table will be emptied when creating a new primary key.\nDo you want to continue?')">
						Primary Key</a>
					</td>
				</tr>
				<%
			Next
			%> 
			<tr>
				<td colspan="13">
					<img src="pictures/arrow.gif" border="0" width="38" height="22" alt="With selected:" />
					<i>With selected:</i>&nbsp;&nbsp;
					<input id="button" type="submit" name="submit" value="Change" />
					&nbsp;<i>Or</i>&nbsp;
					<input id="button" type="submit" name="submit" value="Drop" />
				</td>
			</tr>
		</table>
	</form>
	
	<!-- Indexes -->
	<br />
	<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<!-- Indexes form -->
				<table border="0" id="memgroup">
					<tr id="tdrow1">
						<th>Keyname</th>
						<th>Action</th>
						<th>Field</th>
						<th>Type</th>
					</tr>     
					<%
					'****************************************************
					'* Loop through the index collection of this table	*
					'****************************************************
			        For Each index In db.objADOX.Tables(table).Indexes
						%>
						<tr>
							<td id="tdrow2" rowspan="1">
								<%= index.Name %>
							</td>
							<td id="tdrow2" rowspan="1">
								<a href="javascript:ConfirmLink('table.asp?action=dropindex&index=<%= index.Name %>&table=<%= table %>', 'Are you sure you want to drop this index?')">Drop</a>
							</td>
							<td id="tdrow2">
								<%= index.Columns(0) %>
							</td>
							<td id="tdrow2">
								<%
								If index.PrimaryKey = True Then
									Response.Write "PRIMARY"
								Else
									Response.Write "INDEX"
								End If
								%>									
							</td>
						</tr>
					<%
		        Next
		        %>
				</table><br />      
			</td>
		</tr>
	</table>
	
	<hr />
	
	<!-- TABLE WORK -->
	<ul>
	
	<!-- Execute custom SQL -->
	<script language="JavaScript">
	<!--
		function parsesql(tform) {
			// check if select is being done
			if (UCase(Left(tform.sql.value, 6)) == "SELECT") {
				document.location = "record.asp?sql=" + tform.sql.value;
				return false;
			}
			
			// check if a drop is being done
			if (UCase(Left(tform.sql.value, 4)) == "DROP") {
				if (confirm('Are you sure about this?')) {
					return true;
				} else {
					return false;
				}
			}
		}
	// -->
	</script>
	<li>
		<form action="table.asp?action=dosql&table=<%= table %>" name="dosql" method="post" OnSubmit="return parsesql(this);">
			<a name="sql"></a>
			Execute SQL: <input size="50" name="sql" type="text" id="textinput" style="vertical-align: middle" value="<%= Request.QueryString("sql") %>"> 
			<input id="button" type="submit" value="Go" style="vertical-align: middle">
			
		</form>
		
	</li>


    <!-- Add some new fields -->
    <li>
        <form method="post" action="table.asp?action=addfield&table=<%= table %>" id="addfield" name="addfield">
            
            
            Add new field&nbsp;:
            <input type="text" name="number" size="2" maxlength="2" value="1" id="textinput" style="vertical-align: middle" onfocus="this.select()" />
            <input id="button" type="submit" value="Go" style="vertical-align: middle" />
        </form>
    </li>

    <!-- Change table name -->
    <li>
        <div style="margin-bottom: 10px">
            <form method="post" action="table.asp?action=rename&table=<%= table %>" id="renametable" name="renametable">
                Rename table to&nbsp;:
                <input type="text" size="20" name="name" value="<%= db.objADOX.Tables(table).Name %>" id="textinput" onfocus="this.select()" />&nbsp;
                <input type="submit" id="button" value="Go" />
            </form>
        </div>
    </li>

    <!-- Deletes the table -->
    <li>
        <a href="javascript:ConfirmLink('table.asp?action=drop&table=<%= table %>', 'Are you sure you want to drop this table?')">
            Drop table</a>
    </li>

	</ul>
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>

	</body>

	</html>
	<%
End Sub

'****************************************************
'* Procedure for emptying a table					*
'****************************************************
Sub EmptyTable
	'****************************************************
	'* If there are fields, empty table					*
	'****************************************************
	If db.objADOX.Tables(table).Columns.Count > 0 Then
		db.Query("DELETE * FROM [" & table & "]")
	End If

	'****************************************************
	'* Redirect back to table page						*
	'****************************************************
	Response.Redirect "table.asp?table=" & table
End Sub

'****************************************************
'* Procedure for dropping a table					*
'****************************************************
Sub DropTable
	'****************************************************
	'* Drop the table									*
	'****************************************************
	db.Query("DROP TABLE [" & table & "]")	

	'****************************************************
	'* Redirect to index page							*
	'****************************************************
	Response.Redirect "index.asp"
End Sub

'****************************************************
'* Procedure for dropping a field					*
'****************************************************
Sub DropField
	'****************************************************
	'* Extract field array								*
	'****************************************************
	field = Request.QueryString("field")
	
	field  = Split(field, ",")
	
	'****************************************************
	'* Start SQL query for dropping	field(s)			*
	'****************************************************
	strQuery = "ALTER TABLE [" & table & "]"
	first = True
	
	For Each fieldloop In field
		'****************************************************
		'* Check if the field in question actually exists	*
		'****************************************************
		If FieldExists(table, Trim(fieldloop)) = True And IsNumeric(field) = False Then
			'****************************************************
			'Check if field is an index or not					*
			'****************************************************
			If db.objADOX.Tables(table).Indexes.Count > 0 Then
				subaction = False
				For Each index in db.objADOX.Tables(table).Indexes
					If LCase(index.Columns(0)) <> LCase(fieldloop) Then
						'Check if field hasn't already been added to query
						If subaction = False Then
							'Add field to SQL query
							If first = True Then
								strQuery = strQuery & " DROP COLUMN [" & Trim(fieldloop) & "]"
								first = False
							Else
								strQuery = strQuery & ", [" & Trim(fieldloop) & "]"
							End If
							subaction = True
							valid = True
						End If
					End If
				Next
			Else
				'Add field to SQL query
				If first = True Then
					strQuery = strQuery & " DROP COLUMN [" & Trim(fieldloop) & "]"
					first = False
				Else
					strQuery = strQuery & ", [" & Trim(fieldloop) & "]"
				End If
				valid = True
			End If
		End If
	Next
	
	'****************************************************
	'* Check if atleast one field has been dropped		*
	'****************************************************
	If valid = False Then
		'Invalid Field
		strError = "An invalid field name(s) has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Field", strError
	
		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5
	
		'Exit procedure
		Exit Sub
	End If

	'****************************************************
	'* Execute query									*
	'****************************************************
	If strQuery <> "ALTER TABLE [" & table & "]" Then
		db.Query(strQuery)
	End If
	
	'****************************************************
	'* Redirect back to table page						*
	'****************************************************
	Response.Redirect "table.asp?table=" & table
End Sub

'****************************************************
'* Procedure for editing field(s)					*
'****************************************************
Sub EditField
	field = Request.QueryString("field")
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
		   changetitle('<%= db.objADOX.Tables(table).Name %> running on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
		   -->
		</script>
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large">table <i><%= db.objADOX.Tables(table).Name %></i> running on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>

	<form name="changefield" method="post" action="table.asp?action=updatefield&table=<%= table %>&field=<%= field %>">
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1">
				<th>Field</th>
				<th>Type</th>
				<th>Null*</th>
				<th>Default</th>
				<th>Extra**</th>
				<th>Size***</th>
			</tr>
	<%
	'****************************************************
	'* Extract field array								*
	'****************************************************
	field  = Split(field, ",")
	
	For Each fieldloop In field
		'****************************************************
		'* Display the fields if they exists				*
		'****************************************************
		If FieldExists(table, Trim(fieldloop)) = True And IsNumeric(field) = False Then
			'****************************************************
			'* Display field in page							*
			'****************************************************
			%>
			<tr>
				<td id="tdrow2">
					<input type="text" name="name_<%= fieldloop %>" id="textinput" size="10" maxlength="64" value="<%= db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Name %>" />
				</td>
				<td id="tdrow2">
					<select id="dropdown" name="type_<%= fieldloop %>">
						<%
						'// Loop through all the columns
						For Each forloop In column.columns
							Response.Write "<option value=""" & forloop(0) & """" & _
							SelectedData(forloop(0), db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Type) & _
							">" & forloop(1) & "</option>"							
						Next						
						%>
					</select>
				</td>
				<td id="tdrow2">
					<select id="dropdown" name="null_<%= fieldloop %>">
						<%
						If db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Properties("Jet OLEDB:Allow Zero Length") = True Then
							%>
							<option value="False">not null</option>
							<option value="True" selected>null</option>
							<%
						Else
							%>
							<option value="no" selected>not null</option>
							<option value="yes">null</option>
							<%
						End If
						%>
					</select>
					</td>
				<td id="tdrow2">
					<input id="textinput" type="text" name="default_<%= fieldloop %>" size="8" value="<%= db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Properties("Default") %>">
		        </td>
		        <td id="tdrow2">
			        <select id="dropdown" name="autoincrement_<%= fieldloop %>">
						<option value="no"<%= SelectedData(False, db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Properties("AutoIncrement")) %>></option>
						<option value="yes"<%= SelectedData(True, db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Properties("AutoIncrement")) %>>auto_increment</option>
					</select>
				</td>
				<td id="tdrow2">
					<input id="textinput" type="text" name="DefinedSize_<%= fieldloop %>" size="8" value="<%= db.objADOX.Tables(table).Columns(Cstr(fieldloop)).DefinedSize %>">
		        </td>
			</tr>
			<%
			valid = True
		End If
	Next
	
	'****************************************************
	'* Check if atleast one field has been displayed	*
	'****************************************************
	If valid = False Then
		'Invalid Field
		strError = "An invalid field name(s) has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Field", strError
	
		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5
	
		'Exit procedure
		Exit Sub
	End If
	%>
	</table>
    <br />
    
	<input id="button" type="submit" name="submit" value="Save" />
	</form>
	<table>
		<tr>
			<td valign="top">*&nbsp;</td>
			<td>
				<b>Certain fieldtypes do not allow null length.</b>
			</td>
		</tr>
		<tr>
			<td valign="top">**&nbsp;</td>
			<td>
				<b>Please note that when you choose auto_increment all the data in this table (<i><%= table %></i>) will be wiped.
				Also, if another field already auto increments, this field will not be changed to auto increment.
				</b>
			</td>
		</tr>
		<tr>
			<td valign="top">***&nbsp;</td>
			<td>
				<b>Please note that size field only apply to text-fields.
				</b>
			</td>
		</tr>
	</table>
	<br />
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>
	</body>

	</html>
	<%	
End Sub

'****************************************************
'* Procedure to update properties of field(s)		*
'****************************************************
Sub UpdateField
	'****************************************************
	'* Extract field array								*
	'****************************************************
	field = Request.QueryString("field")
	field  = Split(field, ",")

	For Each fieldloop In field
		'****************************************************
		'* Check if the field in question actually exists	*
		'****************************************************
		If FieldExists(table, Trim(fieldloop)) = True And IsNumeric(field) = False Then
			'****************************************************
			'* Retrieve field values and add SQL query			*
			'****************************************************
			strQuery = "ALTER TABLE [" & table & "] ALTER COLUMN [" & fieldloop & "] "

			
			'// Add field type
			strQuery = strQuery & column.ColumnDatabase(Request.Form("type_" & fieldloop))
			
			response.Write("--------->"&Request.Form("type_" & fieldloop))
			
			'// Add field size (only on textfields)
			If IsNumeric(Request.Form("DefinedSize_" & fieldloop)) = True _
			And CLng(Request.Form("type_" & fieldloop)) = adVarWChar _
			Then
				'// Check if it's a valid size
				If Request.Form("DefinedSize_" & fieldloop) < 256 _
				And Request.Form("DefinedSize_" & fieldloop) > 0 _
				Then
					strQuery = strQuery & "(" & Request.Form("DefinedSize_" & fieldloop) & ")"
				End If
			End If
		
			'Check if this field should change to autoincrement
			If Request.Form("autoincrement_" & fieldloop) = "yes" Then
				valid = True
			
				'Check if other fields are already autoincrement
				For Each i In db.objADOX.Tables(table).Columns
					If LCase(i.Name) <> LCase(fieldloop) And i.Properties("AutoIncrement") = True Then
						valid = False
					End If
				Next
	
				If valid = True Then		
					'Empty table
					strQuery = "DELETE FROM [" & table & "]"
					db.Query(strQuery)
				
					'Change field to autoincrement
					strQuery = "ALTER TABLE [" & table & "] ALTER COLUMN [" & fieldloop & "] Autoincrement"
				End If
			End If
			
			'Execute SQL query
			db.Query(strQuery)

				
			temp = Empty
			temp = fieldloop
			
			'// Check if there isn't another field already named
			'// like this.
			If FieldExists(table, Request.Form("name_" & fieldloop)) = False Then
				'Change name
				If IsBlank(Request.Form("name_" & fieldloop)) = False Then
					db.objADOX.Tables(table).Columns(Cstr(fieldloop)).Name = Request.Form("name_" & fieldloop)
					temp = Request.Form("name_" & fieldloop)
				End If
			End If
			
			'Change default value
			db.objADOX.Tables(table).Columns(Cstr(temp)).Properties("Default") = Request.Form("default_" & fieldloop)
			
			'Change null property
			If Request.Form("null_" & fieldloop) = "yes" Then
				db.objADOX.Tables(table).Columns(Cstr(temp)).Properties("Jet OLEDB:Allow Zero Length") = True
			Else
				db.objADOX.Tables(table).Columns(Cstr(temp)).Properties("Jet OLEDB:Allow Zero Length") = False
			End If
			
			valid = True
		End If
	Next	
	
	'****************************************************
	'* Check if atleast one field has been updated		*
	'****************************************************
	If valid = False Then
		'Invalid Field
		strError = "An invalid field name(s) has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Field", strError
	
		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5
	
		'Exit procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Redirect to table page							*
	'****************************************************
	'Response.Redirect "table.asp?table=" & table
End Sub

'****************************************************
'* Procedure to add a new index						*
'****************************************************
Sub AddIndex
	field = Request.QueryString("field")
	subaction = Request.QueryString("subaction")
	
	'****************************************************
	'* Check if the field in question actually exists	*
	'****************************************************
	If FieldExists(table, Trim(field)) = False Or IsNumeric(field) = True Then
		'Invalid Field
		strError = "An invalid field name has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Field", strError

		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5

		'Exit procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Check if the index doesn't already exist			*
	'****************************************************
	For Each index in db.objADOX.Tables(table).Indexes
		If LCase(index.Columns(0)) = LCase(fieldloop) Then
			'Index already exists
			strError = "An index on this field already exists in this table. " & _
			"Please go back and choose another field to create an index on."
		
			'Display error
			ErrorMessage "Index already exists", strError

			'Redirect to table
			JSRedirect "table.asp?table=" & table, 5

			'Exit procedure
			Exit Sub
		End If
	Next
	
	'****************************************************
	'* Check if this table already has a primary key	*
	'****************************************************
	If LCase(subaction) = "primary" Then
		For Each index in db.objADOX.Tables(table).Indexes
			If index.PrimaryKey = True Then
				'Already a primary key
				strError = "A primary key already exists in this table. " & _
				"Please go back and create a regular index on this field."
		
				'Display error
				ErrorMessage "Primary key already exists", strError

				'Redirect to table
				JSRedirect "table.asp?table=" & table, 5

				'Exit procedure
				Exit Sub
			End If
		Next
	End If
	
	'****************************************************
	'* Determine what to do, create an index or primary	*
	'****************************************************
	If LCase(subaction) = "primary" Then
		'Empty table
		strQuery = "DELETE FROM [" & table & "]"
		db.Query(strQuery)
	
		'Create primary key
		strQuery = "CREATE INDEX [PRIMARY] ON [" & table & "] ([" & field & "]) WITH PRIMARY"
	Else
		'Create ordinary index
		strQuery = "CREATE INDEX [" & field & "Index] ON [" & table & "] ([" & field & "])"
	End If
	db.Query(strQuery)
	
	'****************************************************
	'* Redirect back to table page						*
	'****************************************************
	Response.Redirect "table.asp?table=" & table		
End Sub

'****************************************************
'* Procedure to delete an index						*
'****************************************************
Sub DropIndex
	'****************************************************
	'* Extract index and check if it exists				*
	'****************************************************
	index = CStr(Request.QueryString("index"))

	If IndexExists(table, index) = False Or IsNumeric(index) = True Then
		'Invalid Index
		strError = "An invalid index name has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."
	
		'Display error
		ErrorMessage "Invalid index", strError
		
		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5
		
		'Exit procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Drop the index									*
	'****************************************************
	strQuery = "DROP INDEX [" & index & "] ON [" & table & "]"
	db.Query(strQuery)
	
	'****************************************************
	'* Redirect back to table page						*
	'****************************************************
	Response.Redirect "table.asp?table=" & table
End Sub

'****************************************************
'* Procedure for mass dropping or editing fields	*
'****************************************************
Sub DoFields
	'****************************************************
	'* Enter the redirect base url						*
	'****************************************************
	If LCase(Request.Form("submit")) = "drop" Then
		redirect = "table.asp?action=dropfield&table=" & table & "&field="
	Else
		redirect = "table.asp?action=editfield&table=" & table & "&field="
	End If
	
	'****************************************************
	'* Add field names to end of url					*
	'****************************************************
	For Each field In db.objADOX.Tables(table).Columns
		If IsBlank(Request.Form("selected_" & field.Name)) = False Then
			redirect = redirect & field.Name & ","
		End If
	Next
	
	'****************************************************
	'* Check if fields have been appended, and redirect	*
	'****************************************************
	If Len(redirect) <> Len("table.asp?action=editfield&table=" & table & "&field=") Then	
		Response.Redirect redirect
	Else
		Response.Redirect "table.asp?table=" & table
	End If
End Sub

'****************************************************
'* Procedure for renaming a table					*
'****************************************************
Sub RenameTable
	'****************************************************
	'* Check if the new name of this table doesn't		*
	'* already exists.									*
	'****************************************************
	If TableExists(Request.Form("name")) = True Then
		'Display Error
		strError = "Another table with this name already exists. " & _
		"Please go back and choose another name for this table."
	
		'Display error
		ErrorMessage "Table already exists", strError
		
		'Redirect to table
		JSRedirect "table.asp?table=" & table, 5
		
		'Exit procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Change name of table to new one					*
	'****************************************************
	If IsBlank(Request.Form("name")) = False Then
		db.objADOX.Tables(table).Name = Request.Form("name")
	End If
	
	'****************************************************
	'* Redirect to index, so table bar is refreshed		*
	'****************************************************
	Response.Redirect "index.asp"
End Sub

'****************************************************
'* Procedure for adding new field(s)				*
'****************************************************
Sub AddField
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
		   changetitle('<%= db.objADOX.Tables(table).Name %> running on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
		   -->
		</script>
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large">table <i><%= db.objADOX.Tables(table).Name %></i> running on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>

	<form name="addfield" method="post" action="table.asp?action=insertfield&table=<%= table %>">
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1">
				<th>Field</th>
				<th>Type</th>
				<th>Null*</th>
				<th>Default</th>
				<th>Extra**</th>
				<th>Size***</th>
			</tr>
	<%
	'****************************************************
	'* Validate number of fields to be added			*
	'****************************************************
	fieldloop = Request.Form("number")
	
	If IsNumeric(fieldloop) = False Then
		fieldloop = 1
	End If
	
	fieldloop = Clng(fieldloop)
	
	For i = 1 To fieldloop
		'****************************************************
		'* Display input boxes to enter values				*
		'****************************************************
		%>
		<tr>
			<td id="tdrow2">
				<input type="text" name="name_<%= i %>" id="textinput" size="10" maxlength="64" />
			</td>
			<td id="tdrow2">
				<select id="dropdown" name="type_<%= i %>">
					<%= column.CreateHTML("<option value=""$column->id"">$column->name</option>")%>
				</select>
			</td>
			<td id="tdrow2">
				<select id="dropdown" name="null_<%= i %>">
					<option value="no" selected>not null</option>
					<option value="yes">null</option>
				</select>
				</td>
			<td id="tdrow2">
				<input id="textinput" type="text" name="default_<%= i %>" size="8">
		       </td>
		       <td id="tdrow2">
		        <select id="dropdown" name="autoincrement_<%= i %>">
					<option value="no" selected></option>
					<option value="yes">auto_increment</option>
				</select>
			</td>
			<td id="tdrow2">
				<input id="textinput" type="text" name="DefinedSize_<%= i %>" size="8">
			</td>
		</tr>
		<%
	Next
	%>
	</table>
    <br />
    
	<input id="button" type="submit" name="submit" value="Add" />
	<input type="hidden" name="number" value="<%= fieldloop %>">
	</form>
	<table>
		<tr>
			<td valign="top">*&nbsp;</td>
			<td>
				<b>Certain fieldtypes do not allow null length.</b>
			</td>
		</tr>
		<tr>
			<td valign="top">**&nbsp;</td>
			<td>
				<b>Please note that when you choose auto_increment all the data in this table (<i><%= table %></i>) will be wiped.
				Also, if another field already auto increments, this field will not be changed to auto increment.
				</b>
			</td>
		</tr>
		<tr>
			<td valign="top">***&nbsp;</td>
			<td>
				<b>Please note that size field only apply to text-fields.
				</b>
			</td>
		</tr>
	</table>
	<br />
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>
	</body>

	</html>
	<%
End Sub

'****************************************************
'* Procedure for inserting new fields into table	*
'****************************************************
Sub InsertField
	'****************************************************
	'* Validate number of fields to be inserted			*
	'****************************************************
	fieldloop = Request.Form("number")
	
	If IsNumeric(fieldloop) = False Then
		fieldloop = 1
	End If
	
	fieldloop = Clng(fieldloop)
	
	strQuery = "ALTER TABLE [" & table & "]"
	
	first = True
	For i = 1 To fieldloop
		'****************************************************
		'* Make sure only valid fields get inserted			*
		'****************************************************
		If IsBlank(Request.Form("name_" & i)) = False Then
			'****************************************************
			'* Check if this field should be autoincrement		*
			'****************************************************
			If Request.Form("autoincrement_" & i) = "yes" Then
				valid = True
			
				'Check if other fields are already autoincrement
				For Each field In db.objADOX.Tables(table).Columns
					If field.Properties("AutoIncrement") = True Then
						valid = False
						Exit For
					End If
				Next
			
				If valid = True Then
					If db.objADOX.Tables(table).Columns.Count <> 0 Then
						'Empty table
						strQuery2 = "DELETE FROM [" & table & "]"
						db.Query(strQuery2)
					End If
					
					strQuery2 = "ALTER TABLE [" & table & "] ADD COLUMN "
					strQuery2 = strQuery2 & "[" & Request.Form("name_" & i) & "] autoincrement"
					db.Query(strQuery2)
				End If
			Else
				'****************************************************
				'* Add name to SQL query							*
				'****************************************************
				If first = True Then
					strQuery = strQuery & " ADD COLUMN [" & Request.Form("name_" & i) & "] "
					first = False
				Else
					strQuery = strQuery & ", [" & Request.Form("name_" & i) & "] "
				End If
				
				'// Add field type
				strQuery = strQuery & column.ColumnDatabase(Request.Form("type_" & fieldloop))
			
				'// Add field size (only on textfields)
				If IsNumeric(Request.Form("DefinedSize_" & i)) = True _
				And CLng(Request.Form("type_" & i)) = adVarWChar _
				Then
					'// Check if it's a valid size
					If Request.Form("DefinedSize_" & i) < 256 _
					And Request.Form("DefinedSize_" & i) > 0 _
					Then
						strQuery = strQuery & "(" & Request.Form("DefinedSize_" & fieldloop) & ")"
					End If
				End If
				
				'****************************************************
				'* Add default value of field to SQL query			*
				'****************************************************
				If IsBlank(Request.Form("default_" & i)) = False Then
					strQuery = strQuery & " DEFAULT " & ReplaceQuery(Request.Form("default_" & i))
				End If
			End If
		End If			
	Next

	'****************************************************
	'* Execute query									*
	'****************************************************
	If strQuery <> "ALTER TABLE [" & table & "]" Then
		db.Query(strQuery)
	End If
	
	'****************************************************
	'* Change NULL boolean manually						*
	'****************************************************
	For i = 1 To fieldloop
		On Error Resume Next
		
		If IsBlank(Request.Form("name_" & i)) = False _
		And Request.Form("autoincrement_" & i) <> "yes" _
		Then
			If Request.Form("null_" & i) = "yes" Then
				db.objADOX.Tables(table).Columns(Cstr(Request.Form("name_" & i))).Properties("Jet OLEDB:Allow Zero Length") = True
			Else
				db.objADOX.Tables(table).Columns(Cstr(Request.Form("name_" & i))).Properties("Jet OLEDB:Allow Zero Length") = False
			End If
		End If
		
		On Error GoTo 0
	Next

	'****************************************************
	'* Redirect back to table page						*
	'****************************************************
	Response.Redirect "table.asp?table=" & table
End Sub

'****************************************************
'* Call ending tasks procedure						*
'****************************************************
IncludeBottom
%>