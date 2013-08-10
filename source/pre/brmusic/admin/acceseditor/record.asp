<!-- #include file="includetop.asp" -->
<%
'****************************************************
'* Record File										*
'****************************************************
'* This file is used to display a list of records	*
'* and individual records. It also allows for		*
'* inserting, editing and deleting records.			*
'****************************************************

'****************************************************
'* Check if a valid SQL query has been passed		*
'****************************************************
strQuery = Request.QueryString("sql")

If IsBlank(strQuery) = True Or IsNumeric(strQuery) = True Then
	'Invalid SQL
	strError = "An invalid SQL query has been passed to this page. " & _
	"Please go back and try again. If you clicked a valid link, please notify the system administrator."

	'Display error
	ErrorMessage "Invalid SQL", strError
	
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
	CASE "addrecord"
		AddRecord
	CASE "insertrecord"
		InsertRecord
	CASE "editrecord"
		EditRecord
	CASE "updaterecord"
		UpdateRecord
	CASE "deleterecord"
		DeleteRecord
	CASE Else
		DisplayList
END SELECT

'****************************************************
'* Procedure for displaying a list of records		*
'****************************************************
Sub DisplayList
	'****************************************************
	'* Make sure a select query is being done			*
	'****************************************************
	If UCase(Left(strQuery, 6)) <> "SELECT" Then
		SQLError 1, "Invalid SELECT query. Only SELECT queries are allowed for this page.", name, strQuery
		Exit Sub
	End If

	'****************************************************
	'* Execute SQL query								*
	'****************************************************
	db.Query(strQuery)
		
	'****************************************************
	'* Put all the field names in an array				*
	'****************************************************
	For Each field in db.objRS.Fields
		fieldloop = fieldloop & Replace(field.Name, ",", ":comma:") & ","
	Next
	
	'- Cut off final comma
	fieldloop = Left(fieldloop, Len(fieldloop)-1)
	
	'- Add all to array again
	field = Split(fieldloop, ",")
		
	'****************************************************
	'* Arrange paging stuff								*
	'****************************************************
	'- Arrange current oage
	intCurrentPage = Request.QueryString("page")
	
	If IsBlank(intCurrentPage) = True or IsNumeric(intCurrentPage) = False Or intCurrentPage < 1 Then
		intCurrentPage = 1
	End If
	
	'- Arrange page size
	intPageSize = Request.QueryString("recordsperpage")

	If IsBlank(intPageSize) Or IsNumeric(intPageSize) = False Then
		intPageSize = 30
	End If

	If db.objRS.EOF = False Then
		'- Tell RecordSet to use these values
		db.objRS.PageSize = intPageSize

		db.objRS.AbsolutePage = intCurrentPage

		'- Retrieve total pages count
		intTotalPages = db.objRS.PageCount
	
		'- Retrieve recordcount
		intRecordCount = db.objRS.RecordCount
	
		'- Retrieve records and put them in an array
		records = db.objRS.GetRows
	End If
	
	'****************************************************
	'* Close SQL query									*
	'****************************************************
	db.QueryClose

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
			changetitle('SQL record list running on <%= Request.ServerVariables("SERVER_NAME") %> - <%= name & " " & version %>');
			//-->
		</script>

    
	</head>


	<body bgcolor="#FFFFD9" class="bodyAdmin">
	<div id="large"><i>SQL record list</i> running on <i><%= Request.ServerVariables("SERVER_NAME") %></i></div>
	<br>
	
	<div align="left">
    <table border="0" cellpadding="5" id="memgroup" width="100%">
    <tr>
        <td id="tdrow1">
            <b>Showing rows (<%= intRecordCount %> total)</b><br />
        </td>
    </tr>
    
    <tr>
        <td id="tdrow2">
            
            SQL-query&nbsp;:&nbsp;<a href="table.asp?table=<%= GetTableFromSQL(strQuery) %>&sql=<%= strQuery %>#sql">Edit</a><br />
            <%= strQuery %>
        </td>
    </tr>
           
    </table>
	</div><br />
	
	<%
	'****************************************************
	'* Check if no rows were returned, and display msg	*
	'****************************************************
	If intRecordCount < 1 Then
		'No records
		Exit Sub	
	End If
	%>
	<!-- Results table -->
	<table border="0" cellpadding="5" id="memgroup" width="100%">
        
	<!-- Results table headers -->
	<tr>
		<td id="tdrow1"></td>
		<td id="tdrow1"></td>
		<%
		For Each fieldloop In field
			%>
		    <th id="tdrow1">
			<%= fieldloop %>
		    </th>
			<%
		Next
		%>                   
	</tr>
	
	<!-- Results table body -->
	<%
	'****************************************************
	'* Check which to use, pagesize of ubound of records*
	'****************************************************
	If UBound(records, 2) <= intPageSize Then
		recordloop = UBound(records, 2)
	Else
		recordloop = intPageSize
	End If
		
	For record = 0 To recordloop
		%>
		<tr>
			<td id="tdrow2">
				<a href="javascript:document.editrecord_<%= record %>.submit();">Edit</a>
		        <form name="editrecord_<%= record %>" method="post" action="record.asp?action=editrecord&table=<%= GetTableFromSQL(strQuery) %>&sql=bypassfilter">
		        <%
		        i = 0
		        For Each fieldloop in field
					Response.Write "<input type=""hidden"" name=""" & fieldloop & """ value=""" & HTMLSafe(records(i, record)) & """>" & vbCrlf
					i = i + 1
		        Next
		        %>	
		        </form>
		    </td>
              
			<td id="tdrow2">
				<a href="javascript:document.deleterecord_<%= record %>.submit();">Delete</a>
		        <form name="deleterecord_<%= record %>" method="post" action="record.asp?action=deleterecord&table=<%= GetTableFromSQL(strQuery) %>&sql=bypassfilter">
				<%
		        i = 0
		        For Each fieldloop in field
					Response.Write "<input type=""hidden"" name=""" & fieldloop & """ value=""" & HTMLSafe(records(i, record)) & """>" & vbCrlf
					i = i + 1
		        Next
		        %>		        
		        </form>
			</td>
		<%
		i = 0
		For Each fieldloop In field
			Response.Write "<td align=""left"" valign=""top"" id=""tdrow2"">"
			
			Response.Write HTMLSafe(Left(records(i, record), 350))
			
									
			Response.Write "</td>"
			i = i + 1
		Next
		%>
		</tr>
		<%
	Next
	%>
	</table>
	<br />

	<!-- Paging Code -->
	<%
	strPagingLink = "record.asp?sql=" & strQuery

	If db.Paging("previous", 2, intCurrentPage, intTotalPages) = True Then
		Response.Write "<a href=""" & strPagingLink & "&page=1"" title=""First Page"">« First</a>&nbsp;"
		
		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage - 1 & """ title=""Previous Page"">«</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage - 2 & """>" & intCurrentPage - 2  & "</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage - 1 & """>" & intCurrentPage - 1  & "</a>&nbsp;"
	ElseIf db.Paging("previous", 1, intCurrentPage, intTotalPages) = True Then
		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage - 1 & """ title=""Previous Page"">«</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage - 1 & """>" & intCurrentPage - 1 & "</a>&nbsp;"
	End If

	If db.Paging("previous", 1, intCurrentPage, intTotalPages) = True _
	Or db.Paging("next", 1, intCurrentPage, intTotalPages) = True Then
		Response.Write "<b>[" & intCurrentPage & "]&nbsp;</b>"
	End If


	If db.Paging("next", 2, intCurrentPage, intTotalPages) = True Then
		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage + 1 & """>" & intCurrentPage + 1 & "</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage + 2 & """>" & intCurrentPage + 2 & "</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage + 1 & """ title=""Next Page"">»</a>&nbsp;"

		Response.Write "<a href=""" & strPagingLink & "&page=" & intTotalPages & """ title=""Last Page"">Last »</a>&nbsp;"

	ElseIf db.Paging("next", 1, intCurrentPage, intTotalPages) = True Then
	Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage + 1 & """>" & intCurrentPage + 1 & "</a>&nbsp;"

	Response.Write "<a href=""" & strPagingLink & "&page=" & intCurrentPage + 1 & """ title=""Next Page"">»</a>&nbsp;"
	End If
	%>
  
	<!-- Insert a new row -->
	<p>
		<a href="record.asp?action=addrecord&sql=bypassfilter&table=<%= GetTableFromSQL(strQuery) %>">
			Insert New Row
		</a>
	</p>
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>


	</body>

	</html>
	<%
End Sub

'****************************************************
'* Procedure for adding a new record				*
'****************************************************
Sub AddRecord
	'****************************************************
	'* Check if a valid table name has been passed		*
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
	
		'Redirect to index
		JSRedirect "index.asp", 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Check if this table has fields					*
	'****************************************************
	If db.objADOX.Tables(table).Columns.Count < 1 Then
		'No fields
		strError = "This table contains no fields, and so you cannot add a new record. " & _
		"Please go back and add a field to this table."

		'Display error
		ErrorMessage "No fields", strError
	
		'Redirect to index
		JSRedirect "table.asp?table=" & table, 5
	
		'End this procedure
		Exit Sub
	End If
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
	
	<!-- Add record field properties form -->
	<form method="post" action="record.asp?action=insertrecord&table=<%= table %>&sql=bypassfilter" name="insertrecord">
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1">
				<th>Field</th>
				<th>Type</th>
				<th>Null</th>
				<th>Value</th>
			</tr>
	<%
	'****************************************************
	'* Loop through each field							*
	'****************************************************
	For Each field In db.objADOX.Tables(table).Columns
	%>
		<tr>
			<td align="center" id="tdrow2"><%= field.Name %></td>
    
			<td align="center" id="tdrow2" nowrap="nowrap">
				<%
				'// Display the field name			
				Response.Write column.ColumnName(field.Type)
				%>
			</td>
			<td align="center" id="tdrow2">
				<%
				'****************************************************
				'* Display if a NULL value is allowed or not		*
				'****************************************************
				If field.Properties("Jet OLEDB:Allow Zero Length") = True Then
					Response.Write "Allowed"
				Else
					Response.Write "&nbsp;"
				End If
				%>
			</td>
			<td align="center" id="tdrow2">
				<%
				'****************************************************
				'* Display the appropriate input box				*
				'****************************************************
				If field.Properties("AutoIncrement") = True Then
					Response.Write "[autoincrement]"
				Else
					SELECT CASE column.ColumnType(field.Type)
						CASE "textinput"
							%><input id="textinput" type="text" name="<%= field.Name %>" size="15" tabindex="1" /><%
						CASE "textarea"
							%><textarea tabindex="1" id="multitext" cols="30" rows="7" name="<%= field.Name %>"></textarea><%
						CASE "yesno"
							Response.Write "<input tabindex=""1"" type=""radio"" name=""" & field.Name & """ value=""True"">True" & vbCrlf
							Response.Write "<input tabindex=""1"" type=""radio"" name=""" & field.Name & """ value=""False"">False" & vbCrlf
						CASE Else
							Response.Write "[" & column.ColumnName(field.Type) & "]"
					END	SELECT
				End If
				%>
			</td>
        
		</tr>
		<%	
	Next
	%>
			<tr>
				<td colspan="4" align="center" valign="middle">
					<input id="button" type="submit" value="Go" tabindex="29" />
				</td>
			</tr>
		</table>
	</form>
	
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>

	</body>

	</html>
	<%
End Sub

'****************************************************
'* Procedure to insert a new record					*
'****************************************************
Sub InsertRecord
	'****************************************************
	'* Check if a valid table name has been passed		*
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
	
		'Redirect to index
		JSRedirect "index.asp", 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Check if this table has fields					*
	'****************************************************
	If db.objADOX.Tables(table).Columns.Count < 1 Then
		'No fields
		strError = "This table contains no fields, and so you cannot add a new record. " & _
		"Please go back and add a field to this table."

		'Display error
		ErrorMessage "No fields", strError
	
		'Redirect to index
		JSRedirect "table.asp?table=" & table, 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Begin query and loop through fields				*
	'****************************************************
	strQuery = "INSERT INTO [" & table & "]"
	
	first = True
	For Each field in db.objADOX.Tables(table).Columns
		'****************************************************
		'* Check if the non-NULL fields haven been entered	*
		'****************************************************
		If field.Properties("Jet OLEDB:Allow Zero Length") = False _
		And IsBlank(Request.Form(field.Name)) = True _
		And field.Properties("AutoIncrement") = False _
		Then
			'Empty field
			strError = "Please fill in all the fields which do not allow a NULL value. " & _
			"Please go back and try again."

			'Display error
			ErrorMessage "Empty field(s)", strError
	
			'Redirect back
			JSGoBack(5)
	
			'End this procedure
			Exit Sub
		End If
		
		'****************************************************
		'* Add fields to SQL query							*
		'****************************************************
		If field.Properties("AutoIncrement") = False _
		And column.ColumnQuotes(field.Type) <> "no" Then
			If first = True Then
				strQuery = strQuery & " ("
				first = False
			Else
				strQuery = strQuery & ", "
			End If
		
			strQuery = strQuery & "[" & field.Name & "]"
		End If
	Next
	
	'****************************************************
	'* Add to SQL query and loop through fields again	*
	'****************************************************
	strQuery = strQuery & ") VALUES "
	
	first = True
	For Each field in db.objADOX.Tables(table).Columns
		'****************************************************
		'* Add to SQL query, depending on the field type	*
		'****************************************************
		If field.Properties("AutoIncrement") = False _
		And column.ColumnQuotes(field.Type) <> "no" Then
			If first = True Then
				strQuery = strQuery & " ("
			Else
				strQuery = strQuery & ", "
			End If
		
			SELECT CASE column.ColumnQuotes(field.Type)
				CASE ""
					strQuery = strQuery & ReplaceQuery(Request.Form(field.Name))
				CASE "quotes"
					strQuery = strQuery & "'" & ReplaceQuery(Request.Form(field.Name)) & "'"
				CASE "date"
					strQuery = strQuery & "#" & ReplaceQuery(Request.Form(field.Name)) & "#"
				CASE Else
					'// Invalid field, delete ( or ,
					If first = True Then
						strQuery = Left(strQuery, Len(strQuery) - Len(" ("))
					Else
						strQuery = Left(strQuery, Len(strQuery) - Len(", "))
					End If
			END	SELECT
			
			first = False
		End If
	Next
	
	'****************************************************
	'* Finish off SQL query								*
	'****************************************************
	strQuery = strQuery & ")"
			
	'****************************************************
	'* Execute SQL query								*
	'****************************************************
	db.Query(strQuery)
	
	'****************************************************
	'* Redirect back to record page						*
	'****************************************************
	Response.Redirect "record.asp?sql=SELECT * FROM [" & table & "]"
End Sub

'****************************************************
'* Procedure for editing a record					*
'****************************************************
Sub EditRecord
	'****************************************************
	'* Check if a valid table name has been passed		*
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
	
		'Redirect to index
		JSRedirect "index.asp", 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Begin SQL query									*
	'****************************************************
	strQuery = "SELECT * FROM [" & table & "]"
	
	'// Check if this table has any primary keys,
	'// and if it has, use that
	For Each index In db.objADOX.Tables(table).Indexes
		If index.PrimaryKey = True Then
			strQuery = strQuery & " WHERE [" & _
			index.Columns(0) & "] = "
				
			'// Determine fieldtype
			SELECT CASE column.ColumnQuotes(db.objADOX.Tables(table).Columns(CStr(index.Columns(0))).Type)
				CASE ""
					strQuery = strQuery & ReplaceQuery(Request.Form(CStr(index.Columns(0))))
					valid = True
				CASE "quotes"
					strQuery = strQuery & "'" & ReplaceQuery(Request.Form(CStr(index.Columns(0)))) & "'"
					valid = True
				CASE Else
					'// Delete where bit again
					strQuery = "SELECT * FROM [" & table & "]"
			END SELECT
		End If
	Next
		
	'// Check if there is already a whereclause
	If valid = False Then '// Build alternate WHERE clause
		first = True
		For Each field in db.objADOX.Tables(table).Columns
			If IsBlank(Request.Form(field)) = False Then
				'// Check if this is the first field
				If InStr(1, LCase(strQuery), "where") > 0 Then
					strQuery = strQuery & " AND "
				Else
					strQuery = strQuery & " WHERE "	
				End If
					
				'// Add field value to SQL query			
				SELECT CASE column.ColumnQuotes(field.Type)
					CASE ""
						strQuery = strQuery & "[" & field.Name & "] = " & ReplaceQuery(Request.Form(field.Name))
					CASE "quotes"
						strQuery = strQuery & "[" & field.Name & "] = " & "'" & ReplaceQuery(Request.Form(field.Name)) & "'"
					CASE Else
						'// Invalid field, delete WHERE or AND
						If InStr(1, LCase(strQuery), "where") > 0 Then
							strQuery = Left(strQuery, Len(strQuery) - Len(" AND "))
						Else
							strQuery = Left(strQuery, Len(strQuery) - Len(" WHERE "))
						End If
				END	SELECT
			End If
		Next
	End If
	
	'****************************************************
	'* Execute SQL query								*
	'****************************************************	
	db.Query(strQuery)
			
	'****************************************************
	'* Check if a record has been returned				*
	'****************************************************
	If db.objRS.EOF = True Then
		'Invalid record, display error
		strError = "Invalid record details have been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Record", strError
	
		'Redirect to index
		JSRedirect "record.asp?sql=SELECT * FROM [" & table & "]", 5
	
		'End this procedure
		Exit Sub
	End If
	
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
	
	
	<!-- Edit record field properties form -->
	<form method="post" action="record.asp?action=updaterecord&table=<%= table %>&sql=bypassfilter" name="updaterecord">
	<input type="hidden" name="where" value="<%= HTMLSafe(Replace(strQuery, "SELECT * FROM [" & table & "] ", "")) %>">
		<table border="0" id="memgroup" width="100%">
			<tr id="tdrow1">
				<th>Field</th>
				<th>Type</th>
				<th>Null</th>
				<th>Value</th>
			</tr>
	<%
	'****************************************************
	'* Loop through each field							*
	'****************************************************
	For Each field In db.objADOX.Tables(table).Columns
	%>
		<tr>
			<td align="center" id="tdrow2"><%= field.Name %></td>
    
			<td align="center" id="tdrow2" nowrap="nowrap">
				<%
				'// Display the field type
				Response.Write column.ColumnName(field.Type)
				%>
			</td>
			<td align="center" id="tdrow2">
				<%
				'****************************************************
				'* Display if a NULL value is allowed or not		*
				'****************************************************
				If field.Properties("Jet OLEDB:Allow Zero Length") = True Then
					Response.Write "Allowed"
				Else
					Response.Write "&nbsp;"
				End If
				%>
			</td>
			<td align="center" id="tdrow2">
				<%
				'****************************************************
				'* Display the appropriate input box				*
				'****************************************************
				If field.Properties("AutoIncrement") = True Then
					Response.Write "[autoincrement]"
				Else
					SELECT CASE column.ColumnType(field.Type)
						CASE "textinput"
							%><input value="<%= db.objRS(field.Name) %>" id="textinput" type="text" name="<%= field.Name %>" size="15" tabindex="1" /><%
						CASE "textarea"
							%><textarea tabindex="1" id="multitext" cols="30" rows="7" name="<%= field.Name %>"><%= db.objRS(field.Name) %></textarea><%
						CASE "yesno"
							Response.Write "<input tabindex=""1"" type=""radio"" name=""" & field.Name & """ value=""True""" & CheckedData(db.objRS(field.Name), "True") & ">True" & vbCrlf
							Response.Write "<input tabindex=""1"" type=""radio"" name=""" & field.Name & """ value=""False""" & CheckedData(db.objRS(field.Name), "False") & ">False" & vbCrlf
						CASE Else
							Response.Write "[" & field.Type & "]"
					END	SELECT
				End If
				%>
			</td>
        
		</tr>
		<%
	Next
	'****************************************************
	'* Close Recordset									*
	'****************************************************
	db.QueryClose
	%>
			<tr>
				<td colspan="4" align="center" valign="middle">
					<input id="button" type="submit" value="Go" tabindex="29" />
				</td>
			</tr>
		</table>
	</form>
	
	<div align="center">Powered by <%= name & " " & version %><br>Copyright ©2002-2003 Dennis Pallett (<a href="http://www.aspit.net" target="_blank">AspIt</a>)</div>

	</body>

	</html>
	<%
End Sub

'****************************************************
'* Procedure used for updating a record				*
'****************************************************
Sub UpdateRecord
	'****************************************************
	'* Check if a valid table name has been passed		*
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
	
		'Redirect to index
		JSRedirect "index.asp", 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Begin SQL query									*
	'****************************************************
	strQuery = "UPDATE [" & table & "]"
	
	first = True
	For Each field in db.objADOX.Tables(table).Columns
		'// Check if the non-NULL fields haven been entered		
		If field.Properties("Jet OLEDB:Allow Zero Length") = False _
		And IsBlank(Request.Form(field.Name)) = True _
		And field.Properties("AutoIncrement") = False _
		Then
			'Empty field
			strError = "Please fill in all the fields which do not allow a NULL value. " & _
			"Please go back and try again."

			'Display error
			ErrorMessage "Empty field(s)", strError
	
			'Redirect back
			JSGoBack(5)
	
			'End this procedure
			Exit Sub
		End If
	
		If field.Properties("AutoIncrement") = False _
		And column.ColumnQuotes(field.Type) <> "no" Then
			'// Check if this is the first field
			If first = True Then
				strQuery = strQuery & " SET "
				first = False
			Else
				strQuery = strQuery & ", "
			End If
		
			'// Add field value to SQL query
			SELECT CASE column.ColumnQuotes(field.Type)
				CASE ""
					strQuery = strQuery & "[" & field.Name & "] = " & ReplaceQuery(Request.Form(field.Name))
				CASE "quotes"
					strQuery = strQuery & "[" & field.Name & "] = " & "'" & ReplaceQuery(Request.Form(field.Name)) & "'"
				CASE "date"
					strQuery = strQuery & "[" & field.Name & "] = " & "#" & ReplaceQuery(Request.Form(field.Name)) & "#"
				CASE Else
					'// Remove SET or ,					
					If first = True Then
						strQuery = Left(strQuery, Len(strQuery) - Len(" SET "))
					Else
						strQuery = Left(strQuery, Len(strQuery) - Len(", "))
					End If
			END	SELECT
		End If	
	Next
	
	
	'****************************************************
	'* Retrieve whereclause and validate it				*
	'****************************************************
	whereclause = Request.Form("where")
	
	If IsBlank(whereclause) = True _
	Or Len(whereclause) < 5 Then
		'Invalid Where Clause
		strError = "An invalid where clause has been passed to this page. " & _
		"Please go back and try again. If you clicked a valid link, please notify the system administrator."

		'Display error
		ErrorMessage "Invalid Where Clause", strError
	
		'Redirect to index
		JSGoBack 5
	
		'End this procedure
		Exit Sub
	End If
	
	
	'****************************************************
	'* Finish off SQL query								*
	'****************************************************
	strQuery = strQuery & " " & whereclause
	
	'****************************************************
	'* Execute SQL query								*
	'****************************************************	
	db.Query(strQuery)
	
	'****************************************************
	'* Redirect back to record page						*
	'****************************************************
	Response.Redirect "record.asp?sql=SELECT * FROM [" & table & "]"
End Sub

'****************************************************
'* Procedure for deleting a record					*
'****************************************************
Sub DeleteRecord
	'****************************************************
	'* Check if a valid table name has been passed		*
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
	
		'Redirect to index
		JSRedirect "index.asp", 5
	
		'End this procedure
		Exit Sub
	End If
	
	'****************************************************
	'* Begin SQL query									*
	'****************************************************
	strQuery = "DELETE FROM [" & table & "]"
	
	'// Check if this table has any primary keys,
	'// and if it has, use that
	For Each index In db.objADOX.Tables(table).Indexes
		If index.PrimaryKey = True Then
			strQuery = strQuery & " WHERE [" & _
			index.Columns(0) & "] = "
				
			'// Determine fieldtype
			SELECT CASE column.ColumnQuotes(db.objADOX.Tables(table).Columns(CStr(index.Columns(0))).Type)
				CASE ""
					strQuery = strQuery & ReplaceQuery(Request.Form(CStr(index.Columns(0))))
					valid = True
				CASE "quotes"
					strQuery = strQuery & "'" & ReplaceQuery(Request.Form(CStr(index.Columns(0)))) & "'"
					valid = True
				CASE Else
					'// Delete where bit again
					strQuery = "SELECT * FROM [" & table & "]"
			END SELECT
		End If
	Next
		
	'// Check if there is already a whereclause
	If valid = False Then '// Build alternate WHERE clause
		first = True
		For Each field in db.objADOX.Tables(table).Columns
			If IsBlank(Request.Form(field)) = False Then
				'// Check if this is the first field
				If InStr(1, LCase(strQuery), "where") > 0 Then
					strQuery = strQuery & " AND "
				Else
					strQuery = strQuery & " WHERE "	
				End If
					
				'// Add field value to SQL query			
				SELECT CASE column.ColumnQuotes(field.Type)
					CASE ""
						strQuery = strQuery & "[" & field.Name & "] = " & ReplaceQuery(Request.Form(field.Name))
					CASE "quotes"
						strQuery = strQuery & "[" & field.Name & "] = " & "'" & ReplaceQuery(Request.Form(field.Name)) & "'"
					CASE Else
						'// Invalid field, delete WHERE or AND
						If InStr(1, LCase(strQuery), "where") > 0 Then
							strQuery = Left(strQuery, Len(strQuery) - Len(" AND "))
						Else
							strQuery = Left(strQuery, Len(strQuery) - Len(" WHERE "))
						End If
				END	SELECT
			End If
		Next
	End If
	
	'****************************************************
	'* Execute SQL query								*
	'****************************************************
	db.Query(strQuery)
	
	'****************************************************
	'* Redirect back to record page						*
	'****************************************************
	Response.Redirect "record.asp?sql=SELECT * FROM [" & table & "]"
End Sub


'****************************************************
'* Call ending tasks procedure						*
'****************************************************
IncludeBottom
%>
