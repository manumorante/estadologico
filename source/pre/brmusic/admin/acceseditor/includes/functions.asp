<%
'****************************************************
'* Functions File									*
'****************************************************
'* This file contains a set of shared functions		*
'* throughout aspAccessEditor.						*
'****************************************************

'****************************************************
'* Function for returning safe SQL queries			*
'****************************************************
Function ReplaceQuery(query)
	ReplaceQuery = Replace(query, "'", "''")
End Function

'****************************************************
'* Procedure for doing a Javascript Go Back			*
'****************************************************
Sub JSGoBack(Byval seconds)
	%>
	<script language="JavaScript">
	<!-- hide from old browser

	function goback()
	{
	history.back(1)
	}

	timer = setTimeout('goback()', '<%= seconds*1000 %>');
	
	-->
	</script>
	<%
End Sub

'****************************************************
'* Procedure for doing a JavaScript redirect, after	*
'* x seconds.										*
'****************************************************
Function JSRedirect(Byval url, seconds)
	%>
	<script language="JavaScript">
	<!--
	
	function redirect() {
	window.location = "<%= url %>";
	}

	timer = setTimeout('redirect()', '<%= seconds*1000 %>');
	
	-->
	</script>
	<%
End Function

'****************************************************
'* Function for checking if a value has to be		*
'* selected.										*
'****************************************************
Function SelectedData(strValue, strToSearchFor)
	If LCase(strValue) = LCase(strToSearchFor) Then
		SelectedData = " selected"
	End If
End Function

'****************************************************
'* Function for checking if a value has to be		*
'* checked.											*
'****************************************************
Function CheckedData(strValue, strToSearchFor)
	If LCase(strValue) = LCase(strToSearchFor) Then
		CheckedData = " checked"
	End If
End Function

'****************************************************
'* Function to check if a variable or array is empty*
'****************************************************
Function IsBlank(byref TempVar)
	IsBlank = False

	select Case VarType(TempVar)
	
		'Empty & Null
		case 0,1
			IsBlank = true
		
		'Strings
		case 8
			if Len(TempVar) = 0 Then
				IsBlank = True
			end If
		
		'Arrays
		case 8204,8209
			'does it have any dimensions?
			if UBound(TempVar) = -1 Then
				IsBlank = True
			end If

		'Anything else		
		case Else
			IsBlank = False
	end Select
End Function

'****************************************************
'* Function to check if a file exists. Should only	*
'* be used once a page.								*
'****************************************************
Function FileExists(Byval path)
	'Create FileSystemObject
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(path) Then
		'File exists
		FileExists = True
	Else
		'File does not exist
		FileExists = False
	End If

	'Cleanup
	Set objFSO = Nothing
End Function

'****************************************************
'* Procedure to delete file. Should only be used	*
'* once a page.										*
'****************************************************
Sub DeleteFile(Byval FilePath)
	'Create FileSystemObject
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	objFSO.DeleteFile(FilePath)

	'Cleanup
	Set objFSO = Nothing
End Sub

'****************************************************
'* Procedure to start ending tasks(sounds weird, ey)*
'****************************************************
Sub IncludeBottom
	'// Destroy database class
	Set db = Nothing
	'// Destroy column class
	Set column = Nothing
End Sub

'****************************************************
'* Function for checking if a table exists			*
'****************************************************
Function TableExists(Byval TableName)
	For Each tableloop in db.objADOX.Tables
		If LCase(TableName) = LCase(tableloop.Name) Then
			'Already exists
			TableExists = True
			Exit Function			
		End If
	Next
	
	TableExists = False
End Function

'****************************************************
'* Function to check if a field exists. It can check*
'* if a field exists in a certain table or in the	*
'* database.										*
'****************************************************
Function FieldExists(Byval TableName, Byval FieldName)
	Dim fieldtemp1, fieldtemp2

	'****************************************************
	'* Check if we have to check a certain table		*
	'****************************************************
	If IsBlank(TableName) = True Then
		'Check in whole database
		
		'Loop through tables
		For Each fieldtemp1 In db.objADOX.Tables
			'Loop through fields
			For Each fieldtemp2 In tableloop.Columns
				'Check if field exists
				If LCase(FieldName) = LCase(fieldtemp2.Name) Then
					FieldExists = True
					Exit Function
				End If
			Next			
		Next		
	Else
		'Check in specific table
		
		'Check if table exists
		If TableExists(TableName) = False Then
			FieldExists = False
			Exit Function
		End If
		
		'Loop through fields
		For Each fieldtemp1 In db.objADOX.Tables(TableName).Columns
			'Check if field exists
			If LCase(FieldName) = LCase(fieldtemp1.Name) Then
				FieldExists = True
				Exit Function
			End If
		Next
	End If
	
	FieldExists = False
End Function

'****************************************************
'* Function to check if an index exists				*
'****************************************************
Function IndexExists(Byval TableName, IndexName)
	'****************************************************
	'* Check if we have to check a certain table		*
	'****************************************************
	If IsBlank(TableName) = True Then
		'Check in whole database
		
		'Loop through tables
		For Each tableloop In db.objADOX.Tables
			'Loop through indexes
			For Each indexloop In tableloop.Indexes
				'Check if field exists
				If LCase(IndexName) = LCase(indexloop.Name) Then
					IndexExists = True
					Exit Function
				End If
			Next			
		Next		
	Else
		'Check in specific table
		
		'Check if table exists
		If TableExists(TableName) = False Then
			IndexExists = False
			Exit Function
		End If
		
		'Loop through indexes
		For Each indexloop In db.objADOX.Tables(TableName).Indexes
			'Check if index exists
			If LCase(IndexName) = LCase(indexloop.Name) Then
				IndexExists = True
				Exit Function
			End If
		Next
	End If
	
	IndexExists = False
End Function

'****************************************************
'* Function to strip out html tags					*
'****************************************************
Function HTMLSafe(Byval Variable)
	If IsBlank(Variable) = False Then
		HTMLSafe = Server.HTMLEncode(Variable)
	End If
End Function

'****************************************************
'* Function to retrieve a table from a SQL query	*
'****************************************************
Function GetTableFromSQL(Byval SQL)
	Dim charPos, charLen, wordlist

	'****************************************************
	'* Change SQL text to lowercase						*
	'****************************************************
	SQL = LCase(SQL)
	
	'****************************************************
	'* Enter words in array which might come after table*
	'****************************************************
	wordlist = Array("where", "inner", "left", "order", "right")

	'****************************************************
	'* Get the begin and length of the table name		*
	'****************************************************
	charPos = InStr(1, SQL, "from") + 5
	
	'****************************************************
	'* Loop through each word and check if it's there	*
	'****************************************************
	For Each fieldloop In wordlist
		charLen = InStr(charPos, SQL, " " & fieldloop)
		
		If charLen > 1 Then
			Exit For
		End If
	Next

	
	'****************************************************
	'* Retrieve table name								*
	'****************************************************
	If charLen > 0 Then
		SQL = Mid(SQL, charPos, charLen)
	Else
		SQL = Mid(SQL, charPos)
	End If
	
	'****************************************************
	'* Remove trailing and [ ] infront					*
	'****************************************************
	If Left(SQL, 1) = "[" Then SQL = Mid(SQL, 2)
	If Right(SQL, 1) = "]" Then SQL = Left(SQL, Len(SQL) - 1)
	
	GetTableFromSQL = SQL
End Function

'// Function to check if a view exists
Function ViewExists(Byval ViewName)
	For Each viewloop in db.objADOX.Views
		If LCase(ViewName) = LCase(viewloop.Name) Then
			'Already exists
			ViewExists = True
			Exit Function			
		End If
	Next
	
	ViewExists = False
End Function
%>