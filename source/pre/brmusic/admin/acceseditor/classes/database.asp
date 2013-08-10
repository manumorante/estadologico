<%
'****************************************************
'* Database Class									*
'****************************************************
'* This class is used for all the database work in  *
'* aspAccessEditor.									*
'****************************************************

'****************************************************
'* Security check: Check if this page has been		*
'* called directly (thus someone tries to exploit	*
'* something). This is an include file so it should	*
'* not be called directly, thus this security check	*
'****************************************************
If Right(Request.ServerVariables("SCRIPT_NAME"), Len("database.asp")) = "database.asp" Then
	Response.Write "What are you trying to find here? There's nothing here, understood?"
	Response.End
End If

Class DBConnect
	Public dbQueriesUsed
	Public objConn
	Public objRS
	Public strQuery
	Public strConnection
	Public intCurrentPage
	Public intTotalPages
	Public intPageSize
	Public strPagingLink
	Public strPaging
	Public objADOX
	
	'****************************************************
	'* This procedure is triggered when the class is	*
	'* initialized.										*
	'****************************************************
	Private Sub Class_Initialize()
		'****************************************************
		'* Create connection, record set object and ADOX 	*
		'* object.											*
		'****************************************************
		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objRS = Server.CreateObject("ADODB.Recordset")
		Set objADOX = Server.CreateObject("ADOX.Catalog")
	End Sub

	'****************************************************
	'* This procedure is triggered when the class is	*
	'* closed/ended.									*
	'****************************************************
	Private Sub Class_Terminate()
		'****************************************************
		'* Destroy connection, record set object and ADOX	*
		'* object.											*
		'****************************************************
		If IsObject(objConn) Then
			Set objConn = Nothing
		End If
		If IsObject(objRS) Then
			Set objRS = Nothing
		End If	
		If IsObject(objADOX) Then
			Set objADOX = Nothing
		End If	
	End Sub
	
	'****************************************************
	'* This function is used to connect to a database.	*
	'* Returns true when a connection has been made.	*
	'* Returns false when a connection hasn't been made.*
	'* Use it like this:								*
	'* -												*
	'If dbstuff.Connect("access", "dbuser", "dbpassword", "dbserverORpath") = False Then
	'Response.Write "Failed to connect to database!"		
	'Response.End										
	'End If												
	'* -																	
	'****************************************************
	Public Function Connect(strUser, strPassword, strServer)
		strConnection = "Provider=MICROSOFT.JET.OLEDB.4.0; DATA SOURCE=" & strServer & ";"
		If strUser <> "" Then
			strConnection = strConnection & " Uid=" & strUser & ";"
		End If
		If strPassword <> "" Then
			strConnection = strConnection & " Pwd=" & strPassword & ";"
		End If
	
		'****************************************************
		'* Open connection									*
		'****************************************************
		objConn.Open strConnection		
		
		'****************************************************
		'* Set active connection of the ADOX object			*
		'****************************************************
		objADOX.ActiveConnection = objConn
		
		Connect = True
	End Function
	
	'****************************************************
	'* Close the connection to the database				*
	'****************************************************
	Public Sub Disconnect
		objConn.Close
	End Sub
	
	'****************************************************
	'* Set recordset cursors							*
	'****************************************************
	Public Sub SetCursors(CursorLocation, CursorType, CacheSize)
		If IsBlank(CursorLocation) = False Then
			objRS.CursorLocation = CursorLocation
		End If
		If IsBlank(CursorType) = False Then
			objRS.CursorType = CursorType
		End If
		If IsBlank(CacheSize) = False Then
			objRS.CacheSize = CacheSize
		End If
	End Sub
	
	'****************************************************
	'* Execute a SQL query to the database				*
	'****************************************************
	Public Sub Query(strQuery)
		On Error Resume Next
		
		objRS.Open CStr(strQuery), objConn, adOpenKeyset , , adCmdText
		dbQueriesUsed = dbQueriesUsed + 1
		
		'****************************************************
		'* Check if any errormessages occured				*
		'****************************************************
		If Err.number <> 0 Then
			SQLError Err.number, Err.Description, Err.Source, strQuery
			
			IncludeBottom
			Response.End
		End If
		
		On Error Goto 0
	End Sub
	
	'****************************************************
	'* Close RecordSet									*
	'****************************************************
	Public Sub QueryClose
		objRS.Close
	End Sub
	
	
	'****************************************************
	'* This function is used for paging. Very useful!	*
	'****************************************************
	Function Paging(Byval strType, NumberOf, CurrentPage, TotalPages)
		SELECT CASE strType
		CASE "previous"
			If CurrentPage - NumberOf => 1 Then
				Paging = True
			End If
		CASE "next"
			If CurrentPage + NumberOf =< TotalPages Then
				Paging = True
			End If
		END SELECT
	End Function
End Class
%>
