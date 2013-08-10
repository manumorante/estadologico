<%@ Language=VBScript %>
<%
'****************************************************
'* Header Include File								*
'****************************************************
'* This file is included in every main file at the	*
'* top. This does all the preparation.				*
'****************************************************

'****************************************************
'* Issue some begin commands, e.g set buffer and	*
'* expiration data.									*
'****************************************************
Option Explicit
Response.Buffer = True
Response.Expires = -10
Response.ExpiresAbsolute = DateAdd("d", -1, Now())

'****************************************************
'* Include database class, functions, adovbs and	*
'* error function.									*
'****************************************************
%>
<!-- #include file="includes/pwdprotect.asp" -->
<!-- #include file="classes/database.asp" -->
<!-- #include file="includes/err.asp" -->
<!-- #include file="includes/functions.asp" -->
<!-- #include file="includes/adovbs.inc" -->
<!-- #include file="classes/columns.asp" -->
<%
'****************************************************
'* Make a check so these files are excepted			*
'****************************************************
If Right(Request.ServerVariables("SCRIPT_NAME"), Len("createdb.asp")) <> "createdb.asp" _
And Right(Request.ServerVariables("SCRIPT_NAME"), Len("loaddb.asp")) <> "loaddb.asp" _
And Right(Request.ServerVariables("SCRIPT_NAME"), Len("selectdb.asp")) <> "selectdb.asp" _
Then
	'****************************************************
	'* Check if database path is empty					*
	'****************************************************
	If IsBlank(Session("dbPath")) = True Then
		Response.Redirect "loaddb.asp"
		Response.End
	End If

	'****************************************************
	'* Create instance of database class				*
	'****************************************************
	Dim db

	Set db = New DBConnect

	On Error Resume Next

	'****************************************************
	'* Connect to database, and also immediately check	*
	'* if a connection was made.						*
	'****************************************************
	If db.Connect(Session("dbUser"), Session("dbPassword"), Session("dbPath")) = False Then
		'Display an error
		strError = "A connection to the database hasn't been established. " & _
		"Please try again by <a href=""loaddb.asp"">clicking here</a> or waiting a few seconds for a redirect."
		ErrorMessage "Connection Failure", strError
	
		'Finish off tasks
		IncludeBottom
	
		'Redirect to other page
		JSRedirect "loaddb.asp", 5
	
		'Finish of processing this page
		Response.End
	End If
	
	On Error GoTo 0

	'****************************************************
	'* Set cursor types and cache size					*
	'****************************************************
	db.SetCursors adUseClient, adOpenStatic, 30
End If
	
'****************************************************
'* Declare global variables							*
'****************************************************
Dim objFSO, action, i, strError
Dim database, table, field, record, records
Dim strQuery, tableloop, index, fieldloop, first, recordloop
Dim strQuery2, indexloop, subaction, redirect, valid
Dim intCurrentPage, intPageSize, intTotalPages, intRecordCount
Dim strPagingLink, whereclause, temp, objFile, temparray
Dim intI, view, objCommand, viewloop, column, forloop

'// Create column class
Set column = New ColumnTypes

'****************************************************
'* Add two global constants used in aspAccessEditor	*
'****************************************************
Const name = "aspAccessEditor"
Const version = "2.0"
%>