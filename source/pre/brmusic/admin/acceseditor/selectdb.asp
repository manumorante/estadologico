<!-- #include file="includetop.asp" -->
<%
'****************************************************
'* Database Select List File						*
'****************************************************
'* This file is used to select a database path		*
'****************************************************
%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN"
		"http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" dir="ltr">
	<head>
		<title><%= name %></title>
		
		<link rel="stylesheet" type="text/css" href="extern/style.css">
	</head>
	
	<body bgcolor="#FFFFFF" class="bodyAdmin">
	
	<%	
	'// Declare some variables
	Dim f, f0, f1, fc, cur,parent
	
	Set objFSO = CreateObject ("Scripting.FileSystemObject")
	
	If Len(Request("current")) > 0 Then
		cur = Replace( Request("current") & "\","\\","\")
		If cur = "\\" Then cur = ""
		parent = ""
		
		If InStrRev(cur,"\") > 0 Then
			parent = Left(cur, InStrRev(cur, "\", Len(cur)-1))
		End If
	Else
		cur = ""
	End If
	
	Set f = objFSO.GetFolder(Server.Mappath("\") & "\" & cur)
%>
current:<%= cur %><BR>
[dir] <b><a href="selectdb.asp?current=<%= parent %>">..</a></b><br>
<%	Set fc = f.SubFolders
	For Each f1 In fc
  %>
[dir] <b><a href="selectdb.asp?current=<%= cur & f1.name %>"><%= f1.name %></a></b><br>
<% Next 
	Set fc = f.Files
	For Each f1 In fc
	If LCase(Right(f1.name,4)) = ".mdb" Then
%>
[database] <b><i><a href="javascript:opener.form1.dbPath.value='<%= Replace(cur & f1.name,"\","\\") %>';opener.form1.server.checked=true;self.close();"><%= f1.name %></i></a></b><br>
<%	End If
 Next 
 
Set objFSO = Nothing %>
<br><br>
Originally created by <i>Enrico Calderini</i><br>
Modified by <i>Dennis Pallett</i>
</body>
</html>

