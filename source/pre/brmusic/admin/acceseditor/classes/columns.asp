<%
'****************************************************
'* Column Types Class								*
'****************************************************
'* This class is used for all the column types in	*
'* aspAccessEditor.									*
'****************************************************

'****************************************************
'* Security check: Check if this page has been		*
'* called directly (thus someone tries to exploit	*
'* something). This is an include file so it should	*
'* not be called directly, thus this security check	*
'****************************************************
If Right(Request.ServerVariables("SCRIPT_NAME"), Len("columns.asp")) = "columns.asp" Then
	Response.Write "What are you trying to find here? There's nothing here, understood?"
	Response.End
End If

Class ColumnTypes
	'// Public variables:
	Public Columns

	'// Private variables:
	Private temp
	Private forloop

	'// This procedure is triggered when the class is initialized
	Private Sub Class_Initialize()
		'// Build columns array, which looks like this
		'// Columns:
		'//		- Number, Name, FormType, SQL, DBName
		'// e.g - adSingle, Single, textinput, , Single	
		
		'// EVERY C0LUMN KNOWN: (NOT USED)	
		'// You can copy a column from here to the real array
		'// if you notice one is missing.
				
		'Columns = Array(_
		'Array(adEmpty, "Empty", "", "no", ""), _
		'Array(adTinyInt, "Byte", "textinput", "", "TinyInt"), _
		'Array(adSmallInt, "Small Number", "textinput", "", "SmallInt"), _
		'Array(adInteger, "Integer", "textinput", "", "Integer"), _
		'Array(adBigInt, "Big Number", "textinput", "", "Number"), _
		'Array(adUnsignedTinyInt, "Unsigned Byte", "textinput", "", "TinyInt"), _
		'Array(adUnsignedSmallInt, "Unsigned Small Number", "textinput", "", "SmallInt"), _
		'Array(adUnsignedInt, "Unsigned Number", "textinput", "", "Int"), _
		'Array(adUnsignedBigInt, "Unsigned Big Number", "textinput", "", "Number"), _
		'Array(adSingle, "Single", "textinput", "", "Single"), _
		'Array(adDouble, "Double", "textinput", "", "Double"), _
		'Array(adCurrency, "Currency", "textinput", "", "Currency"), _
		'Array(adDecimal, "Decimal", "textinput", "", "Number"), _
		'Array(adNumeric, "Numeric", "textinput", "", "Number"), _
		'Array(adBoolean, "True/False", "yesno", "", "YesNo"), _
		'Array(adError, "Error", "", "no", ""), _
		'Array(adUserDefined, "UserDefined", "", "no", ""), _
		'Array(adVariant, "Variant", "textinput", "quotes", ""), _
		'Array(adIDispatch, "IDispatch", "", "no", ""), _
		'Array(adIUnknown, "Unknown", "", "no", ""), _
		'Array(adGUID, "GUID", "textinput", "", "GUID"), _
		'Array(adDate, "Date", "textinput", "date", "Date"), _
		'Array(adDBDate, "(DB)Date", "textinput", "date", "Date"), _
		'Array(adDBTime, "(DB)Time", "textinput", "date", "Date"), _
		'Array(adDBTimeStamp, "(DB)TimeStamp", "textinput", "date", "Date"), _
		'Array(adBSTR, "String", "textinput", "quotes", "String"), _
		'Array(adChar, "Char", "textinput", "quotes", "Char"), _
		'Array(adVarChar, "(Var)Text", "textinput", "quotes", "VarChar"), _
		'Array(adLongVarChar, "Long (Var)Text", "textinput", "quotes", "VarChar"), _
		'Array(adWChar, "(W)Text", "textinput", "quotes", "Varchar"), _
		'Array(adVarWChar, "(VarW)Text", "textinput", "quotes", "Varchar"), _
		'Array(adLongVarWChar, "Memo", "textarea", "quotes", "Text"), _
		'Array(adBinary, "Binary", "", "no", "Binary"), _
		'Array(adVarBinary, "(Var)Binary", "", "no", "Binary"), _
		'Array(adLongVarBinary, "Long Binary", "", "no", "LongBinary"), _
		'Array(adChapter, "Chapter", "", "no", ""), _
		'Array(adFileTime, "FileTime", "", "no", ""), _
		'Array(adDBFileTime, "(DB)FileTime", "", "no", ""), _
		'Array(adPropVariant, "Proper Variant", "textinput", "quotes", ""), _
		'Array(adVarNumeric, "(Var)Number", "textinput", "number", "Number") _
		')
		
		'// REAL ARRAY: (IS USED)
		Columns = Array(_
		Array(adSmallInt, "Small Number", "textinput", "", "SmallInt"), _
		Array(adInteger, "Integer", "textinput", "", "Integer"), _
		Array(adUnsignedTinyInt, "Byte", "textinput", "", "TinyInt"), _
		Array(adSingle, "Single", "textinput", "", "Single"), _
		Array(adDouble, "Double", "textinput", "", "Double"), _
		Array(adCurrency, "Currency", "textinput", "", "Currency"), _
		Array(adBoolean, "True/False", "yesno", "", "YesNo"), _
		Array(adGUID, "GUID", "textinput", "", "GUID"), _
		Array(adDate, "Date", "textinput", "date", "Date"), _
		Array(adVarWChar, "Text", "textinput", "quotes", "Varchar"), _
		Array(adLongVarWChar, "Memo", "textarea", "quotes", "Text"), _
		Array(adVarBinary, "Short Binary", "", "no", "Binary"), _
		Array(adLongVarBinary, "Long Binary", "", "no", "LongBinary") _
		)
	End Sub

	'// This procedure is triggered when the class is closed/ended
	Private Sub Class_Terminate()
		'// Destroy columns array
		Columns = Empty
	End Sub
	
	'// Function to return the name of a column
	Public Function ColumnName(Byval intColumn)
		'// Clear any previous value of temp
		temp = Empty
	
		'// Loop through columns array
		For Each forloop In Columns
			If forloop(0) = Clng(intColumn) Then
				temp = forloop(1)
			End If
		Next
		
		'// Check if temp is empty
		If IsBlank(temp) Then temp = "Unknown"
		
		'// Return value
		ColumnName = temp
	End Function
	
	'// Function to return the type of a column
	Public Function ColumnType(Byval intColumn)
		'// Clear any previous value of temp
		temp = Empty
	
		'// Loop through columns array
		For Each forloop In Columns
			If forloop(0) = Clng(intColumn) Then
				temp = forloop(2)
			End If
		Next
		
		'// Check if temp is empty
		If IsBlank(temp) Then temp = ""
		
		'// Return value
		ColumnType = temp
	End Function
	
	'// Function to return what quotes must go around this field
	Public Function ColumnQuotes(Byval intColumn)
		'// Clear any previous value of temp
		temp = Empty
	
		'// Loop through columns array
		For Each forloop In Columns
			If forloop(0) = Clng(intColumn) Then
				temp = forloop(3)
			End If
		Next
		
		'// Check if temp is empty
		If IsBlank(temp) Then temp = ""
		
		'// Return value
		ColumnQuotes = temp
	End Function
	
	'// Function to return the database name of a column
	Public Function ColumnDatabase(Byval intColumn)
		'// Clear any previous value of temp
		temp = Empty
	
		'// Loop through columns array
		For Each forloop In Columns
			If forloop(0) = Clng(intColumn) Then
				temp = forloop(4)
			End If
		Next
		
		'// Check if temp is empty
		If IsBlank(temp) Then temp = "Varchar"
		
		'// Return value
		ColumnDatabase = temp
	End Function
	
	'// Function to create HTML of all columns
	Public Function CreateHTML(Byval HTML)
		'// Clear any previous value of temp
		temp = Empty
		
		For Each forloop In Columns
			temp = temp & HTML
			temp = Replace(temp, "$column->id", forloop(0))
			temp = Replace(temp, "$column->name", forloop(1))
			temp = Replace(temp, "$column->type", forloop(2))
			temp = Replace(temp, "$column->db", forloop(3))
		Next
				
		'// Return value
		CreateHTML = temp
	End Function	
End Class
%>