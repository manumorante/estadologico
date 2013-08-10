<%
' AscAt
' HexAt
' isJPG
' isPNG
' isGIF
' isBMP
' isWMF
' isWebImg

' ReadImg
' ReadJPG
' ReadPNG
' ReadGIF
' ReadWMF
' ReadBMP

' isDigit
' isHex
' HexToDec
Dim HW

Function AscAt(s, n)
	AscAt = Asc(Mid(s, n, 1))
End Function

Function HexAt(s, n)
	HexAt = Hex(AscAt(s, n))
End Function


Function isJPG(fichero)
	dim archivo
	set fso = CreateObject("Scripting.FileSystemObject")
	set archivo = fso.getFile(fichero)
	If inStr(uCase(archivo.type), "JPG") <> 0 Then
		isJPG = true
	Else
		isJPG = false
	End If
	set fso = nothing
	set archivo = nothing
End Function


Function isPNG(fichero)
	dim archivo
	set fso = CreateObject("Scripting.FileSystemObject")
	set archivo = fso.getFile(fichero)
	If inStr(uCase(archivo.type), "PNG") <> 0 Then
		isPNG = true
	Else
		isPNG = false
	End If
	set fso = nothing
	set archivo = nothing
End Function


Function isGIF(fichero)
	dim archivo
	set fso = CreateObject("Scripting.FileSystemObject")
	set archivo = fso.getFile(fichero)
	If inStr(uCase(archivo.type), "GIF") <> 0 Then
		isGIF = true
	Else
		isGIF = false
	End If
	set fso = nothing
	set archivo = nothing
End Function


Function isBMP(fichero)
	dim archivo
	set fso = CreateObject("Scripting.FileSystemObject")
	set archivo = fso.getFile(fichero)
	If inStr(uCase(archivo.type), "BMP") <> 0 Then
		isBMP = true
	Else
		isBMP = false
	End If
	set fso = nothing
	set archivo = nothing
End Function


Function isWMF(fichero)
	dim archivo
	set fso = CreateObject("Scripting.FileSystemObject")
	set archivo = fso.getFile(fichero)
	If inStr(uCase(archivo.type), "WMF") <> 0 Then
		isWMF = true
	Else
		isWMF = false
	End If
	set fso = nothing
	set archivo = nothing
End Function


Function isWebImg(f)
'	If isGIF(f) Or isJPG(f) Or isPNG(f) Or isBMP(f) Or isWMF(f) Then
	If isGIF(f) Or isJPG(f) Or isPNG(f) Then
		isWebImg = true
	Else
		isWebImg = true
	End If
End Function


Function ReadImg(fichero)
	If isGIF(fichero) Then
		ReadImg = ReadGIF(fichero)
	ElseIf isJPG(fichero) Then
		ReadImg = ReadJPG(fichero)
	ElseIf isPNG(fichero) Then
		ReadImg = ReadPNG(fichero)
	ElseIf isBMP(fichero) Then
		ReadImg = ReadPNG(fichero)
	ElseIf isWMF(fichero) Then
		ReadImg = ReadWMF(fichero)
	Else
		ReadImg = Array(0,0)
	End If
End Function


Function ReadJPG(fichero)
	Dim fso, ts, s, HW, nbytes
	HW = Array("","")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fichero, 1)
	s = Right(ts.Read(167), 4)
	HW(0) = HexToDec(HexAt(s,3) & HexAt(s,4))
	HW(1) = HexToDec(HexAt(s,1) & HexAt(s,2))
	ts.Close
	ReadJPG = HW
End Function


Function ReadPNG(fichero)
	Dim fso, ts, s, HW, nbytes
	HW = Array("","")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fichero, 1)
	s = Right(ts.Read(24), 8)
	HW(0) = HexToDec(HexAt(s,3) & HexAt(s,4))
	HW(1) = HexToDec(HexAt(s,7) & HexAt(s,8))
	ts.Close
	ReadPNG = HW
End Function


Function ReadGIF(fichero)
	Dim fso, ts, s, HW, nbytes
	HW = Array("","")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fichero, 1)
	s = Right(ts.Read(10), 4)
	HW(0) = HexToDec(HexAt(s,2) & HexAt(s,1))
	HW(1) = HexToDec(HexAt(s,4) & HexAt(s,3))
	ts.Close
	ReadGIF = HW
End Function


Function ReadWMF(fichero)
	Dim fso, ts, s, HW, nbytes
	HW = Array("","")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fichero, 1)
	s = Right(ts.Read(14), 4)
	HW(0) = HexToDec(HexAt(s,2) & HexAt(s,1))
	HW(1) = HexToDec(HexAt(s,4) & HexAt(s,3))
	ts.Close
	ReadWMF = HW
End Function


Function ReadBMP(fichero)
	Dim fso, ts, s, HW, nbytes
	HW = Array("","")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fichero, 1)
	s = Right(ts.Read(24), 8)
	HW(0) = HexToDec(HexAt(s,4) & HexAt(s,3))
	HW(1) = HexToDec(HexAt(s,8) & HexAt(s,7))
	ts.Close
	ReadBMP = HW
End Function


Function isDigit(c)
	If inStr("0123456789", c) <> 0 Then
		isDigit = true
	Else
		isDigit = false
	End If
End Function


Function isHex(c)
	If inStr("0123456789ABCDEFabcdef", c) <> 0 Then
		isHex = true
	Else
		ishex = false
	End If
End Function


Function HexToDec(cadhex)
	Dim n, i, ch, decimal
	decimal = 0
	n = Len(cadhex)
	For i=1 To n
		ch = Mid(cadhex, i, 1)
		If isHex(ch) Then
			decimal = decimal * 16
			If isDigit(ch) Then
				decimal = decimal + ch
			Else
				decimal = decimal + Asc(uCase(ch)) - Asc("A")
			End If
		Else
			HexToDec = -1
		End If
	Next
	HexToDec = decimal
End Function
%>