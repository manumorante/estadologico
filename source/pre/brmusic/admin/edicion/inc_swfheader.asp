<%
'-------------------------------------------------------------
'  Create Date : 17/10/2001 (dd/mm/yyyy)
'  Mod. Date   : 17/10/2001
'  Author      : Claudio Heidel (heidel@f256.com)
'-------------------------------------------------------------

Class SWFDump

  Private header
  Private RECTdata
  Private nBits
  Private mversion
  Private mfilelen
  Private mxMin
  Private mxMax
  Private myMin
  Private myMax
  Private mheigt
  Private mwidth
  Private mframerate
  Private mframecount

  Private Sub Class_Initialize()

  End Sub

  Private Sub Class_Terminate()

  End Sub


  Private Function ReadHeader (filename)
     Const ForReading = 1, ForWriting = 2, ForAppending = 8
     Dim fso, f
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set f = fso.OpenTextFile(filename, ForReading)
     ReadHeader = f.Read(21)
  End Function

  Private Function ToBin(inNumber, OutLenStr )
    Dim binary
    binary = ""
    do while inNumber >= 1
      binary = binary & inNumber mod 2
      inNumber = inNumber \ 2
    loop
    binary = binary & String(OutLenStr - len(binary), "0")
    ToBin = StrReverse(binary)
  End Function

  Private Function Bin2Decimal(inBin)
    Dim counter
    Dim temp
    Dim Value
    inBin = StrReverse(inBin)
    temp = 0
    For counter = 1 to Len(inBin)
      If counter = 1 then
        Value = 1
      Else
        Value = Value  * 2
      End If
      temp = temp + mid(inBin, counter ,1)  *  Value
    Next
    Bin2Decimal = temp
  End Function

  Public Function SWFDump(fileName)

    header = ReadHeader (fileName)
    mversion = asc(mid(header,4,1))
    mfilelen = asc(mid(header,5,1))
    mfilelen = mfilelen + asc(mid(header,6,1)) * 256
    mfilelen = mfilelen + asc(mid(header,7,1)) * 256 * 256
    mfilelen = mfilelen + asc(mid(header,8,1)) * 256 * 256 * 256

    RECTdata = ToBin(asc(mid(header,9,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,10,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,11,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,12,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,13,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,14,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,15,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,16,1)),8)
    RECTdata = RECTdata & ToBin(asc(mid(header,17,1)),8)

    nBits = Mid(RECTdata,1,5)
    nBits = Bin2Decimal(nBits)

    mxMin =  Bin2Decimal(Mid(RECTdata,6,nBits))
    mxMax =  Bin2Decimal(Mid(RECTdata,6 + nBits * 1 ,nBits))
    myMin =  Bin2Decimal(Mid(RECTdata,6 + nBits * 2 ,nBits))
    myMax =  Bin2Decimal(Mid(RECTdata,6 + nBits * 3 ,nBits))

    mheigt = (myMax - myMin) / 20
    mwidth = (mxMax - mxMin) / 20

    mframerate = asc(mid(header,18,1))

    mframecount = asc(mid(header,19,1))
    mframecount = mframecount + asc(mid(header,20,1)) * 256

  End Function


  Public Property Get Heigt()
    Heigt = mheigt
  End Property

  Public Property Get Width()
    Width = mwidth
  End Property

  Public Property Get Version()
    Version = mversion
  End Property

  Public Property Get FileLen()
    FileLen = mfilelen
  End Property

  Public Property Get xMin()
    xMin = mxMin
  End Property

  Public Property Get xMax()
    xMax = mxMax
  End Property

  Public Property Get yMin()
    yMin = myMin
  End Property

  Public Property Get yMax()
    yMax = myMax
  End Property

  Public Property Get Framerate()
    Framerate = mframerate
  End Property

  Public Property Get Framecount()
    Framecount = mframecount
  End Property
End Class
%>