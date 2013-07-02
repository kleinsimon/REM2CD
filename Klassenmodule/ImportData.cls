VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Path As String
Public filename As String
Public calibration As String

Public Sub SetImg(Path As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim cstring As String

    Me.Path = Path
    Me.filename = fso.GetFileName(Path)
    
    Set ts = fso.GetFile(Path).OpenAsTextStream(ForReading, TristateFalse)
    
    Do Until ts.AtEndOfStream
        If (InStr(1, ts.ReadLine, "AP_IMAGE_PIXEL_SIZE", vbBinaryCompare) > 0) Then
            cstring = ts.ReadLine
            Exit Do
        End If
    Loop
    
    cstring = Split(cstring, "=", 2, vbBinaryCompare)(1)
    Me.calibration = CStr(convertToMicron(cstring))
End Sub

Private Function convertToMicron(s As String) As Double
    Dim Data() As String
    Dim val As Double
    val = 1
    s = Trim(s)
    s = Replace(s, ".", getSysSeperator)
    Data = Split(s, " ", 2)
    val = CDbl(Data(0))
    
    Data(1) = Trim(Data(1))

    If (InStr(Data(1), "nm")) Then val = val / 1000#
    If (InStr(Data(1), "mm")) Then val = val * 1000#
    If (InStr(Data(1), "pm")) Then val = val / 1000000#
    
    convertToMicron = val
End Function


Private Function getSysSeperator() As String
    getSysSeperator = Mid$(1 / 2, 2, 1)
End Function
