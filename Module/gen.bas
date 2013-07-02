Attribute VB_Name = "gen"
'Tool for adding micron-bars to microscopic images in Corel Draw, written by Simon Klein (mail@simonklein.de)
'This Tool may be used or altered for private, non-commercial or educational use.
Sub showCalibration()
    On Error Resume Next
    rem2cdCalib.Show (vbModeless)
End Sub

'Shows the Parameter input form for transformation and calibration
Sub makeContainer()
    On Error Resume Next

    
    'If (ActiveSelection.Shapes(1).Type = cdrBitmapShape) Then
    rem2cdValues.Show (vbModeless)
    'rem2cdValues.setImg ActiveSelection.Shapes

End Sub


'transforms a image to a given size and adds a micron-bar and a border, dependend on calibration info
Function doContTransform( _
    ByVal CImg As Shape, _
     Mode As Integer, _
     calibration As String, _
     Text As String, _
     Length As String, _
     widthstr As String, _
     heightstr As String, _
     TextOL As String, _
     TextOR As String, _
     TextUL As String, _
     Lth As String, _
     LBth As String, _
     TextBold As Boolean, _
     TextSize As String _
    ) As Shape
    
    On Error GoTo ErrHandler
    Dim Brect As Shape
    Dim Lrect, Wline, Ttext, Gr As Shape
    Dim SR As ShapeRange
    Dim width, height, Left, Bottom As Double
    Dim LminW, LmaxW, LmarginH, LmarginV, TmarginH, TmarginV As Double
    Dim Lwidth, WWidth, WHeight, LHeight As Double
    Dim white, black As Color
    Dim textShapes As New ShapeRange
    Dim noBar As Boolean
    noBar = False
    Dim oldname As String
    
    If (CImg.Properties.Exists("rem2cd", PropID.isBalkenGroup)) Then
        Dim sh, tmpSh As Shape
        oldname = CImg.Name
        Set tmpSh = CImg
        For Each sh In tmpSh.PowerClip.Shapes
            If (sh.Name = "balkenImage") Then
                Set CImg = sh
            Else
                sh.Delete
            End If
        Next
        tmpSh.PowerClip.ExtractShapes
        tmpSh.Delete
    End If
    
    Set white = New Color
    Set black = New Color
    black.CMYKAssign 0, 0, 0, 100
    white.CMYKAssign 0, 0, 0, 0

    Set SR = New ShapeRange
    If (GetSetting("rem2cd", "settings", "LastAbort") = "true") Then Exit Function
    
    ActiveDocument.BeginCommandGroup "rem2cd"
    Application.Optimization = True

    'Text = GetSetting("rem2cd", "settings", "LastText", "N/A")
    Lwidth = ConvertUnits(Length, cdrCentimeter, ActiveDocument.Unit)
    
    'widthstr = GetSetting("rem2cd", "settings", "LastWidth", "11")
    'heightstr = GetSetting("rem2cd", "settings", "LastHeight", "")
    
    'TextOL = GetSetting("rem2cd", "settings", "LastTextOL", "")
    'TextOR = GetSetting("rem2cd", "settings", "LastTextOR", "")
    'TextUL = GetSetting("rem2cd", "settings", "LastTextUL", "")
    'Lth = CDbl(GetSetting("rem2cd", "settings", "LastLine", "1,5"))
    'LBth = CDbl(GetSetting("rem2cd", "settings", "LastLineB", "1,5"))
    'TextBold = CBool(GetSetting("rem2cd", "settings", "LastBold", "false"))
        
    Left = CImg.LeftX
    Bottom = CImg.BottomY
    
    With CImg
        cimgratio = .SizeHeight / .SizeWidth
        If (widthstr = "" And heightstr = "") Then
            width = .SizeWidth
            height = .SizeHeight
        ElseIf (widthstr = "") Then
            height = ConvertUnits(heightstr, cdrCentimeter, ActiveDocument.Unit)
            width = height / cimgratio
        ElseIf (heightstr = "") Then
            width = ConvertUnits(widthstr, cdrCentimeter, ActiveDocument.Unit)
            height = width * cimgratio
        Else
            width = ConvertUnits(widthstr, cdrCentimeter, ActiveDocument.Unit)
            height = ConvertUnits(heightstr, cdrCentimeter, ActiveDocument.Unit)
        End If
        
        brdratio = height / width
        If (brdratio <= cimgratio) Then
            .SizeWidth = width
            .SizeHeight = width * cimgratio
        Else
            .SizeHeight = height
            .SizeWidth = height / cimgratio
        End If
        
        .LeftX = Left
        .BottomY = Bottom
    End With
    
    LminW = ConvertUnits(GetSetting("rem2cd", "settings", "LastBalkenMin", "2"), cdrCentimeter, ActiveDocument.Unit)
    LmaxW = ConvertUnits(GetSetting("rem2cd", "settings", "LastBalkenMax", width / 3), cdrCentimeter, ActiveDocument.Unit)

    If (Mode = 0) Then
        res = getScale(ConvertUnits(LminW, ActiveDocument.Unit, cdrCentimeter), ConvertUnits(LmaxW, ActiveDocument.Unit, cdrCentimeter), calibration, CImg)
        Lwidth = ConvertUnits(res(1), cdrCentimeter, ActiveDocument.Unit)
        Text = getScaleText(res(2))
    ElseIf (Text = "" Or Lwidth = 0) Then
        noBar = True
    End If
    
    LHeight = ConvertUnits(8, cdrPoint, ActiveDocument.Unit)
    LmarginH = ConvertUnits(8, cdrPoint, ActiveDocument.Unit) 'Abstand der Maßlinie, Horizontal
    LmarginV = ConvertUnits(4, cdrPoint, ActiveDocument.Unit) 'Abstand der Maßlinie, Vertikal
    TmarginH = ConvertUnits(4, cdrPoint, ActiveDocument.Unit) 'Abstand der Maßzahl, Horizontal
    TmarginV = ConvertUnits(4, cdrPoint, ActiveDocument.Unit) 'Abstand der Maßzahl, Vertikal
    
    'TextSize = CDbl(GetSetting("rem2cd", "settings", "LastTextSize", "10"))
    TextHeight = ConvertUnits(TextSize, cdrPoint, ActiveDocument.Unit)
    
    WWidth = Lwidth + 2 * LmarginH
    WHeight = LHeight + LmarginV + TmarginV
    THeight = TextHeight
    
    If (Not noBar) Then
        Set Lrect = ActiveLayer.CreateRectangle2(Left + width - WWidth, Bottom, WWidth, WHeight + TextHeight)
        Set Wline = ActiveLayer.CreateLineSegment(Left + width - LmarginH - Lwidth, Bottom + WHeight / 2, Left + width - LmarginH, Bottom + WHeight / 2)
        Set Ttext = ActiveLayer.CreateArtisticText(Left + width - WWidth / 2, Bottom + WHeight / 2 + TmarginV, Text, Alignment:=cdrCenterAlignment, Size:=TextSize)
    End If
    Set Brect = ActiveLayer.CreateRectangle2(Left, Bottom, width, height)
    
       
    If (TextOL <> "") Then
        textShapes.Add ActiveLayer.CreateArtisticText(Left + TmarginH, Bottom + height - THeight - TmarginV, TextOL, Alignment:=cdrLeftAlignment, Size:=TextSize)
        textShapes.Add ActiveLayer.CreateRectangle2(Left, Bottom + height - THeight - 2 * TmarginV, textShapes(textShapes.Count).SizeWidth + 2 * TmarginH, THeight + 2 * TmarginV)
        textShapes(textShapes.Count).Fill.ApplyUniformFill (white)
        textShapes(textShapes.Count).Outline.width = 0
        textShapes(textShapes.Count).OrderBackOne
        textShapes(textShapes.Count - 1).Text.Story.Bold = TextBold
    End If
    
    If (TextOR <> "") Then
        textShapes.Add ActiveLayer.CreateArtisticText(Left + width - TmarginH, Bottom + height - THeight - TmarginV, TextOR, Alignment:=cdrRightAlignment, Size:=TextSize)
        textShapes.Add ActiveLayer.CreateRectangle2(Left + width - textShapes(textShapes.Count).SizeWidth - 2 * TmarginH, Bottom + height - THeight - 2 * TmarginV, textShapes(textShapes.Count).SizeWidth + 2 * TmarginH, THeight + 2 * TmarginV)
        textShapes(textShapes.Count).Fill.ApplyUniformFill (white)
        textShapes(textShapes.Count).Outline.width = 0
        textShapes(textShapes.Count).OrderBackOne
        textShapes(textShapes.Count - 1).Text.Story.Bold = TextBold
    End If
    
    If (TextUL <> "") Then
        textShapes.Add ActiveLayer.CreateArtisticText(Left + TmarginH, Bottom + TmarginV, TextUL, Alignment:=cdrLeftAlignment, Size:=TextSize)
        textShapes.Add ActiveLayer.CreateRectangle2(Left, Bottom, textShapes(textShapes.Count).SizeWidth + 2 * TmarginH, THeight + 2 * TmarginV)
        textShapes(textShapes.Count).Fill.ApplyUniformFill (white)
        textShapes(textShapes.Count).Outline.width = 0
        textShapes(textShapes.Count).OrderBackOne
        textShapes(textShapes.Count - 1).Text.Story.Bold = TextBold
    End If
    
    If (Not noBar) Then
        With Ttext
            Ttext.Text.Story.Bold = TextBold
        End With
        
        With Lrect
            .Outline.width = 0
            .Fill.ApplyUniformFill (white)
        End With
        
        With Wline
            .Outline.width = ConvertUnits(Lth, cdrPoint, ActiveDocument.Unit)
            .Outline.Color = black
            .Outline.EndArrow = ArrowHeads(59)
            .Outline.StartArrow = ArrowHeads(59)
        End With
    End If
    
    With Brect
        .Outline.Color = black
        .Outline.width = ConvertUnits(LBth, cdrPoint, ActiveDocument.Unit)
    End With
    
    If (Not noBar) Then
        SR.Add (Lrect)
        SR.Add (Wline)
        SR.Add (Ttext)
        Set Gr = SR.Group
    End If
    
    CImg.Name = "balkenImage"
    
    CImg.AddToPowerClip Brect, cdrFalse
    If (Not noBar) Then Gr.AddToPowerClip Brect, cdrFalse
    textShapes.AddToPowerClip Brect, cdrFalse
    
    Brect.Properties("rem2cd", PropID.isBalkenGroup) = True
    Brect.Properties("rem2cd", PropID.calibration) = calibration
    Brect.Properties("rem2cd", PropID.width) = widthstr
    Brect.Properties("rem2cd", PropID.height) = heightstr
    Brect.Properties("rem2cd", PropID.Length) = Length
    Brect.Properties("rem2cd", PropID.lineW) = Lth
    Brect.Properties("rem2cd", PropID.lineBW) = LBth
    Brect.Properties("rem2cd", PropID.Text) = Text
    Brect.Properties("rem2cd", PropID.Mode) = Mode
    Brect.Properties("rem2cd", PropID.TextBold) = TextBold
    Brect.Properties("rem2cd", PropID.TextSize) = TextSize
    Brect.Properties("rem2cd", PropID.TextOL) = TextOL
    Brect.Properties("rem2cd", PropID.TextOR) = TextOR
    Brect.Properties("rem2cd", PropID.TextUL) = TextUL
    Brect.Properties("rem2cd", PropID.filename) = oldname
    Brect.Name = oldname

    'Dim df As DataField
    'Dim di As DataItem
    
    'Set df = ActiveDocument.DataFields.Add("rem2cd_Settings")
    'Set di = Brect.ObjectData.Add(df)
    'di.Value = ""
    
    ActiveWindow.Refresh
    Application.Refresh
    ActiveDocument.ClearSelection
    
    Brect.Selected = True
    Set doContTransform = Brect
    
ExitFunction:
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Exit Function
    
ErrHandler:
    MsgBox "Error: " & Err.Description
    Resume ExitFunction
End Function

'Returns near full number in a range of line widths. (0) is the corrosponding length of a line calculated with calibration parameter, (1) is the actual value
Function getScale(ByVal minW As Double, ByVal maxW As Double, ByVal calibration As Double, ByRef img As Shape) As Double()
    Dim minVal As Double
    Dim resMm As Double
    Dim pot As Double
    Dim res(2) As Double
    Dim i As Integer
    
    resMm = ConvertUnits(img.Bitmap.SizeWidth / img.SizeWidth, cdrCentimeter, ActiveDocument.Unit)
    minVal = minW * resMm * calibration
    
    pot = (10 ^ Floor(Log10(minVal)))
    
    For i = 1 To 19
        Dim l As Double
        Dim z As Double
        
        z = i * pot
        l = CDbl(z) / (resMm * calibration)

        If (l >= minW And l <= maxW) Then
            res(1) = l
            res(2) = z
            getScale = res
            Exit For
        Else
            z = 10 * i * pot
            l = CDbl(z) / (resMm * calibration)
            If (l >= minW And l <= maxW) Then
                res(1) = l
                res(2) = z
                getScale = res
                Exit For
            End If
        End If
    Next i
End Function

' Returns unit for a given scale between pm and m
Function getScaleText(ByVal sc As Double) As String
    Dim pot As Double
    Dim res As String
    
    Select Case (sc)
        Case Is >= 1000000000
            res = CStr(sc / 1000000000) + " km"
            
        Case Is >= 1000000
            res = CStr(sc / 1000000) + " m"
            
        Case Is >= 10000
            res = CStr(sc / 10000) + " cm"
            
        Case Is >= 1000
            res = CStr(sc / 1000) + " mm"
        
        Case Is < 0.001
            res = CStr(sc * 1000000) + " pm"
            
        Case Is < 1
            res = CStr(sc * 1000) + " nm"
        
        Case Else
            res = CStr(sc) + " µm"
    End Select
    
    getScaleText = res
End Function

'Returns the next lower, full number of a value
Function Floor(ByVal num As Double) As Double
    r = Round(num, 0)
    If (r > num) Then r = r - 1
    
    Floor = r
End Function

'Returns the Logarithm of base 10
Static Function Log10(x)
    Log10 = Log(x) / Log(10#)
End Function
