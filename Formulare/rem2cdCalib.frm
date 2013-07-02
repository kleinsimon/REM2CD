VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rem2cdCalib 
   Caption         =   "Kalibration messen"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5085
   OleObjectBlob   =   "rem2cdCalib.frx":0000
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "rem2cdCalib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calibration Parameter measuring tool for Corel Draw, written by Simon Klein (mail@simonklein.de)
'This Tool may be used or altered for private, non-commercial or educational use.
Dim mes As Shape
Dim img As Shape
Dim calib As Double
Public parentWindow As Variant
Private val As New Validater

Private Sub CommandButton1_Click()
    On Error Resume Next
    If (Not val.validate(Me.TextBox1, valPosNumber)) Then Exit Sub

    If (mes Is Nothing Or img Is Nothing) Then
        Unload Me
        Exit Sub
    End If
    
    Dim w, res, pxDist As Double
    
    res = CDbl(img.Bitmap.SizeWidth) / img.SizeWidth
    pxDist = mes.SizeWidth * res
    w = CDbl(Me.TextBox1.Value)
    calib = w / pxDist
    
    'MsgBox calib

    TextOut.Text = Round(calib, 10)
    If (Not parentWindow = Null) Then CommandButton2.Visible = True
End Sub

Private Sub CommandButton2_Click()
    If (Not val.validate(Me.TextOut, valPosNumber)) Then Exit Sub
    parentWindow.setcalib (calib)
    Unload Me
End Sub


Private Sub CommandButton3_Click()
    abort
End Sub

Private Sub UserForm_Initialize()
    Me.Left = CDbl(GetSetting("rem2cd", "settings", "CalibLeft", "500"))
    Me.Top = CDbl(GetSetting("rem2cd", "settings", "CalibTop", "500"))

    If (ActiveSelection.Shapes.Count = 0) Then
        MsgBox "Kein Objekt ausgewählt"
        Unload Me
        Exit Sub
    Else
        If (ActiveSelection.Shapes(1).Type = cdrBitmapShape) Then
            Set img = ActiveSelection.Shapes(1)
        ElseIf (ActiveSelection.Shapes(1).Name = "balkenGroup") Then
            Dim sh, tmpSh As Shape
            Set tmpSh = ActiveSelection.Shapes(1)
            For Each sh In tmpSh.PowerClip.Shapes
                If (sh.Name = "balkenImage") Then
                    Set img = sh
                End If
            Next
        End If
    End If

    If (img Is Nothing) Then
        MsgBox "Kein Bild ausgewählt"
        Unload Me
        Exit Sub
    End If
    
    
        
    Set mes = ActiveLayer.CreateRectangle2( _
        ActiveWindow.ActiveView.OriginX, _
        ActiveWindow.ActiveView.OriginY, _
        ConvertUnits(3, cdrCentimeter, ActiveDocument.Unit), _
        ConvertUnits(1, cdrCentimeter, ActiveDocument.Unit) _
        )
    
    Dim back As Color
    Set back = New Color
    back.CMYKAssign 0, 60, 100, 0
    
    With mes
        .Selected = True
        .Outline.width = 0
        .Fill.ApplyUniformFill back
        .Transparency.ApplyUniformTransparency 0
        .Transparency.MergeMode = cdrMergeXOR
    End With
    
    LabelInfo.Caption = "Bewegen Sie das Rechteck auf die Referenzstrecke und passen Sie die Größe an"
End Sub

Sub abort()
    On Error Resume Next
    Me.Hide
    parentWindow.setcalib -1
    SaveSetting "rem2cd", "settings", "CalibLeft", CStr(Me.Left)
    SaveSetting "rem2cd", "settings", "CalibTop", CStr(Me.Top)
    'If (Not img Is Nothing) Then img.Selected = True
    If (Not mes Is Nothing) Then mes.Delete
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    abort
End Sub


