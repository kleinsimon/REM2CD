VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rem2cdValues 
   Caption         =   "Maﬂbalken hinzuf¸gen"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10650
   OleObjectBlob   =   "rem2cdValues.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "rem2cdValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tool for adding micron-bars to microscopic images in Corel Draw, written by Simon Klein (mail@simonklein.de)
'This Tool may be used or altered for private, non-commercial or educational use.
Option Explicit
Private img As New ShapeRange
Private val As New Validater
Private impImg As New Collection

Private Sub CommandButtonClearList_Click()
    Me.ListView1.ListItems.Clear
    Set impImg = Nothing
    Set impImg = New Collection
End Sub

Private Sub CommandButtonDelItemFromList_Click()
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    removeImg Me.ListView1.SelectedItem.Text
    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim f As Variant
    For Each f In Data.Files
        addFile CStr(f)
    Next
End Sub

Private Sub ListView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If (Data.Files.Count >= 1) Then
        Effect = 1
        State = 1
    End If
End Sub

Private Sub removeImg(filename As String)
    Dim i As Integer
    For i = 1 To impImg.Count
        If impImg(i).filename = filename Then impImg.Remove i: Exit For
    Next
End Sub


Private Sub UserForm_Initialize()
    With Me.ListView1
        .View = lvwReport
        .ColumnHeaders.Add , , "Datei", 2 * .width / 3 - 2
        .ColumnHeaders.Add , , "Kalibrationswert", .width / 3 - 2
    End With
    SaveSetting "rem2cd", "settings", "LastAbort", "true"
    SetImg ActiveSelection.Shapes
    LoadValues
End Sub

Public Function SetImg(im As Shapes)
    Dim i As Shape
    For Each i In im
        If i.Properties.Exists("rem2cd", PropID.isBalkenGroup) And im.Count > 1 Then
            MsgBox "Es kann nur 1 gerahmtes Bild bearbeitet werden"
            Unload Me
            Exit Function
        End If
        img.Add i
    Next
End Function

Private Function validate() As Boolean
    Dim res As Boolean
    res = True
    res = res And val.validate(Me.TextBoxHeight, valPosNumOrEmpty)
    res = res And val.validate(Me.TextBoxWidth, valPosNumOrEmpty)
    res = res And val.validate(Me.TextBoxLen, valPosNumber)
    res = res And val.validate(Me.TextBoxLine, valNumber)
    res = res And val.validate(Me.TextBoxLineB, valNumber)
    validate = res
End Function


Private Sub CommandButton1_Click()
    If img.Count = 0 Then SetImg ActiveSelection.Shapes

    SaveSettings
    If (Not validate) Then Exit Sub
    
    If (MultiPageMode.Value = 2) Then
        If impImg.Count = 0 Then Exit Sub

        Dim f As Variant
        Dim fi As Integer
        Dim CurImg, res As Shape
        For fi = 1 To impImg.Count
            
            ActiveDocument.ClearSelection
            ActiveDocument.ActiveLayer.Import impImg(fi).Path
            Set CurImg = ActiveSelection.Shapes(1)
            Set res = doContTransform( _
                CurImg, 0, impImg(fi).calibration, _
                TextBoxTxt.Text, TextBoxLen.Text, TextBoxWidth.Text, TextBoxHeight.Text, _
                TextBoxOL.Text, TextBoxOR.Text, TextBoxUL.Text, TextBoxLine.Text, TextBoxLineB.Text, CheckBoxBold.Value, TextBoxSize.Text)
                
            res.Properties("rem2cd", PropID.filename) = impImg(fi).filename
            res.Name = impImg(fi).filename
            'ActiveLayer.CreateArtisticText img.LeftX, img.BottomY, impImg(fi).filename
        Next
    Else
       If (ActiveSelection.Shapes.Count < 1) Then
               MsgBox "Kein Objekt ausgew‰hlt"
               Exit Sub
       End If
       
       'ActiveDocument.ClearSelection
    
       Dim im As Shape
       If (img Is Nothing Or img.Count = 0) Then Exit Sub
       
       For Each im In img
           doContTransform _
               im, MultiPageMode.Value, TextBoxCalib.Text, _
               TextBoxTxt.Text, TextBoxLen.Text, TextBoxWidth.Text, TextBoxHeight.Text, _
               TextBoxOL.Text, TextBoxOR.Text, TextBoxUL.Text, TextBoxLine.Text, TextBoxLineB.Text, CheckBoxBold.Value, TextBoxSize.Text
       Next
    End If
    
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo ErrHandler
    
    Dim Path() As String
    Dim p As Variant
    
    'Path = CorelScriptTools.GetFileBox("Tiff|*.tif;*.tiff", "REM-Bild ˆffnen", 0)
    With Me.CommonDialog1
        .Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
        .Filter = "Tiff|*.tif;*.tiff"
        .CancelError = True
        .ShowOpen
        If .filename = "" Then Exit Sub
        Path = Split(.filename & vbNullChar, vbNullChar)
    End With
    
    For Each p In Path
        addFile CStr(p)
    Next
    
ErrHandler:

    Err.Clear
    Exit Sub

End Sub

Private Sub addFile(Path As String)
    If Not (EndsWith(Path, ".tif") Or EndsWith(Path, ".tiff")) Then Exit Sub
    Dim iD As New ImportData
    
    iD.SetImg (Path)
    
    
    With Me.ListView1.ListItems.Add(, , iD.filename)
        .SubItems(1) = iD.calibration
    End With

    impImg.Add iD
End Sub

Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Private Function getSysSeperator() As String
    getSysSeperator = Mid$(1 / 2, 2, 1)
End Function

Private Sub CommandButtonList_Click()
    img(1).Selected = True
    rem2cdList.Show (vbModeless)
    Set rem2cdList.parentWindow = Me
    'Me.Enabled = False
    Me.Hide
End Sub

Private Sub CommandButtonMes_Click()
    img(1).Selected = True
    rem2cdCalib.Show (vbModeless)
    Set rem2cdCalib.parentWindow = Me
    'Me.Enabled = False
    Me.Hide
End Sub

Public Sub setcalib(calibration As Double)
    If (calibration > 0) Then Me.TextBoxCalib.Text = CStr(calibration)
    'Me.Enabled = True
    Me.Show vbModeless
    'ActiveDocument.ClearSelection
End Sub

Private Sub LoadValues()
    If (ActiveSelection.Shapes.Count = 1) Then
        If ActiveSelection.Shapes(1).Properties.Exists("rem2cd", PropID.isBalkenGroup) Then
        'EDIT MODE
            Me.Caption = "Maﬂbalken bearbeiten"
            Me.MultiPageMode.Value = 0
            Me.MultiPageMode.Pages(2).Enabled = False
            With ActiveSelection.Shapes(1)
                Me.TextBoxCalib.Text = .Properties("rem2cd", PropID.calibration)
                Me.TextBoxLen.Text = .Properties("rem2cd", PropID.Length)
                Me.TextBoxTxt.Text = .Properties("rem2cd", PropID.Text)
                Me.TextBoxWidth.Text = .Properties("rem2cd", PropID.width)
                Me.TextBoxHeight.Text = .Properties("rem2cd", PropID.height)
                Me.TextBoxLine.Text = .Properties("rem2cd", PropID.lineW)
                Me.TextBoxLineB.Text = .Properties("rem2cd", PropID.lineBW)
                Me.TextBoxOL.Text = .Properties("rem2cd", PropID.TextOL)
                Me.TextBoxOR.Text = .Properties("rem2cd", PropID.TextOR)
                Me.TextBoxUL.Text = .Properties("rem2cd", PropID.TextUL)
                Me.TextBoxSize = .Properties("rem2cd", PropID.TextSize)
                
                Me.Left = CDbl(GetSetting("rem2cd", "settings", "LastL", "500"))
                Me.Top = CDbl(GetSetting("rem2cd", "settings", "LastT", "500"))
                
                Me.CheckBoxBold.Value = CBool(.Properties("rem2cd", PropID.TextBold))
        
                Me.MultiPageMode.Value = .Properties("rem2cd", PropID.Mode)
                
                Exit Sub
            End With
        End If
    End If
    
    Me.TextBoxLen.Text = GetSetting("rem2cd", "settings", "LastLength", "3")
    Me.TextBoxCalib.Text = GetSetting("rem2cd", "settings", "LastCalib", "1")
    Me.TextBoxLen.Text = GetSetting("rem2cd", "settings", "LastLength", "3")
    Me.TextBoxTxt.Text = GetSetting("rem2cd", "settings", "LastText", "N/A")
    Me.TextBoxWidth.Text = GetSetting("rem2cd", "settings", "LastWidth", "11")
    Me.TextBoxHeight.Text = GetSetting("rem2cd", "settings", "LastHeight", "8")
    Me.TextBoxLine.Text = GetSetting("rem2cd", "settings", "LastLine", "1,5")
    Me.TextBoxLineB.Text = GetSetting("rem2cd", "settings", "LastLineB", "1,5")
    Me.TextBoxOL.Text = GetSetting("rem2cd", "settings", "LastTextOL", "")
    Me.TextBoxOR.Text = GetSetting("rem2cd", "settings", "LastTextOR", "")
    Me.TextBoxUL.Text = GetSetting("rem2cd", "settings", "LastTextUL", "")
    Me.TextBoxSize = GetSetting("rem2cd", "settings", "LastTextSize", "10")
    
    Me.Left = CDbl(GetSetting("rem2cd", "settings", "LastL", "500"))
    Me.Top = CDbl(GetSetting("rem2cd", "settings", "LastT", "500"))
    
    Me.TextBoxBalkenMax.Text = GetSetting("rem2cd", "settings", "LastBalkenMax", "5")
    Me.TextBoxBalkenMin.Text = GetSetting("rem2cd", "settings", "LastBalkenMin", "2")
    
    Me.CheckBoxBold.Value = CBool(GetSetting("rem2cd", "settings", "LastBold", "false"))
    
    If (IsNumeric(GetSetting("rem2cd", "settings", "lastMode", "0"))) Then
        MultiPageMode.Value = GetSetting("rem2cd", "settings", "lastMode", "0")
    End If
End Sub


Private Sub SaveSettings()
    SaveSetting "rem2cd", "settings", "LastLength", TextBoxLen.Text
    SaveSetting "rem2cd", "settings", "LastCalib", TextBoxCalib.Text
    SaveSetting "rem2cd", "settings", "LastText", TextBoxTxt.Text
    SaveSetting "rem2cd", "settings", "LastWidth", TextBoxWidth.Text
    SaveSetting "rem2cd", "settings", "LastHeight", TextBoxHeight.Text
    SaveSetting "rem2cd", "settings", "LastLine", TextBoxLine.Text
    SaveSetting "rem2cd", "settings", "LastLineB", TextBoxLineB.Text
    SaveSetting "rem2cd", "settings", "LastTextOL", Me.TextBoxOL.Text
    SaveSetting "rem2cd", "settings", "LastTextOR", Me.TextBoxOR.Text
    SaveSetting "rem2cd", "settings", "LastTextUL", Me.TextBoxUL.Text
    SaveSetting "rem2cd", "settings", "LastTextSize", Me.TextBoxSize.Text
    
    SaveSetting "rem2cd", "settings", "LastBalkenMin", Me.TextBoxBalkenMin.Text
    SaveSetting "rem2cd", "settings", "LastBalkenMax", Me.TextBoxBalkenMax.Text
    
    SaveSetting "rem2cd", "settings", "LastMode", MultiPageMode.Value
    SaveSetting "rem2cd", "settings", "LastBold", CStr(Me.CheckBoxBold.Value)
    SaveSetting "rem2cd", "settings", "LastAbort", "false"
        
    SaveSetting "rem2cd", "settings", "LastL", Me.Left
    SaveSetting "rem2cd", "settings", "LastT", Me.Top
End Sub

