VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rem2cdList 
   Caption         =   "Kalibration Liste"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   OleObjectBlob   =   "rem2cdList.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "rem2cdList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calibration Parameter listing tool for Corel Draw, written by Simon Klein (mail@simonklein.de)
'This Tool may be used or altered for private, non-commercial or educational use.
Public parentWindow As Variant
Private val As New Validater

Public Sub setcalib(calib As Double)
    If (calib > 0) Then
        TextBoxVal.Value = CStr(calib)
    End If
    
    Me.Show vbModeless
End Sub

Private Function validate() As Boolean
    Dim res As Boolean
    res = True
    
    res = res And val.validate(Me.TextBoxVal, valPosNumber)
    res = res And val.validate(Me.TextBoxName, valString)
    
    validate = res
End Function
    
Private Sub ButtonAdd_Click()
    If (Not validate) Then
        Exit Sub
    End If
    With ListBox1
        .AddItem
        .List(.ListCount - 1, 0) = TextBoxName.Text
        .List(.ListCount - 1, 1) = TextBoxVal.Text
    End With
    TextBoxVal.Value = ""
    TextBoxName.Value = ""
End Sub

Private Sub ButtonDel_Click()
    If (ListBox1.ListIndex <> -1) Then _
        ListBox1.RemoveItem (ListBox1.ListIndex)
End Sub

Private Sub ButtonEsc_Click()
    abort
End Sub

Private Sub ButtonLoad_Click()
    LoadData
End Sub

Private Sub ButtonOk_Click()
    SubmitValue
End Sub

Private Sub ButtonSave_Click()
    SaveData
End Sub

Private Sub LoadData()
    FillList _
        GetSetting("rem2cd", "values", "ListNames", ""), _
        GetSetting("rem2cd", "values", "ListValues", "")
End Sub

Private Sub FillList(inNames As String, inVals As String, Optional delimiter As String = "|")
    Dim names, vals As Variant
    
    names = Split(inNames, delimiter)
    vals = Split(inVals, delimiter)

    With ListBox1
        .Clear

        For i = 0 To UBound(names)
                If (names(i) <> "" And vals(i) <> "") Then
                    .AddItem
                    .List(.ListCount - 1, 0) = unEscapeString(names(i))
                    .List(.ListCount - 1, 1) = unEscapeString(vals(i))
                End If
        Next i
    End With
End Sub


Private Sub SaveData()
    Dim names, vals As String

    With ListBox1
        For i = 0 To .ListCount - 1
            names = names & "|" & escapeString(.List(i, 0))
            vals = vals & "|" & escapeString(.List(i, 1))
        Next i
    End With
    
    SaveSetting "rem2cd", "values", "ListNames", names
    SaveSetting "rem2cd", "values", "ListValues", vals
End Sub

Private Sub CommandButtonMes_Click()
    Me.Hide
    rem2cdCalib.Show (vbModeless)
    Set rem2cdCalib.parentWindow = Me
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SubmitValue
End Sub

Private Sub SubmitValue()
    If (ListBox1.ListIndex = -1) Then
        MsgBox "nichts ausgewählt"
        Exit Sub
    End If
    parentWindow.setcalib (CDbl(ListBox1.List(ListBox1.ListIndex, 1)))
    SaveData
    Unload Me
End Sub

Private Function escapeString(ByVal inputString As String, Optional delimiter As String = "|") As String
    escapeString = Replace(inputString, delimiter, "\" & delimiter)
End Function

Private Function unEscapeString(ByVal inputString As String, Optional delimiter As String = "|") As String
    unEscapeString = Replace(inputString, "\" & delimiter, delimiter)
End Function


Private Sub UserForm_Activate()
    'Set val = New Validater
    LoadData
End Sub

Sub abort()
    Me.Hide
    parentWindow.setcalib -1
    Unload Me
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    abort
End Sub
