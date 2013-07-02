VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Validater Class for validating (numerical) input in textboxes. Written by Simon Klein (mail@simonklein.de)
Option Explicit

Private highlightColor As Long
Private OldVals As New Dictionary

Public Enum valType
    valEmpty
    valNumber
    valPosNumber
    valNegNumber
    valNatNumber
    valString
    valPosNumOrEmpty
End Enum

Public Function validate(contr As Control, typ As valType) As Boolean
    If (Not check(contr, typ)) Then
        fail contr
        validate = False
        Exit Function
    Else
        restore contr
        validate = True
        Exit Function
    End If
    validate = True
End Function

Public Function test(bla As String) As Boolean
    test = False
End Function

Private Function check(contr As Control, typ As valType)
    Dim t As String
    t = contr.Value
    
    If (typ = valString) Then
        If (Len(t) > 0) Then
            check = True
            Exit Function
        Else
            check = False
            contr.ControlTipText = "Eingabe erforderlich"
            Exit Function
        End If
    End If
    
    If (typ = valEmpty) Then
        If (Len(t) = 0) Then
            check = True
            Exit Function
        Else
            check = False
            contr.ControlTipText = "Leeres Feld erforderlich"
            Exit Function
        End If
    End If
    
    If (typ = valPosNumOrEmpty) Then
        If (Len(t) = 0 Or check(contr, valPosNumber)) Then
            check = True
            Exit Function
        Else
            check = False
            Exit Function
        End If
    End If
    
    If (IsNumeric(t)) Then
        Select Case typ
            Case valNumber
                check = True
                Exit Function
            
            Case valPosNumber
                If (CDbl(t) >= 0) Then
                    check = True
                    Exit Function
                Else
                    contr.ControlTipText = "Die Eingabe muss eine positive Zahl sei"
                End If
                    
            Case valNegNumber
                If (CDbl(t) <= 0) Then
                    check = True
                    Exit Function
                Else
                    contr.ControlTipText = "Die Eingabe muss eine negative Zahl sei"
                End If
                
            Case valNatNumber
                If (Round(CDbl(t), 0) = CDbl(t)) Then
                    check = True
                    Exit Function
                Else
                    contr.ControlTipText = "Die Eingabe muss eine natürliche Zahl sei"
                End If
        End Select
    Else
        contr.ControlTipText = "Die Eingabe muss eine Zahl sein"
    End If
    check = False
End Function

Private Sub fail(contr As Control)
    save contr
    contr.BackColor = highlightColor
End Sub

Private Sub save(contr As Control)
    If (Not OldVals.Exists(contr)) Then _
        OldVals.Add contr, contr.BackColor
End Sub

Private Sub restore(contr As Control)
    contr.ControlTipText = ""
    If (OldVals.Exists(contr)) Then
        contr.BackColor = OldVals(contr)
        OldVals.Remove contr
    End If
End Sub

Private Sub Class_Initialize()
    'Set OldVals = New Dictionary
    highlightColor = vbRed
End Sub
