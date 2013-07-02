VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_Start()
    'On Error Resume Next
    Dim bar As CommandBar
    Dim ButCalib, ButApply As CommandBarControl
    
    Set bar = Application.CommandBars("Toolbox")
    
    Set ButCalib = bar.Controls.AddCustomButton(vgcore.UICategories.cdrCmdCategoryMacros, "REM2Cd.gen.showCalibration", temporary:=True)
    Set ButApply = bar.Controls.AddCustomButton(vgcore.UICategories.cdrCmdCategoryMacros, "REM2Cd.gen.makeContainer", temporary:=True)

    ButCalib.ToolTipText = "Kalibration messen"
    ButCalib.DescriptionText = "Kalibration messen"
    ButCalib.Caption = "Kalibration messen"
    ButCalib.SetIcon2 ("guid://753696a1-5d8d-4adf-8120-7a423f425a5a")
    
    ButApply.ToolTipText = "µ-Balken hinzufügen"
    ButApply.DescriptionText = "µ-Balken hinzufügen"
    ButApply.Caption = "µ-Balken hinzufügen"
    ButApply.SetIcon2 ("guid://46327bd4-8bad-41c5-aba1-efa770b8e2c8")

End Sub
