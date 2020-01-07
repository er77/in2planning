Attribute VB_Name = "mActReportAndCalc"
Sub p_in2plnCalculationPlanning(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
ActiveSheet.Cells(1, 1).Select
    X = f_execSVMenu("Planning->Business Rules")
    If X <> 0 Then
       p_execSVMenu ("Planning Ad Hoc->Business Rules")
    End If
End Sub

Sub p_in2plnCalculationForms(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

   ActiveSheet.Cells(1, 1).Select
     X = f_execSVMenu("Planning->Rules on Form")
     If (X <> 0) Then
       MsgBox "Please open form"
     End If
     
   
End Sub


Sub p_in2plnQueryDesigner(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  ActiveSheet.Cells(1, 1).Select
    X = HypShowPov(True)
    X = HypExecuteMenu(Empty, "Planning Ad Hoc->Query Designer")
 
End Sub


