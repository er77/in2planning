Attribute VB_Name = "mSVPlanningAct"

Option Explicit
 

Sub p_CalculationForms(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

ActiveSheet.Cells(1, 1).Select
   
   If HypConnected(Empty) Then
    X = HypExecuteMenu(Empty, "Planning->Rules on Form")
   Else
     MsgBox "Please make a connection ", vbExclamation
   End If
  
 
l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  X = 0
 End If

  
End Sub


Sub p_CalculationPlanning(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
ActiveSheet.Cells(1, 1).Select

   If HypConnected(Empty) Then
    X = HypExecuteMenu(Empty, "Planning->Business Rules")
    If X <> 0 Then
      X = HypExecuteMenu(Empty, "Planning Ad Hoc->Business Rules")
    End If
   Else
     MsgBox "Please make a connection ", vbExclamation
   End If
 
l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  X = 0
 End If
 
End Sub


