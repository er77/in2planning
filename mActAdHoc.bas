Attribute VB_Name = "mActAdHoc"

Sub p_in2plnPivot(ByVal vIRibbonControl As IRibbonControl)
Dim namelist As Variant
Dim vallist As Variant
On Error GoTo ErrorHandler
  
 Call p_CheckConnection
 Call p_setExcelCalcOff

  X = HypMenuVPivot()
  
   If X = 0 Then
     vIsSVEnabled = False
     X = HypSetMenu(vIsSVEnabled)
   Else
    If vIsSVEnabled Then
      SendKeys "%y3y1"
      SendKeys "%YQAYQW"
       DoEvents
     vIsSVEnabled = False
     X = HypSetMenu(vIsSVEnabled)
     isHypShowPov = True
     X = HypShowPov(isHypShowPov)
    Else
     vIsSVEnabled = True
     X = HypSetMenu(vIsSVEnabled)
     isHypShowPov = True
     X = HypShowPov(isHypShowPov)
    End If
   End If
   
Call p_setExcelCalcOn

l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_in2plnPivot")
  
End Sub


Sub p_in2plnZoomOut(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

    Call p_CheckConnection
    Call p_setExcelCalcOff
    
      X = HypMenuVZoomOut()
       If X <> 0 Then GoTo ErrorHandler
     X = HypShowPov(isHypShowPov)
  Call p_setExcelCalcOn
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_in2plnZoomOut")
   X = 0
End Sub

Sub p_in2plnZoomIn(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

      Call p_CheckConnection
      Call p_setExcelCalcOff
      
       X = HypMenuVZoomIn()
        If X <> 0 Then GoTo ErrorHandler
       X = HypShowPov(isHypShowPov)
      Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_in2plnZoomIn")
 X = 0
End Sub

Sub p_in2plnKeepOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
      Call p_CheckConnection
      Call p_setExcelCalcOff
      
      X = HypMenuVKeepOnly()
        If X <> 0 Then GoTo ErrorHandler
      X = HypShowPov(isHypShowPov)
     Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_in2plnKeepOnly")
 X = 0
End Sub

Sub p_in2plnRemoveOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  
      Call p_CheckConnection
      Call p_setExcelCalcOff
      
        X = HypMenuVRemoveOnly()
        
        If X <> 0 Then GoTo ErrorHandler
          X = HypShowPov(isHypShowPov)
          X = 0
        Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(X, "p_in2plnRemoveOnly")
   X = 0
End Sub

Sub p_in2plnMemberSelect(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

    Call p_CheckConnection
    Call p_setExcelCalcOff
 
 Application.EnableCancelKey = xlErrorHandler
     X = HypMenuVMemberSelection()
 Application.EnableCancelKey = xlInterrupt
     If X <> 0 Then GoTo ErrorHandler
   Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(X, "p_in2plnMemberSelect")

 
End Sub
 

Sub p_in2plnAttributes(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  
      Call p_execSVMenu("Planning Ad Hoc->Insert Attributes")
    
l_exit:
    Exit Sub
ErrorHandler:
   Call p_ErrorHandler(X, "p_in2plnAttributes")
End Sub

