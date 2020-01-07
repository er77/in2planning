Attribute VB_Name = "mActRefresh"

Sub p_in2plnRetrieve(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
      Call p_CheckConnection
      Call p_RefreshRibbonNow
      
      Call p_setExcelCalcOff
      
        X = HypShowPov(isHypShowPov)
        
     If vIsFirstRetrive Then
      If vIsUseNameDefault Then
       X = HypSetAliasTable(Empty, "Default")
      Else
       X = HypSetAliasTable(Empty, "none")
      End If
       vIsFirstRetrive = False
       p_RefreshRibbonNow
     Else
        X = HypMenuVRefresh()
     End If
         If X <> 0 Then GoTo ErrorHandler
        X = HypShowPov(isHypShowPov)
        X = 0
     
 Call p_setExcelCalcOn
 Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_in2plnRetrieve")
End Sub
 

Sub p_in2plnUndo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
   
  Call p_CheckConnection
  Call p_setExcelCalcOff
  
       X = HypMenuVUndo()
  Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_in2plnUndo")
End Sub

Sub p_in2plnRedo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
     
 Call p_CheckConnection
 Call p_setExcelCalcOff

      X = HypMenuVRedo()
Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
 
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_in2plnRedo")
X = 0
End Sub

