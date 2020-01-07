Attribute VB_Name = "mActAdditionalService"
Option Explicit
 
    
           
 Sub p_in2plnConnections(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 Application.EnableCancelKey = xlErrorHandler
 vIsSVEnabled = Not vIsSVEnabled
 HypSetMenu (vIsSVEnabled)
 X = 0
 Err.Clear
  Application.EnableCancelKey = xlInterrupt
End Sub


 Sub p_in2plnPOVManager(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 Application.EnableCancelKey = xlErrorHandler
  isHypShowPov = Not isHypShowPov
  X = HypShowPov(isHypShowPov)
    'X = HypMenuVPOVManager()
 Err.Clear
  Application.EnableCancelKey = xlInterrupt
End Sub



Sub p_in2plnOptions(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
ActiveSheet.Cells(1, 1).Select

    If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        GoTo l_exit
    End If
    
    Application.EnableCancelKey = xlErrorHandler
     X = HypMenuVOptions()
    Application.EnableCancelKey = xlInterrupt
    
       If X <> 0 And X <> -55 Then GoTo ErrorHandler
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Options")
 
End Sub



Sub p_in2plnSheetInfo(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
    p_execSVMenu ("Smart View->Sheet Info")
End Sub

 

Sub p_in2plnsetAliasTable(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
   p_execSVMenu ("Planning Ad Hoc->Change Alias")
End Sub

 
 
Sub p_in2plnHypMenuVFunctionBuilder(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 Call p_CheckConnection
 
         Call p_setExcelCalcOff
            X = HypMenuVFunctionBuilder()
         Call p_setExcelCalcOn
  
 
End Sub



Sub p_in2plnAbout(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
    MsgBox "Hyperion Planning Custom Ribbon Panel " _
          & vbNewLine & " developer: er@essbase.ru "
    X = HypMenuVAbout
     ThisWorkbook.FollowHyperlink "https://github.com/er77/in2planning/issues"
 Err.Clear
 
End Sub

Sub p_in2plnAbout2(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  ThisWorkbook.FollowHyperlink "http://amegatech.ru/services.html"
 Err.Clear
 
End Sub







               

