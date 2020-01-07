Attribute VB_Name = "mActSubmitData"
Sub p_in2plnSubmitData(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

 Call p_CheckConnection
  
   X = MsgBox(" Upload data ?", vbOKCancel, "Planning Save Data")
       If X = 1 Then
        Application.EnableCancelKey = xlErrorHandler
         Call p_setExcelCalcOff
            X = HypMenuVSubmitData() 'HypExecuteMenu(ActiveSheet.Name, "Essbase->Submit Data") ' HypSubmitData(Empty) ' 'HypMenuVSubmitData()
         Call p_setExcelCalcOn
        Application.EnableCancelKey = xlInterrupt
       End If
  X = 0
  
End Sub

 Sub p_in2plnCellInfoComments(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 X = f_execSVMenu("Planning Ad Hoc->Comments")
  If X <> 0 Then
      p_execSVMenu ("Planning->Comments")
  End If
End Sub

 Sub p_in2plnCellInfoSupportingDetail(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 X = f_execSVMenu("Planning Ad Hoc->Supporting Details")
 If X <> 0 Then
     p_execSVMenu ("Planning->Supporting Details")
  End If
End Sub

 Sub p_in2plnCellInfoAttachmnet(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 X = f_execSVMenu("Planning Ad Hoc->Attachment")
  If X <> 0 Then
    p_execSVMenu ("Planning->Attachment")
  End If
End Sub
 Sub p_in2plnCellInfoHistory(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  X = f_execSVMenu("Planning Ad Hoc->History")
 If X <> 0 Then
    p_execSVMenu ("Planning->History")
  End If
End Sub



