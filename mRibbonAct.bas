Attribute VB_Name = "mRibbonAct"
Option Explicit
 
 
 
 Public StartExcelTime   As Single
 Public vIsSVEnabled As Boolean
 Public isMDXSlice As Boolean
 
 Public isHypShowPov As Boolean

 


Public Sub p_CheckConnectionINT()

 If ActiveSheet Is Nothing Then
      MsgBox "active sheet is not determinated ", vbExclamation
    End
 End If

If Not isConnectPresent Then
    MsgBox "Can't find connection link. Please create Quick Connect", vbExclamation
    Call p_RefreshRibbonNow
  End
End If

 

End Sub

 
 


 

 



Sub p_SVon(ByVal vIRibbonControl As IRibbonControl)
 On Error GoTo ErrorHandler

  X = HypSetMenu(True)
    If X <> 0 Then GoTo ErrorHandler
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_SVon")
 
End Sub
 
 
           
 Sub p_Connections(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 Application.EnableCancelKey = xlErrorHandler
 vIsSVEnabled = Not vIsSVEnabled
 HypSetMenu (vIsSVEnabled)
 X = HypExecuteMenu(Empty, "Smart View->Panel")
 X = 0
 Err.Clear
  Application.EnableCancelKey = xlInterrupt
End Sub
 Sub p_ShowPanel(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 Application.EnableCancelKey = xlErrorHandler
 X = HypExecuteMenu(Empty, "Smart View->Panel")
 X = 0
 Application.Visible = False
' Application.Wait DateAdd("s", 2, Now)
 Err.Clear
 Application.Visible = True
  Application.EnableCancelKey = xlInterrupt
End Sub

 Sub p_POVManager(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
 Application.EnableCancelKey = xlErrorHandler
  isHypShowPov = Not isHypShowPov
  X = HypShowPov(isHypShowPov)
    'X = HypMenuVPOVManager()
 Err.Clear
  Application.EnableCancelKey = xlInterrupt
End Sub





 Sub p_CellInfoComments(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
Application.EnableCancelKey = xlErrorHandler
 X = HypExecuteMenu(Empty, "Planning Ad Hoc->Comments")
  If X <> 0 Then
    X = HypExecuteMenu(Empty, "Planning->Comments")
  End If
 Err.Clear
Application.EnableCancelKey = xlInterrupt
End Sub

 Sub p_CellInfoSupportingDetail(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
Application.EnableCancelKey = xlErrorHandler
 X = HypExecuteMenu(Empty, "Planning Ad Hoc->Supporting Details")
 If X <> 0 Then
    X = HypExecuteMenu(Empty, "Planning->Supporting Details")
  End If
 X = 0
 Err.Clear
Application.EnableCancelKey = xlInterrupt
End Sub

 Sub p_CellInfoAttachmnet(ByVal vIRibbonControl As IRibbonControl)    '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
Application.EnableCancelKey = xlErrorHandler
 X = HypExecuteMenu(Empty, "Planning Ad Hoc->Attachmnet")
 If X <> 0 Then
    X = HypExecuteMenu(Empty, "Planning->Attachmnet")
  End If
 X = 0
 Err.Clear
Application.EnableCancelKey = xlInterrupt
End Sub
 Sub p_CellInfoHistory(ByVal vIRibbonControl As IRibbonControl)  '(ByVal vIRibbonControl As IRibbonControl, ByRef vPressed)
 
On Error Resume Next
Application.EnableCancelKey = xlErrorHandler
 X = HypExecuteMenu(Empty, "Planning Ad Hoc->History")
 If X <> 0 Then
    X = HypExecuteMenu(Empty, "Planning->History")
  End If
 X = 0
 Err.Clear
Application.EnableCancelKey = xlInterrupt
End Sub

Sub p_Options(ByVal vIRibbonControl As IRibbonControl)
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




Sub p_Disconnect(ByVal vIRibbonControl As IRibbonControl)

   p_svcDisconnect
 
  Call p_SheetInfo(vIRibbonControl)
   
End Sub


Sub p_DisconnectOld(ByVal vIRibbonControl As IRibbonControl)
  
  X = HypDisconnect(Empty, True)
  
  Call p_SheetInfo(vIRibbonControl)
End Sub


Sub p_DisconnectAll(ByVal vIRibbonControl As IRibbonControl)
  
         Dim WS_Count As Integer
         Dim i As Integer
         
         WS_Count = ActiveWorkbook.Worksheets.Count
 
         For i = 1 To WS_Count
             X = HypDisconnect(ActiveWorkbook.Worksheets(i).Name, True)
         Next i
    Call p_SheetInfo(vIRibbonControl)
End Sub
 Sub p_setPOVe(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

ActiveSheet.Cells(1, 1).Select
Application.EnableCancelKey = xlErrorHandler
   X = HypMenuVPOVManager() 'HypExecuteMenu(Empty, "Essbase->POV")
Application.EnableCancelKey = xlInterrupt
l_exit:
    Exit Sub
ErrorHandler:
 If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
   Call p_ErrorHandler(X, "p_setPOVe")
 End If

 
End Sub
Sub p_SheetInfo(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
Application.EnableCancelKey = xlErrorHandler
vIsSVEnabled = False
X = HypSetMenu(False)

        X = HypExecuteMenu(Empty, "Smart View->Sheet Info")
        
   Err.Clear
 If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  If X <> 0 Then
   Call p_ErrorHandler(X, "p_SheetInfo")
  End If
 End If

 Application.EnableCancelKey = xlInterrupt
End Sub

Sub p_CheckConnection()

  If Not HypConnected(Empty) Then
      X = HypMenuVConnect()
  End If
 
End Sub



 

Sub p_setAliasTable(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

   p_CheckConnection
 
   Call p_setExcelCalcOff
     X = HypExecuteMenu(Empty, "Planning Ad Hoc->Change Alias")
  Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
   Call p_ErrorHandler(X, "p_setAliasTable")
 End If
 
End Sub

Sub p_Pivot(ByVal vIRibbonControl As IRibbonControl)
Dim namelist As Variant
Dim vallist As Variant
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
 
 Call p_setExcelCalcOff

  X = HypMenuVPivot()
   
   ' MsgBox "x1=" & X
   
   If X = 0 Then
     vIsSVEnabled = False
     X = HypSetMenu(vIsSVEnabled)
   Else
    If vIsSVEnabled Then
      SendKeys "%y3y1"
      SendKeys "%YQAYQW"
       DoEvents
      ' Application.Wait DateAdd("s", 2, Now)
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
 
 isMDXSlice = False
 
  
Call p_setExcelCalcOn

l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(X, "p_Pivot")
  
End Sub

Sub p_ZoomOut(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
 Call p_setExcelCalcOff
      X = HypMenuVZoomOut()
    
    If X <> 0 Then GoTo ErrorHandler
 
  X = HypShowPov(isHypShowPov)
   isMDXSlice = False
   
  Call p_setExcelCalcOn
   
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_ZoomOut")
   X = 0
End Sub

Sub p_ZoomIn(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
 
  Call p_setExcelCalcOff
       X = HypMenuVZoomIn()
     If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(isHypShowPov)
   isMDXSlice = False
   
      Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_ZoomIn")
 X = 0
End Sub

Sub p_KeepOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

   If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
 
  Call p_setExcelCalcOff
      X = HypMenuVKeepOnly()
     If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(isHypShowPov)
  
     Call p_setExcelCalcOn
     
   isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_KeepOnly")
 X = 0
End Sub

Sub p_RemoveOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
 
Call p_setExcelCalcOff
 X = HypMenuVRemoveOnly()

  If X <> 0 Then GoTo ErrorHandler
  X = HypShowPov(isHypShowPov)
  isMDXSlice = False

Call p_setExcelCalcOn

l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_RemoveOnly")
   X = 0
End Sub

Sub p_MemberSelect(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
 
 Application.EnableCancelKey = xlErrorHandler
     X = HypMenuVMemberSelection() ' HypExecuteMenu(Empty, "Planning Ad Hoc->Member Selection")
 Application.EnableCancelKey = xlInterrupt
     If X <> 0 Then GoTo ErrorHandler
     

  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  ' Call p_ErrorHandler(x, "p_MemberSelect")
 End If
 
End Sub
 

Sub p_Attributes(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
  Application.EnableCancelKey = xlErrorHandler
     X = HypExecuteMenu(Empty, "Planning Ad Hoc->Insert Attributes")
  Application.EnableCancelKey = xlInterrupt
     If X <> 0 Then GoTo ErrorHandler
     
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
End If
 
End Sub

Sub p_QueryDesigner(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
  If Not HypConnected(Empty) Then
      Call p_CheckConnection
  End If
  
  Application.EnableCancelKey = xlErrorHandler
     X = HypExecuteMenu(Empty, "Planning Ad Hoc->Query Designer")
  Application.EnableCancelKey = xlInterrupt
     If X <> 0 Then GoTo ErrorHandler
     

  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

  If (X = -15) And InStr(UCase(ActiveSheet.Name), "QUERY") Then
      'MsgBox "You can't connect from Query page. Please use other sheets", vbExclamation
       X = 0
    End
  End If


If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
End If
 
End Sub

Sub p_Retrieve(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 
 'Call p_CheckConnection
 Call p_setExcelCalcOff
 
   If Not HypConnected(ActiveSheet.Name) Then
      Call p_CheckConnection
   End If
  
 Call p_RefreshRibbonNow
   
        
        X = HypShowPov(isHypShowPov)
        
     If vIsFirstRetrive Then
      If vIsUseNameDefault Then
       X = HypSetAliasTable(Empty, "Default")
      Else
       X = HypSetAliasTable(Empty, "none")
      End If
       vIsFirstRetrive = False
     Else
        X = HypMenuVRefresh()
       ' x = HypRetrieve(ActiveSheet.Name)
     End If
         If X <> 0 Then GoTo ErrorHandler
        X = HypShowPov(isHypShowPov)
  
     If X <> 0 Then GoTo ErrorHandler
 Call p_setExcelCalcOn
   isMDXSlice = False
 
ErrorHandler:
'Call p_ErrorHandler(x, "p_Retrieve")
 X = 0
  Call p_setExcelCalcOn
End Sub
 

Sub p_Undo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
   
  Call p_CheckConnection
 
Call p_setExcelCalcOff
       X = HypMenuVUndo()
Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
 
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Undo")
    X = 0
End Sub

Sub p_Redo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
     
 Call p_CheckConnection
Call p_setExcelCalcOff

      X = HypMenuVRedo()
Call p_setExcelCalcOn

     If X <> 0 Then GoTo ErrorHandler
 
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:
'Call p_ErrorHandler(x, "p_Redo")
X = 0
End Sub
'

Sub p_SplitReports(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
     

 Call p_CheckConnection
  
   Call p_setExcelCalcOff_
    X = HypExecuteMenu(Empty, "Planning Ad Hoc->Visualize")
  Call p_setExcelCalcOn_

l_exit:
    Exit Sub
ErrorHandler:
If X = -15 Then
   MsgBox " Probably you need to change SmartView Language to English. Go To Additionals->Options ", vbExclamation
   X = 0
 Else
  ' Call p_ErrorHandler(x, "p_CalculationEssBase")
  X = 0
 End If
 
End Sub

 

 




 
Sub p_SubmitData(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 Application.EnableCancelKey = xlErrorHandler
  
   X = MsgBox(" Upload data ?", vbOKCancel, "Essbase Save Data")
       If X = 1 Then
         
         Call p_setExcelCalcOff
            X = HypMenuVSubmitData() 'HypExecuteMenu(ActiveSheet.Name, "Essbase->Submit Data") ' HypSubmitData(Empty) ' 'HypMenuVSubmitData()
         Call p_setExcelCalcOn
       Else
         X = 0
       End If
       If X <> 0 Then GoTo ErrorHandler
  
  isMDXSlice = False
 
ErrorHandler:
End Sub
 
 
Sub p_HypMenuVFunctionBuilder(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
  isMDXSlice = False
         Call p_setExcelCalcOff
            X = HypMenuVFunctionBuilder()
         Call p_setExcelCalcOn
l_exit:
    Exit Sub
ErrorHandler:

 
End Sub

Sub p_SubmitDataVORefresh(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

 Call p_CheckConnection
 
  
   X = MsgBox(" Upload data ?", vbOKCancel, "Essbase Save Data")
       If X = 1 Then
         Call p_setExcelCalcOff
            X = HypSubmitSelectedRangeWithoutRefresh(Null, False, True, True)
         Call p_setExcelCalcOn
       Else
         X = 0
       End If
       If X <> 0 Then GoTo ErrorHandler
  
  isMDXSlice = False
l_exit:
    Exit Sub
ErrorHandler:

 
End Sub

Sub p_About(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

 
    MsgBox "Essbase Custom Panel " _
          & vbNewLine & " developer: er@essbase.ru " & vbNewLine _
          & vbNewLine & " https://github.com/er77/EssBase.ACT/issues "
       
    X = HypMenuVAbout
    
 Err.Clear
 
End Sub

Sub p_About2(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

  
  ThisWorkbook.FollowHyperlink "http://amegatech.ru/services.html"
    
 Err.Clear
 
End Sub
               

