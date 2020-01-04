Attribute VB_Name = "mSVCCommon"
Option Explicit
Option Compare Text

Dim vCurrQueryTime
  
Global Const isDebug = True
 
Public vIsFirstRetrive  As Boolean
Public vItWasOtlPage  As Boolean
 Public vActiveCell
 Public vLastSheetName As String
 
 
 Sub p_svcDisconnect()

    Dim objWorkbook As Workbook
    Set objWorkbook = ActiveWorkbook
 
    X = HypDeleteMetaData(objWorkbook, True, True)

   Set objWorkbook = Nothing
End Sub

 Sub p_svcDisconnectSheet()

    Dim objSheet As Worksheet
    Set objSheet = ActiveSheet
 
    X = HypDeleteMetaData(objSheet, False, True)

   Set objSheet = Nothing
End Sub

 
 

Sub p_WriteStatusBarTime()
 'vCurrQueryTime = (Now - vCurrQueryTime)
 
Application.StatusBar = " Exec Time: " & DateDiff("s", vCurrQueryTime, Now) & " sec "
'Application.OnTime Now + TimeSerial(0, 0, 90), "p_ClearStatusBar"
End Sub

Sub p_ClearStatusBar()
On Error Resume Next
    Application.StatusBar = False
End Sub


Sub p_setExcelCalcOff_()
  On Error Resume Next
  
  If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        End
    End If
    
    vCurrQueryTime = Now
     Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.UsedRange.EntireRow.Hidden = False
    Set vActiveCell = Range(ActiveCell.Address)
        
    If Err.Number <> 0 Then
       Err.Clear
    End If
End Sub

Sub p_setExcelCalcOn_()
   On Error Resume Next

    Call p_WriteStatusBarTime
      
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.UsedRange.EntireRow.Hidden = False
    
    If Err.Number <> 0 Then
       Err.Clear
    End If

End Sub

Sub p_setExcelCalcOff()

  If ActiveSheet Is Nothing Then
        MsgBox "active sheet is not determinated "
        End
    End If
    
  On Error Resume Next
    vCurrQueryTime = Now
     Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
   ' Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.UsedRange.EntireRow.Hidden = False
    Set vActiveCell = Range(ActiveCell.Address)
    
   
     
    If Err.Number <> 0 Then
       Err.Clear
    End If
End Sub

Sub p_setExcelCalcOn_INT()
   On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    Application.ScreenUpdating = True
   ' Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.UsedRange.EntireRow.Hidden = False
    
    ActiveSheet.UsedRange.Select
    Selection.NumberFormat = "#,##0.00"
    
    Application.CutCopyMode = False
    vActiveCell.Select
    
    If Err.Number <> 0 Then
       Err.Clear
    End If
    
End Sub

Sub p_setExcelCalcOn()
   On Error Resume Next
   
    Call p_WriteStatusBarTime
      
      X = HypShowPov(isHypShowPov)
      X = 0
    Call p_setExcelCalcOn_INT
    
     ActiveSheet.Outline.ShowLevels RowLevels:=1
End Sub
Function CheckIfSheetExists(SheetName As String) As Boolean
      CheckIfSheetExists = False
    Dim ws As Worksheet
      For Each ws In Worksheets
        If SheetName = ws.Name Then
          CheckIfSheetExists = True
          Exit Function
        End If
      Next ws
End Function



Sub p_BackOutl(ByVal vIRibbonControl As IRibbonControl)

     Dim vCurrSheetName
       vCurrSheetName = ActiveSheet.Name
       
      If (InStr(UCase(vCurrSheetName), "OTL") > 0) Then
        If CheckIfSheetExists(vLastSheetName) Then
          On Error Resume Next
            Worksheets(vLastSheetName).Activate
             Err.Clear
             vLastSheetName = ""
          Exit Sub
        End If
         Exit Sub
       End If
       
        vLastSheetName = vCurrSheetName
         If CheckIfSheetExists("OTL") Then
          On Error Resume Next
             Worksheets("OTL").Activate
          Err.Clear
          Exit Sub
         'Else
         '  MsgBox "This button will activate  ""OTL"" page"
         End If
         
      
       
End Sub

Sub p_FreezePanes(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
 
End Sub


Sub p_AutoFilter(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  Selection.AutoFilter
 
End Sub

 
Sub p_ErrorHandler(vErrNum As Long, p As String)

Dim vErrStr

Dim vButton
Dim vErrorHandler
 vErrStr = ""
 

 
   If (Err.Number <> 0) Then
       vErrStr = vErrStr & "Error # " & Str(Err.Number) & " was generated by " _
         & Err.source & vbCrLf & "Error Line: " & Erl & vbCrLf & Err.Description
       vButton = vbCritical
    End If
      
   If vErrNum = -4 And (Err.Number = 0) Then
      MsgBox "The connection was lost. Please make new one  " & p, vbExclamation
      p_svcDisconnect
   End
   End If
   
  If vErrNum = 4 And (Err.Number = 0) Then
      MsgBox " Unknown SmartView Error . Please Restart Excel ", vbExclamation
      p_svcDisconnect
   End
   End If

  If isDebug Then
    If vErrNum <> 0 And (vErrNum <> 4) And (vErrNum <> 1020021) Then
      vErrStr = vErrStr & vbCrLf & getErrorText(vErrNum)
    End If
   vErrStr = vErrStr & vbCrLf & p
  End If

 vButton = vbCritical
 vErrorHandler = "Error"
 
 If (Err.Number = 0) Then
   vButton = vbExclamation
   vErrorHandler = "Warning"
 End If
 
 If vErrStr <> "" Then
    MsgBox vErrStr, vButton, vErrorHandler
 End If
 
X = 0

If Err.Number <> 0 Then
   Err.Clear
End If
 
 p_setExcelCalcOn_

End

End Sub





 
