Attribute VB_Name = "mRibbonSVC"
 Option Explicit
  
 Public isFirstOptionQ As Boolean
 Public iSHstbar As Boolean
 Public vIsSuppresPressed As Boolean
 
 
  Public vIRibbonUI As IRibbonUI
 
Sub p_iSHstbar()
On Error Resume Next
iSHstbar = False
    If AddIns("Hstbar").Installed Then  'Hyperion.CommonAddin
        iSHstbar = True
    End If
   If Err.Number <> 0 Then
     iSHstbar = False
   Err.Clear
   End If
End Sub

Sub p_RefreshRibbonNow()
On Error Resume Next
  Call p_restoreOptions
   vIRibbonUI.Invalidate
    If Err.Number > 0 Then
         Err.Clear
    End If
    isFirstOptionQ = True
End Sub

Sub p_OnRibbonLoad(vRibbon As IRibbonUI)
   Call p_iSHstbar
 
 Application.MultiThreadedCalculation.Enabled = True
 Application.AutoRecover.Time = 7
 Application.EnableEvents = True
 
    Set vIRibbonUI = vRibbon
    vItWasOtlPage = False
    isHypShowPov = True
      
      vIsSVEnabled = False
      vIsSuppresPressed = False
 
       X = HypSetMenu(iSHstbar)
 
      vIsFirstRetrive = True
      vIsUseNameDefault = True
      isFirstOptionQ = True
  
 
End Sub

 


Function isSheetOTL() As Boolean
isSheetOTL = False
  If (InStr(UCase(ActiveSheet.Name), "OTL") > 0) Then
   isSheetOTL = True
  End If
   
End Function


 

 
 
  Sub p_IsVisible(ByVal vIRibbonControl As IRibbonControl, ByRef vReturnValue)
  
    vReturnValue = False
  
  If vModeAnalyse = 0 Then
    Select Case vIRibbonControl.ID
        Case "grp_RData"
           vReturnValue = True
        Case "b_SheetInfo"
            vReturnValue = True
        Case "grp_Options"
            vReturnValue = True
        Case "grp_Main0"
            vReturnValue = True
        Case "grp_Refresh"
          vReturnValue = True
     End Select
  End If
     
 End Sub
 




