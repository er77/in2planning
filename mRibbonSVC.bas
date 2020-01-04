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
      vConnName = ""
      vAppName = ""
      vDbName = ""
      vUserName = ""
      vPassword = ""
      vFriendlyName = ""
      vIsSVEnabled = False
      vIsSuppresPressed = False
      vCurrEnv = 0
       X = HypSetMenu(iSHstbar)
      ' X = HypExecuteMenu(Empty, "Smart View->Panel")
      vIsFirstRetrive = True
      vIsUseNameDefault = True
      isFirstOptionQ = True
      vCurrXMLGlobal = ""
 
End Sub

 







