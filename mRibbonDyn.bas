Attribute VB_Name = "mRibbonDyn"
Option Explicit

Public vCurrXMLGlobal

Function isSheetOTL() As Boolean
isSheetOTL = False
  If (InStr(UCase(ActiveSheet.Name), "OTL") > 0) Then
   isSheetOTL = True
  End If
   
End Function


Function isConnectPresent() As Boolean
 isConnectPresent = True 'isTextBoxPresent("ConnectQ")
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
 
 
 
 

 


