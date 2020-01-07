Attribute VB_Name = "mActRefreshOpt"
 Option Explicit
 
 Public vHSV_SUPPRESSCOLUMNS_MISSING As Boolean
 Public vHSV_SUPPRESSCOLUMNS_ZEROS As Boolean
 Public vHSV_SUPPRESS_MISSINGBLOCKS As Boolean
 Public vHSV_SUPPRESSROWS_MISSING As Boolean
 Public vHSV_SUPPRESSROWS_ZEROS As Boolean
 Public vHSV_MEMBER_DISPLAY  As Integer
 Public vHSV_ZOOMIN As Integer
 Public vHSV_ANCESTOR_POSITION As Integer

 Public vHSV_MISSING_LABEL As String
 Public vHSV_INDENTATION As Integer
 Public vHSV_INCLUDE_SELECTION As Boolean
 Public vIsDeleteOrphans As Boolean
 Public vIsSuppressOnPivot As Boolean
 
 Public vIsUseNameDefault As Boolean
 
 
 Public vIsAlreadyHide As Boolean
 
 Public vCurrEnv As Integer
 Public vModeAnalyse As Integer
 
 Public X As Long
 
 

 Sub p_onActionINT(vIRibbonControlID As String, ByVal vSelectedValue)
On Error GoTo ErrorHandler
Dim isEssOption

isEssOption = True
isFirstOptionQ = True

    Select Case vSelectedValue
     '   Case "mn_Supr0"
     '       vHSV_SUPPRESS_MISSINGBLOCKS = True
     '       vHSV_SUPPRESSROWS_MISSING = False

        Case "mn_Supr1"
            vHSV_SUPPRESSROWS_MISSING = True
            vHSV_SUPPRESS_MISSINGBLOCKS = True
            vHSV_SUPPRESSROWS_ZEROS = True
            vHSV_SUPPRESSCOLUMNS_MISSING = True
            vHSV_SUPPRESSCOLUMNS_ZEROS = True
            
        Case "mn_Supr2"
            vHSV_SUPPRESSROWS_MISSING = True
            vHSV_SUPPRESS_MISSINGBLOCKS = False
            vHSV_SUPPRESSROWS_ZEROS = False
             Call p_mnMissInit
             
        Case "mn_Supr6"
            vHSV_SUPPRESS_MISSINGBLOCKS = False
            vHSV_SUPPRESSROWS_MISSING = False
              Call p_mnMissInit
              
        Case "mn_SuprClmn2"
            vHSV_SUPPRESSCOLUMNS_MISSING = True

        Case "mn_SuprClmn6"
            vHSV_SUPPRESSCOLUMNS_MISSING = False

        Case "mn_Zoom0"
            vHSV_ZOOMIN = 2

        Case "mn_Zoom1"
            vHSV_ZOOMIN = 0

        Case "mn_Zoom2"
            vHSV_ZOOMIN = 1

        Case "mn_Intend0"
            vHSV_INDENTATION = 0

        Case "mn_Intend1"
            vHSV_INDENTATION = 1

        Case "mn_Intend2"
            vHSV_INDENTATION = 2
            
        Case "mn_Show1"
            vHSV_MEMBER_DISPLAY = 2
                
        Case "mn_Show2"
            vHSV_MEMBER_DISPLAY = 1
            
      Case "mn_SubTot0"
            vHSV_ANCESTOR_POSITION = 0

        Case "mn_SubTot1"
            vHSV_ANCESTOR_POSITION = 1

         Case "mn_Selection0"
            vHSV_INCLUDE_SELECTION = True
 
        Case "mn_Selection1"
            vHSV_INCLUDE_SELECTION = False
 

         Case "mn_DelOrp0"
            vIsDeleteOrphans = True
            
        Case "mn_DelOrp1"
            vIsDeleteOrphans = False
 
         Case "mn_Env0"
            vCurrEnv = 0
            isEssOption = False
            
         Case "mn_Env1"
            vCurrEnv = 1
            isEssOption = False
                    
         Case "mn_Env2"
            vCurrEnv = 2
            isEssOption = False
            
         Case "mn_Env3"
            vCurrEnv = 3
            isEssOption = False
            
         Case "mn_Mode0"
            vModeAnalyse = 0
            p_RefreshRibbonNow
            isEssOption = False
            
         Case "mn_Mode1"
            vModeAnalyse = 1
            p_RefreshRibbonNow
            isEssOption = False
            
         Case "mn_Mode2"
            vModeAnalyse = 2
            p_RefreshRibbonNow
            isEssOption = False

    End Select

  Call p_WriteGlobalProperty(vIRibbonControlID, vSelectedValue)

    If ActiveSheet Is Nothing Then
        GoTo l_exit
    End If
 
    If isEssOption Then
      Call p_setCurrentOptions(vIRibbonControlID, vSelectedValue)
    End If

l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_onAction")
End Sub
 
  Sub pb_onAction_supr(vIRibbonControl As IRibbonControl, ByRef returnedVal)
 
 returnedVal = vIsSuppresPressed
 
   
 
End Sub
 
 Sub p_onAction_supr(vIRibbonControl As IRibbonControl, ByVal vSelectedValue, Optional vOptional)
 
 vIsSuppresPressed = Not vIsSuppresPressed
 
  If vIsSuppresPressed Then
   Call p_onActionINT(vIRibbonControl.ID, "mn_Supr1")
  Else
   Call p_onActionINT(vIRibbonControl.ID, "mn_Supr6")
  End If
 
End Sub
 
 Sub p_in2plnOnAction(vIRibbonControl As IRibbonControl, ByVal vSelectedValue, Optional vOptional)

   Call p_onActionINT(vIRibbonControl.ID, vSelectedValue)
 
End Sub
 
Sub p_setCurrentOptions(vIRibbonControlID As String, ByVal vSelectedValue)

   
     Select Case vIRibbonControlID
        Case "mn_Supr"
             Call p_SetOptionSuppress
            
        Case "mn_Zoom"
            Call p_SetOption(HSV_ZOOMIN, vHSV_ZOOMIN)
            
        Case "mn_Intend"
             Call p_SetOption(HSV_INDENTATION, vHSV_INDENTATION)

        Case "mn_Show"
              Call p_SetOption(HSV_MEMBER_DISPLAY, vHSV_MEMBER_DISPLAY)
              
        Case "mn_SubTot"
              Call p_SetOption(HSV_ANCESTOR_POSITION, vHSV_ANCESTOR_POSITION)
        Case "mn_Selection"
              Call p_SetOption(HSV_INCLUDE_SELECTION, vHSV_INCLUDE_SELECTION)
    End Select
End Sub

Sub p_getSelectedItemIDINT(vIRibbonControlID As String, ByRef itemID As Variant)
Dim isEssOption

    If ActiveSheet Is Nothing Then

      If Not vIsAlreadyHide Then
       vIsAlreadyHide = True
        vIsSVEnabled = False
        X = HypSetMenu(False)
      End If
      
       isFirstOptionQ = False
       Exit Sub
    End If
    
   itemID = Null
 'DoEvents
 
   itemID = f_ReadGlobalProperty(vIRibbonControlID)
   
   isEssOption = True
   If (IsNull(itemID)) Then
     Select Case vIRibbonControlID
        Case "mn_Supr"
              
              itemID = "mn_Supr6"
              
              vIsSuppresPressed = False
              
              vHSV_SUPPRESS_MISSINGBLOCKS = False
              vHSV_SUPPRESSROWS_MISSING = False
              Call p_SetOption(HSV_NUMBER_OF_UNDO_ACTION, 4)
               If vIsSVEnabled = True Then
                X = HypSetMenu(True)
                 Else
               If iSHstbar Then
                 X = HypSetMenu(False)
               End If
                End If
                             
               Call p_mnMissInit
               
        Case "mn_SuprClmn"
              itemID = "mn_SuprClmn6"
              vHSV_SUPPRESSCOLUMNS_MISSING = False
               
        Case "mn_Zoom"
             itemID = "mn_Zoom0"
             vHSV_ZOOMIN = 2
             
        Case "mn_Intend"
              itemID = "mn_Intend0"
                vHSV_INDENTATION = 0
                
        Case "mn_Show"
              itemID = "mn_Show1"
              vHSV_MEMBER_DISPLAY = 2
              
        Case "mn_AddSupr"
              itemID = "mn_Supr5"
              vIsSuppressOnPivot = False
              
        Case "mn_SubTot"
               itemID = "mn_SubTot0"
               vHSV_ANCESTOR_POSITION = 0
               
        Case "mn_Mode"
               itemID = "mn_Mode0"
               vModeAnalyse = 0
               
        ' Case "mn_Alias"
        '
        '    If vIsUseNameDefault Then
        '      itemID = "mn_Alias0"
        '    Else
        '      itemID = "mn_Alias1"
        '    End If
               
         Case "mn_Selection"
            itemID = "mn_Selection0"
            vHSV_INCLUDE_SELECTION = True
            
         Case "mn_DelOrp"
            itemID = "mn_DelOrp0"
            vIsDeleteOrphans = True
            
    End Select
    
    If isEssOption Then
      Call p_setCurrentOptions(vIRibbonControlID, itemID)
    End If
Else
     If InStr(vIRibbonControlID, "mn_Show") = 0 Then
        Call p_onActionINT(vIRibbonControlID, itemID)
     End If
 End If
    
End Sub

Sub p_getSelectedItemID(vIRibbonControl As IRibbonControl, ByRef itemID As Variant, Optional vOptional)
 

Call p_getSelectedItemIDINT(vIRibbonControl.ID, itemID)
    
  StartExcelTime = Timer
     
End Sub

 
Sub p_mnMissInit()
 
    vHSV_MISSING_LABEL = "" ' "#NumericZero"
    Call p_SetOption(HSV_MISSING_LABEL, vHSV_MISSING_LABEL)
    Call p_SetOption(HSV_NOACCESS_LABEL, "'-")
End Sub
 
 Sub p_SetOptionSuppress()
    Call p_SetOption(HSV_SUPPRESSCOLUMNS_MISSING, vHSV_SUPPRESSCOLUMNS_MISSING)
    Call p_SetOption(HSV_SUPPRESSCOLUMNS_ZEROS, vHSV_SUPPRESSCOLUMNS_ZEROS)
    Call p_SetOption(HSV_SUPPRESS_MISSINGBLOCKS, vHSV_SUPPRESS_MISSINGBLOCKS)
    Call p_SetOption(HSV_SUPPRESSROWS_MISSING, vHSV_SUPPRESSROWS_MISSING)
    Call p_SetOption(HSV_SUPPRESSROWS_ZEROS, vHSV_SUPPRESSROWS_ZEROS)
 End Sub
             
Sub p_restoreOptions()
Dim itemID
 'ActiveSheet.Cells(1, 1).Select
  itemID = 0
        Call p_getSelectedItemIDINT("mn_Supr", itemID)
        Call p_getSelectedItemIDINT("mn_Zoom", itemID)
        Call p_getSelectedItemIDINT("mn_Intend", itemID)
        Call p_getSelectedItemIDINT("mn_Show", itemID)
        Call p_getSelectedItemIDINT("mn_Alias", itemID)
        Call p_getSelectedItemIDINT("mn_AddSupr", itemID)
        Call p_getSelectedItemIDINT("mn_SubTot", itemID)
        Call p_getSelectedItemIDINT("mn_Mode", itemID)
        Call p_getSelectedItemIDINT("mn_Selection", itemID)
        Call p_getSelectedItemIDINT("mn_Env", itemID)
        Call p_getSelectedItemIDINT("mn_DelOrp", itemID)
End Sub
 

Sub p_SetOption(vStrOption As Variant, vOptionValue As Variant)
Dim strErr As String
'Dim cnt As Integer
'https://docs.oracle.com/cd/E38438_01/epm.111223/sv_developer/frameset.htm?ch13s06.html



  
    If Not (vStrOption = 999) Then
     X = HypSetOption(vStrOption, vOptionValue, ActiveSheet.Name)
     
      If X <> 0 Then
        strErr = X & " HypSetOption1 "
         GoTo ErrorHandler
      End If
    End If
      
      If (vStrOption < 17) And Not (vStrOption = 999) Then
        X = HypSetSheetOption(Empty, vStrOption, vOptionValue)
        
        If X <> 0 Then
          strErr = X & " HypSetSheetOption "
           GoTo ErrorHandler
        End If
       End If
 
    'And (cnt > 0)

   If ((vStrOption < 10) And ((vStrOption <> 1))) And (vStrOption <> 999) And iSHstbar Then
    X = HypSetGlobalOption(vStrOption, vOptionValue)
    If X <> 0 Then
        strErr = X & " HypSetGlobalOption "
         GoTo ErrorHandler
      End If
   End If
   ' Not (cnt = 0) And
    If (vStrOption <> 999) And iSHstbar Then
          X = HypSetOption(vStrOption, vOptionValue, "")
            If X <> 0 Then
             strErr = X & " HypSetOption2 "
         GoTo ErrorHandler
    End If
   
  
 End If
l_exit:
     Exit Sub
ErrorHandler:
On Error Resume Next
 If X = -74 Then
   ActiveSheet.Cells(1, 1).Select
 Else
   Call p_ErrorHandler(X, "p_SetOption" & vStrOption & "  " & vOptionValue)
 End If
 
End Sub
Function getCountWB() As Integer
  
     getCountWB = ThisWorkbook.Sheets.Count
 End Function

Sub p_SetAllOptions()

           Call p_SetOptionSuppress
           Call p_SetOption(HSV_MEMBER_DISPLAY, vHSV_MEMBER_DISPLAY)
           Call p_SetOption(HSV_ZOOMIN, vHSV_ZOOMIN)
           Call p_SetOption(HSV_ANCESTOR_POSITION, vHSV_ANCESTOR_POSITION)
           Call p_SetOption(HSV_MISSING_LABEL, vHSV_MISSING_LABEL)
           
l_exit:
     Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_SetAllOptions")
     
End Sub
