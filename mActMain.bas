Attribute VB_Name = "mActMain"
 Public StartExcelTime   As Single
 Public vIsSVEnabled As Boolean
 Public isHypShowPov As Boolean

 Sub p_in2plnShowPanel(ByVal vIRibbonControl As IRibbonControl)
    On Error Resume Next
     Application.EnableCancelKey = xlErrorHandler
     X = HypExecuteMenu(Empty, "Smart View->Panel")
     X = 0
     Application.Visible = False
     Err.Clear
     Application.Visible = True
     Application.EnableCancelKey = xlInterrupt
     
  
  ' Dim vtGrid As Variant
  ' Dim server As Variant
  ' Dim user As Variant
  ' Dim app As Variant
  ' Dim db As Variant
  ' Dim provider As Variant
  ' Dim conn As Variant
  ' Dim url As Variant
  ' sts = HypRetrieve(Empty)
  ' sts = HypGetSourceGrid(Empty, vtGrid)
  ' sts = HypGetConnectionInfo(server, user, pwd, app, db, conn, url, provider)
 

 
End Sub

Sub p_in2plnBackOutl(ByVal vIRibbonControl As IRibbonControl)
  On Error Resume Next
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
         End If
End Sub

Sub p_in2plnFreezePanes(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Sub


Sub p_in2plnAutoFilter(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  Selection.AutoFilter
End Sub


