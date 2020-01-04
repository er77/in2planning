Attribute VB_Name = "mSVCComment"

Sub p_setFormulaToComment(ByVal vSheetName As String, ByVal vCurrCellAdrees As String, ByVal vCommentText As String)
    On Error Resume Next
  With Worksheets(vSheetName).Range(vCurrCellAdrees)
        If Not .Comment Is Nothing Then .Comment.Delete
        .AddComment "er"
        .Comment.Text Now() & Chr(10) & Chr(13) & vCommentText    'No need the assignment sign "=" after .Comment.Text
    End With
  If Err.Number <> 0 Then
       Worksheets(vSheetName).Range(vCurrCellAdrees).Comment.Delete
  End If
   Err.Clear
End Sub

Function getCellComment(ByVal vSheetName As String, ByVal vCurrCellAdrees As String)
    On Error Resume Next
      getCellComment = Worksheets(vSheetName).Range(vCurrCellAdrees).Comment.Text
   Err.Clear
End Function

Sub p_addToCommentFormula(ByVal vSheetName As String, ByVal vCurrCellAdrees As String, ByVal vCommentText As String)
    On Error Resume Next
  With Worksheets(vSheetName).Range(vCurrCellAdrees)
         vCommentText = .Comment.Text & Chr(10) & Chr(13) & Chr(10) & Chr(13) & vCommentText
        .Comment.Text vCommentText
    End With
  If Err.Number <> 0 Then
       Worksheets(vSheetName).Range(vCurrCellAdrees).Comment.Delete
  End If
   Err.Clear
End Sub

Sub p_deleteNamedRange(vCurrNameRangeName As String)
 On Error Resume Next
 
 Scope.Names(vCurrNameRangeName).Delete

End Sub

Function getClearInt(ByVal vCurrStr As String)
  Dim i As Long, vCodesToClean As Variant
  vCodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 96, 126, 127, 127, 129, 141, 143, 144, 157, 160)
  For i = LBound(vCodesToClean) To UBound(vCodesToClean)
    If InStr(vCurrStr, Chr(vCodesToClean(i))) Then vCurrStr = Replace(vCurrStr, Chr(vCodesToClean(i)), "")
  Next
    getClearInt = Trim(vCurrStr)
End Function

Function getClearString(ByVal vCurrStr As String)
  Dim i As Long
   
  vCurrStr = getClearInt(vCurrStr)
  For i = 128 To 255
    If InStr(vCurrStr, Chr(i)) Then vCurrStr = Replace(vCurrStr, Chr(i), "")
  Next
    vCurrStr = Application.WorksheetFunction.Clean(vCurrStr)
    getClearString = Trim(vCurrStr)
End Function


Function getNamedRange(vNewSheetName As String, vCurrCellAddress As String)
 
 Dim myNamedRange
  Call p_deleteNamedRange(vCurrNameRangeName & "")
  getNamedRange = getClearRangeName(vNewSheetName, vCurrCellAddress)
  ActiveWorkbook.Names.Add Name:=getClearString(getNamedRange), RefersTo:=Worksheets(vNewSheetName).Range(vCurrCellAddress)
 
End Function

Function getHsClearFormula(vCurStr As String)
         vCurStr = UCase(vCurStr)
         vCurStr = getClearString(vCurStr)
         vCurStr = Replace(vCurStr, """", "")
         'vCurStr = Replace(vCurStr, "&", "")
         getHsClearFormula = Replace(vCurStr, "$", "")
End Function

Function getClearRangeName(vCurrSheetName As String, vCurrCellAddress As String)
   getClearRangeName = getHsClearFormula("l_" & Worksheets(vCurrSheetName).Index & vCurrCellAddress)
   
End Function

