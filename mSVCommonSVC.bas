Attribute VB_Name = "mSVCommonSVC"
Option Explicit

Public vConnName As String
Public vServerName As String
Public vAPSServerName As String
Public vAppName As String
Public vDbName As String
Public vUserName As String
Public vPassword As String
Public vFriendlyName As String

Public vConnName_stored As String
Public vServerName_stored  As String
Public vAPSServerName_stored As String
Public vAppName_stored As String
Public vDbName_stored As String
Public vUserName_stored As String
Public vPassword_stored As String
Public vFriendlyName_stored As String
Public vArrFormulas() As Variant

Public vCurrConnectQ As String
  
  
 

 
 
 


Function getErrorText(vErrNum As Long) As String
 
    getErrorText = vbNewLine & "SmartView error num is " & vErrNum & " : " & GetReturnCodeMessage(vErrNum) & vbNewLine
    
    If vErrNum = -4 Then
      getErrorText = getErrorText & vbNewLine & vbNewLine & "Check connection credentionals information" & vbNewLine & vbNewLine
    End If
    
    If vErrNum = 41 Then
      getErrorText = getErrorText & vbNewLine & vbNewLine & "Check members in the retive slice and database connection "
    End If
    
End Function
 

 

