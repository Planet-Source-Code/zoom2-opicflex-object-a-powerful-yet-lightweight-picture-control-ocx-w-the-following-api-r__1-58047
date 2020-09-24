VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiErrorStringVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'Specifies that the function should search the system message-table resource(s) for the requested message
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
   
   
Function ApiErrorText(ByVal ErrNum As Long) As String
Dim msg As String
Dim nRet As Long

   msg = Space$(1024)
   nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, msg, Len(msg), ByVal 0&)
   If nRet Then
      ApiErrorText = Left$(msg, nRet)
   Else
      ApiErrorText = "Error (" & ErrNum & ") not defined."
   End If
End Function

