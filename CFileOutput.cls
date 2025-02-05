VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CFileErrorHandler"
'=======================================================================
'Class:                 CFileOutput
'
'Synopsis:              This class presents an interface through
'                       which a crlf delimeted file can be written.
'=======================================================================
Option Explicit

Private m_bIsOpen As Boolean
Private m_sFileName As String
Private m_nFileHandle As Integer

Private m_bTrimTrailingBlanks As Boolean
Private m_lExcludeFromTrim As Long 'If trimming trailing chars, exclude from trimming the first n chars.

'Preserved error data
Private m_lErr As Long
Private m_sErrSrc As String
Private m_sErrDesc As String
Private m_sHelpContext As String
Private m_sHelpFile As String

'Error handlers
Private m_oFileErrorHandler As CFileErrorHandler
Private m_oErrorHandler As CGeneralErrorHandler

Private Sub Class_Initialize()
   Set m_oFileErrorHandler = New CFileErrorHandler
   Set m_oErrorHandler = New CGeneralErrorHandler
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   If m_bIsOpen Then
      Close #m_nFileHandle
   End If
   Set m_oFileErrorHandler = Nothing
   Set m_oErrorHandler = Nothing
End Sub

Public Property Get ExcludeFromTrim() As Long
   ExcludeFromTrim = m_lExcludeFromTrim
End Property

Public Property Let ExcludeFromTrim(ByVal lNewValue As Long)
   m_lExcludeFromTrim = lNewValue
End Property

Public Property Get TrimTrailingBlanks() As Boolean
   TrimTrailingBlanks = m_bTrimTrailingBlanks
End Property

Public Property Let TrimTrailingBlanks(ByVal bNewValue As Boolean)
   m_bTrimTrailingBlanks = bNewValue
End Property

Public Property Get FileHandle() As Integer
   FileHandle = m_nFileHandle
End Property

Public Function OpenFile(ByVal sFileName As String) As Boolean
   
   On Error GoTo EHandler
   
   m_sFileName = sFileName
   m_nFileHandle = FreeFile
   Open m_sFileName For Output As #m_nFileHandle
   
   m_bIsOpen = True
   OpenFile = True
   
   Exit Function

EHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.Source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in Method CFileOutput.OpenFile."
  
   'Handle error
   Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
      Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
         Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
            Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
               Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
            Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
               Resume
            Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
               Resume Next
            Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
               GoSub Rollback
               'Signal failure
               OpenFile = False
               Exit Function
         End Select
      Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
         Resume
      Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
         Resume Next
      Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
         GoSub Rollback
         'Signal failure
         OpenFile = False
         Exit Function
   End Select
   
Rollback:
   
   m_sFileName = vbNullString
   m_nFileHandle = 0
   m_bIsOpen = False
   
   Return
   
End Function

Public Sub CloseFile()
   On Error Resume Next
   If m_bIsOpen Then
      Close #m_nFileHandle
   End If
   m_sFileName = vbNullString
   m_nFileHandle = 0
   m_bIsOpen = False
End Sub

Public Sub WriteLine(ByVal sRecord As String)
   Dim nPos As Integer
   
   On Error GoTo EHandler
   
   If m_bTrimTrailingBlanks Then
      sRecord = RTrim$(sRecord)
      If Len(sRecord) < m_lExcludeFromTrim Then
         sRecord = sRecord & String(m_lExcludeFromTrim - Len(sRecord), " ")
      End If
   End If
   
   Print #m_nFileHandle, sRecord
   
   Exit Sub

EHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.Source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in Method CFileOutput.WriteLine."
  
   'Handle error
   Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
      Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
         Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
            Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
               Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
            Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
               Resume
            Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
               Resume Next
            Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
               Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
               Exit Sub
         End Select
      Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
         Resume
      Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
         Resume Next
      Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
         Exit Sub
   End Select
   
End Sub

Public Property Let Position(ByVal lPosition As Long)
   On Error GoTo EHandler
   Seek m_nFileHandle, lPosition
   Exit Property

EHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.Source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in PropertyLet CFileOutput.Position."
  
   'Handle error
   Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
      Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
         Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
            Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
               Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
            Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
               Resume
            Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
               Resume Next
            Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
               Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
               Exit Property
         End Select
      Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
         Resume
      Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
         Resume Next
      Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
         Exit Property
   End Select

End Property

Public Property Get Position() As Long
   
   On Error GoTo EHandler
   Position = Seek(m_nFileHandle)
   Exit Property

EHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.Source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in PropertyGet CFileOutput.Position."
  
   'Handle error
   Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
      Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
         Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
            Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
               Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
            Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
               Resume
            Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
               Resume Next
            Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
               Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
               Exit Property
         End Select
      Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
         Resume
      Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
         Resume Next
      Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
         Exit Property
   End Select

End Property

