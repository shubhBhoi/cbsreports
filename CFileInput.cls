VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CFileErrorHandler"
'=======================================================================
'Class:                 CFileInput
'
'Synopsis:              This class presents an interface through
'                       which a crlf delimeted file can be read.
'=======================================================================
Option Explicit

'This module requires that the constants g_icErrApplicationCanceled and
'g_scErrApplicationCanceled be declared in another public module.  This are
'used for raising an error if the user responds to a file error
'by requesting abort.

Private Const m_scClassName As String = "CFileInput"

Private m_bIsOpen As Boolean
Private m_sFileName As String
Private m_nFileHandle As Integer
Private m_lFileSize As Long
Private m_lFileSizeBytes As Long
Private m_sRecord As String
Private m_lRecordNumber As Long
Private m_lLineNumber As Long
Private m_bAtEnd As Boolean

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
   CloseFile
   Set m_oFileErrorHandler = Nothing
   Set m_oErrorHandler = Nothing
End Sub

Public Property Get FileHandle() As Integer
   FileHandle = m_nFileHandle
End Property

Public Property Get EndOfFile() As Boolean
   EndOfFile = EOF(m_nFileHandle)
End Property

'AtEnd differs from EndOfFile in that AtEnd is true only if we've attempted
'a read when EOF() is true whereas EndOfFile is true as soon as EOF() is true.
Public Property Get AtEnd() As Boolean
   AtEnd = m_bAtEnd
End Property

Public Property Get LineNumber() As Long
   LineNumber = m_lLineNumber
End Property

Public Property Get Name() As String
   Name = m_sFileName
End Property

Public Property Let Position(ByVal lPosition As Long)
   
   On Error GoTo EHandler
   
   If lPosition < 2 Then
      m_lRecordNumber = 0
      m_lLineNumber = 0
   End If
   
   Seek m_nFileHandle, lPosition
   m_bAtEnd = EndOfFile
   
   Exit Property
   
EHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in PropertyLet CFileInput.Postion."
  
   On Error GoTo 0
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
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in PropertyGet CFileInput.Postion."
  
   On Error GoTo 0
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

Public Property Get PosPercent() As Long
   Dim lFilePos As Long
   Dim cPercent As Currency
   lFilePos = Loc(m_nFileHandle)
   cPercent = lFilePos / m_lFileSize
   PosPercent = cPercent * 100
   If PosPercent > 100 Then
      PosPercent = 100
   End If
End Property

'Returns the record count.  Only valid if user
'is doing sequential reads.  Setting the position
'property will invalidate this property unless
'the Position is set to 1, in which case the
'RecordNumber will be reset to 0.
Public Property Get RecordNumber() As Long
   RecordNumber = m_lRecordNumber
End Property

Public Property Get FileSize() As Long
   FileSize = m_lFileSizeBytes
End Property

Public Property Get Record() As String
   Record = m_sRecord
End Property

Public Sub CloseFile()
   On Error Resume Next
   If m_bIsOpen Then
      Close #m_nFileHandle
      m_sFileName = vbNullString
      m_nFileHandle = 0
      m_lFileSizeBytes = 0
      m_lFileSize = 0
      m_sRecord = vbNullString
   End If
End Sub

Public Function OpenFile(ByVal vsFileName As String) As Boolean
   
   On Error GoTo ErrHandler
   
   m_sFileName = vsFileName
   m_nFileHandle = FreeFile
   
   Open m_sFileName For Input Access Read As #m_nFileHandle
   
   m_lFileSizeBytes = LOF(m_nFileHandle)
   m_lFileSize = LOF(m_nFileHandle) / 128
   
   m_bAtEnd = False
   
   OpenFile = True
   
   Exit Function
   
ErrHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in Metho CFileInput.OpenFile."
  
   On Error GoTo 0
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

Public Function ReadRecord() As String
   Dim nPos As Integer
   
   On Error GoTo ErrHandler
   
   m_sRecord = vbNullString
   ReadRecord = vbNullString
   
   If m_bAtEnd = False Then
      If EndOfFile Then
         m_bAtEnd = True
         Exit Function
      End If
   End If
   
   Line Input #m_nFileHandle, m_sRecord
   ReadRecord = m_sRecord

   On Error Resume Next
   m_lRecordNumber = m_lRecordNumber + 1
   
   Exit Function

ErrHandler:
  
   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpFile = Err.HelpFile
   m_sHelpContext = Err.HelpContext
   
   Const scErrorContext = "Error in Method CFileInput.ReadRecord."
  
   On Error GoTo 0
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
               Exit Function
         End Select
      Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
         Resume
      Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
         Resume Next
      Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
         Exit Function
   End Select
   
End Function
