VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'=======================================================================
'Class:                 CFileErrorHandler
'
'Synopsis:              Handles file access errors.
'
'                       'Handle' means to present the error to the
'                       user and ask for a response.  How the user
'                       responded (Abort, Retry, Ignore, etc.) is
'                       returned from the HandleError method.  It
'                       is the responsibility of the calling
'                       procedure to act appropriately on the response.
'=======================================================================
Option Explicit

Private Const ErrBadFileNameOrNumber As Integer = 52
Private Const ErrFileNotFound As Integer = 53
Private Const ErrBadFileMode As Integer = 54
Private Const ErrFileAlreadyOpen As Integer = 55
Private Const ErrDeviceIO As Integer = 57
Private Const ErrFileAlreadyExists As Integer = 58
Private Const ErrBadRecordLength As Integer = 59
Private Const ErrDiskFull As Integer = 61
Private Const ErrInputPastEndOfFile As Integer = 62
Private Const ErrBadRecordNumber As Integer = 63
Private Const ErrBadFileName As Integer = 64
Private Const ErrTooManyFiles As Integer = 67
Private Const ErrDeviceUnavailable As Integer = 68
Private Const ErrPermissionDenied As Integer = 70
Private Const ErrDiskNotReady As Integer = 71
Private Const ErrPathFileAccess As Integer = 75
Private Const ErrPathDoesNotExist As Integer = 76
Private Const ErrInvalidFileFormat As Integer = 321 'new

'Begin of the Code Change : CPTPK-2513
Private m_oInteractionProvider As IInteractionProvider

Public Sub SetInteractionProvider(ByRef oInteractionProvider As IInteractionProvider)
    Set m_oInteractionProvider = oInteractionProvider
End Sub
'End   of the Code Change : CPTPK-2513

Public Function HandleError(ByVal sTitle As String, _
                  ByVal lErrNumber As Long, ByVal sErrDescription As String, ByVal sSource As String, _
                  ByVal sHelpFile As String, ByVal sHelpContext As String, _
                  Optional ByVal sErrorIntroduction As String) As Long
   Dim nMsgType As Integer
   Dim sMsg As String
   Dim nResponse As Integer

   HandleError = ErrorDispNotHandled


   nMsgType = vbExclamation + vbOKOnly

   Select Case lErrNumber
      Case g_icErrApplicationCanceled
         HandleError = ErrorDispAbortRaiseError
         Exit Function
      Case ErrBadFileNameOrNumber
         sMsg = "Illegal filename or handle."
         nMsgType = vbCritical + vbOKOnly
      Case ErrFileNotFound
         sMsg = "File doesn't exist."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrBadFileMode
         sMsg = "Can't open your file for that type of access."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrFileAlreadyOpen
         sMsg = "This file is already open."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrDeviceIO
         sMsg = "Internal disk error."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrFileAlreadyExists
         sMsg = "File already exists."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrBadRecordLength
         sMsg = "Bad record length."
         nMsgType = vbExclamation + vbAbortRetryIgnore
      Case ErrDiskFull
         sMsg = "Disk is full."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrInputPastEndOfFile
         sMsg = "This file has a nonstandard end-of-file marker, "
         sMsg = sMsg & "or an attempt was made to read beyond "
         sMsg = sMsg & "the end-of-file marker."
         nMsgType = vbExclamation + vbAbortRetryIgnore
      Case ErrBadRecordNumber
         sMsg = "Bad record number."
         nMsgType = vbExclamation + vbAbortRetryIgnore
      Case ErrBadFileName
         sMsg = "Illegal filename."
         nMsgType = vbCritical + vbOKOnly
      Case ErrTooManyFiles
         sMsg = "Too many files open at one time."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrDeviceUnavailable
         sMsg = "Device is unavailable."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrPermissionDenied
         sMsg = "Can't open file.  Permission denied."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrDiskNotReady
         sMsg = "Insert a disk in the drive and close the door."
      Case ErrPathFileAccess
         sMsg = "Path/File access error.  Path may not be formatted correctly or you may be attempting to write to a read-only file."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrPathDoesNotExist
         sMsg = "That path doesn't exist."
         nMsgType = vbExclamation + vbRetryCancel
      Case ErrInvalidFileFormat
         sMsg = "Invalid file format."
         nMsgType = vbExclamation + vbAbortRetryIgnore
      Case Else
         HandleError = ErrorDispNotHandled

         Exit Function
   End Select
   
   If Len(Trim$(sErrorIntroduction)) Then
      sMsg = sErrorIntroduction & vbCrLf & vbCrLf & sMsg
   End If
   
    'Begin of the Code Change : CPTPK-2513
    If Not m_oInteractionProvider Is Nothing Then
        nResponse = m_oInteractionProvider.MessageBox(sMsg, nMsgType, sTitle)
    Else
       nResponse = MsgBox(sMsg, nMsgType, sTitle)
    End If
   'End   of the Code Change : CPTPK-2513
   
   Select Case nResponse
      Case vbRetry
         HandleError = ErrorDispRetryResume
      Case vbIgnore
         HandleError = ErrorDispIgnoreResumeNext
      Case vbOK, vbCancel, vbAbort
         HandleError = ErrorDispAbortRaiseError
      Case Else
         HandleError = ErrorDispNotHandled
   End Select

End Function
