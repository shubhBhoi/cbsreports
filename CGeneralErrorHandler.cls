VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGeneralErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'=======================================================================
'Class:                 CGeneralErrorHandler
'
'Synopsis:              Handles errors of a general nature (as opposed
'                       to a DAO error or a file access error.)
'
'                       'Handle' means to present the error to the
'                       user and ask for a response.  How the user
'                       responded (Abort, Retry, Ignore, etc.) is
'                       returned from the HandleError method.  It
'                       is the responsibility of the calling
'                       procedure to act appropriately on the response.
'=======================================================================
Option Explicit

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
   Dim nResponse
   
   If lErrNumber = g_icErrApplicationCanceled Then
      HandleError = ErrorDispAbortRaiseError
      Exit Function
   End If

   HandleError = ErrorDispNotHandled

   If sErrorIntroduction <> vbNullString Then
      sMsg = sErrorIntroduction & vbCrLf & vbCrLf
   End If
   sMsg = sMsg & "Error " & ErrorNumberString(lErrNumber) & _
                       IIf(Trim$(sSource) = vbNullString, vbNullString, " in " & sSource) & _
                       ": " & sErrDescription
   
   
   nMsgType = vbExclamation + vbAbortRetryIgnore
   If Len(Trim$(sHelpFile)) > 0 Then
      nMsgType = nMsgType + vbMsgBoxHelpButton
   End If
   
    'Begin of the Code Change : CPTPK-2513
     If Not m_oInteractionProvider Is Nothing Then
         nResponse = m_oInteractionProvider.MessageBox(sMsg, nMsgType, sTitle, sHelpFile, sHelpContext)
     Else
         nResponse = MsgBox(sMsg, nMsgType, sTitle, sHelpFile, sHelpContext)
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
