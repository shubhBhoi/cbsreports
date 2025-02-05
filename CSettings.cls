VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oSettingsRootNode As MSXML2.IXMLDOMNode

'Preserved error data
Private m_lErr As Long
Private m_sErrSrc As String
Private m_sErrDesc As String
Private m_sHelpContext As String
Private m_sHelpFile As String

'Error handlers
Private m_oErrorHandler As CGeneralErrorHandler

Private Sub Class_Initialize()
   Set m_oErrorHandler = New CGeneralErrorHandler
End Sub

Private Sub Class_Terminate()
   Set m_oErrorHandler = Nothing
End Sub

'==============================================================================
'Function:              GetDirectory
'
'Synopsis:              Returns the specified directory.
'
'==============================================================================
Public Function GetDirectory(ByVal sID As String) As String
    GetDirectory = vbNullString
    XMLGetNodeText "//Directory[@ID = '" & sID & "']", GetDirectory
End Function

'==============================================================================
'Function:              XMLLoadSettings
'
'Synopsis:              Loads the SettingsRoot node into an XML string.
'
'==============================================================================
Public Function XMLLoadSettings(ByVal sXML As String) As Boolean
    Dim oDocument As MSXML2.DOMDocument60
    Dim oBatchRootNode As MSXML2.IXMLDOMNode
    Dim oPaymentsSummaryNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    Dim lResult As VbMsgBoxResult
    
    On Error GoTo XMLLoadSettings_Error
    
    Set m_oSettingsRootNode = Nothing
    
    Set oDocument = New MSXML2.DOMDocument60
    
    On Error Resume Next
    oDocument.async = False
    oDocument.validateOnParse = True
    bSuccess = oDocument.loadXML(sXML)
    On Error GoTo XMLLoadSettings_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML document." & vbCrLf & vbCrLf
        sMsg = sMsg & oDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDocument.parseError.Line & vbCrLf & _
            oDocument.parseError.srcText & vbCrLf & vbCrLf
        Err.Raise 5, App.Title, sMsg
        GoTo XMLLoadSettings_ErrorExit
    End If
    
    Set m_oSettingsRootNode = oDocument.selectSingleNode("SettingsRoot")
    If m_oSettingsRootNode Is Nothing Then
        GoTo XMLLoadSettings_ErrorExit
    End If
    
    XMLLoadSettings = True
    Exit Function
    
XMLLoadSettings_ErrorExit:
    
    XMLLoadSettings = False
    Set m_oSettingsRootNode = Nothing
    Exit Function

XMLLoadSettings_Error:

    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "Error loading the settings XML string.  (Proc=CSettings.XMLLoadSettings)"
    
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
            Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            Resume XMLLoadSettings_ErrorExit
    End Select
    
    Exit Function

End Function

Private Function XMLGetNodeText(ByVal sNodeName As String, ByRef sText As String) As Boolean
    Dim oNode As MSXML2.IXMLDOMNode
    sText = vbNullString
    Set oNode = m_oSettingsRootNode.selectSingleNode(sNodeName)
    If oNode Is Nothing Then
        XMLGetNodeText = False
        Exit Function
    End If
    sText = oNode.Text
    XMLGetNodeText = True
End Function





