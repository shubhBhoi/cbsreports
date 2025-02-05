VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const FiservMailRateCategorySelect As String = "%FiservMailRateCategorySelect"
Private Const FiservMailRateCategoryCase As String = "%FiservMailRateCategoryCase:"
Private Const FiservMailRateCategoryDefault As String = "%FiservMailRateCategoryDefault"
Private Const FiservMailRateCategoryEndSelect As String = "%FiservMailRateCategoryEndSelect"
Private Const FiservMailRateCategoryConcatenator As String = "%%+"

Private m_oJob As IJob2
Private m_sTag As String

Implements IRunJob
Implements ISettings
Implements IInteractive

' Begin DDA 2.0 upgrade
Public Enum EddPlatformType
    Legacy = 0
    DDA_20 = 1
End Enum
Private m_lXML_EddPlatformType As EddPlatformType
' End DDA 2.0 upgrade
    
Private Const m_scFileName As String = "FILE NAME-"

Private Const m_scPaperTypePressureSeal As String = "PSGN1"
Private Const m_scPaperTypePlain As String = "BND20"
Private Const m_scPaperTypeBottomPerf As String = "BottomPerf"

Private Const m_lcPressureSealRetAddrY As Long = 757 'Base line of highest return address line on pressure seal envelope.
Private Const m_lcPressureSealOuterLogoY As Long = 765 'Top of indicia on pressure seal envelope.

'Private Const m_scXmlAse As String = "RETURN SERVICE REQUESTED"
Private m_sXmlAse As String

Private Const m_scInst_29175100_FinFactors As String = "Finance Factors"

Private m_sXmlInst As String
Private m_sXml_FileGroup As String

Private m_lXmlAddressPositionX As Long ' Left edge of flat-stock address blocks in points.
Private m_lXmlAddressPositionY As Long ' Lower edge of flat-stock address blocks in points.

Private m_sXmlReportId As String ' This is used only if there is no file header in the client print stream.
Private m_sXmlFormType As String ' This is used only if there is no file header in the client print stream.

Private m_sXmlDocument As String

Private Const scDocument_TM0330P1 As String = "TM0330P1"

Private m_bXML_Archive As Boolean
Private m_bXML_EddEstatements As Boolean

Private Const m_scEddJobNamePrefix As String = "EDD:"
Private m_sXML_EddJobKey As String
Private m_sXML_EddProductId As String
Private m_sXML_EddDocType As String
Private m_sXml_EddClientId As String
Private m_bXML_EddCompanionFile As Boolean
Private m_sEddCycleDate As Date

Private m_sXML_BulkShipList As String
Private m_sBulkShipList As String
Private m_saBulkShipAccounts() As String
Private m_lBulkShipAccountsIDX As Long
Private m_bXML_ReturnToBank As Boolean

Private m_sXML_AccountType As String

Private m_bXML_ArchiveNameInFile As Boolean

Private m_oInteractionProvider As IInteractionProvider 'CPTPK-2513

' This is used for new Signature eNotices
' This will be created client by client
' These are like eStatements - but different
Private m_sXML_EddEnoticeJobKey As String
Private m_sXML_EddEnoticeClientId As String

Private Enum EnoticeType
    EnoticeTypeUnknown = 0
    EnoticeType10and20
    EnoticeType30
    EnoticeType50
End Enum

' Forms that are communicated via the Settings XML tag - these have to be manually referenced in the code
Private m_oXml_InnerLogo As PostscriptImaging.FormInfo
Private m_oXml_OuterLogo As PostscriptImaging.FormInfo
Private m_oXml_FlatStockLogo As PostscriptImaging.FormInfo
Private m_oXml_Backer As PostscriptImaging.FormInfo

Private m_lEddEnoticeType As EnoticeType
Private m_sPieceEddEnoticeType As String

Private m_oDneFile As CFileOutput
Private m_sDneDate As String
Private m_sDneBaseFileName As String
Private m_sDneFirstAccount As String
Private m_sDneLastAccount As String
Private m_lDneStatementCount As Long
Private m_lDneImageCount As Long

Private m_sOMR As String

Private m_oSettingsSink As ISettingsSink
Private m_oWizardSettings As CSettings

Private m_sDocumentFolder As String
Private m_sFontsFolder As String
Private m_sSettingsFolder As String
Private m_sOutputFolder As String

Private m_oXML_BankAddress As CAddress

Private m_oBPPExteriorLogo(1) As PostscriptImaging.FormInfo
Private m_oBPBInteriorLogo(1) As PostscriptImaging.FormInfo
Private m_oBPBFlatStockLogo As PostscriptImaging.FormInfo
Private m_oBPBBackers(1) As PostscriptImaging.FormInfo
Private m_oBPP102Form As PostscriptImaging.FormInfo

' HigherOne logos
Private m_oHigherOneExteriorLogo(1) As PostscriptImaging.FormInfo
Private m_oHigherOneInteriorLogo(1) As PostscriptImaging.FormInfo

' Sterling national bank logos and backers
Private m_oSNBExteriorLogo(2) As PostscriptImaging.FormInfo
Private m_oSNBInteriorLogo(2) As PostscriptImaging.FormInfo
Private m_oSNBFlatStockLogo(2) As PostscriptImaging.FormInfo
Private m_oSNBBackers(2) As PostscriptImaging.FormInfo

' DEPRECATED
' The below variables have been deprecated in favor of settings XML elements
' -------------------------------------------------------------------
' Nationwide bank logos and backers
' Index 0 is the lending/original logo
' Index 1 is the deposits logo
'Private m_oNationwideExteriorLogo(1) As PostscriptImaging.FormInfo
'Private m_oNationwideInteriorLogo(1) As PostscriptImaging.FormInfo
'Private m_oNationwideFlatStockLogo(1) As PostscriptImaging.FormInfo
'Private m_oNationwideBackers(1) As PostscriptImaging.FormInfo
' -------------------------------------------------------------------

' El Dorado bank logos and backers
Private m_oElDoradoExteriorLogo As PostscriptImaging.FormInfo
Private m_oElDoradoInteriorLogo As PostscriptImaging.FormInfo
Private m_oElDoradoFlatStockLogo As PostscriptImaging.FormInfo
Private m_oElDoradoBackers(0) As PostscriptImaging.FormInfo

' Investors bank logos and backers
Private m_oInvestorsExteriorLogo As PostscriptImaging.FormInfo
Private m_oInvestorsInteriorLogo As PostscriptImaging.FormInfo
Private m_oInvestorsFlatStockLogo As PostscriptImaging.FormInfo
Private m_oInvestorsBackers(1) As PostscriptImaging.FormInfo

Private m_oRABacker As PostscriptImaging.FormInfo
Private m_oBackerInfo As PostscriptImaging.FormInfo

Private Const m_scIndiciaFile As String = "FDS\Indicia_PresortedFirstClassFDS_FormCache.xml"
Private m_oIndicia As PostscriptImaging.FormInfo

Private Enum StackItemsEnum
    StackItemsUndefined = 0
    StackItemsBypass
    StackItemsPressureSeal
    StackItemsNotice
    StackItemsFlatStock
    StackItemsFlatStock8LPI
    StackItemsFlatStock9LPI
    StackItemsFlatStockLegal
    StackItemsFlatStock42
    StackItemsFlatStock428LPI
    StackItemsLN0473
    StackItemsLN0473_152
End Enum
Private m_lStackItems As StackItemsEnum

Private Const m_scStockTrayPlainWhite20 = "BND20"
Private Const m_scStockTrayPressureSeal = "PSGEN"
Private Const m_scStockTrayPressureSeal2 = "PSGN1"
Private Const m_scEnvelope42 = "42"

'Private Const m_scStockTrayBPB101 = "BPB-101" 'BPB Letterhead w/ top perf - no backer. 'REM RM16020
Private Const m_scStockTrayBPB101 = "BPB-101-2" 'BPB Letterhead w/ top perf - no backer. 'RM16020
Private Const m_scStockTrayBPB601 = "BPB-601" 'Borel Letterhead w/ top perf - no backer.
Private Const m_scStockTrayBPBpex = "BPB-201-3"

'Private Const m_scStockTrayBPB201CreditLine = "BPB-201-1" 'BPB Letterhead w/ top perf - Credit Line Backer 'REM RM16020
Private Const m_scStockTrayBPB201CreditLine = "BPB-201-2" 'BPB Letterhead w/ top perf - Credit Line Backer 'RM16020
Private Const m_scStockTrayBPB701CreditLine = "BPB-701" 'Borel Letterhead w/ top perf - Credit Line Backer

Private Const m_scStockTrayBPBTopPerf = "BND20" 'BPB Letterhead w/ top perf (seems to be same as 101)
Private Const m_scStockTrayBorelTopPerf = "BND20" 'Borel Letterhead w/ top perf (seems to be same as 601)

'Private Const m_scStockTrayBPBLetterhead = "BPB-102" 'BPB Letterhead - no pref, no backer\ 'REM RM16020
Private Const m_scStockTrayBPBLetterhead = "BPB-102-2" 'BPB Letterhead - no pref, no backer 'RM16020
Private Const m_scStockTrayBorelLetterhead = "BPB-602" 'Borel Letterhead - no pref, no backer

'Private Const m_scStockTrayBPB103 = "BPB-103" 'BPB Legal Size Letterhead 'REM RM16020
Private Const m_scStockTrayBPB103 = "BPB-103-2" 'BPB Legal Size Letterhead 'RM16020
Private Const m_scStockTrayBPB603 = "BPB-603" 'Borel Legal Size Letterhead

'Private Const m_scEnclosureBPBEL9 As String = "BPBEL9" 'REM RM16020
Private Const m_scEnclosureBPBEL9 As String = "BPBEL9-3" 'RM16020
Private Const m_scEnclosureBPBER9 As String = "BPBER9"
Private Const m_scEnclosureBPBBORELER9 As String = "BOREL-ER9"

Private Const m_scEnvelopeBorel As String = "BOREL-E10"
Private Const m_scEnvelopeBPB As String = "BPB-E10"

' Woori America Bank - T152Q2 BRE
Private Const m_scWooriT152Q2BRE As String = "ENVLN1-WAB"

' Nationwide Bank T152Q2 BRE
Private Const m_scNationwideBRE As String = "#9-BRE"

' Investors Custom Stock codes
Private Const m_scInvestorsINVSTM = "INV-STM"
Private Const m_scInvestorsLEGAL = "AA-5736"
Private Const m_scInvestorsBRE = "#9-BRE"
Private Const m_scInvestorsBREMORE = "#9-BRE + More"

Private m_sInputFiles As String
Private m_sInputFile() As String
Private m_lInputFileIdx As Long
Private m_lInputFilePtr As Long
' Controls whether virtual files are supported in the server
Const m_bcVirtualFileSupport As Boolean = False

Private m_InputFileFileRecordCount As Long

Private m_oPrintStream As FDS_IO_PrintStream.IPrintStreamReader

Private m_sStmtBodyFileName As String
Private m_oStmtBodyFile As CFileOutput

Private m_sStmtBodyDatabaseName As String
Private m_oStmtBodyDatabase As FfgSbfDb.CDbStmtBodyFile

Private m_oCassSink As ICassSink

Private m_lLocation As Long 'Postion in SBF

Private m_lOutputSheets As Long
Private m_lOutputImages As Long
Private m_lInputImages As Long

Private m_lInputSequence As Long
Private m_lOutputImageSeq As Long

'--------------------------------------------------
'Nautilus indexing fields
Private m_saIndexValues() As String
Private m_saIndexKeys() As String
Private m_lIndexIdx As Long

' RM15930
'SecureDestruction indexing fields
Private m_saSecureDestructionIndexValues() As String
Private m_saSecureDestructionIndexKeys() As String
Private m_lSecureDestructionIndexIdx As Long
 ' End RM15930

Private Const m_scAccountNumberKey As String = "Account_Number"
Private Const m_scStatementDateKey As String = "Statement_Date"
Private Const m_scDocumentTypeKey As String = "Document_Type"
Private Const m_scPrimaryNameKey As String = "Primary_Name"
Private Const m_scFileTypeKey As String = "File_Type"
'--------------------------------------------------

Private m_oReturnAddress As CAddress
Private m_oFlatStockLogoFormInfo As PostscriptImaging.FormInfo
Private m_oExteriorLogoFormInfo As PostscriptImaging.FormInfo
Private m_oInteriorLogoFormInfo As PostscriptImaging.FormInfo
Private m_oPieceBacker As PostscriptImaging.FormInfo
'Private m_oContactInfoFormInfo As PostscriptImaging.FormInfo
Private m_oDeliveryAddress As CAddress
Private m_oCassAddress As CAddress
Private m_sAccountNumber As String
Private m_sDocumentDate As String
Private m_sRegion As String
Private m_bTrailerPage As Boolean
Private m_bNoPrint As Boolean
Private m_bDuplex As Boolean
Private m_bReturnToBank As Boolean
Private m_sStockTray As String
Private m_sEnvelopeCatCode As String
Private m_sEnclosureCatCode As String
Private m_cPaymentAmount As Currency
Private m_sOCR As String
Private m_sMICR As String
Private m_cAmountDue As Currency

'ToDo: paramaterize coordinates.
Private Const m_sngcOCR_X As Single = 54
Private Const m_sngcOCR_Y As Single = 708
Private m_sngOCR_X As Single
Private m_sngOCR_Y As Single

Private m_sLetterDate As String

Private m_saBulkShipFileName() As String
Private m_saBulkShipFormType() As String
Private m_lBulkShipFileIdx As Long

Private m_saUnrecognizedFileName() As String
Private m_saUnrecognizedFormType() As String
Private m_lUnrecognizedFileIdx As Long

Private m_lNextNoPrintStack As Long
Private m_lCurrentNoPrintStack As Long
Private m_lCurrentSampleStack As Long
Private m_lCurrentBulkShipStack As Long

Private m_bSamples As Boolean
Private m_saNoPrintStack() As String
Private m_lNoPrintStackIdx As Long

Private m_lCurrentFileNumber As Long
Private m_sCurrentFileName As String
Private m_sCurrentFileFormType As String
Private m_lCurrentFileExpectedPages As Long
Private m_bCurrentFileIsUnrecognized As Boolean

Private m_lCurrentFilePieces As Long
Private m_lCurrentFileImages As Long

'Current page buffer.
Private m_saCurrentPageLine() As String
Private m_saCurrentPageLineBold() As String
Private m_lCurrentPageIdx As Long

'Read ahead page buffer.
Private m_saReadAheadPageLine() As String
Private m_saReadAheadPageLineBold() As String
Private m_lReadAheadPageIdx As Long

Private m_bPrintStreamPrimed As Boolean
Private m_bPrintStreamAtEnd As Boolean

'Top panel
Private Const m_lcSelfSealFormPanel_x As Long = 48
Private Const m_lcSelfSealFormPanel_y As Long = 273
Private Const m_lcSelfSealFormPanel_w As Long = 516
Private Const m_lcSelfSealFormPanel_h As Long = 192

'Bottom panel
Private Const m_lcSelfSealInstructionPanel_x As Long = 48
Private Const m_lcSelfSealInstructionPanel_y As Long = 48
Private Const m_lcSelfSealInstructionPanel_w As Long = 516
Private Const m_lcSelfSealInstructionPanel_h As Long = 207

'Both panels combined
Private Const m_lcSelfSealPanels_x As Long = 48
Private Const m_lcSelfSealPanels_y As Long = 48
Private Const m_lcSelfSealPanels_w As Long = 516
Private Const m_lcSelfSealPanels_h As Long = 432

'Both panels on reverse side
Private Const m_lcSelfSealBackPanels_x As Long = 48
Private Const m_lcSelfSealBackPanels_y As Long = 312
Private Const m_lcSelfSealBackPanels_w As Long = 516
Private Const m_lcSelfSealBackPanels_h As Long = 432

Private Const m_lcSize001_BottomMargin As Long = -9
Private Const m_lcSize001_DTL As Long = 9

Private m_oAddressUtility As FfgStdAddress.CAddressUtility

Private m_oRegExp As RegExp

'Preserved error data
Private m_lErr As Long
Private m_sErrSrc As String
Private m_sErrDesc As String
Private m_sHelpContext As String
Private m_sHelpFile As String

'Error handlers
Private m_oFileErrorHandler As CFileErrorHandler
Private m_oErrorHandler As CGeneralErrorHandler

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Private Sub Class_Initialize()
    Set m_oFileErrorHandler = New CFileErrorHandler
    Set m_oErrorHandler = New CGeneralErrorHandler
    Set m_oAddressUtility = New CAddressUtility
    m_lInputFileIdx = -1
    m_lInputFilePtr = -1
    Set m_oRegExp = New RegExp
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   Set m_oFileErrorHandler = Nothing
   Set m_oErrorHandler = Nothing
    Set m_oAddressUtility = Nothing
End Sub

'Begin of the Code Change : CPTPK-2513
Private Sub IInteractive_SetInteractionProvider(oInteractionProvider As FFGPrintOlb.IInteractionProvider)
    Set m_oInteractionProvider = oInteractionProvider
End Sub
'End   of the Code Change : CPTPK-2513

Private Property Get IRunJob_SbfName() As String
   IRunJob_SbfName = m_sStmtBodyFileName
End Property

Private Property Let IRunJob_SbfName(ByVal sNewValue As String)
   m_sStmtBodyFileName = sNewValue
End Property

Private Property Get IRunJob_DatabaseName() As String
   IRunJob_DatabaseName = m_sStmtBodyDatabaseName
End Property

Private Property Let IRunJob_DatabaseName(ByVal sNewValue As String)
   m_sStmtBodyDatabaseName = sNewValue
End Property

Private Property Get IRunJob_InputFiles() As String
   IRunJob_InputFiles = m_sInputFiles
End Property

Private Property Let IRunJob_InputFiles(ByVal sNewValue As String)
   m_sInputFiles = sNewValue
End Property

Private Property Get IRunJob_MultipleInputFiles() As Boolean
   IRunJob_MultipleInputFiles = True
End Property

Private Property Get IRunJob_CassType() As JobCassTypeEnum
'   IRunJob_CassType = jctPerPiece
    IRunJob_CassType = jctNoCass
End Property

Private Property Get IRunJob_CassSink() As ICassSink
'    Set IRunJob_CassSink = m_oCassSink
    Set IRunJob_CassSink = Nothing
End Property

Private Property Set IRunJob_CassSink(ByRef oCassSink As ICassSink)
'    Set m_oCassSink = oCassSink
End Property

Private Function IRunJob_RunJob(ByRef oJob As IJob, oJobStatusSink As IJobStatusSink) As Boolean
    
    Dim lPieceCount As Long
    Dim bCanceled As Boolean
    Dim oStack As CStack
    Dim oPiece As CPiece
    Dim lTotalInputPieces As Long
    Dim lTotalInputImages As Long
    Dim lTotalOutputImages As Long
    Dim bMorePages As Boolean
    Dim bAnotherFile As Boolean
    Dim lCutColumns As Long
    Dim i As Long
    Dim sXMLSettings As String
    Dim oJobStatusSink2 As IJobStatusSink2
    
    Dim bCass As Boolean
    Dim bCertified As Boolean
    Dim sDPBC As String
    Dim bBulkShip As Boolean
    Dim bNoPrint As Boolean
    Dim lForeignGroup As ForeignGroupEnum
    Dim bIndiciaRequired As Boolean
    
    'Begin of the Code Change : CPTPK-2513
     m_oErrorHandler.SetInteractionProvider m_oInteractionProvider
     m_oFileErrorHandler.SetInteractionProvider m_oInteractionProvider
    'End   of the Code Change : CPTPK-2513
    
    On Error GoTo RunJob_Error
    
    m_lBulkShipAccountsIDX = -1
    m_bXML_ArchiveNameInFile = False
    
    Set oJobStatusSink2 = oJobStatusSink
    Set m_oJob = oJob
    m_sTag = Trim(oJob.Tag)
    
    m_bSamples = False
    Erase m_saNoPrintStack
    m_lNoPrintStackIdx = -1
    
    Erase m_saBulkShipFileName
    Erase m_saBulkShipFormType
    m_lBulkShipFileIdx = -1
    
    Erase m_saUnrecognizedFileName
    Erase m_saUnrecognizedFormType
    m_lUnrecognizedFileIdx = -1
    
    If m_oSettingsSink Is Nothing Then
        MessageBox "Cannot obtain the fonts database location.", vbCritical + vbOKOnly, App.Title
        GoTo RunJob_ErrorExit
    End If
    m_oSettingsSink.Settings sXMLSettings
    Set m_oWizardSettings = New CSettings
    If m_oWizardSettings.XMLLoadSettings(sXMLSettings) = False Then
        MessageBox "Cannot obtain the settings database location.", vbCritical + vbOKOnly, App.Title
        GoTo RunJob_ErrorExit
    End If
    m_sDocumentFolder = m_oWizardSettings.GetDirectory("Document")
    m_sFontsFolder = m_oWizardSettings.GetDirectory("Font")
    m_sSettingsFolder = m_oWizardSettings.GetDirectory("Settings")
    m_sOutputFolder = m_oWizardSettings.GetDirectory("Output")
    
    If ParseXmlInst(m_sTag, m_sXmlInst) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXmlBulkShipList(m_sTag, m_sXML_BulkShipList) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXmlFileGroup(m_sTag, m_sXml_FileGroup) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLAse(m_sTag, m_sXmlAse) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLDocument(m_sTag, m_sXmlDocument) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLArchive(m_sTag, m_bXML_Archive, m_bXML_ArchiveNameInFile) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLFileHeader(m_sTag, m_sXmlReportId, m_sXmlFormType) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLEddSettings(m_sTag, m_sXML_EddJobKey, m_sXML_EddProductId, m_sXML_EddDocType, m_sXml_EddClientId, m_bXML_EddCompanionFile) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If ParseXMLAccountType(m_sTag, m_sXML_AccountType) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If Len(m_sXML_EddJobKey) = 0 Then

        If ParseXMLEdd(m_sTag, m_bXML_EddEstatements) = False Then
            MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
            GoTo RunJob_ErrorExit
        End If
    End If
    
    If Len(m_sXML_EddJobKey) > 0 Then
            
        If ParseXMLEddPlatform(m_sTag, m_lXML_EddPlatformType) = False Then
            MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
            GoTo RunJob_ErrorExit
        End If
    End If
    
    
    If ParseXMLEddEnoticeSettings(m_sTag, m_sXML_EddEnoticeJobKey, m_sXML_EddEnoticeClientId) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    ' Verify that a job cannot be configured for eStatement AND eNotice AND Archive - or some combination of the 3.  Must be 1 and only 1
    Dim bIsDigital As Boolean
    bIsDigital = False
    ' Assign the archive
    bIsDigital = m_bXML_Archive
    ' Assign the eStatements
    If Len(Trim(m_sXML_EddJobKey)) > 0 Then
        If bIsDigital = True Then
            MessageBox "Cannot have edd, enotice, and/or archive configured all as one run.  Must be only 1 config.", vbOKOnly + vbCritical, App.Title
            GoTo RunJob_ErrorExit
        End If
        bIsDigital = True
    End If
    ' Assign the eNotices
    If Len(Trim(m_sXML_EddEnoticeJobKey)) > 0 Then
        If bIsDigital = True Then
            MessageBox "Cannot have edd, enotice, and/or archive configured all as one run.  Must be only 1 config.", vbOKOnly + vbCritical, App.Title
            GoTo RunJob_ErrorExit
        End If
        bIsDigital = True
    End If
    
    If ParseXMLAddressPosition(m_sTag, m_lXmlAddressPositionX, m_lXmlAddressPositionY) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
    
'RM12308
    If ParseXmlReturnToBank(m_sTag, m_bXML_ReturnToBank) = False Then
        MessageBox "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid.", vbOKOnly + vbCritical, App.Title
        GoTo RunJob_ErrorExit
    End If
'END RM12308
    
    m_lInputFileIdx = SplitMultiSelect(m_sInputFiles, m_sInputFile())
    If IncrementInputFile = False Then
        MessageBox "Cannot open input file.", vbExclamation + vbOKOnly, App.Title
        GoTo RunJob_ErrorExit
        Exit Function
    End If
    
    If Len(m_sXML_BulkShipList) > 0 Then
        m_sBulkShipList = AddBackslash(GetFolder(m_sInputFile(0))) & m_sXML_BulkShipList
        Dim cfm As CFileInput
        Set cfm = New CFileInput
        If cfm.OpenFile(m_sBulkShipList) = False Then
            MessageBox "Unable to open Bulk ship list file " & m_sBulkShipList & ".", vbOKOnly + vbCritical, App.Title
            Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
        End If
        
        Dim sRecord As String
        m_lBulkShipAccountsIDX = -1
        Do While Not cfm.EndOfFile
            sRecord = cfm.ReadRecord
            m_lBulkShipAccountsIDX = m_lBulkShipAccountsIDX + 1
            ReDim Preserve m_saBulkShipAccounts(m_lBulkShipAccountsIDX)
            m_saBulkShipAccounts(m_lBulkShipAccountsIDX) = Left(sRecord, 13) ' currently only setup for Finance Factors
        Loop
        cfm.CloseFile
        Set cfm = Nothing
        
    End If
    
    If OpenSbfDatabase() = False Then
        MessageBox "Cannot open the statement body database.", vbExclamation + vbOKOnly, App.Title
        GoTo RunJob_ErrorExit
    End If
    
    If Not m_bSamples Then
        If m_bXML_Archive Then
            SetupStacks_Archive
        ElseIf Len(m_sXML_EddJobKey) > 0 Then
            SetupStacks_Edd
        ElseIf Len(m_sXML_EddEnoticeJobKey) > 0 Then
            SetupStacks_EddEnotice
        ElseIf m_bXML_EddEstatements Then 'Rivercity - depracated
            SetupStacks_EddEstatements
        Else
            SetupStacks
        End If
    End If
    
    If OpenSbfFile() = False Then
       MessageBox "Cannot open the statement body file.", vbExclamation + vbOKOnly, App.Title
       GoTo RunJob_ErrorExit
    End If
    
    PsWriteHeader
    PsWriteProlog_Begin
    PsWriteProlog_FlatStock
    PsWriteProlog_Notice
    PsWriteProlog_End
    PsWriteSetup
    
    Do
        
        If OpenPrintStream() = False Then
            MessageBox "Cannot open input file.", vbExclamation + vbOKOnly, App.Title
            GoTo RunJob_ErrorExit
            Exit Function
        End If
        
        If Not oJobStatusSink2 Is Nothing Then
            oJobStatusSink.Status 0, jstPercentage, bCanceled
            oJobStatusSink2.StatusCaption "Processing " & GetFile(m_sInputFile(m_lInputFilePtr)) & " ..."
        End If
    
        '----------------------------
        'Initialize the client data buffer data.
        m_bPrintStreamPrimed = False
        m_bPrintStreamAtEnd = False
        
        Erase m_saReadAheadPageLine
        Erase m_saReadAheadPageLineBold
        m_lReadAheadPageIdx = -1
        'm_bReadAheadPageIsFileNameRecord = False
        
'        m_sReadAheadFileName = vbNullString
'        m_sReadAheadFileFormType = vbNullString
'        m_lReadAheadFileExpectedPages = 0
        
        'Note: File number not cleared, it carries forward to the next file.
        m_sCurrentFileName = vbNullString
        m_sCurrentFileFormType = vbNullString
        m_lCurrentFileExpectedPages = 0
        
        m_lCurrentFilePieces = 0
        m_lCurrentFileImages = 0
        m_lCurrentNoPrintStack = 0
        m_lCurrentSampleStack = 0
        m_lCurrentBulkShipStack = 0
        '----------------------------
        bMorePages = GetPage_Indirect()
        
        If bMorePages = False Then
            If m_bPrintStreamAtEnd = False Then
                MessageBox "Cannot locate the first data page in the input file.", vbExclamation + vbOKOnly, App.Title
                GoTo RunJob_ErrorExit
            End If
        End If
    
        'Process all statements
        Do While bMorePages
        
            ' If we are the ENotices - we may need to render the document twice
            ' We go through this part of the run job instead of the standard
            If Len(m_sXML_EddEnoticeJobKey) > 0 Then
            
                ' Re-sets the m_sPieceEddEnoticeType
                m_sPieceEddEnoticeType = vbNullString
            
                ' Verify that the ENoticeType is valid
                If m_lEddEnoticeType = EnoticeTypeUnknown Then
                    Err.Raise g_lcErrBase + g_lcErrGenericAssertion, "CAdmin.RunJob()", "A ENoticeType must be specified in the 'GetPage_Indirect'.  Rendering depends on it for Signature ENotices."
                    GoTo RunJob_ErrorExit
                End If
                
                Select Case m_lEddEnoticeType
                    Case EnoticeType.EnoticeType10and20
                        m_sPieceEddEnoticeType = "06"
                    Case EnoticeType.EnoticeType50
                        m_sPieceEddEnoticeType = "03"
                    Case EnoticeType.EnoticeType30
                        m_sPieceEddEnoticeType = "05"
                End Select
                
                '------------------------------------------
                'Place one piece (statement) in the statement body file
                If PrintPiece_Indirect() = False Then
                    GoTo RunJob_ErrorExit
                End If
                '------------------------------------------
                
                bNoPrint = False
                bCass = False
                bBulkShip = False
                bCertified = False
                sDPBC = vbNullString
                lForeignGroup = fgUSA
                bIndiciaRequired = False
                
                'Determine print/no print status
                bNoPrint = m_bNoPrint
                
                'Determine mail/no mail status
                If bNoPrint Then
                    bBulkShip = True
                End If
                bBulkShip = IIf(m_bXML_ReturnToBank, True, bBulkShip) 'RM12308
                If m_bReturnToBank Then
                    bBulkShip = True
                End If
                If Not bBulkShip Then
                    bBulkShip = IsBulkShip(m_oCassAddress)
                End If
                If Not bBulkShip And m_lBulkShipAccountsIDX > -1 Then
                    For i = 0 To m_lBulkShipAccountsIDX
                        If Trim(m_sAccountNumber) = Trim(m_saBulkShipAccounts(i)) Then
                            bBulkShip = True
                            If m_lEddEnoticeType > EnoticeTypeUnknown Then
                                bNoPrint = True
                            End If
                            Exit For
                        End If
                    Next
                End If
                
                'Determine cass/no cass status
                bCass = True
                If bBulkShip Then
                    bCass = False
                ElseIf bNoPrint Then
                    bCass = False
                End If
                
                'CASS certify
'                If bCass Then
'                    If lForeignGroup = fgUSA Then
'                        If Not m_oCassSink Is Nothing Then
'                            bCertified = m_oCassSink.GetDPBC(m_oCassAddress, sDPBC)
'                        End If
'                    End If
'                    If bCertified = False Then
'                        lForeignGroup = m_oAddressUtility.ForeignGroup(m_oCassAddress.CityStateZip)
'                    End If
'                End If
                
                '------------------------------------------
                'Add item to piece table in SBF database.
                ' This logic is special depending on whether the EddEnotice is running or not since 2 copies may
                ' Need to be inserted instead of just 1
                Set oPiece = New CPiece
                With oPiece
                    .Position = m_lLocation
                    .AccountNumber = m_sAccountNumber
                    .Duplex = m_bDuplex
                    .UserKey = Format(m_lLocation, "000000000")
                    .BarCode = vbNullString
'                    If m_oCassAddress Is Nothing Then
'                        .Name = vbNullString
'                        .AddressLine1 = vbNullString
'                        .StreetAddress = vbNullString
'                        .CityStateZip = vbNullString
'                    Else
                        .FullAddress = m_oCassAddress
'                    End If
                    
                    If m_lStackItems = StackItemsBypass Then
                        .StackId = m_lCurrentNoPrintStack + 5000
                    ElseIf m_bSamples Then
                        bIndiciaRequired = (m_lStackItems = StackItemsPressureSeal)
                        .BarCode = sDPBC
                        .StackId = m_lCurrentSampleStack
                    ElseIf bNoPrint Then
                        .StackId = 5
                    Else
                        Select Case m_lEddEnoticeType
                            Case EnoticeType.EnoticeType10and20
                                .StackId = 1
                            Case EnoticeType.EnoticeType50
                                .StackId = 3
                            Case EnoticeType.EnoticeType30
                                .StackId = 4
                            Case Else
                                ' Error message here
                        End Select
                    End If
                   
                   .InputSequence = m_lInputSequence
                   .OutputImages = m_lOutputImages
                   .OutputSheets = m_lOutputSheets
                   If .Duplex Then
                     .OutputClicks = m_lOutputSheets * 2
                   Else
                     .OutputClicks = m_lOutputSheets
                   End If
                
                End With
                
                If m_oStmtBodyDatabase.InsertPiece(oPiece) = False Then
                    GoTo RunJob_ErrorExit
                End If
                '------------------------------------------
            
                'Write applicable indicia and close out the piece.
                With m_oStmtBodyFile
                    If bIndiciaRequired Then
                        .WriteLine FiservMailRateCategorySelect
                        .WriteLine FiservMailRateCategoryCase & " International*"
                        .WriteLine FiservMailRateCategoryDefault
                            .WriteLine FiservMailRateCategoryConcatenator & PsString(m_sXmlAse, True) & " ITM42ASE" 'Ancillary Service Endorsement
                            .WriteLine FiservMailRateCategoryConcatenator & m_oIndicia.ExecutionCommand
                        .WriteLine FiservMailRateCategoryEndSelect
                        'm_lIndicia(m_lFormIdx) = m_lIndicia(m_lFormIdx) + 1
                    End If
                    PsPageEnd
                    .WriteLine "%FiservEndPiece:"
                    '.WriteLine "%FDSEndPiece:"
                End With
                
                '------------------------------------------
                'Accumulate input stack totals
                lTotalInputPieces = lTotalInputPieces + 1
                lTotalInputImages = lTotalInputImages + m_lInputImages
                m_lCurrentFilePieces = m_lCurrentFilePieces + 1
                m_lCurrentFileImages = m_lCurrentFileImages + m_lInputImages
                lTotalOutputImages = lTotalOutputImages + m_lOutputImages
                '------------------------------------------
                
                If m_lEddEnoticeType = EnoticeType10and20 Then
                    ' A copy of the statement is being inserted into the Wizard
                    Dim lCopyLocation As Long
                    lCopyLocation = m_oStmtBodyFile.Position
                
                    ' Set the index piece for the copy - "06" has already been written.  This will write "01"
                    m_sPieceEddEnoticeType = "01"
                    
                    ' Write out the new piece indexes for this file
                    With m_oStmtBodyFile
                        .WriteLine "%FiservBeginPieceCopy: " & CStr(m_lLocation)
                        For i = 0 To m_lIndexIdx
                            .WriteLine "%FiservPieceCopyIndex: " & m_saIndexKeys(i) & " " & m_saIndexValues(i)
                        Next
                        .WriteLine "%FiservPieceCopyIndex: EDDDocumentType " & m_sPieceEddEnoticeType & "-" & m_sAccountNumber
                        .WriteLine "%FiservEndPieceCopy:"
                    End With
                                                            
                    '------------------------------------------
                    'Add item to piece table in SBF database.
                    ' This logic is special depending on whether the EddEnotice is running or not since 2 copies may
                    ' Need to be inserted instead of just 1
                    With oPiece
                        If m_lStackItems = StackItemsBypass Then
                            .StackId = m_lCurrentNoPrintStack + 5000
                        Else
                            .StackId = 2
                        End If
                        .Position = lCopyLocation
                        .UserKey = Format(lCopyLocation, "000000000")
                    End With
                    If m_oStmtBodyDatabase.InsertPiece(oPiece) = False Then
                        GoTo RunJob_ErrorExit
                    End If
                    '------------------------------------------
                    
                    '------------------------------------------
                    'Accumulate input stack totals
                    lTotalInputPieces = lTotalInputPieces + 1
                    lTotalInputImages = lTotalInputImages + m_lInputImages
                    m_lCurrentFilePieces = m_lCurrentFilePieces + 1
                    m_lCurrentFileImages = m_lCurrentFileImages + m_lInputImages
                    lTotalOutputImages = lTotalOutputImages + m_lOutputImages
                    '------------------------------------------
                End If
                
                'Update the wizard with current completion status
                lPieceCount = lPieceCount + 1
                oJobStatusSink.Status m_oPrintStream.PosPercent, jstPercentage, bCanceled
                
                If bCanceled Then
                    'Cancel was clicked in the Wizard interface, it is asking us to close down now.
                    GoTo RunJob_ErrorExit
                End If
            Else
            
                '------------------------------------------
                'Place one piece (statement) in the statement body file
                If PrintPiece_Indirect() = False Then
                    GoTo RunJob_ErrorExit
                End If
                '------------------------------------------
                
                bNoPrint = False
                bCass = False
                bBulkShip = False
                bCertified = False
                sDPBC = vbNullString
                lForeignGroup = fgUSA
                bIndiciaRequired = False
                
                'Determine print/no print status
                bNoPrint = m_bNoPrint
                
                'Determine mail/no mail status
                If bNoPrint Then
                    bBulkShip = True
                End If
                bBulkShip = IIf(m_bXML_ReturnToBank, True, bBulkShip) 'RM12308
                If m_bReturnToBank Then
                    bBulkShip = True
                End If
                If Not bBulkShip Then
                    bBulkShip = IsBulkShip(m_oCassAddress)
                End If
                If Not bBulkShip And m_lBulkShipAccountsIDX > -1 Then
                    For i = 0 To m_lBulkShipAccountsIDX
                        If Trim(m_sAccountNumber) = Trim(m_saBulkShipAccounts(i)) Then
                            bBulkShip = True
                            Exit For
                        End If
                    Next
                End If
                
                'Determine cass/no cass status
                bCass = True
                If bBulkShip Then
                    bCass = False
                ElseIf bNoPrint Then
                    bCass = False
                End If
                
                'CASS certify
'                If bCass Then
'                    If lForeignGroup = fgUSA Then
'                        If Not m_oCassSink Is Nothing Then
'                            bCertified = m_oCassSink.GetDPBC(m_oCassAddress, sDPBC)
'                        End If
'                    End If
'                    If bCertified = False Then
'                        lForeignGroup = m_oAddressUtility.ForeignGroup(m_oCassAddress.CityStateZip)
'                    End If
'                End If
                
                '------------------------------------------

                'Add item to piece table in SBF database.
                Set oPiece = New CPiece
                With oPiece
                    .Position = m_lLocation
                    .AccountNumber = m_sAccountNumber
                    .Duplex = m_bDuplex
                    .UserKey = Format(m_lLocation, "000000000")
                    .BarCode = vbNullString
'                    If m_oCassAddress Is Nothing Then
'                        .Name = vbNullString
'                        .AddressLine1 = vbNullString
'                        .StreetAddress = vbNullString
'                        .CityStateZip = vbNullString
'                    Else
                        .FullAddress = m_oCassAddress
'                    End If
                    
                    If m_lStackItems = StackItemsBypass Then
                        .StackId = m_lCurrentNoPrintStack + 5000
                    ElseIf m_bSamples Then
                        bIndiciaRequired = (m_lStackItems = StackItemsPressureSeal)
                        .BarCode = sDPBC
                        .StackId = m_lCurrentSampleStack
                    ElseIf bNoPrint Then
                        .StackId = 2
                    ElseIf Len(m_sXML_EddJobKey) > 0 Then
                        .StackId = 1
                    ElseIf m_bXML_Archive Or m_bXML_EddEstatements Then
                        .StackId = 1
                    ElseIf bBulkShip Then
                        .StackId = m_lCurrentBulkShipStack
                    Else
                        Select Case m_lStackItems
                            Case StackItemsPressureSeal
                                .StackId = 108
                                 bIndiciaRequired = True
'                                If lForeignGroup = fgCanada Then
'                                    .StackId = 103
'                                ElseIf lForeignGroup = fgMexico Then
'                                    .StackId = 104
'                                ElseIf lForeignGroup = fgAllOthers Then
'                                    .StackId = 105
'                                ElseIf bCertified Then
'                                    bIndiciaRequired = True
'                                    .BarCode = sDPBC
'                                    .StackId = 108
'                                Else
'                                    bIndiciaRequired = True
'                                   .StackId = 107
'                                End If
                            Case StackItemsFlatStock, StackItemsFlatStock8LPI, StackItemsFlatStock9LPI, StackItemsFlatStockLegal, StackItemsNotice, StackItemsFlatStock42, StackItemsFlatStock428LPI, StackItemsLN0473, StackItemsLN0473_152
                                Dim lStackOffset As Long
                                lStackOffset = 200
                                If lForeignGroup <> fgUSA Then
                                    .StackId = m_oStmtBodyDatabase.GetCleansedAddressStackId(lStackOffset + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets), International)
                                Else
                                    .StackId = lStackOffset + 50 + IIf(m_lOutputSheets > 9, 10, m_lOutputSheets)
                                End If
'                                If lForeignGroup = fgCanada Then
'                                    .StackId = lStackOffset + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets)
'                                ElseIf lForeignGroup = fgMexico Then
'                                    .StackId = lStackOffset + 10 + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets)
'                                ElseIf lForeignGroup = fgAllOthers Then
'                                    .StackId = lStackOffset + 20 + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets)
'                                ElseIf bCertified Then
'                                    .BarCode = sDPBC
'                                    .StackId = lStackOffset + 50 + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets)
'                                Else
'                                   .StackId = lStackOffset + 40 + IIf(m_lOutputSheets >= 10, 10, m_lOutputSheets)
'                                End If
                                If m_sXmlInst = "BPB" Then
                                    If m_lStackItems = StackItemsFlatStockLegal Then
                                        If m_sEnvelopeCatCode = m_scEnvelopeBorel Then
                                            .StackId = .StackId + 700
                                        Else
                                            .StackId = .StackId + 600
                                        End If
                                        Select Case m_sEnclosureCatCode
                                            Case ""
                                            Case Else
                                                Stop
                                        End Select
                                    Else
                                        If m_sEnvelopeCatCode = m_scEnvelope42 Then
                                            .StackId = .StackId + 700
                                        End If
                                        Select Case m_sEnclosureCatCode
                                            Case ""
                                            Case m_scEnclosureBPBEL9
                                                .StackId = .StackId + 100
                                            Case m_scEnclosureBPBER9
                                                .StackId = .StackId + 200
                                            Case m_scEnclosureBPBBORELER9
                                                .StackId = .StackId + 500

                                            Case Else
                                                Stop
                                        End Select
                                    End If
                                End If
                                
                                If m_sXmlInst = "WooriAmerica" Then
                                    Select Case m_sEnclosureCatCode
                                        Case m_scWooriT152Q2BRE
                                            .StackId = .StackId + 100
                                        Case Else
                                    
                                    End Select
                                End If
                                
                                If m_sXmlInst = "Nationwide" Then
                                    Select Case m_sEnclosureCatCode
                                        Case m_scNationwideBRE
                                            .StackId = .StackId + 100 ' Stack 300 is for T152Q2 and LN0480Q1
                                    End Select
                                End If
                                
                                If m_sXmlInst = "Investors" Then
                                    If m_lStackItems = StackItemsFlatStockLegal Then
                                        .StackId = .StackId + 100 ' Stack 300 is for Legal-Sized Pieces
                                    Else
                                        Select Case m_sEnclosureCatCode
                                            Case m_scInvestorsBRE
                                                .StackId = .StackId + 200 ' Stack 400 is for TM0851P1 BRE envelopes
                                            Case m_scInvestorsBREMORE
                                                .StackId = .StackId + 300 ' Stack 500 is for TM0847P1 BRE envelopes + extra enclosures
                                        End Select
                                    End If
                                End If
                        End Select
                    End If
                    
                   .InputSequence = m_lInputSequence
                   .OutputImages = m_lOutputImages
                   .OutputSheets = m_lOutputSheets
                   If .Duplex Then
                     .OutputClicks = m_lOutputSheets * 2
                   Else
                     .OutputClicks = m_lOutputSheets
                   End If
                
                End With
                
                If m_oStmtBodyDatabase.InsertPiece(oPiece) = False Then
                    GoTo RunJob_ErrorExit
                End If
                '------------------------------------------
            
                'Write applicable indicia and close out the piece.
                With m_oStmtBodyFile
                    If bIndiciaRequired Then
                        .WriteLine FiservMailRateCategorySelect
                        .WriteLine FiservMailRateCategoryCase & " International*"
                            'Don't apply indicia and ase to international
                        .WriteLine FiservMailRateCategoryDefault
                            .WriteLine FiservMailRateCategoryConcatenator & PsString(m_sXmlAse, True) & " ITM42ASE" 'Ancillary Service Endorsement
                            .WriteLine FiservMailRateCategoryConcatenator & m_oIndicia.ExecutionCommand
                        .WriteLine FiservMailRateCategoryEndSelect
                        'm_lIndicia(m_lFormIdx) = m_lIndicia(m_lFormIdx) + 1
                    End If
                    PsPageEnd
                    .WriteLine "%FiservEndPiece:"
                    '.WriteLine "%FDSEndPiece:"
                End With
                
                If Len(m_sXML_EddJobKey) > 0 Or m_bXML_EddEstatements Then
                    If Not bNoPrint Then
                        If Len(m_sDneFirstAccount) = 0 Then
                            m_sDneFirstAccount = m_sAccountNumber
                        End If
                        m_sDneLastAccount = m_sAccountNumber
                        m_lDneStatementCount = m_lDneStatementCount + 1
                        m_lDneImageCount = m_lDneImageCount + oPiece.OutputImages
                    End If
        
                    If Len(m_sDneDate) = 0 Then
                        If Len(m_sDocumentDate) > 0 Then
                            m_sDneDate = Format(CDate(m_sDocumentDate), "yyyymmdd")
                        End If
                    End If
                End If
                
                '------------------------------------------
                'Accumulate input stack totals
                lTotalInputPieces = lTotalInputPieces + 1
                lTotalInputImages = lTotalInputImages + m_lInputImages
                m_lCurrentFilePieces = m_lCurrentFilePieces + 1
                m_lCurrentFileImages = m_lCurrentFileImages + m_lInputImages
                lTotalOutputImages = lTotalOutputImages + m_lOutputImages
                '------------------------------------------
                
                'Update the wizard with current completion status
                lPieceCount = lPieceCount + 1
                oJobStatusSink.Status m_oPrintStream.PosPercent, jstPercentage, bCanceled
                
                If bCanceled Then
                    'Cancel was clicked in the Wizard interface, it is asking us to close down now.
                    GoTo RunJob_ErrorExit
                End If
                
            End If
        
            bMorePages = GetPage_Indirect()
        
        Loop 'End of processing all statements
        
        'Update input stack totals
        If m_InputFileFileRecordCount > 0 Then
            If m_lCurrentFileExpectedPages >= 0 Then
                ' Adjusts the expected file pages with eNotices - certain types require 2 output pieces
                If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                    If m_lEddEnoticeType = EnoticeType10and20 Then
                        m_lCurrentFileExpectedPages = m_lCurrentFileExpectedPages * 2
                    End If
                End If
                If m_lCurrentFileImages <> m_lCurrentFileExpectedPages Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "File does not have the specified number of pages - " & m_sCurrentFileName & "'."
                End If
            End If
        End If
        
        If m_bcVirtualFileSupport Then
            'First, record this input file in the stack table.
            Set oStack = New CStack
            oStack.SetProperties 10000 + m_lCurrentFileNumber, StackInput, m_sCurrentFileName & " " & m_sCurrentFileFormType & IIf(m_lCurrentNoPrintStack > 0, " (Bypassed)", ""), BalGroupInput, StackSimplex, False, False, 0
            m_oStmtBodyDatabase.InsertStack oStack
            Set oStack = Nothing
        
            'Insert the input stack totals into stack totals table.
            m_oStmtBodyDatabase.InsertInputTotals 10000 + m_lCurrentFileNumber, m_lCurrentFilePieces, m_lCurrentFileImages
        Else
            'First, record this input file in the stack table.
            Set oStack = New CStack
            oStack.SetProperties 10000 + m_lCurrentFileNumber, StackInput, m_sInputFile(m_lInputFilePtr), BalGroupInput, StackSimplex, False, False, 0
            m_oStmtBodyDatabase.InsertStack oStack
            Set oStack = Nothing
        
            'Insert the input stack totals into stack totals table.
            m_oStmtBodyDatabase.InsertInputTotals 10000 + m_lCurrentFileNumber, m_lCurrentFilePieces, m_lCurrentFileImages
        End If
        
        bAnotherFile = IncrementInputFile()
        
    Loop While bAnotherFile
    
    PsWriteTrailer lTotalOutputImages
    
    '---------------------------------------------------------------------
    'Write applicable finish directives
    Dim oFinishDirectivesDom As MSXML2.DOMDocument60
    If Len(m_sXML_EddJobKey) > 0 Then
        If m_sEddCycleDate = 0 Then
            m_sEddCycleDate = Now
        End If
        WriteEddJobSettings oFinishDirectivesDom, m_sXML_EddJobKey, m_sEddCycleDate, m_sXml_EddClientId, m_sXML_EddProductId, m_bXML_EddCompanionFile, m_sXML_EddDocType
    ElseIf Len(m_sXML_EddEnoticeJobKey) > 0 Then
        If m_sEddCycleDate = 0 Then
            m_sEddCycleDate = Now
        End If
        'Writes the finishing directives for all 4 files
        WriteEddJobSettings oFinishDirectivesDom, m_sXML_EddEnoticeJobKey & "06", m_sEddCycleDate, m_sXML_EddEnoticeClientId, "06", True, m_sXML_EddDocType
        WriteEddJobSettings oFinishDirectivesDom, m_sXML_EddEnoticeJobKey & "01", m_sEddCycleDate, m_sXML_EddEnoticeClientId, "01", True, m_sXML_EddDocType
        WriteEddJobSettings oFinishDirectivesDom, m_sXML_EddEnoticeJobKey & "03", m_sEddCycleDate, m_sXML_EddEnoticeClientId, "03", True, m_sXML_EddDocType
        WriteEddJobSettings oFinishDirectivesDom, m_sXML_EddEnoticeJobKey & "05", m_sEddCycleDate, m_sXML_EddEnoticeClientId, "05", True, m_sXML_EddDocType
    End If
    
    If Not oFinishDirectivesDom Is Nothing Then
        Dim sFinishDirectivesPath As String
        sFinishDirectivesPath = AddExtension(RemoveExtension(m_sStmtBodyFileName) & "_fd", ".xml")
        oFinishDirectivesDom.save sFinishDirectivesPath
    End If
    '---------------------------------------------------------------------
    
    If m_bXML_EddEstatements Then
        If OpenDneFile(Replace(oJob.Mnemonic, "_", "."), m_sDneDate) Then
            With m_oDneFile
                .WriteLine "Job Name: " & m_oJob.OutFileName
                .WriteLine "Statement Count: " & CStr(m_lDneStatementCount)
                .WriteLine "Image Count: " & CStr(m_lDneImageCount)
                .WriteLine "First Account: " & m_sDneFirstAccount
                .WriteLine "Last Account: " & m_sDneLastAccount
            End With
            m_oDneFile.CloseFile
            Set m_oDneFile = Nothing
        End If
    End If
    
    If Not m_oPrintStream Is Nothing Then
        m_oPrintStream.Close
        Set m_oPrintStream = Nothing
    End If
    Set m_oStmtBodyFile = Nothing
    Set m_oStmtBodyDatabase = Nothing
    
    IRunJob_RunJob = True
    
    Exit Function
    
RunJob_ErrorExit:
    
    On Error Resume Next
    
    If Not m_oPrintStream Is Nothing Then
        m_oPrintStream.Close
        Set m_oPrintStream = Nothing
    End If
    Set m_oStmtBodyFile = Nothing
    Set m_oStmtBodyDatabase = Nothing
    
    IRunJob_RunJob = False
    Exit Function
    
RunJob_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
       Case 1 'User chose Retry
          On Error GoTo RunJob_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo RunJob_Error
          Resume Next
       Case 3 'User chose Abort
          GoTo RunJob_ErrorExit
    End Select
    
    Exit Function
    
End Function

Private Function PrintPiece_Bypass() As Boolean
    Dim sngPsCurrPosY As Single
    Dim bEndOfPiece As Boolean
    Dim i As Long
    
    On Error GoTo PrintPiece_Bypass_Error
    
    m_lInputImages = 1
    m_lOutputSheets = 1
    m_lOutputImages = 1
    'm_lOutputClicks = 1
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
        
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_Bypass = True
    
    Exit Function
   
PrintPiece_Bypass_ErrorExit:
    
    PrintPiece_Bypass = False
    Exit Function
   
PrintPiece_Bypass_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_Bypass()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_Bypass_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_Bypass_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_Bypass_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_Bypass_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_Bypass_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_Bypass_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_Bypass = False
    Exit Function
    
End Function

Private Function PrintPiece_Notice() As Boolean
    Dim sngPsCurrPosY As Single
    Dim bEndOfPiece As Boolean
    Dim i As Long
    Dim sngLineSpacing As Single
    
    Const lcTopMargin As Long = 441
    Const lcBottomMargin As Long = 0
    Const sngcLineSpacing As Single = 9.2
    
    On Error GoTo PrintPiece_Notice_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    sngLineSpacing = sngcLineSpacing
    'm_lOutputClicks = 0
    
    ' Sets up proper rendering coordinates for this print run
    Select Case m_sXmlInst
        Case "Nationwide"
            Select Case m_sCurrentFileName
                Case "CU1610P1"
                    sngLineSpacing = 8.7
                Case Else
                    sngLineSpacing = 9.2
            End Select
        Case Else
            Select Case m_sCurrentFileName
                Case Else
                    sngLineSpacing = 9.2
            End Select
    End Select

    If m_sXmlInst = "BPB" Then
        If m_sCurrentFileName = "CU1610P1" Then
            CondensePageArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx
        End If
    End If
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
        
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    m_lInputImages = m_lInputImages + 1
    
    PsPageBegin "FDSNotice"
    
    With m_oStmtBodyFile
        
'        .WriteLine "userdict /FDSNotice get begin"
        
        If Not m_oExteriorLogoFormInfo Is Nothing Then
            '.WriteLine "gs"
            .WriteLine m_oExteriorLogoFormInfo.ExecutionCommand
            '.WriteLine "gr"
        End If
        
        If m_lStackItems = StackItemsPressureSeal Then
            If Not m_oInteriorLogoFormInfo Is Nothing Then
                '.WriteLine "gs"
                .WriteLine m_oInteriorLogoFormInfo.ExecutionCommand
                '.WriteLine "gr"
            End If
    
            .WriteLine PsProcITM42_DELADDR(m_oDeliveryAddress)
            .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
            .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) ITM42PostNet"
            .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) ITM42SEQ"
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCDPS"
        Else
            .WriteLine PsProcSTD10_DELADDR(m_oDeliveryAddress)
            .WriteLine "%FiservMailPieceAddress: 72 581 10 6"
            .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) STD10PostNet"
            .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) STD10SEQ"
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
        End If
            

'        .WriteLine "gs"
'        .WriteLine "54 306 translate"
        
        sngPsCurrPosY = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
        
        If (m_lStackItems = StackItemsPressureSeal) And (Not m_oInteriorLogoFormInfo Is Nothing) Then
            sngPsCurrPosY = sngPsCurrPosY - m_oInteriorLogoFormInfo.ImageBoundingBox.Height
            sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing * 2 'White space
        Else
            If Not m_oXML_BankAddress Is Nothing Then
                For i = 1 To m_oXML_BankAddress.CondensedAddressLines
                    sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                    .WriteLine PsString(m_oXML_BankAddress.CondensedAddress(i), True) & " 0 " & CStr(sngPsCurrPosY) & " DTL"
                Next
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing * 2 'White space
            End If
        End If
        
        For i = 0 To m_lCurrentPageIdx
            sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
            If Len(m_saCurrentPageLineBold(i)) > 0 Then
                m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
            End If
            If Len(m_saCurrentPageLine(i)) > 0 Then
                m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
            End If
        Next
        
        If sngPsCurrPosY < m_lcSelfSealPanels_y Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Line will print off bottom of page."
        End If
         
'        .WriteLine "gr"
    
    End With
        
   
    ' Page gets ended in RunJob after the address is cassed and gets indicia
    'PsPageEnd
    'm_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_Notice = True
    
    Exit Function
   
PrintPiece_Notice_ErrorExit:
    
    PrintPiece_Notice = False
    Exit Function
   
PrintPiece_Notice_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_Notice()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_Notice_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_Notice_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_Notice_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_Notice_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_Notice_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_Notice_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_Notice = False
    Exit Function
    
End Function

Private Function IsSectionSep(ByVal lIndex As Long, ByRef sSectionDescription As String, ByVal sLine As String) As Boolean
    Dim oMatches As MatchCollection
    IsSectionSep = False
    sSectionDescription = vbNullString
    
    m_oRegExp.Pattern = "^\s*(?:-)+\s+(.+?)\s+(?:-)+\s*$"
    If m_oRegExp.Test(sLine) Then
        Set oMatches = m_oRegExp.Execute(sLine)
        If oMatches.Count > 0 Then
            sSectionDescription = Trim(oMatches.Item(0).SubMatches(0))
            IsSectionSep = True
            Exit Function
        End If
    End If
End Function

Private Function PrintPiece_FlatStock8LPI() As Boolean
    Dim sngPsCurrPosY As Single
  '  Const sngLineSpacing As Single = 10
    Const sngLineSpacing As Single = 9
'    Dim sngLineSpacing As Single
'    sngLineSpacing = 9
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStock8LPI_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    'm_lOutputClicks = 0
    'm_bDuplex = False
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Dim sngRiverCityInfoPosX As Single
    
    
    Do
    
        sngRiverCityInfoPosX = 0
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile
            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            'Note: Allstate gets logo and address.
            

            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
            End If

            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            Select Case m_sXmlInst
                Case "WooriAmerica"
                    Select Case m_sCurrentFileName
                        Case "T152Q2"
                            sngPsCurrPosY = 765
                            'sngLineSpacing = 8.5
                        Case Else
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Woori America Bank - No top margin set for " & m_sCurrentFileName
                    End Select
                    .WriteLine "66 0 translate"
                Case "Investors"
                    Select Case m_sCurrentFileName
                        Case "TA0050P1"
                            sngPsCurrPosY = 591
'                        Case "PRINT1"
'                            sngPsCurrPosY = 774
                        Case Else
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Investors Bank - No top margin set for " & m_sCurrentFileName
                    End Select
                    .WriteLine "54 0 translate"
                    'CPTPK-1472
                Case "BPB"
                Select Case m_sCurrentFileName
                'CPTPK-1712
                        Case "TA9544P1"
                            sngPsCurrPosY = 810
                        Case "TA9544P2"
                            sngPsCurrPosY = 810
                            'End CPTPK-1712
                        Case "TA9544P3"
                            sngPsCurrPosY = 810
                        Case Else
                            sngPsCurrPosY = 792
                    End Select
'                    CPTPK -1472
                Case Else
                    sngPsCurrPosY = 792
                    .WriteLine "66 0 translate"
            End Select
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
            
                If m_sXmlInst = "RiverCity" Then
                    If m_sCurrentFileName = "T152Q2" Then
                        If m_lInputImages = 1 Then
                            If i > 22 Then
                                If sngRiverCityInfoPosX = 0 Then
                                    If InStr(1, m_saCurrentPageLine(i), "*----------Analysis Summary----------*", vbTextCompare) > 0 Then
                                        Dim k As Long
                                        For k = i - 1 To i - 6 Step -1
                                            If Len(Trim(m_saCurrentPageLine(k))) > 0 Then
                                                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "The six lines above the Account Analysis sub title need to be blank ino order to place the customer service information text for River City.   The lines are not blank."
                                            End If
                                        Next
                                        sngRiverCityInfoPosX = sngPsCurrPosY
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                
                If Len(m_saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                End If
                If Len(m_saCurrentPageLine(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
                       
'            If Not m_oContactInfoFormInfo Is Nothing Then
'                If sngRiverCityInfoPosX <> 0 Then
'                    .WriteLine "gs 0 " & CStr(sngRiverCityInfoPosX + (sngLineSpacing * 6)) & " translate"
'                    .WriteLine m_oContactInfoFormInfo.ExecutionCommand
'                    .WriteLine "gr"
'                End If
'            End If
        
            .WriteLine "gr"
            
            If sngRiverCityInfoPosX <> 0 Then
                .WriteLine CStr(sngRiverCityInfoPosX) & " RCBANK"
            End If
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_FlatStock8LPI = True
    
    Exit Function
   
PrintPiece_FlatStock8LPI_ErrorExit:
    
    PrintPiece_FlatStock8LPI = False
    Exit Function
   
PrintPiece_FlatStock8LPI_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStock8LPI()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStock8LPI_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStock8LPI_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStock8LPI_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStock8LPI_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStock8LPI_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStock8LPI_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStock8LPI = False
    Exit Function
    
End Function

Private Function PrintPiece_FlatStock9LPI() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 8
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStock9LPI_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile
            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            'Note: Allstate gets logo and address.
            

            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
            End If

            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            sngPsCurrPosY = 688
            
            .WriteLine "gs"
            .WriteLine "100 0 translate"
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
            
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                
                If Len(m_saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB8(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                End If
                If Len(m_saCurrentPageLine(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTL8(sngPsCurrPosY, m_saCurrentPageLine(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
         
    PrintPiece_FlatStock9LPI = True
    
    Exit Function
   
PrintPiece_FlatStock9LPI_ErrorExit:
    
    PrintPiece_FlatStock9LPI = False
    Exit Function
   
PrintPiece_FlatStock9LPI_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStock9LPI()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStock9LPI_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStock9LPI_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStock9LPI_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStock9LPI_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStock9LPI_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStock9LPI_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStock9LPI = False
    Exit Function
    
End Function

Private Function PrintPiece_FlatStock() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 10
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStock_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    'm_lOutputClicks = 0
    'm_bDuplex = False
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            If Not m_oReturnAddress Is Nothing Then
                If m_sXmlInst = "WooriAmerica" Then
                    .WriteLine PsProc_RETADDRWoori(m_oReturnAddress)
                Else
                    .WriteLine PsProc_RETADDR(m_oReturnAddress)
                        
                End If
            End If
            

            'Note: Allstate gets logo and address.
            
             Select Case m_sXmlInst
                Case "BPB"
                    Select Case m_sCurrentFileName
                        Case "TM0330P1"
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
                            End If
                            
                            If m_lOutputImages = 1 Then
                                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
                            End If
                        
                        Case Else
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
                            End If
                            
                            If m_lOutputImages = 1 Then
                                If Len(m_sOCR) > 0 Then
                                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                End If
                                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
                            End If
                    End Select
                'RM5096
                Case "SNB"
                    Select Case m_sCurrentFileName
                        Case "LN0470P1"
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
                            End If

                            If m_lOutputImages = 1 Then
                                If Len(m_sOCR) > 0 Then
                                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                End If
                                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
                            End If
                        Case "E_LN0470P1"
                            If Len(m_sXML_EddJobKey) > 0 Then
                               If Not m_oDeliveryAddress Is Nothing Then
                                    .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
                                End If

                                If m_lOutputImages = 1 Then
                                    If Len(m_sOCR) > 0 Then
                                        .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                    End If
                                    .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                                    .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                                    .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
                                End If
                            End If
                            
                        Case "SD0095P3"
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
                            End If

                            If m_lOutputImages = 1 Then
                                If Len(m_sOCR) > 0 Then
                                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                End If
                                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
                            End If
                            
                        Case "SD0095P2"
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
                            End If

                            If m_lOutputImages = 1 Then
                                If Len(m_sOCR) > 0 Then
                                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                End If
                                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
                            End If
                            
                        Case Else
                            If Not m_oDeliveryAddress Is Nothing Then
                                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
                            End If

                            If m_lOutputImages = 1 Then
                                If Len(m_sOCR) > 0 Then
                                
                                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                                End If
                                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
                            End If
                    End Select
                Case Else
                    If Not m_oDeliveryAddress Is Nothing Then
                        .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
                    End If
                    
                    If m_lOutputImages = 1 Then
                        If Len(m_sOCR) > 0 Then
                            .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                        End If
                        .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                        .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                        .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
                    End If
            End Select
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            Select Case m_sXmlInst
                Case "WooriAmerica"
                    Select Case m_sCurrentFileName
                        Case "LN0458P1", "LN0460P1", "LN0470P1"
                            .WriteLine "60 0 translate"
                            sngPsCurrPosY = 750
                        Case "TA0185P1"
                            .WriteLine "66 0 translate"
                            sngPsCurrPosY = 765
                        Case Else
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "No top margin set for " & m_sCurrentFileName
                    End Select
                Case "BPB"
                    Select Case m_sCurrentFileName
                        Case "TM0360P2", "TM0360P4", "TM0360P5", "TM0360P7", "TM0360P8"
                            sngPsCurrPosY = 535
                        Case "LN93357P2"
                            sngPsCurrPosY = 720 ' Leave room for BPB OCR line above data.
                        Case "CH0202P4"
                            sngPsCurrPosY = 500
                        Case Else
                            sngPsCurrPosY = 728
                    End Select
                    .WriteLine "66 0 translate"
                Case "Investors"
                    Dim lXtranslate As Long
                    lXtranslate = 54
                    Select Case m_sCurrentFileName
                        Case "TA955289P"
                            sngPsCurrPosY = 591
                            lXtranslate = 45
                        Case "TA0434P1"
                            sngPsCurrPosY = 591
                        Case "TM0365P1"
                            sngPsCurrPosY = 591
                        Case "CH0202P1", "CH0202P4"
                            sngPsCurrPosY = 591
                        Case "TM0847P1"
                            sngPsCurrPosY = 710
                        Case "CU1610P1", "CU1620P1"
                            sngPsCurrPosY = 710
                        Case "TM0851P1"
                            sngPsCurrPosY = 710
                        Case "TA1161P1"
                            sngPsCurrPosY = 591
'                        Case "PRINT1"
'                            sngPsCurrPosY = 774
                        Case "FT0160P1", "FT0160P2"
                            sngPsCurrPosY = 591
                        Case "SD0095P2", "SD0095P3", "SD0095P4"
                            sngPsCurrPosY = 591
                        Case Else
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Investors Bank - No top margin set for " & m_sCurrentFileName
                    End Select
                    .WriteLine CStr(lXtranslate) & " 0 translate"
                Case Else
                    Select Case m_sCurrentFileName
                        Case "LN93357P2"
                            sngPsCurrPosY = 720 ' Leave room for BPB OCR line above data.
                        Case "CH0202P4"
                            sngPsCurrPosY = 500
                        Case Else
                            sngPsCurrPosY = 591
                    End Select
                    .WriteLine "66 0 translate"
            End Select
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                If Len(m_saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                End If
                If Len(m_saCurrentPageLine(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_FlatStock = True
    
    Exit Function
   
PrintPiece_FlatStock_ErrorExit:
    
    PrintPiece_FlatStock = False
    Exit Function
   
PrintPiece_FlatStock_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStock()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStock_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStock_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStock_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStock_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStock_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStock_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStock = False
    Exit Function
    
End Function

Private Function PrintPiece_FlatStock42() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 10
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStock42_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    'm_lOutputClicks = 0
    'm_bDuplex = False
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If

            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
            End If
            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            ' Sets up proper rendering coordinates for this print run
            Select Case m_sXmlInst
                Case "SNB"
                    Select Case m_sCurrentFileName
                        Case "TA990770P1"
                            sngPsCurrPosY = 594
                            .WriteLine "72 0 translate"
                        Case "TA990771P1"
                            sngPsCurrPosY = 600
                            .WriteLine "18 0 translate"
                        Case "T152Q2"
                            sngPsCurrPosY = 774
                            .WriteLine "54 0 translate"
                        Case "TM0330P1"
                            sngPsCurrPosY = 720
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case "Nationwide"
                    Select Case m_sCurrentFileName
                        Case "CU1610P1"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case "LN0480Q1", "E_QSYSPRT"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case "TM0847P1"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case "TM0360P2", "TM0360P7"
                            sngPsCurrPosY = 582
                            .WriteLine "36 0 translate"
                        Case "TM0330P1"
                            sngPsCurrPosY = 720
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 582
                            .WriteLine "54 0 translate"
                    End Select
                Case "RiverCity"
                    Select Case m_sCurrentFileName
                        Case "LN0480Q1"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case "ElDorado"
                    Select Case m_sCurrentFileName
                        Case "TM0330P1"
                            sngPsCurrPosY = 720
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case "Investors"
                    Select Case m_sCurrentFileName
                        Case "CU1610P1"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case m_scInst_29175100_FinFactors
                    Select Case m_sCurrentFileName
                        Case "LN0480Q1"
                            sngPsCurrPosY = 710
                            .WriteLine "54 0 translate"
                        Case "TM0330P1"
                            .WriteLine "Form_TM0330P1_29175100"
                            sngPsCurrPosY = 720
                            .WriteLine "54 0 translate"
                    End Select
                Case "55810000", "75535000"
                    Select Case m_sCurrentFileName
                        Case "TM0330P1"
                            sngPsCurrPosY = 720
                            .WriteLine "54 0 translate"
                    End Select
                Case Else
                    Select Case m_sCurrentFileName
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
            End Select
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                If Len(m_saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, Replace(m_saCurrentPageLineBold(i), Chr(26), ""))
                End If
                If Len(m_saCurrentPageLine(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, Replace(m_saCurrentPageLine(i), Chr(26), ""))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_FlatStock42 = True
    
    Exit Function
   
PrintPiece_FlatStock42_ErrorExit:
    
    PrintPiece_FlatStock42 = False
    Exit Function
   
PrintPiece_FlatStock42_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStock42()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStock42_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStock42_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStock42_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStock42_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStock42_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStock42_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStock42 = False
    Exit Function
    
End Function

Private Function PrintPiece_FlatStock428LPI() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 9
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStock428LPI_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    'm_lOutputClicks = 0
    'm_bDuplex = False
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            
            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR42(m_oDeliveryAddress)
            End If
            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: 54 634 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet42"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ42"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            ' Sets up proper rendering coordinates for this print run
            Select Case m_sXmlInst
                Case "SNB"
                    Select Case m_sCurrentFileName
                        Case "TA990770P1"
                            sngPsCurrPosY = 600
                            .WriteLine "18 0 translate"
                        Case "TA990771P1"
                            sngPsCurrPosY = 600
                            .WriteLine "18 0 translate"
                        Case "T152Q2"
                            sngPsCurrPosY = 774
                            .WriteLine "54 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case "Nationwide"
                    Select Case m_sCurrentFileName
                        Case "T152Q2"
                            sngPsCurrPosY = 738
                            .WriteLine "54 0 translate"
                        Case "LN0111Q1"
                            sngPsCurrPosY = 564
                            .WriteLine "63 0 translate"
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
                Case Else
                    Select Case m_sCurrentFileName
                        Case Else
                            sngPsCurrPosY = 600
                            .WriteLine "54 0 translate"
                    End Select
            End Select
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                ' Check we are nationwide LN0111Q1 - use a smaller font
                If m_sXmlInst = "Nationwide" And (m_sCurrentFileName = "LN0111Q1") Then
                    If Len(m_saCurrentPageLineBold(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTLB9(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                    End If
                    If Len(m_saCurrentPageLine(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTL9(sngPsCurrPosY, m_saCurrentPageLine(i))
                    End If
                ' Use the standard font size
                Else
                    If Len(m_saCurrentPageLineBold(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                    End If
                    If Len(m_saCurrentPageLine(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
                    End If
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
            If Len(Trim(m_sMICR)) > 0 Then
                .WriteLine PsString(m_sMICR, True) & " MICR"
            End If
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_FlatStock428LPI = True
    
    Exit Function
   
PrintPiece_FlatStock428LPI_ErrorExit:
    
    PrintPiece_FlatStock428LPI = False
    Exit Function
   
PrintPiece_FlatStock428LPI_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStock()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStock428LPI_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStock428LPI_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStock428LPI_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStock428LPI_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStock428LPI_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStock428LPI_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStock428LPI = False
    Exit Function
    
End Function


Private Function PrintPiece_FlatStockLegal() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 9
    Dim bEndOfPiece As Boolean
    'Dim bLetter As Boolean
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_FlatStockLegal_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    'm_lOutputClicks = 0
    m_bDuplex = False
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    'm_oStmtBodyFile.WriteLine "%FDSBeginPiece:"
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    If m_sXmlInst = "Investors" And m_sCurrentFileName = "LN0860P1" Then
'        Dim sAccountNumber As String
'        If Len(m_sAccountNumber) < 10 Then
'            sAccountNumber = Right(String(10, "0") & m_sAccountNumber, 10)
'        Else
'            sAccountNumber = m_sAccountNumber
'        End If
        m_oStmtBodyFile.WriteLine "%FiservDocumentKey: " & m_sAccountNumber & "~" & GetGUID()
    End If
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStockLegal "FDSFlatStock", m_bDuplex
            
            .WriteLine "gs 0 216 translate" 'Translate address elements up for Legal
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            'Note: Allstate gets logo and address.
            
            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
            End If
            
            If m_lOutputImages = 1 Then
                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gr" 'Translate address elements up for Legal
            
            .WriteLine "gs"
            
            'sngPsCurrPosY = 738
            sngPsCurrPosY = 954
            
            Select Case m_sXmlInst
                Case "Investors"
                    Select Case m_sCurrentFileName
                        Case "LN0860P1"
                            sngPsCurrPosY = IIf(m_lOutputImages > 1, 1000, 981)
                            .WriteLine "72 0 translate"
'                        Case "PRINT1"
'                            sngPsCurrPosY = 909
'                            .WriteLine "81 0 translate"
                        Case Else
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "No top margin set for " & m_sCurrentFileName
                    End Select
                Case "BPB"
                    Select Case m_sCurrentFileName
                        Case "LN0860P1"
                            sngPsCurrPosY = 963
                            .WriteLine "72 0 translate"
                        Case Else
                            sngPsCurrPosY = 954
                            .WriteLine "72 0 translate"
                    End Select
                Case Else
                    sngPsCurrPosY = 954
                    .WriteLine "72 0 translate"
            End Select
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                If Len(m_saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB9(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                End If
                If Len(m_saCurrentPageLine(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTL9(sngPsCurrPosY, m_saCurrentPageLine(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
    Loop Until GetTrailerPage_Indirect = False

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
    'm_oStmtBodyFile.WriteLine "%FDSEndPiece:"
         
    PrintPiece_FlatStockLegal = True
    
    Exit Function
   
PrintPiece_FlatStockLegal_ErrorExit:
    
    PrintPiece_FlatStockLegal = False
    Exit Function
   
PrintPiece_FlatStockLegal_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_FlatStockLegal()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_FlatStockLegal_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_FlatStockLegal_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_FlatStockLegal_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_FlatStockLegal_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_FlatStockLegal_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_FlatStockLegal_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_FlatStockLegal = False
    Exit Function
    
End Function

Private Function PrintPiece_LN0473() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 11
    Dim bEndOfPiece As Boolean
    Dim oMatch As Match
    Dim oMatches As MatchCollection
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_LN0473_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            
            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
            End If
            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            .WriteLine "66 0 translate"
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            Dim bInBoldPara As Boolean
            bInBoldPara = False
            Dim saCurrentPageLineBold() As String 'Added for doing bold sections of a line.
            ReDim saCurrentPageLineBold(m_lCurrentPageIdx)
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineBold(i) = Space(Len(m_saCurrentPageLine(i)))
            Next
            
            Dim saCurrentPageLineUnderline() As String 'Added for doing underline sections of a line.
            ReDim saCurrentPageLineUnderline(m_lCurrentPageIdx)
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineUnderline(i) = Space(Len(m_saCurrentPageLine(i)))
            Next
            
            For i = 0 To m_lCurrentPageIdx
                If bInBoldPara Then
                    If Len(Trim(m_saCurrentPageLine(i))) = 0 Then
                        bInBoldPara = False
                    End If
                Else
                    If InStr(1, m_saCurrentPageLine(i), "**Under the terms of your Adjustable-Rate Mortgage", vbTextCompare) > 0 Then
                        bInBoldPara = True
                    ElseIf InStr(1, m_saCurrentPageLine(i), "However, if the ", vbTextCompare) > 0 Then
                        bInBoldPara = True
                    End If
                End If
                
                If bInBoldPara Then
                    If InStr(1, m_saCurrentPageLine(i), "with the exact amount", vbTextCompare) > 0 Then
                        m_oRegExp.Pattern = "\bexact\b"
                        If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                            Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                            For Each oMatch In oMatches
                                Mid(saCurrentPageLineUnderline(i), oMatch.FirstIndex + 1) = String(oMatch.length, "_")
                            Next
                        End If
                    End If
                    
                    m_oRegExp.Pattern = "\bestimate\b"
                    If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                        Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                        For Each oMatch In oMatches
                            Mid(saCurrentPageLineUnderline(i), oMatch.FirstIndex + 1) = String(oMatch.length, "_")
                        Next
                    End If
                
                    saCurrentPageLineBold(i) = m_saCurrentPageLine(i)
                    m_saCurrentPageLine(i) = Space(Len(m_saCurrentPageLine(i)))
                End If
            Next
            
             
            Dim bInInterestRateSection As Boolean
            bInInterestRateSection = False
            For i = 0 To m_lCurrentPageIdx
                If bInInterestRateSection Then
                    If InStr(1, m_saCurrentPageLine(i), "  ------", vbTextCompare) = 1 Then
                        Exit For
                    End If
                    m_oRegExp.Pattern = "\bestimate\b"
                    If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                        Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                        For Each oMatch In oMatches
                            Mid(saCurrentPageLineUnderline(i), oMatch.FirstIndex + 1) = String(oMatch.length, "_")
                        Next
                    End If
                Else
                    bInInterestRateSection = InStr(1, m_saCurrentPageLine(i), "-- NEW INTEREST RATE", vbTextCompare) > 0
                End If
            Next
           
            'These items from the IF YOU ANTICIPATE PROBLEMS section get underlined.
            Const scUnderlineTest1Val As String = "  Refinance your loan "
            Const scUnderlineTest1Und As String = "  ___________________ "
            
            Const scUnderlineTest2Val As String = "  Modify your loan terms "
            Const scUnderlineTest2Und As String = "  ______________________ "
            
            Const scUnderlineTest3Val As String = "  Sell your home "
            Const scUnderlineTest3Und As String = "  ______________ "
            
            Const scUnderlineTest4Val As String = "  Payment forbearance "
            Const scUnderlineTest4Und As String = "  ___________________ "
            
            Const scBoldTest1Val As String = "  The following options may be possible and may be subject to lender"
            Const scBoldTest1Reg As String = "  The following options may be possible and     be subject to lender"
            Const scBoldTest1Bld As String = "                                            may                     "
            
            Const scUnderlineUrl1Val As String = "www.hud.gov/offices/hsg/sfh/hcc/hcs.cfm"
            Const scUnderlineUrl1Und As String = "_______________________________________"
            
            Const scUnderlineUrl2Val As String = "www.consumerfinance.gov/mortgagehelp"
            Const scUnderlineUrl2Und As String = "___________________________________"
            
            Dim bInNewPaymentGrid As Boolean
            bInNewPaymentGrid = False
            
            Const scGridTitle1Val As String = "                                             CURRENT RATE    ESTIMATED NEW RATE"
            Const scGridTitle1Reg As String = "                                                     RATE                  RATE"
            Const scGridTitle1Bld As String = "                                             CURRENT         ESTIMATED NEW     "
            
            Const scGridTitle2Val As String = "                                             CURRENT RATE              NEW RATE"
            Const scGridTitle2Reg As String = "                                                     RATE                  RATE"
            Const scGridTitle2Bld As String = "                                             CURRENT                   NEW     "
            
            Dim bTotalPymtFound As Boolean
            bTotalPymtFound = False
            
            For i = 0 To m_lCurrentPageIdx
                If bInNewPaymentGrid Then
                    If Not bTotalPymtFound Then
                        If InStr(1, m_saCurrentPageLine(i), "  TOTAL PAYMENT DUE ", vbTextCompare) = 1 Then
                            bTotalPymtFound = True
                        Else
'                            m_oRegExp.Pattern = "(^\s+Principal\s+)(\$0?.00)\b(.+)$"
'                            If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
'                                Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
'                                Set oMatch = oMatches.Item(0)
'                                m_saCurrentPageLine(i) = oMatch.SubMatches(0) & Space(Len(oMatch.SubMatches(1))) & oMatch.SubMatches(2)
'                                saCurrentPageLineBold(i) = Space(Len(oMatch.SubMatches(0))) & oMatch.SubMatches(1) & Space(Len(oMatch.SubMatches(2)))
'                            End If
                            m_oRegExp.Pattern = "^\s+Principal\s+\$[\d,.-]+\s+\$[\d,.-]+\s*$"
                            If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                                Const scZeroAmount As String = "$.00"
                                Dim lZeroPos As Integer
                                lZeroPos = InStr(1, m_saCurrentPageLine(i), scZeroAmount, vbTextCompare)
                                While (lZeroPos > 0)
                                    Mid(m_saCurrentPageLine(i), lZeroPos) = Space(Len(scZeroAmount))
                                    Mid(saCurrentPageLineBold(i), lZeroPos, Len(scZeroAmount)) = scZeroAmount
                                    lZeroPos = InStr(1, m_saCurrentPageLine(i), scZeroAmount, vbTextCompare)
                                Wend
                            End If
                        End If
                    End If
                    If bTotalPymtFound Then
                        If InStr(1, m_saCurrentPageLine(i), "  -----------------------------------------------------------------------------", vbTextCompare) = 1 Then
                            Exit For
                        End If
                        saCurrentPageLineBold(i) = m_saCurrentPageLine(i)
                        m_saCurrentPageLine(i) = Space(Len(m_saCurrentPageLine(i)))
                    End If
                Else
                    If InStr(1, m_saCurrentPageLine(i), scGridTitle1Val, vbTextCompare) = 1 Then
                        bInNewPaymentGrid = True
                        Mid(m_saCurrentPageLine(i), 1) = scGridTitle1Reg
                        Mid(saCurrentPageLineBold(i), 1) = scGridTitle1Bld
                    ElseIf InStr(1, m_saCurrentPageLine(i), scGridTitle2Val, vbTextCompare) = 1 Then
                        bInNewPaymentGrid = True
                        Mid(m_saCurrentPageLine(i), 1) = scGridTitle2Reg
                        Mid(saCurrentPageLineBold(i), 1) = scGridTitle2Bld
                    End If
                End If
            Next
            
            Dim bInAnticipateSection As Boolean
            bInAnticipateSection = False
            For i = 0 To m_lCurrentPageIdx
                If bInAnticipateSection Then
                    If InStr(1, m_saCurrentPageLine(i), scUnderlineTest1Val, vbTextCompare) = 1 Then
                        Mid(saCurrentPageLineUnderline(i), 1) = scUnderlineTest1Und
                    ElseIf InStr(1, m_saCurrentPageLine(i), scUnderlineTest2Val, vbTextCompare) = 1 Then
                        Mid(saCurrentPageLineUnderline(i), 1) = scUnderlineTest2Und
                    ElseIf InStr(1, m_saCurrentPageLine(i), scUnderlineTest3Val, vbTextCompare) = 1 Then
                        Mid(saCurrentPageLineUnderline(i), 1) = scUnderlineTest3Und
                    ElseIf InStr(1, m_saCurrentPageLine(i), scUnderlineTest4Val, vbTextCompare) = 1 Then
                        Mid(saCurrentPageLineUnderline(i), 1) = scUnderlineTest4Und
                    ElseIf InStr(1, m_saCurrentPageLine(i), scBoldTest1Val, vbTextCompare) = 1 Then
                        Mid(m_saCurrentPageLine(i), 1) = scBoldTest1Reg
                        Mid(saCurrentPageLineBold(i), 1) = scBoldTest1Bld
                    End If
                    
                    Dim lPos As Integer
                    lPos = InStr(1, m_saCurrentPageLine(i), scUnderlineUrl1Val, vbTextCompare)
                    If lPos > 0 Then
                        Mid(saCurrentPageLineUnderline(i), lPos) = scUnderlineUrl1Und
                    End If
                    
                    lPos = InStr(1, m_saCurrentPageLine(i), scUnderlineUrl2Val, vbTextCompare)
                    If lPos > 0 Then
                        Mid(saCurrentPageLineUnderline(i), lPos) = scUnderlineUrl2Und
                    End If
                Else
                    bInAnticipateSection = InStr(1, m_saCurrentPageLine(i), "- IF YOU ANTICIPATE PROBLEMS MAKING YOUR PAYMENTS -", vbTextCompare) > 0
                End If
            Next
            
            sngPsCurrPosY = 728
            
            Dim sSectionDescription As String
            
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineBold(i) = RTrim(saCurrentPageLineBold(i))
                saCurrentPageLineUnderline(i) = RTrim(saCurrentPageLineUnderline(i))
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                
                If IsPymtChangeGrid(i) Then
                    .WriteLine PsProc_PYMTCHGGRID(sngPsCurrPosY, sngLineSpacing)
                End If
                
                If IsSectionSep(i, sSectionDescription, m_saCurrentPageLine(i)) Then
                    .WriteLine PsProc_SECH(sngPsCurrPosY, sSectionDescription)
                Else
                    If Len(m_saCurrentPageLineBold(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                    End If
                    If Len(m_saCurrentPageLine(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
                    End If
                End If
            
                If Len(saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, saCurrentPageLineBold(i))
                End If
                
                If Len(saCurrentPageLineUnderline(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, saCurrentPageLineUnderline(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
         
    PrintPiece_LN0473 = True
    
    Exit Function
   
PrintPiece_LN0473_ErrorExit:
    
    PrintPiece_LN0473 = False
    Exit Function
   
PrintPiece_LN0473_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_LN0473()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_LN0473_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_LN0473_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_LN0473_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_LN0473_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_LN0473_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_LN0473_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_LN0473 = False
    Exit Function
    
End Function

Private Function PrintPiece_LN0473_152() As Boolean
    Dim sngPsCurrPosY As Single
    Const sngLineSpacing As Single = 11
    Dim bEndOfPiece As Boolean
    Dim oMatch As Match
    Dim oMatches As MatchCollection
    Dim i As Long
    
    Dim bDuplex As Boolean
    
    On Error GoTo PrintPiece_LN0473_152_Error
    
    m_lInputImages = 0
    m_lOutputSheets = 0
    m_lOutputImages = 0
    
    m_lInputSequence = m_lInputSequence + 1
    
    m_lLocation = m_oStmtBodyFile.Position
    
    m_oStmtBodyFile.WriteLine "%FiservBeginPiece:"
    
    Do
    
        m_lInputImages = m_lInputImages + 1
        
        With m_oStmtBodyFile

            
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
        
            If Not m_oFlatStockLogoFormInfo Is Nothing Then
                .WriteLine m_oFlatStockLogoFormInfo.ExecutionCommand
            End If
            
            If Not m_oReturnAddress Is Nothing Then
                .WriteLine PsProc_RETADDR(m_oReturnAddress)
            End If
            
            If Not m_oDeliveryAddress Is Nothing Then
                .WriteLine PsProc_DELADDR(m_oDeliveryAddress)
            End If
            
            If m_lOutputImages = 1 Then
                If Len(m_sOCR) > 0 Then
                    .WriteLine PsString(m_sOCR, True) & " " & CStr(m_sngOCR_X) & " " & CStr(m_sngOCR_Y) & " OCR"
                End If
                .WriteLine "%FiservMailPieceAddress: " & m_lXmlAddressPositionX & " " & m_lXmlAddressPositionY + 14 & " 10 6"
                .WriteLine "%FDS_DPBC: (*<%FDS_DPBC>*) PostNet"
                .WriteLine "%FDS_SEQ:(<%FDS_SEQ>) SEQ"
            End If
            
            .WriteLine PsString(m_sCurrentFileName & " " & m_sCurrentFileFormType, True) & " FRMCD"
            
            .WriteLine "gs"
            
            .WriteLine "66 0 translate"
            
            For i = m_lCurrentPageIdx To 0 Step -1 'Trim blanks from end of page
                If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
                    m_lCurrentPageIdx = i
                    Exit For
                End If
            Next
            
            Dim bInBoldPara As Boolean
            bInBoldPara = False
            Dim saCurrentPageLineBold() As String 'Added for doing bold sections of a line.
            ReDim saCurrentPageLineBold(m_lCurrentPageIdx)
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineBold(i) = Space(Len(m_saCurrentPageLine(i)))
            Next
            
            Dim saCurrentPageLineUnderline() As String 'Added for doing underline sections of a line.
            ReDim saCurrentPageLineUnderline(m_lCurrentPageIdx)
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineUnderline(i) = Space(Len(m_saCurrentPageLine(i)))
            Next
            
            For i = 0 To m_lCurrentPageIdx
                If bInBoldPara Then
                    If Trim(m_saCurrentPageLine(i)) = "|                                                                             |" Then
                        bInBoldPara = False
                    End If
                Else
                    If InStr(1, m_saCurrentPageLine(i), "**Under the terms of your Adjustable-Rate Mortgage", vbTextCompare) > 0 Then
                        bInBoldPara = True
                    End If
                End If
                
                If bInBoldPara Then
                    saCurrentPageLineBold(i) = Space(2) & Mid(m_saCurrentPageLine(i), 3, Len(m_saCurrentPageLine(i)) - 3) & Space(1)
                    m_saCurrentPageLine(i) = " |" & Space(Len(m_saCurrentPageLine(i)) - 3) & "|"
                End If
            Next
            
            Dim bInRateChangeSection As Boolean
            bInRateChangeSection = False
            For i = 0 To m_lCurrentPageIdx
                ' First check if we are entering or leaving the rate change section
                If bInRateChangeSection Then
                    If Trim(m_saCurrentPageLine(i)) = "|                                                                             |" Then
                        bInRateChangeSection = False
                    End If
                Else
                    If Trim(m_saCurrentPageLine(i)) = "| |-----------------------------|----------------------|--------------------| |" Then
                        bInRateChangeSection = True
                    End If
                End If
                
                ' If we are in the rate change section, selectively bold some text here and there
                If bInRateChangeSection Then
                    ' Bold current and new balance text in line
                    m_oRegExp.Pattern = "^( \| \|\s+\|         )(CURRENT)( RATE \|           )(NEW)( RATE \| \|)$"
                    If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                        Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                        Set oMatch = oMatches.Item(0)
                        m_saCurrentPageLine(i) = oMatch.SubMatches(0) & Space(Len(oMatch.SubMatches(1))) & oMatch.SubMatches(2) & Space(Len(oMatch.SubMatches(3))) & oMatch.SubMatches(4)
                        saCurrentPageLineBold(i) = Space(Len(oMatch.SubMatches(0))) & oMatch.SubMatches(1) & Space(Len(oMatch.SubMatches(2))) & oMatch.SubMatches(3) & Space(Len(oMatch.SubMatches(4)))
                    End If
                    
                    m_oRegExp.Pattern = "^( \| \|)(TOTAL PAYMENT DUE)(\s+\|\s+)(\$[\d\.,]+?)( \|\s+ )(\$[\d\.,]+?)( \| \|)$"
                    If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                        Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                        Set oMatch = oMatches.Item(0)
                        m_saCurrentPageLine(i) = oMatch.SubMatches(0) & Space(Len(oMatch.SubMatches(1))) & oMatch.SubMatches(2) & Space(Len(oMatch.SubMatches(3))) & oMatch.SubMatches(4) & Space(Len(oMatch.SubMatches(5))) & oMatch.SubMatches(6)
                        saCurrentPageLineBold(i) = Space(Len(oMatch.SubMatches(0))) & oMatch.SubMatches(1) & Space(Len(oMatch.SubMatches(2))) & oMatch.SubMatches(3) & Space(Len(oMatch.SubMatches(4))) & oMatch.SubMatches(5) & Space(Len(oMatch.SubMatches(6)))
                    End If
                    
                    m_oRegExp.Pattern = "^( \| \|)(MONTHLY)(\s+\|\s+\|\s+)(\(DUE\s+\d{1,2}-\d{2}-\d{4}\))(\s+\| \|)$"
                    If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                        Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                        Set oMatch = oMatches.Item(0)
                        m_saCurrentPageLine(i) = oMatch.SubMatches(0) & Space(Len(oMatch.SubMatches(1))) & oMatch.SubMatches(2) & Space(Len(oMatch.SubMatches(3))) & oMatch.SubMatches(4)
                        saCurrentPageLineBold(i) = Space(Len(oMatch.SubMatches(0))) & oMatch.SubMatches(1) & Space(Len(oMatch.SubMatches(2))) & oMatch.SubMatches(3) & Space(Len(oMatch.SubMatches(4)))
                    End If
                End If
            Next
           
            ' All regular expres
            m_oRegExp.Pattern = "^( \|-+ )(.+?)( -+\|)$"
            For i = 0 To m_lCurrentPageIdx
                If m_oRegExp.Test(m_saCurrentPageLine(i)) = True Then
                    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(i))
                    Set oMatch = oMatches.Item(0)
                    m_saCurrentPageLine(i) = oMatch.SubMatches(0) & Space(Len(oMatch.SubMatches(1))) & oMatch.SubMatches(2)
                    saCurrentPageLineBold(i) = Space(Len(oMatch.SubMatches(0))) & oMatch.SubMatches(1) & Space(Len(oMatch.SubMatches(2)))
                    saCurrentPageLineUnderline(i) = Space(Len(oMatch.SubMatches(0))) & Underscore(Len(oMatch.SubMatches(1))) & Space(Len(oMatch.SubMatches(2)))
                End If
            Next
            
            sngPsCurrPosY = 728
            
            Dim sSectionDescription As String
            
            For i = 0 To m_lCurrentPageIdx
                saCurrentPageLineBold(i) = RTrim(saCurrentPageLineBold(i))
                saCurrentPageLineUnderline(i) = RTrim(saCurrentPageLineUnderline(i))
            Next
            
            For i = 0 To m_lCurrentPageIdx
                sngPsCurrPosY = sngPsCurrPosY - sngLineSpacing
                
                If IsPymtChangeGrid(i) Then
                    .WriteLine PsProc_PYMTCHGGRID(sngPsCurrPosY, sngLineSpacing)
                End If
                
                If IsSectionSep(i, sSectionDescription, m_saCurrentPageLine(i)) Then
                    .WriteLine PsProc_SECH(sngPsCurrPosY, sSectionDescription)
                Else
                    If Len(m_saCurrentPageLineBold(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, m_saCurrentPageLineBold(i))
                    End If
                    If Len(m_saCurrentPageLine(i)) > 0 Then
                        m_oStmtBodyFile.WriteLine PsProc001_DTL(sngPsCurrPosY, m_saCurrentPageLine(i))
                    End If
                End If
            
                If Len(saCurrentPageLineBold(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, saCurrentPageLineBold(i))
                End If
                
                If Len(saCurrentPageLineUnderline(i)) > 0 Then
                    m_oStmtBodyFile.WriteLine PsProc001_DTLB(sngPsCurrPosY, saCurrentPageLineUnderline(i))
                End If
            Next
            
            If sngPsCurrPosY < 0 Then
               Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page overflow error - data may print off the bottom of the page."
            End If
        
            .WriteLine "gr"
            
        End With
        
        PsPageEnd
        
        If Not m_oPieceBacker Is Nothing Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            m_oStmtBodyFile.WriteLine m_oPieceBacker.ExecutionCommand
            PsPageEnd
        End If
        
    Loop Until GetTrailerPage_Indirect = False
        
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then
            PsPageBegin_FlatStock "FDSFlatStock", m_bDuplex
            PsPageEnd
        End If
    End If

    m_oStmtBodyFile.WriteLine "%FiservEndPiece:"
         
    PrintPiece_LN0473_152 = True
    
    Exit Function
   
PrintPiece_LN0473_152_ErrorExit:
    
    PrintPiece_LN0473_152 = False
    Exit Function
   
PrintPiece_LN0473_152_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.PrintPiece_LN0473_152()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oFileErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
        Case ErrorDispatchEnum.ErrorDispNotHandled 'Not a file error - use general error handler
            Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
               Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
                   Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
               Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
                   On Error GoTo PrintPiece_LN0473_152_Error
                   Resume
               Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
                   On Error GoTo PrintPiece_LN0473_152_Error
                   Resume Next
               Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
                   GoTo PrintPiece_LN0473_152_ErrorExit
                   Exit Function
            End Select
        Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
            On Error GoTo PrintPiece_LN0473_152_Error
            Resume
        Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
            On Error GoTo PrintPiece_LN0473_152_Error
            Resume Next
        Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
            GoTo PrintPiece_LN0473_152_ErrorExit
            Exit Function
    End Select
     
    On Error Resume Next
    PrintPiece_LN0473_152 = False
    Exit Function
    
End Function

Private Function IsPymtChangeGrid(ByVal lIndex As Long) As Boolean
    IsPymtChangeGrid = False
    Const scVertBar As String = "  -----------------------------------------------------------------------------"
    'Const scHeading1 As String = "                                             CURRENT RATE    ESTIMATED NEW RATE"
    Const scHeading2 As String = "                                          AND PAYMENT DUE       AND PAYMENT DUE"
    'Const scHeading3 As String = "                                                  MONTHLY               MONTHLY"
    
    If lIndex + 13 > m_lCurrentPageIdx Then
        Exit Function
    End If
    If m_saCurrentPageLine(lIndex) = scVertBar Then
    'If m_saCurrentPageLine(lIndex + 1) = scHeading1 Then
    If m_saCurrentPageLine(lIndex + 2) = scHeading2 Then
    'If m_saCurrentPageLine(lIndex + 3) = scHeading3 Then
    If m_saCurrentPageLine(lIndex + 4) = scVertBar Then
    If m_saCurrentPageLine(lIndex + 6) = scVertBar Then
    If m_saCurrentPageLine(lIndex + 10) = scVertBar Then
    If m_saCurrentPageLine(lIndex + 13) = scVertBar Then
        IsPymtChangeGrid = True
        m_saCurrentPageLine(lIndex) = vbNullString
        m_saCurrentPageLine(lIndex + 4) = vbNullString
        m_saCurrentPageLine(lIndex + 6) = vbNullString
        m_saCurrentPageLine(lIndex + 10) = vbNullString
        m_saCurrentPageLine(lIndex + 13) = vbNullString
    End If
    End If
    End If
    End If
    'End If
    End If
    'End If
    End If
End Function

Private Function GetTrailerPage_Indirect() As Boolean
    
    If m_lReadAheadPageIdx < 0 Then
        Exit Function
    End If

    If Left(m_saReadAheadPageLine(0), Len(m_scFileName)) = m_scFileName Then
        GetTrailerPage_Indirect = False
        Exit Function
    End If
    
    Select Case m_sCurrentFileName
    
        Case "LN395Q1"
            If m_sCurrentFileFormType = "CLET" Then
                GetTrailerPage_Indirect = False
            Else
                GetTrailerPage_Indirect = GetPage_LN395Q1Trailer
            End If
            Exit Function
    
        Case "LN0472P1"
            'These are always one-pagers
            GetTrailerPage_Indirect = False
            Exit Function
    
        Case "LN0473P1"
            If m_sXmlInst = "Investors" Then
                GetTrailerPage_Indirect = GetPage_LN0473P1Trailer(True)
            Else
                GetTrailerPage_Indirect = GetPage_LN0473P1Trailer
            End If
            Exit Function
    
        Case "LN0860P1"
            If m_sXmlInst = "Investors" Then
                GetTrailerPage_Indirect = GetPage_LN0860P1Trailer(True, False)
            Else
                GetTrailerPage_Indirect = GetPage_LN0860P1Trailer
            End If
            Exit Function
            
'        Case "PRINT1"
'            If m_sXmlInst = "Investors" Then
'                GetTrailerPage_Indirect = GetPage_PRINT1Trailer(True, False)
'            Else
'                GetTrailerPage_Indirect = GetPage_PRINT1Trailer
'            End If
'            Exit Function
    
        Case "LN0480Q1", "E_QSYSPRT"
            GetTrailerPage_Indirect = GetPage_LN93357PTrailer
            Exit Function
    
        Case "LN93357P1"
            GetTrailerPage_Indirect = GetPage_LN93357PTrailer
            Exit Function
    
        Case "LN93357P2"
            GetTrailerPage_Indirect = GetPage_LN93357PTrailer
            Exit Function
    
        Case "LN93357P3"
            GetTrailerPage_Indirect = GetPage_LN93357PTrailer
            Exit Function
    
        Case "RA0080P1"
            GetTrailerPage_Indirect = GetPage_RA0080P1Trailer
            Exit Function
    
        Case "TA0050P1"
            'These are always one-pagers
            GetTrailerPage_Indirect = False
            Exit Function
             
        Case "TA9544P1"
            GetTrailerPage_Indirect = GetPage_TA9544PTrailer
            Exit Function
             
        Case "TA9544P2"
            GetTrailerPage_Indirect = GetPage_TA9544PTrailer
            Exit Function
            
            'CPTPK-1472
        Case "TA9544P3"
            GetTrailerPage_Indirect = GetPage_TA9544PTrailer
            Exit Function
             'CPTPK-1472
             
        Case "TA93357P3", "TA93357P2"
            GetTrailerPage_Indirect = GetPage_TA93357PTrailer
            Exit Function
             
        Case "T152Q2"
            GetTrailerPage_Indirect = GetPage_T152Q2Trailer
            Exit Function
            
        Case "CH0202P4"
            GetTrailerPage_Indirect = GetPage_NoticeTrailer
            Exit Function
            
        Case "TA990770P1"
            GetTrailerPage_Indirect = GetPage_TA990771PTrailer
            If GetTrailerPage_Indirect = True And m_sXmlInst = "SNB" Then
                ReformatPage_TA990771
            End If
            Exit Function
            
        Case "TA990771P1"
            GetTrailerPage_Indirect = GetPage_TA990771PTrailer
            If GetTrailerPage_Indirect = True And m_sXmlInst = "SNB" Then
                ReformatPage_TA990771
            End If
            Exit Function
            
        Case "TM0330P1"
            GetTrailerPage_Indirect = GetPage_TM0330P1Trailer
'            If m_sXmlInst = "ElDorado" Or m_sXmlInst = "Nationwide" Or m_sXmlInst = "SNB" Then 'REM RM12761
            If m_sXmlInst = "ElDorado" Or m_sXmlInst = "Nationwide" Or m_sXmlInst = "SNB" Or m_sXmlInst = "BPB" Then 'RM12761
                If GetTrailerPage_Indirect = True Then
                    ReformatTM0330P1
                End If
            End If
            Exit Function
            
        Case "CH0202P1"
            GetTrailerPage_Indirect = GetPage_CH0202P1Trailer
            Exit Function
            
        Case "TM0851P1"
            GetTrailerPage_Indirect = False
            ' Turn off duplexing if Investors TM0851P1
            If m_sXmlInst = "Investors" Then
                m_bDuplex = False
            End If
            Exit Function
            
        Case Else
            GetTrailerPage_Indirect = False
            Exit Function
            
  End Select
    
End Function

Private Function GetPage_Indirect() As Boolean
    Dim oStack As CStack
    
    On Error GoTo GetPage_Indirect_Error
    
    If ReadPage() = False Then
        GetPage_Indirect = False
        Exit Function
    End If
    
    m_sOMR = vbNullString
    
    ' Institution settings
    Select Case m_sXmlInst
    
        Case "BPB"
            Select Case m_sCurrentFileName
                
                'RM13669
                Case "CU1610P1"
                    GetPage_Notice 1, 12
                    m_lStackItems = StackItemsPressureSeal
                    m_sStockTray = m_scPaperTypePressureSeal
                'END RM13669
                    
                Case "LN395Q1"
                    If m_sCurrentFileFormType = "CLET" Then
                        GetPage_LN395Q1CLET
                        m_lStackItems = StackItemsPressureSeal
                    Else
                        GetPage_LN395Q1
                        m_lStackItems = StackItemsFlatStock
                        ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                        ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                        
                        ' 09/04/2012 - All Boston Private Bank stock and envelopes
                        m_sStockTray = m_scStockTrayBPBLetterhead
                        m_sEnvelopeCatCode = m_scEnvelopeBPB
                    End If
                    
                Case "LN0458P1"
                    GetPage_Notice 10, 13
                    m_lStackItems = StackItemsPressureSeal
                    
                Case "LN0470P1"
                    GetPage_Notice 9, 13, False, True, True
                    Call Reformat_BPB_LN0470
                    m_lStackItems = StackItemsPressureSeal
                    
                Case "LN0470P2"
                    GetPage_Notice 9, 13, False, False, True
                    Call Reformat_BPB_LN0470
                    m_lStackItems = StackItemsPressureSeal
                    
                Case "LN0472P1"
                    GetPage_LN0472P1
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    
                Case "LN0473P1"
                    GetPage_LN0473P1
                    'm_lStackItems = StackItemsFlatStock
                    m_lStackItems = StackItemsLN0473_152
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    If m_bXML_Archive Then
                        Set m_oFlatStockLogoFormInfo = m_oBPP102Form
                    End If
                    
                Case "LN0480Q1"
                    GetPage_LN93357P
                    m_lStackItems = StackItemsFlatStock
                    
                    ' m_sStockTray = m_scStockTrayBPB701CreditLine
                    ' m_sEnvelopeCatCode = m_scEnvelopeBorel
                    
                    m_sStockTray = m_scStockTrayBPB201CreditLine
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9

                    m_sRegion = "002"
                    
                Case "LN0860P1"
                    GetPage_LN0860P1
                    m_lStackItems = StackItemsFlatStockLegal
                    ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBPB603, m_scStockTrayBPB103)
                    ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
             
                    m_sStockTray = m_scStockTrayBPB103
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
             
                Case "LN93357P1"
                    GetPage_LN93357P
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPB201CreditLine
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9
                    
'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If
            
                Case "LN93357P2"
                    GetPage_LN93357P
                    m_lStackItems = StackItemsFlatStock
                    
                    ' m_sStockTray = m_scStockTrayBPB701CreditLine
                    ' m_sEnvelopeCatCode = m_scEnvelopeBorel
                    
                    m_sStockTray = m_scStockTrayBPB201CreditLine
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9
                    
'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If
                
                Case "LN93357P3"
                    GetPage_LN93357P
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPB201CreditLine
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9
                    
'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If
            
                Case "RA0080P1"
                    GetPage_RA0080P1
                    m_bReturnToBank = False
                    m_lStackItems = StackItemsFlatStock
                    m_bDuplex = True
                    'Set m_oPieceBacker = m_oRABacker
                    
                    ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                    ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                   
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                   
                Case "SD0095P2"
                    GetPage_Notice 16, 16
                    'm_bReturnToBank = True
                    
                    m_lStackItems = StackItemsNotice
                    If m_sRegion = "004" Then
                        m_bNoPrint = True
                    Else
                        m_sStockTray = m_scStockTrayBPBLetterhead
                        m_sEnvelopeCatCode = m_scEnvelopeBPB
                        m_sEnclosureCatCode = m_scEnclosureBPBER9
                    End If
                    
                Case "SD0095P3"
                    GetPage_Notice 16, 16
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsNotice
                    
                    ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                    ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                    ' m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                    
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBER9
                    
'                This needs to be tested.
                Case "SD0095P4"
                    GetPage_Notice 16, 15, False, False, True
                    m_lStackItems = StackItemsPressureSeal
             
'                Case "T152Q2"
'                    GetPage_TA93357P
'                    m_bReturnToBank = True
'                    m_lStackItems = StackItemsFlatStock
'                    m_sStockTray = "TopPerf"
''                    m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBPB701CreditLine, m_scStockTrayBPB201CreditLine)
''                    m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
''                    m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                    
                Case "TA0050P1"
                    GetPage_TA0050P1
                    m_bReturnToBank = False
                    m_lStackItems = StackItemsFlatStock
                    
                    ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                    ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                    ' m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                    
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                    
                Case "TA0180P1"
                    GetPage_Notice 1, 13 + IIf(Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And InStr(1, m_saCurrentPageLine(1), "***NOTICE OF ") > 0, 1, 0)  'bug fix for Signature release 17.2
                    m_lStackItems = StackItemsPressureSeal
                    'm_bReturnToBank = True
                    'm_lStackItems = StackItemsNotice
'                    m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
'                    m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                    
                Case "TA0185P1"
                    GetPage_Notice 6, 15
                    m_lStackItems = StackItemsPressureSeal
                    
                Case "TA0190P1"
                    GetPage_Notice 6, 13
                    m_lStackItems = StackItemsPressureSeal
'                    m_bReturnToBank = True
'                    m_lStackItems = StackItemsNotice
'                    m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
'                    m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                
                Case "TA0441P1"
                    GetPage_Notice 9, 14, True, False, True
                    Dim bBankByMail As Boolean
                    bBankByMail = False
                    If m_lCurrentPageIdx >= 10 Then
                        If InStr(1, m_saCurrentPageLine(10), "BANK-BY-MAIL DEPOSIT", vbTextCompare) > 0 Then
                            bBankByMail = True
                        End If
                    End If
                    m_lStackItems = StackItemsNotice
                    
                    ' m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                    ' m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                    ' If bBankByMail Then
                    '     m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                    ' End If
                
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    If bBankByMail Then
                        m_sEnclosureCatCode = m_scEnclosureBPBER9
                    End If
                
                
'                Case "TA0441P2"
'                    'ToDo: Test
'                    GetPage_Notice 9, 14, True, False, True
'                    m_lStackItems = StackItemsNotice
'                    m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
'                    m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
'                    m_sEnclosureCatCode = IIf(m_sRegion = "002", m_scEnclosureBPBBORELER9, m_scEnclosureBPBER9)
                
                Case "TA1161P1"
                    GetPage_Notice 11, 16
                    m_lStackItems = StackItemsPressureSeal
                    
                Case "TA1171P1"
                    GetPage_Notice 11, 20
                    m_lStackItems = StackItemsPressureSeal
             
                Case "TA9544P1"
                    GetPage_TA9544P
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsFlatStock8LPI
'                    'REM CPTPK-1712
'                    m_sStockTray = m_scStockTrayBPB101
'                    'END REM CPTPK-1712
                    'CPTPK-1712
                    m_sStockTray = m_scStockTrayBPBpex
                    'End CPTPK-1712
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9

'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If
             
                Case "TA9544P2"
                    GetPage_TA9544P
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsFlatStock8LPI

                    ' m_sStockTray = m_scStockTrayBPB601
'                    'REM CPTPK-1712
'                    m_sStockTray = m_scStockTrayBPB101
'                    'END REM CPTPK-1712
                    'CPTPK-1712
                    m_sStockTray = m_scStockTrayBPBpex
                    'End CPTPK-1712
                    ' m_sEnvelopeCatCode = m_scEnvelopeBorel
                    m_sEnvelopeCatCode = m_scEnvelopeBPB

                    m_sEnclosureCatCode = m_scEnclosureBPBEL9
                    
'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If
             
                Case "TA9544P3"
                    GetPage_TA9544P
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsFlatStock8LPI
'                    'REM CPTPK-1712
'                    m_sStockTray = m_scStockTrayBPB101
'                    'End REM CPTPK-1712
                    'CPTPK-1712
                    m_sStockTray = m_scStockTrayBPBpex
                    'End CPTPK-1712
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    m_sEnclosureCatCode = m_scEnclosureBPBEL9

'                    If Len(m_sAccountNumber) > 0 Then
'                        m_sOCR = OcrLine(m_sAccountNumber, m_cPaymentAmount)
'                    End If

                Case "TA93357P2"
                    GetPage_TA93357P
                
''REM RM11759
'                    'm_bReturnToBank = True
'                    m_lStackItems = StackItemsFlatStock8LPI
'                    m_sStockTray = m_scStockTrayBPB101
'                    m_sEnvelopeCatCode = m_scEnvelopeBPB
'                    m_sEnclosureCatCode = m_scEnclosureBPBEL9
''END REM RM11759

'RM11759
                    If Mid(m_saCurrentPageLine(23), 75, 5) = " 771 " Then
                        m_bNoPrint = True
                        Do While GetTrailerPage_Indirect = True
                            m_lCurrentFileImages = m_lCurrentFileImages + 1
                        Loop

                        m_lStackItems = StackItemsBypass
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    Else
                        'm_bReturnToBank = True
                        m_lStackItems = StackItemsFlatStock8LPI
                        m_sStockTray = m_scStockTrayBPB101
                        m_sEnvelopeCatCode = m_scEnvelopeBPB
                        m_sEnclosureCatCode = m_scEnclosureBPBEL9
                    End If
'END RM11759

                Case "TM0360P2"
                    'GetPage_Notice 9, 14
                    'm_lStackItems = StackItemsPressureSeal
                    GetPage_FlatStock 9, 14
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    'Set m_oPieceBacker = m_oBPBBackers(1)
                    m_bDuplex = True

                    
                Case "TM0360P4"
                    'GetPage_Notice 9, 13
                    'm_lStackItems = StackItemsPressureSeal
                    GetPage_FlatStock 9, 13
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    Set m_oPieceBacker = m_oBPBBackers(1)
                    m_bDuplex = True
                    
                Case "TM0360P5"
                    'GetPage_Notice 9, 15
                    'm_lStackItems = StackItemsPressureSeal
                    GetPage_FlatStock 9, 15
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    Set m_oPieceBacker = m_oBPBBackers(1)
                    m_bDuplex = True
                    
                Case "TM0360P7"
                    'GetPage_Notice 9, 14
                    'm_lStackItems = StackItemsPressureSeal
                    GetPage_FlatStock 9, 14
                    If m_sRegion = "004" Then
                        m_bNoPrint = True
                    End If
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    'Set m_oPieceBacker = m_oBPBBackers(1)
                    m_bDuplex = True
                    
                Case "TM0360P8"
                    'GetPage_Notice 9, 15
                    'm_lStackItems = StackItemsPressureSeal
                    GetPage_FlatStock 9, 15
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                    'Set m_oPieceBacker = m_oBPBBackers(1)
                    m_bDuplex = True
                    
                Case "CH0202P1"
                    ' Added 03/30/2012 - DB
                    GetPage_Notice 5, 15
                    ' m_lStackItems = StackItemsPressureSeal
                    ' 10/23/2012 - Changed output to BPB-102 - DB
                    m_lStackItems = StackItemsNotice
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                
                Case "CH0202P4"
                    ' Added 03/30/2012 - DB
                    GetPage_Notice 5, 15
                    ' m_lStackItems = StackItemsPressureSeal
                    ' 10/23/2012 - Changed output to BPB-102 - DB
                    'm_lStackItems = StackItemsNotice
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = m_scStockTrayBPBLetterhead
                    m_sEnvelopeCatCode = m_scEnvelopeBPB
                
                Case "TA0420P1"
                    ' Added 03/30/2012 - DB
                    GetPage_Notice 9, 13
                    m_lStackItems = StackItemsPressureSeal
                
                Case "LN0465P1"
                    'Suppress these pieces.
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    If m_lCurrentNoPrintStack = 0 Then
                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                    End If
                    
                Case "LN0460P1"
                    ' Added 03/30/2012 - DB
                    'Suppress these pieces.
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    If m_lCurrentNoPrintStack = 0 Then
                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                    End If
                
                Case "TM0847P1"
                    GetPage_TM0847P1
                    ' Trim the leading blank elements from the page array
                    m_lCurrentPageIdx = TrimLeadingBlanks(m_saCurrentPageLine)
                    m_lStackItems = StackItemsPressureSeal
                
                Case "TM0330P1"
                    ' Added 12/28/2012 - DB
                    GetPage_TM0330P1 7, 8, True
                    m_lStackItems = StackItemsFlatStock
                    
                    ' #42 envelope setup
                    m_sStockTray = m_scStockTrayPlainWhite20
                    m_sEnvelopeCatCode = m_scEnvelope42
                    Set m_oFlatStockLogoFormInfo = m_oBPBFlatStockLogo
                
                Case "TI0530P1"
                    ' Added 12/28/2012 - DB
                    GetPage_Notice 7, 8
                    m_lStackItems = StackItemsPressureSeal
                
                
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
            End Select
            
            Set m_oExteriorLogoFormInfo = Nothing
            Set m_oInteriorLogoFormInfo = Nothing
            
            If m_sStockTray = m_scPaperTypePressureSeal Then
            
                Select Case m_sRegion
                    Case "", "001", "002", "003", "004"
                        ' Added region codes 003 and 004 which process the same as Boston Private Bank - 03/21/2012 - DB
                        Set m_oExteriorLogoFormInfo = m_oBPPExteriorLogo(0)
                        Set m_oInteriorLogoFormInfo = m_oBPBInteriorLogo(0)
                        
                    ' 09/04/2012 - All pressure seal pieces get the Boston Private logo
'                    Case "002"
'                        Set m_oExteriorLogoFormInfo = m_oBPPExteriorLogo(1)
'                        Set m_oInteriorLogoFormInfo = m_oBPBInteriorLogo(1)
                    Case Else
                        Stop
                End Select
                
            End If
            
            Select Case m_sEnclosureCatCode
                Case ""
                Case m_scEnclosureBPBEL9
                    m_sOMR = "Y"
                Case m_scEnclosureBPBER9
                    m_sOMR = "NY"
                Case m_scEnclosureBPBBORELER9
                    m_sOMR = "NNY"

                Case Else
                    Stop
            End Select

            
        Case "SNB" ' Sterling National Bank
            Dim lCurIdx As Long
                Const scRe As String = "RE:  Loan "
                Const scIf As String = "IF"
                Dim bBypassNotice As Boolean
            Select Case m_sCurrentFileName
                
                Case "E_CH0202P1"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        If m_sCurrentFileFormType = "NTCE" Then
                            'GetPage_NoticeScrape 11, 11
                            GetPage_CH0202P
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            ' No programming in place to handle this form type
                            If Not m_bCurrentFileIsUnrecognized Then
                                If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                                End If
                                m_bCurrentFileIsUnrecognized = True
                                If m_lCurrentNoPrintStack = 0 Then
                                    m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                    m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                    SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                                End If
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                    Else
                        ' No programming in place to handle this file type
                        If Not m_bCurrentFileIsUnrecognized Then
                            If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                            End If
                            m_bCurrentFileIsUnrecognized = True
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                            End If
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                
                ' Approved
                Case "LN395Q1"
                    ' Suppress form type HIST from being printed
                    If m_sCurrentFileFormType = "HIST" Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' If we are running the archive job - suppress these from being archived
                        If m_bXML_Archive = True Then
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        Else
                            ' Calculate where address block scrape should begin.  This handles situations where not all address lines
                            ' are 3 lines AND deals with inconsistent starting location of delivery address blocks on page
                            For lCurIdx = m_lCurrentPageIdx To 0 Step -1
                                If Len(Trim(m_saCurrentPageLine(lCurIdx))) = 0 Then
                                    Exit For
                                End If
                            Next
                            
                            'Ignore bogus notice that doesn't have a mailing address.   Repeated attempts to get these
                            'shut off have failed.

                            bBypassNotice = False
                            
                            ' This checks for the "FIRST REQUEST" format
                            If m_lCurrentPageIdx >= 1 Then
                                If Left(m_saCurrentPageLine(1), Len(scRe)) = scRe Then
                                    If Right(m_saCurrentPageLine(1), Len(scIf)) = scIf Then
                                        GetPage_NoPrint
                                        m_lStackItems = StackItemsPressureSeal
                                        m_bNoPrint = True
                                        bBypassNotice = True
                                    End If
                                End If
                            End If
                            
                            ' This checks for the "SECOND REQUEST" format
                            If m_lCurrentPageIdx >= 3 And Not bBypassNotice Then
                                If Left(m_saCurrentPageLine(3), Len(scRe)) = scRe Then
                                    If Right(m_saCurrentPageLine(1), Len(scIf)) = scIf Then
                                        GetPage_NoPrint
                                        m_lStackItems = StackItemsPressureSeal
                                        m_bNoPrint = True
                                        bBypassNotice = True
                                    End If
                                End If
                            End If
                            
                            If Not bBypassNotice Then
                                If InStr(1, m_saCurrentPageLine(1), "**LOAN NOTICE**") > 0 Then
                                    ' New format that has not been encountered before - scrape delivery address from different location
                                    GetPage_Notice 6, 13, False, False, False
                                    m_lStackItems = StackItemsPressureSeal
                                Else
                                    GetPage_Notice 6, lCurIdx + 2, False, False, True
                                    'GetPage_Notice 6, 12, False, False, False
                                    m_lStackItems = StackItemsPressureSeal
                                End If
                            End If
                        End If
                    End If
                Case "E_LN395Q1"
                    ' Suppress form type HIST from being printed
                    If m_sCurrentFileFormType = "HIST" Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' If we are running the archive job - suppress these from being archived
                        If m_bXML_Archive = True Then
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        Else
                            ' Calculate where address block scrape should begin.  This handles situations where not all address lines
                            ' are 3 lines AND deals with inconsistent starting location of delivery address blocks on page
                            
                            For lCurIdx = m_lCurrentPageIdx To 0 Step -1
                                If Len(Trim(m_saCurrentPageLine(lCurIdx))) = 0 Then
                                    Exit For
                                End If
                            Next
                            
                            'Ignore bogus notice that doesn't have a mailing address.   Repeated attempts to get these
                            'shut off have failed.

                            bBypassNotice = False
                            
                            ' This checks for the "FIRST REQUEST" format
                            If m_lCurrentPageIdx >= 1 Then
                                If Left(m_saCurrentPageLine(1), Len(scRe)) = scRe Then
                                    If Right(m_saCurrentPageLine(1), Len(scIf)) = scIf Then
                                        GetPage_NoPrint
                                        m_lStackItems = StackItemsPressureSeal
                                        m_bNoPrint = True
                                        bBypassNotice = True
                                    End If
                                End If
                            End If
                            
                            ' This checks for the "SECOND REQUEST" format
                            If m_lCurrentPageIdx >= 3 And Not bBypassNotice Then
                                If Left(m_saCurrentPageLine(3), Len(scRe)) = scRe Then
                                    If Right(m_saCurrentPageLine(1), Len(scIf)) = scIf Then
                                        GetPage_NoPrint
                                        m_lStackItems = StackItemsPressureSeal
                                        m_bNoPrint = True
                                        bBypassNotice = True
                                    End If
                                End If
                            End If
                            
                            If Not bBypassNotice Then
                                If InStr(1, m_saCurrentPageLine(1), "**LOAN NOTICE**") > 0 Then
                                    ' New format that has not been encountered before - scrape delivery address from different location
                                    GetPage_Notice 6, 13, False, False, False
                                    If Len(m_sXML_EddJobKey) > 0 Then
                                        BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                    End If
                                    m_lStackItems = StackItemsPressureSeal
                                Else
                                    GetPage_Notice 6, lCurIdx + 2, False, False, True
                                    'GetPage_Notice 6, 12, False, False, False
                                    If Len(m_sXML_EddJobKey) > 0 Then
                                        BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                    End If
                                    m_lStackItems = StackItemsPressureSeal
                                End If
                            End If
                        End If
                    End If

                   
                ' Approved
                Case "SD0095P2"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 16, 16
                      '  m_lStackItems = StackItemsPressureSeal
                    End If
                       m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                     Set m_oPieceBacker = m_oSNBBackers(2)
                     m_bDuplex = True
                ' Approved
                Case "SD0095P3"
                      
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 16, 16
                    '    m_lStackItems = StackItemsPressureSeal
                    End If
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                     Set m_oPieceBacker = m_oSNBBackers(2)
                     m_bDuplex = True
                  
                
                ' Approved
                Case "SD0095P4"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 16, 15, False, False, True
                        m_lStackItems = StackItemsPressureSeal
                    End If
             
                'RM5096
                Case "LN0470P1"
                    GetPage_NoticeScrape 9, 13, True, False, True
                    If m_bXML_Archive = True Then
                        BuildSNBPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                    End If
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                    Set m_oPieceBacker = m_oSNBBackers(0)
                    m_bDuplex = True
                
                Case "E_LN0470P1"
                    GetPage_NoticeScrape 9, 13, True, False, True
                    If m_bXML_Archive = True Then
                        BuildSNBPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                    End If
                    If Len(m_sXML_EddJobKey) > 0 Then
                        BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                    End If
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                    Set m_oPieceBacker = m_oSNBBackers(0)
                    m_bDuplex = True
                    
                 ' These run through PrimeDoc: 73100500\73100500_MixedDocs
'                ' Approved
'                Case "T152Q2"
'                    ' If we are running the archive job - suppress these from being archived
'                    If m_bXML_Archive = True Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_T152Q2
'                        m_lStackItems = StackItemsFlatStock428LPI
'                        m_sStockTray = m_scPaperTypeBottomPerf
'                        Set m_oPieceBacker = m_oSNBBackers(0)
'                        m_bDuplex = True
'                    End If
                    
'                Case "TA0050P1"
'                    GetPage_TA0050P1
'                    m_lStackItems = StackItemsFlatStock
'                    m_sStockTray = m_scStockTrayPlainWhite20
                    
                ' Approved
                Case "TA0180P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 1, 13 + IIf(Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And InStr(1, m_saCurrentPageLine(1), "***NOTICE OF ") > 0, 1, 0)  'bug fix for Signature release 17.2
                        m_lStackItems = StackItemsPressureSeal
                    End If

                    'm_bReturnToBank = True
                    'm_lStackItems = StackItemsNotice
                    'm_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
                    'm_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                Case "E_TA0180P1"
                        If Len(m_sXML_EddJobKey) > 0 Then
                            GetPage_Notice 1, 13 + IIf(Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And InStr(1, m_saCurrentPageLine(1), "***NOTICE OF ") > 0, 1, 0)  'bug fix for Signature release 17.2
                            BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsPressureSeal
                       
                        
                    ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TA0180P
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType10and20
                    ElseIf m_bXML_Archive = True Then
                        ' Archive job - do not process for archive, only do for CH0202P1
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' Print job - suppress from print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                Case "TA0185P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        'GetPage_Notice 6, 14
                        GetPage_TA0185P1 True
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
                Case "E_TA0185P1"
                    If Len(m_sXML_EddJobKey) > 0 Then
                            GetPage_TA0185P1 True
                            BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsPressureSeal
                        
                    ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TA0185P1
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType10and20
                    ElseIf m_bXML_Archive = True Then
                        ' Only archive CH0202P1 for this client
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' E Notice only - do not print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                Case "TA0183P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 1, 19
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
'                Case "TA0190P1"
'                    GetPage_Notice 6, 14
'                    m_lStackItems = StackItemsPressureSeal
                
                Case "E_TA0190P1"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        If m_sCurrentFileFormType = "CLET" Then
                            'GetPage_NoticeScrape 11, 11
                            GetPage_TA0190P
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            ' No programming in place to handle this form type
                            If Not m_bCurrentFileIsUnrecognized Then
                                If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                                End If
                                m_bCurrentFileIsUnrecognized = True
                                If m_lCurrentNoPrintStack = 0 Then
                                    m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                    m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                    SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                                End If
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                    Else
                        ' No programming in place to handle this file type
                        If Not m_bCurrentFileIsUnrecognized Then
                            If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                            End If
                            m_bCurrentFileIsUnrecognized = True
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                            End If
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                                
                Case "E_QSYSPRT"
                
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_LN0480
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType50
                    ElseIf m_bXML_Archive = True Then
                        ' Only archive CH0202P1 for this client
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' E Notice only - do not print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                Case "CU1610P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 1, 12
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                    End If
                
                ' Approved
                Case "TA0434P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 6, 4
                        m_lStackItems = StackItemsPressureSeal
                    End If
'                    m_bReturnToBank = True
'                    m_lStackItems = StackItemsNotice
'                    m_sStockTray = IIf(m_sRegion = "002", m_scStockTrayBorelLetterhead, m_scStockTrayBPBLetterhead)
'                    m_sEnvelopeCatCode = IIf(m_sRegion = "002", m_scEnvelopeBorel, m_scEnvelopeBPB)
                
                ' Approved
                Case "TA0441P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 13, True, False, True
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
                ' Approved
                Case "TA990771P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA990771P
                        ReformatPage_TA990771
                        m_lStackItems = StackItemsFlatStock428LPI
                        m_sStockTray = m_scStockTrayPlainWhite20
                        'RM10773
                        Select Case m_sAccountNumber
                            Case "7500000331", "7500000380", "7500000792", "7500001256", "7500001431", "7500001470", "7500001869", "7500001932", "7500003121", "7500003144", "7500003255", "7500003260", "7500003291"
                                m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType & " - Westchester Property Management Group")
                                m_bReturnToBank = True
                            Case Else
                                m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
                        End Select
                        'END RM10773
                    End If
                    
                ' Approved
                Case "TA990770P1"
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA990771P
                        ReformatPage_TA990771
                        m_lStackItems = StackItemsFlatStock428LPI
                        m_sStockTray = m_scStockTrayPlainWhite20
                        'RM10773
                        Select Case m_sAccountNumber
                            Case "7500000331", "7500000380", "7500000792", "7500001256", "7500001431", "7500001470", "7500001869", "7500001932", "7500003121", "7500003144", "7500003255", "7500003260", "7500003291"
                                m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType & " - Westchester Property Management Group", "BULK-GRPA")
                                m_bReturnToBank = True
                            Case Else
                                m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
                        End Select
                        'END RM10773
                    End If
                
                ' Approved
                Case "TA1161P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 11, 16, False, False, True
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
                ' Approved
                Case "TM0360P2"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 14
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
                ' Approved
                Case "TM0360P4"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 13
                        m_lStackItems = StackItemsPressureSeal
                    End If
                Case "E_TM0360P4"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TM0360P
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType30
                    ElseIf m_bXML_Archive = True Then
                        ' Only archive CH0202P1 for this client
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' E Notice only - do not print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                Case "TM0360P5"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 15
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                    End If
                Case "E_TM0360P5"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 15
                        If Len(m_sXML_EddJobKey) > 0 Then
                               BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        End If
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                    End If
                
                ' Approved
                Case "TM0360P7"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 14
                        m_sStockTray = m_scPaperTypePlain
                        m_lStackItems = StackItemsFlatStock42
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                    End If
                Case "E_TM0360P7"
                    If Len(m_sXML_EddJobKey) > 0 Then
                            GetPage_Notice 9, 14
                            BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_sStockTray = m_scPaperTypePlain
                            m_lStackItems = StackItemsFlatStock42
                            Set m_oPieceBacker = m_oSNBBackers(1)
                            m_bDuplex = True
                        
                    ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TM0360P
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                        m_lEddEnoticeType = EnoticeType30
                        Set m_oReturnAddress = Nothing
                    ElseIf m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                 Case "TM0360P8"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        'GetPage_Notice 9, 15 'REM RM12831
                        GetPage_Notice 9, 14 'RM12831
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                    End If
                Case "E_TM0360P8"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        'GetPage_Notice 9, 15 'REM RM12831
                        GetPage_Notice 9, 14 'RM12831
                        If Len(m_sXML_EddJobKey) > 0 Then
                               BuildSNBEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        End If
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        Set m_oPieceBacker = m_oSNBBackers(1)
                        m_bDuplex = True
                    End If
                    
                ' Approved
                Case "TA0420P1"
                    'CPTPK-2110
                     GetPage_Notice 9, 13
                    If m_bXML_Archive = True Then
                        BuildSNBPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx

                        End If
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = "BND20"
                    'End CPTPK-2110
'                    'REM CPTPK-2110
'                    ' If we are running the archive job - suppress these from being archived
'                    If m_bXML_Archive = True Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_Notice 9, 13    ' REM prin-275
'                        'GetPage_TA0420P True    ' prin-275
'                        m_lStackItems = StackItemsPressureSeal
'                    End If
'                     'End REM CPTPK-2110
                Case "E_TA0420P1"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TA0420P '   REM prin-275
                        'GetPage_TA0420P True    ' prin-275
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType10and20
                    ElseIf m_bXML_Archive = True Then
                        ' Only archive CH0202P1 for this client
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' E Notice only - do not print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved
                Case "TA0420P2"
                    ' If we are running the archive job - suppress these from print
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 9, 13    '   REM prin-275
                        'GetPage_TA0420P True    ' prin-275
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
                Case "E_TA0420P2"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_TA0420P '   REM prin-275
                        'GetPage_TA0420P True    ' prin-275
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsPressureSeal
                        m_lEddEnoticeType = EnoticeType10and20
                    ElseIf m_bXML_Archive = True Then
                        ' Only archive CH0202P1 for this client
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        ' E Notice only - do not print
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                
                ' Approved
                Case "TM0847P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TM0847P1
                        ' Trim the leading blank elements from the page array
                        m_lCurrentPageIdx = TrimLeadingBlanks(m_saCurrentPageLine)
                        m_lStackItems = StackItemsPressureSeal
                    End If
    
                ' Approved
                Case "CU1620P1"
                    ' If we are not running the archive job - suppress print for this notice
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_Notice 1, 12
                        m_sStockTray = m_scPaperTypePlain
                        m_lStackItems = StackItemsFlatStock42
                    End If
                      
                ' Approved
                Case "CH0202P1"
                    ' Only process this piece if we are archiving (CH0202P1 job)
                    If m_bXML_Archive = True Then
                        GetPage_FlatStock 5, 17
                        BuildSNBPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                    Else
                        ' Suppress print for this notice if running print job
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                ' Approved 07-19-2016
                Case "TA95602P1"
                    ' If we are running the archive job - suppress these from being archived
                    If m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA95602 True
                        m_sStockTray = m_scPaperTypePlain
                        m_lStackItems = StackItemsFlatStock42
                    End If
                    
                Case "TA0050P1"
                        ' Bypass these without warning.
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                        
                  

                    
'Begin RM 5213
'                Case "TA956024P"
'                    ' If we are running the archive job - suppress these from being archived
'                    If m_bXML_Archive = True Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_TA956024P
'                        m_lStackItems = StackItemsPressureSeal
'                    End If
'End RM 5213

                Case "E_LN0460P1"
                    If Len(m_sXML_EddEnoticeClientId) > 0 Then
                        GetPage_LN0460P 10, 13, True, False, False, False
                        BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_lEddEnoticeType = EnoticeType50
                        Set m_oReturnAddress = Nothing
                    ElseIf m_bXML_Archive = True Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    End If
                    
                'RM9344
                Case "TM0330P1"
                    GetPage_TM0330P1 7, 9, True
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = m_scPaperTypePlain
                'END RM9344
                        
'                'RM14945
'                Case "TM0365P1"
'                    GetPage_TM0365P1
'                    Set m_oReturnAddress = Nothing
''                    BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
'                    m_lStackItems = StackItemsFlatStock42
'                    m_sStockTray = m_scPaperTypePlain
'                    m_sEnvelopeCatCode = m_scEnvelope42
'                'END RM14945
                    
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
            End Select
            
            ' Default logos
            Set m_oExteriorLogoFormInfo = m_oSNBExteriorLogo(0)
            Set m_oInteriorLogoFormInfo = m_oSNBInteriorLogo(0)
            Set m_oFlatStockLogoFormInfo = m_oSNBFlatStockLogo(0)
            
            'RM9344
            Select Case m_sCurrentFileName
                Case "TM0330P1"
                    Set m_oExteriorLogoFormInfo = m_oSNBExteriorLogo(1)
                    Set m_oInteriorLogoFormInfo = m_oSNBInteriorLogo(1)
                    Set m_oFlatStockLogoFormInfo = m_oSNBFlatStockLogo(1)
                Case "LN0470P1"
                    Set m_oExteriorLogoFormInfo = m_oSNBExteriorLogo(2)
                    Set m_oInteriorLogoFormInfo = m_oSNBInteriorLogo(2)
                    Set m_oFlatStockLogoFormInfo = m_oSNBFlatStockLogo(2)
            End Select
            'END RM9344

            
        Case "Nationwide"
            If m_sXml_FileGroup = "LOANS" Then
                ' This file group is for the LOANS group - initial implementation of Nationwide
                Select Case m_sCurrentFileName
                    ' Approved 05-06-2015
                    Case "LN395Q1"
                        ' Suppress form type HIST from being printed
                        If m_sCurrentFileFormType = "HIST" Then
                            'GetPage_NoticeScrape 11, 11
                            GetPage_LN395Q1 True
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_lEddEnoticeType = EnoticeType50
                        ElseIf m_sCurrentFileFormType = "CLET" Then
                            GetPage_LN395Q1CLET
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsPressureSeal
                        Else
                            ' No programming in place to handle this file type
                            If Not m_bCurrentFileIsUnrecognized Then
                                If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                                End If
                                m_bCurrentFileIsUnrecognized = True
                                If m_lCurrentNoPrintStack = 0 Then
                                    m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                    m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                    SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                                End If
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                    Case "E_LN395Q1"
                        If m_bXML_Archive Then
                            ' Suppress form type HIST from being printed
                            If m_sCurrentFileFormType = "HIST" Then
                                'GetPage_NoticeScrape 11, 11
                                GetPage_LN395Q1 True
                                BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsFlatStock42
                            ElseIf m_sCurrentFileFormType = "CLET" Then
                                GetPage_LN395Q1CLET
                                BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsPressureSeal
                            Else
                                ' No programming in place to handle this file type
                                If Not m_bCurrentFileIsUnrecognized Then
                                    If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                                    End If
                                    m_bCurrentFileIsUnrecognized = True
                                    If m_lCurrentNoPrintStack = 0 Then
                                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                                    End If
                                End If
                                GetPage_NoPrint
                                m_lStackItems = StackItemsBypass
                            End If
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            ' Suppress form type HIST from being printed
                            If m_sCurrentFileFormType = "HIST" Then
                                'GetPage_NoticeScrape 11, 11
                                GetPage_LN395Q1 True
                                BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsFlatStock42
                                m_lEddEnoticeType = EnoticeType50
                            ElseIf m_sCurrentFileFormType = "CLET" Then
                                GetPage_LN395Q1CLET
                                BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsPressureSeal
                                m_lEddEnoticeType = EnoticeType50
                            Else
                                ' No programming in place to handle this file type
                                If Not m_bCurrentFileIsUnrecognized Then
                                    If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                                    End If
                                    m_bCurrentFileIsUnrecognized = True
                                    If m_lCurrentNoPrintStack = 0 Then
                                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                                    End If
                                End If
                                GetPage_NoPrint
                                m_lStackItems = StackItemsBypass
                            End If
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
    
                    ' Approved 05-06-2015
                    Case "LN0480Q1"
                        GetPage_LN0480
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        NationwideReformatLN0480
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = "TopPerf"
                        m_sEnvelopeCatCode = "42"
                        m_sOCR = Nationwide_LN0480_OCRLine(m_sAccountNumber, m_cAmountDue)
                        m_sngOCR_X = 54
                        m_sngOCR_Y = 559
                        m_bDuplex = True
                        Set m_oPieceBacker = m_oXml_Backer
                        m_sEnclosureCatCode = m_scNationwideBRE
                    Case "E_QSYSPRT"
                        If m_bXML_Archive Then
                            GetPage_LN0480
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            NationwideReformatLN0480
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = "TopPerf"
                            m_sEnvelopeCatCode = "42"
                            m_sOCR = Nationwide_LN0480_OCRLine(m_sAccountNumber, m_cAmountDue)
                            m_sngOCR_X = 54
                            m_sngOCR_Y = 559
                            m_bDuplex = True
                            Set m_oPieceBacker = m_oXml_Backer
                            m_sEnclosureCatCode = m_scNationwideBRE
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_LN0480
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            NationwideReformatLN0480
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = "TopPerf"
                            m_sEnvelopeCatCode = "42"
                            m_sOCR = Nationwide_LN0480_OCRLine(m_sAccountNumber, m_cAmountDue)
                            m_sngOCR_X = 54
                            m_sngOCR_Y = 559
                            m_bDuplex = True
                            Set m_oPieceBacker = m_oXml_Backer
                            m_sEnclosureCatCode = m_scNationwideBRE
                            m_lEddEnoticeType = EnoticeType50
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    ' Approved 05-06-2015
                    Case "T152Q2"
                        GetPage_T152Q2
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock428LPI
                        m_sStockTray = "TopPerf"
                        m_sEnvelopeCatCode = "42"
                        m_sOCR = Nationwide_LN0480_OCRLine(m_sAccountNumber, m_cAmountDue)
                        m_sngOCR_X = 54
                        m_sngOCR_Y = 559
                        'Set m_oFlatStockLogoFormInfo = m_oNationwideFlatStockLogo(0)
                        m_bDuplex = True
                        Set m_oPieceBacker = m_oXml_Backer
                        m_sEnclosureCatCode = m_scNationwideBRE
                        
                    ' Approved 05-06-2015
                    ' Moving to deposits
'                    Case "CU1610P1"
'                        GetPage_CU1610P1
'                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
'                        m_lStackItems = StackItemsFlatStock42
'                        m_sStockTray = m_scPaperTypePlain
'                        m_sEnvelopeCatCode = m_scEnvelope42
    
                    ' Approved 05-06-2015
                    Case "LN0111Q1"
                        GetPage_TA0050P1
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        NationwideReformatLN0111Q1
                        m_lStackItems = StackItemsFlatStock428LPI
                        m_sStockTray = m_scStockTrayPlainWhite20
                    Case "E_LN0111Q1"
                        If m_bXML_Archive Then
                            GetPage_TA0050P1
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            NationwideReformatLN0111Q1
                            m_lStackItems = StackItemsFlatStock428LPI
                            m_sStockTray = m_scStockTrayPlainWhite20
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA0050P1
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            NationwideReformatLN0111Q1
                            m_lStackItems = StackItemsFlatStock428LPI
                            m_sStockTray = m_scStockTrayPlainWhite20
                            m_lEddEnoticeType = EnoticeType50
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                     
                    Case Else
                        'This error message is raised if any new files are detected.   The code will
                        'need to be changed to process or suppress the file.
                        If Not m_bCurrentFileIsUnrecognized Then
                            If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                            End If
                            m_bCurrentFileIsUnrecognized = True
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                            End If
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                End Select ' End Select Case m_sCurrentFileName

            ElseIf m_sXml_FileGroup = "DEPOSITS" Then
                Select Case m_sCurrentFileName
                    Case "CU1610P1"
                        GetPage_CU1610P1
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                
                    Case "TM0847P1"
                        GetPage_TM0847P1
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
    
                    Case "TA0180P1"
                        If m_lCurrentPageIdx = 0 And Len(Trim(m_saCurrentPageLine(0))) = 0 Then
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        Else
                            GetPage_TA0180P
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        End If
                    Case "E_TA0180P1"
                        If m_lCurrentPageIdx = 0 And Len(Trim(m_saCurrentPageLine(0))) = 0 Then
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                            If Len(m_sXML_EddEnoticeClientId) > 0 Then
                                m_lEddEnoticeType = EnoticeType10and20
                            End If
                        Else
                            If m_bXML_Archive Then
                                GetPage_TA0180P
                                BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsFlatStock42
                                m_sStockTray = m_scPaperTypePlain
                                m_sEnvelopeCatCode = m_scEnvelope42
                            ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                                GetPage_TA0180P
                                Set m_oReturnAddress = Nothing
                                BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                                m_lStackItems = StackItemsFlatStock42
                                m_sStockTray = m_scPaperTypePlain
                                m_sEnvelopeCatCode = m_scEnvelope42
                                m_lEddEnoticeType = EnoticeType10and20
                            Else
                                If m_lCurrentNoPrintStack = 0 Then
                                    m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                    m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                    SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                                End If
                                GetPage_NoPrint
                                m_lStackItems = StackItemsBypass
                            End If
                        End If

                        
                    Case "TA0183P1"
                        GetPage_TA0180P
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
    
                    Case "TA0420P1"
                        GetPage_TA0420P
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TA0420P1"
                        If m_bXML_Archive Then
                            GetPage_TA0420P
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA0420P
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lEddEnoticeType = EnoticeType10and20
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA0420P2"
                        GetPage_TA0420P
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TA0420P2"
                        If m_bXML_Archive Then
                            GetPage_TA0420P
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA0420P
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM0360P2"
                        GetPage_TM0360P
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TM0360P2"
                        If m_bXML_Archive Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM0360P7"
                        GetPage_TM0360P
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TM0360P7"
                        If m_bXML_Archive Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM0360P8"
                        GetPage_TM0360P
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TM0360P8"
                        If m_bXML_Archive Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TM0360P
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA1161P1"
                        GetPage_TA1161P1
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TA1161P1"
                        If m_bXML_Archive Then
                            GetPage_TA1161P1
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA1161P1
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM0330P1"
                        GetPage_TM0330P1 7, 9, True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "TM0365P1"
                        GetPage_TM0365P1
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "E_TM0365P1"
                        If m_bXML_Archive Then
                            GetPage_TM0365P1
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TM0365P1
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95603P1"
                        GetPage_TA956023 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA956031"
                        If m_bXML_Archive Then
                            GetPage_TA956023 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA956023 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95602P1"
                        GetPage_TA956022 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA956021"
                        If m_bXML_Archive Then
                            GetPage_TA956022 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA956022 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95602P3"
                        GetPage_TA956023 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA956023"
                        If m_bXML_Archive Then
                            GetPage_TA956023 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA956023 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95602P4"
                        GetPage_TA956024 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA956024"
                        If m_bXML_Archive Then
                            GetPage_TA956024 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA956024 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "CH931613P1"
                        GetPage_CH931613 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "TA0050P1"
                        GetPage_HC5185
                        Set m_oReturnAddress = Nothing
                        ReformatPage_TA0050P1
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA0050P1"
                        If m_bXML_Archive Then
                            GetPage_HC5185
                            Set m_oReturnAddress = Nothing
                            ReformatPage_TA0050P1
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_HC5185
                            Set m_oReturnAddress = Nothing
                            ReformatPage_TA0050P1
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P1"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955991"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P2"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955992"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P3"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955993"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P4"
                        GetPage_TA955992 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955994"
                        If m_bXML_Archive Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P5"
                        GetPage_TA955992 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955995"
                        If m_bXML_Archive Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TA95599P6"
                        GetPage_TA955992 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TA955996"
                        If m_bXML_Archive Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955992 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType10and20
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM95604P1"
                        GetPage_TM956042 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TM956041"
                        If m_bXML_Archive Then
                            GetPage_TM956042 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TM956042 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM95599P1"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TM955991"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM95599P2"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TM955992"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "TM95599P3"
                        GetPage_TA955991 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                    Case "E_TM955993"
                        If m_bXML_Archive Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                        ElseIf Len(m_sXML_EddEnoticeClientId) > 0 Then
                            GetPage_TA955991 True
                            Set m_oReturnAddress = Nothing
                            BuildEddPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                            m_lStackItems = StackItemsFlatStock42
                            m_sStockTray = m_scPaperTypePlain
                            m_sEnvelopeCatCode = m_scEnvelope42
                            m_lEddEnoticeType = EnoticeType30
                        Else
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                        
                    Case "FT0160P1"
                        GetPage_FT0161 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "FT0160P2"
                        GetPage_FT0161 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case "FT0160P3"
                        GetPage_FT0161 True
                        Set m_oReturnAddress = Nothing
                        BuildNationwidePieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                        m_lStackItems = StackItemsFlatStock42
                        m_sStockTray = m_scPaperTypePlain
                        m_sEnvelopeCatCode = m_scEnvelope42
                        
                    Case Else
                        'This error message is raised if any new files are detected.   The code will
                        'need to be changed to process or suppress the file.
                        If Not m_bCurrentFileIsUnrecognized Then
                            If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                            End If
                            m_bCurrentFileIsUnrecognized = True
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                            End If
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass

                End Select ' End Select Case m_sCurrentFileName
            Else
                ' TODO Write error condition - nationwide must have a file group declared in the settings XML
                MessageBox "Invalid Nationwide config - xml Settings 'FileGroup' is not declared.  A file group must be declared on the settings record to process this client's files.", vbCritical + vbOKOnly, App.Title
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            
            Set m_oExteriorLogoFormInfo = Nothing
            Set m_oInteriorLogoFormInfo = Nothing
            
            Set m_oExteriorLogoFormInfo = m_oXml_OuterLogo
            Set m_oInteriorLogoFormInfo = m_oXml_InnerLogo
            Set m_oFlatStockLogoFormInfo = m_oXml_FlatStockLogo
            
        Case "ElDorado"
            Select Case m_sCurrentFileName
                    
                ' Approved 12/16/14
                Case "TA0180P1"
                    GetPage_Notice 1, 13 + IIf(Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And InStr(1, m_saCurrentPageLine(1), "***NOTICE OF ") > 0, 1, 0)  'bug fix for Signature release 17.2
                    m_lStackItems = StackItemsPressureSeal
                    
                ' Approved 12/16/14
                Case "TA0190P1"
                    GetPage_Notice 6, 13
                    m_lStackItems = StackItemsPressureSeal
                    
                ' Approved 12/16/14
                Case "TA0441P1"
                    GetPage_Notice 9, 13, True, False, False
                    m_lStackItems = StackItemsPressureSeal
                    
                ' Approved 02/10/15
                Case "SD0095P2"
                    GetPage_Notice 16, 16
                    m_lStackItems = StackItemsPressureSeal
                    
                ' Approved 02/10/15
                Case "SD0095P3"
                    GetPage_Notice 16, 16
                    m_lStackItems = StackItemsPressureSeal
                
                ' Approved 02/10/15
                Case "SD0095P4"
                    GetPage_Notice 16, 17, False, False, True
                    m_lStackItems = StackItemsPressureSeal
                    
                ' Approved 12/17/14
                Case "TA1161P1"
                    GetPage_Notice 11, 20
                    If m_oDeliveryAddress.CondensedAddressLines = 0 Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        m_lStackItems = StackItemsBypass
                    Else
                        m_lStackItems = StackItemsPressureSeal
                    End If
                    
'                Case "TM0330P1"
'                    GetPage_TM0330P1 7, 8, True
'                    ReformatTM0330P1
'                    m_lStackItems = StackItemsFlatStock42
'
'                    ' #42 envelope setup
'                    m_sStockTray = m_scStockTrayPlainWhite20
'                    m_sEnvelopeCatCode = m_scEnvelope42
                
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
            End Select
            
            Set m_oExteriorLogoFormInfo = Nothing
            Set m_oInteriorLogoFormInfo = Nothing
            
            Set m_oExteriorLogoFormInfo = m_oElDoradoExteriorLogo
            Set m_oInteriorLogoFormInfo = m_oElDoradoInteriorLogo
            Set m_oFlatStockLogoFormInfo = m_oElDoradoFlatStockLogo
            
            
        Case "Investors"
            Select Case m_sCurrentFileName
                                        
                ' Approved 08/24/2015
                Case "LN0473P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_LN0473P1 True
                        ReformatPage_LN0473P1
                        m_lStackItems = StackItemsLN0473
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                        'm_bReturnToBank = True
                    End If

                ' Approved 08/24/2015
                Case "LN0860P1"
                    If m_bXML_EddEstatements Then
                        ' Check if we are the LN0860 EDD job
                        If m_oJob.id <> 3498 Then
                            ' We are not the LN0860 job - suppress this file from print
                            If m_lCurrentNoPrintStack = 0 Then
                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                            End If
                            GetPage_NoPrint
                            m_lStackItems = StackItemsBypass
                        End If
                    End If
                    GetPage_LN0860P1 True, False
                    ReformatPage_LN0860P1
                    m_lStackItems = StackItemsFlatStockLegal
                    m_sStockTray = m_scInvestorsLEGAL
                    m_sEnvelopeCatCode = m_scEnvelope42
                    
                    If m_oJob.id = 3678 Then ' set up job record to send jobs to print.
                        m_bReturnToBank = False
                    Else
                        m_bReturnToBank = True
                    End If
                    
                ' Approved 09/08/2015
                ' Moved to PrimeDoc: 44253500\44253500_MixedDocs 10/26/2015
'                Case "PRINT1"
'                    ' "T152Q2" file, not
'                    If m_bXML_EddEstatements Then
'                        ' Check if we are the T152Q2 EDD job
'                        If m_oJob.id <> 3563 Then
'                            ' We are not the T152Q2 job - suppress this file from print
'                            If m_lCurrentNoPrintStack = 0 Then
'                                m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                                m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                                SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                            End If
'                            GetPage_NoPrint
'                            m_lStackItems = StackItemsBypass
'                        End If
'                    End If
'                    GetPage_PRINT1 True, False
'                    ReformatPage_PRINT1
'                    m_bReturnToBank = True
'                    m_lStackItems = StackItemsFlatStock9LPI ' StackItemsFlatStockLegal
'                    m_sStockTray = m_scInvestorsINVSTM 'm_scInvestorsLEGAL
'                    m_sEnvelopeCatCode = m_scEnvelope42
                                                            
                ' Approved 08/24/2015
                Case "SD0095P2"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_SD0095P True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                    End If
                    
                ' Approved 08/24/2015
                Case "SD0095P3"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_SD0095P True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                    End If
                    
                ' Approved 08/24/2015
                Case "SD0095P4"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_SD0095P True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                    End If

                    
                ' Cancelled by client
'                Case "TA0050P1"
'                    If m_lCurrentNoPrintStack = 0 Then
'                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                    End If
'                    GetPage_NoPrint
'                    m_lStackItems = StackItemsBypass
'                    GetPage_TA0050P1 True
'                    ReformatPage_TA0050P1
'                    m_lStackItems = StackItemsFlatStock8LPI
'                    m_sStockTray = m_scInvestorsINVSTM
                    
                ' Approved 08/24/2015
                Case "TA0180P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA0180P True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                    End If
                    
                ' Approved 08/26/2015
                Case "TA0185P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA0185P1 True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                    End If
                    
                ' Cancelled by client
'                Case "TA0190P1"
'                    If m_lCurrentNoPrintStack = 0 Then
'                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                    End If
'                    GetPage_NoPrint
'                    m_lStackItems = StackItemsBypass
'                    GetPage_TA0190P True, True
'                    m_lStackItems = StackItemsPressureSeal
'                    Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
'                    Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo

                ' Cancelled by client
'                Case "TA914863P"
'                    If m_bXML_EddEstatements Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    End If
'                    GetPage_TA914863P
'                    m_lStackItems = StackItemsFlatStock
'                    m_sStockTray = m_scInvestorsINVSTM
'                    m_sEnvelopeCatCode = m_scEnvelope42
                                
                ' Approved 08/24/2015
                Case "TA1161P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA1161P1 True
                        Set m_oReturnAddress = Nothing
                        m_sStockTray = m_scInvestorsINVSTM
                        m_lStackItems = StackItemsFlatStock
'                        m_bReturnToBank = True  'REM RM13442
                    End If
                    
                ' Approved 08/24/2015
                Case "TA0434P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA0434P1 True
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                    End If

                ' Cancelled by client
'                Case "TM0365P1"
'                    If m_lCurrentNoPrintStack = 0 Then
'                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                    End If
'                    GetPage_NoPrint
'                    m_lStackItems = StackItemsBypass
'                    GetPage_TM0365P1 True
'                    m_lStackItems = StackItemsFlatStock
'                    m_sStockTray = m_scInvestorsINVSTM
'                    m_sEnvelopeCatCode = m_scEnvelope42
                    
'                ' Not approved yet
'                Case "CH0202P1"
'                    If m_bXML_EddEstatements Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_CH0202P
'                        m_lStackItems = StackItemsFlatStock
'                        m_sStockTray = m_scInvestorsINVSTM
'                        m_bReturnToBank = True
'                    End If
'
'                ' Not approved yet
'                Case "CH0202P4"
'                    If m_bXML_EddEstatements Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_CH0202P
'                        m_lStackItems = StackItemsFlatStock
'                        m_sStockTray = m_scInvestorsINVSTM
'                        m_bReturnToBank = True
'                    End If
                
                ' Approved 08/24/2015
                Case "TA0420P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA0420P True
                        Set m_oReturnAddress = Nothing
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
'                        m_bReturnToBank = True  'REM RM13442
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                    End If

                ' Approved 08/24/2015
                Case "TA0420P2"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA0420P True
                        Set m_oReturnAddress = Nothing
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
'                        m_bReturnToBank = True  'REM RM13442
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                    End If
                    
                ' Not approved yet
'                Case "FT0160P1"
'                    If m_bXML_EddEstatements Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_FT0160P True, True
'                        m_lStackItems = StackItemsFlatStock
'                        m_sStockTray = m_scInvestorsINVSTM
'                        m_bReturnToBank = True
'                    End If
                    
                ' Approved 08/24/2015
                Case "FT0160P2"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_FT0160P True, True
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
'                        m_bReturnToBank = True 'REM RM10701
                    End If
                                
                ' Approved 08/24/2015
                Case "TM0847P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TM0847P1_Investors
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        If Mid(m_saCurrentPageLine(13), 1, 24) = "According to our records" Then
                            m_sEnclosureCatCode = m_scInvestorsBREMORE
                        ElseIf Mid(m_saCurrentPageLine(14), 1, 33) = "Our records indicate that you are" Then
                            m_sEnclosureCatCode = m_scInvestorsBRE
                        End If
                    End If
                    
                ' Approved 08/24/2015
                Case "TM0851P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TM0851
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                        m_sEnclosureCatCode = m_scInvestorsBRE
                        Set m_oPieceBacker = m_oInvestorsBackers(1)
                        m_bDuplex = True
                    End If
                
                ' Approved 08/24/2015
                Case "CU1610P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_CU1610P1
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                        m_bReturnToBank = True
                    End If
                    
                ' Approved 08/24/2015
                Case "CU1620P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_CU1620P1
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                    End If
                    
                ' Not approved yet
'                Case "TA0096P1"
'                    If m_bXML_EddEstatements Then
'                        If m_lCurrentNoPrintStack = 0 Then
'                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
'                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
'                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
'                        End If
'                        GetPage_NoPrint
'                        m_lStackItems = StackItemsBypass
'                    Else
'                        GetPage_TA0096P True, True
'                        m_lStackItems = StackItemsPressureSeal
'                        m_sStockTray = m_scStockTrayPressureSeal2
'                    End If
                    
                ' Approved 08/24/2015
                Case "TM0096P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TM0096P True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                    End If
                    
                ' Approved 08/24/2015
                Case "LN0096P1"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_LN0096P True, True, True
                        m_lStackItems = StackItemsPressureSeal
                        m_sStockTray = m_scStockTrayPressureSeal2
                        Set m_oInteriorLogoFormInfo = m_oInvestorsInteriorLogo
                        Set m_oExteriorLogoFormInfo = m_oInvestorsExteriorLogo
                    End If
                    
                Case "TA955289P"
                    If m_bXML_EddEstatements Then
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName & "(" & m_sCurrentFileFormType & ")"
                        End If
                        GetPage_NoPrint
                        m_lStackItems = StackItemsBypass
                    Else
                        GetPage_TA955289
                        m_bReturnToBank = True
                        m_lStackItems = StackItemsFlatStock
                        m_sStockTray = m_scInvestorsINVSTM
                        m_sEnvelopeCatCode = m_scEnvelope42
                    End If
                
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
            End Select
            
            ' RM15930
            If Not m_bXML_Archive Then
                'create index.  Per 1DC JIRA PRIN-251, 3 fields are:
                '1) Check Number
                '2) Account Number
                '3) Document Type
            
                m_lSecureDestructionIndexIdx = 2
                ReDim m_saSecureDestructionIndexKeys(m_lSecureDestructionIndexIdx)
                ReDim m_saSecureDestructionIndexValues(m_lSecureDestructionIndexIdx)
                
                ' Check Number
                m_saSecureDestructionIndexKeys(0) = "UserDef1"
                m_saSecureDestructionIndexValues(0) = "" ' No check numbers on notices

                ' Account Number
                m_saSecureDestructionIndexKeys(1) = "UserDef2"
                m_saSecureDestructionIndexValues(1) = TrimLeadingZeros2(Trim(m_sAccountNumber)) 'PRIN-251
'                m_saSecureDestructionIndexValues(1) = "" 'REM PRIN-251

                ' Document Type
                m_saSecureDestructionIndexKeys(2) = "UserDef3"
                m_saSecureDestructionIndexValues(2) = m_sCurrentFileName
            End If
            'END RM15930
        

            
        Case "The Palmetto Bank"
            Select Case m_sCurrentFileName
            
                Case "T152Q2"
                    GetPage_T152Q2
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsFlatStock8LPI
                    m_sStockTray = "TopPerf"
                    m_sEnvelopeCatCode = "42"
                    Set m_oPieceBacker = m_oBackerInfo
                    Set m_oFlatStockLogoFormInfo = m_oBPBFlatStockLogo
                    m_bDuplex = True
                
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    
            End Select
            
        Case "Allstate"
        
            Stop 'Until this is tested.
            Select Case m_sCurrentFileName
                Case "TM0847P1"
                    GetPage_TM0847P1
                Case Else
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    If m_lCurrentNoPrintStack = 0 Then
                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                    End If
                    'Send this to unrecognized report buffer.
            End Select
            
        Case "RiverCity"
            Select Case m_sCurrentFileName
            
                Case "T152Q2"
                    GetPage_T152Q2
                    'm_bReturnToBank = True
                    m_lStackItems = StackItemsFlatStock8LPI
                    m_sStockTray = "RC-STM"
                    m_sEnvelopeCatCode = "42"
                    
                ' Approved 01/23/2015
                Case "LN0480Q1"
                    GetPage_FlatStock 5, 11
                    NationwideReformatLN0480
                    BuildRiverCityPieceIndex m_saCurrentPageLine, m_lCurrentPageIdx
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = "RCB-HELOCP"
                    m_sEnvelopeCatCode = "42"
                
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    
            End Select
            
        Case "WooriAmerica"
            Debug.Print m_sCurrentFileName
            Select Case m_sCurrentFileName
                Case "LN0458P1"
                    GetPage_NoticeScrapeWoori 10, 14, 2, 1, 3, 36, True
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                
                Case "LN0460P1"
                    GetPage_NoticeScrapeWoori 10, 13, 2, 1, 3, 18, True
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                
                Case "LN0470P1"
                    GetPage_NoticeScrapeWoori 9, 13, 2, 1, 3, 28, True, False, True
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
                    
                Case "TA0185P1"
                    GetPage_NoticeScrape 6, 14, True
                    m_lStackItems = StackItemsFlatStock
                    m_sStockTray = "BND20"
            
                Case "LN395P1"
                    ' This type was not found in any of the test files during setup but was included
                    ' on the setup programming request. - 08/09/2012 - DB
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Type LN395P1 not configured. Contact CSR"
                    ' GetPage_Notice 10, 13
                    ' m_lStackItems = StackItemsFlatStock
                    ' m_sStockTray = "Z990857-LBS"
                
                Case "T152Q2"
                    If m_sCurrentFileFormType = "TIMESORT" Then
                        GetPage_NoticeScrape 15, 23, True
                        m_sStockTray = "BND20"
                        m_sEnclosureCatCode = ""
                    Else
                        GetPage_NoticeScrape 14, 18, True
                        m_sStockTray = "Z990857-LBS"
                        m_sEnclosureCatCode = m_scWooriT152Q2BRE
                    End If
                    m_lStackItems = StackItemsFlatStock8LPI
                    
                Case "LN395Q1"
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    If m_lCurrentNoPrintStack = 0 Then
                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                    End If
                
                Case "LN0480Q1"
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    If m_lCurrentNoPrintStack = 0 Then
                        m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                        m_lCurrentNoPrintStack = m_lNextNoPrintStack
                        SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                    End If
                Case Else
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
            End Select
        
        Case "HigherOne"
            Select Case m_sCurrentFileName
                Case "TA0180P1"
                    GetPage_Notice 1, 13 + IIf(Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And InStr(1, m_saCurrentPageLine(1), "***NOTICE OF ") > 0, 1, 0)  'bug fix for Signature release 17.2
                    m_lStackItems = StackItemsPressureSeal
                Case Else
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
            End Select
            
            Set m_oExteriorLogoFormInfo = m_oHigherOneExteriorLogo(0)
            Set m_oInteriorLogoFormInfo = m_oHigherOneInteriorLogo(0)
        
        Case m_scInst_29175100_FinFactors
            Select Case m_sCurrentFileName
            
                Case "LN0480Q1"
                    GetPage_LN0480
                    NationwideReformatLN0480
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = "TopPerf"
                    m_sEnvelopeCatCode = "42"
                    
                    Set m_oFlatStockLogoFormInfo = m_oXml_FlatStockLogo
                    Set m_oPieceBacker = m_oXml_Backer
                    m_bDuplex = True
                    If m_sXML_EddEnoticeClientId <> vbNullString Then
                        m_lEddEnoticeType = EnoticeType50
'                        Dim i As Long
'                        For i = 0 To m_lBulkShipAccountsIDX
'                            If Trim(m_sAccountNumber) = Trim(m_saBulkShipAccounts(i)) Then
''                                m_lStackItems = StackItemsBypass
'                                m_lEddEnoticeType = EnoticeTypeUnknown
'                                Exit For
'                            End If
'                        Next
                    End If
                Case "TM0330P1"
                    GetPage_TM0330P1 7, 8, True
                    'NationwideReformatLN0480
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = "TopPerf"
                    m_sEnvelopeCatCode = "42"
                    
                    Set m_oFlatStockLogoFormInfo = m_oXml_FlatStockLogo
                    Set m_oPieceBacker = m_oXml_Backer
                    m_bDuplex = True
                    If m_sXML_EddEnoticeClientId <> vbNullString Then
                        m_lEddEnoticeType = EnoticeType30
'                        Dim i As Long
'                        For i = 0 To m_lBulkShipAccountsIDX
'                            If Trim(m_sAccountNumber) = Trim(m_saBulkShipAccounts(i)) Then
''                                m_lStackItems = StackItemsBypass
'                                m_lEddEnoticeType = EnoticeTypeUnknown
'                                Exit For
'                            End If
'                        Next
                    End If
                Case Else
                    'This error message is raised if any new files are detected.   The code will
                    'need to be changed to process or suppress the file.
                    If Not m_bCurrentFileIsUnrecognized Then
                        If UnrecognizedErrorReqd(m_sCurrentFileName, m_sCurrentFileFormType) Then
                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unrecognized file - '" & m_sCurrentFileName & "'.  Programming and the CSR need to be notified so that the program can be coded to properly handle this file."
                        End If
                        m_bCurrentFileIsUnrecognized = True
                        If m_lCurrentNoPrintStack = 0 Then
                            m_lNextNoPrintStack = m_lNextNoPrintStack + 1
                            m_lCurrentNoPrintStack = m_lNextNoPrintStack
                            SetupNoPrintStack m_lCurrentNoPrintStack, m_sCurrentFileName
                        End If
                    End If
                    GetPage_NoPrint
                    m_lStackItems = StackItemsBypass
                    
            End Select
            
        Case "55810000", "75535000"
            Select Case m_sCurrentFileName
            
                Case "TM0330P1"
                    GetPage_TM0330P1 7, 8, True, bKeepReturnAddress:=True
                    If m_bXML_Archive Then
                        BuildIRAPieceIndex "", "", "IRA", m_oDeliveryAddress.Name
                    End If
                    m_lStackItems = StackItemsFlatStock42
                    m_sStockTray = "BND20"
                    m_sEnvelopeCatCode = "42"
                    
            End Select
        
        Case Else
            Stop
            
    End Select
    
    GetPage_Indirect = True
    
    Exit Function
    
GetPage_Indirect_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_NoPrint"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_Indirect_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_Indirect_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

  
    
End Function

Private Function GetPage_NoPrint() As Boolean
    Const lcAddrX As Long = 1
    Const lcAddrLen As Long = 40
    
    Dim sAssertionString As String
    
    On Error GoTo GetPage_NoPrint_Error
    
    InitializePieceFields
    
    GetPage_NoPrint = True
    
    Exit Function
    
GetPage_NoPrint_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_NoPrint"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_NoPrint_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_NoPrint_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function GetPage_SD0095() As Boolean
'    Const lcAddrX As Long = 16
'    Const lcAddrLen As Long = 40
'
'    Const scAccountNbr As String = "BRANCH   BOX"
'
'    Dim sAssertionString As String
'
'    On Error GoTo GetPage_SD0095_Error
'
'    InitializePieceFields
'
'    If m_lCurrentPageIdx < 20 Then
'        m_lCurrentPageIdx = 20
'        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
'    End If
'
'    StringAssertion m_saCurrentPageLine(5), scAccountNbr, 59, 6
'
'    m_sAccountNumber = ReadString(m_saCurrentPageLine(6), 59, 0, PvTrimOutput)
'
'    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, 16, lcAddrX, lcAddrLen, False)
'
''    m_oDeliveryAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(15), lcAddrX, lcAddrLen, PvRTrimOutput)
''    m_oDeliveryAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(16), lcAddrX, lcAddrLen, PvRTrimOutput)
''    m_oDeliveryAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(17), lcAddrX, lcAddrLen, PvRTrimOutput)
''    m_oDeliveryAddress.AddressLine(4) = ScrapeString(m_saCurrentPageLine(18), lcAddrX, lcAddrLen, PvRTrimOutput)
''    m_oDeliveryAddress.AddressLine(5) = ScrapeString(m_saCurrentPageLine(19), lcAddrX, lcAddrLen, PvRTrimOutput)
''    m_oDeliveryAddress.AddressLine(6) = ScrapeString(m_saCurrentPageLine(20), lcAddrX, lcAddrLen, PvRTrimOutput)
'
'    Set m_oCassAddress = m_oDeliveryAddress.Clone
'
'    GetPage_SD0095 = True
'
'    Exit Function
'
'GetPage_SD0095_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_SD0095"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_SD0095_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_SD0095_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Sub InitializePieceFields()
    Set m_oReturnAddress = New CAddress
    Set m_oFlatStockLogoFormInfo = Nothing
    Set m_oExteriorLogoFormInfo = Nothing
    Set m_oInteriorLogoFormInfo = Nothing
    Set m_oPieceBacker = Nothing
    Set m_oDeliveryAddress = New CAddress
    m_oDeliveryAddress.Alignment = fsaAlignBottom
    Set m_oCassAddress = New CAddress
    m_sAccountNumber = vbNullString
    m_sDocumentDate = vbNullString
    m_sRegion = vbNullString
    m_bTrailerPage = False
    m_bNoPrint = False
    m_bDuplex = False
    m_bReturnToBank = False
    m_lStackItems = StackItemsUndefined
    m_sStockTray = m_scPaperTypePressureSeal
    m_sEnvelopeCatCode = vbNullString
    m_sEnclosureCatCode = vbNullString
    m_sOCR = vbNullString
    m_sMICR = vbNullString
    m_sngOCR_X = m_sngcOCR_X
    m_sngOCR_Y = m_sngcOCR_Y

        
    Erase m_saIndexValues
    Erase m_saIndexKeys
    m_lIndexIdx = -1

    ' RM15930
    Erase m_saSecureDestructionIndexValues
    Erase m_saSecureDestructionIndexKeys
    m_lSecureDestructionIndexIdx = -1
    ' END RM15930

End Sub

Private Function GetPage_FlatStock(lAddrX As Long, lAddrLine As Long, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean) As Boolean
    
    On Error GoTo GetPage_FlatStock_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, bNoRightAddressMargin, bNoBottomAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    GetPage_FlatStock = True
    
    Exit Function
    
GetPage_FlatStock_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_FlatStock"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_FlatStock_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_FlatStock_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0480() As Boolean
    
    On Error GoTo GetPage_LN0480_Error
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oDocDate As Date
    
    lAddrX = 5
    lAddrLine = 11
    
    ' Put client specific coordinates in here
    Select Case m_sXmlInst
        
    End Select
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Gets the account number - account number is the first number on line 8
    oRegex.Pattern = "^\s+(\d+)\s+\d{1,2}-\d{2}-\d{2}\s+PAYMENT DUE$"
    If oRegex.Test(m_saCurrentPageLine(8)) = True Then
        Set oMatches = oRegex.Execute(m_saCurrentPageLine(8))
        m_sAccountNumber = oMatches.Item(0).SubMatches(0)
    Else
        Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
    End If
    
    ' Gets the document date
    oDocDate = FileDateTime(m_sInputFile(m_lInputFilePtr))
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    If m_sXmlInst = "Nationwide" Then
        ' Reads the amount due - this is used for creating OCR line
        ' Checks to see if the payment due field is in the proper location
        If Mid(m_saCurrentPageLine(8), 68, 11) = "PAYMENT DUE" Then
            ' If it is then the value underneath is the current amount due
            m_cAmountDue = CCur(Trim(Mid(m_saCurrentPageLine(9), 68, 11)))
        Else
            Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find amount due field.  This is required for OCR."
        End If
    End If
    

'RM9155
If m_bXML_Archive Then
    Dim i As Long
    m_lIndexIdx = 3

    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)

    m_saIndexKeys(0) = m_scDocumentTypeKey
    m_saIndexValues(0) = "LNS Line of Credit Statement - Fiserv"

    m_saIndexKeys(1) = m_scPrimaryNameKey
    m_saIndexValues(1) = m_oDeliveryAddress.Name

    m_saIndexKeys(2) = m_scStatementDateKey
    m_saIndexValues(2) = m_sDocumentDate

    m_saIndexKeys(3) = m_scAccountNumberKey
    m_saIndexValues(3) = m_sAccountNumber

    If m_sXmlInst = m_scInst_29175100_FinFactors Then
        m_saIndexValues(2) = ReformatDate(m_saIndexValues(2), "MM/dd/yyyy")

        m_lIndexIdx = 5

        ReDim Preserve m_saIndexKeys(m_lIndexIdx)
        ReDim Preserve m_saIndexValues(m_lIndexIdx)

        m_saIndexKeys(4) = m_scFileTypeKey
        m_saIndexValues(4) = "16"   'Finance Factors

        m_saIndexKeys(5) = "File_Name"
'        m_saIndexValues(5) = Right("0000000000" & m_lCurrentFilePieces, 10) & ".pdf"
'        m_saIndexValues(5) = "<Sequence>.pdf"
        m_saIndexValues(5) = Right("0000000000" & m_lCurrentFilePieces + 1, 10) & ".pdf"
    End If
End If
'END RM9155
'
'RM9153
    If Len(m_sXML_EddJobKey) > 0 Then
        m_lIndexIdx = 4
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)

        m_saIndexKeys(0) = "AccountNumber"
        m_saIndexValues(0) = m_sAccountNumber

        m_saIndexKeys(1) = "CustomerName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name

        m_saIndexKeys(2) = "StatementDate"
        'm_saIndexValues(2) = m_sDocumentDate
        m_saIndexValues(2) = Format(oDocDate, "yyyyMMdd")

        m_saIndexKeys(3) = "AccountTypeNumber"
        m_saIndexValues(3) = m_sXML_EddProductId & "-" & m_sAccountNumber

        m_saIndexKeys(4) = "NoticeType"
        m_saIndexValues(4) = "LN0480"

    End If
'END RM9153
    
    GetPage_LN0480 = True
    
    Exit Function
    
GetPage_LN0480_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0480"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0480_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0480_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function ReformatDate(ByVal sDate As String, ByVal sFormat As String) As String
    If Len(sDate) = 0 Then
        ReformatDate = vbNullString
        Exit Function
    End If
    
    ReformatDate = Format(CDate(sDate), sFormat)
End Function

Private Function GetPage_TM0851() As Boolean
    
    On Error GoTo GetPage_TM0851_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 4
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Removes the return address from the page since the corner card resource should contain it
    'Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
    'Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
    'Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
'    Set oRegexTest = New RegExp
'    oRegexTest.IgnoreCase = True
'    oRegexTest.Pattern = "^.{40}\s+Date:\s+(\d{1,2}/\d{1,2}/\d{4})"
'    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TM0851 page encountered - document date value not found."
'        GetPage_TM0851 = False
'        Exit Function
'    End If
    
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    sRawDate = oRegexMatch.SubMatches.Item(0)
'    oDocDate = CDate(sRawDate)
'    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TM0851 = True
    
    Exit Function
    
GetPage_TM0851_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0851"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0851_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0851_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_CU1610P1() As Boolean
    
    On Error GoTo GetPage_CU1610P1_Error
    
    Dim lAddrLine As Long
    Dim lAddrX As Long
    Dim lRetAddrLine As Long
    Dim lRetAddrX As Long
    
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
    lAddrLine = 12
    lAddrX = 1
    lRetAddrLine = 5
    lRetAddrX = 1
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Set m_oReturnAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lRetAddrLine, lRetAddrX, 40, True, False, False)
    Set m_oReturnAddress = Nothing
    
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^.{35}\s+(\d{1,2}\-\d{1,2}\-\d{2})"
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - document date value not found."
        GetPage_CU1610P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_CU1610P1 = True
    
    Exit Function
    
GetPage_CU1610P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_CU1610P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_CU1610P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_CU1610P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA955289() As Boolean
    
    On Error GoTo GetPage_TA955289_Error
    
    Dim lAddrLine As Long
    Dim lAddrX As Long
    Dim lRetAddrLine As Long
    Dim lRetAddrX As Long
    
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
    lAddrLine = 2
    lAddrX = 2
'    lRetAddrLine = 0
'    lRetAddrX = 0
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
            
            
    Const scPage As String = "PAGE "
    Dim lPos As Long
    lPos = InStr(50, m_saCurrentPageLine(1), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 1 is missing."
    End If
    
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saCurrentPageLine(1), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 1 is missing."
    End If
    
    If sWorkPage <> "1" Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected multi-page document.  These have not been coded or tested."
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
'    Set m_oReturnAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lRetAddrLine, lRetAddrX, 40, True, False, False)
'    Set m_oReturnAddress = Nothing
    
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^.{35}\s*FROM\s+(\d{1,2}/\d{1,2}/\d{2})"
    If oRegexTest.Test(m_saCurrentPageLine(3)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - document date value not found."
        GetPage_TA955289 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(3))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA955289 = True
    
    Exit Function
    
GetPage_TA955289_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA955289"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA955289_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA955289_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_CU1620P1() As Boolean
    
    On Error GoTo GetPage_CU1620P1_Error
    
    Dim lAddrLine As Long
    Dim lAddrX As Long
    Dim lRetAddrLine As Long
    Dim lRetAddrX As Long
    
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
    lAddrLine = 12
    lAddrX = 1
    lRetAddrLine = 5
    lRetAddrX = 1
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Set m_oReturnAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lRetAddrLine, lRetAddrX, 40, True, False, False)
    Set m_oReturnAddress = Nothing
    
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s*(\d{1,2}\-\d{1,2}\-\d{2})"
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - document date value not found."
        GetPage_CU1620P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    m_saCurrentPageLine(1) = Right(Space(73) & sRawDate, 73)
    
    GetPage_CU1620P1 = True
    
    Exit Function
    
GetPage_CU1620P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_CU1620P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_CU1620P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_CU1620P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function NationwideReformatLN0480() As Boolean

    Dim tempStr As String
    Dim currLine As Long
    Dim moveToLine As Long
    Dim offset As Long
    Dim currLineLen As Long
    Dim currLineOffset As Long
    
    Const lcLineLen = 79
    
    On Error GoTo NationwideReformatLN0480_Error
    
    For currLine = 6 To 8
        currLineLen = 20
        currLineOffset = 1
        If Len(Trim(Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen))) > 0 Then
            'm_saCurrentPageLine(currLine - 6) = Space(lcLineLen - currLineLen) & Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen)
            m_saCurrentPageLine(currLine - 6) = Right(Space(78) & Trim(Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen)), 78)
            'm_saCurrentPageLine(currLine - 6) = Right(Space(78) & Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen), 78)
            Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen) = Space(currLineLen)
            

'            tempStr = m_saCurrentPageLineBold(currLine)
'            m_saCurrentPageLineBold(currLine - 6) = m_saCurrentPageLineBold(currLine)
'            m_saCurrentPageLineBold(currLine) = tempStr
        End If
        
        If Len(Trim(Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen))) > 0 Then
            m_saCurrentPageLineBold(currLine - 6) = Right(Space(78) & Trim(Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen)), 78)
            Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen) = Space(currLineLen)
            
        End If
        
        currLineLen = 30
        currLineOffset = 20
        If Len(Trim(Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen))) > 0 Then
            'm_saCurrentPageLine(currLine - 3) = Space(lcLineLen - currLineLen) & Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen)
            m_saCurrentPageLine(currLine - 3) = Right(Space(78) & Trim(Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen)), 78)
            'm_saCurrentPageLine(currLine - 3) = Right(Space(78) & Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen), 78)
            Mid(m_saCurrentPageLine(currLine), currLineOffset, currLineLen) = Space(currLineLen)
'            tempStr = m_saCurrentPageLineBold(currLine)
'            m_saCurrentPageLineBold(currLine - 3) = m_saCurrentPageLineBold(currLine)
'            m_saCurrentPageLineBold(currLine) = tempStr
        End If
        
        If Len(Trim(Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen))) > 0 Then
            m_saCurrentPageLineBold(currLine - 3) = Right(Space(78) & Trim(Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen)), 78)
            Mid(m_saCurrentPageLineBold(currLine), currLineOffset, currLineLen) = Space(currLineLen)
        End If
        
    Next
    
    NationwideReformatLN0480 = True
    
    Exit Function
    
NationwideReformatLN0480_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.NationwideReformatLN0480"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo NationwideReformatLN0480_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo NationwideReformatLN0480_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function NationwideReformatT152Q2() As Boolean
  
    Const lcPageLen = 70
    Const lcNumInsert = 5
    
    Dim lPageLen As Long
    Dim lNumInsert As Long
    
    lPageLen = lcPageLen
    lNumInsert = lcNumInsert
    
    On Error GoTo NationwideReformatT152Q2_Error
    
    If m_lCurrentPageIdx + lNumInsert > lPageLen Then
        Err.Raise g_lcErrInvalidPage, App.Title, "There is not enough space on the page to move the contents of the page down. Account number: " & m_sAccountNumber
        NationwideReformatT152Q2 = False
        Exit Function
    End If
    
    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 14, lNumInsert
    
    NationwideReformatT152Q2 = True
    
    Exit Function
    
NationwideReformatT152Q2_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.NationwideReformatT152Q2"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo NationwideReformatT152Q2_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo NationwideReformatT152Q2_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function NationwideReformatLN0111Q1() As Boolean
    
    Dim lPageDeleteStart As Long
    Dim lNumLinesToDelete As Long
    
    On Error GoTo NationwideReformatLN0111Q1_Error
    
    lPageDeleteStart = 9
    lNumLinesToDelete = 7
    
    ' Removes one of the two division lines
    If Trim(m_saCurrentPageLine(8)) = "_______________________________________________________________________" Then
        m_saCurrentPageLine(8) = vbNullString
        lPageDeleteStart = 8
        lNumLinesToDelete = 8
    End If
    
    ' Deletes lines 8 through 16 - moving the contents up the page so it will be less likely to run off the bottom
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, lPageDeleteStart, lNumLinesToDelete
    
    NationwideReformatLN0111Q1 = True
    
    Exit Function
    
NationwideReformatLN0111Q1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.NationwideReformatLN0111Q1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo NationwideReformatLN0111Q1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo NationwideReformatLN0111Q1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function ReformatPage_TA0050P1() As Boolean
    
    Dim lPageDeleteStart As Long
    Dim lPageDeleteEnd As Long
    Dim lCurLine As Long
    
    On Error GoTo ReformatPage_TA0050P1_Error
    
    ' Loops through the page array until it hits the first line
    lPageDeleteStart = -1
    lPageDeleteEnd = -1
    For lCurLine = 0 To m_lCurrentPageIdx
        ' Checks if we have hit one of the dividers for the delivery address
        If Trim(m_saCurrentPageLine(lCurLine)) = "_______________________________________________________________________" Then
            If lPageDeleteStart = -1 Then
                lPageDeleteStart = lCurLine
            ElseIf lPageDeleteEnd = -1 And lPageDeleteStart > -1 Then
                lPageDeleteEnd = lCurLine
            End If
        End If
        If lPageDeleteStart > -1 Then
            m_saCurrentPageLine(lCurLine) = vbNullString
            m_saCurrentPageLineBold(lCurLine) = vbNullString
        End If
        If lPageDeleteEnd > -1 Then
            Exit For
        End If
    Next
    
    'lPageDeleteStart = 9
    'lNumLinesToDelete = 7
    
    ' Removes one of the two division lines
    'If Trim(m_saCurrentPageLine(8)) = "_______________________________________________________________________" Then
    '    m_saCurrentPageLine(8) = vbNullString
    '    lPageDeleteStart = 8
    '    lNumLinesToDelete = 8
    'End If
    
    ' Deletes lines 8 through 16 - moving the contents up the page so it will be less likely to run off the bottom
    If lPageDeleteStart > -1 And lPageDeleteEnd > -1 Then
        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, lPageDeleteStart, (lPageDeleteEnd - lPageDeleteStart + 1)
    ElseIf lPageDeleteStart > -1 And lPageDeleteEnd = -1 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - No closing bar found aroun delivery address for TA0050P1 page."
        ReformatPage_TA0050P1 = False
        Exit Function
    End If
        
    
    ReformatPage_TA0050P1 = True
    
    Exit Function
    
ReformatPage_TA0050P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatPage_TA0050P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatPage_TA0050P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatPage_TA0050P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function ReformatPage_LN0860P1() As Boolean
    
    Dim lPageDeleteStart As Long
    Dim lPageDeleteEnd As Long
    Dim lCurLine As Long
    Dim sLoanNumber As String
    Dim sDocTitle As String
    Dim sDocDate As String
    
    On Error GoTo ReformatPage_LN0860P1_Error
    
    sLoanNumber = Trim(m_saCurrentPageLine(15))
    m_saCurrentPageLine(15) = vbNullString
    sDocTitle = Trim(m_saCurrentPageLineBold(7))
    m_saCurrentPageLineBold(7) = vbNullString
    sDocDate = Trim(m_saCurrentPageLine(8))
    m_saCurrentPageLine(8) = vbNullString
    
    ' Re-locates the document title and loan number to the same line
    If Len(Trim(m_saCurrentPageLine(19))) = 0 And Len(Trim(m_saCurrentPageLineBold(19))) = 0 Then
        m_saCurrentPageLine(19) = Right(Space(83) & sLoanNumber, 83)
        m_saCurrentPageLineBold(19) = Space(1) & sDocTitle
    Else
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to construct document title page.  Content found on line 19 of document"
        ReformatPage_LN0860P1 = False
        Exit Function
    End If
    
    If Len(Trim(m_saCurrentPageLine(20))) = 0 And Len(Trim(m_saCurrentPageLineBold(20))) = 0 Then
        m_saCurrentPageLine(20) = Space(1) & sDocDate
    Else
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to construct document title page line 2.  Content found on line 20 of document"
        ReformatPage_LN0860P1 = False
        Exit Function
    End If

    ReformatPage_LN0860P1 = True
    
    Exit Function
    
ReformatPage_LN0860P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatPage_LN0860P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatPage_LN0860P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatPage_LN0860P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function ReformatPage_PRINT1() As Boolean
'
'    Dim lPageDeleteStart As Long
'    Dim lPageDeleteEnd As Long
'    Dim lCurLine As Long
'    Dim sLoanNumber As String
'    Dim sDocTitle As String
'    Dim sDocDate As String
'
'    On Error GoTo ReformatPage_PRINT1_Error
'
'    ' Find which line 'Balance Information' starts on - all lines above this will be aligned to the right
'    Dim lBalanceLine As Long
'    lBalanceLine = -1
'    For lCurLine = 0 To m_lCurrentPageIdx
'        If InStr(1, m_saCurrentPageLine(lCurLine), " Account:") > 0 Then
'            lBalanceLine = lCurLine
'            Exit For
'        End If
'    Next
'
'    If lBalanceLine < 0 Then
'        Err.Raise g_lcErrBase + g_lcErrGenericAssertion, App.Title, "Invalid PRINT1 page encountered."
'    End If
'
'    ' We want to right align all data that is above the balance line
'    For lCurLine = 0 To lBalanceLine
'        m_saCurrentPageLine(lCurLine) = Right(Space(82) & Trim(m_saCurrentPageLine(lCurLine)), 82)
'    Next
'
'
'
'    ReformatPage_PRINT1 = True
'
'    Exit Function
'
'ReformatPage_PRINT1_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.ReformatPage_PRINT1"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo ReformatPage_PRINT1_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo ReformatPage_PRINT1_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function


Private Function ReformatPage_LN0473P1() As Boolean
    
    On Error GoTo ReformatPage_LN0473P1_Error
    
    If Len(Trim(m_saCurrentPageLine(16))) = 0 And Len(Trim(m_saCurrentPageLineBold(16))) = 0 Then
        m_saCurrentPageLine(16) = Space(43) & ScrapeString(m_saCurrentPageLine(4), 43)
        m_saCurrentPageLineBold(16) = Space(43) & ScrapeString(m_saCurrentPageLineBold(4), 43)
    Else
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - data in area for relocating property address LN0473P1."
        ReformatPage_LN0473P1 = False
        Exit Function
    End If

    If Len(Trim(m_saCurrentPageLine(17))) = 0 And Len(Trim(m_saCurrentPageLineBold(17))) = 0 Then
        m_saCurrentPageLine(17) = Space(43) & ScrapeString(m_saCurrentPageLine(5), 43)
        m_saCurrentPageLineBold(17) = Space(43) & ScrapeString(m_saCurrentPageLineBold(5), 43)
    Else
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - data in area for relocating property address LN0473P1."
        ReformatPage_LN0473P1 = False
        Exit Function
    End If
    
    If Len(Trim(m_saCurrentPageLine(18))) = 0 And Len(Trim(m_saCurrentPageLineBold(18))) = 0 Then
        m_saCurrentPageLine(18) = Space(43) & ScrapeString(m_saCurrentPageLine(6), 43)
        m_saCurrentPageLineBold(18) = Space(43) & ScrapeString(m_saCurrentPageLineBold(6), 43)
    Else
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - data in area for relocating property address LN0473P1."
        ReformatPage_LN0473P1 = False
        Exit Function
    End If

    ReformatPage_LN0473P1 = True
    
    Exit Function
    
ReformatPage_LN0473P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatPage_LN0473P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatPage_LN0473P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatPage_LN0473P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


Private Function ReformatPage_TA990771() As Boolean
    
    On Error GoTo ReformatPage_TA990771_Error
    
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 8, 5
    
    Dim i As Long
    For i = 0 To 6
        m_saCurrentPageLine(i) = Right(Space(87) & Trim(m_saCurrentPageLine(i)), 87)
    Next
    

    ReformatPage_TA990771 = True
    
    Exit Function
    
ReformatPage_TA990771_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatPage_TA990771"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatPage_TA990771_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatPage_TA990771_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function ReformatPage_TM0847P1_Investors() As Boolean
    
    Dim lPageDeleteStart As Long
    Dim lPageDeleteEnd As Long
    Dim lCurLine As Long
    
    On Error GoTo ReformatPage_TM0847P1_Investors_Error
    
    ' Loops through the page array until it hits the first line
    lPageDeleteStart = -1
    lPageDeleteEnd = -1
    For lCurLine = 0 To m_lCurrentPageIdx
        If lPageDeleteStart = -1 Then
            lPageDeleteStart = lCurLine
        End If
        If Len(Trim(m_saCurrentPageLine(lCurLine))) > 0 Then
            If m_saCurrentPageLine(lCurLine) = "________________________________________________________________________________" Then
                lPageDeleteEnd = lCurLine - 1
                Exit For
            Else
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - non-blank data found in deletion area for TM0874P1."
                ReformatPage_TM0847P1_Investors = False
                Exit Function
            End If
        End If
    Next
    
    'lPageDeleteStart = 9
    'lNumLinesToDelete = 7
    
    ' Removes one of the two division lines
    'If Trim(m_saCurrentPageLine(8)) = "_______________________________________________________________________" Then
    '    m_saCurrentPageLine(8) = vbNullString
    '    lPageDeleteStart = 8
    '    lNumLinesToDelete = 8
    'End If
    
    ' Deletes lines 8 through 16 - moving the contents up the page so it will be less likely to run off the bottom
    If lPageDeleteStart > -1 And lPageDeleteEnd > -1 Then
        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, lPageDeleteStart, (lPageDeleteEnd - lPageDeleteStart + 1)
    ElseIf lPageDeleteStart > -1 And lPageDeleteEnd = -1 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - No closing bar found aroun delivery address for TM0847P1_Investors page."
        ReformatPage_TM0847P1_Investors = False
        Exit Function
    End If
        
    
    ReformatPage_TM0847P1_Investors = True
    
    Exit Function
    
ReformatPage_TM0847P1_Investors_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatPage_TM0847P1_Investors"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatPage_TM0847P1_Investors_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatPage_TM0847P1_Investors_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


'Private Function GetPage_T152Q2() As Boolean
'
'    Const lAddrX As Long = 10
'    Const lAddrLine As Long = 20
'    On Error GoTo GetPage_T152Q2_Error
'
'    InitializePieceFields
'
'    If m_lCurrentPageIdx < lAddrLine + 4 Then
'        m_lCurrentPageIdx = lAddrLine + 4
'        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
'        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
'    End If
'
'    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
'
'    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, True)
'    Set m_oCassAddress = m_oDeliveryAddress.Clone
'
'    Dim i As Long
''    For i = 0 To 8
''        If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
''            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Line 1 through 9 must be blank on a T152Q2.  The following data will be dropped:" & vbCrLf & vbCrLf & "'" & Trim(m_saCurrentPageLine(i)) & "'."
''        End If
''    Next
''    If m_lCurrentPageIdx > 76 Then
''        Stop
''    End If
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 19, 3
'
''
''    '--------------------------------
''    'Move letter text down to get it out of the envelope window.
''    If Len(Trim(m_saCurrentPageLine(17))) > 0 Then
''        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data found on line 18 on an LN395Q1."
''    End If
''
'''    If Left(Trim(m_saCurrentPageLine(18)), 4) <> "RE: " Then
'''        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data not found on line 19 on an LN395Q1."
'''    End If
''
''    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 18, 2
''    '--------------------------------
'
'    GetPage_T152Q2 = True
'
'    Exit Function
'
'GetPage_T152Q2_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_T152Q2"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_T152Q2_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_T152Q2_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Function GetPage_T152Q2() As Boolean
    On Error GoTo GetPage_T152Q2_Error
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim oRegexFinder As RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Select Case m_sXmlInst
        Case "The Palmetto Bank"
            lAddrX = 22
            lAddrLine = 21
        Case "SNB"
            lAddrLine = 21
            lAddrX = 20
        Case "Nationwide"
            lAddrX = 7
            lAddrLine = 9
        Case Else
            lAddrX = 20
            lAddrLine = 12
    End Select
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Dim i As Long
    For i = 12 To 17
        'Remove OMR marks.
        If i > m_lCurrentPageIdx Then
            Exit For
        End If
        If Mid(m_saCurrentPageLine(i), 6, 4) = "----" Then
            Mid(m_saCurrentPageLine(i), 6) = "    "
        End If
    Next
    
    Select Case m_sXmlInst
        Case "The Palmetto Bank"
            For i = 5 To 12
                ScrapeString m_saCurrentPageLine(i), 7, 37
            Next
    End Select
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    If m_sCurrentFileFormType = "HIST" Then
        If Trim(Mid(m_saCurrentPageLine(8), 55, 15)) = "Statement Date" Then
            m_sDocumentDate = Trim(Mid(m_saCurrentPageLine(8), 70, 8))
        End If
        
        If Not IsDate(m_sDocumentDate) Then
            m_sDocumentDate = vbNullString
        Else
            m_sDocumentDate = Format(CDate(m_sDocumentDate), "yyyy-MM-dd")
            m_sEddCycleDate = CDate(m_sDocumentDate)
        End If
        
        ' Inits the regular expression
        Set oRegexFinder = New RegExp
        oRegexFinder.IgnoreCase = True
        oRegexFinder.Pattern = " Account Number\s+(\d+)"
        If oRegexFinder.Test(m_saCurrentPageLine(11)) = True Then
            Set oMatches = oRegexFinder.Execute(m_saCurrentPageLine(11))
            Set oMatch = oMatches.Item(0)
            m_sAccountNumber = oMatch.SubMatches(0)
        End If
        
'        If Mid(m_saCurrentPageLine(11), 49, 14) = "Account Number" Then
'            m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(11), 63, 14))
'        End If
        
        If Len(m_sDocumentDate) = 0 Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Statement date not found."
        End If
        
        If Len(m_sAccountNumber) = 0 Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Account number not found."
        End If
        
        
    ElseIf m_sCurrentFileFormType = "ANALFRM" Then
        If Mid(m_saCurrentPageLine(11), 63, 14) = "Account Number" Then
            m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(12), 63, 14))
        End If
        If Trim(Mid(m_saCurrentPageLine(8), 63, 15)) = "Statement Date" Then
            m_sDocumentDate = Trim(Mid(m_saCurrentPageLine(9), 63, 15))
        End If
        
        If Not IsDate(m_sDocumentDate) Then
            m_sDocumentDate = vbNullString
        Else
            m_sDocumentDate = Format(CDate(m_sDocumentDate), "yyyy-MM-dd")
            m_sEddCycleDate = CDate(m_sDocumentDate)
        End If
    End If
        
    If Len(m_sXML_EddJobKey) > 0 Or m_bXML_EddEstatements Then
        Dim sAccountNumber As String
        If Len(m_sAccountNumber) < 10 Then
            sAccountNumber = Right(String(10, "0") & m_sAccountNumber, 10)
        Else
            sAccountNumber = m_sAccountNumber
        End If
        
        m_lIndexIdx = 2
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
        
        m_saIndexKeys(0) = "Account_Number"
        m_saIndexValues(0) = sAccountNumber
        
        m_saIndexKeys(1) = "PrimaryName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name
        
        m_saIndexKeys(2) = "Statement_Date"
        m_saIndexValues(2) = m_sDocumentDate
        
        If Len(m_sXML_AccountType) > 0 Then
            m_lIndexIdx = m_lIndexIdx + 2
            ReDim Preserve m_saIndexKeys(m_lIndexIdx)
            ReDim Preserve m_saIndexValues(m_lIndexIdx)
            m_saIndexKeys(m_lIndexIdx - 1) = "Account_Type"
            m_saIndexValues(m_lIndexIdx - 1) = m_sXML_AccountType
            m_saIndexKeys(m_lIndexIdx) = "EDDDocumentType"
            m_saIndexValues(m_lIndexIdx) = m_sXML_AccountType & "-" & sAccountNumber
        End If
    Else
        m_lIndexIdx = 3
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
        
        m_saIndexKeys(0) = "File_Name"
        m_saIndexValues(0) = m_sCurrentFileName
        
        m_saIndexKeys(1) = "Account_Number"
        m_saIndexValues(1) = m_sAccountNumber
        
        m_saIndexKeys(2) = "Document_Date"
        m_saIndexValues(2) = m_sDocumentDate
        
        m_saIndexKeys(3) = "Name"
        m_saIndexValues(3) = m_oDeliveryAddress.Name
    End If
    
    ' Trims extra blank lines from SNB T152Q2
    Select Case m_sXmlInst
        Case "SNB"
            DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 5
            'Checks if there is a remit portion on this piece
            If InStr(1, m_saCurrentPageLine(67), "To ensure proper credit, please return this portion with payment.", vbTextCompare) > 0 Then
                m_sMICR = BuildMICRLinePs(vbNullString, "5999 0025", vbNullString, vbNullString, 0)
            End If
        Case "The Palmetto Bank"
            ' Check if this is the "print" job for The Palmetto Bank - 2947
            ' Tim Feierabend 03/17/2015
            If m_oJob.id = 2947 Then
                ' loops through the delivery address looking for "NO MAIL" condition - if found then suppress from print
                ' Tim Feierabend 03/19/2015
                For i = 1 To m_oDeliveryAddress.CondensedAddressLines
                    If InStr(1, m_oDeliveryAddress.CondensedAddress(i), "NO MAIL", vbTextCompare) > 0 Then
                        m_bNoPrint = True
                    End If
                Next
            End If
    End Select
    
    GetPage_T152Q2 = True
    
    Exit Function
    
GetPage_T152Q2_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_T152Q2"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_T152Q2_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_T152Q2_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_T152Q2Trailer() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_T152Q2Trailer = False
    
    On Error GoTo GetPage_T152Q2Trailer_Error
    
    Dim i As Long
    Dim bContinued As Boolean
''    Can't just look for END OF STATEMENT or CONTINUED ON NEXT PAGE, that would be too easy.  Signature has a bug
''    where they put END OF STATEMENT on every page, even if it is not the end of the statement.
''
''    For i = m_lCurrentPageIdx To 0 Step -1
''        If InStr(1, m_saCurrentPageLine(i), "END OF STATEMENT", vbTextCompare) > 0 Then
''            Exit Function
''        End If
''
''        If InStr(1, m_saCurrentPageLine(i), "CONTINUED ON NEXT PAGE", vbTextCompare) > 0 Then
''            bContinued = True
''            Exit For
''        End If
''    Next
''
''    If Not bContinued Then
''        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "'END OF STATEMENT' or 'CONTINUED ON NEXT PAGE' not found."
''    End If

    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    Select Case m_sXmlInst
        Case "The Palmetto Bank"
            Const scPageNum As String = " Page "
            If Mid(m_saReadAheadPageLine(5), 70, Len(scPageNum)) = scPageNum Then
                ScrapeString m_saReadAheadPageLine(5), 7, 37
            End If
    End Select
    
    Dim lPageLineNum As Long

    Select Case m_sXmlInst
        Case "RiverCity"
            Const scPage As String = "Page "
            Dim lPos As Long
            lPos = InStr(50, m_saReadAheadPageLine(6), scPage, vbTextCompare)
            If lPos = 0 Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 7 is missing."
            End If
            
            Dim sWorkPage As String
            sWorkPage = Trim(Mid(m_saReadAheadPageLine(6), lPos + Len(scPage)))
            
            If Not IsNumeric(sWorkPage) Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 7 is missing."
            End If
            
            If Val(sWorkPage) <= 1 Then
                Exit Function
            End If
        Case "SNB"
            lPageLineNum = 5
            
            lPos = InStr(50, m_saReadAheadPageLine(lPageLineNum), scPage, vbTextCompare)
            If lPos = 0 Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
            End If
            
            sWorkPage = Trim(Mid(m_saReadAheadPageLine(lPageLineNum), lPos + Len(scPage)))
            
            If Not IsNumeric(sWorkPage) Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
            End If
            
            If Val(sWorkPage) <= 1 Then
                Exit Function
            Else
                DeleteLinesFromArray m_saReadAheadPageLine, m_saReadAheadPageLineBold, m_lReadAheadPageIdx, 0, 5
                Set m_oFlatStockLogoFormInfo = Nothing
                Set m_oPieceBacker = Nothing
                m_bDuplex = False
                'Checks if there is a remit portion on this piece
                If InStr(1, m_saReadAheadPageLine(67), "To ensure proper credit, please return this portion with payment.", vbTextCompare) > 0 Then
                    m_sMICR = BuildMICRLinePs(vbNullString, "5999 0025", vbNullString, vbNullString, 0)
                End If
            End If
            
        Case "Nationwide"
            lPageLineNum = 6
            
            lPos = InStr(1, m_saReadAheadPageLine(lPageLineNum), scPage, vbTextCompare)
            If lPos = 0 Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
            End If
            
            sWorkPage = Trim(Mid(m_saReadAheadPageLine(lPageLineNum), lPos + Len(scPage)))
            
            If Not IsNumeric(sWorkPage) Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
            End If
            
            If Val(sWorkPage) <= 1 Then
                Exit Function
            Else
                Set m_oPieceBacker = Nothing
                m_bDuplex = False
            End If
            
        Case Else
            For i = m_lCurrentPageIdx To 0 Step -1
                If InStr(1, m_saCurrentPageLine(i), "END OF STATEMENT", vbTextCompare) > 0 Then
                    Exit Function
                End If
        
                If InStr(1, m_saCurrentPageLine(i), "CONTINUED ON NEXT PAGE", vbTextCompare) > 0 Then
                    bContinued = True
                    Exit For
                End If
            Next
        
            If Not bContinued Then
                'If neither END OF STATEMENT or CONTINUED ON NEXT PAGE found, assume this is a trailer page, but
                'assert that page number > 1.
                lPageLineNum = 5
                lPos = InStr(1, m_saReadAheadPageLine(lPageLineNum), scPage, vbTextCompare)
                If lPos = 0 Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
                End If
                
                sWorkPage = Trim(Mid(m_saReadAheadPageLine(lPageLineNum), lPos + Len(scPage)))
                
                If Not IsNumeric(sWorkPage) Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line " & CStr(lPageLineNum + 1) & " is missing."
                End If
                
                If Val(sWorkPage) <= 1 Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "'END OF STATEMENT' or 'CONTINUED ON NEXT PAGE' not found."
                End If
                
            End If
        
    End Select
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_T152Q2Trailer = True
    
    Exit Function
    
GetPage_T152Q2Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_T152Q2Trailer"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_T152Q2Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_T152Q2Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA93357P() As Boolean
    
    'Const lcAddrX As Long = 10
    Const lcAddrX As Long = 22
    Const lcAddrLine As Long = 21
    On Error GoTo GetPage_TA93357P_Error
    
    If m_lCurrentPageIdx >= 10 Then
        'If Mid(m_saCurrentPageLine(10), 42, 8) = "OFFICER:" Then
        If InStr(42, m_saCurrentPageLine(10), "OFFICER:", vbTextCompare) > 0 Then
            'Fix issue where Signature erroneously places a TA95544P1 notice in the TA93357P1 file.
            'ToDo: This may not work if the erroneously placed notice is a multi-page document.
            GetPage_TA93357P = GetPage_TA9544P()
            Exit Function
        End If
    End If
    
    If m_lCurrentPageIdx >= 11 Then
        'If Mid(m_saCurrentPageLine(10), 42, 8) = "OFFICER:" Then
        If InStr(42, m_saCurrentPageLine(11), "OFFICER:", vbTextCompare) > 0 Then
            'Fix issue where Signature erroneously places a TA95544P1 notice in the TA93357P1 file.
            'ToDo: This may not work if the erroneously placed notice is a multi-page document.
            GetPage_TA93357P = GetPage_TA9544P()
            Exit Function
        End If
    End If
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrLine + 4 Then
        m_lCurrentPageIdx = lcAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 19, 3
    
    GetPage_TA93357P = True
    
    Exit Function
    
GetPage_TA93357P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA93357P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA93357P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function GetPage_TA93357PTrailer() As Boolean
'
'    Const lcAddrX = 5
'    Const lcAddrLine = 11
'    Const lcAddrLen = 40
'
'    GetPage_TA93357PTrailer = False
'
'    On Error GoTo GetPage_TA93357PTrailer_Error
'
'    'Assume a trailer page unless page 1 is found.
'
'    If m_lReadAheadPageIdx >= 27 Then
'        If Trim(Mid(m_saReadAheadPageLine(27), 73)) = "1" Then
'            Exit Function
'        End If
'    End If
'
'    ReadPageAhead
'
'    Set m_oDeliveryAddress = Nothing
'    Set m_oReturnAddress = Nothing
'
''    Dim i As Long
''    For i = 0 To 8
''        If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
''            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Line 1 through 9 must be blank on a T152Q2.  The following data will be dropped:" & vbCrLf & vbCrLf & "'" & Trim(m_saCurrentPageLine(i)) & "'."
''        End If
''    Next
''    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
'
'    GetPage_TA93357PTrailer = True
'
'    Exit Function
'
'GetPage_TA93357PTrailer_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_TA93357PTrailer"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_TA93357PTrailer_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_TA93357PTrailer_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Function GetPage_TA93357PTrailer() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_TA93357PTrailer = False
    
    On Error GoTo GetPage_TA93357PTrailer_Error
    
    Dim bContinued As Boolean
    
    Dim i As Long
    For i = m_lCurrentPageIdx To 0 Step -1
        If InStr(1, m_saCurrentPageLine(i), "END OF STATEMENT", vbTextCompare) > 0 Then
            Exit Function
        End If
        
        If InStr(1, m_saCurrentPageLine(i), "CONTINUED ON NEXT PAGE", vbTextCompare) > 0 Then
            bContinued = True
            Exit For
        End If
    Next
    
    If Not bContinued Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "'END OF STATEMENT' or 'CONTINUED ON NEXT PAGE' not found."
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_TA93357PTrailer = True
    
    Exit Function
    
GetPage_TA93357PTrailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA93357PTrailer"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA93357PTrailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA93357PTrailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA990771PTrailer() As Boolean
    GetPage_TA990771PTrailer = False
    
    Dim oRegex As RegExp
    Dim oMatch As Match
    Dim lPageNum As Long
    Dim lTotalPageNum As Long
    Dim sReadAcctNum As String
    
    On Error GoTo GetPage_TA990771PTrailer_Error

    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    ' Check if read ahead page is another page in this piece, exit function if it is
    sReadAcctNum = Trim(m_saReadAheadPageLine(1))
    
    If sReadAcctNum = m_sAccountNumber Then
        ' If they match - scrape the address
        Set m_oDeliveryAddress = GetDeliveryAddress(m_saReadAheadPageLine, m_lReadAheadPageIdx, 8, 16, 40, True, False, False, True)
        'Set m_oDeliveryAddress = Nothing
    Else
        ' Account numbers don't match, end this piece
        Exit Function
    End If
    
    ReadPageAhead
    
    'Set m_oDeliveryAddress = Nothing
    'Set m_oReturnAddress = Nothing
    
    GetPage_TA990771PTrailer = True
    
    Exit Function
    
GetPage_TA990771PTrailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA990771PTrailer"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA990771PTrailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA990771PTrailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function GetPage_TA9544P() As Boolean
'
'    Const lcAddrX As Long = 10
'    Const lcAddrLine As Long = 20
'    On Error GoTo GetPage_TA9544P_Error
'
'    InitializePieceFields
'
'    If m_lCurrentPageIdx < lcAddrLine + 4 Then
'        m_lCurrentPageIdx = lcAddrLine + 4
'        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
'        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
'    End If
'
'    If InStr(40, m_saCurrentPageLine(1), "OFFICER: ", vbTextCompare) > 0 Then
'        'Fixing signature bug where these lines are missing on the first form.
'        InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 1, 9
'    End If
'
'    If m_lCurrentPageIdx < 27 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 28 is missing."
'    End If
'
'    If Not IsNumeric(Trim(Mid(m_saCurrentPageLine(27), 73))) Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 28 is missing."
'    End If
'
'    If Val(Trim(Mid(m_saCurrentPageLine(27), 73))) <> 1 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 28 is missing."
'    End If
'
'    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
'
'    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
'    Set m_oCassAddress = m_oDeliveryAddress.Clone
'
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
'
'    If m_lCurrentPageIdx >= 19 Then
'        If InStr(1, m_saCurrentPageLine(19), "PLEASE RETURN THE TOP PORTION OF THIS BILLING WITH YOUR PAYMENT.  IF", vbTextCompare) > 0 Then
'            'Fixing Signature bug where needed blank line is sometimes missing.  Move this down one line.
'            InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 19, 1
'        End If
'    End If
'
'    Dim i As Long
'    Dim lDeleteLineCount As Long
'    For i = 20 To m_lCurrentPageIdx
'        If Len(Trim(m_saCurrentPageLine(i))) = 0 Then
'            lDeleteLineCount = lDeleteLineCount + 1
'        Else
''            If InStr(1, m_saCurrentPageLine(i), "PLEASE RETURN THE TOP PORTION OF THIS BILLING WITH YOUR PAYMENT.  IF", vbTextCompare) = 0 Then
''                'Only delete lines if the "please return the top portion" text is found.
''                lDeleteLineCount = 0
''            End If
'
'            Exit For
'        End If
'    Next
'
'    If lDeleteLineCount > 0 Then
'        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 20, lDeleteLineCount
'    End If
'
'
'    GetPage_TA9544P = True
'
'    Exit Function
'
'GetPage_TA9544P_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_TA9544P"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_TA9544P_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_TA9544P_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Function GetPage_TA9544P() As Boolean
    Dim i As Long
    
    On Error GoTo GetPage_TA9544P_Error
    
    InitializePieceFields
    
    '------------------------------------------------------------------------------------------------------
    ' The following insert lines logic fixes a Signature bug where the first piece in a TA95442P2 file
    ' is missing blank lines at the top of the page.  Subsequent pieces in the same file are fine.
    If m_lCurrentPageIdx >= 1 Then
        If InStr(1, m_saCurrentPageLine(1), "OFFICER:", vbTextCompare) > 0 Then
           InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 1, 9
            If m_lCurrentPageIdx >= 28 Then
                If InStr(1, m_saCurrentPageLine(28), "PLEASE RETURN THE TOP PORTION OF THIS BILLING WITH YOUR PAYMENT.", vbTextCompare) > 0 Then
                   InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 28, 3
                End If
            End If
        End If
    End If
    
    If m_lCurrentPageIdx >= 1 Then
        If InStr(1, m_saCurrentPageLine(1), "MONTHLY MORTGAGE STATEMENT", vbTextCompare) > 0 Then
           InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 1, 8
            If m_lCurrentPageIdx >= 28 Then
                If InStr(1, m_saCurrentPageLine(28), "PLEASE RETURN THE TOP PORTION OF THIS BILLING WITH YOUR PAYMENT.", vbTextCompare) > 0 Then
                   InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 28, 3
                End If
            End If
        End If
    End If
    
    If m_lCurrentPageIdx >= 1 Then
        If InStr(1, m_saCurrentPageLine(1), "MORTGAGE STATEMENT", vbTextCompare) > 0 Then
           'This kludge fixes an issue where all the vertical white space is missing in a MORTGAGE STATEMENT
           'Move the title down to line 6 where we expect it.
           InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 1, 4
           'Move the address block down to line 20 where we expect it.
           InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 14, 5
           'Find the PLEASE RETURN line and move it down to line 33 where we expect it.
           For i = 20 To 31
            If i > m_lCurrentPageIdx Then
              Exit For
            End If
            If InStr(1, m_saCurrentPageLine(i), "PLEASE RETURN THE TOP PORTION OF THIS BILLING WITH YOUR PAYMENT.", vbTextCompare) > 0 Then
                If Trim(m_saCurrentPageLine(i - 1)) = "* * * IMPORTANT NOTICE * * *" Then
                    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, i - 1, 32 - 22 - IIf(m_lCurrentPageIdx > 76, (m_lCurrentPageIdx - 76), 0) 'after 76 we run off the page
                Else
                    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, i, 32 - 23
               End If
               Exit For
            End If
           Next
        End If
    End If
    '------------------------------------------------------------------------------------------------------
    
'    Dim lAddressLine1Idx As Long
'    Dim lAddressLinePos As Long
'    Dim lPageNumberLineIdx As Long
    
'    lAddressLine1Idx = -1
'    lAddressLinePos = -1
    
    
'    lPageNumberLineIdx = -1
    
'    For i = 5 To m_lCurrentPageIdx
'        If InStr(50, m_saCurrentPageLine(i), "PAYMENT DUE DATE:", vbTextCompare) > 0 Then
'            'lAddressLine1Idx = i - 4
'            lPageNumberLineIdx = i + 3
'            Exit For
'        End If
'    Next
    
'    For i = 19 To 25
'        If Len(Trim(Left(m_saCurrentPageLine(i), 30))) > 0 Then
'            'If there is something other than blanks in the first 30 characters it's a good GUESS that this
'            'is the first address line.  Signature moves the address up and down randomly just to be cruel --
'            'reminds me of Lucy pulling the football away from Charlie Brown at the last second.
'            lAddressLine1Idx = i
'
'            'Oh, and just to make things interesting, lets move it left and right at random too!  After all,
'            'everyone has lots a people sitting around with nothing better to do than to find envelopes that work
'            'with random address placements, right?  Might as well give them something to do.
'            Dim sLetsFindWhereItStarts As String
'            sLetsFindWhereItStarts = Left(m_saCurrentPageLine(i), 30)
'            For lAddressLinePos = 1 To Len(sLetsFindWhereItStarts)
'                If Len(Trim(Mid(sLetsFindWhereItStarts, lAddressLinePos, 1))) > 0 Then
'                    Exit For
'                End If
'            Next
'
'            Exit For
'        End If
'    Next
'
'
'    If lAddressLine1Idx < 0 Then
'        'Hey Signature, I see your programmers have been busy again.  Don't worry about telling us where you moved the address block to
'        'this time.  We have an abundance of progammmers sitting around with nothing to do so we'll take care of it.  You just relax and
'        'enjoy your umbrella drink.
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Can't find the address block."
'    End If
    
    'Try to find the account number.   Good luck, they call it anything and put it wherever -- doesn't seem to matter.
    Const scLoanNumber As String = "Loan Number"
    Const scAccountNumber As String = "ACCOUNT NUMBER"
    Const scAccountNumber2 As String = "Account Number"
    Dim lPos As Long
    'For i = lPageNumberLineIdx To 0 Step -1
    For i = 1 To 25
        If i >= m_lCurrentPageIdx Then
            Exit For
        End If
        lPos = InStr(1, m_saCurrentPageLine(i), scLoanNumber, vbTextCompare)
        If lPos > 0 Then
            m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scLoanNumber)))
            Exit For
        End If
        lPos = InStr(1, m_saCurrentPageLine(i), scAccountNumber, vbTextCompare)
        If lPos > 0 Then
            m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scAccountNumber)))
            Exit For
        End If
        lPos = InStr(1, m_saCurrentPageLine(i), scAccountNumber2, vbTextCompare)
        If lPos > 0 Then
            m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scAccountNumber2)))
            Exit For
        End If
    Next
    
    If Len(Trim(m_sAccountNumber)) = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Account number not found."
    End If
    
    'Try to find the payment amount.   Good luck, they call it anything and put it wherever -- doesn't seem to matter.  Who
    'reads their statements anyway.
    Const scPaymentAmount As String = "PAYMENT AMOUNT:"
    Const scTotalDue As String = "TOTAL DUE:"
    Const scAmountDue As String = "** AMOUNT DUE"
    Dim bPaymentAmountFound As Boolean
    bPaymentAmountFound = False
    For i = 1 To 25
        lPos = InStr(1, m_saCurrentPageLine(i), scPaymentAmount, vbTextCompare)
        If lPos > 0 Then
            bPaymentAmountFound = True
            m_cPaymentAmount = CCur(Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scPaymentAmount))))
            Exit For
        End If
        lPos = InStr(1, m_saCurrentPageLine(i), scTotalDue, vbTextCompare)
        If lPos > 0 Then
            bPaymentAmountFound = True
            m_cPaymentAmount = CCur(Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scTotalDue))))
            Exit For
        End If
        lPos = InStr(1, m_saCurrentPageLine(i), scAmountDue, vbTextCompare)
        If lPos > 0 Then
            m_cPaymentAmount = CCur(Trim(Mid(m_saCurrentPageLine(i), lPos + Len(scAmountDue), 19)))
            bPaymentAmountFound = True
            Exit For
        End If
    Next
    
    If Not bPaymentAmountFound Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Payment amount not found."
    End If
    
'    If lAddressLine1Idx < 0 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "PAYMENT DUE DATE: is missing."
'        GetPage_TA9544P = False
'        Exit Function
'    End If
    
'    If Val(Trim(Mid(m_saCurrentPageLine(lPageNumberLineIdx), 73))) <> 1 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number not equal to one on first page of piece."
'    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Const lcAddrX As Long = 10
    Const lcAddrLine As Long = 20
    
    Dim lAddrX As Long
    lAddrX = 22
    
    
    'Address block is shifting around.  Look for it in 10 or 22.
    Dim lAddrLineIdx As Long
    lAddrLineIdx = lcAddrLine - 1
    For i = lAddrLineIdx To lAddrLineIdx + 5
     If i > m_lCurrentPageIdx Then
       Exit For
     End If
     If Len(Trim(Left(m_saCurrentPageLine(i), lAddrX - 1))) > 0 Then
        lAddrX = lcAddrX
        Exit For
     End If
    Next
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    GetPage_TA9544P = True
    
    Exit Function
    
GetPage_TA9544P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA9544P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA9544P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA9544P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function GetPage_TA9544PTrailer() As Boolean
'
'    Const lcAddrX = 5
'    Const lcAddrLine = 11
'    Const lcAddrLen = 40
'
'    GetPage_TA9544PTrailer = False
'
'    On Error GoTo GetPage_TA9544PTrailer_Error
'
'    'Assume a trailer page unless page 1 is found.
'
'    If m_lReadAheadPageIdx > 27 Then
'        If IsNumeric(Trim(Mid(m_saReadAheadPageLine(27), 73))) Then
'            If Val(Trim(Mid(m_saReadAheadPageLine(27), 73))) < 2 Then
'                Exit Function
'            End If
'        End If
'    End If
'
'    ReadPageAhead
'
'    Set m_oDeliveryAddress = Nothing
'    Set m_oReturnAddress = Nothing
'
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
'
'    GetPage_TA9544PTrailer = True
'
'    Exit Function
'
'GetPage_TA9544PTrailer_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_TA9544PTrailer"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_TA9544PTrailer_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_TA9544PTrailer_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Function GetPage_TA9544PTrailer() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_TA9544PTrailer = False
    
    On Error GoTo GetPage_TA9544PTrailer_Error
    
    Dim bContinued As Boolean
    
    Dim i As Long
    For i = m_lCurrentPageIdx To 0 Step -1
        If InStr(1, m_saCurrentPageLine(i), "END OF STATEMENT", vbTextCompare) > 0 Then
            Exit Function
        End If
        
        If InStr(1, m_saCurrentPageLine(i), "CONTINUED ON NEXT PAGE", vbTextCompare) > 0 Then
            bContinued = True
            Exit For
        End If
    Next
    
    If Not bContinued Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "'END OF STATEMENT' or 'CONTINUED ON NEXT PAGE' not found."
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
'    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 9
    
    GetPage_TA9544PTrailer = True
    
    Exit Function
    
GetPage_TA9544PTrailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA9544PTrailer"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA9544PTrailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA9544PTrailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0860P1(Optional ByVal bRemoveBankAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    Const lcAddrX As Long = 10
    Const lcAddrLine As Long = 18
       
    
    On Error GoTo GetPage_LN0860P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrLine + 5 Then
        m_lCurrentPageIdx = lcAddrLine + 5
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Validates that a page number is found on the piece
    Const scPage As String = "PAGE "
    Dim lPos As Long
    lPos = InStr(75, m_saCurrentPageLine(2), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
    End If
    
    ' Validates that the page we are working on is a lead page and not a trailer page
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saCurrentPageLine(2), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
    End If
    If Val(sWorkPage) > 1 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number 1 expected on line 3."
    End If

    If InStr(1, m_saCurrentPageLine(17), "LOAN NUMBER", vbTextCompare) = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "LOAN NUMBER not found on line 16."
    End If
    
    m_saCurrentPageLine(17) = "    " & m_saCurrentPageLine(17)
    
    If bRemoveBankAddr Then
        Mid(m_saCurrentPageLine(2), 15, 40) = Space(40)
        Mid(m_saCurrentPageLine(3), 15, 40) = Space(40)
        Mid(m_saCurrentPageLine(4), 15, 40) = Space(40)
        Mid(m_saCurrentPageLine(5), 15, 40) = Space(40)
        Mid(m_saCurrentPageLine(6), 15, 40) = Space(40)
    End If
    
    ' Moves the page number down IF we are Investors Bank
    If m_sXmlInst = "Investors" Then
        m_saCurrentPageLine(6) = m_saCurrentPageLine(2)
        m_saCurrentPageLine(2) = vbNullString
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex object
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
    ' Gets the account number off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+LOAN NUMBER:\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(17)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on LN0860P1 page."
        GetPage_LN0860P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(17))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(17) = Replace(m_saCurrentPageLine(17), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}/\d{1,2}/\d{2})"
    If oRegexTest.Test(m_saCurrentPageLine(13)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on LN0860P1 page."
        GetPage_LN0860P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(13))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 2
    'ScrollLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 7, 2
    
    ScrollLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 5, 8, 3
    'ScrollLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 7, 10, 3
    
    If Len(m_sXML_EddJobKey) > 0 Then
'        Dim sAccountNumber As String
'        If Len(m_sAccountNumber) < 10 Then
'            sAccountNumber = Right(String(10, "0") & m_sAccountNumber, 10)
'        Else
'            sAccountNumber = m_sAccountNumber
'        End If
        
        m_lIndexIdx = 3
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
        
        m_saIndexKeys(0) = "AccountNumber"
        m_saIndexValues(0) = m_sAccountNumber
        
        m_saIndexKeys(1) = "CustomerName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name
        
        m_saIndexKeys(2) = "StatementDate"
        'm_saIndexValues(2) = m_sDocumentDate
        m_saIndexValues(2) = Format(oDocDate, "yyyyMMdd")
        
        m_saIndexKeys(3) = "AccountTypeNumber"
        m_saIndexValues(3) = m_sXML_EddProductId & "-" & m_sAccountNumber
        
    End If
    
    GetPage_LN0860P1 = True
    
    Exit Function
    
GetPage_LN0860P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0860P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0860P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0860P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0860P1Trailer(Optional ByVal bRemoveBankAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_LN0860P1Trailer = False
    
    On Error GoTo GetPage_LN0860P1Trailer_Error
    
    If m_lReadAheadPageIdx < 2 Then
        Exit Function
    End If
    
    Const scPage As String = "PAGE "
    Dim lPos As Long
    lPos = InStr(74, m_saReadAheadPageLine(2), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
    End If
    
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saReadAheadPageLine(2), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
    End If
    If Val(sWorkPage) < 2 Then
        Exit Function
    End If
    
    Dim sMaskedAcctNum As String
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saReadAheadPageLine(8) = Replace(m_saReadAheadPageLine(8), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    If bRemoveBankAddr Then
        Mid(m_saReadAheadPageLine(2), 15, 40) = Space(40)
        Mid(m_saReadAheadPageLine(3), 15, 40) = Space(40)
        Mid(m_saReadAheadPageLine(4), 15, 40) = Space(40)
        Mid(m_saReadAheadPageLine(5), 15, 40) = Space(40)
        Mid(m_saReadAheadPageLine(6), 15, 40) = Space(40)
    End If
    
    ' If we are investors - relocate the page number so it doesn't overprint on their custom stock
    
    If m_sXmlInst = "Investors" Then
        m_saReadAheadPageLine(4) = m_saReadAheadPageLine(2)
        m_saReadAheadPageLine(2) = vbNullString
    End If
    
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_LN0860P1Trailer = True
    
    Exit Function
    
GetPage_LN0860P1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0860P1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0860P1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'Private Function GetPage_PRINT1(Optional ByVal bRemoveBankAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean
'
'    Const lcAddrX As Long = 15
'    Const lcAddrLine As Long = 10
'
'
'    On Error GoTo GetPage_PRINT1_Error
'
'    InitializePieceFields
'
'    If m_lCurrentPageIdx < lcAddrLine + 5 Then
'        m_lCurrentPageIdx = lcAddrLine + 5
'        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
'        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
'    End If
'
'    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, False, True, False)
'    Set m_oCassAddress = m_oDeliveryAddress.Clone
'
'    ' Validates that a page number is found on the piece
'    Const scPage As String = "PAGE "
'    Dim lPos As Long
'    lPos = InStr(68, m_saCurrentPageLine(1), scPage, vbTextCompare)
'    If lPos = 0 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
'    End If
'
'    ' Validates that the page we are working on is a lead page and not a trailer page
'    Dim sWorkPage As String
'    sWorkPage = Trim(Mid(m_saCurrentPageLine(1), lPos + Len(scPage)))
'
'    If Not IsNumeric(sWorkPage) Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 3 is missing."
'    End If
'    If Val(sWorkPage) > 1 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number 1 expected on line 3."
'    End If
'
'    If InStr(1, m_saCurrentPageLine(9), "Account:", vbTextCompare) = 0 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Account Number not found on line 10."
'    End If
'
'    If bRemoveBankAddr Then
'        Mid(m_saCurrentPageLine(1), 5, 40) = Space(40)
'        Mid(m_saCurrentPageLine(2), 5, 40) = Space(40)
'        Mid(m_saCurrentPageLine(3), 5, 40) = Space(40)
'    End If
'
'    ' Variables used for validating the page
'    Dim oRegexTest As RegExp
'    Dim oRegexMatches As MatchCollection
'    Dim oRegexMatch As Match
'    Dim sMaskedAcctNum As String
'    Dim sRawDate As String
'    Dim oDocDate As Date
'
'    ' Inits the regex object
'    Set oRegexTest = New RegExp
'    oRegexTest.IgnoreCase = True
'
'    ' Gets the account number off the page and verifies that it can be found
'    oRegexTest.Pattern = "^\s+Account:\s+(\d+)"
'    If oRegexTest.Test(m_saCurrentPageLine(9)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on PRINT1 page."
'        GetPage_PRINT1 = False
'        Exit Function
'    End If
'
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(9))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    m_sAccountNumber = oRegexMatch.SubMatches(0)
'
'    ' Masks the account number if we are to mask it
'    If bMaskAcctNum Then
'        If Len(m_sAccountNumber) > 4 Then
'            sMaskedAcctNum = m_sAccountNumber
'            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
'            m_saCurrentPageLine(9) = Replace(m_saCurrentPageLine(9), m_sAccountNumber, sMaskedAcctNum)
'        End If
'    End If
'
'    ' Check to see if there is a "summary of accounts" area - this means there is a lead account number
'    Dim i As Long
'    For i = 0 To m_lCurrentPageIdx
'        oRegexTest.Pattern = "ANALYSIS STATEMENT"
'
'        ' If we have hit the Analysis statement title - then there is no summary of accounts section
'        If oRegexTest.Test(m_saCurrentPageLine(i)) = True Then
'            Exit For
'        End If
'
'        ' Check for if we have hit the summary of accounts
'        oRegexTest.Pattern = "\s+Summary of Accounts:"
'        If oRegexTest.Test(m_saCurrentPageLine(i)) = True Then
'            oRegexTest.Pattern = "\s+Lead Account:\s+(\d+)"
'
'            ' Check if there is a lead account on the next line and it isn't off the page
'            If i + 1 > m_lCurrentPageIdx Then
'                Err.Raise g_lcErrBase + g_lcErrInvalidPage, "GetPage_PRINT1", "Invalid page encountered - Summary of Accounts found but no lead account could be found."
'            End If
'
'            Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(i + 1))
'            Set oRegexMatch = oRegexMatches.Item(0)
'            m_sAccountNumber = oRegexMatch.SubMatches(0)
'
'            Exit For
'        End If
'    Next
'
'    ' Gets the date received/document date off the page and verifies that it can be found
'    oRegexTest.Pattern = " To:\s+(\d{1,2}-\d{1,2}-\d{2})"
'    If oRegexTest.Test(m_saCurrentPageLine(6)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on PRINT1 page."
'        GetPage_PRINT1 = False
'        Exit Function
'    End If
'
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(6))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    sRawDate = oRegexMatch.SubMatches.Item(0)
'    oDocDate = CDate(sRawDate)
'    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
'
'    If Len(m_sXML_EddJobKey) > 0 Then
'        Dim sAccountNumber As String
'        If Len(m_sAccountNumber) < 10 Then
'            sAccountNumber = Right(String(10, "0") & m_sAccountNumber, 10)
'        Else
'            sAccountNumber = m_sAccountNumber
'        End If
'
'        m_lIndexIdx = 3
'        ReDim m_saIndexKeys(m_lIndexIdx)
'        ReDim m_saIndexValues(m_lIndexIdx)
'
'        m_saIndexKeys(0) = "AccountNumber"
'        m_saIndexValues(0) = sAccountNumber
'
'        m_saIndexKeys(1) = "CustomerName"
'        m_saIndexValues(1) = m_oDeliveryAddress.Name
'
'        m_saIndexKeys(2) = "StatementDate"
'        'm_saIndexValues(2) = m_sDocumentDate
'        m_saIndexValues(2) = Format(oDocDate, "yyyyMMdd")
'
'        m_saIndexKeys(3) = "AccountTypeNumber"
'        m_saIndexValues(3) = m_sXML_EddProductId & "-" & sAccountNumber
'
'    End If
'
''    oRegexTest.Pattern = "(ANALYSIS STATEMENT)|(Lead Account:)|(Summary of Accounts:)"
''
''    If oRegexTest.Test(m_saCurrentPageLine(16)) = False Then
''        Err.Raise g_lcErrBase + g_lcErrInvalidPage, "GetPage_PRINT1", " INDEX 16 DOES NOT EQUAL ANALYSIS STATEMENT."
''    Else
''
''    End If
'
'
'    GetPage_PRINT1 = True
'
'    Exit Function
'
'GetPage_PRINT1_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_PRINT1"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_PRINT1_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_PRINT1_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

'Private Function GetPage_PRINT1Trailer(Optional ByVal bRemoveBankAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean
'
'    Const lcAddrX = 15
'    Const lcAddrLine = 10
'    Const lcAddrLen = 40
'
'    GetPage_PRINT1Trailer = False
'
'    On Error GoTo GetPage_PRINT1Trailer_Error
'
'    ' Variables used for validating the page
'    Dim oRegexTest As RegExp
'    Dim oRegexMatches As MatchCollection
'    Dim oRegexMatch As Match
'
'    Set oRegexTest = New RegExp
'    oRegexTest.IgnoreCase = True
'
'    Dim bIsGrouped As Boolean
'    bIsGrouped = False
'
'    If m_lReadAheadPageIdx < 2 Then
'        Exit Function
'    End If
'
'    Const scPage As String = "PAGE "
'    Dim lPos As Long
'    Dim lLine As Long
'    ' This checks for the page on line 2
'    lPos = InStr(68, m_saReadAheadPageLine(1), scPage, vbTextCompare)
'    If lPos = 0 Then
'        ' Check for the page on line 1 - trailer pages have their page number on line 1
'        lPos = InStr(68, m_saReadAheadPageLine(0), scPage, vbTextCompare)
'        If lPos = 0 Then
'            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 1 or 2 is missing."
'        Else
'            lLine = 0
'        End If
'    Else
'        lLine = 1
'    End If
'
'    Dim sWorkPage As String
'    sWorkPage = Trim(Mid(m_saReadAheadPageLine(lLine), lPos + Len(scPage)))
'
'    If Not IsNumeric(sWorkPage) Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 2 is missing."
'    End If
'
'    ' If we are a page 1 - check if this account should be grouped
'    If Val(sWorkPage) = 1 Then
'        ' Check to see if there is a "summary of accounts" area - this means there is a lead account number
'        Dim i As Long
'        For i = 0 To m_lReadAheadPageIdx
'            oRegexTest.Pattern = "\s+ANALYSIS STATEMENT"
'
'            ' If we have hit the Analysis statement title - then there is no summary of accounts section
'            If oRegexTest.Test(m_saReadAheadPageLine(i)) = True Then
'                Exit For
'            End If
'
'            ' Check for if we have hit the summary of accounts
'            oRegexTest.Pattern = "\s+Summary of Accounts:"
'            If oRegexTest.Test(m_saReadAheadPageLine(i)) = True Then
'                oRegexTest.Pattern = "\s+Lead Account:\s+(\d+)"
'
'                ' Check if there is a lead account on the next line and it isn't off the page
'                If i + 1 > m_lCurrentPageIdx Then
'                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, "GetPage_PRINT1", "Invalid page encountered - Summary of Accounts but was at the end of a piece."
'                End If
'                If oRegexTest.Test(m_saReadAheadPageLine(i + 1)) = False Then
'                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, "GetPage_PRINT1", "Invalid page encountered - No Lead Account could be found."
'                End If
'
'                Set oRegexMatches = oRegexTest.Execute(m_saReadAheadPageLine(i + 1))
'                Set oRegexMatch = oRegexMatches.Item(0)
'
'                If m_sAccountNumber = oRegexMatch.SubMatches(0) Then
'                    bIsGrouped = True
'                End If
'
'                Exit For
'            End If
'        Next
'    End If
'
'    If Val(sWorkPage) < 2 And Not bIsGrouped Then
'        Exit Function
'    End If
'
'    Dim sMaskedAcctNum As String
'
'    ' Masks the account number if we are to mask it
'    If bMaskAcctNum Then
'        If Len(m_sAccountNumber) > 4 Then
'            sMaskedAcctNum = m_sAccountNumber
'            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
'            m_saReadAheadPageLine(9) = Replace(m_saReadAheadPageLine(9), m_sAccountNumber, sMaskedAcctNum)
'        End If
'    End If
'
'    If bRemoveBankAddr Then
'        Mid(m_saReadAheadPageLine(1), 5, 40) = Space(40)
'        Mid(m_saReadAheadPageLine(2), 5, 40) = Space(40)
'        Mid(m_saReadAheadPageLine(3), 5, 40) = Space(40)
'    End If
'
'    ' If we are investors - relocate the page number so it doesn't overprint on their custom stock
'
'    ReadPageAhead
'
'    Set m_oDeliveryAddress = Nothing
'    Set m_oReturnAddress = Nothing
'
'    GetPage_PRINT1Trailer = True
'
'    Exit Function
'
'GetPage_PRINT1Trailer_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo GetPage_PRINT1Trailer_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo GetPage_PRINT1Trailer_Error
'          Resume Next
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

Private Function GetPage_LN0472P1() As Boolean
    
    Const lcAddrX As Long = 10
    Const lcAddrLine As Long = 14
    On Error GoTo GetPage_LN0472P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrLine + 5 Then
        m_lCurrentPageIdx = lcAddrLine + 5
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    m_saCurrentPageLine(6) = Mid(m_saCurrentPageLine(6), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(6), 37)
    m_saCurrentPageLine(7) = Mid(m_saCurrentPageLine(7), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(7), 37)
    m_saCurrentPageLine(8) = Mid(m_saCurrentPageLine(8), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(8), 37)
    
    If InStr(1, m_saCurrentPageLine(61), "End of Notice", vbTextCompare) = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "End of Notice not found."
    End If
    
    'ToDo: Assert 'End of Notice' present.
    '       Erase return address.
    
    
    GetPage_LN0472P1 = True
    
    Exit Function
    
GetPage_LN0472P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0472P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0472P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0472P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0473P1(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    Const lcAddrX As Long = 10
    Const lcAddrLine As Long = 14
    
    Dim sMaskedAcctNum As String
    
    On Error GoTo GetPage_LN0473P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrLine + 5 Then
        m_lCurrentPageIdx = lcAddrLine + 5
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
            
    Const scAccountNumber = "Account Number: "
    If Mid(m_saCurrentPageLine(1), 52, Len(scAccountNumber)) = scAccountNumber Then
        m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(1), 52 + Len(scAccountNumber)))
        
        ' If we are to mask the account number - do it here
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(1) = Replace(m_saCurrentPageLine(1), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
    End If
    
    Const scNoticeDate = "Notice Date:"
    If Trim(Mid(m_saCurrentPageLine(0), 52, 12)) = scNoticeDate Then
        m_sDocumentDate = Trim(Mid(m_saCurrentPageLine(0), 52 + Len(scNoticeDate)))
    End If
    
    If Not IsDate(m_sDocumentDate) Then
        m_sDocumentDate = vbNullString
    Else
        m_sDocumentDate = Format(CDate(m_sDocumentDate), "yyyy-MM-dd")
        m_sEddCycleDate = CDate(m_sDocumentDate)
    End If

    If Len(m_sAccountNumber) = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Account number not found."
    End If
    
    'Erase return address
    m_saCurrentPageLine(0) = Mid(m_saCurrentPageLine(0), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(0), 37)
    m_saCurrentPageLine(1) = Mid(m_saCurrentPageLine(1), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(1), 37)
    m_saCurrentPageLine(2) = Mid(m_saCurrentPageLine(2), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(2), 37)
    
    If m_sXmlInst = "BPB" Then
        BuildBpbPieceIndex m_sCurrentFileName, m_sAccountNumber, m_sDocumentDate, "LN0473-ARM", m_oDeliveryAddress.Name
    Else
        m_lIndexIdx = 3
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
    
        m_saIndexKeys(0) = "Account_Number"
        m_saIndexValues(0) = m_sAccountNumber
    
        m_saIndexKeys(1) = "PrimaryName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name
    
        m_saIndexKeys(2) = "Notice_Date"
        m_saIndexValues(2) = m_sDocumentDate
    
        m_saIndexKeys(3) = "Notice_Type"
        m_saIndexValues(3) = "LN0473P1"
    End If
    
    GetPage_LN0473P1 = True
    
    Exit Function
    
GetPage_LN0473P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0473P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0473P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0473P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0473P1Trailer(Optional ByVal bMaskAcctNum As Boolean) As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_LN0473P1Trailer = False
    
    On Error GoTo GetPage_LN0473P1Trailer_Error
    
    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    Const scPage As String = "Page "
    Dim lPage As Long
    If Mid(m_saReadAheadPageLine(0), 71, Len(scPage)) = scPage Then
        lPage = CLng(Trim(Mid(m_saReadAheadPageLine(0), 71 + Len(scPage))))
        If lPage < 2 Then
            Exit Function
        End If
    Else
        Exit Function
    End If
            
    Const scAccountNumber = "Account Number:"
    Dim sAccountNumber As String
    If Mid(m_saReadAheadPageLine(1), 52, Len(scAccountNumber)) = scAccountNumber Then
        sAccountNumber = Trim(Mid(m_saReadAheadPageLine(1), 52 + Len(scAccountNumber)))
    End If
    
    If sAccountNumber <> m_sAccountNumber Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Account number mismatch."
    End If
    
    Dim sMaskedAcctNum As String
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saReadAheadPageLine(1) = Replace(m_saReadAheadPageLine(1), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ReadPageAhead
    
    'Erase return address
    m_saCurrentPageLine(0) = Mid(m_saCurrentPageLine(0), 1, 1) & Space(35) & Mid(m_saCurrentPageLine(0), 37)
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_LN0473P1Trailer = True
    
    Exit Function
    
GetPage_LN0473P1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0473P1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0473P1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN395Q1(Optional ByVal bSuppressRetAddr As Boolean) As Boolean
    
    Const lcAddrX As Long = 11
    Const lcAddrLine As Long = 11
    On Error GoTo GetPage_LN395Q1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrLine + 4 Then
        m_lCurrentPageIdx = lcAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Should we suppress return-address printing
    If bSuppressRetAddr Then
        ScrapeString m_saCurrentPageLine(0), 11, 40
        ScrapeString m_saCurrentPageLine(1), 11, 40
        ScrapeString m_saCurrentPageLine(2), 11, 40
    End If
    
    ' Custom processing code for Boston Private Bank
    If m_sXmlInst = "BPB" Then
        Dim i As Long
        If InStr(1, m_saCurrentPageLine(10), "PROPERTY ADDRESS", vbTextCompare) > 0 Then
            For i = 10 To 16
                'Move property address to the right to get it out of the envelope window
                m_saCurrentPageLine(i) = "   " & m_saCurrentPageLine(i)
            Next
        End If
        
        '--------------------------------
        'Move letter text down to get it out of the envelope window.
        If Len(Trim(m_saCurrentPageLine(17))) > 0 Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data found on line 18 on an LN395Q1."
        End If
        
    '    If Left(Trim(m_saCurrentPageLine(18)), 4) <> "RE: " Then
    '        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data not found on line 19 on an LN395Q1."
    '    End If
        
        InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 18, 2
        '--------------------------------
    End If
        
    GetPage_LN395Q1 = True
    
    Exit Function
    
GetPage_LN395Q1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN395Q1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN395Q1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN395Q1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN395Q1CLET() As Boolean
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim bNoBottomMargin As Boolean
    Dim oRegExp As RegExp
    Dim oMatchCollection As MatchCollection
    On Error GoTo GetPage_LN395Q1CLET_Error
    
    lAddrX = 6
    lAddrLine = 13
    bNoBottomMargin = False
    Set oRegExp = New RegExp
    oRegExp.IgnoreCase = True
    
    Select Case m_sXmlInst
        Case "Nationwide"
            lAddrX = 1
            lAddrLine = 6
            bNoBottomMargin = True
    End Select
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, bNoBottomMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    If m_sXmlInst = "Nationwide" Then
        ' Get the acccount number
        oRegExp.Pattern = "Account Number\s+(\d+)"
        If oRegExp.Test(m_saCurrentPageLine(3)) = False Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate account number on LN395Q1 CLET page."
        End If
        
        Set oMatchCollection = oRegExp.Execute(m_saCurrentPageLine(3))
        m_sAccountNumber = oMatchCollection.Item(0).SubMatches(0)
        
        ' Get the document date
        oRegExp.Pattern = "Notice Date\s+(\d{2}/\d{2}/\d{2})"
        If oRegExp.Test(m_saCurrentPageLine(2)) = False Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate notice date on LN395Q1 CLET page."
        End If
        
        Set oMatchCollection = oRegExp.Execute(m_saCurrentPageLine(2))
        m_sDocumentDate = oMatchCollection.Item(0).SubMatches(0)
        
    End If
    
    GetPage_LN395Q1CLET = True
    
    Exit Function
    
GetPage_LN395Q1CLET_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN395Q1CLET"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN395Q1CLET_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN395Q1CLET_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN395Q1Trailer() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_LN395Q1Trailer = False
    
    On Error GoTo GetPage_LN395Q1Trailer_Error
    
'    If m_lReadAheadPageIdx < 0 Then
'        Exit Function
'    End If
    
    If InStr(1, m_saReadAheadPageLine(0), "PAGE ", vbTextCompare) = 0 Then
        Exit Function
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_LN395Q1Trailer = True
    
    Exit Function
    
GetPage_LN395Q1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN395Q1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN395Q1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_RA0080P1() As Boolean
    
    Const lcAddrX As Long = 9
    Const lcAddrLine As Long = 16
    On Error GoTo GetPage_RA0080P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < 20 Then
        m_lCurrentPageIdx = 20
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(7), 78, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, 40, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Const scPage As String = "Page "
    Dim lPos As Long
    lPos = InStr(50, m_saCurrentPageLine(20), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saCurrentPageLine(20), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
'    Dim i As Long
'    If InStr(1, m_saCurrentPageLine(10), "PROPERTY ADDRESS", vbTextCompare) > 0 Then
'        For i = 10 To 16
'            'Move property address to the right to get it out of the envelope window
'            m_saCurrentPageLine(i) = "   " & m_saCurrentPageLine(i)
'        Next
'    End If
    
'    '--------------------------------
'    'Move letter text down to get it out of the envelope window.
'    If Len(Trim(m_saCurrentPageLine(17))) > 0 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data found on line 18 on an RA0080P1."
'    End If
    
'    If Left(Trim(m_saCurrentPageLine(18)), 4) <> "RE: " Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unexpected data not found on line 19 on an RA0080P1."
'    End If
    
'    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 18, 2
    '--------------------------------
    
    GetPage_RA0080P1 = True
    
    Exit Function
    
GetPage_RA0080P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_RA0080P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_RA0080P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_RA0080P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_RA0080P1Trailer() As Boolean
    
    GetPage_RA0080P1Trailer = False
    
    On Error GoTo GetPage_RA0080P1Trailer_Error
    
    If m_lReadAheadPageIdx < 20 Then
        Exit Function
    End If
    
    Const scPage As String = "Page "
    Dim lPos As Long
    lPos = InStr(50, m_saReadAheadPageLine(20), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saReadAheadPageLine(20), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
    If Val(sWorkPage) <= 1 Then
        Exit Function
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_RA0080P1Trailer = True
    
    Exit Function
    
GetPage_RA0080P1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_RA0080P1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_RA0080P1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM0330P1Trailer() As Boolean
    
    GetPage_TM0330P1Trailer = False
    
    On Error GoTo GetPage_TM0330P1Trailer_Error
    
    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    If m_lStackItems = StackItemsPressureSeal Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Muti-page notice found for Pressure Seal job."
    End If
    
    Const scPage As String = "Page "
    Dim lPos As Long
    Dim sWorkPage As String
    
    lPos = InStr(50, m_saReadAheadPageLine(11), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 11 is missing."
    End If

    sWorkPage = Trim(Mid(m_saReadAheadPageLine(11), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number found on line 11 is not a numeric value.  Value=" & sWorkPage
    End If
    
    If Val(sWorkPage) <= 1 Then
        Exit Function
    End If
    
    ' If we get to this point then the read ahead page is part of the current piece - we read it into the buffer
    ' and return true
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    If m_sXmlInst = m_scInst_29175100_FinFactors Then
        ' Shift text right to get it out of the envelope window
        m_saCurrentPageLine(4) = Right$(Space(80) & Trim(m_saCurrentPageLine(4)), 80)
        m_saCurrentPageLine(5) = Right$(Space(80) & Trim(m_saCurrentPageLine(5)), 80)
        
        ' Scrape out the return address
        Const lcAddrX = 7
        Const lcAddrLen = 40
        ScrapeString m_saCurrentPageLine(0), lcAddrX, lcAddrLen, PvRTrimOutput
        ScrapeString m_saCurrentPageLine(1), lcAddrX, lcAddrLen, PvRTrimOutput
        ScrapeString m_saCurrentPageLine(2), lcAddrX, lcAddrLen, PvRTrimOutput
'
'        ' Prevent the return address from printing
'        If bKeepReturnAddress = False Then
'            Set m_oReturnAddress = Nothing
'        End If
    End If
    
    GetPage_TM0330P1Trailer = True
    
    Exit Function
    
GetPage_TM0330P1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0330P1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0330P1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_NoticeTrailer() As Boolean
    
    GetPage_NoticeTrailer = False
    
    On Error GoTo GetPage_NoticeTrailer_Error
    
    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    If m_lStackItems = StackItemsPressureSeal Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Muti-page notice found for Pressure Seal job."
    End If
    
    Const scPage As String = "Page "
    Dim lPos As Long
    lPos = InStr(50, m_saReadAheadPageLine(1), scPage, vbTextCompare)
    If lPos = 0 Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 1 is missing."
    End If
    
    Dim sWorkPage As String
    sWorkPage = Trim(Mid(m_saReadAheadPageLine(1), lPos + Len(scPage)))
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
    If Val(sWorkPage) <= 1 Then
        Exit Function
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_NoticeTrailer = True
    
    Exit Function
    
GetPage_NoticeTrailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_NoticeTrailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_NoticeTrailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_CH0202P1Trailer() As Boolean
    
    GetPage_CH0202P1Trailer = False
    
    On Error GoTo GetPage_CH0202P1Trailer_Error
    
    If m_lReadAheadPageIdx < 1 Then
        Exit Function
    End If
    
    If m_lStackItems = StackItemsPressureSeal Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Muti-page notice found for Pressure Seal job."
    End If
    
    Const scPage As String = "PAGE "
    Dim lPos As Long
    Dim sWorkPage As String
    
    lPos = InStr(50, m_saReadAheadPageLine(1), scPage, vbTextCompare)
    If lPos = 0 Then
        ' check line 2 for the page number
        lPos = InStr(50, m_saReadAheadPageLine(2), scPage, vbTextCompare)
        If lPos = 0 Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 1 is missing."
        Else
            sWorkPage = Trim(Mid(m_saReadAheadPageLine(2), lPos + Len(scPage)))
        End If
    Else
        sWorkPage = Trim(Mid(m_saReadAheadPageLine(1), lPos + Len(scPage)))
    End If
    
    If Not IsNumeric(sWorkPage) Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Page number expected on line 21 is missing."
    End If
    
    If Val(sWorkPage) <= 1 Then
        Exit Function
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    GetPage_CH0202P1Trailer = True
    
    Exit Function
    
GetPage_CH0202P1Trailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_CH0202P1Trailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_CH0202P1Trailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_SD0095P(Optional ByVal bMaskAcctNum As Boolean, Optional ByVal bSuppressRetAddr As Boolean) As Boolean
    
    On Error GoTo GetPage_SD0095P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    Select Case m_sCurrentFileName
        Case "SD0095P2", "SD0095P3"
            lAddrX = 16
            lAddrLine = 16
        Case "SD0095P4"
            lAddrX = 16
            lAddrLine = 15
    End Select
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Suppresses the bank address
    If bSuppressRetAddr Then
        Mid(m_saCurrentPageLine(0), 2, 40) = Space(40)
        Mid(m_saCurrentPageLine(1), 2, 40) = Space(40)
        Mid(m_saCurrentPageLine(2), 2, 40) = Space(40)
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim lWorkPage As Long
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex tester
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
'    ' Gets the page number and verifies that we are on the first page
'    oRegexTest.Pattern = "^\s+PAGE\s+(\d+)"
'    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find page number on SD0095P page."
'        GetPage_SD0095P = False
'        Exit Function
'    End If
'
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    lWorkPage = CLng(oRegexMatch.SubMatches.Item(0))
'
'    If lWorkPage > 1 Then
'        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - piece cannot begin with non-page 1 page on SD0095P page."
'        GetPage_SD0095P = False
'        Exit Function
'    End If
    
    ' Gets the account number
    If m_sCurrentFileName = "SD0095P4" Then
        oRegexTest.Pattern = " ACCOUNT\s+((\d+\s?)+)"
        If oRegexTest.Test(m_saCurrentPageLine(13)) = False Then
            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on SD0095P4 page."
            GetPage_SD0095P = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(13))
        Set oRegexMatch = oRegexMatches.Item(0)
        m_sAccountNumber = oRegexMatch.SubMatches(0)
    
        ' Masks the account number if we are to mask it
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(13) = Replace(m_saCurrentPageLine(13), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+NOTICE DATE"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date label on SD0095P page."
        GetPage_SD0095P = False
        Exit Function
    End If
    
    ' The date pattern varies from file to file...
    Select Case m_sCurrentFileName
        Case "SD0095P2"
            oRegexTest.Pattern = "^\s+(\d{1,2}/\d{2}/\d{2})"
        Case "SD0095P3", "SD0095P4"
            oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{2}\-\d{2})"
        Case Else
            Err.Raise g_lcErrBase, App.Title, "Setup error - file '" & m_sCurrentFileName & "' not configured for GetPage_SD0095P date extraction."
            GetPage_SD0095P = False
            Exit Function
    End Select
    
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on SD0095P page."
        GetPage_SD0095P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_SD0095P = True
    
    Exit Function
    
GetPage_SD0095P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_SD0095P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_SD0095P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_SD0095P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA95602(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA95602_Title = "CAdmin.GetPage_TA95602(...)"
    
    On Error GoTo GetPage_TA95602_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 2
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Address for this file is at top left corner.  Get delivery address has an issue with this location,
    ' Insert a blank line to get it to work properly
    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrapeAddress, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Removes the blank line at the beginning used for helping out the address scrape
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    
    ' Check for the account number
    m_oRegExp.Pattern = "^ {29}ACCOUNT NUMBER:\s+(\d+)"
    If m_oRegExp.Test(m_saCurrentPageLine(8)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA95602_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(8))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "^ {29}(\d{2}/\d{2}/\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(10)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA95602_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(10))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA95602 = True
    
    Exit Function
    
GetPage_TA95602_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA95602"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA95602_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA95602_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA956023(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA956023_Title = "CAdmin.GetPage_TA956023(...)"
    
    On Error GoTo GetPage_TA956023_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 5
    sAcctNumRegex = "on account (\d+)"
    lAcctNumLine = 11
    
    Select Case m_sCurrentFileName
        Case "TA95602P3", "E_TA956023"
            lAddrX = 1
            lAddrLine = 5
            sAcctNumRegex = "Your account (\d+)"
            lAcctNumLine = 11
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956023_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}\-\d{2}\-\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956023_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA956023 = True
    
    Exit Function
    
GetPage_TA956023_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA956023"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA956023_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA956023_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA956022(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA956022_Title = "CAdmin.GetPage_TA956022(...)"
    
    On Error GoTo GetPage_TA956022_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 5
    sAcctNumRegex = " Acct (\d+)"
    lAcctNumLine = 10
    
    Select Case m_sCurrentFileName
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956022_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}\-\d{2}\-\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956022_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA956022 = True
    
    Exit Function
    
GetPage_TA956022_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA956022"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA956022_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA956022_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA956024(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA956024_Title = "CAdmin.GetPage_TA956024(...)"
    
    On Error GoTo GetPage_TA956024_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
        
    Select Case m_sCurrentFileName
        Case Else
            lAddrX = 1
            lAddrLine = 5
            sAcctNumRegex = "Acct (\d+)"
            lAcctNumLine = 10
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956024_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}\-\d{2}\-\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956024_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA956024 = True
    
    Exit Function
    
GetPage_TA956024_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA956024"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA956024_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA956024_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA955991(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA955991_Title = "CAdmin.GetPage_TA955991(...)"
    
    On Error GoTo GetPage_TA955991_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 5
    sAcctNumRegex = "Acct (\d+)"
    lAcctNumLine = 10
    
    Select Case m_sCurrentFileName
        ' Add cases per file name in here if file differs
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA955991_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}[\-/]\d{2}[\-/]\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA955991_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA955991 = True
    
    Exit Function
    
GetPage_TA955991_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA955991"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA955991_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA955991_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_FT0161(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_FT0161_Title = "CAdmin.GetPage_FT0161(...)"
    
    On Error GoTo GetPage_FT0161_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 8
    lAddrLine = 14
    sAcctNumRegex = "ACCT NUMBER"
    lAcctNumLine = 8
    
    Select Case m_sCurrentFileName
        ' Add cases for files that don't match the above
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 8, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 8, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 8, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_FT0161_Title, "Cannot find account number on page."
    End If
    
    m_oRegExp.Pattern = " (\d+) "
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine + 1)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_FT0161_Title, "Cannot find account number value on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine + 1))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "(\d{2}[\-/]\d{2}[\-/]\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_FT0161_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_FT0161 = True
    
    Exit Function
    
GetPage_FT0161_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_FT0161"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_FT0161_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_FT0161_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA955992(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA955992_Title = "CAdmin.GetPage_TA955992(...)"
    
    On Error GoTo GetPage_TA955992_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 5
    sAcctNumRegex = "Acct (\d+)$"
    lAcctNumLine = 10
    
    Select Case m_sCurrentFileName
        ' Add cases per file name here if this differs
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA955992_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}\-\d{2}\-\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA955992_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA955992 = True
    
    Exit Function
    
GetPage_TA955992_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA955992"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA955992_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA955992_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM956042(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TM956042_Title = "CAdmin.GetPage_TM956042(...)"
    
    On Error GoTo GetPage_TM956042_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 5
    sAcctNumRegex = "Account:\s+(\d+)"
    lAcctNumLine = 10
    
    Select Case m_sCurrentFileName
        Case "TM95604P1", "E_TM956041"
            sAcctNumRegex = "Acct\s+(\d+)"
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TM956042_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "Date:\s+(\d{2}/\d{2}/\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TM956042_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TM956042 = True
    
    Exit Function
    
GetPage_TM956042_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM956042"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM956042_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM956042_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


Private Function GetPage_CH931613(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_CH931613_Title = "CAdmin.GetPage_CH931613(...)"
    
    On Error GoTo GetPage_CH931613_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sAcctNumRegex As String
    Dim lAcctNumLine As Long
    Dim sDateRegex As String
    Dim lDateLine As Long
    
    ' Defaults
    lAddrX = 1
    lAddrLine = 2
    sAcctNumRegex = "Account:\s+(\d+)"
    lAcctNumLine = 9
    sDateRegex = "Date:\s+(\d{2}\-\d{2}\-\d{4})"
    lDateLine = 8
    
    Select Case m_sCurrentFileName
        ' Add cases here if the file differs from what's above
    End Select

    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Gets the return address
    Set m_oReturnAddress = New CAddress
'    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40, PvRTrimOutput)
'    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40, PvRTrimOutput)
'    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40, PvRTrimOutput)

    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 45, bScrapeAddress, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
        
    ' Check for the account number
    m_oRegExp.Pattern = sAcctNumRegex
    m_oRegExp.IgnoreCase = True
    If m_oRegExp.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_CH931613_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lAcctNumLine))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = sDateRegex
    If m_oRegExp.Test(m_saCurrentPageLine(lDateLine)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_CH931613_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(lDateLine))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_CH931613 = True
    
    Exit Function
    
GetPage_CH931613_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_CH931613"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_CH931613_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_CH931613_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA956024P(Optional bScrapeAddress As Boolean) As Boolean
    Const GetPage_TA956024P_Title = "CAdmin.GetPage_TA956024P(...)"
    
    On Error GoTo GetPage_TA956024P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 2
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Address for this file is at top left corner.  Get delivery address has an issue with this location,
    ' Insert a blank line to get it to work properly
    InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    
    ' Scrape the delivery address
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrapeAddress, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Removes the blank line at the beginning used for helping out the address scrape
    DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    
    ' Check for the account number
    m_oRegExp.Pattern = "^ {29}ACCOUNT NUMBER:\s*(\d+)?"
    If m_oRegExp.Test(m_saCurrentPageLine(9)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956024P_Title, "Cannot find account number on page."
    End If
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(9))
    Set oMatch = oMatches.Item(0)
    
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Check for the document date
    m_oRegExp.Pattern = "^ {29}(\d{2}-\d{2}-\d{4})"
    If m_oRegExp.Test(m_saCurrentPageLine(13)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, GetPage_TA956024P_Title, "Cannot find document date on page."
    End If
    
    Set oMatches = m_oRegExp.Execute(m_saCurrentPageLine(13))
    Set oMatch = oMatches.Item(0)
    
    Dim sRawDate As String
    Dim dtDate As Date
    
    sRawDate = oMatch.SubMatches(0)
    dtDate = CDate(sRawDate)
    
    m_sDocumentDate = Format(dtDate, "mm-dd-yyyy")
    
    GetPage_TA956024P = True
    
    Exit Function
    
GetPage_TA956024P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA956024P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA956024P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA956024P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function



Private Function GetPage_Notice(lAddrX As Long, lAddrLine As Long, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean, Optional ByVal bNoTopAddressMargin As Boolean) As Boolean
    
    On Error GoTo GetPage_Notice_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, False, bNoRightAddressMargin, bNoBottomAddressMargin, bNoTopAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    GetPage_Notice = True
    
    Exit Function
    
GetPage_Notice_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_Notice"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_Notice_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_Notice_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM0847P1_Investors() As Boolean
    
    On Error GoTo GetPage_TM0847P1_Investors_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 4
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    GetPage_TM0847P1_Investors = True
    
    Exit Function
    
GetPage_TM0847P1_Investors_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0847P1_Investors"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0847P1_Investors_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0847P1_Investors_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0420P(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    On Error GoTo GetPage_TA0420P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 9
    lAddrLine = 13
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Suppresses the bank address
'    Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
'    Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
'    Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40)
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim lWorkPage As Long
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex tester
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
'    ' Gets the page number and verifies that we are on the first page
'    oRegexTest.Pattern = "^\s+PAGE\s+(\d+)"
'    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find page number on TA0420P page."
'        GetPage_TA0420P = False
'        Exit Function
'    End If
'
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    lWorkPage = CLng(oRegexMatch.SubMatches.Item(0))
'
'    If lWorkPage > 1 Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - piece cannot begin with non-page 1 page on TA0420P page."
'        GetPage_TA0420P = False
'        Exit Function
'    End If
    
    ' Gets the account number
    oRegexTest.Pattern = "^\s+ACCOUNT$"
    If oRegexTest.Test(m_saCurrentPageLine(12)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number label on TA0420P page."
        GetPage_TA0420P = False
        Exit Function
    End If
    oRegexTest.Pattern = "^\s+NUMBER$"
    If oRegexTest.Test(m_saCurrentPageLine(13)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number label on TA0420P page."
        GetPage_TA0420P = False
        Exit Function
    End If
    oRegexTest.Pattern = "^\s+([\d ]+)$"
    If oRegexTest.Test(m_saCurrentPageLine(14)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TA0420P page."
        GetPage_TA0420P = False
        Exit Function
    End If

    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(14))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)

    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(14) = Replace(m_saCurrentPageLine(14), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Removes spaces out of the account number - we have masked it at this point in the page array
    ' For some reason this document prints with spaces in the account number that are not significant
    m_sAccountNumber = Replace(m_sAccountNumber, " ", "")
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+DATE"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date label on TA0420P page."
        GetPage_TA0420P = False
        Exit Function
    End If
    
    oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TA0420P page."
        GetPage_TA0420P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA0420P = True
    
    Exit Function
    
GetPage_TA0420P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0420P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0420P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0420P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_FT0160P(Optional ByVal bMaskAcctNum As Boolean, Optional ByVal bSuppressRetAddr As Boolean) As Boolean
    
    On Error GoTo GetPage_FT0160P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim bDummyPage As Boolean
    Dim bAcctNumberFound As Boolean
    
    lAddrX = 9
    lAddrLine = 14
    bDummyPage = False
    bAcctNumberFound = False
    
    Select Case m_sCurrentFileName
        Case "FT0160P1"
            ' Do nothing
        Case Else
            If m_sXmlInst = "Investors" Then
'                lAddrLine = 15 'RM12007 fix for Signature 19.1 release
            End If
    End Select
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Check if this page is a dummy page
    If Trim(m_oDeliveryAddress.AddressLine(1)) = "NO NAME   NO CIF NAME & ADDRESS FOUND" Then
        bDummyPage = True
    End If
    
    ' Suppresses the bank address
    If bSuppressRetAddr Then
        Mid(m_saCurrentPageLine(0), 8, 40) = Space(40)
        Mid(m_saCurrentPageLine(1), 8, 40) = Space(40)
        Mid(m_saCurrentPageLine(2), 8, 40) = Space(40)
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim lWorkPage As Long
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex tester
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
'    ' Gets the page number and verifies that we are on the first page
'    oRegexTest.Pattern = "^\s+PAGE\s+(\d+)"
'    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find page number on FT0160P page."
'        GetPage_FT0160P = False
'        Exit Function
'    End If
'
'    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
'    Set oRegexMatch = oRegexMatches.Item(0)
'    lWorkPage = CLng(oRegexMatch.SubMatches.Item(0))
'
'    If lWorkPage > 1 Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - piece cannot begin with non-page 1 page on FT0160P page."
'        GetPage_FT0160P = False
'        Exit Function
'    End If

    ' Sets the location of where to scrape values depending on which file type is called
    Dim lAcctNumLine As Long
    Dim lDateLine As Long
    Select Case m_sCurrentFileName
        Case "FT0160P1"
            lAcctNumLine = 9
            lDateLine = 0
        Case Else
            lAcctNumLine = 10
            lDateLine = 0
            If m_sXmlInst = "Investors" Then
'                lAcctNumLine = 11 'RM12007 fix for Signature 19.1 release
            End If
    End Select
    
    ' Gets the account number
    oRegexTest.Pattern = "^ ACCT NUMBER"
    If oRegexTest.Test(m_saCurrentPageLine(lAcctNumLine - 1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number label on FT0160P page."
        GetPage_FT0160P = False
        Exit Function
    End If
    oRegexTest.Pattern = "^\s+(\d+) "
    If oRegexTest.Test(m_saCurrentPageLine(lAcctNumLine)) = False Then
        ' If we are not a dummy page then go ahead and raise the error
        If bDummyPage = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on FT0160P page."
            GetPage_FT0160P = False
            Exit Function
        End If
    Else
        bAcctNumberFound = True
    End If
    
    If bAcctNumberFound Then
    
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(lAcctNumLine))
        Set oRegexMatch = oRegexMatches.Item(0)
        m_sAccountNumber = oRegexMatch.SubMatches(0)
    
        ' Masks the account number if we are to mask it
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(lAcctNumLine) = Replace(m_saCurrentPageLine(lAcctNumLine), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
    
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{2}\-\d{2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(lDateLine)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on FT0160P page."
        GetPage_FT0160P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(lDateLine))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_FT0160P = True
    
    Exit Function
    
GetPage_FT0160P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_FT0160P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_FT0160P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_FT0160P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


Private Function GetPage_CH0202P() As Boolean
    
    On Error GoTo GetPage_CH0202P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 5
    lAddrLine = 17
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Suppresses the bank address
    Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
    Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
    Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim lWorkPage As Long
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex tester
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
    ' Gets the page number and verifies that we are on the first page
    oRegexTest.Pattern = "^\s+PAGE\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find page number on CH0202P page."
        GetPage_CH0202P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    lWorkPage = CLng(oRegexMatch.SubMatches.Item(0))
    
    If lWorkPage > 1 Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - piece cannot begin with non-page 1 page on CH0202P page."
        GetPage_CH0202P = False
        Exit Function
    End If
    
    ' Gets the account number
    oRegexTest.Pattern = "\s+DEMAND DEPOSIT ACCOUNT\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(16)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on CH0202P page."
        GetPage_CH0202P = False
        Exit Function
    End If

    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(16))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)

'    ' Masks the account number if we are to mask it
'    If bMaskAcctNum Then
'        If Len(m_sAccountNumber) > 4 Then
'            sMaskedAcctNum = m_sAccountNumber
'            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
'            m_saCurrentPageLine(6) = Replace(m_saCurrentPageLine(16), m_sAccountNumber, sMaskedAcctNum)
'        End If
'    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on CH0202P page."
        GetPage_CH0202P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_CH0202P = True
    
    Exit Function
    
GetPage_CH0202P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_CH0202P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_CH0202P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_CH0202P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


Private Function GetPage_TA1161P1(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    On Error GoTo GetPage_TA1161P1_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim bNoTopAddressVerify As Boolean
    
    Select Case m_sXmlInst
        Case "Nationwide"
            lAddrX = 11
            lAddrLine = m_lCurrentPageIdx - 4
            bNoTopAddressVerify = True
        Case "Investors"
            lAddrX = 11
            lAddrLine = 16
            bNoTopAddressVerify = True
        Case Else
            Err.Raise g_lcErrBase + g_lcErrParameterError, "CAdmin.GetPage_TA1161P1", "Inst " & m_sXmlInst & " is not configured."
    End Select
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, bNoTopAddressVerify)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Suppresses the bank address
'    Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
'    Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
'    Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 1, 40)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 1, 40)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 1, 40)
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+ACCOUNT NUMBER:\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(6)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TA1161P1 page."
        GetPage_TA1161P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(6))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(6) = Replace(m_saCurrentPageLine(6), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+DATE:\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(5)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TA1161P1 page."
        GetPage_TA1161P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(5))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA1161P1 = True
    
    Exit Function
    
GetPage_TA1161P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA1161P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA1161P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA1161P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0185P1(Optional ByVal bMaskAcctNum As Boolean, Optional ByVal bSuppressReturnAddr As Boolean) As Boolean
    
    On Error GoTo GetPage_TA0185P1_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 6
    lAddrLine = 14
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Suppresses the bank return address
    If bSuppressReturnAddr Then
        Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
        Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
        Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    End If
    
    ' Gets the account numbers on the piece so we can suppress them
        ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim dtDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    'oRegexTest.Pattern = "^\s+TRANSACTION ACCOUNT.+(?:HELOC|PERSONAL OVERDRAFT LOC|BL BUSINESS OVERDRAFT)$"
    oRegexTest.Pattern = "^\s+TRANSACTION ACCOUNT\b"
    If oRegexTest.Test(m_saCurrentPageLine(9)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number labels on TA0185P1 page."
        GetPage_TA0185P1 = False
        Exit Function
    End If
    
    oRegexTest.Pattern = "\s+(\d+-\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(10)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number 1 on TA0185P1 page."
        GetPage_TA0185P1 = False
        Exit Function
    End If
   
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(10))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    Dim sUnmaskedAcctNum As String
    sUnmaskedAcctNum = m_sAccountNumber 'Preserve original format (with dash) for masking code.
    m_sAccountNumber = Replace(m_sAccountNumber, "-", "")
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(sUnmaskedAcctNum) > 4 Then
            sMaskedAcctNum = sUnmaskedAcctNum
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(10) = Replace(m_saCurrentPageLine(10), sUnmaskedAcctNum, sMaskedAcctNum)
        End If
    End If
    
    If bMaskAcctNum Then
        'Mask the transfer account number.
        oRegexTest.Pattern = ".{49}\s+(\d+)\s*$"
        If oRegexTest.Test(m_saCurrentPageLine(10)) = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find the transfer account number on TA0185P1 page."
            GetPage_TA0185P1 = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(10))
        Set oRegexMatch = oRegexMatches.Item(0)
        sUnmaskedAcctNum = oRegexMatch.SubMatches(0)
        sMaskedAcctNum = sUnmaskedAcctNum
        Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
        m_saCurrentPageLine(10) = Replace(m_saCurrentPageLine(10), sUnmaskedAcctNum, sMaskedAcctNum)
    End If
    
'    oRegexTest.Pattern = ".{50}\s+(\d+)"
'    If oRegexTest.Test(m_saCurrentPageLine(10)) = False Then
'        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number 2 on TA0185P1 page."
'        GetPage_TA0185P1 = False
'        Exit Function
'    End If
'
''    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(10))
''    Set oRegexMatch = oRegexMatches.Item(0)
'    'm_sAccountNumber = oRegexMatch.SubMatches(0)
'
'    ' Masks the account number if we are to mask it
'    If bMaskAcctNum Then
'        If Len(m_sAccountNumber) > 4 Then
'            sMaskedAcctNum = oRegexMatch.SubMatches(0)
'            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
'            m_saCurrentPageLine(10) = Replace(m_saCurrentPageLine(10), oRegexMatch.SubMatches(0), sMaskedAcctNum)
'        End If
'    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "\b(\d{1,2}\-\d{1,2}\-\d{4})\s*$"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TM0365P1 page."
        GetPage_TA0185P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    dtDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(dtDocDate, "MM-dd-yyyy")
    
    GetPage_TA0185P1 = True
    
    Exit Function
    
GetPage_TA0185P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0185P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0185P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0185P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM0365P1(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    On Error GoTo GetPage_TM0365P1_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sMaskedAcctNum As String
        
    Select Case m_sXmlInst
        Case "Nationwide"
            lAddrX = 9
            lAddrLine = 13
'        'RM14945
'        Case "SNB"
'            lAddrX = 3
'            lAddrLine = 9
'        'END RM14945
        Case Else
            lAddrX = 9
            lAddrLine = 12
    End Select
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' We scrape the return address for this notice because it is in the same corner as the corner card resource,
    ' which SHOULD have the return address in it
    Mid(m_saCurrentPageLine(0), 7, 40) = Space(40)
    Mid(m_saCurrentPageLine(1), 7, 40) = Space(40)
    Mid(m_saCurrentPageLine(2), 7, 40) = Space(40)
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Inits the regex object
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    
    ' Gets the certificate number off the page and verifies that it can be found
    Dim bCertificateFound As Boolean
    Dim lCertificateIdx As Long
    oRegexTest.Pattern = "^\s+CERTIFICATE:\s+(\d+)"
    If bCertificateFound = False And oRegexTest.Test(m_saCurrentPageLine(11)) = True Then
        bCertificateFound = True
        lCertificateIdx = 11
    End If
    If bCertificateFound = False And oRegexTest.Test(m_saCurrentPageLine(12)) = True Then
        bCertificateFound = True
        lCertificateIdx = 12
    End If
    
'    'RM14945
'    If m_sXmlInst = "SNB" Then
'        If bCertificateFound = False And oRegexTest.Test(m_saCurrentPageLine(31)) = True Then
'            bCertificateFound = True
'            lCertificateIdx = 31
'        End If
'    End If
'    'END RM14945
    
    If bCertificateFound = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TM0365P1 page."
        GetPage_TM0365P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(lCertificateIdx))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(11) = Replace(m_saCurrentPageLine(11), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TM0365P1 page."
        GetPage_TM0365P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TM0365P1 = True
    
    Exit Function
    
GetPage_TM0365P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0365P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0365P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0365P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0434P1(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    On Error GoTo GetPage_TA0434P1_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim sMaskedAcctNum As String
    
    lAddrX = 6
    lAddrLine = 3
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+Account Number:\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(10)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TA0434P1 page."
        GetPage_TA0434P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(10))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(10) = Replace(m_saCurrentPageLine(10), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TA0050P1 page."
        GetPage_TA0434P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA0434P1 = True
    
    Exit Function
    
GetPage_TA0434P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0434P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0434P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0434P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA914863P() As Boolean
    
    On Error GoTo GetPage_TA914863P_Error
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 1
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    GetPage_TA914863P = True
    
    Exit Function
    
GetPage_TA914863P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA914863P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA914863P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA914863P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0180P(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    Const sc183Notice = "***NOTICE"
    
    On Error GoTo GetPage_TA0180P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 1
    lAddrLine = 13
    
    ' Fix for Signature release 17.2 that adds a blank line to the top of the TA0180P1 file type
    If Len(Trim(Mid(m_saCurrentPageLine(0), 1, 94))) = 0 And _
        (InStr(1, m_saCurrentPageLine(1), sc183Notice) > 0 Or (Trim(m_saCurrentPageLine(1)) = "NOTICE OF INSUFFICIENT FUNDS" Or Trim(m_saCurrentPageLine(1)) = "NOTICE OF UNCOLLECTED FUNDS")) Then
        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 1
    End If
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, True, False, True, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Variables used for validating
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Sanity check to make sure this notice is correct
    
    If Trim(m_saCurrentPageLine(0)) = "NOTICE OF INSUFFICIENT FUNDS" Or Trim(m_saCurrentPageLine(0)) = "NOTICE OF UNCOLLECTED FUNDS" Then
        Set oRegexTest = New RegExp
        oRegexTest.IgnoreCase = True
        oRegexTest.Pattern = "^ACCOUNT:\s+(\d+)$"
        If oRegexTest.Test(m_saCurrentPageLine(3)) = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - account number not found for TA0180P."
            GetPage_TA0180P = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(3))
        Set oRegexMatch = oRegexMatches.Item(0)
        m_sAccountNumber = oRegexMatch.SubMatches.Item(0)
        
        ' Masks the account number if we are to mask it
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(3) = Replace(m_saCurrentPageLine(3), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
        
        oRegexTest.Pattern = "^DATE:\s+(\d{2}/\d{2}/\d{4})$"
        If oRegexTest.Test(m_saCurrentPageLine(2)) = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - document date could not be found for TA0180P."
            GetPage_TA0180P = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(2))
        Set oRegexMatch = oRegexMatches.Item(0)
        
        ' Formats the date in a consistent way
        sRawDate = oRegexMatch.SubMatches.Item(0)
        oDocDate = CDate(sRawDate)
        m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
        
    ElseIf Left(Trim(m_saCurrentPageLine(0)), Len(sc183Notice)) = sc183Notice Then
        Set oRegexTest = New RegExp
        oRegexTest.IgnoreCase = True
        oRegexTest.Pattern = "^\s+ACCT NBR:\s+(\d+)$"
        If oRegexTest.Test(m_saCurrentPageLine(2)) = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - account number not found for TA0180P."
            GetPage_TA0180P = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(2))
        Set oRegexMatch = oRegexMatches.Item(0)
        m_sAccountNumber = oRegexMatch.SubMatches.Item(0)
        
        ' Masks the account number if we are to mask it
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(2) = Replace(m_saCurrentPageLine(2), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
        
        oRegexTest.Pattern = "^\s+DATE:\s+(\d{2}/\d{2}/\d{4})$"
        If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - document date could not be found for TA0180P."
            GetPage_TA0180P = False
            Exit Function
        End If
        Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
        Set oRegexMatch = oRegexMatches.Item(0)
        
        ' Formats the date in a consistent way
        sRawDate = oRegexMatch.SubMatches.Item(0)
        oDocDate = CDate(sRawDate)
        m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    Else
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - incorrect title found on line 0."
        GetPage_TA0180P = False
        Exit Function
    End If
        
    GetPage_TA0180P = True
    
    Exit Function
    
GetPage_TA0180P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0180P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0180P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0180P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function


Private Function GetPage_TA0190P(Optional ByVal bMaskAcctNum As Boolean, Optional ByVal bSuppressBankAddr As Boolean) As Boolean
    
    On Error GoTo GetPage_TA0190P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    
    lAddrX = 6
    lAddrLine = 14
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, False, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' We scrape the return address for this notice because it is in the same corner as the corner card resource,
    ' which SHOULD have the return address in it
    If bSuppressBankAddr Then
        Mid(m_saCurrentPageLine(0), 1, 40) = Space(40)
        Mid(m_saCurrentPageLine(1), 1, 40) = Space(40)
        Mid(m_saCurrentPageLine(2), 1, 40) = Space(40)
    End If
    
    ' Gets the TA0190P date
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    Set oRegexTest = New RegExp
    oRegexTest.Pattern = "^.{35}\s+DATE"
    oRegexTest.IgnoreCase = True
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - document date label not found."
        GetPage_TA0190P = False
        Exit Function
    End If
    
    oRegexTest.Pattern = "^.{35}\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(1)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - document date value not found."
        GetPage_TA0190P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(1))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    ' Gets the account number
    Dim sAccountNumber As String
    If Trim(Left(m_saCurrentPageLine(6), 15)) <> "ACCOUNT" Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - account number label not found."
        GetPage_TA0190P = False
        Exit Function
    End If
    
    If Trim(Left(m_saCurrentPageLine(7), 15)) <> "NUMBER" Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - account number label not found."
        GetPage_TA0190P = False
        Exit Function
    End If
    
    sAccountNumber = Trim(Left(m_saCurrentPageLine(9), 15))
    If m_sXmlInst = "SNB" Then
        'May want to do this for all clients.
        sAccountNumber = Replace(sAccountNumber, " ", "")
    End If
    If Len(sAccountNumber) > 0 Then
        m_sAccountNumber = sAccountNumber
        
        ' Masks the account number if we are to mask it
        If bMaskAcctNum Then
            If Len(m_sAccountNumber) > 4 Then
                sMaskedAcctNum = m_sAccountNumber
                Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
                m_saCurrentPageLine(9) = Replace(m_saCurrentPageLine(9), m_sAccountNumber, sMaskedAcctNum)
            End If
        End If
    Else
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid TA0190P page encountered - account number value not found."
        GetPage_TA0190P = False
        Exit Function
    End If
    
    GetPage_TA0190P = True
    
    Exit Function
    
GetPage_TA0190P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0190P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0190P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0190P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA990771P() As Boolean
    
    On Error GoTo GetPage_TA990771P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < 20 Then
        m_lCurrentPageIdx = 20
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, 8, 16, 40, True, False, False, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    m_sAccountNumber = Trim(m_saCurrentPageLine(1))
    
    GetPage_TA990771P = True
    
    Exit Function
    
GetPage_TA990771P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA990771P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA990771P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA990771P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

' GetPage_NoticeScrape - GetPage_Notice with address scrape parameter added
Private Function GetPage_NoticeScrape(lAddrX As Long, lAddrLine As Long, bScrape As Boolean, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean, Optional ByVal bNoTopAddressMargin As Boolean) As Boolean
    
    On Error GoTo GetPage_NoticeScrape_Error
    
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrape, bNoRightAddressMargin, bNoBottomAddressMargin, bNoTopAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
        
    
    GetPage_NoticeScrape = True
    
    Exit Function
    
GetPage_NoticeScrape_Error:
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_NoticeScrape"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_NoticeScrape_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_NoticeScrape_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function
' GetPage_NoticeScrapeWoori - GetPage_NoticeWoori with address scrape parameter added
Private Function GetPage_NoticeScrapeWoori(lAddrX As Long, lAddrLine As Long, lRetAddrX As Long, lRetAddrLine As Long, lRetLineCount As Long, lRetAddressLen As Long, bScrape As Boolean, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean, Optional ByVal bNoTopAddressMargin As Boolean) As Boolean
    
    On Error GoTo GetPage_NoticeScrapeWoori_Error
    
        
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrape, bNoRightAddressMargin, bNoBottomAddressMargin, bNoTopAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Set m_oReturnAddress = GetReturnAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lRetAddrLine, lRetLineCount, lRetAddrX, lRetAddressLen, False, False)
    
    
    GetPage_NoticeScrapeWoori = True
    
    Exit Function
    
GetPage_NoticeScrapeWoori_Error:
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_NoticeScrapeWoori"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_NoticeScrapeWoori_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_NoticeScrapeWoori_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function

Private Function GetPage_LN0460P(lAddrX As Long, lAddrLine As Long, bScrape As Boolean, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean, Optional ByVal bNoTopAddressMargin As Boolean) As Boolean
    
    On Error GoTo GetPage_LN0460P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrape, bNoRightAddressMargin, bNoBottomAddressMargin, bNoTopAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Dim oRegex As RegExp
    Dim oMatch As Match
    Set oRegex = New RegExp
    
    ' Get the account number
    oRegex.Pattern = "NOTE NUMBER\s+(\d+)"
    If oRegex.Test(m_saCurrentPageLine(2)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate NOTE NUMBER on LN0460P page"
    End If
    Set oMatch = oRegex.Execute(m_saCurrentPageLine(2)).Item(0)
    m_sAccountNumber = oMatch.SubMatches(0)
    
    ' Get the document date
    oRegex.Pattern = "(\d{2}\-\d{2}\-\d{2})"
    If oRegex.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate notice date on LN0460P page"
    End If
    Set oMatch = oRegex.Execute(m_saCurrentPageLine(0)).Item(0)
    m_sDocumentDate = oMatch.SubMatches(0)
    
    GetPage_LN0460P = True
    
    Exit Function
    
GetPage_LN0460P_Error:
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0460P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0460P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0460P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function

Private Function GetPage_TM0360P() As Boolean
    
    On Error GoTo GetPage_TM0360P_Error
    
    InitializePieceFields
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim bScrape As Boolean
    Dim oRegexTest As RegExp
    
    Dim sCertificateRegex As String
    Dim lCertificateLine As Long
    
    lAddrX = 5
    lAddrLine = 14
    bScrape = True
    
    Select Case m_sCurrentFileName
        Case "TM0360P2", "E_TM0360P2"
            lAddrLine = 13
            lCertificateLine = 12
            sCertificateRegex = "^\s+CERTIFICATE:\s+(\d+)"
        Case "TM0360P7", "E_TM0360P7"
            lCertificateLine = 13
            sCertificateRegex = "^\s+CERTIFICATE\s+(\d+)"
        Case "TM0360P8", "E_TM0360P8"
            lAddrLine = 15
            lCertificateLine = 14
            sCertificateRegex = "^\s+CERTIFICATE\s+(\d+)"
        Case "TM0360P4", "E_TM0360P4"
            lAddrLine = 13
            lCertificateLine = 6
            sCertificateRegex = "^\s+C\.D\. NUMBER"
        Case Else
            Err.Raise g_lcErrParameterError, "CAdmin.GetPage_TM0360P()", "Report type " & m_sCurrentFileName & " is not configured."
    End Select
    

    Set oRegexTest = New RegExp
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 5, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 5, 40, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 5, 40, PvRTrimOutput)
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrape, False, False, False)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim sRawDate As String
    
    ' Gets the certificate number off the page and verifies that it can be found
    oRegexTest.Pattern = sCertificateRegex
    If oRegexTest.Test(m_saCurrentPageLine(lCertificateLine)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TM0360P page."
        GetPage_TM0360P = False
        Exit Function
    End If
    
    ' TM0360P4 has the certificate number on a separate line than the certificate label
    Select Case m_sCurrentFileName
        Case "TM0360P4", "E_TM0360P4"
            sCertificateRegex = "^\s+(\d+)"
            oRegexTest.Pattern = sCertificateRegex
            lCertificateLine = lCertificateLine + 1
    End Select
    
    Set oMatches = oRegexTest.Execute(m_saCurrentPageLine(lCertificateLine))
    Set oMatch = oMatches.Item(0)
    m_sAccountNumber = oMatch.SubMatches(0)
        
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}\-\d{1,2}\-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TM0360P page."
        GetPage_TM0360P = False
        Exit Function
    End If
    
    Set oMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oMatch = oMatches.Item(0)
    m_sDocumentDate = Format(CDate(oMatch.SubMatches(0)), "mm-dd-yyyy")
    
    GetPage_TM0360P = True
    
    Exit Function
    
GetPage_TM0360P_Error:
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0360P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0360P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0360P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function

' GetPage_TM0330P1
Private Function GetPage_TM0330P1(lAddrX As Long, lAddrLine As Long, bScrape As Boolean, Optional ByVal bNoRightAddressMargin As Boolean, Optional ByVal bNoBottomAddressMargin As Boolean, Optional ByVal bNoTopAddressMargin As Boolean, Optional ByVal bKeepReturnAddress As Boolean) As Boolean
    
    On Error GoTo GetPage_TM0330P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lAddrLine + 4 Then
        m_lCurrentPageIdx = lAddrLine + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    ' Shift text right to get it out of the envelope window
    m_saCurrentPageLine(4) = Right$(Space(80) & Trim(m_saCurrentPageLine(4)), 80)
    m_saCurrentPageLine(5) = Right$(Space(80) & Trim(m_saCurrentPageLine(5)), 80)
    
    ' Scrape out the return address
    Const lcAddrX = 7
    Const lcAddrLen = 40
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), lcAddrX, lcAddrLen, PvRTrimOutput)
    
    ' Prevent the return address from printing
    If bKeepReturnAddress = False Then
        Set m_oReturnAddress = Nothing
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, 40, bScrape, bNoRightAddressMargin, bNoBottomAddressMargin, bNoTopAddressMargin)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    If m_bXML_Archive Then
        Dim oRegex As New RegExp
        Dim oMatches As MatchCollection
    
        m_sAccountNumber = vbNullString
'        If Len(m_sAccountNumber) = 0 Then
'            Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
'        End If
        
'        ' Gets the document date
'        oDocDate = FileDateTime(m_sInputFile(m_lInputFilePtr))
'        m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
        ' Gets the document date
        If m_lCurrentPageIdx >= 9 Then
            oRegex.Pattern = "\b(\d{1,2}-\d{2}-\d{2})\s*$"
            If oRegex.Test(m_saCurrentPageLine(9)) = True Then
                Set oMatches = oRegex.Execute(m_saCurrentPageLine(9))
                m_sDocumentDate = oMatches.Item(0).SubMatches(0)
            End If
        End If
        If Len(m_sDocumentDate) = 0 Then
            Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
        End If
        
        Dim i As Long
        m_lIndexIdx = 2
    
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
    
        m_saIndexKeys(0) = m_scDocumentTypeKey
        m_saIndexValues(0) = "Retirement Account Statement - Fiserv"
    
        m_saIndexKeys(1) = m_scPrimaryNameKey
        m_saIndexValues(1) = m_oDeliveryAddress.Name
    
        m_saIndexKeys(2) = m_scStatementDateKey
        m_saIndexValues(2) = m_sDocumentDate
'
'        m_saIndexKeys(3) = m_scAccountNumberKey
'        m_saIndexValues(3) = m_sAccountNumber
    
        If m_sXmlInst = m_scInst_29175100_FinFactors Then
            m_saIndexValues(0) = "TIM Retirement Account Statement - Fiserv" 'RM10760
            
            m_saIndexValues(2) = ReformatDate(m_saIndexValues(2), "MM/dd/yyyy")
    
            m_lIndexIdx = 4
    
            ReDim Preserve m_saIndexKeys(m_lIndexIdx)
            ReDim Preserve m_saIndexValues(m_lIndexIdx)
    
            m_saIndexKeys(3) = m_scFileTypeKey
            m_saIndexValues(3) = "16"   'Finance Factors
    
            m_saIndexKeys(4) = "File_Name"
    '        m_saIndexValues(5) = Right("0000000000" & m_lCurrentFilePieces, 10) & ".pdf"
    '        m_saIndexValues(5) = "<Sequence>.pdf"
            m_saIndexValues(4) = Right("0000000000" & m_lCurrentFilePieces + 1, 10) & ".pdf"
        End If
    End If
    
    GetPage_TM0330P1 = True
    
    Exit Function
    
GetPage_TM0330P1_Error:
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0330P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0330P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0330P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function
Private Sub InsertLinesInArray(ByRef saArray() As String, ByRef saBoldArray() As String, ByRef lArrayIdx As Long, ByVal lInsertPoint As Long, ByVal lLines As Long)
    Dim lOldArrayIdx As Long
    Dim i As Long
    lOldArrayIdx = lArrayIdx
    lArrayIdx = lArrayIdx + lLines
    ReDim Preserve saArray(lArrayIdx)
    ReDim Preserve saBoldArray(lArrayIdx)
    For i = lOldArrayIdx To lInsertPoint Step -1
        saArray(i + lLines) = saArray(i)
        saBoldArray(i + lLines) = saBoldArray(i)
    Next
    For i = lInsertPoint To lInsertPoint + lLines - 1
        saArray(i) = vbNullString
        saBoldArray(i) = vbNullString
    Next
End Sub

Private Sub DeleteLinesFromArray(ByRef saArray() As String, ByRef saBoldArray() As String, ByRef lArrayIdx As Long, ByVal lDeletionPoint As Long, ByVal lLines As Long)
    Dim i As Long
    Dim lDeletedLines As Long
    For i = lDeletionPoint + lLines To lArrayIdx
        lDeletedLines = lDeletedLines + 1
        If lDeletedLines <= lLines Then
            If Len(Trim(saArray(i - lLines))) > 0 Then
                Err.Raise 7, App.Title, "Deleting a non blank line is not allowed."
            End If
        End If
        saArray(i - lLines) = saArray(i)
        saBoldArray(i - lLines) = saBoldArray(i)
    Next
    lArrayIdx = lArrayIdx - lLines
    ReDim Preserve saArray(lArrayIdx)
    ReDim Preserve saBoldArray(lArrayIdx)
End Sub

Private Sub ScrollLinesInArray(ByRef saArray() As String, ByRef saBoldArray() As String, ByVal lArrayIdx As Long, ByVal lStartIndex As Long, ByVal lLength As Long, ByVal lScrollLines As Long)
    
    If lScrollLines < 2 Then
        Err.Raise 7, App.Title, "lScrollLines argument must be 2 or greater."
    End If
    
    If lScrollLines >= lLength Then
        Err.Raise 7, App.Title, "lScrollLines argument must be less than the lLength argument."
    End If
    
    If lStartIndex + lLength - 1 > lArrayIdx Then
        Err.Raise 7, App.Title, "lStartIndex plus lLength argument must be with the bounds of the array."
    End If
    
    Dim i As Long
    For i = lStartIndex + lScrollLines To lStartIndex + lLength - 1
        If Len(Trim(saArray(i - lScrollLines))) > 0 Then
            Err.Raise 7, App.Title, "Scrolling away a non blank line is not allowed."
        End If
        saArray(i - lScrollLines) = saArray(i)
        saBoldArray(i - lScrollLines) = saBoldArray(i)
        saArray(i) = vbNullString
        saBoldArray(i) = vbNullString
    Next
    For i = lStartIndex + lLength - lScrollLines To lStartIndex + lLength - 1
        saArray(i) = vbNullString
        saBoldArray(i) = vbNullString
    Next
End Sub

'==============================================================================
'Function:              BuildMICRLinePs
'
'Synopsis:              Returns a 65 character MICR line.
'
'==============================================================================
Private Function BuildMICRLinePs(ByVal sAuxOnUs As String, ByVal sTransitNumber As String, ByVal sAccountNumber As String, ByVal sTranCode As String, ByVal cAmount As Currency) As String
    Dim sTransit As String
    
    Dim sMICRAuxOnUs As String
    Dim sMICRTr As String
    Dim sMICROnUs As String
    Dim sMICRAmount As String
    Dim sChar As String
    Dim sMicrLine As String
    Dim i As Long
    
    Const lcDim66 As Long = 66
    
    Const scTransitSymbol As String = ":"
    Const scOnUsSymbol As String = "<"
    Const scDashSymbol As String = "="
    Const scAmountSymbol As String = ";"
    Const scOneDotSpace As String = "[" 'Used to insert half a space every other character.
    
    'Strip non numeric chars
    sAuxOnUs = Trim(sAuxOnUs)
    For i = 1 To Len(sAuxOnUs)
        sChar = Mid(sAuxOnUs, i, 1)
        Select Case Asc(sChar)
            Case 48 To 57
                sMICRAuxOnUs = sMICRAuxOnUs & sChar
        End Select
    Next
    If Len(sMICRAuxOnUs) > 0 Then
        sMICRAuxOnUs = scOnUsSymbol & sMICRAuxOnUs & scOnUsSymbol
    End If
    If Len(sMICRAuxOnUs) > 21 Then
       Err.Raise 5, App.Title, "Invalid procedure call in BuildMICRLinePs.  Aux OnUs must be twenty-one or less characters."
    End If
    sMICRAuxOnUs = Right(Space(21) & sMICRAuxOnUs, 21)
    
    Select Case Len(sTransitNumber)
        Case Is <= 8
            sTransitNumber = Right(String(8, "0") & sTransitNumber, 8)
            sMICRTr = Format(sTransitNumber, "0000=0000")
        Case 9
            sMICRTr = Format(sTransitNumber, "000000000")
        Case Else
    End Select
    sMICRTr = scTransitSymbol & sMICRTr & scTransitSymbol
    
    'Strip non numeric chars
    sAccountNumber = Trim(sAccountNumber)
    sMICROnUs = Space(1)
    For i = 1 To Len(sAccountNumber)
        sChar = Mid(sAccountNumber, i, 1)
        Select Case Asc(sChar)
            Case 48 To 57
                sMICROnUs = sMICROnUs & sChar
        End Select
    Next
    If Len(sMICROnUs) > 1 Then
        sMICROnUs = sMICROnUs & scOnUsSymbol
    End If
   
    If Len(sMICROnUs) > 14 Then
       Err.Raise 5, App.Title, "Invalid procedure call in BuildMICRLinePs.  Account number must be twelve or less characters."
    End If
    
    If Len(sTranCode) > 0 Then
        If Len(sTranCode) > 3 Then
           Err.Raise 5, App.Title, "Invalid procedure call in BuildMICRLine.  Tran Code must be between 0 and 999."
        End If
        If Not IsNumeric(sTranCode) Then
           Err.Raise 5, App.Title, "Invalid procedure call in BuildMICRLine.  Tran Code must be between 0 and 999."
        End If
    End If
    
    If Len(sTranCode) > 0 Then
        sMICROnUs = sMICROnUs & Space(2) & Right(Space(3) & sTranCode, 3)
    End If
    If Len(sMICROnUs) > 19 Then
       Err.Raise 5, App.Title, "Invalid procedure call in BuildMICRLinePs.  OnUs field must be nineteen or less characters."
    End If


'    If cAmount < 0 Then
'        cAmount = 0
'    End If
    If cAmount > 0 Then
        sMICRAmount = Format(cAmount * 100, "0000000000")
        sMICRAmount = scAmountSymbol & sMICRAmount & scAmountSymbol
        If Len(sMICRAmount) > 12 Then
            Err.Raise 5, App.Title, "Invalid procedure call.  Amount too large."
        End If
    End If
    
    sMicrLine = Space(lcDim66 - 1)
    Mid(sMicrLine, lcDim66 - 65) = sMICRAuxOnUs
    Mid(sMicrLine, lcDim66 - 43) = sMICRTr
    Mid(sMicrLine, lcDim66 - 32) = sMICROnUs
    Mid(sMicrLine, lcDim66 - 12) = sMICRAmount

    BuildMICRLinePs = sMicrLine
    
End Function

Private Function BuildNationwidePieceIndex(sCurrentPageArray() As String, lCurrentPageIdx As Long) As Boolean
    On Error GoTo BuildNationwidePieceIndex_Error

    ' Constants
    Const lcIndex As Long = 3
    Const scAccountNumber As String = "   ACCOUNT NBR"
    Const scActivityHeader As String = "* - - - - - - - - - - - - -ACTIVITY ON YOUR ACCOUNT- - - - - - - - - - - - - - *"
    Const scAccountNumberAct As String = "     ACCOUNT"
    
    'Regular expression used to pull information out of the piece
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date

    ' Set default return value
    BuildNationwidePieceIndex = False
    m_lIndexIdx = lcIndex
    
    ' Initializes the piece index arrays
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    m_saIndexKeys(0) = "File_Name"
    m_saIndexKeys(1) = "Account_Number"
    m_saIndexKeys(2) = "Document_Date"
    m_saIndexKeys(3) = "Name"
    
    ' Validate that we are Nationwide Bank
    If m_sXmlInst <> "Nationwide" Then
        Err.Raise g_lcErrParameterError, App.Title, "Inst does not match 'Nationwide' - cannot continue."
        Exit Function
    End If
    
    ' Each file and form type varies in layout - look in different locations for the required information
    If m_sXml_FileGroup = "LOANS" Then
        Select Case m_sCurrentFileName
            Case "LN395Q1", "E_LN395Q1"
                If m_sCurrentFileFormType = "HIST" Then
                    m_saIndexValues(0) = "FOS-Rate Review Notice (LN0395-HIST)"
                    
                    m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                    
                    ' Check if we have an invalid page
                    If m_lCurrentPageIdx < 20 Then
                        Err.Raise g_lcErrInvalidPage, App.Title, "Current page is abormally short.  Expecting at least 20 lines."
                        BuildNationwidePieceIndex = False
                        Exit Function
                    End If
                    
                    ' Gets the account number
                    oRegex.Pattern = " RE: Loan Number\s+(\d+)"
                    If oRegex.Test(sCurrentPageArray(18)) = True Then
                        Set oMatches = oRegex.Execute(sCurrentPageArray(18))
                        Set oMatch = oMatches.Item(0)
                        m_sAccountNumber = oMatch.SubMatches(0)
                        m_saIndexValues(1) = m_sAccountNumber
                    Else
                        ' Check if it is a payoff summary
                        If Trim(sCurrentPageArray(18)) = "*------ PAYOFF SUMMARY ------*" Then
                            oRegex.Pattern = "^\s+LOAN NUMBER:\s+(\d+)$"
                            If oRegex.Test(sCurrentPageArray(19)) = True Then
                                Set oMatches = oRegex.Execute(sCurrentPageArray(19))
                                Set oMatch = oMatches.Item(0)
                                m_sAccountNumber = oMatch.SubMatches(0)
                                m_saIndexValues(1) = m_sAccountNumber
                            Else
                                Err.Raise g_lcErrBase, App.Title, "Cannot find account number on page."
                                BuildNationwidePieceIndex = False
                                Exit Function
                            End If
                        Else
                            Err.Raise g_lcErrBase, App.Title, "Cannot find Payoff summary on page."
                            BuildNationwidePieceIndex = False
                            Exit Function
                        End If
                    End If
                    
                    ' Gets the document date
                    sRawDate = Trim(sCurrentPageArray(6))
                    oDocDate = CDate(sRawDate)
                    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
                    m_saIndexValues(2) = m_sDocumentDate
                    
                    ' Gets the first line of the address
                    m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                ElseIf m_sCurrentFileFormType = "CLET" Then
                    m_saIndexValues(0) = "FOS-Rate Change Notice (LN0395-CLET)"
                    
                    m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                    
                    ' Gets the account number
                    oRegex.Pattern = " Account Number\s+(\d+)"
                    If oRegex.Test(sCurrentPageArray(3)) = True Then
                        Set oMatches = oRegex.Execute(sCurrentPageArray(3))
                        Set oMatch = oMatches.Item(0)
                        m_sAccountNumber = oMatch.SubMatches(0)
                        m_saIndexValues(1) = m_sAccountNumber
                    Else
                        Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
                        BuildNationwidePieceIndex = False
                        Exit Function
                    End If
                    
                    ' Gets the document date
                    oRegex.Pattern = " Notice Date\s+(\d{2}/\d{2}/\d{2})"
                    If oRegex.Test(sCurrentPageArray(2)) = True Then
                        Set oMatches = oRegex.Execute(sCurrentPageArray(2))
                        Set oMatch = oMatches.Item(0)
                        sRawDate = oMatch.SubMatches(0)
                        oDocDate = CDate(sRawDate)
                        m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
                        m_saIndexValues(2) = m_sDocumentDate
                    Else
                        Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find notice date on page."
                        BuildNationwidePieceIndex = False
                        Exit Function
                    End If
                    
                    
                    ' Gets the first line of the address
                    m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                    
                Else
                    Err.Raise g_lcErrInvalidPage, App.Title, "Unrecognized form type for " & m_sCurrentFileName & ": " & m_sCurrentFileFormType
                    BuildNationwidePieceIndex = False
                    Exit Function
                End If
                
    '        Case "LN0470P1"
    '            m_saIndexValues(0) = m_sCurrentFileName
    '
    '            ' Gets the account number
    '            oRegex.Pattern = " NOTE NUMBER\s+(\d+)"
    '            If oRegex.Test(sCurrentPageArray(2)) = True Then
    '                Set oMatches = oRegex.Execute(sCurrentPageArray(2))
    '                Set oMatch = oMatches.Item(0)
    '                m_sAccountNumber = oMatch.SubMatches(0)
    '                m_saIndexValues(1) = m_sAccountNumber
    '            Else
    '                Err.Raise g_lcErrBase, g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
    '            End If
    '
    '            ' Gets the document date
    '            sRawDate = Mid(sCurrentPageArray(0), 42, 8)
    '            oDocDate = CDate(sRawDate)
    '            m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    '            m_saIndexValues(2) = m_sDocumentDate
    '
    '            ' Gets the first line of the address
    '            m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
    '        Case "LN0470P2"
    '            m_saIndexValues(0) = m_sCurrentFileName
    '
    '            ' Gets the account number
    '            oRegex.Pattern = " TYPE/NOTE NBR\s+\d+\s+(\d+)"
    '            If oRegex.Test(sCurrentPageArray(8)) = True Then
    '                Set oMatches = oRegex.Execute(sCurrentPageArray(8))
    '                Set oMatch = oMatches.Item(0)
    '                m_sAccountNumber = oMatch.SubMatches(0)
    '                m_saIndexValues(1) = m_sAccountNumber
    '            Else
    '                Err.Raise g_lcErrBase, g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
    '            End If
    '
    '            ' Gets the document date
    '            sRawDate = Mid(sCurrentPageArray(0), 42, 6)
    '            sRawDate = Mid(sRawDate, 1, 2) & "/" & Mid(sRawDate, 3, 2) & "/" & Mid(sRawDate, 5, 2)
    '            oDocDate = CDate(sRawDate)
    '            m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    '            m_saIndexValues(2) = m_sDocumentDate
    '
    '            ' Gets the first line of the address
    '            m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
            Case "LN0480Q1", "E_QSYSPRT"
                ' Client requested file name index
                m_saIndexValues(0) = "FOS-Billing Statement (LN0480)"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number - account number is the first number on line 8
                m_saIndexValues(1) = m_sAccountNumber
                    
                ' Gets the document date
                m_saIndexValues(2) = m_sDocumentDate
                
                ' Gets the first name
                m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                
            Case "T152Q2"
                ' Client-requested file name index
                m_saIndexValues(0) = "FOS-Billing Statement (TA0152)"
                
                ' Gets the account number
                oRegex.Pattern = "^\s+ Account Number:\s+(\d+)$"
                If oRegex.Test(sCurrentPageArray(8)) = True Then
                    Set oMatches = oRegex.Execute(sCurrentPageArray(8))
                    Set oMatch = oMatches.Item(0)
                    m_sAccountNumber = oMatch.SubMatches(0)
                    m_saIndexValues(1) = m_sAccountNumber
                Else
                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
                End If
                
                ' Gets the document date
                sRawDate = Mid(sCurrentPageArray(5), 72, 8)
                oDocDate = CDate(sRawDate)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
                m_saIndexValues(2) = m_sDocumentDate
                
                ' Gets the name that this document is addressed to
                m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                
                ' Gets the amount due - this is used in the OCR line
                oRegex.Pattern = "^\s+ Amount Due:\s+(\d.*\.\d{2}|\.\d{2})$"
                If oRegex.Test(sCurrentPageArray(10)) = True Then
                    Set oMatches = oRegex.Execute(sCurrentPageArray(10))
                    Set oMatch = oMatches.Item(0)
                    m_cAmountDue = CCur(oMatch.SubMatches(0))
                Else
                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find amount due field.  This is required for OCR."
                End If
                
'            Case "CU1610P1"
'                ' Client requested file name index
'                m_saIndexValues(0) = "FOS-Notice of Address Change"
'
'                ' Get the account number
'                m_saIndexValues(1) = vbNullString ' no information about accounts in this notice
'
'                ' Get the document date
'                oRegex.Pattern = "^\s+(\d{1,2}-\d{1,2}-\d{2})$"
'                If oRegex.Test(sCurrentPageArray(1)) = True Then
'                    Set oMatches = oRegex.Execute(sCurrentPageArray(1))
'                    Set oMatch = oMatches.Item(0)
'                    sRawDate = oMatch.SubMatches(0)
'                    oDocDate = CDate(sRawDate)
'                    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
'                    m_saIndexValues(2) = m_sDocumentDate
'                Else
'                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find notice date on page."
'                    BuildNationwidePieceIndex = False
'                    Exit Function
'                End If
'
'                ' Get the name that the document is addressed to
'                m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                
            Case "LN0111Q1", "E_LN0111Q1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Stop Payment Notice (LN111Q1)"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                oRegex.Pattern = "^\s+ACCOUNT NBR\s+(\d+)\s.*$"
                If oRegex.Test(sCurrentPageArray(19)) = True Then
                    Set oMatches = oRegex.Execute(sCurrentPageArray(19))
                    Set oMatch = oMatches.Item(0)
                    m_sAccountNumber = oMatch.SubMatches(0)
                    m_saIndexValues(1) = m_sAccountNumber
                Else
                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
                    BuildNationwidePieceIndex = False
                    Exit Function
                End If
                
                ' Gets the document date
                oDocDate = FileDateTime(m_sInputFile(m_lInputFilePtr))
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
                m_saIndexValues(2) = m_sDocumentDate
                
                ' Gets the name that the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name

        End Select
    ElseIf m_sXml_FileGroup = "DEPOSITS" Then
        Select Case m_sCurrentFileName
        
        '--------------------------------------------------------------------------------------------------
        ' THIS IS THE BEGINNING OF THE DEPOSIT NOTICES
        '--------------------------------------------------------------------------------------------------
        
            Case "CU1610P1"
                ' Client requested file name index
                m_saIndexValues(0) = "FOS-Notice of Address Change"
                
                ' Get the account number
                m_saIndexValues(1) = vbNullString ' no information about accounts in this notice
                
                ' Get the document date
                oRegex.Pattern = "^\s+(\d{1,2}-\d{1,2}-\d{2})$"
                If oRegex.Test(sCurrentPageArray(1)) = True Then
                    Set oMatches = oRegex.Execute(sCurrentPageArray(1))
                    Set oMatch = oMatches.Item(0)
                    sRawDate = oMatch.SubMatches(0)
                    oDocDate = CDate(sRawDate)
                    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
                    m_saIndexValues(2) = m_sDocumentDate
                Else
                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find notice date on page."
                    BuildNationwidePieceIndex = False
                    Exit Function
                End If
                
                ' Get the name that the document is addressed to
                m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
                
            Case "TA0180P1", "E_TA0180P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Notice of Insufficient Funds"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' GEts the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA0183P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Notice of Returned Item"
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' GEts the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA0420P1", "E_TA0420P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Money Market Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA0420P2", "E_TA0420P2"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Savings Excessive Withdrawal Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0360P2", "E_TM0360P2"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Automatic Renewal Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0360P7", "E_TM0360P7"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Pre-Renewal Notice - Regulation DD Accounts <12M"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0360P8", "E_TM0360P8"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Pre-Renewal Notice Regulation DD Accounts >12M"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the document date
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA1161P1", "E_TA1161P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-ODP Notice of Transfer"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                ' The get page routine already takes care of getting the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                ' The get page routine already takes care of getting the account number
                oDocDate = CDate(m_sDocumentDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0330P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Retirement Account Statement"
                
                ' Gets the account number
                If Mid(m_saCurrentPageLine(16), 1, Len(scAccountNumber)) = scAccountNumber Then
                    m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(17), 1, Len(scAccountNumber)))
                End If
                
                ' Check to see if the account number was initiall found - some bryn mawr statements are stupid and don't have a
                ' "Summary of your investment" section at the beginning of page 1
                Dim i As Long
                If Len(m_sAccountNumber) = 0 Then
                    For i = 0 To m_lCurrentPageIdx
                        If InStr(1, m_saCurrentPageLine(i), scActivityHeader) Then
                            If InStr(1, m_saCurrentPageLine(i + 1), scAccountNumberAct) > 0 Then
                                m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(i + 2), InStr(1, m_saCurrentPageLine(i + 1), scAccountNumberAct), Len(scAccountNumberAct)))
                            End If
                            Exit For
                        End If
                        
                    Next
                End If
                
                If Len(Trim(m_sAccountNumber)) = 0 Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate account number on TM0330P1 page."
                End If
                
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                oRegex.Pattern = "^\s+(\d{1,2}\-\d{2}\-\d{2})$"
                If oRegex.Test(sCurrentPageArray(9)) = False Then
                    Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find document date on page."
                    BuildNationwidePieceIndex = False
                    Exit Function
                End If
                
                Set oMatches = oRegex.Execute(sCurrentPageArray(9))
                Set oMatch = oMatches.Item(0)
                sRawDate = oMatch.SubMatches(0)
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0365P1", "E_TM0365P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Certificate of Deposit Reinvestment Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
            
            Case "TA95603P1", "E_TA956031"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Product type change notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95602P3", "E_TA956023"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Dormant Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95602P4", "E_TA956024"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Active Notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "CH931613P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Charge Back Notice"
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA0050P1", "E_TA0050P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Stop Payment Letter"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95599P1", "E_TA955991"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran unfunded zero balance notice 1"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
    
            Case "TA95599P2", "E_TA955992"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran unfunded zero balance notice 2"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95599P3", "E_TA955993"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran unfunded zero balance notice 3"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95599P4", "E_TA955994"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran zero balance notice 1"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95599P5", "E_TA955995"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran zero balance notice 2"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95599P6", "E_TA955996"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Tran zero balance notice 3"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM95604P1", "E_TM956041"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Time CD funded notice"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM95599P1", "E_TM955991"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Time unfunded zero balance notice 1"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM95599P2", "E_TM955992"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Time unfunded zero balance notice 2"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM95599P3", "E_TM955993"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Time unfunded zero balance notice 3"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "FT0160P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Extended Hold Notice"
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "FT0160P2"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Extended Hold Notice"
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "FT0160P3"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Extended Hold Notice"
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TM0847P1"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Required Minimum Distribution Notice"
                
                ' Gets the account number
                m_saIndexValues(1) = vbNullString
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case "TA95602P1", "E_TA956021"
                ' Gets the file name
                m_saIndexValues(0) = "FOS-Active Notice 1"
                m_saIndexValues(0) = IIf(Left(m_sCurrentFileName, 2) = "E_", m_saIndexValues(0) & " Electronic", m_saIndexValues(0))
                
                ' Gets the account number
                m_saIndexValues(1) = m_sAccountNumber
                
                ' Gets the document date
                sRawDate = m_sDocumentDate
                oDocDate = CDate(sRawDate)
                m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
                
                ' Gets the name the document is addressed to
                m_saIndexValues(3) = m_oDeliveryAddress.Name
                
            Case Else
                Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Cannot generate Nationwide Bank piece index - file name not recognized: " & m_sCurrentFileName
                Exit Function
        
        End Select
    Else
        Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Cannot generate Nationwide Bank piece index - unrecognized file group is configured: " & m_sXml_FileGroup
        Exit Function
    End If
    
    ' Success
    BuildNationwidePieceIndex = True
    
    Exit Function
    
BuildNationwidePieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildNationwidePieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildNationwidePieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildNationwidePieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function

Private Function BuildEddPieceIndex(sCurrentPageArray() As String, lCurrentPageIdx As Long) As Boolean
    On Error GoTo BuildEddPieceIndex_Error

    ' Constants
    Const lcIndex As Long = 3
    
    'Regular expression used to pull information out of the piece
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date

    ' Set default return value
    BuildEddPieceIndex = False
    m_lIndexIdx = lcIndex
    
    ' Initializes the piece index arrays
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    m_saIndexKeys(0) = "AccountNumber"
    m_saIndexKeys(1) = "CustomerName"
    m_saIndexKeys(2) = "StatementDate"
    m_saIndexKeys(3) = "ReportID"
    
    ' Validate that we are Nationwide Bank
    If Not (m_sXmlInst = "Nationwide" Or m_sXmlInst = "SNB") Then
        Err.Raise g_lcErrParameterError, App.Title, "Inst does not match 'Nationwide' or 'SNB' - cannot continue."
        Exit Function
    End If
    
    ' Validate that this is a e-notice file
    If Left(m_sCurrentFileName, Len("E_")) <> "E_" Then
        Err.Raise g_lcErrParameterError, App.Title, "Current file is not an e-notice file - must start with 'E_'.  Cannot continue."
        Exit Function
    End If
    
    ' Each file and form type varies in layout - look in different locations for the required information
    Select Case m_sCurrentFileName
        Case Else
            ' Gets the account number
            If Len(Trim(m_sAccountNumber)) = 0 Then
                MessageBox "Invalid page - account number cannot be blank for e notice file.  File is " & GetFile(m_sInputFile(m_lInputFilePtr)) & " (" & m_sCurrentFileName & ")"
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            m_saIndexValues(0) = TrimLeadingZeros(m_sAccountNumber)
            
            ' Gets the customer name
            If Len(Trim(m_oDeliveryAddress.Name)) = 0 Then
                MessageBox "Invalid page - delivery address name cannot be blank for e notice file.  File is " & GetFile(m_sInputFile(m_lInputFilePtr)) & " (" & m_sCurrentFileName & ")"
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            m_saIndexValues(1) = m_oDeliveryAddress.Name
            
            ' Gets the statement date
            If Len(Trim(m_sDocumentDate)) = 0 Then
                MessageBox "Invalid page - document date cannot be blank for e notice file.  File is " & GetFile(m_sInputFile(m_lInputFilePtr)) & " (" & m_sCurrentFileName & ")"
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            sRawDate = m_sDocumentDate
            oDocDate = CDate(sRawDate)
            m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
    End Select
    
    ' Adds the current file report type
    m_saIndexValues(3) = m_sCurrentFileName
    
    ' Sets the EDD Cycle Date if it hasn't been set already
    If m_sEddCycleDate = 0 Then
        m_sEddCycleDate = CDate(m_saIndexValues(2))
    End If
    
    ' Success
    BuildEddPieceIndex = True
    
    Exit Function
    
BuildEddPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildEddPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildEddPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildEddPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function

Private Function BuildBpbPieceIndex( _
            ByVal sFileName As String, _
            ByVal sAccountNumber As String, _
            ByVal sDocumentDate As String, _
            ByVal sDocumentType As String, _
            ByVal sName As String) As Boolean
            
    On Error GoTo BuildBpbPieceIndex_Error

    Const lcIndex As Long = 4

    BuildBpbPieceIndex = False
    m_lIndexIdx = lcIndex
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    
    If m_sXmlInst <> "BPB" Then
        Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Inst does not match 'BPB' - cannot continue."
        Exit Function
    End If
    
    m_saIndexKeys(0) = "File_Name"
    m_saIndexValues(0) = sFileName
    
    m_saIndexKeys(1) = "Account_Number"
    m_saIndexValues(1) = sAccountNumber
    
    m_saIndexKeys(2) = "Document_Date"
    m_saIndexValues(2) = sDocumentDate
    
    m_saIndexKeys(3) = "Document_Type"
    m_saIndexValues(3) = sDocumentType
    
    m_saIndexKeys(4) = "Name"
    m_saIndexValues(4) = sName
    
    BuildBpbPieceIndex = True
    
    Exit Function
    
BuildBpbPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildBpbPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildBpbPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildBpbPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function

Private Function BuildIRAPieceIndex(ByVal sAccountNumber As String, ByVal sDocumentDate As String, ByVal sDocumentType As String, ByVal sName As String) As Boolean
            
    On Error GoTo BuildIRAPieceIndex_Error

    Const lcIndex As Long = 5

    BuildIRAPieceIndex = False
    m_lIndexIdx = lcIndex
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    
    m_saIndexKeys(0) = "Form_Type"
    m_saIndexValues(0) = sDocumentType
    
    m_saIndexKeys(1) = "Primary_TIN"
    m_saIndexValues(1) = ""
    
    m_saIndexKeys(2) = "Primary_Name"
    m_saIndexValues(2) = sName
    
    m_saIndexKeys(3) = "Tax_Year"
    m_saIndexValues(3) = sDocumentDate
    
    m_saIndexKeys(4) = "Account_Number"
    m_saIndexValues(4) = sAccountNumber
    
    m_saIndexKeys(5) = "File_Name"
    m_saIndexValues(5) = Right(Replace(Space(10), " ", "0") & Trim(Str(m_lInputSequence) + 1) & ".pdf", 14)
    
    
    BuildIRAPieceIndex = True
    
    Exit Function
    
BuildIRAPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildIRAPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildIRAPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildIRAPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function

Private Function BuildPieceIndex( _
            ByVal sFileName As String, _
            ByVal sAccountNumber As String, _
            ByVal sDocumentDate As String, _
            ByVal sName As String) As Boolean
            
    On Error GoTo BuildPieceIndex_Error

    Const lcIndex As Long = 3

    BuildPieceIndex = False
    m_lIndexIdx = lcIndex
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    
    m_saIndexKeys(0) = "File_Name"
    m_saIndexValues(0) = sFileName
    
    m_saIndexKeys(1) = "Document_Date"
    m_saIndexValues(1) = sDocumentDate
    
    m_saIndexKeys(2) = "Account_Number"
    m_saIndexValues(2) = sAccountNumber
    
    m_saIndexKeys(3) = "Name"
    m_saIndexValues(3) = sName
    
    BuildPieceIndex = True
    
    Exit Function
    
BuildPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function

Private Function BuildRiverCityPieceIndex(sCurrentPageArray() As String, lCurrentPageIdx As Long) As Boolean
    On Error GoTo BuildRiverCityPieceIndex_Error

    ' Constants
    Const lcIndex As Long = 3
    
    'Regular expression used to pull information out of the piece
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim sRawDate As String
    Dim oDocDate As Date

    ' Set default return value
    BuildRiverCityPieceIndex = False
    m_lIndexIdx = lcIndex
    
    ' Initializes the piece index arrays
'    ReDim m_saIndexKeys(m_lIndexIdx)
'    ReDim m_saIndexValues(m_lIndexIdx)
'    m_saIndexKeys(0) = "File_Name"
'    m_saIndexKeys(1) = "Account_Number"
'    m_saIndexKeys(2) = "Document_Date"
'    m_saIndexKeys(3) = "Name"
    
    ' Validate that we are RiverCity Bank
    If m_sXmlInst <> "RiverCity" Then
        Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Inst does not match 'RiverCity' - cannot continue."
        Exit Function
    End If
    
    ' Each file and form type varies in layout - look in different locations for the required information
    Select Case m_sCurrentFileName
        Case "LN0480Q1"
            'm_saIndexValues(0) = m_sCurrentFileName
            
            ' Gets the account number
            If Trim(sCurrentPageArray(1)) = "ACCOUNT NUMBER" Then
                m_sAccountNumber = Trim(Mid(sCurrentPageArray(2), 65))
                'm_saIndexValues(1) = m_sAccountNumber
            Else
                Err.Raise g_lcErrBase, g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
            
            ' Gets the document date
            sRawDate = Mid(sCurrentPageArray(5), 66, 8)
            'sRawDate = Mid(sRawDate, 1, 2) & "/" & Mid(sRawDate, 3, 2) & "/" & Mid(sRawDate, 5, 2)
            oDocDate = CDate(sRawDate)
            m_sDocumentDate = Format(oDocDate, "yyyy-MM-dd")
            'm_saIndexValues(2) = m_sDocumentDate
            
            ' Gets the first line of the address
            'm_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
        Case Else
            Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Cannot generate RiverCity Bank piece index - file name not recognized: " & m_sCurrentFileName
            Exit Function
    End Select
    
    If Len(m_sXML_EddJobKey) > 0 Or m_bXML_EddEstatements Then
        Dim sAccountNumber As String
        If Len(m_sAccountNumber) < 10 Then
            sAccountNumber = Right(String(10, "0") & m_sAccountNumber, 10)
        Else
            sAccountNumber = m_sAccountNumber
        End If
        
        m_lIndexIdx = 2
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
        
        m_saIndexKeys(0) = "Account_Number"
        m_saIndexValues(0) = sAccountNumber
        
        m_saIndexKeys(1) = "PrimaryName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name
        
        m_saIndexKeys(2) = "Statement_Date"
        m_saIndexValues(2) = m_sDocumentDate
        
        If Len(m_sXML_AccountType) > 0 Then
            m_lIndexIdx = m_lIndexIdx + 2
            ReDim Preserve m_saIndexKeys(m_lIndexIdx)
            ReDim Preserve m_saIndexValues(m_lIndexIdx)
            m_saIndexKeys(m_lIndexIdx - 1) = "Account_Type"
            m_saIndexValues(m_lIndexIdx - 1) = m_sXML_AccountType
            m_saIndexKeys(m_lIndexIdx) = "EDDDocumentType"
            m_saIndexValues(m_lIndexIdx) = m_sXML_AccountType & "-" & sAccountNumber
        End If
    Else
        m_lIndexIdx = 3
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)
        
        m_saIndexKeys(0) = "File_Name"
        m_saIndexValues(0) = m_sCurrentFileName
        
        m_saIndexKeys(1) = "Account_Number"
        m_saIndexValues(1) = m_sAccountNumber
        
        m_saIndexKeys(2) = "Document_Date"
        m_saIndexValues(2) = m_sDocumentDate
        
        m_saIndexKeys(3) = "Name"
        m_saIndexValues(3) = m_oDeliveryAddress.Name
    End If
        
    
    ' Success
    BuildRiverCityPieceIndex = True
    
    Exit Function
    
BuildRiverCityPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildRiverCityPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildRiverCityPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildRiverCityPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function
Private Function BuildSNBEddPieceIndex(sCurrentPageArray() As String, lCurrentPageIdx As Long) As Boolean

    On Error GoTo BuildSNBEddPieceIndex_Error



    
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oDocDate As Date
    Dim DashAccount As String
    ' Set default return value
    BuildSNBEddPieceIndex = False

    Select Case m_sCurrentFileName

            
     Case "E_LN0470P1"

            ' Gets the account number
            oRegex.Pattern = "\s+NOTE NUMBER\s+(\d+)\s*$"
            If oRegex.Test(sCurrentPageArray(2)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(2))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "(\s+\d{1,2}\/\d{1,2}\/\d{1,2})\s+"
            If oRegex.Test(sCurrentPageArray(0)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(0))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
        
      Case "E_LN395Q1"
    
            ' Gets the account number
            oRegex.Pattern = "Loan Number\s+(\d+)"
            If oRegex.Test(sCurrentPageArray(7)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(7))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "\s+(\d{1,2}\/\d{1,2}\/\d{2})$"
            If oRegex.Test(sCurrentPageArray(1)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(1))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
                   
        Case "E_TA0185P1"

            ' Gets the document date
            oRegex.Pattern = "\s*(\d{1,2}-\d{1,2}-\d{4})$"
            If oRegex.Test(sCurrentPageArray(0)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(0))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
        
        Case "E_TA0180P1"
        
            ' Gets the account number
            oRegex.Pattern = "ACCT NBR:\s+(\d+)$"
            If oRegex.Test(sCurrentPageArray(2)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(2))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "\s+(\d{1,2}\/\d{1,2}\/\d{4})$"
            If oRegex.Test(sCurrentPageArray(1)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(1))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
        
        Case "E_TM0360P5"
        
            ' Gets the account number
            oRegex.Pattern = "CERTIFICATE\s+(\d+)$"
            If oRegex.Test(sCurrentPageArray(14)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(14))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "\s*(\d{1,2}-\d{1,2}-\d{4})$"
            If oRegex.Test(sCurrentPageArray(0)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(0))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If

        Case "E_TM0360P8"
        
            ' Gets the account number
            oRegex.Pattern = "CERTIFICATE\s+(\d+)$"
            If oRegex.Test(sCurrentPageArray(13)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(13))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "\s*(\d{1,2}-\d{1,2}-\d{4})$"
            If oRegex.Test(sCurrentPageArray(0)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(0))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
        Case "E_TM0360P7"
        
            ' Gets the account number
            oRegex.Pattern = "CERTIFICATE\s+(\d+)$"
            If oRegex.Test(sCurrentPageArray(13)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(13))
                m_sAccountNumber = oMatches.Item(0).SubMatches(0)
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
    
            ' Gets the document date
            oRegex.Pattern = "\s*(\d{1,2}-\d{1,2}-\d{4})$"
            If oRegex.Test(sCurrentPageArray(0)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(0))
                oDocDate = oMatches.Item(0).SubMatches(0)
                m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
            Else
                Err.Raise g_lcErrInvalidPage, App.Title, "Cannot find date on page."
            End If
        
    Case Else
            Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Cannot generate SNB EDD piece index - file name not recognized: " & m_sCurrentFileName
         Exit Function
    End Select
    
    If Len(m_sXML_EddJobKey) > 0 Then
        m_lIndexIdx = 4
        ReDim m_saIndexKeys(m_lIndexIdx)
        ReDim m_saIndexValues(m_lIndexIdx)

        m_saIndexKeys(0) = "AccountNumber"
        m_saIndexValues(0) = m_sAccountNumber

        m_saIndexKeys(1) = "CustomerName"
        m_saIndexValues(1) = m_oDeliveryAddress.Name

        m_saIndexKeys(2) = "StatementDate"
        m_saIndexValues(2) = Format(oDocDate, "MM-dd-yyyy")
        
        m_saIndexKeys(3) = "ReportID"
        m_saIndexValues(3) = m_sCurrentFileName

        m_saIndexKeys(4) = "EDDDocumentType"
        m_saIndexValues(4) = m_sXML_EddProductId & "-" & m_sAccountNumber


    End If
    
    ' Success
    BuildSNBEddPieceIndex = True
    
    Exit Function
    
BuildSNBEddPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildSNBEddPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildSNBEddPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildSNBEddPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function


Private Function BuildSNBPieceIndex(sCurrentPageArray() As String, lCurrentPageIdx As Long) As Boolean
    On Error GoTo BuildSNBPieceIndex_Error

    ' Constants
    Const lcIndex As Long = 3
    
    'Regular expression used to pull information out of the piece
    Dim oRegex As RegExp
    Set oRegex = New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim sRawDate As String
    Dim dtDocDate As Date

    ' Set default return value
    BuildSNBPieceIndex = False
    m_lIndexIdx = lcIndex
    
    ' Initializes the piece index arrays
    ReDim m_saIndexKeys(m_lIndexIdx)
    ReDim m_saIndexValues(m_lIndexIdx)
    m_saIndexKeys(0) = "File_Name"
    m_saIndexKeys(1) = "Account_Number"
    m_saIndexKeys(2) = "Document_Date"
    m_saIndexKeys(3) = "Name"
    
    ' Validate that we are Nationwide Bank
    If m_sXmlInst <> "SNB" Then
        Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Inst does not match 'SNB' - cannot continue."
        Exit Function
    End If
    
    ' Each file and form type varies in layout - look in different locations for the required information
    Select Case m_sCurrentFileName
        Case "CH0202P1"
            m_saIndexValues(0) = m_sCurrentFileName
            
            ' Gets the account number
            oRegex.Pattern = " ACCOUNT\s+(\d+)"
            If oRegex.Test(sCurrentPageArray(16)) = True Then
                Set oMatches = oRegex.Execute(sCurrentPageArray(16))
                Set oMatch = oMatches.Item(0)
                m_sAccountNumber = oMatch.SubMatches(0)
                m_saIndexValues(1) = m_sAccountNumber
            Else
                Err.Raise g_lcErrBase, g_lcErrInvalidPage, App.Title, "Cannot find account number on page."
            End If
            
            ' Gets the document date
            sRawDate = Trim(Mid(sCurrentPageArray(0), 70))
            dtDocDate = CDate(sRawDate)
            m_sDocumentDate = Format(dtDocDate, "MM-dd-yyyy")
            m_saIndexValues(2) = m_sDocumentDate
            
            ' Gets the first line of the address
            m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
            
        Case "LN0470P1"
            
            ' Gets the account number
            oRegex.Pattern = "NOTE NUMBER\s+(\d+)"
            If oRegex.Test(m_saCurrentPageLine(2)) = False Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate NOTE NUMBER on LN0470P page"
            End If
            Set oMatch = oRegex.Execute(m_saCurrentPageLine(2)).Item(0)
            m_sAccountNumber = oMatch.SubMatches(0)
            
            ' Get the document date
            oRegex.Pattern = "(\d{1,2}/\d{1,2}/\d{2})"
            If oRegex.Test(m_saCurrentPageLine(0)) = False Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate notice date on LN0470P page"
            End If
            sRawDate = oRegex.Execute(m_saCurrentPageLine(0)).Item(0)
            dtDocDate = CDate(sRawDate)
            m_sDocumentDate = Format(dtDocDate, "MM-dd-yyyy")
            
            m_saIndexValues(0) = m_sCurrentFileName
            m_saIndexValues(1) = m_sAccountNumber
            m_saIndexValues(2) = m_sDocumentDate
            m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
            
        Case "TA0420P1"
            
            ' Gets the account number
            If Mid(sCurrentPageArray(13), 59, 6) = "NUMBER" Then
                m_sAccountNumber = Trim(Mid(sCurrentPageArray(14), 50, 16))
            End If
            
            If Mid(sCurrentPageArray(0), 68, 4) = "DATE" Then
                oRegex.Pattern = "(\d{1,2}-\d{1,2}-\d{4})"
                If oRegex.Test(sCurrentPageArray(1)) = False Then
                    Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Unable to locate notice date on TA0420P1 page"
                End If
                sRawDate = oRegex.Execute(sCurrentPageArray(1)).Item(0)
                dtDocDate = CDate(sRawDate)
                m_sDocumentDate = Format(dtDocDate, "MM-dd-yyyy")
            End If
            
            m_saIndexValues(0) = m_sCurrentFileName
            m_saIndexValues(1) = m_sAccountNumber
            m_saIndexValues(2) = m_sDocumentDate
            m_saIndexValues(3) = Trim(m_oDeliveryAddress.Name)
            
        Case Else
            Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Cannot generate Nationwide Bank piece index - file name not recognized: " & m_sCurrentFileName
            Exit Function
    End Select
        
    
    ' Success
    BuildSNBPieceIndex = True
    
    Exit Function
    
BuildSNBPieceIndex_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.BuildSNBPieceIndex"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo BuildSNBPieceIndex_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo BuildSNBPieceIndex_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
    
End Function


Private Function GetPage_LN93357P() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    On Error GoTo GetPage_LN93357P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrX + 4 Then
        m_lCurrentPageIdx = lcAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    If Trim(m_saCurrentPageLine(1)) = "ACCOUNT NUMBER         PAYMENT DUE DATE                            MINIMUM" Then
        'This code fixes a problem at CBS where they are not including all needed blank lines.
        InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 10, 5
        InsertLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 6
    End If
    
    m_sAccountNumber = Trim(Mid(m_saCurrentPageLine(8), 1, 22))
    m_cPaymentAmount = CCur(Trim(Mid(m_saCurrentPageLine(9), 60)))
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 95, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, lcAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    Dim bRaiseLines As Boolean
    bRaiseLines = True
    Dim i As Long
    For i = 0 To 2
        If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
            bRaiseLines = False
        End If
    Next
    
    If bRaiseLines Then
        ScrollLinesInArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 18, 3
    End If
    
    GetPage_LN93357P = True
    
    Exit Function
    
GetPage_LN93357P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN93357P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN93357P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN93357PTrailer() As Boolean

    Const lcAddrX = 5
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    GetPage_LN93357PTrailer = False
    
    On Error GoTo GetPage_LN93357PTrailer_Error
    
    If m_lReadAheadPageIdx < 10 Then
        Exit Function
    End If
    
    If Trim(m_saReadAheadPageLine(7)) = "ACCOUNT NUMBER         PAYMENT DUE DATE                            MINIMUM" Then
        Exit Function
    End If
    If Trim(m_saReadAheadPageLine(7)) = "ACCOUNT NUMBER         PAYMENT DUE DATE                             TOTAL" Then
        Exit Function
    End If
    
    ReadPageAhead
    
    Set m_oDeliveryAddress = Nothing
    Set m_oReturnAddress = Nothing
    
    Select Case m_sXmlInst
        Case "Nationwide"
            If m_sCurrentFileName = "LN0480Q1" Or m_sCurrentFileName = "E_QSYSPRT" Then
                Set m_oPieceBacker = Nothing
                m_bDuplex = False
            End If
        Case m_scInst_29175100_FinFactors
            If m_sCurrentFileName = "LN0480Q1" Then
                Set m_oPieceBacker = Nothing
                m_bDuplex = False
            End If
    End Select
    
    GetPage_LN93357PTrailer = True
    
    Exit Function
    
GetPage_LN93357PTrailer_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN93357P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN93357PTrailer_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN93357PTrailer_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0050P1(Optional ByVal bMaskAcctNum As Boolean) As Boolean

    Const lcAddrX = 11
    Const lcAddrLine = 11
    Const lcAddrLen = 40
    
    On Error GoTo GetPage_TA0050P1_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrX + 4 Then
        m_lCurrentPageIdx = lcAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 76, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, lcAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+ACCOUNT NBR\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(19)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TA0050P1 page."
        GetPage_TA0050P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(19))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(19) = Replace(m_saCurrentPageLine(19), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+DATE RECEIVED:\s+(\d{1,2}/\d{1,2}/\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(21)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TA0050P1 page."
        GetPage_TA0050P1 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(21))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA0050P1 = True
    
    Exit Function
    
GetPage_TA0050P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0050P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0050P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0050P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_HC5185(Optional ByVal bMaskAcctNum As Boolean) As Boolean
    
    Dim lAddrX As Long
    Dim lAddrLine As Long
    Dim lAddrLen As Long

    On Error GoTo GetPage_HC5185_Error
    
    InitializePieceFields
    
    Select Case m_sXmlInst
        Case Else
            lAddrX = 11
            lAddrLine = 11
            lAddrLen = 40
    End Select
    
    If m_sXmlInst = "Nationwide" Then
        ' Remove the return address and delete top 4 lines
        m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), 8, 40, PvRTrimOutput)
        m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), 8, 40, PvRTrimOutput)
        m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), 8, 40, PvRTrimOutput)
        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 0, 4
    End If
    
    If m_lCurrentPageIdx < lAddrX + 4 Then
        m_lCurrentPageIdx = lAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 76, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lAddrLine, lAddrX, lAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+ACCOUNT NBR\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(19)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on HC5185 page."
        GetPage_HC5185 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(19))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(19) = Replace(m_saCurrentPageLine(19), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+DATE RECEIVED:\s+(\d{1,2}/\d{1,2}/\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(21)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on HC5185 page."
        GetPage_HC5185 = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(21))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_HC5185 = True
    
    Exit Function
    
GetPage_HC5185_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_HC5185"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_HC5185_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_HC5185_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TA0096P(Optional ByVal bSuppressRetAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean

    Const lcAddrX = 7
    Const lcAddrLine = 14
    Const lcAddrLen = 40
    
    On Error GoTo GetPage_TA0096P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrX + 4 Then
        m_lCurrentPageIdx = lcAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, lcAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Should we suppress return-address printing
    If bSuppressRetAddr Then
        ScrapeString m_saCurrentPageLine(0), 1, 40
        ScrapeString m_saCurrentPageLine(1), 1, 40
        ScrapeString m_saCurrentPageLine(2), 1, 40
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+ACCOUNT NUMBER:\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(13)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TA0096P page."
        GetPage_TA0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(13))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(13) = Replace(m_saCurrentPageLine(13), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}-\d{1,2}-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TA0096P page."
        GetPage_TA0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TA0096P = True
    
    Exit Function
    
GetPage_TA0096P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TA0096P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TA0096P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TA0096P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM0096P(Optional ByVal bSuppressRetAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean) As Boolean

    Const lcAddrX = 7
    Const lcAddrLine = 14
    Const lcAddrLen = 40
    
    On Error GoTo GetPage_TM0096P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrX + 4 Then
        m_lCurrentPageIdx = lcAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    m_sRegion = Trim(Mid(m_saCurrentPageLine(0), 76, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, lcAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Should we suppress return-address printing
    If bSuppressRetAddr Then
        ScrapeString m_saCurrentPageLine(0), 1, 40
        ScrapeString m_saCurrentPageLine(1), 1, 40
        ScrapeString m_saCurrentPageLine(2), 1, 40
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+ACCOUNT NUMBER:\s+(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(13)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find account number on TM0096P page."
        GetPage_TM0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(13))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(13) = Replace(m_saCurrentPageLine(13), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}-\d{1,2}-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on TM0096P page."
        GetPage_TM0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    GetPage_TM0096P = True
    
    Exit Function
    
GetPage_TM0096P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0096P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0096P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0096P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_LN0096P(Optional ByVal bSuppressRetAddr As Boolean, Optional ByVal bMaskAcctNum As Boolean, Optional ByVal bSqueezeNotice As Boolean) As Boolean

    Const lcAddrX = 1
    Const lcAddrLine = 14
    Const lcAddrLen = 40
    
    On Error GoTo GetPage_LN0096P_Error
    
    InitializePieceFields
    
    If m_lCurrentPageIdx < lcAddrX + 4 Then
        m_lCurrentPageIdx = lcAddrX + 4
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    'm_sRegion = Trim(Mid(m_saCurrentPageLine(0), 76, 3))
    
    Set m_oDeliveryAddress = GetDeliveryAddress(m_saCurrentPageLine, m_lCurrentPageIdx, lcAddrLine, lcAddrX, lcAddrLen, True)
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    ' Should we suppress return-address printing
    If bSuppressRetAddr Then
        ScrapeString m_saCurrentPageLine(0), 1, 40
        ScrapeString m_saCurrentPageLine(1), 1, 40
        ScrapeString m_saCurrentPageLine(2), 1, 40
    End If
    
    ' Variables used for validating the page
    Dim oRegexTest As RegExp
    Dim oRegexMatches As MatchCollection
    Dim oRegexMatch As Match
    Dim sMaskedAcctNum As String
    Dim sRawDate As String
    Dim oDocDate As Date
    
    ' Gets the account number off the page and verifies that it can be found
    Set oRegexTest = New RegExp
    oRegexTest.IgnoreCase = True
    oRegexTest.Pattern = "^\s+LOAN NUMBER"
    If oRegexTest.Test(m_saCurrentPageLine(27)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find loan number label on LN0096P page."
        GetPage_LN0096P = False
        Exit Function
    End If
    oRegexTest.Pattern = "^\s*(\d+)"
    If oRegexTest.Test(m_saCurrentPageLine(30)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find loan number on LN0096P page."
        GetPage_LN0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(30))
    Set oRegexMatch = oRegexMatches.Item(0)
    m_sAccountNumber = oRegexMatch.SubMatches(0)
    
    ' Masks the account number if we are to mask it
    If bMaskAcctNum Then
        If Len(m_sAccountNumber) > 4 Then
            sMaskedAcctNum = m_sAccountNumber
            Mid(sMaskedAcctNum, 1, Len(sMaskedAcctNum) - 4) = String(Len(sMaskedAcctNum) - 4, "#")
            m_saCurrentPageLine(13) = Replace(m_saCurrentPageLine(13), m_sAccountNumber, sMaskedAcctNum)
        End If
    End If
    
    ' Gets the date received/document date off the page and verifies that it can be found
    oRegexTest.Pattern = "^\s+(\d{1,2}-\d{1,2}-\d{4})"
    If oRegexTest.Test(m_saCurrentPageLine(0)) = False Then
        Err.Raise g_lcErrInvalidPage, App.Title, "Invalid page encountered - cannot find date received on LN0096P page."
        GetPage_LN0096P = False
        Exit Function
    End If
    
    Set oRegexMatches = oRegexTest.Execute(m_saCurrentPageLine(0))
    Set oRegexMatch = oRegexMatches.Item(0)
    sRawDate = oRegexMatch.SubMatches.Item(0)
    oDocDate = CDate(sRawDate)
    m_sDocumentDate = Format(oDocDate, "MM-dd-yyyy")
    
    ' Squeezes out blank lines in the middle of the notice body
    If bSqueezeNotice Then
        DeleteLinesFromArray m_saCurrentPageLine, m_saCurrentPageLineBold, m_lCurrentPageIdx, 2, 18
    End If
    
    GetPage_LN0096P = True
    
    Exit Function
    
GetPage_LN0096P_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_LN0096P"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_LN0096P_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_LN0096P_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function GetPage_TM0847P1() As Boolean
    Const lcAddrX As Long = 1
    Const lcAddrLen As Long = 40
'
'    Const scAccountNbr As String = "BRANCH   BOX"
    
    Dim sAssertionString As String
    
    On Error GoTo GetPage_TM0847P1_Error
    
    If m_lCurrentPageIdx < 20 Then
        m_lCurrentPageIdx = 20
        ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
        ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    End If
    
    InitializePieceFields
    
    'For Boston Private Bank
    m_sRegion = ScrapeString(m_saCurrentPageLine(0), 95, 3, PvRTrimOutput)
    
    
    m_sLetterDate = vbNullString
    
'
'    StringAssertion m_saCurrentPageLine(5), scAccountNbr, 59, 6
'
'    m_sAccountNumber = ReadString(m_saCurrentPageLine(6), 59, 0, PvTrimOutput)
    
    m_oReturnAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(0), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oReturnAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(1), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oReturnAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(2), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oReturnAddress.AddressLine(4) = ScrapeString(m_saCurrentPageLine(3), lcAddrX, lcAddrLen, PvRTrimOutput)
    
    m_sLetterDate = ScrapeString(m_saCurrentPageLine(6), 1, 20, PvRTrimOutput)
    m_sDocumentDate = Right(m_sLetterDate, 10)
    
    m_oDeliveryAddress.AddressLine(1) = ScrapeString(m_saCurrentPageLine(11), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oDeliveryAddress.AddressLine(2) = ScrapeString(m_saCurrentPageLine(12), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oDeliveryAddress.AddressLine(3) = ScrapeString(m_saCurrentPageLine(13), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oDeliveryAddress.AddressLine(4) = ScrapeString(m_saCurrentPageLine(14), lcAddrX, lcAddrLen, PvRTrimOutput)
    m_oDeliveryAddress.AddressLine(5) = ScrapeString(m_saCurrentPageLine(15), lcAddrX, lcAddrLen, PvRTrimOutput)
    
    'Validate all data scraped from top of page.
    Dim i As Long
    For i = 0 To 17
        If Len(Trim(m_saCurrentPageLine(i))) > 0 Then
            Stop
        End If
    Next
    
    m_saCurrentPageLine(15) = Space(76 - Len(m_sLetterDate)) & m_sLetterDate
    
    Set m_oCassAddress = m_oDeliveryAddress.Clone
    
    m_lStackItems = StackItemsFlatStock
    
    GetPage_TM0847P1 = True
    
    Exit Function
    
GetPage_TM0847P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.GetPage_TM0847P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo GetPage_TM0847P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo GetPage_TM0847P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function PrintPiece_Indirect() As Boolean
    Select Case m_lStackItems
        Case StackItemsPressureSeal
            PrintPiece_Indirect = PrintPiece_Notice()
        Case StackItemsNotice
            PrintPiece_Indirect = PrintPiece_Notice()
        Case StackItemsBypass
            PrintPiece_Indirect = PrintPiece_Bypass()
        Case StackItemsFlatStock
            PrintPiece_Indirect = PrintPiece_FlatStock()
        Case StackItemsFlatStock8LPI
            PrintPiece_Indirect = PrintPiece_FlatStock8LPI()
        Case StackItemsFlatStock9LPI
            PrintPiece_Indirect = PrintPiece_FlatStock9LPI()
        Case StackItemsFlatStockLegal
            PrintPiece_Indirect = PrintPiece_FlatStockLegal()
        Case StackItemsFlatStock42
            PrintPiece_Indirect = PrintPiece_FlatStock42()
        Case StackItemsFlatStock428LPI
            PrintPiece_Indirect = PrintPiece_FlatStock428LPI()
        Case StackItemsLN0473
            PrintPiece_Indirect = PrintPiece_LN0473()
        Case StackItemsLN0473_152
            PrintPiece_Indirect = PrintPiece_LN0473_152()
        Case Else
            MessageBox "Assertion error:  Unrecognized report name in PrintPiece_Indirect - cannot continue.", vbCritical + vbOKOnly, App.Title
            Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select
End Function

'==============================================================================
'Function:              OpenPrintStream
'
'Synopsis:              Opens the input file.  Returns true if opened successfully.
'
'==============================================================================
Private Function OpenPrintStream() As Boolean

    On Error GoTo OpenPrintStream_Error
    
    Set m_oPrintStream = New FDS_IO_PrintStream.ANSIStreamReader
    m_oPrintStream.Open m_sInputFile(m_lInputFilePtr)
    OpenPrintStream = True
    
    Exit Function

OpenPrintStream_ErrorExit:

   On Error Resume Next
   Set m_oPrintStream = Nothing
   OpenPrintStream = False
   Exit Function

OpenPrintStream_Error:

   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpContext = Err.HelpContext
   m_sHelpFile = Err.HelpFile

   Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
      Case 1 'User chose Retry
         Resume
      Case 2 'User chose Ignore
         Resume Next
      Case 3 'User chose Abort
         GoTo OpenPrintStream_ErrorExit
   End Select

   Exit Function

End Function

'==============================================================================
'Function:              OpenDneFile
'
'Synopsis:              Opens the DNE file.  Returns true if opened successfully.
'
'==============================================================================
Private Function OpenDneFile(ByVal sDneId As String, ByVal sDneDate As String) As Boolean
    
    On Error GoTo EHandler
    
    'm_sDneBaseFileName = GetFile(AddRemoveExtension(m_sStmtBodyFileName, vbNullString)) & "_" & m_sDneDate
    Dim sDneFileName As String
    sDneFileName = sDneId & "." & sDneDate & ".dne"
    
    Set m_oDneFile = New CFileOutput
    If m_oDneFile.OpenFile(m_sOutputFolder & "\" & sDneFileName) = False Then
        GoSub Rollback
        'Signal failure
        OpenDneFile = False
        Exit Function
    End If
    
    OpenDneFile = True
    
    Exit Function
    
EHandler:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
       Case 1 'User chose Retry
          On Error GoTo EHandler
          Resume
       Case 2 'User chose Ignore
          On Error GoTo EHandler
          Resume Next
       Case 3 'User chose Abort
          GoSub Rollback
          'Signal failure.
          OpenDneFile = False
    End Select
    
    Exit Function
    
Rollback:
    
    On Error Resume Next
    Set m_oStmtBodyFile = Nothing
    
    Return
    
End Function

'==============================================================================
'Function:              OpenSbfFile
'
'Synopsis:              Opens the Sbf file.  Returns true if opened successfully.
'
'==============================================================================
Private Function OpenSbfFile() As Boolean
    
    On Error GoTo EHandler
    
    Set m_oStmtBodyFile = New CFileOutput
    If m_oStmtBodyFile.OpenFile(m_sStmtBodyFileName) = False Then
        GoSub Rollback
        'Signal failure
        OpenSbfFile = False
        Exit Function
    End If
    
    With m_oStmtBodyFile
        .TrimTrailingBlanks = True
        .ExcludeFromTrim = 2
    End With
    
    OpenSbfFile = True
    
    Exit Function
    
EHandler:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
       Case 1 'User chose Retry
          On Error GoTo EHandler
          Resume
       Case 2 'User chose Ignore
          On Error GoTo EHandler
          Resume Next
       Case 3 'User chose Abort
          GoSub Rollback
          'Signal failure.
          OpenSbfFile = False
    End Select
    
    Exit Function
    
Rollback:
    
    On Error Resume Next
    Set m_oStmtBodyFile = Nothing
    
    Return
    
End Function

'==============================================================================
'Function:              OpenSbfDatabase
'
'Synopsis:              Opens (creates) the Sbf database.
'                       Returns true if opened successfully.
'
'==============================================================================
Private Function OpenSbfDatabase() As Boolean

   On Error GoTo EHandler

   Set m_oStmtBodyDatabase = New FfgSbfDb.CDbStmtBodyFile
   If m_oStmtBodyDatabase.CreateDatabase(m_sStmtBodyDatabaseName) = False Then
      GoSub Rollback
      'Signal failure
      OpenSbfDatabase = False
      Exit Function
   End If

   OpenSbfDatabase = True
   
   Exit Function

EHandler:

   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpContext = Err.HelpContext
   m_sHelpFile = Err.HelpFile

   Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
      Case 1 'User chose Retry
         Resume
      Case 2 'User chose Ignore
         Resume Next
      Case 3 'User chose Abort
         GoSub Rollback
         'Signal failure.
         OpenSbfDatabase = False
   End Select

   Exit Function

Rollback:

   On Error Resume Next
   Set m_oStmtBodyDatabase = Nothing

   Return

End Function

Private Function ReadPage() As Boolean
    On Error GoTo ReadPage_Error
    
    Dim oStack As CStack
    Dim sNote As String
    Dim bFileHeaderFound As Boolean
    
    Dim oRegex As RegExp
    Dim oMatchCollection As MatchCollection
    Dim oMatch As Match
    
    If Not m_bPrintStreamPrimed Then
        'Prime read.
        ReadPageAhead
    End If
    
    Do
    
        ReadPageAhead 'This transfers the read-ahead page to the current page and loads the read-ahead page.
        
        If m_lCurrentPageIdx < 0 Then
            'End of file
            ReadPage = False
            Exit Function
        End If
        
        Const scAS400 = "S400  "
        bFileHeaderFound = False
        
        If m_sXmlInst = "BPB" And Left(m_saCurrentPageLine(0), Len(m_scFileName)) = m_scFileName Then
            bFileHeaderFound = True
            
            If m_lCurrentFilePieces > 0 And m_bcVirtualFileSupport = False Then
                MessageBox "A second virtual file was detected in the input file " & GetFile(m_sInputFile(m_lInputFilePtr)) & ".  This virtual file needs to be separated into it's own file for processing.", vbOKOnly + vbCritical, App.Title
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            
            If m_lCurrentPageIdx > 0 Then
                'Shouldn't ever get here unless the code is broken.
                Stop
            End If
            
            If m_InputFileFileRecordCount > 0 Then
                If m_lCurrentFileExpectedPages >= 0 Then
                    ' Some eNotices requires 2 output pieces per input piece - this doubles the expected number of pages
                    ' to account for this
                    If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                        If m_lEddEnoticeType = EnoticeType10and20 Then
                            m_lCurrentFileExpectedPages = m_lCurrentFileExpectedPages * 2
                        End If
                    End If
                    If m_lCurrentFileImages <> m_lCurrentFileExpectedPages Then
                        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "File does not have the specified number of pages - " & m_sCurrentFileName & "'."
                    End If
                End If

                'Update input stack totals.
                Set oStack = New CStack
                If m_bCurrentFileIsUnrecognized Then
                    sNote = " (Unrecognized)"
                ElseIf m_lCurrentNoPrintStack > 0 Then
                    sNote = " (Suppressed)"
                Else
                    sNote = vbNullString
                End If
                oStack.SetProperties 10000 + m_lCurrentFileNumber, StackInput, RTrim(m_sCurrentFileName & " " & m_sCurrentFileFormType & " " & sNote), BalGroupInput, StackSimplex, False, False, 0
                m_oStmtBodyDatabase.InsertStack oStack
                Set oStack = Nothing
                
                'Insert the input stack totals into stack totals table.
                m_oStmtBodyDatabase.InsertInputTotals 10000 + m_lCurrentFileNumber, m_lCurrentFilePieces, m_lCurrentFileImages
            End If
            
            m_InputFileFileRecordCount = m_InputFileFileRecordCount + 1
            m_lCurrentFileNumber = m_lCurrentFileNumber + 1
            m_sCurrentFileName = RTrim(Mid(m_saCurrentPageLine(0), 11, 9))
            If Len(Trim(Mid(m_saCurrentPageLine(0), 31, 6))) = 0 Then
                m_lCurrentFileExpectedPages = -1
            Else
                m_lCurrentFileExpectedPages = CLng(Mid(m_saCurrentPageLine(0), 31, 6))
            End If
            m_bCurrentFileIsUnrecognized = False
            m_sCurrentFileFormType = RTrim(Mid(m_saCurrentPageLine(0), 50))
            m_lCurrentFilePieces = 0
            m_lCurrentFileImages = 0
            m_lCurrentNoPrintStack = 0
            m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
            
        ElseIf m_sXmlInst = "ElDorado" And m_lCurrentFilePieces = 0 Then
            ' ElDorado is InHouse Signature - they don't have any file records in their file.  Filename is used instead
            ' THe only time current file pieces will be 0 is when processing a new file from the beginning
            Dim fileNameExtractor As RegExp
            Dim fileNameMatch As MatchCollection
            Dim tempMatch As Match
            Set fileNameExtractor = New RegExp
            fileNameExtractor.Pattern = "^92351\.(.+)\.(\d{8})"
            
            If fileNameExtractor.Test(GetFile(m_sInputFile(m_lInputFilePtr))) = False Then
                Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Input file name is not valid for El Dorado: " & GetFile(m_sInputFile(m_lInputFilePtr))
                ReadPage = False
                Exit Function
            End If
            
            Set fileNameMatch = fileNameExtractor.Execute(GetFile(m_sInputFile(m_lInputFilePtr)))
            Set tempMatch = fileNameMatch.Item(0)
                  
            m_InputFileFileRecordCount = m_InputFileFileRecordCount + 1
            m_lCurrentFileNumber = m_lCurrentFileNumber + 1
            m_sCurrentFileName = tempMatch.SubMatches(0)
            m_lCurrentFileExpectedPages = -1
            m_bCurrentFileIsUnrecognized = False
            m_sCurrentFileFormType = "ELDORADO"
            m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
        ElseIf m_sXmlInst = m_scInst_29175100_FinFactors And m_lCurrentFilePieces = 0 Then
            ' Finance Factors is InHouse Signature - they don't have any file records in their file.  For now hard coding ln0480
            ' THe only time current file pieces will be 0 is when processing a new file from the beginning
        
            m_InputFileFileRecordCount = m_InputFileFileRecordCount + 1
            m_lCurrentFileNumber = m_lCurrentFileNumber + 1
            m_sCurrentFileName = "LN0480Q1"
            'Can't assume LN0480Q1 anymore.
            If InStr(1, GetFile(m_sInputFile(m_lInputFilePtr)), "TM0330P1", vbTextCompare) > 0 Then
                m_sCurrentFileName = "TM0330P1"
            End If
            m_lCurrentFileExpectedPages = -1
            m_bCurrentFileIsUnrecognized = False
            m_sCurrentFileFormType = "STMT"
            m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
            
        ElseIf m_sXmlInst = "55810000" Or m_sXmlInst = "75535000" Then
            Set oRegex = New RegExp
            oRegex.IgnoreCase = True
            oRegex.Pattern = "^FILE NAME-(.+)PAGES-(.+)FORM TYPE-(.*)$"
                
            If oRegex.Test(m_saCurrentPageLine(0)) = True Then
                Set oMatchCollection = oRegex.Execute(m_saCurrentPageLine(0))
                Set oMatch = oMatchCollection.Item(0)
                
                m_InputFileFileRecordCount = m_InputFileFileRecordCount + 1
                m_lCurrentFileNumber = m_lCurrentFileNumber + 1
                m_sCurrentFileName = RTrim(oMatch.SubMatches(0))
                If Len(RTrim(oMatch.SubMatches(1))) > 0 Then
                    m_lCurrentFileExpectedPages = CLng(oMatch.SubMatches(1))
                Else
                    m_lCurrentFileExpectedPages = -1
                End If
                m_bCurrentFileIsUnrecognized = False
                m_sCurrentFileFormType = RTrim(oMatch.SubMatches(2))
                m_lCurrentFilePieces = 0
                m_lCurrentFileImages = 0
                m_lCurrentNoPrintStack = 0
                m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
                bFileHeaderFound = True
            End If
        
        ElseIf Left(m_saCurrentPageLine(0), Len(scAS400)) = scAS400 Then
            bFileHeaderFound = True
            
            If m_lCurrentFilePieces > 0 And m_bcVirtualFileSupport = False Then
                MessageBox "A second virtual file was detected in the input file " & GetFile(m_sInputFile(m_lInputFilePtr)) & ".  This virtual file needs to be separated into it's own file for processing.", vbOKOnly + vbCritical, App.Title
                On Error GoTo 0
                Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
            End If
            
            If m_lCurrentPageIdx > 0 Then
                'Shouldn't ever get here unless the code is broken.
                Stop
            End If
            
            If m_InputFileFileRecordCount > 0 Then
                If m_lCurrentFileExpectedPages >= 0 Then
                    ' Some eNotices requires 2 output pieces per input piece - this doubles the expected number of pages
                    ' to account for this
                    If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                        If m_lEddEnoticeType = EnoticeType10and20 Then
                            m_lCurrentFileExpectedPages = m_lCurrentFileExpectedPages * 2
                        End If
                    End If
                    If m_lCurrentFileImages <> m_lCurrentFileExpectedPages Then
                        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "File does not have the specified number of pages - " & m_sCurrentFileName & "'."
                    End If
                End If
                    
                'Update input stack totals.
                Set oStack = New CStack
                If m_bCurrentFileIsUnrecognized Then
                    sNote = " (Unrecognized)"
                ElseIf m_lCurrentNoPrintStack > 0 Then
                    sNote = " (Suppressed)"
                Else
                    sNote = vbNullString
                End If
                oStack.SetProperties 10000 + m_lCurrentFileNumber, StackInput, RTrim(m_sCurrentFileName & " " & m_sCurrentFileFormType & " " & sNote), BalGroupInput, StackSimplex, False, False, 0
                m_oStmtBodyDatabase.InsertStack oStack
                Set oStack = Nothing
                
                'Insert the input stack totals into stack totals table.
                m_oStmtBodyDatabase.InsertInputTotals 10000 + m_lCurrentFileNumber, m_lCurrentFilePieces, m_lCurrentFileImages
            End If

            Set oRegex = New RegExp
            oRegex.IgnoreCase = True
            oRegex.Pattern = "^S400 {2}CB9410 {6}FILE NAME- ([\w ]{10})([\w\* ]{10}).+?PAGES-([ 0-9]{10})$"
            
            If oRegex.Test(m_saCurrentPageLine(0)) = False Then
                Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid file header detected - file header must match the following regular expression: ^AS400 +CB9410 +FILE NAME- (?'FileName'[\w ]{10})(?'FormType'[\w\* ]{10}).+?PAGES-(?'Pages'[ 0-9]{10})$"
                ReadPage = False
                Exit Function
            End If
            
            Set oMatchCollection = oRegex.Execute(m_saCurrentPageLine(0))
            Set oMatch = oMatchCollection.Item(0)
            
            m_InputFileFileRecordCount = m_InputFileFileRecordCount + 1
            m_lCurrentFileNumber = m_lCurrentFileNumber + 1
            m_sCurrentFileName = RTrim(oMatch.SubMatches(0))
            m_lCurrentFileExpectedPages = CLng(oMatch.SubMatches(2))
            m_bCurrentFileIsUnrecognized = False
            m_sCurrentFileFormType = RTrim(oMatch.SubMatches(1))
            m_lCurrentFilePieces = 0
            m_lCurrentFileImages = 0
            m_lCurrentNoPrintStack = 0
            m_lCurrentBulkShipStack = GetBulkShipIndex(m_sCurrentFileName, m_sCurrentFileFormType)
        End If
    
    Loop Until (m_lCurrentPageIdx >= 0) And (bFileHeaderFound = False) 'Ignore blank pages and keep looping after file name records.
    
    ReadPage = True
    
    Exit Function
    
ReadPage_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReadPage"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReadPage_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReadPage_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

'==============================================================================
'Function:              ParseXMLAddressPosition
'
'Synopsis:              Parses XML options from the tag property.  Returns
'                       true if tag is a valid XML document for this DLL.
'==============================================================================
Private Function ParseXMLAddressPosition(ByVal sTag As String, ByRef lAddressPositionX As Long, ByRef lAddressPositionY As Long) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLAddressPosition_Error
    
    'Defaults
    ParseXMLAddressPosition = False
'    lAddressPositionX = 63
'    lAddressPositionY = 621
    lAddressPositionX = 72
    lAddressPositionY = 567
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLAddressPosition_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLAddressPosition_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If oRootNode.nodeName <> "Options" Then
        MessageBox "Error loading the XML options string from the Tag property.  The m_oChecksRoot node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case oNode.nodeName
            Case "AddressPosition"
                Set oAttribute = oNode.Attributes.getNamedItem("x")
                If Not oAttribute Is Nothing Then
                    lAddressPositionX = CLng(oAttribute.Text)
                End If
                Set oAttribute = oNode.Attributes.getNamedItem("y")
                If Not oAttribute Is Nothing Then
                    lAddressPositionY = CLng(oAttribute.Text)
                End If
        End Select
    Next
    
    ParseXMLAddressPosition = True
    Exit Function

ParseXMLAddressPosition_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLAddressPosition = False
    Exit Function

ParseXMLAddressPosition_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

'==============================================================================
'Function:              ParseXMLFileHeader
'
'Synopsis:              Parses XML options from the tag property.  Returns
'                       true if tag is a valid XML document for this DLL.
'==============================================================================
Private Function ParseXMLFileHeader(ByVal sTag As String, ByRef sReportId As String, ByRef sFormType As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLFileHeader_Error
    
    'Defaults
    ParseXMLFileHeader = False
    sReportId = vbNullString
    sFormType = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLFileHeader_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLFileHeader_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If oRootNode.nodeName <> "Options" Then
        MessageBox "Error loading the XML options string from the Tag property.  The m_oChecksRoot node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Dim oFileHeaderNode As MSXML2.IXMLDOMNode
    Set oFileHeaderNode = oRootNode.selectSingleNode("FileHeader/ReportId")
    If Not oFileHeaderNode Is Nothing Then
        sReportId = oFileHeaderNode.Text
        Set oFileHeaderNode = oRootNode.selectSingleNode("FileHeader/FormType")
        If Not oFileHeaderNode Is Nothing Then
            sFormType = oFileHeaderNode.Text
        End If
    End If
    
    ParseXMLFileHeader = True
    Exit Function

ParseXMLFileHeader_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLFileHeader = False
    Exit Function

ParseXMLFileHeader_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

'==============================================================================
'Function:              ParseXMLDocument
'
'Synopsis:              Parses XML options from the tag property.  Returns
'                       true if tag is a valid XML document for this DLL.
'==============================================================================
Private Function ParseXMLDocument(ByVal sTag As String, ByRef sDocument As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLDocument_Error
    
    'Defaults
    ParseXMLDocument = False
    sDocument = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLDocument_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLDocument_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If oRootNode.nodeName <> "Options" Then
        MessageBox "Error loading the XML options string from the Tag property.  The m_oChecksRoot node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Dim oFileHeaderNode As MSXML2.IXMLDOMNode
    Set oFileHeaderNode = oRootNode.selectSingleNode("Document")
    If Not oFileHeaderNode Is Nothing Then
        sDocument = oFileHeaderNode.Text
    End If
    
    ParseXMLDocument = True
    Exit Function

ParseXMLDocument_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLDocument = False
    Exit Function

ParseXMLDocument_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Sub WriteEddJobSettings(ByRef oDom As MSXML2.DOMDocument60, ByVal sJobKey As String, ByVal dtCycleDate As Date, ByVal sClientId As String, ByVal sClientProductId As String, ByVal bCompanionFile As Boolean, ByVal sDocType As String)
    Dim oNode As MSXML2.IXMLDOMNode
    
    If sDocType = vbNullString Then
        sDocType = "NOT"
    End If
    
    Dim oRoot As MSXML2.IXMLDOMElement
    If oDom Is Nothing Then
        Set oDom = New MSXML2.DOMDocument60
        oDom.validateOnParse = False
        oDom.resolveExternals = False
        oDom.preserveWhiteSpace = True
    
        ' Create processing instruction and document m_oChecksRoot
        Set oNode = oDom.createProcessingInstruction("xml", "version='1.0'")
        oDom.appendChild oNode
        Set oNode = Nothing
        
        ' Create a comment for the document.
        Set oNode = oDom.createComment("FDSPrint Wizard Job Finish Directives.")
        'oDom.appendChild oDom.createTextNode(vbNewLine)
        oDom.appendChild oNode
        Set oNode = Nothing
        
        ' Create the root element
        Set oRoot = oDom.createElement("FinishDirectives")
        
        ' Add the root element to the DOM instance.
        'oDom.appendChild oDom.createTextNode(vbNewLine)
        oDom.appendChild oRoot
    Else
        Set oRoot = oDom.documentElement
    End If
    
    Dim oEddJobsNode As MSXML2.IXMLDOMElement
    
    Set oEddJobsNode = oRoot.selectSingleNode("//FinishDirectives/EddJobs")
    If oEddJobsNode Is Nothing Then
        Set oEddJobsNode = oDom.createElement("EddJobs")
        'oRoot.appendChild oDom.createTextNode(vbNewLine + vbTab)
        oRoot.appendChild oEddJobsNode
    End If
    
    Dim sJobName As String
    sJobName = sJobKey & "." & Format(dtCycleDate, "yyyyMMdd")
    
    Dim oEddJobNode As MSXML2.IXMLDOMElement
    Set oEddJobNode = oDom.createElement("EddJob")
    'oEddJobsNode.appendChild oDom.createTextNode(vbNewLine + vbTab + vbTab)
    oEddJobsNode.appendChild oEddJobNode
    'Dim oAttr As MSXML2.IXMLDOMAttribute
    oEddJobNode.setAttribute "key", sJobKey
    oEddJobNode.setAttribute "dne", LCase(CStr(True))
    oEddJobNode.setAttribute "companionFile", LCase(CStr(bCompanionFile))
    
     ' DDA 2.0 upgrade
        If m_lXML_EddPlatformType = DDA_20 Then
            oEddJobNode.setAttribute "platform", CStr(m_lXML_EddPlatformType)
        End If
    
    Set oNode = oDom.createElement("JobName")
    'oEddJobsNode.appendChild oDom.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
    oEddJobNode.appendChild oNode
    oNode.Text = sJobName
    Set oNode = Nothing
    
    Set oNode = oDom.createElement("ClientId")
    'oEddJobsNode.appendChild oDom.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
    oEddJobNode.appendChild oNode
    oNode.Text = sClientId
    Set oNode = Nothing
    
    Set oNode = oDom.createElement("ClientProductId")
    'oEddJobsNode.appendChild oDom.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
    oEddJobNode.appendChild oNode
    oNode.Text = sClientProductId
    Set oNode = Nothing
    
    Set oNode = oDom.createElement("PDF")
    'oEddJobsNode.appendChild oDom.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
    oEddJobNode.appendChild oNode
    Set oNode = Nothing
    
     ' DDA 2.0 upgrade
        If m_lXML_EddPlatformType = DDA_20 Then
            Set oNode = oDom.createElement("DocumentType")
            oEddJobNode.appendChild oNode
            oNode.Text = sDocType
            
            Set oNode = Nothing
            Set oNode = oDom.createElement("DataDate")
            oEddJobNode.appendChild oNode
            oNode.Text = Format(dtCycleDate, "MM/dd/yyyy")
            Set oNode = Nothing
        End If
    
End Sub

'Depracated - Use ParseXMLEddSettings
Private Function ParseXMLEdd(ByVal sTag As String, ByRef bEstatements As Boolean) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAddressNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    Dim lAddressLine As Long
    Dim bInstructions As Boolean
    
    On Error GoTo ParseXMLEdd_Error
    
    'Defaults
    ParseXMLEdd = False
    bEstatements = False
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLEdd_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLEdd_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case UCase(oNode.nodeName)
            Case "ESTATEMENTS", "EDD"
                Select Case UCase(Trim(oNode.Text))
                    Case "NO"
                        bEstatements = False
                    Case "FALSE"
                        bEstatements = False
                    Case "OFF"
                        bEstatements = False
                    Case "YES"
                        bEstatements = True
                    Case "TRUE"
                        bEstatements = True
                    Case "ON"
                        bEstatements = True
                End Select
       End Select
    Next
    
    ParseXMLEdd = True
    Exit Function

ParseXMLEdd_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLEdd = False
    Exit Function

ParseXMLEdd_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Function ParseXMLEddSettings(ByVal sTag As String, ByRef sEddJobName As String, ByRef sEddProductId As String, ByRef sEddDocType As String, ByRef sEddClientId As String, ByRef bEddCompanionFile As Boolean) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAddressNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    Dim lAddressLine As Long
    Dim bInstructions As Boolean
    
    On Error GoTo ParseXMLEddSettings_Error
    
    'Defaults
    sEddDocType = vbNullString
    ParseXMLEddSettings = False
    sEddJobName = vbNullString
    sEddProductId = vbNullString
    sEddClientId = vbNullString
    bEddCompanionFile = False
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLEddSettings_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLEddSettings_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Dim oEddNode As MSXML2.IXMLDOMNode
    Set oEddNode = oRootNode.selectSingleNode("EDD")
    If Not oEddNode Is Nothing Then
        Set oAttribute = oEddNode.Attributes.getNamedItem("jobKey")
        If Not oAttribute Is Nothing Then
            sEddJobName = Trim(oAttribute.Text)
        End If
        If Len(sEddJobName) > 0 Then
            Set oAttribute = oEddNode.Attributes.getNamedItem("productID")
            If Not oAttribute Is Nothing Then
                sEddProductId = Trim(oAttribute.Text)
            End If
            Set oAttribute = oEddNode.Attributes.getNamedItem("eddDocType")
            If Not oAttribute Is Nothing Then
                sEddDocType = Trim(oAttribute.Text)
            End If
            Set oAttribute = oEddNode.Attributes.getNamedItem("clientID")
            If Not oAttribute Is Nothing Then
                sEddClientId = Trim(oAttribute.Text)
            End If
            bEddCompanionFile = True ' Default is true
            Set oAttribute = oEddNode.Attributes.getNamedItem("companionFile")
            If Not oAttribute Is Nothing Then
                bEddCompanionFile = CBool(oAttribute.Text)
            End If
        End If
    End If
    
    ParseXMLEddSettings = True
    Exit Function

ParseXMLEddSettings_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLEddSettings = False
    Exit Function

ParseXMLEddSettings_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Function ParseXMLEddEnoticeSettings(ByVal sTag As String, ByRef sEddJobName As String, ByRef sEddClientId As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAddressNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    Dim lAddressLine As Long
    Dim bInstructions As Boolean
    
    On Error GoTo ParseXMLEddEnoticeSettings_Error
    
    'Defaults
    ParseXMLEddEnoticeSettings = False
    sEddJobName = vbNullString
    sEddClientId = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLEddEnoticeSettings_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLEddEnoticeSettings_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
        
    For Each oNode In oRootNode.childNodes
        Select Case UCase(oNode.nodeName)
            Case "ENOTICE"
                Set oAttribute = oNode.Attributes.getNamedItem("jobKey")
                If Not oAttribute Is Nothing Then
                    sEddJobName = Trim(oAttribute.Text)
                End If
                If Len(sEddJobName) > 0 Then
                    Set oAttribute = oNode.Attributes.getNamedItem("clientID")
                    If Not oAttribute Is Nothing Then
                        sEddClientId = Trim(oAttribute.Text)
                    End If
                End If
        End Select
    Next
    
    ParseXMLEddEnoticeSettings = True
    Exit Function

ParseXMLEddEnoticeSettings_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLEddEnoticeSettings = False
    Exit Function

ParseXMLEddEnoticeSettings_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Function ParseXMLAccountType(ByVal sTag As String, ByRef sAccountType As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLAccountType_Error
    
    'Defaults
    ParseXMLAccountType = False
    sAccountType = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLAccountType_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLAccountType_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = oRootNode.selectSingleNode("AccountType")
    If Not oNode Is Nothing Then
        sAccountType = Trim(oNode.Text)
    End If
    
    ParseXMLAccountType = True
    Exit Function

ParseXMLAccountType_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLAccountType = False
    Exit Function

ParseXMLAccountType_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function
' DDA 2.0 Upgrade
Private Function ParseXMLEddPlatform(ByVal sTag As String, ByRef lEddPlatformType As EddPlatformType) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAddressNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String

    On Error GoTo ParseXMLEddPlatform_Error
    
    'Defaults
    ParseXMLEddPlatform = False
    lEddPlatformType = EddPlatformType.Legacy
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLEddPlatform_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLEddPlatform_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MsgBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MsgBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Dim oEddNode As MSXML2.IXMLDOMNode
    Set oEddNode = oRootNode.selectSingleNode("EddPlatform")
    If Not oEddNode Is Nothing Then
        Select Case Trim(UCase(oEddNode.Text))
            Case "DDA20"
                lEddPlatformType = DDA_20
            Case "LEGACY"
                lEddPlatformType = Legacy
            Case Else
                Err.Raise 1, App.Title, "Invalid EddPlatform value provided: " & oEddNode.Text
        End Select
    End If
    
    ParseXMLEddPlatform = True
    Exit Function

ParseXMLEddPlatform_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MsgBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLEddPlatform = False
    Exit Function

ParseXMLEddPlatform_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MsgBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

'==============================================================================
'Function:              ParseXMLArchive
'
'Synopsis:              Parses XML options from the tag property.  Returns
'                       true if tag is a valid XML document for this DLL.
'
'Remarks:               Sample XML:
'                       <Options>
'                           <IRSForm>1098</IRSForm>
'                           <ReturnAddress>No</ReturnAddress>
'                           <XeroxPDE>1098</XeroxPDE>
'                           <XeroxForm>NONE</XeroxForm>
'                       </Options>
'==============================================================================
Private Function ParseXMLArchive(ByVal sTag As String, ByRef bArchive As Boolean, ByRef bArchiveNameInFile As Boolean) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLArchive_Error
    
    'Defaults
    ParseXMLArchive = False
    bArchive = False
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLArchive_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLArchive_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case UCase(oNode.nodeName)
            Case "ARCHIVE"
                Select Case UCase(Trim(oNode.Text))
                    Case "NO"
                        bArchive = False
                    Case "FALSE"
                        bArchive = False
                    Case "OFF"
                        bArchive = False
                    Case "YES"
                        bArchive = True
                    Case "TRUE"
                        bArchive = True
                    Case "ON"
                        bArchive = True
                End Select
                
                If bArchive Then
                    Set oAttribute = oNode.Attributes.getNamedItem("ArchiveNameInFile")
                    If Not oAttribute Is Nothing Then
                        Select Case UCase(Trim(oAttribute.Text))
                            Case "NO"
                                bArchiveNameInFile = False
                            Case "FALSE"
                                bArchiveNameInFile = False
                            Case "OFF"
                                bArchiveNameInFile = False
                            Case "YES"
                                bArchiveNameInFile = True
                            Case "TRUE"
                                bArchiveNameInFile = True
                            Case "ON"
                                bArchiveNameInFile = True
                        End Select
                    End If
                End If
       End Select
    Next
    
    ParseXMLArchive = True
    Exit Function

ParseXMLArchive_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLArchive = False
    Exit Function

ParseXMLArchive_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Function ParseXMLAse(ByVal sTag As String, ByRef sAse As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAddressNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    On Error GoTo ParseXMLAse_Error
    
    'Defaults
    ParseXMLAse = False
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXMLAse_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXMLAse_Error
    
    If Not bSuccess Then
        sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
        sMsg = sMsg & oDOMDocument.parseError.reason & vbCrLf & _
            " Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
            oDOMDocument.parseError.srcText
        MessageBox sMsg, vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case UCase(oNode.nodeName)
            Case "ASE"
                sAse = oNode.Text
       End Select
    Next
    
    ParseXMLAse = True
    Exit Function

ParseXMLAse_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXMLAse = False
    Exit Function

ParseXMLAse_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function


Private Function ParseXmlInst(ByVal sTag As String, ByRef sInst As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    Dim oListNode As MSXML2.IXMLDOMNode
    
    On Error GoTo ParseXmlInst_Error
    
    'Defaults
    ParseXmlInst = False
    sInst = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXmlInst_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXmlInst_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case oNode.nodeName
            Case "Inst"
                sInst = Trim(oNode.Text)
        End Select
    Next
    
    ParseXmlInst = True
    Exit Function

ParseXmlInst_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXmlInst = False
    Exit Function

ParseXmlInst_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function
Private Function ParseXmlBulkShipList(ByVal sTag As String, ByRef sBulkList As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    Dim oListNode As MSXML2.IXMLDOMNode
    
    On Error GoTo ParseXmlBulkShipList_Error
    
    'Defaults
    ParseXmlBulkShipList = False
    sBulkList = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXmlBulkShipList_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXmlBulkShipList_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case UCase(oNode.nodeName)
            Case "BULKSHIPLIST"
                sBulkList = Trim(oNode.Text)
        End Select
    Next
    
    ParseXmlBulkShipList = True
    Exit Function

ParseXmlBulkShipList_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXmlBulkShipList = False
    Exit Function

ParseXmlBulkShipList_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function
Private Function ParseXmlFileGroup(ByVal sTag As String, ByRef sFileGroup As String) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    Dim oListNode As MSXML2.IXMLDOMNode
    
    On Error GoTo ParseXmlFileGroup_Error
    
    'Defaults
    ParseXmlFileGroup = False
    sFileGroup = vbNullString
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXmlFileGroup_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXmlFileGroup_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case oNode.nodeName
            Case "FileGroup"
                sFileGroup = UCase(Trim(oNode.Text))
        End Select
    Next
    
    ParseXmlFileGroup = True
    Exit Function

ParseXmlFileGroup_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXmlFileGroup = False
    Exit Function

ParseXmlFileGroup_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function

Private Function ParseXmlReturnToBank(ByVal sTag As String, ByRef bReturnToBank As Boolean) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    Dim oListNode As MSXML2.IXMLDOMNode
    
    On Error GoTo ParseXmlReturnToBank_Error
    
    'Defaults
    ParseXmlReturnToBank = False
    bReturnToBank = False
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXmlReturnToBank_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXmlReturnToBank_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        Select Case oNode.nodeName
            Case "ReturnToBank"
                bReturnToBank = CBool(Trim(oNode.Text))
        End Select
    Next
    
    ParseXmlReturnToBank = True
    Exit Function

ParseXmlReturnToBank_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXmlReturnToBank = False
    Exit Function

ParseXmlReturnToBank_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function



Private Function ParseXmlForm(ByVal sTag As String, ByVal sFormElementName, ByRef sImageFile As String, ByRef sDictionaryName As String, ByRef oViewport As PostscriptImaging.Viewport, Optional ByVal lRectangleCorner As PostscriptImaging.RectangleCorner = RectangleCorner_LowerLeft) As Boolean
    Dim oDOMDocument As MSXML2.DOMDocument60
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttribute As IXMLDOMNode
    Dim bSuccess As Boolean
    Dim sMsg As String
    
    Dim lX As Long
    Dim lY As Long
    Dim lW As Long
    Dim lH As Long
    Dim bCenter As Boolean
    Dim sValue As String
    
    On Error GoTo ParseXmlForm_Error
    
    'Defaults
    ParseXmlForm = False
    sImageFile = vbNullString
    sDictionaryName = vbNullString
    Set oViewport = Nothing
    
    sFormElementName = UCase(sFormElementName)
    
    Set oDOMDocument = New MSXML2.DOMDocument60
    
    On Error GoTo ParseXmlForm_XMLError
    oDOMDocument.async = False
    oDOMDocument.validateOnParse = True
    bSuccess = oDOMDocument.loadXML(sTag)
    On Error GoTo ParseXmlForm_Error
    
    If Not bSuccess Then
        Exit Function
    End If
    
    Set oRootNode = oDOMDocument.documentElement
    If UCase(oRootNode.nodeName) <> "OPTIONS" Then
        MessageBox "Error loading the XML options string from the Tag property.  The root node is not named 'Options'", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    For Each oNode In oRootNode.childNodes
        If UCase(oNode.nodeName) = sFormElementName Then
            sImageFile = Trim(oNode.Text)
            Set oViewport = New Viewport
            Set oAttribute = oNode.Attributes.getNamedItem("Name")
            If Not oAttribute Is Nothing Then
                sDictionaryName = oAttribute.Text
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("RectangleCorner")
            If Not oAttribute Is Nothing Then
                sValue = oAttribute.Text
                Select Case sValue
                    Case "LowerLeft"
                        lRectangleCorner = RectangleCorner_LowerLeft
                    Case "UpperLeft"
                        lRectangleCorner = RectangleCorner_UpperLeft
                    Case "UpperRight"
                        lRectangleCorner = RectangleCorner_UpperLeft
                    Case "LowerRight"
                        lRectangleCorner = RectangleCorner_LowerRight
                    Case Else
                        Err.Raise g_lcErrBase, App.Title, "Invalid value for RectangleCorner in job settings."
                End Select
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("x")
            If Not oAttribute Is Nothing Then
                lX = CLng(oAttribute.Text)
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("y")
            If Not oAttribute Is Nothing Then
                lY = CLng(oAttribute.Text)
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("w")
            If Not oAttribute Is Nothing Then
                lW = CLng(oAttribute.Text)
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("h")
            If Not oAttribute Is Nothing Then
                lH = CLng(oAttribute.Text)
            End If
            Set oAttribute = oNode.Attributes.getNamedItem("Center")
            If Not oAttribute Is Nothing Then
                bCenter = CBool(oAttribute.Text)
            End If
        End If
    Next
    
    If Not oViewport Is Nothing Then
        With oViewport
            .ReferenceCorner.Corner = lRectangleCorner
            .ReferenceCorner.Location.X = lX
            .ReferenceCorner.Location.Y = lY
            .Width = lW
            .Height = lH
            .CenterImage = bCenter
        End With
    End If
    
    ParseXmlForm = True
    Exit Function

ParseXmlForm_XMLError:
    
    'If an error occurs, display the Source of the error, the line number
    'in the XML Document that the error occurred on and the text of the error:
    sMsg = oDOMDocument.parseError.reason & vbCrLf & _
        "Line Number In Document: " & oDOMDocument.parseError.Line & vbCrLf & _
        oDOMDocument.parseError.srcText
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    ParseXmlForm = False
    Exit Function

ParseXmlForm_Error:

    sMsg = "Error loading the XML options string from the Tag property." & vbCrLf & vbCrLf
    sMsg = sMsg & Err.Description
    MessageBox sMsg, vbCritical + vbOKOnly, App.Title
    Exit Function

End Function


Private Sub PsPageBegin_FlatStock(ByVal sDictionary As String, Optional ByVal bDuplex As Boolean)
    Dim bFront As Boolean
    bFront = False
    Dim sMediaType As String
    sMediaType = PsString(m_sStockTray, True)
    m_lOutputImages = m_lOutputImages + 1
    m_lOutputImageSeq = m_lOutputImageSeq + 1
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then  'Every odd page
           m_lOutputSheets = m_lOutputSheets + 1
           bFront = True
        End If
    Else
       m_lOutputSheets = m_lOutputSheets + 1
       bFront = True
    End If
    With m_oStmtBodyFile
        .WriteLine "%%Page: " & CStr(m_lOutputImages) & " " & CStr(m_lOutputImageSeq)
        Dim i As Long
        If m_lOutputImages = 1 Then
            For i = 0 To m_lIndexIdx
                .WriteLine "%FiservPieceIndex: " & m_saIndexKeys(i) & " " & m_saIndexValues(i)
            Next
            If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                .WriteLine "%FiservPieceIndex: EDDDocumentType " & m_sPieceEddEnoticeType & "-" & TrimLeadingZeros(m_sAccountNumber)
            End If
            ' RM15930
            For i = 0 To m_lSecureDestructionIndexIdx
                If i = 0 Then
                    .WriteLine "%FiservSecureDestructionIndex: PieceID " & m_lLocation
                End If
                .WriteLine "%FiservSecureDestructionIndex: " & m_saSecureDestructionIndexKeys(i) & " " & m_saSecureDestructionIndexValues(i)
            Next
            ' End RM15930
        End If
        .WriteLine "%FiservVerifySecurity: " & m_sAccountNumber
        .WriteLine "%%PageBoundingBox: 1 1 612 792"
        .WriteLine "%%BeginPageSetup"
        .WriteLine "<< /Duplex " & LCase(CStr(bDuplex)) & "  /PageSize [612 792] /MediaColor null /MediaType " & sMediaType & " >> setpagedevice    "
        .WriteLine "/pgsave save def"
        .WriteLine "%%EndPageSetup"
        .WriteLine "userdict /" & sDictionary & " get begin"
        If Not (m_bXML_Archive Or m_bXML_EddEstatements Or Len(m_sXML_EddEnoticeJobKey) > 0) Then
            If bFront Then
                .WriteLine PDIOMR(m_lOutputSheets, m_sOMR)
            End If
        End If
    End With
End Sub

Private Sub PsPageBegin_FlatStockLegal(ByVal sDictionary As String, Optional ByVal bDuplex As Boolean)
    Dim bFront As Boolean
    bFront = False
    Dim sMediaType As String
    sMediaType = PsString(m_sStockTray, True)
    m_lOutputImages = m_lOutputImages + 1
    m_lOutputImageSeq = m_lOutputImageSeq + 1
    If bDuplex Then
        If m_lOutputImages Mod 2 > 0 Then  'Every odd page
           m_lOutputSheets = m_lOutputSheets + 1
           bFront = True
        End If
    Else
       m_lOutputSheets = m_lOutputSheets + 1
       bFront = True
    End If
    With m_oStmtBodyFile
        .WriteLine "%%Page: " & CStr(m_lOutputImages) & " " & CStr(m_lOutputImageSeq)
        Dim i As Long
        If m_lOutputImages = 1 Then
            For i = 0 To m_lIndexIdx
                .WriteLine "%FiservPieceIndex: " & m_saIndexKeys(i) & " " & m_saIndexValues(i)
            Next
            If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                .WriteLine "%FiservPieceIndex: EDDDocumentType " & m_sPieceEddEnoticeType & "-" & TrimLeadingZeros(m_sAccountNumber)
            End If
            ' RM15930
            For i = 0 To m_lSecureDestructionIndexIdx
                If i = 0 Then
                    .WriteLine "%FiservSecureDestructionIndex: PieceID " & m_lLocation
                End If
                .WriteLine "%FiservSecureDestructionIndex: " & m_saSecureDestructionIndexKeys(i) & " " & m_saSecureDestructionIndexValues(i)
            Next
            ' END RM15930
        End If
        .WriteLine "%FiservVerifySecurity: " & m_sAccountNumber
        .WriteLine "%%PageBoundingBox: 1 1 612 1008"
        .WriteLine "%%BeginPageSetup"
        .WriteLine "<< /Duplex false /PageSize [612 1008] /MediaColor null /MediaType " & sMediaType & " >> setpagedevice    "
        .WriteLine "/pgsave save def"
        .WriteLine "%%EndPageSetup"
        .WriteLine "userdict /" & sDictionary & " get begin"
        If Not (m_bXML_Archive Or m_bXML_EddEstatements Or Len(m_sXML_EddEnoticeJobKey) > 0) Then
            If bFront Then
                .WriteLine "gs 0 216 translate"
                .WriteLine PDIOMR(m_lOutputSheets, m_sOMR)
                .WriteLine "gr"
            End If
        End If
    End With
End Sub

Private Sub PsPageBegin(ByVal sDictionary As String)
    Dim sMediaType As String
    Dim bFront As Boolean
    sMediaType = PsString(m_sStockTray, True)
    m_lOutputImages = m_lOutputImages + 1
    'If m_lOutputImages Mod 2 > 0 Then  'Every odd page
       m_lOutputSheets = m_lOutputSheets + 1
       bFront = True
    'End If
    'm_lOutputClicks = m_lOutputClicks + 1
    m_lOutputImageSeq = m_lOutputImageSeq + 1
    With m_oStmtBodyFile
        .WriteLine "%%Page: " & CStr(m_lOutputImages) & " " & CStr(m_lOutputImageSeq)
        Dim i As Long
        If m_lOutputImages = 1 Then
            For i = 0 To m_lIndexIdx
                .WriteLine "%FiservPieceIndex: " & m_saIndexKeys(i) & " " & m_saIndexValues(i)
            Next
            If Len(m_sXML_EddEnoticeJobKey) > 0 Then
                .WriteLine "%FiservPieceIndex: EDDDocumentType " & m_sPieceEddEnoticeType & "-" & TrimLeadingZeros(m_sAccountNumber)
            End If
            ' RM15930
            For i = 0 To m_lSecureDestructionIndexIdx
                If i = 0 Then
                    .WriteLine "%FiservSecureDestructionIndex: PieceID " & m_lLocation
                End If
                .WriteLine "%FiservSecureDestructionIndex: " & m_saSecureDestructionIndexKeys(i) & " " & m_saSecureDestructionIndexValues(i)
            Next
            ' END RM15930
        End If
        .WriteLine "%FiservVerifySecurity: " & m_sAccountNumber
        .WriteLine "%%PageBoundingBox: 1 1 612 792"
        .WriteLine "%%BeginPageSetup"
        .WriteLine "<< /Duplex false /PageSize [612 792] /MediaColor null /MediaType " & sMediaType & " >> setpagedevice    "
        .WriteLine "/pgsave save def"
        .WriteLine "%%EndPageSetup"
        .WriteLine "userdict /" & sDictionary & " get begin"
        If Not (m_lStackItems = StackItemsPressureSeal Or m_bXML_Archive Or m_bXML_EddEstatements Or Len(m_sXML_EddEnoticeJobKey) > 0) Then
            If bFront Then
                .WriteLine PDIOMR(m_lOutputSheets, m_sOMR)
            End If
        End If

    End With
End Sub

Private Sub PsPageEnd()
    With m_oStmtBodyFile
        .WriteLine "end % current notice procedure dictionary"
        .WriteLine "pgsave restore"
        .WriteLine "showpage"
    End With
End Sub

Private Sub PsWriteHeader()
    With m_oStmtBodyFile
        .WriteLine "%!PS-Adobe-3.0"
        .WriteLine "%%Title: " & App.Title
        .WriteLine "%%Creator: " & App.EXEName
        .WriteLine "%%CreationDate: " & Format(Now, "h:m m/d/yyyy")
        .WriteLine "%%Pages: (atend)"
        '.WriteLine "%%PageOrder: Special"
        .WriteLine "%%BoundingBox: 1 1 612 792"
        .WriteLine "%%LanguageLevel: 2"
        .WriteLine "%%DocumentNeededFonts: (atend)"
        .WriteLine "%%DocumentSuppliedFonts: (atend)"
        If m_bXML_ArchiveNameInFile Then
            .WriteLine "%FdsDigArOption: OutputNameInPostScript True"
        End If
        If m_sXmlInst = m_scInst_29175100_FinFactors Then
            .WriteLine "%FdsDigArOption: Delimiter |" ' ddh - incase the spelling is fixed.
            .WriteLine "%FdsDigArOption: Delimeter |" ' ddh - because the spelchekur didn't catch it.
            .WriteLine "%FdsDigArOption: NoQuote true"
            .WriteLine "%FdsDigArOption: GuidFileName False" ' This doesn't work.  GuidFileName isn't anywhere in PSL
            .WriteLine "%FdsDigArOption: OutputNameInPostScript True"
        Else
            .WriteLine "%FdsDigArOption: GuidFileName True" ' This doesn't work.  GuidFileName isn't anywhere in PSL
        End If
        .WriteLine "%%EndComments"
    End With
End Sub

Private Sub PsWriteSetup()
    Dim sImageFile As String
    Dim sDictionaryName As String
    Dim oViewport As PostscriptImaging.Viewport
    Dim oForm As New PostscriptImaging.Form
    oForm.ImageRootDirectory = m_sDocumentFolder
    Dim oFormReader As FormReader
    Dim sCodeLine As String
    
    With m_oStmtBodyFile
        .WriteLine "%%BeginSetup"
        .WriteLine "% set MaxFormItem to be equivalent to MaxFormCache"
        .WriteLine "<< /MaxFormItem currentsystemparams /MaxFormCache get >> setuserparams"
        .WriteLine ""
        .WriteLine "FDSFormsOps begin"
        .WriteLine "FDSGlobal begin"
        '.WriteLine "FDSLocal begin"
        'Note: Each print routine responsible for pushing and popping its routine-specific dictionary on and off the dictionary stack.
        
        'Indicia
        Set oViewport = New Viewport
        With oViewport
            .ReferenceCorner.Corner = RectangleCorner_UpperRight
            .ReferenceCorner.Location.X = 558
            .ReferenceCorner.Location.Y = m_lcPressureSealOuterLogoY
        End With
        
        Set oFormReader = oForm.RenderToCodeReader(m_scIndiciaFile, "Indicia", oViewport)
        Set m_oIndicia = oFormReader.FormInfo
        Do Until oFormReader.EndOfForm
            sCodeLine = oFormReader.ReadLine()
            .WriteLine sCodeLine
        Loop
        
        ' Load XML settings here
        If ParseXmlForm(m_sTag, "FLATSTOCKLOGO", sImageFile, sDictionaryName, oViewport) = False Then
            Err.Raise g_lcErrBase, App.Title, "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid."
        End If
        If Not oViewport Is Nothing Then
            Set oFormReader = oForm.RenderToCodeReader(sImageFile, sDictionaryName, oViewport)
            Set m_oXml_FlatStockLogo = oFormReader.FormInfo
            Do Until oFormReader.EndOfForm
                sCodeLine = oFormReader.ReadLine()
                .WriteLine sCodeLine
            Loop
        End If
        
        If ParseXmlForm(m_sTag, "OUTERLOGO", sImageFile, sDictionaryName, oViewport) = False Then
            Err.Raise g_lcErrBase, App.Title, "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid."
        End If
        If Not oViewport Is Nothing Then
            Set oFormReader = oForm.RenderToCodeReader(sImageFile, sDictionaryName, oViewport)
            Set m_oXml_OuterLogo = oFormReader.FormInfo
            Do Until oFormReader.EndOfForm
                sCodeLine = oFormReader.ReadLine()
                .WriteLine sCodeLine
            Loop
        End If
        
        If ParseXmlForm(m_sTag, "INNERLOGO", sImageFile, sDictionaryName, oViewport) = False Then
            Err.Raise g_lcErrBase, App.Title, "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid."
        End If
        If Not oViewport Is Nothing Then
            Set oFormReader = oForm.RenderToCodeReader(sImageFile, sDictionaryName, oViewport)
            Set m_oXml_InnerLogo = oFormReader.FormInfo
            Do Until oFormReader.EndOfForm
                sCodeLine = oFormReader.ReadLine()
                .WriteLine sCodeLine
            Loop
        End If
        
        If ParseXmlForm(m_sTag, "BACKER", sImageFile, sDictionaryName, oViewport) = False Then
            Err.Raise g_lcErrBase, App.Title, "The Tag '" & m_sTag & "' as defined in the job setting's tag property is invalid."
        End If
        If Not oViewport Is Nothing Then
            Set oFormReader = oForm.RenderToCodeReader(sImageFile, sDictionaryName, oViewport)
            Set m_oXml_Backer = oFormReader.FormInfo
            Do Until oFormReader.EndOfForm
                sCodeLine = oFormReader.ReadLine()
                .WriteLine sCodeLine
            Loop
        End If
        
        ' ToDo: Make these selections parameters
        ' Institution settings
        Select Case m_sXmlInst
            
            Case "BPB"
            
                Const scBPBLogo = "BPB\BPBLogo.eps" 'CPTPK-3026
                Const scBorelLogo = "BPB\BPBLogo.eps"
                Const scBPB_TM0360PBacker = "BPB\TM0360P_Backer.tif"
                Const scBPB_102Form = "9900000\BPB-102.tif" 'CPTPK-3026
                
                If m_bXML_Archive Then
                    Set oViewport = New Viewport
                    Set oFormReader = oForm.RenderToCodeReader(scBPB_102Form, "BPB102", oViewport)
                    Set m_oBPP102Form = oFormReader.FormInfo
                    Do Until oFormReader.EndOfForm
                        sCodeLine = oFormReader.ReadLine()
                        .WriteLine sCodeLine
                    Loop
                End If
               
                'Exterior Logos
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 780
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                
                Set oFormReader = oForm.RenderToCodeReader(scBPBLogo, "BPBExteriorLogo", oViewport)
                Set m_oBPPExteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oFormReader = oForm.RenderToCodeReader(scBorelLogo, "BorelExteriorLogo", oViewport)
                Set m_oBPPExteriorLogo(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Interior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = m_lcSelfSealPanels_x
                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scBPBLogo, "BPBInteriorLogo", oViewport)
                Set m_oBPBInteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oFormReader = oForm.RenderToCodeReader(scBorelLogo, "BorelInteriorLogo", oViewport)
                Set m_oBPBInteriorLogo(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oViewport = New Viewport
                With oViewport
                    .Width = 612
                    .Height = 792
                    .ReferenceCorner.Location.X = 0
                    .ReferenceCorner.Location.Y = -150 'RM16209
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                End With
                Set oFormReader = oForm.RenderToCodeReader(scBPB_TM0360PBacker, "BPB_TM0360P_Backer", oViewport)
                Set m_oBPBBackers(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oViewport = New Viewport
                oViewport.ReferenceCorner.Location.Y = -100 'RM16212
                Set oFormReader = oForm.RenderToCodeReader("BPB\BPBRepurchaseAgreement_FormCache.xml", "RABacker", oViewport)
                Set m_oRABacker = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Flat stock Logo
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 711
                    .Size.Width = 216
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scBPBLogo, "FlatStockLogo", oViewport)
                Set m_oBPBFlatStockLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
            Case "SNB"
            
                Const scSNBLogo As String = "SNB\SNBLogo.eps"
                Const scSNBLogo2 As String = "SNB\SNBLogo2.tif"
                Const scSNBLogo_Jericho As String = "SNB\SNBLogo.eps"  '+RM14942
                'Const scSNBBacker1 As String = "SNB\SNBDisclosure.tif"        '-RM14942
                Const scSNBBacker1 As String = "SNB\SNB-LoanBacker530.tif"     '+RM14942
                Const scSNBBacker2 As String = "SNB\SNBTimeDeposit.tif"
                 Const scSNBBacker3 As String = "SNB\SNBBacker.tif"
                'Exterior Logos
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 780
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo, "SNBExteriorLogo", oViewport)
                Set m_oSNBExteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo2, "SNBExteriorLogo2", oViewport)
                Set m_oSNBExteriorLogo(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'END RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo_Jericho, "SNBExteriorLogoJericho", oViewport)
                Set m_oSNBExteriorLogo(2) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                
                'Interior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = m_lcSelfSealPanels_x
                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo, "SNBInteriorLogo", oViewport)
                Set m_oSNBInteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo2, "SNBInteriorLogo2", oViewport)
                Set m_oSNBInteriorLogo(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'END RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo_Jericho, "SNBInteriorLogoJericho", oViewport)
                Set m_oSNBInteriorLogo(2) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Flat stock Logo
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 711
                    .Size.Width = 216
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo, "FlatStockLogo", oViewport)
                Set m_oSNBFlatStockLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo2, "FlatStockLogo2", oViewport)
                Set m_oSNBFlatStockLogo(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                'END RM9344
                Set oFormReader = oForm.RenderToCodeReader(scSNBLogo_Jericho, "FlatStockLogoJericho", oViewport)
                Set m_oSNBFlatStockLogo(2) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Backer
                Set oViewport = New Viewport
                Set oFormReader = oForm.RenderToCodeReader(scSNBBacker1, "Backer", oViewport)
                Set m_oSNBBackers(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oViewport = New Viewport
                Set oFormReader = oForm.RenderToCodeReader(scSNBBacker2, "Backer2", oViewport)
                Set m_oSNBBackers(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                Set oViewport = New Viewport
                Set oFormReader = oForm.RenderToCodeReader(scSNBBacker3, "Backer3", oViewport)
                Set m_oSNBBackers(2) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
            Case "Nationwide"
            
'                Const scNationwideLogo As String = "Nationwide\Nationwide_Logo.tif"
'                Const scNationwideLogo2 As String = "Nationwide\Nationwide_Deposit_Logo.tif"
'                Const scNationwideBRE9Backer As String = "Nationwide\Nationwide_BRE9Backer.tif"
'
'                'Exterior Logos
'                Set oViewport = New Viewport
'                With oViewport
'                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
'                    .ReferenceCorner.Location.X = 54
'                    .ReferenceCorner.Location.Y = 780
'                    .Size.Width = 0
'                    .Size.Height = 63
'                End With
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo, "ExteriorLogo0", oViewport)
'                Set m_oNationwideExteriorLogo(0) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo2, "ExteriorLogo1", oViewport)
'                Set m_oNationwideExteriorLogo(1) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'
'                'Interior Logo
'                Set oViewport = New Viewport
'                With oViewport
'                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
'                    .ReferenceCorner.Location.X = 54
'                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
'                    .Size.Width = 0
'                    .Size.Height = 63
'                End With
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo, "InteriorLogo0", oViewport)
'                Set m_oNationwideInteriorLogo(0) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo2, "InteriorLogo1", oViewport)
'                Set m_oNationwideInteriorLogo(1) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'                'Flat stock Logo
'                With oViewport
'                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
'                    .ReferenceCorner.Location.X = 54
'                    .ReferenceCorner.Location.Y = 711
'                    .Size.Width = 216
'                    .Size.Height = 63
'                End With
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo, "FlatStockLogo0", oViewport)
'                Set m_oNationwideFlatStockLogo(0) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideLogo2, "FlatStockLogo1", oViewport)
'                Set m_oNationwideFlatStockLogo(1) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
'
'                'Backer - currently there are no backers for Nationwide
'                Set oViewport = New Viewport
'                Set oFormReader = oForm.RenderToCodeReader(scNationwideBRE9Backer, "Backer1", oViewport)
'                Set m_oNationwideBackers(1) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop


            Case "ElDorado"
            
                Const scElDoradoLogo As String = "ElDorado\ElDorado_Logo.tif"
                
                'Exterior Logos
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 780
                    .Size.Width = 180
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scElDoradoLogo, "ElDoradoExteriorLogo", oViewport)
                Set m_oElDoradoExteriorLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                
                'Interior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
                    .Size.Width = 180
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scElDoradoLogo, "ElDoradoInteriorLogo", oViewport)
                Set m_oElDoradoInteriorLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Flat stock Logo
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 711
                    .Size.Width = 180
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(scElDoradoLogo, "FlatStockLogo", oViewport)
                Set m_oElDoradoFlatStockLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Backer - currently there are no backers for Nationwide
'                Set oViewport = New Viewport
'                Set oFormReader = oForm.RenderToCodeReader(scSNBBacker1, "Backer", oViewport)
'                Set m_oSNBBackers(0) = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop

            Case "Investors"
            
                Const scInvestorsLogo As String = "Investors\Investors_Logo.tif"
                Const scInvestorsLogoEdd As String = "Investors\Investors_Logo.eps"
                Const scInvestorsTM0847P1Backer As String = "Investors\Investors_AnnualNoticeWithholding_Backer.tif"
                
                'Exterior Logos
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 780
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(IIf(Len(m_sXML_EddJobKey) > 0, scInvestorsLogoEdd, scInvestorsLogo), "ExteriorLogo", oViewport)
                Set m_oInvestorsExteriorLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                
                'Interior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(IIf(Len(m_sXML_EddJobKey) > 0, scInvestorsLogoEdd, scInvestorsLogo), "InteriorLogo", oViewport)
                Set m_oInvestorsInteriorLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Flat stock Logo
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 711
                    .Size.Width = 216
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader(IIf(Len(m_sXML_EddJobKey) > 0, scInvestorsLogoEdd, scInvestorsLogo), "FlatStockLogo", oViewport)
                Set m_oInvestorsFlatStockLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
                'Backer - currently there are no backers for Nationwide
                Set oViewport = New Viewport
                Set oFormReader = oForm.RenderToCodeReader(scInvestorsTM0847P1Backer, "Backer1", oViewport)
                Set m_oInvestorsBackers(1) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop

            Case "Allstate"
                
                'Flat stock Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperRight
                    .ReferenceCorner.Location.X = 549
                    .ReferenceCorner.Location.Y = 738
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader("Allstate\AllstateLogo.tif", "FlatStockLogo", oViewport)
                Set m_oFlatStockLogoFormInfo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
                
            Case "RiverCity"
            
'                Set oViewport = New Viewport
'                With oViewport
'                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
'                    .ReferenceCorner.Location.X = 0
'                    .ReferenceCorner.Location.Y = 0
'                    .Size.Width = 0
'                    .Size.Height = 0
'                End With
'                Set oFormReader = oForm.RenderToCodeReader("RiverCity\RiverCityContactInfo.tif", "ContactInfo", oViewport)
'                Set m_oContactInfoFormInfo = oFormReader.FormInfo
'                Do Until oFormReader.EndOfForm
'                    sCodeLine = oFormReader.ReadLine()
'                    .WriteLine sCodeLine
'                Loop
           
            Case "The Palmetto Bank"
               
                'Flat stock Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 711
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader("CBS\ThePalmettoBankLogo.tif", "Logo", oViewport)
                Set m_oBPBFlatStockLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
               
                'Backer
                Set oViewport = New Viewport
                Set oFormReader = oForm.RenderToCodeReader("CBS\ThePalmettoBankLoanBacker.tif", "Backer", oViewport)
                Set m_oBackerInfo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
            
            Case "WooriAmerica"
                ' No graphics for Woori America Bank
                
            Case "HigherOne"
                ' Graphics for HigherOne here
                Const scHigherOneLogo = "CBS\HigherOneLogo.tif"
                
                'Exterior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 770
                    .Size.Width = 0
                    .Size.Height = 40
                End With
                
                Set oFormReader = oForm.RenderToCodeReader(scHigherOneLogo, "HigherOneExteriorLogo", oViewport)
                Set m_oHigherOneExteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    ' sCodeLine = oFormReader.ReadLine()
                    ' .WriteLine sCodeLine
                    .WriteLine oFormReader.ReadLine()
                Loop
                
                'Interior Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_UpperLeft
                    .ReferenceCorner.Location.X = m_lcSelfSealPanels_x
                    .ReferenceCorner.Location.Y = m_lcSelfSealPanels_y + m_lcSelfSealPanels_h
                    .Size.Width = 0
                    .Size.Height = 35
                End With
                
                Set oFormReader = oForm.RenderToCodeReader(scHigherOneLogo, "HigherOneInteriorLogo", oViewport)
                Set m_oHigherOneInteriorLogo(0) = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    ' sCodeLine = oFormReader.ReadLine()
                    ' .WriteLine sCodeLine
                    .WriteLine oFormReader.ReadLine()
                Loop
                
            
            Case m_scInst_29175100_FinFactors
                'Flat stock Logo
                Set oViewport = New Viewport
                With oViewport
                    .ReferenceCorner.Corner = RectangleCorner_LowerLeft
                    .ReferenceCorner.Location.X = 54
                    .ReferenceCorner.Location.Y = 710
                    .Size.Width = 0
                    .Size.Height = 63
                End With
                Set oFormReader = oForm.RenderToCodeReader("29175100\29175100_Logo.eps", "Logo", oViewport)
                Set m_oXml_FlatStockLogo = oFormReader.FormInfo
                Do Until oFormReader.EndOfForm
                    sCodeLine = oFormReader.ReadLine()
                    .WriteLine sCodeLine
                Loop
               
                'Backer
                Select Case m_sXmlDocument
                    Case scDocument_TM0330P1
                        Set oViewport = New Viewport
                        Set oFormReader = oForm.RenderToCodeReader("29175100\29175100_TM0330_Backer.tif", "Backer", oViewport)
                        Set m_oXml_Backer = oFormReader.FormInfo
                        Do Until oFormReader.EndOfForm
                            sCodeLine = oFormReader.ReadLine()
                            .WriteLine sCodeLine
                        Loop
                    Case Else
                        Set oViewport = New Viewport
                        Set oFormReader = oForm.RenderToCodeReader("29175100\29175100_HELOCBacker.tif", "Backer", oViewport)
                        Set m_oXml_Backer = oFormReader.FormInfo
                        Do Until oFormReader.EndOfForm
                            sCodeLine = oFormReader.ReadLine()
                            .WriteLine sCodeLine
                        Loop
                End Select
            
            
            Case "55810000", "75535000"
            
            Case Else
                Stop
                
        End Select
        
        .WriteLine ""
        .WriteLine "[{0"
        .WriteLine "/languagelevel where{pop languagelevel 2 ge}{false}ifelse"
        .WriteLine "{1 dict dup/JobTimeout 4 -1 roll put setuserparams}"
        .WriteLine "{statusdict/setjobtimeout get exec}ifelse"
        .WriteLine "}stopped cleartomark"
        .WriteLine "[{120"
        .WriteLine "/languagelevel where{pop languagelevel 2 ge}{false}ifelse"
        .WriteLine "{1 dict dup/WaitTimeout 4 -1 roll put setuserparams}"
        .WriteLine "{statusdict/waittimeout 3 -1 roll put}ifelse"
        .WriteLine "}stopped cleartomark"
        .WriteLine "<< /NumCopies 1  >> setpagedevice"
        .WriteLine "%%EndSetup"
    End With
End Sub

Private Sub PsWriteTrailer(ByVal lPageCount As Long)
    With m_oStmtBodyFile
        .WriteLine "%%Trailer"
        .WriteLine "%%DocumentNeededFonts:"
        .WriteLine "%%+ Helvetica"
        .WriteLine "%%+ Helvetica-Bold"
        .WriteLine "%%DocumentSuppliedFonts:"
        .WriteLine "%%Pages: " & CStr(lPageCount)
        .WriteLine "%%EOF"
    End With
End Sub

Private Function PsProc001_DTL(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTL = PsString(sPrintLine, True) & " " & CStr(sngPos) & " DTL"
End Function

Private Function PsProc001_DTLB(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTLB = PsString(sPrintLine, True) & " " & CStr(sngPos) & " DTLB"
End Function

Private Function PsProc_SECH(ByRef sngPos As Single, ByVal sSubHeading As String) As String
    PsProc_SECH = PsString(sSubHeading, True) & " " & CStr(sngPos) & " SECH"
End Function

Private Function PsProc_PYMTCHGGRID(ByVal sngPos As Single, ByVal sngLineSpacing As Single) As String
    PsProc_PYMTCHGGRID = CStr(sngLineSpacing) & " " & CStr(sngPos) & " PYMTCHGGRID"
End Function

Private Function PsProc001_DTL9(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTL9 = PsString(sPrintLine, True) & CStr(sngPos) & " DTL9"
End Function

Private Function PsProc001_DTLB9(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTLB9 = PsString(sPrintLine, True) & CStr(sngPos) & " DTLB9"
End Function

Private Function PsProc001_DTL8(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTL8 = PsString(sPrintLine, True) & CStr(sngPos) & " DTL8"
End Function

Private Function PsProc001_DTLB8(ByRef sngPos As Single, ByRef sPrintLine As String) As String
    PsProc001_DTLB8 = PsString(sPrintLine, True) & CStr(sngPos) & " DTLB8"
End Function

Private Function PsProc001_DTL3(ByRef sngX As Single, ByRef sngY As Single, ByRef sPrintLine As String) As String
    PsProc001_DTL3 = PsString(sPrintLine, True) & Space(1) & CStr(sngX) & Space(1) & CStr(sngY) & " DTL3"
End Function

Private Function PsProc_RETADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc_RETADDR = PsString(.AddressLine(1), True) & PsString(.AddressLine(2), True) & PsString(.AddressLine(3), True) & PsString(.AddressLine(4), True) & PsString(.AddressLine(5), True) & PsString(.AddressLine(6), True) & " RETADDR"
    End With
End Function
Private Function PsProc_RETADDRWoori(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc_RETADDRWoori = "%" & PsString(.AddressLine(1), True) & PsString(.AddressLine(2), True) & PsString(.AddressLine(3), True) & PsString(.AddressLine(4), True) & PsString(.AddressLine(5), True) & PsString(.AddressLine(6), True) & " RETADDR"
    End With
End Function

Private Function PsProcITM42_DELADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProcITM42_DELADDR = PsString(RTrim(.AlignedAddress(1)), True) & Space(1) & PsString(RTrim(.AlignedAddress(2)), True) & Space(1) & PsString(RTrim(.AlignedAddress(3)), True) & Space(1) & PsString(RTrim(.AlignedAddress(4)), True) & Space(1) & PsString(RTrim(.AlignedAddress(5)), True) & Space(1) & PsString(RTrim(.AlignedAddress(6)), True) & " ITM42DELADDR"
    End With
End Function

Private Function PsProcITM42_RETADDR(ByRef oAddress As CAddress, Optional ByVal bForeign As Boolean) As String
    If oAddress.CondensedAddressLines > 5 Then
        Err.Raise 7, App.Title, "Invalid procedure call.  Address argument has to many lines."
    End If
    If bForeign Then
        Dim oAddressWithUSA As CAddress
        Set oAddressWithUSA = AddUSA(oAddress)
        With oAddressWithUSA
            PsProcITM42_RETADDR = PsString(RTrim(.AlignedAddress(1)), True) & Space(1) & PsString(RTrim(.AlignedAddress(2)), True) & Space(1) & PsString(RTrim(.AlignedAddress(3)), True) & Space(1) & PsString(RTrim(.AlignedAddress(4)), True) & Space(1) & PsString(RTrim(.AlignedAddress(5)), True) & " RETADDR"
        End With
    Else
        With oAddress
            PsProcITM42_RETADDR = PsString(RTrim(.CondensedAddress(1)), True) & Space(1) & PsString(RTrim(.CondensedAddress(2)), True) & Space(1) & PsString(RTrim(.CondensedAddress(3)), True) & Space(1) & PsString(RTrim(.CondensedAddress(4)), True) & Space(1) & PsString(RTrim(.AlignedAddress(5)), True) & " ITM42RETADDR"
        End With
    End If
End Function

Private Function PsProcSTD10_RETADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProcSTD10_RETADDR = PsString(RTrim(.AlignedAddress(1)), True) & Space(1) & PsString(RTrim(.AlignedAddress(2)), True) & Space(1) & PsString(RTrim(.AlignedAddress(3)), True) & Space(1) & PsString(RTrim(.AlignedAddress(4)), True) & Space(1) & PsString(RTrim(.AlignedAddress(5)), True) & Space(1) & PsString(RTrim(.AlignedAddress(6)), True) & " STD10RETADDR"
    End With
End Function

Private Function PsProcSTD10_DELADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProcSTD10_DELADDR = PsString(RTrim(.AlignedAddress(1)), True) & PsString(RTrim(.AlignedAddress(2)), True) & PsString(RTrim(.AlignedAddress(3)), True) & PsString(RTrim(.AlignedAddress(4)), True) & PsString(RTrim(.AlignedAddress(5)), True) & Space(1) & PsString(RTrim(.AlignedAddress(6)), True) & " STD10DELADDR"
    End With
End Function

Private Function PsProc001_DELADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc001_DELADDR = PsString(.AlignedAddress(1), True) & PsString(.AlignedAddress(2), True) & PsString(.AlignedAddress(3), True) & PsString(.AlignedAddress(4), True) & PsString(.AlignedAddress(5), True) & PsString(.AlignedAddress(6), True) & " DELADDR"
    End With
End Function

Private Function PsProc_DELADDR(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc_DELADDR = PsString(.AlignedAddress(1), True) & PsString(.AlignedAddress(2), True) & PsString(.AlignedAddress(3), True) & PsString(.AlignedAddress(4), True) & PsString(.AlignedAddress(5), True) & PsString(.AlignedAddress(6), True) & " DELADDR"
    End With
End Function
Private Function PsProc_DELADDR42(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc_DELADDR42 = PsString(.AlignedAddress(1), True) & PsString(.AlignedAddress(2), True) & PsString(.AlignedAddress(3), True) & PsString(.AlignedAddress(4), True) & PsString(.AlignedAddress(5), True) & PsString(.AlignedAddress(6), True) & " DELADDR42"
    End With
End Function

Private Function PsProc001_DELADDR3(ByRef oAddress As CAddress) As String
    With oAddress
        PsProc001_DELADDR3 = PsString(.AlignedAddress(1), True) & PsString(.AlignedAddress(2), True) & PsString(.AlignedAddress(3), True) & PsString(.AlignedAddress(4), True) & PsString(.AlignedAddress(5), True) & PsString(.AlignedAddress(6), True) & " DELADDR3"
    End With
End Function

Private Function IsBulkShip(ByRef oAddress As CAddress) As Boolean
    Dim i As Long
    
    If Len(oAddress.CityStateZip) = 0 Then
        IsBulkShip = True
    End If
    
    For i = 1 To oAddress.CondensedAddressLines
        If InStr(1, oAddress.CondensedAddress(i), "NO MAIL", vbTextCompare) > 0 Then
            IsBulkShip = True
            Exit Function
        End If
        If InStr(1, oAddress.CondensedAddress(i), "SEND TO", vbTextCompare) > 0 Then
            IsBulkShip = True
            Exit Function
        End If
    Next
   
    IsBulkShip = False
    
End Function

Private Function SetupNoPrintStack(ByVal lStackNumber As Long, ByVal sStackDesc As String) As Boolean
    Dim oStack As CStack
    
    On Error GoTo SetupNoPrintStack_Error
        
    Set oStack = New CStack
    
'    If lStackNumber > 1000 Then
'        Stop
'    End If
    
    With oStack
        
        .SetPropertiesEx lStackNumber + 5000, StackNoPrint, "Bypassed - " & sStackDesc, BalGroupNoPrint, ffgStackDuplexType.StackSimplex, False, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupNoPrintStack_ErrorExit
        End If
        
    End With
    
    SetupNoPrintStack = True
    
    Exit Function
   
SetupNoPrintStack_ErrorExit:
    
    SetupNoPrintStack = False
    Exit Function
   
SetupNoPrintStack_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupNoPrintStack()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupNoPrintStack_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupNoPrintStack_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupNoPrintStack_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupNoPrintStack_ErrorExit


End Function

Private Function GetSampleStack(ByVal sReportName As String) As Long
    Dim i As Long
    For i = 0 To m_lNoPrintStackIdx
        If m_saNoPrintStack(i) = sReportName Then
            GetSampleStack = i + 1
            Exit Function
        End If
    Next
    
    m_lNoPrintStackIdx = m_lNoPrintStackIdx + 1
    ReDim Preserve m_saNoPrintStack(m_lNoPrintStackIdx)
    m_saNoPrintStack(m_lNoPrintStackIdx) = sReportName
    
    SetupSampleStack m_lNoPrintStackIdx + 1, sReportName
    
    GetSampleStack = m_lNoPrintStackIdx + 1
End Function

Private Function SetupSampleStack(ByVal lStackNumber As Long, ByVal sReportName As String) As Boolean
    Dim oStack As CStack
    
    On Error GoTo SetupSampleStack_Error
        
    Set oStack = New CStack
    
'    If lStackNumber > 1000 Then
'        Stop
'    End If
    
    With oStack
        
        .SetPropertiesEx lStackNumber, StackNoMail, "Sample - " & sReportName, BalGroupPrint, ffgStackDuplexType.StackSimplex, False, False, 0, 0
        .SetPropertiesFileGroup sReportName, ""
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupSampleStack_ErrorExit
        End If
        
    End With
    
    SetupSampleStack = True
    
    Exit Function
   
SetupSampleStack_ErrorExit:
    
    SetupSampleStack = False
    Exit Function
   
SetupSampleStack_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupSampleStack()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupSampleStack_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupSampleStack_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupSampleStack_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupSampleStack_ErrorExit


End Function

Private Function SetupBulkShipStack(ByVal lStackNumber As Long, ByVal sFileName As String, ByVal sFormType As String, Optional sSrfKey As String) As Boolean
    Dim oStack As CStack
    
    On Error GoTo SetupBulkShipStack_Error
        
    Set oStack = New CStack
    
    With oStack
    
        Dim sStackDesc As String
        sStackDesc = "Bulk Ship - " & sFileName & " " & sFormType & " " & IIf(Len(sSrfKey) > 0, " ~SrfKey:" & sSrfKey, "")
        
        .SetPropertiesEx lStackNumber, IIf(Len(sSrfKey) > 0, StackNoMail, StackAutoMail), sStackDesc, BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, False, False, 0, 0
        If Len(sSrfKey) > 0 Then
            .SetPropertiesFileGroup sSrfKey & "_0000_00_FB", sStackDesc
        Else
            .SetPropertiesFileGroup "FB", sStackDesc
        End If
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupBulkShipStack_ErrorExit
        End If
        
    End With
    
    SetupBulkShipStack = True
    
    Exit Function
   
SetupBulkShipStack_ErrorExit:
    
    SetupBulkShipStack = False
    Exit Function
   
SetupBulkShipStack_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupBulkShipStack()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupBulkShipStack_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupBulkShipStack_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupBulkShipStack_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupBulkShipStack_ErrorExit


End Function

Private Function SetupStacks_Archive() As Boolean
    Dim oStack As CStack
    Set oStack = New CStack
    
    With oStack
        
        .SetPropertiesEx 1, StackNoMail, "Archive", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_Archive_ErrorExit
        End If
    
        .SetPropertiesEx 2, StackNoPrint, "Suppressed Accounts", BalGroupNoPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_Archive_ErrorExit
        End If
    
    End With
    
    SetupStacks_Archive = True
    
    Exit Function
   
SetupStacks_Archive_ErrorExit:
    
    SetupStacks_Archive = False
    Exit Function
   
SetupStacks_Archive_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupStacks_Archive()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupStacks_Archive_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupStacks_Archive_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupStacks_Archive_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupStacks_Archive_ErrorExit


End Function


'Depracated - Use SetupStacks_Edd
Private Function SetupStacks_EddEstatements() As Boolean
    Dim oStack As CStack
    Set oStack = New CStack
    
    With oStack
        
        .SetPropertiesEx 1, StackNoMail, "Estatements", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEstatements_ErrorExit
        End If
    
        .SetPropertiesEx 2, StackNoPrint, "Suppressed Accounts", BalGroupNoPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEstatements_ErrorExit
        End If
    
    End With
    
    SetupStacks_EddEstatements = True
    
    Exit Function
   
SetupStacks_EddEstatements_ErrorExit:
    
    SetupStacks_EddEstatements = False
    Exit Function
   
SetupStacks_EddEstatements_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupStacks_EddEstatements()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupStacks_EddEstatements_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupStacks_EddEstatements_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupStacks_EddEstatements_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupStacks_EddEstatements_ErrorExit


End Function

Private Function SetupStacks_Edd() As Boolean
    Dim oStack As CStack
    Set oStack = New CStack
    
    With oStack
        
        .SetPropertiesEx 1, StackNoMail, "Estatements", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup m_scEddJobNamePrefix + m_sXML_EddJobKey, vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_Edd_ErrorExit
        End If
        .SetPropertiesEx 2, StackNoPrint, "Suppressed Accounts", BalGroupNoPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_Edd_ErrorExit
        End If
    
    End With
    
    SetupStacks_Edd = True
    
    Exit Function
   
SetupStacks_Edd_ErrorExit:
    
    SetupStacks_Edd = False
    Exit Function
   
SetupStacks_Edd_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupStacks_Edd()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupStacks_Edd_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupStacks_Edd_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupStacks_Edd_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupStacks_Edd_ErrorExit


End Function

Private Function SetupStacks_EddEnotice() As Boolean
    Dim oStack As CStack
    Set oStack = New CStack
    
    With oStack
        
        .SetPropertiesEx 1, StackNoMail, "ENotice - 06 DD Checking or Combined 10", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup m_scEddJobNamePrefix + m_sXML_EddEnoticeJobKey + "06", vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEnotice_ErrorExit
        End If
        
        .SetPropertiesEx 2, StackNoMail, "ENotice - 01 SV Savings 20", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup m_scEddJobNamePrefix + m_sXML_EddEnoticeJobKey + "01", vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEnotice_ErrorExit
        End If
        
        .SetPropertiesEx 3, StackNoMail, "ENotice - 03 SV Savings 50", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup m_scEddJobNamePrefix + m_sXML_EddEnoticeJobKey + "03", vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEnotice_ErrorExit
        End If
        
        .SetPropertiesEx 4, StackNoMail, "ENotice - 05 TM Time (CD) 30", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup m_scEddJobNamePrefix + m_sXML_EddEnoticeJobKey + "05", vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEnotice_ErrorExit
        End If
        
        .SetPropertiesEx 5, StackNoPrint, "Suppressed Accounts", BalGroupNoPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_EddEnotice_ErrorExit
        End If
    
    End With
    
    SetupStacks_EddEnotice = True
    
    Exit Function
   
SetupStacks_EddEnotice_ErrorExit:
    
    SetupStacks_EddEnotice = False
    Exit Function
   
SetupStacks_EddEnotice_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupStacks_EddEnotice()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupStacks_EddEnotice_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupStacks_EddEnotice_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupStacks_EddEnotice_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupStacks_EddEnotice_ErrorExit


End Function

Private Function SetupStacks() As Boolean
    Dim oStack As CStack
    Dim oMailGroupOptions As CMailStackGroupOptions
    Dim lFlatFormPlexCode As ffgStackDuplexType
    Dim i As Long
    
    On Error GoTo SetupStacks_Error
        
    Set oStack = New CStack
    
    
    
    With oStack
    
        'Generic stacks
        
        .SetPropertiesEx 2, StackNoPrint, "No Print", BalGroupNoPrint, ffgStackDuplexType.StackSimplexDuplexMixed, False, False, 0, 0
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
        
        'Pressure Seal
        
        .SetPropertiesEx 102, StackNoMail, "Pressure Seal - Bulk Ship", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
        .SetPropertiesFileGroup "SB", vbNullString
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
        
'        .SetPropertiesEx 103, StackAutoMail, "Pressure Seal - Canada", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
'        .SetPropertiesFileGroup "SM", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
'
'        .SetPropertiesEx 104, StackAutoMail, "Pressure Seal - Mexico", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
'        .SetPropertiesFileGroup "SM", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
'
'        .SetPropertiesEx 105, StackAutoMail, "Pressure Seal - Foreign", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
'        .SetPropertiesFileGroup "SM", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
'
'        .SetPropertiesEx 106, StackAutoMail, "Pressure Seal - Full Rate", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
'        .SetPropertiesFileGroup "SM", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
'
'        .SetPropertiesEx 107, StackAutoMail, "Pressure Seal - Single Rate", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, False, 0, 0
'        .SetPropertiesFileGroup "SP", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
'
'        .SetPropertiesEx 108, StackAutoMail, "Pressure Seal - Auto Mail", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, True, 11, 0
'        .SetPropertiesFileGroup "SP", vbNullString
'        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'            GoTo SetupStacks_ErrorExit
'        End If
        
        Set oMailGroupOptions = New CMailStackGroupOptions
        oMailGroupOptions.StackIdInterval = 1

        .SetPropertiesEx 108, StackAutoMail, "Pressure Seal - ~", BalGroupPrint, ffgStackDuplexType.StackSimplexDuplexMixed, True, True, 0, 0
        .SetPropertiesFileGroup "S~", vbNullString
        If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
            GoTo SetupStacks_ErrorExit
        End If
        
        '------------------ Flat Stock Stacks
        lFlatFormPlexCode = StackSimplexDuplexMixed
        Dim lStackOffset As Long
        
        Dim sAssemblyInstructions As String
        
       If m_sXmlInst = "BPB" Then
            
            lStackOffset = 200
            sAssemblyInstructions = " in " & m_scEnvelopeBPB
            
            .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
                GoTo SetupStacks_ErrorExit
            End If
                
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
            
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
            
            lStackOffset = 300
            sAssemblyInstructions = " in " & m_scEnvelopeBPB & " w/ " & m_scEnclosureBPBEL9
        
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
        
       
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
            
            lStackOffset = 400
            sAssemblyInstructions = " in " & m_scEnvelopeBPB & " w/ " & m_scEnclosureBPBER9
        
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
       
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
            
            lStackOffset = 700
            ' sAssemblyInstructions = " in " & m_scEnvelopeBorel & " w/ " & m_scEnclosureBPBBORELER9
            sAssemblyInstructions = " in " & m_scEnvelopeBPB & " w/ " & m_scEnclosureBPBBORELER9
        
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
       
            lStackOffset = 800
            sAssemblyInstructions = " in " & m_scEnvelopeBPB & " Legal"
        
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                .PaperKind = "Legal"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                .PaperKind = "Legal"
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
       
            lStackOffset = 900
            sAssemblyInstructions = " in " & m_scEnvelope42
        
'            For i = 1 To 10
'                .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
'
'            For i = 1 To 10
'                .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'                .SetPropertiesFileGroup "IM", "Metered Mail"
'                If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                    GoTo SetupStacks_ErrorExit
'                End If
'            Next
            Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
            
            For i = 1 To 10
                .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
                .SetPropertiesFileGroup "I~", vbNullString
                If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                    GoTo SetupStacks_ErrorExit
                End If
            Next
       
            
            
    ElseIf m_sXmlInst = "WooriAmerica" Then
        'Woori America stacks
        lStackOffset = 200
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        ' T152Q2 stack
        lStackOffset = 300
        sAssemblyInstructions = " with " & m_scWooriT152Q2BRE
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
    ElseIf m_sXmlInst = "Nationwide" Then
        'Nationwide stacks
        lStackOffset = 200
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        
        ' #9 BRE stack
        lStackOffset = 300
        sAssemblyInstructions = " with " & m_scNationwideBRE
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        
    ElseIf m_sXmlInst = "Investors" Then
        'Investors Bank Stacks
        lStackOffset = 200
        .PaperKind = "Letter"
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        
        ' Legal sized paper stacks
        
        lStackOffset = 300
        .PaperKind = "Legal"
        
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & " Legal", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        
        ' #9 BRE stack
        lStackOffset = 400
        sAssemblyInstructions = " with " & m_scInvestorsBRE
        .PaperKind = "Letter"
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        
        ' #9 BRE stack + More
        lStackOffset = 500
        sAssemblyInstructions = " with " & m_scInvestorsBREMORE
        .PaperKind = "Letter"
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString) & sAssemblyInstructions, BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
        

        
    Else
        
        lStackOffset = 200
        
        .SetPropertiesEx lStackOffset, StackNoMail, "Bulk Ship", BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
        .SetPropertiesFileGroup "FB", "Bulk Ship Flat"
        If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
            GoTo SetupStacks_ErrorExit
        End If
            
'        For i = 1 To 10
'            .SetProperties lStackOffset + i, StackAutoMail, "Canada " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 10 + i, StackAutoMail, "Mexico " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 20 + i, StackAutoMail, "Foreign " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 30 + i, StackAutoMail, "Full Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 40 + i, StackAutoMail, "Single Rate " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
'
'        For i = 1 To 10
'            .SetProperties lStackOffset + 50 + i, StackAutoMail, "Auto Mail " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0
'            .SetPropertiesFileGroup "IM", "Metered Mail"
'            If m_oStmtBodyDatabase.InsertStack(oStack) = False Then
'                GoTo SetupStacks_ErrorExit
'            End If
'        Next
        Set oMailGroupOptions = New CMailStackGroupOptions
'            oMailGroupOptions.StackIdInterval = 1
        
        For i = 1 To 10
            .SetPropertiesEx lStackOffset + 50 + i, StackAutoMail, "~ " & CStr(i) & " sheet" & IIf(i = 10, "s+", vbNullString), BalGroupPrint, lFlatFormPlexCode, True, False, 0, 0
            .SetPropertiesFileGroup "I~", vbNullString
            If m_oStmtBodyDatabase.InsertMailStackGroup(oStack, oMailGroupOptions) = False Then
                GoTo SetupStacks_ErrorExit
            End If
        Next
        
    End If
        
        
    End With
    
    SetupStacks = True
    
    Exit Function
   
SetupStacks_ErrorExit:
    
    SetupStacks = False
    Exit Function
   
SetupStacks_Error:
    
    'Preserve the context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scErrorContext = "The following error occured in CAdmin.SetupStacks()"
    
    On Error GoTo 0
    'Handle error
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scErrorContext)
       Case ErrorDispatchEnum.ErrorDispNotHandled 'Error not handled - regenerate error
           Err.Raise m_lErr, m_sErrSrc, m_sErrDesc, m_sHelpFile, m_sHelpContext
       Case ErrorDispatchEnum.ErrorDispRetryResume 'User chose Retry
           On Error GoTo SetupStacks_Error
           Resume
       Case ErrorDispatchEnum.ErrorDispIgnoreResumeNext 'User chose Ignore
           On Error GoTo SetupStacks_Error
           Resume Next
       Case ErrorDispatchEnum.ErrorDispAbortRaiseError 'User chose Abort
           GoTo SetupStacks_ErrorExit
           Exit Function
    End Select
    
    GoTo SetupStacks_ErrorExit


End Function

'Private Sub PsWriteProlog()
'    With m_oStmtBodyFile
'        .WriteLine "%%BeginProlog"
'        .WriteLine "%%IncludeResource: procset FDSGlobal 1.0 0"
'        .WriteLine "%%BeginResource: procset FDSLocal 1.0 0"
'        .WriteLine "FDSGlobal begin"
'        .WriteLine "userdict /FDSLocal 100 dict dup begin put"
'        .WriteLine "%%IncludeResource: procset OMRPDI1 1.0 0"
'        .WriteLine "/IncludePrintMediaForm false def"
'        .WriteLine "%Delivery Address Block for Pressure Seal"
'        .WriteLine "/PostNet{ %def () PostNet"
'        .WriteLine "    gs"
'        .WriteLine "    216 72 translate"
'        .WriteLine "    PostNetFont setfont"
'        .WriteLine "    0 0 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/DELADDR{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    216 72 translate"
'        .WriteLine "    Helv10 setfont"
'        .WriteLine "    0 14 MS  % line 6"
'        .WriteLine "    0 24 MS  % line 5"
'        .WriteLine "    0 34 MS  % line 4"
'        .WriteLine "    0 44 MS  % line 3"
'        .WriteLine "    0 54 MS  % line 2"
'        .WriteLine "    0 64 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "    gs"
''        .WriteLine "    HelvBold10 setfont"
''        .WriteLine "    (FIRST-CLASS MAIL) 306 216 MS"
''        .WriteLine "    HelvBold9 setfont"
''        .WriteLine "    (Important Tax Return) 306 204 MS"
''        .WriteLine "    (Document Enclosed) 306 194 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/SEQ { %def"
'        .WriteLine "    gs"
'        .WriteLine "    216 72 translate"
'        .WriteLine "    Helv6 setfont"
'        .WriteLine "    270 0 RS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "%Return Address Block"
'        .WriteLine "/RETADDR{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    42 216 translate"
'        .WriteLine "    Helv8 setfont"
'        .WriteLine "    0 -30 MS  % line 4"
'        .WriteLine "    0 -20 MS  % line 3"
'        .WriteLine "    0 -10 MS  % line 2"
'        .WriteLine "    HelvBold8 setfont"
'        .WriteLine "    0 0 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "% Ancillary Service Endorsement - Return Service"
'        .WriteLine "/ASER{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    42 216 translate"
'        .WriteLine "    HelvBold8 setfont"
'        .WriteLine "    " & PsString(m_scXmlAse, True) & " 0 -50 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/DTL { %def  x y () DTL"
'        .WriteLine "    gs"
'        .WriteLine "    /Start_y exch def"
'        .WriteLine "    ArialMT setfont"
'        .WriteLine "    0 Start_y MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'
'
'
'        .WriteLine "%Delivery Address Block for 3 up"
'        .WriteLine "/PostNet3{ %def () PostNet3"
'        .WriteLine "    gs"
'        .WriteLine "    72 40 translate"
'        .WriteLine "    PostNetFont setfont"
'        .WriteLine "    0 0 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/DELADDR3{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    72 40 translate"
'        .WriteLine "    Helv10 setfont"
'        .WriteLine "    0 14 MS  % line 6"
'        .WriteLine "    0 24 MS  % line 5"
'        .WriteLine "    0 34 MS  % line 4"
'        .WriteLine "    0 44 MS  % line 3"
'        .WriteLine "    0 54 MS  % line 2"
'        .WriteLine "    0 64 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/SEQ3 { %def"
'        .WriteLine "    gs"
'        .WriteLine "    72 40 translate"
'        .WriteLine "    Helv6 setfont"
'        .WriteLine "    270 0 RS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "%Return Address Block"
'        .WriteLine "/RETADDR3{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    18 245 translate"
'        .WriteLine "    Helv8 setfont"
'        .WriteLine "    0 -30 MS  % line 4"
'        .WriteLine "    0 -20 MS  % line 3"
'        .WriteLine "    0 -10 MS  % line 2"
'        .WriteLine "    HelvBold8 setfont"
'        .WriteLine "    0 0 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/DTL3 { %def  () x y DTL3"
'        .WriteLine "    gs"
'        .WriteLine "    /Start_y exch def"
'        .WriteLine "    /Start_x exch def"
'        .WriteLine "    ArialMT setfont"
'        .WriteLine "    Start_x Start_y MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'
'
'        'Flat Stock procedures ----
'        .WriteLine "/PostNetFS{ %def () PostNetFS"
'        .WriteLine "    gs"
'        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
'        .WriteLine "    PostNetFont setfont"
'        .WriteLine "    0 0 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/DELADDRFS{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    Helv10 setfont"
'        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
'        .WriteLine "    0 14 MS  % line 6"
'        .WriteLine "    0 24 MS  % line 5"
'        .WriteLine "    0 34 MS  % line 4"
'        .WriteLine "    0 44 MS  % line 3"
'        .WriteLine "    0 54 MS  % line 2"
'        .WriteLine "    0 64 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/SEQFS { %def"
'        .WriteLine "    gs"
'        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
'        .WriteLine "    Helv6 setfont"
'        .WriteLine "    270 0 RS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "%Return Address Block"
'        .WriteLine "/RETADDRFS{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " 765 translate"
'        .WriteLine "    Helv10 setfont"
'        .WriteLine "    0 -45 MS  % line 6"
'        .WriteLine "    0 -36 MS  % line 5"
'        .WriteLine "    0 -27 MS  % line 4"
'        .WriteLine "    0 -18 MS  % line 3"
'        .WriteLine "    0 -9 MS  % line 2"
'        .WriteLine "    0 -0 MS  % line 1"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'        .WriteLine "/ASEFS { %def"
'        .WriteLine "    gs"
'        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
'        .WriteLine "    Helv10 setfont"
'        .WriteLine "    0 86 MS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
'
'        .WriteLine "/PostNetFont /MB043 findfont 12 scalefont def"
'        .WriteLine "%ArialMonospacedMT fonts"
'        .WriteLine "/ArialMonospacedMT findfont"
'        .WriteLine "dup 10 scalefont /ArialMT exch def"
'        .WriteLine "pop"
'        .WriteLine "%Helvetica fonts"
'        .WriteLine "/Helvetica findfont"
'        .WriteLine "dup 6 scalefont /Helv6 exch def"
'        .WriteLine "dup 8 scalefont /Helv8 exch def"
'        '.WriteLine "dup 9 scalefont /Helv9 exch def"
'        .WriteLine "dup 10 scalefont /Helv10 exch def"
'        .WriteLine "pop"
'        .WriteLine "%Helvetica-Bold fonts"
'        .WriteLine "/Helvetica-Bold findfont"
'        .WriteLine "dup 8 scalefont /HelvBold8 exch def"
'        .WriteLine "dup 9 scalefont /HelvBold9 exch def"
'        .WriteLine "dup 10 scalefont /HelvBold10 exch def"
'        '.WriteLine "dup 14 scalefont /HelvBold14 exch def"
'        .WriteLine "pop"
'        .WriteLine "end % FDSLocal"
'        .WriteLine "end % FDSGlobal"
'        .WriteLine "%%EndResource"
'
'        .WriteLine "%%IncludeResource: procset FDSFormsOps 1.0 0"
'
'        If Not m_oFlatStockLogoFormInfo Is Nothing Then
'            .WriteLine m_oFlatStockLogoFormInfo.DictionaryString
'        End If
'
'        If Not m_oFlatStockLogoFormInfo Is Nothing Then
'            .WriteLine m_oFlatStockLogoFormInfo.DictionaryString
'        End If
'
'        If Not m_oIndicia Is Nothing Then
'            .WriteLine m_oIndicia.DictionaryString
'        End If
'
'        If Not m_oPrintMediaForm Is Nothing Then
'            .WriteLine m_oPrintMediaForm.DictionaryString
'        End If
'
'        .WriteLine "%%EndProlog"
'    End With
'End Sub

Private Sub PsWriteProlog_Begin()
    With m_oStmtBodyFile
        .WriteLine "%%BeginProlog"
        .WriteLine "%%IncludeResource: procset FDSGlobal 1.0 0"
        .WriteLine "%%IncludeResource: font ArialMonospacedMT.PS"
        .WriteLine "%%IncludeResource: font ArialMonospacedMT-Bold.PS"
        .WriteLine "%%IncludeResource: font OCRA.PS"
        .WriteLine "%%IncludeResource: font E13B.PS"
    End With
End Sub

Private Sub PsWriteProlog_End()
    With m_oStmtBodyFile
        .WriteLine "%%IncludeResource: procset FDSFormsOps 1.0 0"
        .WriteLine "%%IncludeResource: procset OMRPDI1 1.0 0"
        
        .WriteLine "%%EndProlog"
    End With
End Sub

Private Sub PsWriteProlog_FlatStock()
    With m_oStmtBodyFile
        .WriteLine "%%BeginResource: procset FDSFlatStock 1.0 0"
        .WriteLine "FDSGlobal begin"
        .WriteLine "userdict /FDSFlatStock 100 dict dup begin put"
        
        'Flat Stock procedures ----
        .WriteLine "/PostNet{ %def () PostNetFS"
        .WriteLine "    gs"
        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
        .WriteLine "    PostNetFont setfont"
        .WriteLine "    0 0 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/DELADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
        .WriteLine "    0 14 MS  % line 6"
        .WriteLine "    0 24 MS  % line 5"
        .WriteLine "    0 34 MS  % line 4"
        .WriteLine "    0 44 MS  % line 3"
        .WriteLine "    0 54 MS  % line 2"
        .WriteLine "    0 64 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/SEQ { %def"
        .WriteLine "    gs"
        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    270 0 RS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        '#42 envelope
        .WriteLine "/PostNet42{ %def () PostNetFS"
        .WriteLine "    gs"
        .WriteLine "    54 620 translate"
        .WriteLine "    PostNetFont setfont"
        .WriteLine "    0 0 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/DELADDR42{ %def"
        .WriteLine "    gs"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    54 620 translate"
        .WriteLine "    0 14 MS  % line 6"
        .WriteLine "    0 24 MS  % line 5"
        .WriteLine "    0 34 MS  % line 4"
        .WriteLine "    0 44 MS  % line 3"
        .WriteLine "    0 54 MS  % line 2"
        .WriteLine "    0 64 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/SEQ42 { %def"
        .WriteLine "    gs"
        .WriteLine "    54 620 translate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    270 0 RS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
         .WriteLine "%MICR Line"
        .WriteLine "/MICR {"
        .WriteLine "    gs"
        .WriteLine "    E13B setfont"
        .WriteLine "    0.5 19.5 moveto [ 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 ] xshow"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "%Return Address Block"
        .WriteLine "/RETADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " 765 translate"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    0 -45 MS  % line 6"
        .WriteLine "    0 -36 MS  % line 5"
        .WriteLine "    0 -27 MS  % line 4"
        .WriteLine "    0 -18 MS  % line 3"
        .WriteLine "    0 -9 MS  % line 2"
        .WriteLine "    0 -0 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/ASE { %def"
        .WriteLine "    gs"
        .WriteLine "    " & CStr(m_lXmlAddressPositionX) & " " & CStr(m_lXmlAddressPositionY) & " translate"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    0 86 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/OCR{ %def x y () OCR"
        .WriteLine "    gs"
        .WriteLine "     OCRA setfont"
        .WriteLine "     MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/FRMCD{ %def () FRMCD"
        .WriteLine "    gs"
        .WriteLine "    612 0 translate 90 rotate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    72 18 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTL { %def  x y () DTL"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMT10 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTL9 { %def  x y () DTL9"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMT9 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTL8 { %def  x y () DTL8"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMT8 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTLB { %def  x y () DTL"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMTBold10 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTLB9 { %def  x y () DTLB9"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMTBold9 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTLB8 { %def  x y () DTLB8"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMTBold8 setfont"
        .WriteLine "    0 Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/SECH { %def  string y SECH"
        .WriteLine "  gs"
        .WriteLine "    0 exch translate"
        .WriteLine "    HelvBold10 setfont"
        .WriteLine "    .5 setlinewidth"
        .WriteLine "    dup stringwidth pop /sw exch def %width of caption"
        .WriteLine "    /linelen 468 sw 24 add sub 2 div def %width of line on each side of caption."
        .WriteLine "    12 4 moveto linelen 0 rlineto 12 0 rmoveto sw 0 rmoveto 12 0 rmoveto linelen 0 rlineto stroke"
        .WriteLine "    12 0 moveto linelen 0 rmoveto 12 0 rmoveto UnderLineString"
        .WriteLine "  gr"
        .WriteLine "}bd"
        
'        .WriteLine "/SECH { %def  string x y SUBH"
'        .WriteLine "    gs"
'        .WriteLine "    HelvBold10 setfont"
'        .WriteLine "    .5 setlinewidth"
'        .WriteLine "    moveto dup stringwidth"
'        .WriteLine "    2 div neg exch 2 div neg exch"
'        .WriteLine "    rmoveto "
'        .WriteLine "    /LineLen 72 def"
'        .WriteLine "    LineLen 0 gt {currentpoint 12 neg 0 rmoveto 0 3 rmoveto LineLen neg 0 rlineto stroke moveto} if"
'        .WriteLine "    show"
'        .WriteLine "    LineLen 0 gt {12 0 rmoveto 0 3 rmoveto LineLen 0 rlineto stroke} if"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
            
        .WriteLine "%Payment change details grid."
        .WriteLine "/PYMTCHGGRID { %def ls y PYMTCHGGRID"
        .WriteLine "  gs"
        .WriteLine "    .5 setlinewidth"
        .WriteLine "    0 exch translate"
        .WriteLine "    /ls exch def"
        .WriteLine "    %Horizontal lines."
        .WriteLine "    6 4 moveto                   474 0 rlineto "
        .WriteLine "    6 ls 4  mul neg 4 add moveto 474 0 rlineto "
        .WriteLine "    6 ls 6  mul neg 4 add moveto 474 0 rlineto "
        .WriteLine "    6 ls 10 mul neg 4 add moveto 474 0 rlineto "
        .WriteLine "    6 ls 13 mul neg 4 add moveto 474 0 rlineto "
        .WriteLine "    %Vertical lines."
        .WriteLine "    6   4 moveto 0 ls 13 mul neg rlineto"
        .WriteLine "    240 4 moveto 0 ls 13 mul neg rlineto"
        .WriteLine "    348 4 moveto 0 ls 13 mul neg rlineto"
        .WriteLine "    480 4 moveto 0 ls 13 mul neg rlineto"
        .WriteLine "    stroke"
        .WriteLine "  gr"
        .WriteLine "}bd"
        
        If m_sXmlInst = "RiverCity" Then
            .WriteLine "/RCBANK { %def  y RCBANK"
            .WriteLine "    gs"
            .WriteLine "    /Start_y exch def"
            .WriteLine "    HelvBold10 setfont"
            .WriteLine ""
            .WriteLine "    (Quick Access) 153 Start_y 33 add CS"
            .WriteLine "    (24Hour Banking Service) 153 Start_y 23 add CS"
            .WriteLine "    ((916) 567-2899) 153 Start_y 13 add CS"
            .WriteLine ""
            .WriteLine "    (Outside Local Area) 306 Start_y 28 add CS"
            .WriteLine "    ((800) 564-7144) 306 Start_y 18 add CS"
            .WriteLine ""
            .WriteLine "    (www.rcbank.com) 459 Start_y 23 add CS"
            .WriteLine ""
            .WriteLine "    gr"
            .WriteLine "}bd"
        End If
        
        If m_sXmlInst = m_scInst_29175100_FinFactors Then
            .WriteLine "/Form_TM0330P1_29175100 { %def Form_TM0330P1_29175100"
            .WriteLine "    gs"
            .WriteLine "    Helv18 setfont"
            .WriteLine "    (Statement of Account) 533 756 RS"
            .WriteLine "    Helv10 setfont"
            .WriteLine "    (Please examine at once and advise us promptly of any questions concerning your account.) 306 18 CS"
            .WriteLine "    HelvBold10 setfont"
            .WriteLine "    (Member FDIC) 533 45 RS"
            .WriteLine "    gr"
            .WriteLine "}bd"
        End If
        
        .WriteLine "/PostNetFont /MB043 findfont 12 scalefont def"
        
        .WriteLine "%ArialMonospacedMT fonts"
        .WriteLine "/ArialMonospacedMT findfont"
        .WriteLine "dup length dict begin"
        .WriteLine " {1 index /FID ne"
        .WriteLine "  {def}"
        .WriteLine "  {pop pop}"
        .WriteLine "  ifelse"
        .WriteLine " }forall"
        .WriteLine " /Encoding ISOLatin1Encoding def"
        .WriteLine " currentdict"
        .WriteLine "end"
        .WriteLine "/ArialMtIso exch definefont pop"
        
        .WriteLine "/ArialMtIso findfont"
        .WriteLine "dup 8 scalefont /ArialMT8 exch def"
        .WriteLine "dup 9 scalefont /ArialMT9 exch def"
        .WriteLine "dup 10 scalefont /ArialMT10 exch def"
        .WriteLine "pop"
        
        .WriteLine "%ArialMonospacedMT-Bold fonts"
        .WriteLine "/ArialMonospacedMT-Bold findfont"
        .WriteLine "dup length dict begin"
        .WriteLine " {1 index /FID ne"
        .WriteLine "  {def}"
        .WriteLine "  {pop pop}"
        .WriteLine "  ifelse"
        .WriteLine " }forall"
        .WriteLine " /Encoding ISOLatin1Encoding def"
        .WriteLine " currentdict"
        .WriteLine "end"
        .WriteLine "/ArialMtIso-Bold exch definefont pop"
        
        .WriteLine "/ArialMtIso-Bold findfont"
        .WriteLine "dup 8 scalefont /ArialMTBold8 exch def"
        .WriteLine "dup 9 scalefont /ArialMTBold9 exch def"
        .WriteLine "dup 10 scalefont /ArialMTBold10 exch def"
        .WriteLine "pop"
        
        .WriteLine "%Helvetica fonts"
        .WriteLine "/Helvetica findfont"
        .WriteLine "dup 6 scalefont /Helv6 exch def"
        .WriteLine "dup 8 scalefont /Helv8 exch def"
        .WriteLine "dup 10 scalefont /Helv10 exch def"
        .WriteLine "dup 18 scalefont /Helv18 exch def"
        .WriteLine "pop"
        .WriteLine "%Helvetica-Bold fonts"
        .WriteLine "/Helvetica-Bold findfont"
        .WriteLine "dup 8 scalefont /HelvBold8 exch def"
        .WriteLine "dup 9 scalefont /HelvBold9 exch def"
        .WriteLine "dup 10 scalefont /HelvBold10 exch def"
        .WriteLine "pop"
        .WriteLine "%OCRA fonts"
        .WriteLine "/OCRA findfont"
        .WriteLine "dup 10 scalefont /OCRA exch def"
        .WriteLine "pop"
        
        .WriteLine "%MICR fonts"
        .WriteLine "/E13B findfont"
        .WriteLine "dup 9 scalefont /E13B exch def"
        .WriteLine "pop"
        
        .WriteLine "end % FDSFlatStock"
        .WriteLine "end % FDSGlobal"
        .WriteLine "%%EndResource"
    End With
End Sub

Private Sub PsWriteProlog_Notice()
    With m_oStmtBodyFile
        .WriteLine "%%BeginResource: procset FDSNotice 1.0 0"
        .WriteLine "FDSGlobal begin"
        .WriteLine "userdict /FDSNotice 100 dict dup begin put"
        .WriteLine "/lm 66 def"
        
        '-----------------------------------------------------
        'ITM42 envelope procedures
        '-----------------------------------------------------
'
'        .WriteLine "%Form code"
'        .WriteLine "/ITM42FRMCD{ %def"
'        .WriteLine "    gs"
'        .WriteLine "    Helv6 setfont"
'        .WriteLine "    27 " & CStr(m_lcPressureSealRetAddrY) & " RS"
'        .WriteLine "    gr"
'        .WriteLine "}bd"
        
        .WriteLine "%Return Address Block"
        .WriteLine "/ITM42RETADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    54 " & CStr(m_lcPressureSealRetAddrY) & " translate"
        .WriteLine "    Helv8 setfont"
        .WriteLine "    0 -40 MS  % line 5"
        .WriteLine "    0 -30 MS  % line 4"
        .WriteLine "    0 -20 MS  % line 3"
        .WriteLine "    0 -10 MS  % line 2"
        .WriteLine "    HelvBold8 setfont"
        .WriteLine "    0 0 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "%Delivery Address Block"
        .WriteLine "/ITM42ASE{ %def"
        .WriteLine "    gs"
        '.WriteLine "    54 " & CStr(m_lcPressureSealRetAddrY) & " translate"
        .WriteLine "    54 620 translate"
        .WriteLine "    HelvBold8 setfont"
        .WriteLine "    0 77 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/ITM42DELADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    54 620 translate"
        .WriteLine "    ArialMT9 setfont"
        .WriteLine "    0 14 MS  % line 6"
        .WriteLine "    0 23 MS  % line 5"
        .WriteLine "    0 32 MS  % line 4"
        .WriteLine "    0 41 MS  % line 3"
        .WriteLine "    0 50 MS  % line 2"
        .WriteLine "    0 59 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/ITM42PostNet{ %def () PostNet"
        .WriteLine "    gs"
        .WriteLine "    54 620 translate"
        .WriteLine "    PostNetFont setfont"
        .WriteLine "    0 0 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/ITM42SEQ { %def"
        .WriteLine "    gs"
        .WriteLine "    54 620 translate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    270 0 RS"
        .WriteLine "    gr"
        .WriteLine "}bd"
'
'        '-----------------------------------------------------
'        'STD10 envelope procedures
'        '-----------------------------------------------------
        .WriteLine "/STD10PostNet{ %def () PostNet"
        .WriteLine "    gs"
        .WriteLine "    72 567 translate"
        .WriteLine "    PostNetFont setfont"
        .WriteLine "    0 0 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/STD10DELADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    72 567 translate"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    0 14 MS  % line 6"
        .WriteLine "    0 24 MS  % line 5"
        .WriteLine "    0 34 MS  % line 4"
        .WriteLine "    0 44 MS  % line 3"
        .WriteLine "    0 54 MS  % line 2"
        .WriteLine "    0 64 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "/STD10SEQ { %def"
        .WriteLine "    gs"
        .WriteLine "    72 567 translate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    270 0 RS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        .WriteLine "%Return Address Block"
        .WriteLine "/STD10RETADDR{ %def"
        .WriteLine "    gs"
        .WriteLine "    36 715 translate"
        .WriteLine "    Helv10 setfont"
        .WriteLine "    0 0 MS  % line 6"
        .WriteLine "    0 10 MS  % line 5"
        .WriteLine "    0 20 MS  % line 4"
        .WriteLine "    0 30 MS  % line 3"
        .WriteLine "    0 40 MS  % line 2"
        .WriteLine "    0 50 MS  % line 1"
        .WriteLine "    gr"
        .WriteLine "}bd"
'        '-----------------------------------------------------
        
        .WriteLine "/FRMCD{ %def () FRMCD"
        .WriteLine "    gs"
        .WriteLine "    612 0 translate 90 rotate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    72 18 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/FRMCDPS{ %def () FRMCD"
        .WriteLine "    gs"
        .WriteLine "    612 0 translate 90 rotate"
        .WriteLine "    Helv6 setfont"
        .WriteLine "    720 18 MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTL { %def  x y () DTL"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMT10 setfont"
        .WriteLine "    lm Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/DTLB { %def  x y () DTLB"
        .WriteLine "    gs"
        .WriteLine "    /Start_y exch def"
        .WriteLine "    ArialMTBold10 setfont"
        .WriteLine "    lm Start_y MS"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "%MICR Line"
        .WriteLine "/MICR {"
        .WriteLine "    gs"
        .WriteLine "    E13B setfont"
        .WriteLine "    0.5 19.5 moveto [ 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 ] xshow"
        .WriteLine "    gr"
        .WriteLine "}bd"
        
        .WriteLine "/PostNetFont /MB043 findfont 12 scalefont def"
        
        .WriteLine "%ArialMonospacedMT fonts"
        
        .WriteLine "/ArialMonospacedMT findfont"
        .WriteLine "dup length dict begin"
        .WriteLine " {1 index /FID ne"
        .WriteLine "  {def}"
        .WriteLine "  {pop pop}"
        .WriteLine "  ifelse"
        .WriteLine " }forall"
        .WriteLine " /Encoding ISOLatin1Encoding def"
        .WriteLine " currentdict"
        .WriteLine "end"
        .WriteLine "/ArialMtIso exch definefont pop"
        
        .WriteLine "/ArialMtIso findfont"
        .WriteLine "dup 9 scalefont /ArialMT9 exch def"
        .WriteLine "dup 10 scalefont /ArialMT10 exch def"
        .WriteLine "pop"
        
        .WriteLine "%ArialMonospacedMT-Bold fonts"
        .WriteLine "/ArialMonospacedMT-Bold findfont"
        .WriteLine "dup length dict begin"
        .WriteLine " {1 index /FID ne"
        .WriteLine "  {def}"
        .WriteLine "  {pop pop}"
        .WriteLine "  ifelse"
        .WriteLine " }forall"
        .WriteLine " /Encoding ISOLatin1Encoding def"
        .WriteLine " currentdict"
        .WriteLine "end"
        .WriteLine "/ArialMtIso-Bold exch definefont pop"
        
        .WriteLine "/ArialMtIso-Bold findfont"
        .WriteLine "dup 8 scalefont /ArialMTBold8 exch def"
        .WriteLine "dup 10 scalefont /ArialMTBold10 exch def"
        .WriteLine "pop"
        
        .WriteLine "%Helvetica fonts"
        .WriteLine "/Helvetica findfont"
        .WriteLine "dup 6 scalefont /Helv6 exch def"
        .WriteLine "dup 8 scalefont /Helv8 exch def"
        .WriteLine "dup 10 scalefont /Helv10 exch def"
        .WriteLine "pop"
        .WriteLine "%Helvetica-Bold fonts"
        .WriteLine "/Helvetica-Bold findfont"
        .WriteLine "dup 8 scalefont /HelvBold8 exch def"
        .WriteLine "dup 9 scalefont /HelvBold9 exch def"
        .WriteLine "dup 10 scalefont /HelvBold10 exch def"
        .WriteLine "pop"
        
        .WriteLine "%MICR fonts"
        .WriteLine "/E13B findfont"
        .WriteLine "dup 9 scalefont /E13B exch def"
        .WriteLine "pop"
        
        .WriteLine "end % FDSNotice"
        .WriteLine "end % FDSGlobal"
        .WriteLine "%%EndResource"
    End With
End Sub

Private Sub ISettings_SettingsSink(oSettingsSink As FFGPrintOlb.ISettingsSink)
    Set m_oSettingsSink = oSettingsSink
End Sub

'
'Private Function LoadImage(ByVal sDocFolder As String, ByRef oFormSpec As FDS_PostScript.FormSpec) As FDS_PostScript.Form
'
'    Dim sPath As String
'
'    On Error GoTo LoadImage_Error
'
'    If Len(sDocFolder) > 0 Then
'        sPath = sDocFolder & "\" & oFormSpec.FileName
'    Else
'        sPath = oFormSpec.FileName
'    End If
'
'    If Not IsFile(sPath) Then
'        Err.Raise g_lcErrBase + g_lcErrParameterError, App.Title, "Image file '" & sPath & "' does not exist."
'        Exit Function
'    End If
'
'    Set LoadImage = New FDS_PostScript.Form
'    With oFormSpec
'        LoadImage.LoadImage .DictionaryName, sPath, .X, .Y, .w, .h, .Center, .Alignment
'    End With
'
'    Exit Function
'
'LoadImage_ErrorExit:
'
'    Set LoadImage = Nothing
'
'    Exit Function
'
'LoadImage_Error:
'
'    'Preserve error context
'    m_lErr = Err.Number
'    m_sErrSrc = Err.source
'    m_sErrDesc = Err.Description
'    m_sHelpContext = Err.HelpContext
'    m_sHelpFile = Err.HelpFile
'
'    Const scIntro = "Error in CAdmin.LoadImage"
'
'    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
'       Case 1 'User chose Retry
'          On Error GoTo LoadImage_Error
'          Resume
'       Case 2 'User chose Ignore
'          On Error GoTo LoadImage_Error
'          GoTo LoadImage_ErrorExit
'       Case 3 'User chose Abort
'          On Error GoTo 0
'          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
'    End Select
'
'End Function

'=======================================================================
'Function:              IsFile
'
'Synopsis:              Tests the past in string to see if it
'                       is an existing file.
'=======================================================================
Private Function IsFile(ByVal sFile As String) As Boolean
   On Error GoTo errIsFile
   IsFile = False
   If Len(Trim$(sFile)) = 0 Then
      IsFile = False
      Exit Function
   End If
   If Len(Dir(sFile, vbNormal)) > 0 Then
      IsFile = True
   End If
   Exit Function
errIsFile:
   IsFile = False
End Function

Private Function IncrementInputFile() As Boolean
    If m_lInputFilePtr < m_lInputFileIdx Then
        m_lInputFilePtr = m_lInputFilePtr + 1
        m_InputFileFileRecordCount = 0
        IncrementInputFile = True
    Else
        IncrementInputFile = False
    End If
End Function

Private Function PDIOMR(ByVal lSheetNumber As Long, Optional ByVal sSelInstFlags As String) As String
    Dim sBarCode As String
    Dim lSeq7 As Long
    sBarCode = "YN~~~PPP" 'Note: The Wizard will set the end-of-piece (second bar) to Y on the last sheet.
    lSeq7 = lSheetNumber Mod 7
    If lSeq7 = 0 Then
        lSeq7 = 7
    End If
    Mid(sBarCode, 6) = IIf((lSeq7 And 1) > 0, "Y", "N")
    Mid(sBarCode, 7) = IIf((lSeq7 And 2) > 0, "Y", "N")
    Mid(sBarCode, 8) = IIf((lSeq7 And 4) > 0, "Y", "N")
    PDIOMR = "%FDSBarcode: PDI1 " & PsString(sBarCode & sSelInstFlags, True) & " OMRPDI1"
End Function

Private Function Nationwide_LN0480_OCRLine(ByVal sAccountNumber As String, ByVal cAmountDue As Currency) As String
    Const lcAcctNumLen As Long = 12
    Const lcAmountLen As Long = 10
    Const lsLockBoxNum As String = "91425"

    Dim sAmountDue As String
    Dim sOCRSegment As String
    Dim lCheckDigit As Long
    
    If cAmountDue < 0 Then
        cAmountDue = 0
    End If
    sAccountNumber = Right(String(lcAcctNumLen, "0") & Trim(sAccountNumber), lcAcctNumLen)
    sAmountDue = Left(Format(cAmountDue * 100, String(lcAmountLen, "0")), lcAmountLen)
    
    sOCRSegment = lsLockBoxNum & sAccountNumber & sAmountDue
    lCheckDigit = Nationwide_GetCheckDigit(sOCRSegment)
    
    Nationwide_LN0480_OCRLine = sOCRSegment & CStr(lCheckDigit)
End Function

Public Function Nationwide_GetCheckDigit(ByVal sOCRSegment As String, Optional ByVal bSumOfValues As Boolean = False) As Long
    Dim i As Long
    Dim lDigit As Long
    Dim lSum As Long
    Dim lRem As Long
    
    Dim bDblToggle As Boolean
    Dim lWeightedValue As Long
    Dim sWeightedDigits As String
    
    Dim sValues As String
    Dim sSum As String
    
    Dim lCheckDigit As Long
    
    Const bDebug As Boolean = False
    
    Nationwide_GetCheckDigit = 0
    sOCRSegment = Replace(sOCRSegment, " ", vbNullString)
    
    If Len(sOCRSegment) = 0 Then
        Exit Function
    End If
    
    If bDebug Then
        Debug.Print
        Debug.Print "Double-Add-Double, " & IIf(bSumOfValues, "Sum Of Values", "Sum Of Digits") & ", Mod 10"
        For i = 1 To Len(sOCRSegment)
            Debug.Print "  " & Mid(sOCRSegment, i, 1);
        Next
        Debug.Print "  <--- OCR Segment"
        sValues = vbNullString
        bDblToggle = True
        For i = 1 To Len(sOCRSegment)
            sValues = sValues & IIf(bDblToggle, "  2", "  1")
            bDblToggle = Not bDblToggle
        Next
         Debug.Print sValues & "  <--- Wieghts"
    End If
    
    bDblToggle = True
    sValues = vbNullString
    For i = 1 To Len(sOCRSegment)
        lWeightedValue = CLng(Mid(sOCRSegment, i, 1))
        If bDblToggle Then
            lWeightedValue = lWeightedValue * 2
        End If
        sWeightedDigits = Format(lWeightedValue, "00")
        sValues = sValues & " " & sWeightedDigits
        bDblToggle = Not bDblToggle
        If bSumOfValues Then
            lSum = lSum + lWeightedValue
            sSum = sSum & "+" & CStr(lWeightedValue)
        Else
            lDigit = CLng(Mid(sWeightedDigits, 2, 1))
            If lDigit > 0 Then
                lSum = lSum + lDigit
                sSum = sSum & "+" & CStr(lDigit)
            End If
            lDigit = CLng(Mid(sWeightedDigits, 1, 1))
            If lDigit > 0 Then
                lSum = lSum + lDigit
                sSum = sSum & "+" & CStr(lDigit)
            End If
        End If
    Next
    
    If bDebug Then
        Debug.Print sValues & "  <--- " & IIf(bSumOfValues, "Values", "Digits")
    End If
    
    If lSum > 0 Then
        lRem = lSum Mod 10
        'If lRem > 0 Then
        '    lRem = 10 - lRem
        'End If
    End If
    
    Nationwide_GetCheckDigit = lRem
    
    If bDebug Then
        Debug.Print Mid(sSum, 2) & "=" & CStr(lSum) & "  <--- " & "Sum"
        Debug.Print CStr(Nationwide_GetCheckDigit) & "  <--- Check Digit"
    End If
    
End Function

Private Function GetDeliveryAddress(ByRef saPageLines() As String, ByVal lPageLineIdx As Long, ByVal lAddressTopLine As Long, ByVal lAddressLeftColumn As Long, ByVal lAddressLen As Long, Optional ByVal bScrape As Boolean, Optional ByVal bNoRightMargin As Boolean, Optional ByVal bNoBottomMargin As Boolean, Optional ByVal bNoTopMargin As Boolean) As CAddress
    
    On Error GoTo GetDeliveryAddress_Error
    
    'The address 'area' is the address block plus a one-space buffer above, below and to the left
    'and right of the actual address block.   The handling code is located in the on address line 1,
    'beyond the 30th character and is looked for by a single character preceeded by 2 or more spaces.
    Dim lAddrAreaCol As Long
    Dim lAddrAreaWidth As Long
    Dim lAddrAreaTopIdx As Long
    Dim lAddrAreaBottomIdx As Long
    
    'The address 'block' is the actual address lines not including the surrounding buffer.
    Dim lAddrBlockTopIdx As Long
    Dim lAddrBlockBottomIdx As Long
    
    Dim sAddressAreaLine As String
    
    Dim sArtificialLeftMargin As String
    
    lAddrAreaCol = lAddressLeftColumn - 1
    If bNoRightMargin Then
        lAddrAreaWidth = lAddressLen + 1
    Else
        lAddrAreaWidth = lAddressLen + 2
    End If
    sArtificialLeftMargin = vbNullString
    If lAddrAreaCol < 0 Then
        Err.Raise 7, App.Title, "Invalid address left column argument."
    ElseIf lAddrAreaCol = 0 Then
        sArtificialLeftMargin = " "
        lAddrAreaCol = lAddrAreaCol + 1
        lAddrAreaWidth = lAddrAreaWidth - 1
    End If
    
    If bNoTopMargin Then
        lAddrAreaTopIdx = lAddressTopLine - 1 ' Point to the first line of the 6-line address block.
    Else
        lAddrAreaTopIdx = lAddressTopLine - 2 ' Point to page line immediately above the 6-line address block.
    End If
    
    If bNoBottomMargin Then
        lAddrAreaBottomIdx = lAddressTopLine + 4 ' Point to the last line of the 6-line address block.
    Else
        lAddrAreaBottomIdx = lAddressTopLine + 5 ' Point to page line immediately below the 6-line address block.
    End If
    
    If lAddrAreaTopIdx < 0 Then
        Err.Raise 7, App.Title, "Invalid address left column argument."
    ElseIf lAddrAreaTopIdx = 0 Then
        lAddrAreaTopIdx = lAddrAreaTopIdx + 1
    End If
    
    If lAddrAreaBottomIdx > lPageLineIdx Then
        lAddrAreaBottomIdx = lPageLineIdx
    End If
    
    lAddrBlockTopIdx = lAddressTopLine - 1
    lAddrBlockBottomIdx = lAddrBlockTopIdx + 5
    
    Set GetDeliveryAddress = New CAddress
    GetDeliveryAddress.Alignment = fsaAlignBottom
    
    Dim i As Long
    Dim lAddressLine As Long
    For i = lAddrAreaTopIdx To lAddrAreaBottomIdx
        If i > lPageLineIdx Then
            Exit For
        End If
        If bScrape Then
            sAddressAreaLine = sArtificialLeftMargin & ScrapeString(saPageLines(i), lAddrAreaCol, lAddrAreaWidth, PvRTrimOutput)
        Else
            sAddressAreaLine = sArtificialLeftMargin & RTrim(Mid(saPageLines(i), lAddrAreaCol, lAddrAreaWidth))
        End If
        If i < lAddrBlockTopIdx Then
            ' Above the address block.
            If Len(Trim(sAddressAreaLine)) > 0 Then
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        ElseIf i > lAddrBlockBottomIdx Then
            ' Below the address block.
            If Len(Trim(sAddressAreaLine)) > 0 Then
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        Else
            ' On an address block line.
            lAddressLine = lAddressLine + 1
            GetDeliveryAddress.AddressLine(lAddressLine) = ScrapeString(sAddressAreaLine, 2, lAddressLen, PvRTrimOutput)
            If Len(Trim(sAddressAreaLine)) Then
                'The one-space buffer to the left or five-space buffer to the right of address block is violated.
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        End If
    Next
    
    Exit Function

GetDeliveryAddress_Error:

   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpContext = Err.HelpContext
   m_sHelpFile = Err.HelpFile

   Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
      Case 1 'User chose Retry
         Resume
      Case 2 'User chose Ignore
         Resume Next
      Case 3 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
   End Select

   Exit Function

End Function
Private Function GetReturnAddress(ByRef saPageLines() As String, ByVal lPageLineIdx As Long, ByRef lAddressTopLine As Long, ByRef lAddressLineCount As Long, ByRef lAddressLeftColumn As Long, ByRef lAddressLen As Long, ByVal bScrape As Boolean, ByVal bSqueeze As Boolean) As CAddress
    Dim lAddressLine As Long
    Dim i As Long
    Set GetReturnAddress = New CAddress
    Dim lOutAddressLine As Long
    For lAddressLine = 1 To lAddressLineCount
        i = lAddressLine + lAddressTopLine - 2
        If i > lPageLineIdx Then
            Exit For
        End If
        Dim sLine As String
        If bScrape Then
            sLine = ScrapeString(saPageLines(i), lAddressLeftColumn, lAddressLen, PvRTrimOutput)
        Else
            sLine = RTrim(Mid(saPageLines(i), lAddressLeftColumn, lAddressLen))
        End If
        
        If bSqueeze Then
            If Len(sLine) > 0 Then
                lOutAddressLine = lOutAddressLine + 1
                GetReturnAddress.AddressLine(lOutAddressLine) = sLine
            End If
        Else
            lOutAddressLine = lOutAddressLine + 1
            GetReturnAddress.AddressLine(lOutAddressLine) = sLine
        End If
    Next
End Function

Private Function GetDeliveryAddress2(ByRef saPageLines() As String, ByVal lPageLineIdx As Long, ByVal lAddressTopLine As Long, ByVal lAddressLeftColumn As Long, ByVal lAddressLen As Long, ByVal lAddrHeight As Long, Optional ByVal bScrape As Boolean, Optional ByVal bNoRightMargin As Boolean, Optional ByVal bNoBottomMargin As Boolean, Optional ByVal bNoTopMargin As Boolean) As CAddress
    
    On Error GoTo GetDeliveryAddress2_Error
    
    'The address 'area' is the address block plus a one-space buffer above, below and to the left
    'and right of the actual address block.   The handling code is located in the on address line 1,
    'beyond the 30th character and is looked for by a single character preceeded by 2 or more spaces.
    Dim lAddrAreaCol As Long
    Dim lAddrAreaWidth As Long
    Dim lAddrAreaTopIdx As Long
    Dim lAddrAreaBottomIdx As Long
    
    'The address 'block' is the actual address lines not including the surrounding buffer.
    Dim lAddrBlockTopIdx As Long
    Dim lAddrBlockBottomIdx As Long
    
    Dim sAddressAreaLine As String
    
    Dim sArtificialLeftMargin As String
    
    lAddrAreaCol = lAddressLeftColumn - 1
    If bNoRightMargin Then
        lAddrAreaWidth = lAddressLen + 1
    Else
        lAddrAreaWidth = lAddressLen + 2
    End If
    sArtificialLeftMargin = vbNullString
    If lAddrAreaCol < 0 Then
        Err.Raise 7, App.Title, "Invalid address left column argument."
    ElseIf lAddrAreaCol = 0 Then
        sArtificialLeftMargin = " "
        lAddrAreaCol = lAddrAreaCol + 1
        lAddrAreaWidth = lAddrAreaWidth - 1
    End If
    
    If bNoTopMargin Then
        lAddrAreaTopIdx = lAddressTopLine - 1 ' Point to the first line of the 6-line address block.
    Else
        lAddrAreaTopIdx = lAddressTopLine - 2 ' Point to page line immediately above the 6-line address block.
    End If
    
    If bNoBottomMargin Then
        lAddrAreaBottomIdx = lAddressTopLine + (lAddrHeight - 2) ' Point to the last line of the 6-line address block.
    Else
        lAddrAreaBottomIdx = lAddressTopLine + (lAddrHeight - 1) ' Point to page line immediately below the 6-line address block.
    End If
    
    If lAddrAreaTopIdx < 0 Then
        Err.Raise 7, App.Title, "Invalid address left column argument."
    ElseIf lAddrAreaTopIdx = 0 Then
        lAddrAreaTopIdx = lAddrAreaTopIdx + 1
    End If
    
    If lAddrAreaBottomIdx > lPageLineIdx Then
        lAddrAreaBottomIdx = lPageLineIdx
    End If
    
    lAddrBlockTopIdx = lAddressTopLine - 1
    lAddrBlockBottomIdx = lAddrBlockTopIdx + (lAddrHeight - 1)
    
    Set GetDeliveryAddress2 = New CAddress
    GetDeliveryAddress2.Alignment = fsaAlignBottom
    
    Dim i As Long
    Dim lAddressLine As Long
    For i = lAddrAreaTopIdx To lAddrAreaBottomIdx
        If i > lPageLineIdx Then
            Exit For
        End If
        If bScrape Then
            sAddressAreaLine = sArtificialLeftMargin & ScrapeString(saPageLines(i), lAddrAreaCol, lAddrAreaWidth, PvRTrimOutput)
        Else
            sAddressAreaLine = sArtificialLeftMargin & RTrim(Mid(saPageLines(i), lAddrAreaCol, lAddrAreaWidth))
        End If
        If i < lAddrBlockTopIdx Then
            ' Above the address block.
            If Len(Trim(sAddressAreaLine)) > 0 Then
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        ElseIf i > lAddrBlockBottomIdx Then
            ' Below the address block.
            If Len(Trim(sAddressAreaLine)) > 0 Then
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        Else
            ' On an address block line.
            lAddressLine = lAddressLine + 1
            GetDeliveryAddress2.AddressLine(lAddressLine) = ScrapeString(sAddressAreaLine, 2, lAddressLen, PvRTrimOutput)
            If Len(Trim(sAddressAreaLine)) Then
                'The one-space buffer to the left or five-space buffer to the right of address block is violated.
                Err.Raise 7, App.Title, "Address block validation error."
            End If
        End If
    Next
    
    Exit Function

GetDeliveryAddress2_Error:

   'Preserve error context
   m_lErr = Err.Number
   m_sErrSrc = Err.source
   m_sErrDesc = Err.Description
   m_sHelpContext = Err.HelpContext
   m_sHelpFile = Err.HelpFile

   Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext)
      Case 1 'User chose Retry
         Resume
      Case 2 'User chose Ignore
         Resume Next
      Case 3 'User chose Abort
         Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
   End Select

   Exit Function

End Function

Private Function AddUSA(ByRef oAddress As CAddress) As CAddress
    Dim i As Long
    Set AddUSA = New CAddress
    For i = 1 To oAddress.CondensedAddressLines
        AddUSA.AddressLine(i) = oAddress.CondensedAddress(i)
    Next
    If i <= 6 Then
        AddUSA.AddressLine(i) = "UNITED STATES OF AMERICA"
    End If
    AddUSA.Alignment = oAddress.Alignment
End Function

Private Function CreateFileHeader(ByVal sReportId As String, ByVal sFormType As String) As String
    Dim sNewHeader As String
    sNewHeader = Space(80)
    Mid(sNewHeader, 1) = "FILE NAME-"
    Mid(sNewHeader, 11) = Left(Trim(sReportId), 13)
    Mid(sNewHeader, 25) = "PAGES-      "
    Mid(sNewHeader, 40) = "FORM TYPE-"
    Mid(sNewHeader, 50) = Left(Trim(sFormType), 9)
    CreateFileHeader = sNewHeader
End Function

Private Sub ReadPageAhead()
    Dim oInputLine As PrintStreamLine
    
    If Not m_bPrintStreamPrimed Then
        m_bPrintStreamPrimed = True
        Erase m_saCurrentPageLine
        m_lCurrentPageIdx = -1
        If m_oPrintStream.EOF Then
            'Empty file.
            m_bPrintStreamAtEnd = True
        Else
            m_oPrintStream.ReadLine
        End If
        If Len(m_sXmlReportId) > 0 Then
            m_lReadAheadPageIdx = m_lReadAheadPageIdx + 1
            ReDim Preserve m_saReadAheadPageLine(m_lReadAheadPageIdx)
            ReDim Preserve m_saReadAheadPageLineBold(m_lReadAheadPageIdx)
            'ReDim Preserve m_baReadAheadPageLineBold(m_lReadAheadPageIdx)
            m_saReadAheadPageLine(m_lReadAheadPageIdx) = CreateFileHeader(m_sXmlReportId, m_sXmlFormType)
            'm_baReadAheadPageLineBold(m_lReadAheadPageIdx) = False
            m_saReadAheadPageLineBold(m_lReadAheadPageIdx) = vbNullString
            Exit Sub
        End If
    Else
        m_saCurrentPageLine = m_saReadAheadPageLine
        m_saCurrentPageLineBold = m_saReadAheadPageLineBold
        m_lCurrentPageIdx = m_lReadAheadPageIdx
    End If
    
    Erase m_saReadAheadPageLine
    'Erase m_baReadAheadPageLineBold
    Erase m_saReadAheadPageLineBold
    m_lReadAheadPageIdx = -1
    'm_bReadAheadPageIsFileNameRecord = False
    
    If m_bPrintStreamAtEnd Then
        Exit Sub
    End If
    
    With m_oPrintStream
    
        '-----------------------------------------------
        'Ensure we're starting at the top-of-form.
        'Set oInputLine = m_oPrintStream.Line
        ' Do Until m_oPrintStream.Line.NewForm Or Left(m_oPrintStream.Line.Data, Len(m_scFileName)) = m_scFileName
        Do Until m_oPrintStream.Line.NewForm Or InStr(m_oPrintStream.Line.Data, m_scFileName)
            If m_oPrintStream.EOF Then
                m_bPrintStreamAtEnd = True
                Exit Sub
            End If
            m_oPrintStream.ReadLine
        Loop
        '-----------------------------------------------
        
        '-----------------------------------------------
        'Exit if at FILE-NAME record
        ' If Left(m_oPrintStream.Line.Data, Len(m_scFileName)) = m_scFileName Then
        If InStr(m_oPrintStream.Line.Data, m_scFileName) Then
            m_lReadAheadPageIdx = m_lReadAheadPageIdx + 1
            ReDim Preserve m_saReadAheadPageLine(m_lReadAheadPageIdx)
            ReDim Preserve m_saReadAheadPageLineBold(m_lReadAheadPageIdx)
            m_saReadAheadPageLine(m_lReadAheadPageIdx) = RTrim(m_oPrintStream.Line.Data)
            m_saReadAheadPageLineBold(m_lReadAheadPageIdx) = vbNullString
            m_oPrintStream.ReadLine
            Exit Sub
        End If
        
        '-----------------------------------------------
        
        Dim sNewChar As String
        Dim sExistingChar As String
        Dim lCurrentLine As Long
        Dim sRecord As String
        Dim i As Long
    
        Do
          
            lCurrentLine = m_oPrintStream.Line.LineNumber
            sRecord = RTrim(m_oPrintStream.Line.Data)
            
'            Debug.Print CStr(oInputLine.lineNumber) & oInputLine.Data
            
'            If Len(sRecord) > 0 Then
                If lCurrentLine - 1 <= m_lReadAheadPageIdx Then
                    If Len(m_saReadAheadPageLine(lCurrentLine - 1)) = 0 Then
                        m_saReadAheadPageLine(lCurrentLine - 1) = sRecord
                    Else
                        'Overstrike new line data onto existing line data.
                        If Len(m_saReadAheadPageLine(lCurrentLine - 1)) < Len(sRecord) Then
                            'Make sure the existing record is at least as long as new record.
                            m_saReadAheadPageLine(lCurrentLine - 1) = Left(m_saReadAheadPageLine(lCurrentLine - 1) & Space(Len(sRecord)), Len(sRecord))
                        End If
                        'm_saReadAheadPageLineBold(lCurrentLine - 1) = True
                        For i = 1 To Len(sRecord)
                            'Go through new record and overstrike characters onto the existing record.
                            'Keep track of bold characters (e.g. a B overstrike on top of existing B).
                            sNewChar = Trim(Mid(sRecord, i, 1))
                            If Len(sNewChar) > 0 Then
                                sExistingChar = Trim(Mid(m_saReadAheadPageLine(lCurrentLine - 1), i, 1))
                                If Len(sExistingChar) = 0 Then
                                    'The base-line character position is a space.  Insert this non-space character there.
                                    Mid(m_saReadAheadPageLine(lCurrentLine - 1), i) = sNewChar
                                Else
                                    'The base-line character position is not a space so we can't just insert the new non-space character there.
                                    If sExistingChar = sNewChar Then
                                        'If the new character matches the base line character, create a bold character.
                                        Dim sExistingBoldChar As String
                                        sExistingBoldChar = Trim(Mid(m_saReadAheadPageLineBold(lCurrentLine - 1), i, 1))
                                        If Len(sExistingBoldChar) > 0 Then
                                            If sExistingBoldChar <> sNewChar Then
                                                Stop 'Shouldn't get here.
                                            End If
                                        End If
                                        If Len(m_saReadAheadPageLineBold(lCurrentLine - 1)) < Len(m_saReadAheadPageLine(lCurrentLine - 1)) Then
                                            'Make sure the bold-char record is at least as long as the base-line record.
                                            m_saReadAheadPageLineBold(lCurrentLine - 1) = Left(m_saReadAheadPageLineBold(lCurrentLine - 1) & Space(Len(m_saReadAheadPageLine(lCurrentLine - 1))), Len(m_saReadAheadPageLine(lCurrentLine - 1)))
                                        End If
                                        'Insert the character into the bold-line character position.
                                        Mid(m_saReadAheadPageLineBold(lCurrentLine - 1), i) = sNewChar
                                    Else
                                        If Not (sExistingChar = "-" And sNewChar = ".") Then
                                            'The base-line and new characters are both non-blank and not equal.
                                            'This condition is currently not supported.
                                            Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page.  Illegal overstrike data - " & sRecord & "'."
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Else
                    If m_lReadAheadPageIdx > 200 Then
                        Err.Raise g_lcErrBase + g_lcErrInvalidPage, App.Title, "Invalid page.  Too many lines."
                    End If
                    m_lReadAheadPageIdx = lCurrentLine - 1
                    ReDim Preserve m_saReadAheadPageLine(m_lReadAheadPageIdx)
                    ReDim Preserve m_saReadAheadPageLineBold(m_lReadAheadPageIdx)
                    m_saReadAheadPageLine(m_lReadAheadPageIdx) = sRecord
                    'm_saReadAheadPageLineBold(m_lReadAheadPageIdx) = False
                    m_saReadAheadPageLineBold(m_lReadAheadPageIdx) = vbNullString
                End If
'            End If
            
            If m_oPrintStream.EOF Then
                m_bPrintStreamAtEnd = True
                Exit Do
            Else
                m_oPrintStream.ReadLine
                ' TF - Redundant, sRecord is re-assigned at top of loop
'                sRecord = m_oPrintStream.Line.Data
'                If m_oPrintStream.Line.Hidden Then
'                    sRecord = vbNullString
'                End If
            End If
        
        ' Loop Until m_oPrintStream.Line.NewForm Or Left(m_oPrintStream.Line.Data, Len(m_scFileName)) = m_scFileName
        Loop Until m_oPrintStream.Line.NewForm Or InStr(m_oPrintStream.Line.Data, m_scFileName)
        
        'Erase base line characters that have a corresponding bold-line character.
        Dim c As Long
        For i = 0 To m_lReadAheadPageIdx
            If Len(m_saReadAheadPageLineBold(i)) > 0 Then
                For c = 1 To Len(m_saReadAheadPageLineBold(i))
                    If Len(Trim(Mid(m_saReadAheadPageLineBold(i), c, 1))) > 0 Then
                        Mid(m_saReadAheadPageLine(i), c) = " "
                    End If
                Next
            End If
        Next

    End With
    
End Sub

Private Function UnrecognizedErrorReqd(ByVal sFileName As String, ByVal sFormType As String)
    Dim i As Long
    For i = 0 To m_lUnrecognizedFileIdx
        If sFileName = m_saUnrecognizedFileName(i) Then
            If sFormType = m_saUnrecognizedFormType(i) Then
                UnrecognizedErrorReqd = False
                Exit Function
            End If
        End If
    Next
    m_lUnrecognizedFileIdx = m_lUnrecognizedFileIdx + 1
    ReDim Preserve m_saUnrecognizedFileName(m_lUnrecognizedFileIdx)
    ReDim Preserve m_saUnrecognizedFormType(m_lUnrecognizedFileIdx)
    m_saUnrecognizedFileName(m_lUnrecognizedFileIdx) = sFileName
    m_saUnrecognizedFormType(m_lUnrecognizedFileIdx) = sFormType
    UnrecognizedErrorReqd = True
    Exit Function
End Function

Private Function GetBulkShipIndex(ByVal sFileName As String, ByVal sFormType As String, Optional sSrfKey As String)
    Dim i As Long
    For i = 0 To m_lBulkShipFileIdx
        If sFileName = m_saBulkShipFileName(i) Then
            If sFormType = m_saBulkShipFormType(i) Then
                GetBulkShipIndex = i + 1000
                Exit Function
            End If
        End If
    Next
    m_lBulkShipFileIdx = m_lBulkShipFileIdx + 1
    ReDim Preserve m_saBulkShipFileName(m_lBulkShipFileIdx)
    ReDim Preserve m_saBulkShipFormType(m_lBulkShipFileIdx)
    m_saBulkShipFileName(m_lBulkShipFileIdx) = sFileName
    m_saBulkShipFormType(m_lBulkShipFileIdx) = sFormType
    GetBulkShipIndex = m_lBulkShipFileIdx + 1000
    SetupBulkShipStack GetBulkShipIndex, sFileName, sFormType, sSrfKey
    Exit Function
End Function
Private Function TrimLeadingBlanks(ByRef saData() As String) As Long
    Dim i As Integer
    Dim iDataStart As Integer
    iDataStart = 0
    'Find the first non blank element
    For i = 0 To UBound(saData)
        If Len(Trim(saData(i))) > 0 Then
            iDataStart = i
            Exit For
        End If
    Next
    
    'Starting at the first element that contains data, copy to a new array
    Dim saCopy() As String
    Dim lCopyIndex As Integer
    lCopyIndex = 0
    For i = iDataStart To UBound(saData)
        ReDim Preserve saCopy(lCopyIndex)
        saCopy(lCopyIndex) = saData(i)
        lCopyIndex = lCopyIndex + 1
    Next
    
    'lCopyIndex is incremented by one too many in loop
    lCopyIndex = lCopyIndex - 1
    
    'Copy the new array over the old array
    CopyArray saCopy, saData
    
    TrimLeadingBlanks = lCopyIndex
End Function
Private Sub CopyArray(ByRef saSource() As String, ByRef saTarget() As String)
    If Not IsArray(saSource) And Not IsArray(saTarget) Then
        Exit Sub
    End If
    
    Dim iSourceLBound As Integer
    Dim iSourceUBound As Integer
    
    iSourceLBound = LBound(saSource)
    iSourceUBound = UBound(saSource)
    
    ReDim saTarget(iSourceLBound To iSourceUBound)
    
    Dim i As Integer
    For i = iSourceLBound To iSourceUBound
        saTarget(i) = saSource(i)
    Next
End Sub

'Boston private OCR scan line
Private Function OCRLine(ByVal sAccountNumber As String, ByVal cAmount As Currency) As String
    Dim sAcntNbr As String
    Dim sPaymentAmt As String
    Dim sOcrLine As String
    sAcntNbr = Right(String(10, "0") & Trim(sAccountNumber), 10)
    sPaymentAmt = Format(cAmount * 100, "000000000000")
    
    sOcrLine = sAcntNbr & " " & sPaymentAmt
    Dim lCheckDigit As Long
    lCheckDigit = GetCheckDigit371(sOcrLine)
    OCRLine = sOcrLine & " " & CStr(lCheckDigit)
End Function

Public Function GetCheckDigit371(ByVal sOCRSegment As String) As Long
    Dim i As Long
    Dim lValue As Long
    Dim lWeight As Long
    Dim lSum As Long
    Dim lRem As Long
    
    GetCheckDigit371 = 0
    
    lWeight = 3
    For i = Len(sOCRSegment) To 1 Step -1
        Dim sDigit As String
        sDigit = Trim(Mid(sOCRSegment, i, 1))
        If Len(sDigit) > 0 Then
            If lWeight = 3 Then
                lWeight = 1
            ElseIf lWeight = 1 Then
                lWeight = 7
            ElseIf lWeight = 7 Then
                lWeight = 3
            End If
            lSum = lSum + CLng(sDigit) * lWeight
        End If
    Next
    
    If lSum > 0 Then
        GetCheckDigit371 = lSum Mod 10
    End If
End Function

Private Function ReformatTM0330P1() As Boolean

    Dim tempBool As Boolean
    Dim currLine As Long
    Dim moveToLine As Long
    Dim offset As Long
    Dim currLineLen As Long
    Dim currLineOffset As Long
    
    Const lcLineLen = 79
    
    On Error GoTo ReformatTM0330P1_Error
    
    m_saCurrentPageLine(0) = vbNullString
    m_saCurrentPageLine(1) = vbNullString
    m_saCurrentPageLine(2) = vbNullString
    
    For currLine = 4 To 7
        currLineLen = 48
        m_saCurrentPageLine(currLine) = Right(Space(lcLineLen) & Trim(m_saCurrentPageLine(currLine)), lcLineLen)
    Next
    
    ReformatTM0330P1 = True
    
    Exit Function
    
ReformatTM0330P1_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.ReformatTM0330P1"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo ReformatTM0330P1_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo ReformatTM0330P1_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Private Function Reformat_BPB_LN0470() As Boolean

'    Dim tempBool As Boolean
'    Dim currLine As Long
'    Dim moveToLine As Long
'    Dim offset As Long
'    Dim currLineLen As Long
'    Dim currLineOffset As Long
'
'    Const lcLineLen = 79
    
    On Error GoTo Reformat_BPB_LN0470_Error
    
    m_lCurrentPageIdx = m_lCurrentPageIdx + 6
    ReDim Preserve m_saCurrentPageLine(m_lCurrentPageIdx)
    ReDim Preserve m_saCurrentPageLineBold(m_lCurrentPageIdx)
    
    m_saCurrentPageLine(m_lCurrentPageIdx - 3) = "      Please contact us at 617-912-3924, 617-912-3902, or 855-211-7781, or"
    m_saCurrentPageLine(m_lCurrentPageIdx - 2) = "      at SVB Private,A Division of First Citizens Bank, Loan Operations Department," 'CPTPK-3026
    m_saCurrentPageLine(m_lCurrentPageIdx - 1) = "      53 state street, Floor 28, Boston, MA 02109. Our office is open Monday "
    m_saCurrentPageLine(m_lCurrentPageIdx) = "      - Friday from 8:30am to 5:00pm, Eastern Time."
    
    Reformat_BPB_LN0470 = True
    
    Exit Function
    
Reformat_BPB_LN0470_Error:
    
    'Preserve error context
    m_lErr = Err.Number
    m_sErrSrc = Err.source
    m_sErrDesc = Err.Description
    m_sHelpContext = Err.HelpContext
    m_sHelpFile = Err.HelpFile
    
    Const scIntro = "Error in CAdmin.Reformat_BPB_LN0470"
    
    Select Case m_oErrorHandler.HandleError(App.Title, m_lErr, m_sErrDesc, m_sErrSrc, m_sHelpFile, m_sHelpContext, scIntro)
       Case 1 'User chose Retry
          On Error GoTo Reformat_BPB_LN0470_Error
          Resume
       Case 2 'User chose Ignore
          On Error GoTo Reformat_BPB_LN0470_Error
          Resume Next
       Case 3 'User chose Abort
          On Error GoTo 0
          Err.Raise g_icErrApplicationCanceled, App.Title, g_scErrApplicationCanceled
    End Select

End Function

Public Function GetGUID() As String
    '(c) 2000 Gus Molina
    
    Dim udtGUID As GUID
    
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        "-" & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        "-" & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        "-" & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        "-" & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
    
End Function

Private Function TrimLeadingZeros(ByVal sString As String) As String
    Dim oRegExp As RegExp
    Dim oMatchCollection As MatchCollection
    
    Set oRegExp = New RegExp
    oRegExp.Pattern = "^0*([1-9][0-9]*)$"
    If oRegExp.Test(sString) = False Then
        Err.Raise g_lcErrBase + g_lcErrGenericAssertion, App.Title, "Invalid account number encountered - must be digits only."
    End If
    Set oMatchCollection = oRegExp.Execute(sString)
    
    TrimLeadingZeros = oMatchCollection.Item(0).SubMatches(0)
End Function

Private Function TrimLeadingZeros2(ByVal sString As String) As String
    ' Variables
    Dim i As Long
    Dim bNonZeroFound As Boolean
    Dim sCurChar As String
    
    ' Code
    TrimLeadingZeros2 = ""
    bNonZeroFound = False
    sCurChar = ""
    sString = LTrim(sString)
    
    For i = 1 To Len(sString)
        ' Get the current char
        sCurChar = Mid(sString, i, 1)
        
        ' Check if we have found a non-zero character
        If Not bNonZeroFound Then
            Select Case sCurChar
                Case "0" ' Ignore
                Case Else
                    bNonZeroFound = True
            End Select
        End If
        
        ' We found non-zero character, add it
        If bNonZeroFound Then
            TrimLeadingZeros2 = TrimLeadingZeros2 & sCurChar
        End If
    Next

End Function

Private Sub CondensePageArray(ByRef saLineData() As String, ByRef saBoldLineData() As String, ByRef lLineDataIdx)
    ' Declare variables
    Dim lConsecutiveBlankLineCount As Long
    Dim saTempArray() As String
    Dim saTempBoldArray() As String
    Dim lTempArrayIdx As Long
    
    ' Set initial values
    lConsecutiveBlankLineCount = 0
    Erase saTempArray
    Erase saTempBoldArray
    lTempArrayIdx = -1
    
    ' Declare code
    Dim i As Long
    For i = 0 To lLineDataIdx
        If Len(Trim(saLineData(i))) = 0 And Len(Trim(saBoldLineData(i))) = 0 Then
            If lTempArrayIdx >= 0 Then
                If lConsecutiveBlankLineCount < 2 Then
                    lConsecutiveBlankLineCount = lConsecutiveBlankLineCount + 1
                End If
            End If
        Else
            Do While lConsecutiveBlankLineCount > 0
                lTempArrayIdx = lTempArrayIdx + 1
                ReDim Preserve saTempArray(lTempArrayIdx)
                ReDim Preserve saTempBoldArray(lTempArrayIdx)
                lConsecutiveBlankLineCount = lConsecutiveBlankLineCount - 1
            Loop
            lTempArrayIdx = lTempArrayIdx + 1
            ReDim Preserve saTempArray(lTempArrayIdx)
            saTempArray(lTempArrayIdx) = saLineData(i)
            ReDim Preserve saTempBoldArray(lTempArrayIdx)
            saTempBoldArray(lTempArrayIdx) = saBoldLineData(i)
        End If
    Next
    
    Erase saLineData
    ReDim saLineData(lTempArrayIdx)
    saLineData = saTempArray
    
    Erase saBoldLineData
    ReDim saBoldLineData(lTempArrayIdx)
    saBoldLineData = saTempBoldArray
    
    lLineDataIdx = lTempArrayIdx
        
End Sub

'Begin of the Code Change : CPTPK-2513
Private Function MessageBox(Prompt, Optional Buttons = vbOKOnly, Optional Title, Optional HelpFile, Optional Context) As VbMsgBoxResult
    
    If IsMissing(Title) Then
        Title = App.Title
    End If
    If IsMissing(HelpFile) Then
        HelpFile = ""
    End If
    If IsMissing(Context) Then
        Context = 0
    End If
    
    If Not m_oInteractionProvider Is Nothing Then
        MessageBox = m_oInteractionProvider.MessageBox(Prompt, Buttons, Title, HelpFile, Context)
    Else
        MessageBox = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    End If
End Function
'End   of the Code Change : CPTPK-2513
