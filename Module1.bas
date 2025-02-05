Attribute VB_Name = "Module1"
Option Explicit

Public Const g_scTOF As String = "1"
Public Const g_lcErrBase As Long = 8000
Public Const g_lcErrInvalidPage As Long = 1
Public Const g_lcErrGenericAssertion As Long = 2
Public Const g_lcErrParameterError As Long = 3

Public Enum PullValueFlags
    PvCheckPresentValue = 1
    PvTrimOutput = 2
    PvRTrimOutput = 4
End Enum

Public Function PsString(ByVal sString As String, Optional ByVal bAddParens As Boolean) As String
    Dim i As Long
    Const scEsc As String = "\"
    Const scLeftParen As String = "("
    Const scRightParen As String = ")"
    
    i = InStr(1, sString, scEsc, vbTextCompare)
    Do Until i = 0
        sString = Left(sString, i - 1) & scEsc & Mid(sString, i)
        i = i + 2
        i = InStr(i, sString, scEsc, vbTextCompare)
    Loop
    
    i = InStr(1, sString, scLeftParen, vbTextCompare)
    Do Until i = 0
        sString = Left(sString, i - 1) & scEsc & Mid(sString, i)
        i = i + 2
        i = InStr(i, sString, scLeftParen, vbTextCompare)
    Loop
    
    i = InStr(1, sString, scRightParen, vbTextCompare)
    Do Until i = 0
        sString = Left(sString, i - 1) & scEsc & Mid(sString, i)
        i = i + 2
        i = InStr(i, sString, scRightParen, vbTextCompare)
    Loop
    
    If bAddParens Then
        PsString = "(" & sString & ")"
    Else
        PsString = sString
    End If
End Function


'==============================================================================
'Function:              ScrapeString
'
'Synopsis:              Returns a sub string of a source string, replacing
'                       the sub string with spaces within the source string.
'
'Remarks:
'==============================================================================
Public Function ScrapeString(ByRef sSourceString As String, ByVal lStart As Long, Optional ByVal lLen As Long, Optional ByVal lFlags As PullValueFlags, Optional ByVal sCheckString As String) As String

    'If optional check string <> "" then exit withouth scraping string.
    If (lFlags And PvCheckPresentValue) = PvCheckPresentValue Then
        If Len(Trim(sCheckString)) > 0 Then
            ScrapeString = sCheckString
            Exit Function
        End If
    End If
    
    'Get value of specified sub string.
    If lLen > 0 Then
        ScrapeString = Mid(sSourceString, lStart, lLen)
    Else
        ScrapeString = Mid(sSourceString, lStart)
    End If
    
    'Replace sub string with spaces.
    If Len(sSourceString) >= lStart Then
        If lLen > 0 Then
            Mid(sSourceString, lStart, lLen) = Space(lLen)
        Else
            Mid(sSourceString, lStart) = Space(Len(sSourceString))
        End If
    End If
    
    'Trim returned sub string if applicable.
    If (lFlags And PvTrimOutput) = PvTrimOutput Then
        ScrapeString = Trim(ScrapeString)
    ElseIf (lFlags And PvRTrimOutput) = PvRTrimOutput Then
        ScrapeString = RTrim(ScrapeString)
    End If
    
End Function

'==============================================================================
'Function:              ReadString
'
'Synopsis:              Returns a sub string of a source string, replacing
'                       the sub string with spaces within the source string.
'
'Remarks:
'==============================================================================
Public Function ReadString(ByRef sSourceString As String, ByVal lStart As Long, Optional ByVal lLen As Long, Optional ByVal lFlags As PullValueFlags, Optional ByVal sCheckString As String) As String

    'If optional check string <> "" then exit withouth scraping string.
    If (lFlags And PvCheckPresentValue) = PvCheckPresentValue Then
        If Len(Trim(sCheckString)) > 0 Then
            ReadString = sCheckString
            Exit Function
        End If
    End If
    
    'Get value of specified sub string.
    If lLen > 0 Then
        ReadString = Mid(sSourceString, lStart, lLen)
    Else
        ReadString = Mid(sSourceString, lStart)
    End If
    
    'Trim returned sub string if applicable.
    If (lFlags And PvTrimOutput) = PvTrimOutput Then
        ReadString = Trim(ReadString)
    ElseIf (lFlags And PvRTrimOutput) = PvRTrimOutput Then
        ReadString = RTrim(ReadString)
    End If
    
End Function

'
'Public Enum GetPathElementEnum
'    peGetDirectoryPath = 1
'    peGetFile = 2
'    peGetExt = 4
'End Enum
'
'Public Function GetPathElement(ByVal lElement As GetPathElementEnum, ByVal sPath As String) As String
'    Const scExtSep As String = "."
'    Const scPathSep As String = "\"
'    Dim lSepPos As Long
'    Dim sDirectoryPath As String
'    Dim sFile As String
'    Dim sExt As String
'
'    GetPathElement = vbNullString
'
'    sPath = Trim(sPath)
'    If Len(sPath) = 0 Then
'        Exit Function
'    End If
'
'    lSepPos = InStrRev(sPath, scPathSep, -1, vbTextCompare)
'    If lSepPos > 0 Then
'        sDirectoryPath = Mid(sPath, 1, lSepPos - 1)
'        sPath = Mid(sPath, lSepPos + 1)
'    End If
'
'    lSepPos = InStrRev(sPath, scExtSep, -1, vbTextCompare)
'    If lSepPos > 0 Then
'        sExt = Mid(sPath, lSepPos + 1)
'        sPath = Mid(sPath, 1, lSepPos - 1)
'    End If
'
'    sFile = sPath
'
'    GetPathElement = vbNullString
'    If lElement And peGetDirectoryPath = peGetDirectoryPath Then
'        GetPathElement = sDirectoryPath
'    End If
'    If lElement And peGetFile = peGetFile Then
'        If Len(GetPathElement) > 0 Then
'            If Len(sFile) > 0 Then
'                GetPathElement = GetPathElement & scPathSep
'            End If
'        End If
'        GetPathElement = GetPathElement & sFile
'    End If
'    If lElement And peGetExt = peGetExt Then
'        If Len(GetPathElement) > 0 Then
'            If Len(sExt) > 0 Then
'                GetPathElement = GetPathElement & scExtSep
'            End If
'        End If
'        GetPathElement = GetPathElement & sExt
'    End If
'
'End Function
'

Public Function Underscore(lLen As Long) As String
    Underscore = String(lLen, "_")
End Function
