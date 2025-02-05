Attribute VB_Name = "modErrorHandler"
Option Explicit

Public Const g_icErrApplicationCanceled As Long = 1100
Public Const g_scErrApplicationCanceled As String = "Application canceled by user request or as the result of an error."

Public Const g_icErrOperationCanceled As Long = 1101
Public Const g_scErrOperationCanceled As String = "Operation canceled by user request or as the result of an error."

Public Enum ErrorDispatchEnum
   ErrorDispNotHandled = 0
   ErrorDispRetryResume
   ErrorDispIgnoreResumeNext
   ErrorDispAbortRaiseError
End Enum

'=======================================================================
'Function:              ErrorNumberString
'
'Synopsis:              Converts an error number (long) to a string.
'                       Errors less than zero are assumed to be an HRESULT
'                       (SCODE) and as such are displayed in hex as well as
'                       decimal to make them easier to read.
'
'Remarks:               If the error facility is FACILITY_ITF (vbObjectError or
'                       &H8004nnnn), then it is converted as a decimal value after
'                       subtracting &H80040000.  This is because with FACILITY_ITF
'                       (an error in an interface call), the low order two bytes
'                       are defined by the object.  So we display only the error
'                       value that is in the low order two bytes.  However,
'                       Microsoft states that the values 0 to &H01FF within
'                       FACILITY_ITF are used by COM so they recommend
'                       developers not use those values.  Therefore, we'll
'                       still display values in that range in Hex with
'                       the leading &H8004.
'
'=======================================================================
Public Function ErrorNumberString(ByVal lErrorNumber As Long) As String
   'Facility codes defined by Microsoft.
'   Const SEVERITY_ERROR as long = &H80000000
'   Const FACILITY_NULL As Long = &H0
'   Const FACILITY_RPC As Long = &H10000
'   Const FACILITY_DISPATCH As Long = &H20000
'   Const FACILITY_STORAGE As Long = &H30000
'   Const FACILITY_ITF As Long = &H40000 'vbobjecterror.
'   Const FACILITY_WIN32 As Long = &H70000
'   Const FACILITY_WINDOWS As Long = &H80000
'   Const FACILITY_SSPI As Long = &H90000
'   Const FACILITY_CONTROL As Long = &HA0000
'   Const FACILITY_CERT As Long = &HB0000
'   Const FACILITY_INTERNET As Long = &HC0000
   
   If lErrorNumber < 0 Then
   
      'Display in hex and decimal if < 0 unless it's a vbObjectError in the
      'range &H200 - &HFFFF.
      
      If (lErrorNumber And vbObjectError) = vbObjectError Then
         
         'vbObjectError
         
         If lErrorNumber >= vbObjectError And lErrorNumber <= vbObjectError + &H1FF Then
            
            'vbObjectError in the range used by COM.
            'Display in hex and decimal.
            ErrorNumberString = "&H" & Hex$(lErrorNumber) & " (" & CStr(lErrorNumber) & ")"
         
         Else
            
            'vbObjectError not in the range used by COM.
            'Display in decimal after subtracting vbObjectError.
            
            ErrorNumberString = CStr(lErrorNumber - vbObjectError)
         
         End If
      
      Else
         
         'HResult or SCODE.  Display in hex and decimal.
         ErrorNumberString = "&H" & Hex$(lErrorNumber) & " (" & CStr(lErrorNumber) & ")"
      
      End If
   
   Else
      
      'Display in decimal all errors greater than zero.
      
      ErrorNumberString = CStr(lErrorNumber)
   
   End If

End Function


