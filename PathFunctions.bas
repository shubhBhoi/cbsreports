Attribute VB_Name = "modPathFunctions"
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const m_lMaxPath As Long = 260
  
Private Declare Function lstrlenW Lib "kernel32" ( _
   ByVal lpString As Long) As Long
  
Private Declare Sub CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" (dest As Any, source As Any, _
   ByVal Bytes As Long)
  
Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

  
' ==== COMPACTING FUNCTIONS ========
Private Declare Function PathCompactPath Lib "Shlwapi" _
   Alias "PathCompactPathW" (ByVal hDc As Long, _
   ByVal lpszPath As Long, ByVal dx As Integer) As Boolean
  
Private Declare Function PathCompactPathEx Lib "Shlwapi" _
   Alias "PathCompactPathExW" (ByVal pszOut As Long, _
   ByVal pszSrc As Long, ByVal cchMax As Integer, _
  dwFlags As Long) As Boolean
    
Private Declare Function PathSetDlgItemPath Lib "Shlwapi" _
   Alias "PathSetDlgItemPathW" (ByVal hDlg As Long, _
   ByVal id As Long, ByVal pszPath As Long) As Boolean
                                    
' ====
Private Declare Function PathFindFileName Lib "Shlwapi" _
   Alias "PathFindFileNameW" (ByVal pPath As Long) As Long
  
Private Declare Function PathRemoveFileSpec Lib "Shlwapi" _
   Alias "PathRemoveFileSpecW" _
   (ByVal pszPath As Long) As Boolean

  
Private Declare Function PathAddBackslash Lib "Shlwapi" _
   Alias "PathAddBackslashW" _
   (ByVal lpszPath As Long) As Long
  
Private Declare Function PathRemoveBackslash _
   Lib "Shlwapi" Alias "PathRemoveBackslashW" _
   (ByVal lpszPath As Long) As Long
  
  
' ===== EXTENSION FUNCTIONS =====
Private Declare Function PathAddExtension Lib "Shlwapi" _
   Alias "PathAddExtensionW" (ByVal lpszPath As Long, _
   ByVal pszExtension As Long) As Boolean
  
Private Declare Sub PathRemoveExtension Lib "Shlwapi" _
   Alias "PathRemoveExtensionW" (ByVal lpszPath As Long)
  
Private Declare Function PathFindExtension Lib "Shlwapi" _
   Alias "PathFindExtensionW" (ByVal pPath As Long) As Long
  
Private Declare Function PathRenameExtension _
   Lib "Shlwapi" Alias "PathRenameExtensionW" _
   (ByVal lpszPath As Long, _
   ByVal pszExtension As Long) As Boolean
  
Private Declare Function PathMatchSpec Lib "Shlwapi" _
   Alias "PathMatchSpecW" (ByVal pszFileParam As Long, _
   ByVal pszSpec As Long) As Boolean
  
Private Declare Function PathIsContentType Lib "Shlwapi" _
   Alias "PathIsContentTypeW" (ByVal pszPath As Long, _
   ByVal pszContentType As Long) As Boolean
    
  
Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public Function BrowseForFolder(ByRef sFolder As String, ByVal hWnd As Long, ByVal sTitle As String) As Boolean

    Dim lIDList As Long
    Dim tBrowseInfo As BrowseInfo
    
    BrowseForFolder = False
    
    With tBrowseInfo
        .hWndOwner = hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lIDList = SHBrowseForFolder(tBrowseInfo)

    If (lIDList) Then
       sFolder = Space(m_lMaxPath)
       SHGetPathFromIDList lIDList, sFolder
       sFolder = Left(sFolder, InStr(sFolder, vbNullChar) - 1)
       BrowseForFolder = True
    End If
    
End Function

  
Public Function CompactPathByChars( _
  Path As String, MaxChars) As String
  ' Truncates the path to fit within the specified
  ' number of characters.
  Dim pAddr As Long
  Dim strBuff As String * m_lMaxPath
  pAddr = StrPtr(strBuff)
  Call PathCompactPathEx( _
          pAddr, StrPtr(Path), MaxChars + 1, 0)
  CompactPathByChars = Ptr2StrU(pAddr)
End Function
  
Public Sub CompactPathIntoLabel(ByVal sPath As String, oForm As Form, oLabel As Label)
  ' Truncates the sPath to fit within the specified
  ' label control.
  
  Dim pAddr As Long
  Dim lMaxPixels As Long
  
  pAddr = StrPtr(sPath)
  lMaxPixels = oLabel.Width \ Screen.TwipsPerPixelX
  
  Call PathCompactPath(oForm.hDc, pAddr, lMaxPixels)
  
  oLabel.Caption = Ptr2StrU(pAddr)

End Sub
  
Public Sub CompactPathIntoTextBox(ByVal sPath As String, oForm As Form, oTextBox As TextBox)
  ' Truncates the sPath to fit within the specified
  ' textbox control.
  
  Dim pAddr As Long
  Dim lMaxPixels As Long
  
  pAddr = StrPtr(sPath)
  lMaxPixels = oTextBox.Width \ Screen.TwipsPerPixelX
  
  Call PathCompactPath(oForm.hDc, pAddr, lMaxPixels)
  
  oTextBox.Text = Ptr2StrU(pAddr)

End Sub
    
Public Sub CompactPathDlgCrl(ByVal Path As String, hDlg As Long, id As Long)
    ' Truncates the path to fit within the available space of
    ' a dialog box control, and sets the text of the control.
    Dim pAddr As Long
    pAddr = StrPtr(Path)
    Call PathSetDlgItemPath(hDlg, id, pAddr)
End Sub

  
' ================================
' Extracting file path components.
' ================================
Public Function GetFile(ByVal Path As String) As String
  ' Returns the file naGetFileme from a path name.
  GetFile = Ptr2StrU(PathFindFileName(StrPtr(Path)))
End Function
  
Public Function GetFolder(ByVal Path As String) As String
  ' Returns the folder name from a path name.
  Dim pAddr As Long
  pAddr = StrPtr(Path)
  Call PathRemoveFileSpec(pAddr)
  GetFolder = Ptr2StrU(pAddr)
End Function
  
Public Function GetExtension(ByVal Path As String) _
   As String
  
  ' Returns the extension name from a file name.
  GetExtension = Ptr2StrU(PathFindExtension(StrPtr(Path)))
End Function

Public Function Ptr2StrU(ByVal pAddr As Long) As String
  Dim lAddr As Long
  lAddr = lstrlenW(pAddr)
  Ptr2StrU = Space(lAddr)
  CopyMemory ByVal StrPtr(Ptr2StrU), ByVal pAddr, lAddr * 2
End Function

  
' =================================================
' Formatting and modifying file and folder strings.
' =================================================
Public Function AddBackslash(ByVal Path As String) As String
    ' Adds a final backslash to the path is there
    ' is no backslash.
    Dim pAddr As Long
    pAddr = StrPtr(PadBuffer(Path))
    Call PathAddBackslash(pAddr)
    AddBackslash = Ptr2StrU(pAddr)
End Function
  
Public Function RemoveBackslash(ByVal Path As String) As String
    ' Removes a final backslash from the path
    ' if there is one.
    Dim pAddr As Long
    pAddr = StrPtr(Path)
    Call PathRemoveBackslash(pAddr)
    RemoveBackslash = Ptr2StrU(pAddr)
End Function
  
Public Function AddExtension(ByVal Path As String, Extension As String) As String
    ' Adds the specified extension to the path
    ' if there is no extension.
    Dim pAddr As Long
    Call QualifyExtension(Extension)
    pAddr = StrPtr(PadBuffer(Path))
    Call PathAddExtension(pAddr, StrPtr(Extension))
    AddExtension = Ptr2StrU(pAddr)
End Function
  
Public Function RemoveExtension(ByVal Path As String) As String
    ' Removes the extension from the path if there is one.
    Dim pAddr As Long
    pAddr = StrPtr(Path)
    Call PathRemoveExtension(pAddr)
    RemoveExtension = Ptr2StrU(pAddr)
End Function
  
Public Function RenameExtension(ByVal Path As String, Extension As String) As String
    ' Renames the extension of the path if there is one, or
    ' adds the extension if there is none.
    Dim pAddr As Long
    Call QualifyExtension(Extension)
    pAddr = StrPtr(PadBuffer(Path))
    Call PathRenameExtension(pAddr, StrPtr(Extension))
    RenameExtension = Ptr2StrU(pAddr)
End Function
  
Public Function AddRemoveExtension(ByVal Path As String, _
     Optional Extension As String, _
     Optional RenameIfExists As Boolean = True)
    ' Combines the three functions above. If Extension is
    ' omitted, any existing extension is removed. If
    ' RenameIfExists is True, the specified extension
    ' replaces any existing extension.
    Dim pAddr As Long
    pAddr = StrPtr(PadBuffer(Path))
    Select Case Extension
     Case vbNullString
       Call PathRemoveExtension(pAddr)
    Case Else
      QualifyExtension Extension
       If RenameIfExists = True Then
         Call PathRenameExtension(pAddr, StrPtr(Extension))
       Else
         Call PathAddExtension(pAddr, StrPtr(Extension))
       End If
    End Select
    AddRemoveExtension = Ptr2StrU(pAddr)
End Function

Private Function PadBuffer(ByVal sPath As String) As String
    PadBuffer = sPath & String(m_lMaxPath - Len(sPath), 0)
End Function
  
Private Sub QualifyExtension(Extension As String)
    If Left(Extension, 1) <> "." Then
        Extension = "." & Extension
    End If
End Sub

Public Function SplitMultiSelect(ByVal sMultiSelect As String, ByRef sFile() As String) As Long
    Dim sFolder As String
    Dim sFileName As String
    Dim i As Long
    
    SplitMultiSelect = -1
    Erase sFile
    
    sMultiSelect = Trim(sMultiSelect)
    If Len(sMultiSelect) = 0 Then
        Exit Function
    End If
'
'    If Right(sMultiSelect, 1) <> vbNullChar Then
'        sMultiSelect = sMultiSelect & vbNullChar
'    End If
    
    i = InStr(1, sMultiSelect, vbNullChar, vbTextCompare)
    
    If i > 0 Then
        'Multiple files.
        'Get the folder.
        sFolder = Left(sMultiSelect, i - 1)
        sMultiSelect = Mid(sMultiSelect, i + 1)
        'Get each file.
        Do Until Len(sMultiSelect) = 0
            i = InStr(1, sMultiSelect, vbNullChar, vbTextCompare)
            If i = 0 Then
                i = Len(sMultiSelect) + 1
            End If
            sFileName = Left(sMultiSelect, i - 1)
            sMultiSelect = Mid(sMultiSelect, i + 1)
            SplitMultiSelect = SplitMultiSelect + 1
            ReDim Preserve sFile(SplitMultiSelect)
            sFile(SplitMultiSelect) = sFolder & "\" & sFileName
        Loop
    Else
        'One file
        SplitMultiSelect = 0
        ReDim sFile(SplitMultiSelect)
        sFile(SplitMultiSelect) = sMultiSelect
    End If
        
End Function


