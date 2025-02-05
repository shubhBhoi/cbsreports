VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddress2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================
'Class:                 CAddress
'
'Synopsis:              Given an address array of up to 6 lines, this object
'                       can parse out the following elements.

'                       Name
'                       AddressLine1
'                       StreetAddress
'                       CityStateZip
'                       CityState
'                       State
'                       ZipCode
'
'Remarks:               The CAddressUtility object's ParseName method
'                       can be used to parse the Name into it's
'                       individual elements if desired.
'
'Algorithm:             The address must be formatted as follows.  Only
'                       the order of the elements is important, not
'                       the actual AddressLine(n) they appear on.
'
'                       <Name> (optional if AddressOnly property = True)
'                       <Extra name line> (optional)
'                       <AddressLine1> (optional)
'                       <Street Address>
'                       <City><State><Zip (99999 or 99999-9999 formats)>
'
'                       Elements are pulled in this order.  If not
'                       enough elements are present then then next
'                       pulled element(s) will be considered blank.
'                       1) If AddressOnly is False then Name is
'                          the first non-blank line.
'                       2) CityStateZip is the last non-blank line.
'                       3) StreetAddress is the second to last non-blank line.
'                       4) AddressLine1 is the third to last non-blank line.
'=======================================================================
Option Explicit

Public Enum FfgStdAddressAlignmentEnum
    fsaAlignTop = 0
    fsaAlignBottom
End Enum

Private Const m_scClassName As String = "CAddress2"

Private m_sAddressLine(1 To 6) As String
Private m_bAddressOnly As Boolean
Private m_oAddressUtility As CAddressUtility
Private m_lAlignment As FfgStdAddressAlignmentEnum

Private Sub Class_Initialize()
    Set m_oAddressUtility = New CAddressUtility
End Sub

Private Sub Class_Terminate()
    Set m_oAddressUtility = Nothing
End Sub

Public Function IsEqual(ByRef oAddress As CAddress2)
    Dim i As Integer
    If oAddress Is Nothing Then
        IsEqual = False
        Exit Function
    End If
    With oAddress
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine)
            If Trim(.AddressLine(i)) <> Trim(m_sAddressLine(i)) Then
                IsEqual = False
                Exit Function
            End If
        Next
    End With
    IsEqual = True
End Function

Public Function Clone() As CAddress2
    Dim i As Integer
    Set Clone = New CAddress2
    With Clone
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine)
            .AddressLine(i) = m_sAddressLine(i)
        Next
        .Alignment = m_lAlignment
        .AddressOnly = m_bAddressOnly
    End With
End Function

'=======================================================================
'Property:              AddressLine()
'
'Synopsis:              Address array of five lines to place the address that
'                       you want to parse.
'
'=======================================================================
Public Property Get AddressLine(ByVal lIndex As Long) As String
    If lIndex < LBound(m_sAddressLine) Or lIndex > UBound(m_sAddressLine) Then
       Err.Raise 5, m_scClassName & ".AddressLine", "Invalid procedure call.  lIndex parameter is out of range."
    End If
    AddressLine = m_sAddressLine(lIndex)
End Property

Public Property Let AddressLine(ByVal lIndex As Long, ByVal sNewValue As String)
    If lIndex < LBound(m_sAddressLine) Or lIndex > UBound(m_sAddressLine) Then
       Err.Raise 5, m_scClassName & ".AddressLine", "Invalid procedure call.  lIndex parameter is out of range."
    End If
    m_sAddressLine(lIndex) = sNewValue
End Property


'=======================================================================
'Property:              Alignment()
'
'Synopsis:              Specifies whether the AlignedAddress property
'                       is compressed upwards towards the top or downwards
'                       towards the bottom.
'
'=======================================================================
Public Property Get Alignment() As FfgStdAddressAlignmentEnum
    Alignment = m_lAlignment
End Property
Public Property Let Alignment(ByVal lNewValue As FfgStdAddressAlignmentEnum)
    m_lAlignment = lNewValue
End Property

Public Property Get NoMail() As Boolean
    Dim i As Long

'    If Len(CityStateZip) = 0 Then
'        NoMail = True
'        Exit Property
'    End If
'
    'Find the line that the name resides on.
    
    If m_lAlignment = fsaAlignTop Then
        If Len(Trim$(AlignedAddress(2))) = 0 Then
            NoMail = True
            Exit Property
        End If
    Else
        If Len(Trim$(AlignedAddress(5))) = 0 Then
            NoMail = True
            Exit Property
        End If
    End If
    
    For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
        If InStr(1, m_sAddressLine(i), "NO MAIL", vbTextCompare) > 0 Then
            NoMail = True
            Exit Property
        End If
        If InStr(1, m_sAddressLine(i), "SEND TO", vbTextCompare) > 0 Then
            NoMail = True
            Exit Property
        End If
    Next
   
    NoMail = False
    
End Property

'=======================================================================
'Property:              AlignedAddress()
'
'Synopsis:              Returns the address line aligned according
'                       to the Alignment property.
'
'=======================================================================
Public Property Get AlignedAddress(ByVal lIndex As Long) As String
    Dim i As Long
    Dim i2 As Long
    If lIndex < LBound(m_sAddressLine) Or lIndex > UBound(m_sAddressLine) Then
       Err.Raise 5, m_scClassName & ".AlignedAddress", "Invalid procedure call.  lIndex parameter is out of range."
    End If
    If m_lAlignment = fsaAlignTop Then
        i2 = 0
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                i2 = i2 + 1
                If i2 = lIndex Then
                    AlignedAddress = m_sAddressLine(i)
                    Exit For
                End If
            End If
        Next
    Else
        i2 = UBound(m_sAddressLine) - LBound(m_sAddressLine) + 1
        For i = UBound(m_sAddressLine) To LBound(m_sAddressLine) Step -1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                i2 = i2 - 1
                If i2 = lIndex Then
                    AlignedAddress = m_sAddressLine(i)
                    Exit For
                End If
            End If
        Next
    End If
End Property

'=======================================================================
'Property:              AddressOnly()
'
'Synopsis:              Set this to true if the address array does
'                       not contain a name line.
'
'Remarks:               The parse algorithm uses this to determine
'                       wether it should consider the first non
'                       blank line a name.   If not then it can
'                       consider that line for one of the other
'                       elements.
'=======================================================================
Public Property Get AddressOnly() As Boolean
    AddressOnly = m_bAddressOnly
End Property
Public Property Let AddressOnly(ByVal bNewValue As Boolean)
    m_bAddressOnly = bNewValue
End Property

'=======================================================================
'Property:              Name
'
'Synopsis:              Returns the name line of the address array.
'=======================================================================
Public Property Get Name() As String
    Dim lNameLine As Long
    Dim i As Long
    Dim x As Long
    
    If m_bAddressOnly Then
    
        Name = vbNullString
        
        Exit Property
        
    End If

    Name = vbNullString
    
    'Find the line that the name resides on.
    For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
        If Len(Trim$(m_sAddressLine(i))) > 0 Then
            Name = m_sAddressLine(i)
            Exit For
        End If
    Next
End Property

'=======================================================================
'Property:              AddressLine1
'
'Synopsis:              Returns the address line 1 of the address array.
'=======================================================================
Public Property Get AddressLine1() As String
    Dim lNameLine As Long
    Dim lAddressLine1Line As Long
    Dim i As Long
    Dim x As Long
    
    If m_bAddressOnly = False Then
        'Find the line that the name resides on.
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                lNameLine = i
                Exit For
            End If
        Next
        If lNameLine = 0 Then
            'No name found, therefore no non blank lines, and therefore no street address
            AddressLine1 = vbNullString
            Exit Property
        End If
        If lNameLine >= UBound(m_sAddressLine) Then
            'Name found on last line, therefore impossible for street address to be there
            AddressLine1 = vbNullString
            Exit Property
        End If
    End If
    
    'Now that we know the name line, we determine the address line 1 by looking
    'for the third non-blank line (starting from bottom) between the bottom
    'and the name line.
    
    For i = UBound(m_sAddressLine) To lNameLine + 1 Step -1
        If Len(Trim$(m_sAddressLine(i))) > 0 Then
            x = x + 1
        End If
        If x = 3 Then
            lAddressLine1Line = i
            Exit For
        End If
    Next
    
    If lAddressLine1Line > 0 Then
        AddressLine1 = m_sAddressLine(lAddressLine1Line)
    Else
        AddressLine1 = vbNullString
    End If

End Property

'=======================================================================
'Property:              StreetAddress
'
'Synopsis:              Returns the street address of the address array.
'=======================================================================
Public Property Get StreetAddress() As String
    Dim lNameLine As Long
    Dim lStreetAddressLine As Long
    Dim i As Long
    Dim x As Long
    
    If m_bAddressOnly = False Then
        'Find the line that the name resides on.
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                lNameLine = i
                Exit For
            End If
        Next
        If lNameLine = 0 Then
            'No name found, therefore no non blank lines, and therefore no street address
            StreetAddress = vbNullString
            Exit Property
        End If
        If lNameLine >= UBound(m_sAddressLine) Then
            'Name found on last line, therefore impossible for street address to be there
            StreetAddress = vbNullString
            Exit Property
        End If
    End If
    
    'Now that we know the name line, we determine the street address by looking
    'for the second non-blank line (starting from bottom) between the bottom
    'and the name line.
    
    For i = UBound(m_sAddressLine) To lNameLine + 1 Step -1
        If Len(Trim$(m_sAddressLine(i))) > 0 Then
            x = x + 1
        End If
        If x = 2 Then
            lStreetAddressLine = i
            Exit For
        End If
    Next
    
    If lStreetAddressLine > 0 Then
        StreetAddress = m_sAddressLine(lStreetAddressLine)
    Else
        StreetAddress = vbNullString
    End If

End Property

'=======================================================================
'Property:              CityStateZip
'
'Synopsis:              Returns the CSZ line of the address array.
'=======================================================================
Public Property Get CityStateZip() As String
    Dim lNameLine As Long
    Dim lCityStateZipLine As Long
    Dim i As Long
    Dim x As Long
    
    If m_bAddressOnly = False Then
        'Find the line that the name resides on.
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                lNameLine = i
                Exit For
            End If
        Next
        If lNameLine = 0 Then
            'No name found, therefore no non blank lines, and therefore no street address
            CityStateZip = vbNullString
            Exit Property
        End If
        If lNameLine >= UBound(m_sAddressLine) Then
            'Name found on last line, therefore impossible for city state zip to be there
            CityStateZip = vbNullString
            Exit Property
        End If
    End If
    
    'Now that we know the name line, we determine the city, state and zip by looking
    'for the first non-blank line (starting from bottom) between the bottom
    'and the name line.
    
    For i = UBound(m_sAddressLine) To lNameLine + 1 Step -1
        If Len(Trim$(m_sAddressLine(i))) > 0 Then
            lCityStateZipLine = i
            Exit For
        End If
    Next
    
    If lCityStateZipLine > 0 Then
        CityStateZip = m_sAddressLine(lCityStateZipLine)
    Else
        CityStateZip = vbNullString
    End If

End Property

Private Function GetCSZLine() As Long
    Dim lNameLine As Long
    Dim lCityStateZipLine As Long
    Dim i As Long
    Dim x As Long
    
    If m_bAddressOnly = False Then
        'Find the line that the name resides on.
        For i = LBound(m_sAddressLine) To UBound(m_sAddressLine) Step 1
            If Len(Trim$(m_sAddressLine(i))) > 0 Then
                lNameLine = i
                Exit For
            End If
        Next
        If lNameLine = 0 Then
            'No name found, therefore no non blank lines, and therefore no street address
            GetCSZLine = 0
            Exit Function
        End If
        If lNameLine >= UBound(m_sAddressLine) Then
            'Name found on last line, therefore impossible for city state zip to be there
            GetCSZLine = 0
            Exit Function
        End If
    End If
    
    'Now that we know the name line, we determine the city, state and zip by looking
    'for the first non-blank line (starting from bottom) between the bottom
    'and the name line.
    
    For i = UBound(m_sAddressLine) To lNameLine + 1 Step -1
        If Len(Trim$(m_sAddressLine(i))) > 0 Then
            GetCSZLine = i
            Exit Function
        End If
    Next
    
    GetCSZLine = 0

End Function

Public Sub FormatCityStateZip()
    Dim lCSZLine As Long
    lCSZLine = GetCSZLine()
    If lCSZLine > 0 Then
        m_sAddressLine(lCSZLine) = m_oAddressUtility.ParseCSZ(m_sAddressLine(lCSZLine), ffgCSZCityStateZip)
    End If
End Sub

'=======================================================================
'Property:              CityStateZip
'
'Synopsis:              Returns the City and State found in the address
'                       array.
'=======================================================================
Public Property Get CityState() As String
    CityState = m_oAddressUtility.ParseCSZ(CityStateZip, ffgCSZCityState)
End Property

'=======================================================================
'Property:              City
'
'Synopsis:              Returns the City found in the address array.
'=======================================================================
Public Property Get City() As String
    City = m_oAddressUtility.ParseCSZ(CityStateZip, ffgCSZCity)
End Property

'=======================================================================
'Property:              State
'
'Synopsis:              Returns the State found in the address array.
'=======================================================================
Public Property Get State() As String
    State = m_oAddressUtility.ParseCSZ(CityStateZip, ffgCSZState)
End Property

'=======================================================================
'Property:              ZipCode
'
'Synopsis:              Returns the ZipCode found in the address array.
'=======================================================================
Public Property Get ZipCode() As String
    ZipCode = m_oAddressUtility.ParseCSZ(CityStateZip, ffgCSZZip)
End Property
