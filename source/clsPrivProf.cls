VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrivProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Class Module clsPrivProf: Private Profile Files services. The componet works
' ========================= autonomous and does not require any other
' components.
'
' Public properties:
' ------------------
' FileFooter  r/w
' FileName    r/w - Specifies a Private Profile file full name
'                 - Returns a valid Private Profile file's full name
' Header      r/w
' Section     r/w - Returns the current specified Section name
'                 - Specifies the Section name valid for all subsequent
'                   methods and properties until another section name is
'                   specified.
' Sections    r
' Value       r/w - Reads from a Private Profile File a value with a
'                   provided value-name from a provided section
'                 - Writes to a Private Profile File a value with a provided
'                   value-name into a provided section
' Public methods:
' ---------------
' NamesRemove     Removes provided value names, in a given Private
'                 Properties File, when provided in a specific section,
'                 else in all sections.
' Reorg
' Exists          Returns TRUE when a given section exists in a given
'                 Private Profile File.
' SectionNames    Returns a Dictionary of all section names [...] in a
'                 Private Profile File whereby each item is a Dictionary of
'                 value-names as key and the value as item, and the key of
'                 the item is the section name.
' SectionsCopy    Copies sections, provided as comma delimites string of
'                 section names from a soure PrivProfile file into a target
'                 Private Profile file, optionally merged.
' SectionNames
' SectionRemove
' SectionsCopy
' SectionsRemove  Removes the sections provided as a comma delimited string
'                 as argument. Sections not existing are ignored.
' ValueNameRename Function replaces an old value name with a new one
'                 either in a specific section or in all sections when no
'                 specific section is provided. Optionally not reorgs the
'                 file, returns True when at least one name has been
'                 replaced.
' ValueNames      Returns a Dictionary of all value-names within given
'                 sections in a Private Profile File with the value-name and
'                 the section name as key (<name>[section]) and the value as
'                 item, the names in ascending order in a Dictionary.
'                 Section names may be provided as a comma delimited string,
'                 a Dictionary or Collection. Non existing sections are
'                 ignored.
' ValueRemove
' Values          Returns the value-names and values of a given section in a
'                 Private Profile File as Dictionary with the value-name as
'                 the key (in ascending order) and the value as item.
'
' Requires:
' ---------
' Reference to Microsoft Scripting Runtine
' Reference to Microsoft VBScript Regular Expressions 5.5
'
' Uses no other components. Will use optionally mErH, fMsg/mMsg when installed
' and activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
'
' W. Rauschenberger, Berlin Jan 2024
' See also https://github.com/warbe-maker/VBA-File-System-Objects.
' ----------------------------------------------------------------------------
Public Fso                      As FileSystemObject

Private Const WIN_NORMAL = 1         'Open Normal
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

#If mMsg = 0 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' approach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument mMsg = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Declare PtrSafe Function WritePrivateProfileString _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpw_ApplicationName As String, _
                ByVal lpw_KeyName As String, _
                ByVal lpw_String As String, _
                ByVal lpw_FileName As String) As Long
                
Private Declare PtrSafe Function GetPrivateProfileString _
                Lib "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpg_ApplicationName As String, _
                ByVal lpg_KeyName As String, _
                ByVal lpg_Default As String, _
                ByVal lpg_ReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpg_FileName As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileSection _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal NoKey As Long, _
                ByVal NoSetting As Long, _
                ByVal Name As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal Name As String) As Long
                 
Private Declare PtrSafe Function GetPrivateProfileSectionNames _
                Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" _
               (ByVal lpszReturnBuffer As String, _
                ByVal nSize As Long, _
                ByVal lpName As String) As Long

Private Declare PtrSafe Function apiShellExecute _
                Lib "shell32.dll" Alias "ShellExecuteA" _
               (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private dctFiles                As Dictionary ' file name as key, item is dctSections
Private dctFilesStrng           As Dictionary ' file as string
Private dctSections             As Dictionary ' section name as key, item is dctValues
Private dctValues               As Dictionary ' value name as key, item is value
Private Reg                     As RegExp
Private sFileName               As String
Private sSection                As String
Private sDelimChar              As String

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    Set Reg = New RegExp
    Set Fso = New FileSystemObject
    Set dctFiles = New Dictionary
    Set dctFilesStrng = New Dictionary
    
    '~~ Initialize the default Private Profile File name
    Select Case Application.Name
        Case "Microsoft Excel": sFileName = ThisWorkbook.Path & "\" & Fso.GetBaseName(ThisWorkbook.Name) & ".dat"
        Case Else:              Err.Raise AppErr(1), ErrSrc(PROC), "Initializing the FileName failed!" & vbLf & _
                                                                   "The current Application.Name is not Excel " & _
                                                                   "and any other application is yet not supported!"
    End Select
    sDelimChar = "="
    
End Sub

Private Sub Class_Terminate()
    Set Fso = Nothing
    Set dctFiles = Nothing
End Sub

Public Property Get FileFooter(Optional ByVal name_file As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' When the file has no footer an empty Collection is returned.
' ----------------------------------------------------------------------------
    Const PROC = "FileFooter-Get"
    
    Dim bActive As Boolean
    Dim cll     As New Collection
    Dim i       As Long
    Dim arr     As Variant
    
    Set FileFooter = cll
    If Not IsValidFileName(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not bActive Then GoTo xt
    
    i = UBound(arr)
    Do While IsComment(arr(i)) And i >= LBound(arr)
        '~~ Any bottom comment up to the first non-comment is regarded a file footer
        cll.Add arr(i)
        i = i - 1
    Loop
    If cll.Count > 0 _
    Then Set FileFooter = RevertOrder(cll)

xt: Set cll = Nothing
    
End Property

Public Property Let FileFooter(Optional ByVal name_file As String = vbNullString, _
                                        ByVal file_footer As Variant)
' ----------------------------------------------------------------------------
' Replaces an old footer in file (name_file) - if any - with the new footer
' (file_footer) - which may be a series of strings delimited by vbCrLf, vbLf,
' "|" or "||".
' ----------------------------------------------------------------------------
    Const PROC = "FileFooter-Let"
    
    Dim arr         As Variant
    Dim arrOut      As Variant
    Dim bActive     As Boolean
    Dim cllFootOld  As Collection
    Dim i           As Long
    Dim v           As Variant
    
    If Not IsValidFileName(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not bActive _
    Then Err.Raise AppErr(2), ErrSrc(PROC), _
         "The file is not ready adding a footer since it apparently does not have any Section nor any Value!"
    
    Set cllFootOld = FileFooter(name_file)
    
    '~~ Transfer all except the old footer (if any)
    For i = LBound(arr) To UBound(arr) - cllFootOld.Count
        ArrayAdd arrOut, arr(i)
    Next i
    
    '~~ Add the file footer in case one has been provided
    For Each v In VarItemsAsCollection(file_footer)
        ArrayAdd arrOut, v
    Next v
    
    '~~ Re-write Private Profile File with new/replaced fime header if any
    FileFromString name_file, Join(arrOut, vbCrLf)

End Property

Public Property Get FileName() As String:               FileName = sFileName:       End Property

Public Property Let FileName(ByVal f_file As String)
' ----------------------------------------------------------------------------
' Note: Only a FileName explicitely provided as Property becomes the default
'       file name for an omitted file argument.
'       Any file name provided as argument along with a property or method
'       call is only valid for the call.
' ----------------------------------------------------------------------------
    Const PROC = "FileName-Let"
    
    If Not IsValidFileName(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(f_file)
    
    sFileName = f_file

End Property

Public Property Get Header(Optional ByVal name_file As String = vbNullString, _
                           Optional ByVal name_section As String = vbNullString, _
                           Optional ByVal name_value As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Header-Get"
    
    Dim cll As New Collection
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
                                    
    Select Case True
        Case name_section = vbNullString And name_value = vbNullString:   Set cll = HeaderFile(name_file)
        Case name_section <> vbNullString And name_value = vbNullString:  Set cll = HeaderSection(name_section, name_file)
        Case name_section <> vbNullString And name_value <> vbNullString: Set cll = HeaderValue(name_value, name_section, name_file)
    End Select
    
    Set Header = cll
    Set cll = Nothing
    
End Property

Public Property Let Header(Optional ByVal name_file As String = vbNullString, _
                           Optional ByVal name_section As String = vbNullString, _
                           Optional ByVal name_value As String = vbNullString, _
                                    ByVal h_head As Variant)
' ------------------------------------------------------------------------------
' Replaces an old file (name_file), section (name_section), or value (name_valuee) header -
' if any - with the a new one (h_head) - which may be a series of strings
' delimited by vbCrLf, vbLf, "|" or "||", or a Collection of strings.
' ------------------------------------------------------------------------------
    Const PROC = "Header-Let"
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
                                        
    Select Case True
        Case name_section = vbNullString And name_value = vbNullString:   HeaderFile(name_file) = h_head
        Case name_section <> vbNullString And name_value = vbNullString:  HeaderSection(name_section, name_file) = h_head
        Case name_section <> vbNullString And name_value <> vbNullString: HeaderValue(name_value, name_section, name_file) = h_head
    End Select
    
    FileAsDictMan name_file ' ensure ascending order
                                    
End Property

Private Property Get HeaderFile(Optional ByVal h_file_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns a Collection with the file header lines (may be empty). The provided
' (h_file_name) may be an array already with the file's content.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderFile-Get"
    
    Dim arr         As Variant
    Dim bActive     As Boolean
    Dim cll         As New Collection
    Dim i           As Long
    Dim j           As Long
    Dim lLBoundTop  As Long
    
    Set HeaderFile = cll
    arr = Split(FileAsString(h_file_name), vbCrLf)
    For i = UBound(arr) To LBound(arr) Step -1
        If IsDelimiterFileHeader(arr(i)) Then
            cll.Add arr(i)
            For j = i - 1 To LBound(arr) Step -1
                cll.Add arr(j)
            Next j
            Set HeaderFile = RevertOrder(cll)
            GoTo xt
        End If
    Next i
        
xt: Set cll = Nothing
    
End Property

Private Property Let HeaderFile(Optional ByVal h_file_name As String, _
                                         ByVal h_file_header As Variant)
' ----------------------------------------------------------------------------
' Replaces any existing file header comments with new ones (h_file_header),
' whereby the new header comments (h_file_header) may be provided as a series
' of strings delimited by vbCrLf, vbLf, "|" or "||", a Collection of
' strings, or a Dictionary with the header comments a key.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderFile-Let"
    
    Dim cllHeaderOld    As Collection
    Dim arr             As Variant
    Dim arrOut          As Variant
    Dim i               As Long
    Dim v               As Variant
    Dim lFrst           As Long
    Dim lLast           As Long
    Dim bActive         As Boolean
    Dim lLenMax         As Long
    Dim sComment        As String
    
    Set cllHeaderOld = HeaderFile(h_file_name)
    If cllHeaderOld.Count = 0 And VarItemsAsCollection(h_file_header).Count = 0 Then Exit Property
        
    '~~ Insert the new file header comments excluding a possibly provided header delimiter
    '~~ and subsequently insert a header delimiter adjusted to the longest haeder comment
    For Each v In VarItemsAsCollection(h_file_header)
        If Not IsDelimiterFileHeader(v) Then
            sComment = Comment(v)
            lLenMax = Max(lLenMax, Len(sComment) - 2)
            ArrayAdd arrOut, sComment
        End If
    Next v
    ArrayAdd arrOut, DelimiterFileHeader(lLenMax)
    
    '~~ Transfer all lines except the old file header
    arr = Split(FileAsString(h_file_name), vbCrLf)
    For i = cllHeaderOld.Count To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    
    FileFromString h_file_name, Join(arrOut, vbCrLf)

End Property

Private Property Get HeaderSection(Optional ByVal h_sect_name As String, _
                                   Optional ByVal h_file_name As String = vbNullString, _
                                   Optional ByRef h_sect_indx As Long) As Variant
' ----------------------------------------------------------------------------
' Returns a Collection with the section's (h_sect_name_name) header lines (possibly
' empty).The provided (name_file) may be an array already with the file's content.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderSection-Get"
    
    Dim arr     As Variant
    Dim bActive As Boolean
    Dim cll     As New Collection
    Dim i       As Long
    
    arr = Split(FileAsString(h_file_name), vbCrLf)
    If SectionExists(arr, h_sect_name, i, cll) Then
        Set HeaderSection = cll
    End If
    
xt: Set cll = Nothing
    
End Property

Private Property Let HeaderSection(Optional ByVal h_sect_name As String, _
                                   Optional ByVal h_file_name As String = vbNullString, _
                                   Optional ByRef h_sect_indx As Long, _
                                            ByVal h_sect_head As Variant)
' ----------------------------------------------------------------------------
' Replaces in file (h_file_name) an old section header - if any - with the a new
' one (h_header) - which may be a series of strings delimited by vbCrLf,
' vbLf, "|" or "||", or a Collection of strings.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderSection-Let"
    
    Dim bActive      As Boolean
    Dim cllHeaderOld As Collection
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim arr          As Variant
    Dim v            As Variant
    Dim arrOut       As Variant
    
    arr = Split(FileAsString(h_file_name), vbCrLf)
    If Not SectionExists(arr, h_sect_name, k, cllHeaderOld) Then GoTo xt
    
    '~~ Transfer down to section excluding an old header
    For i = LBound(arr) To k - 1 - cllHeaderOld.Count
        ArrayAdd arrOut, arr(i)
    Next i
    
    '~~ Insert new header
    For Each v In VarItemsAsCollection(h_sect_head)
        ArrayAdd arrOut, Comment(v)
    Next v
    
    '~~ Transfer all remaining lines
    For i = k To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    
    FileFromString h_file_name, Join(arrOut, vbCrLf)
    
xt:
End Property

Private Property Get HeaderValue(Optional ByVal h_valu_name As String, _
                                 Optional ByVal h_sect_name As String = vbNullString, _
                                 Optional ByVal h_file_name As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Returns a Collection with the value's (h_sect_name) header lines (may be empty).
' The provided (name_file) may be an array already with the file's content.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderValue-Get"
                                  
    On Error GoTo eh
    Dim arr     As Variant
    Dim bActive As Boolean
    Dim cll     As New Collection
    Dim i       As Long
    Dim sValue  As String
    
    Set HeaderValue = cll
    arr = Split(FileAsString(h_file_name), vbCrLf)
    If ValueExists(arr, h_sect_name, h_valu_name, sValue, i, cll) Then
        Set HeaderValue = cll
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let HeaderValue(Optional ByVal h_valu_name As String, _
                                 Optional ByVal h_sect_name As String = vbNullString, _
                                 Optional ByVal h_file_name As String = vbNullString, _
                                          ByVal h_valu_head As Variant)
' ----------------------------------------------------------------------------
' Replaces in an - optionally provided file's - section's (s_section) value
' (h_value_name) a header with a provided header (h_header) which may be
' provided as a series of strings delimited by a | (vertical bar) or as a
' collection of strings.
' ----------------------------------------------------------------------------
    Const PROC = "HeaderValue-Let"
    
    On Error GoTo eh
    Dim arrOut          As Variant
    Dim arr             As Variant
    Dim bActive         As Boolean
    Dim i               As Long
    Dim j               As Long
    Dim cllHeaderOld    As Collection
    Dim sValue          As String
    Dim v               As Variant
    
    arr = Split(FileAsString(h_file_name), vbCrLf)
    If Not IsValidSectName(h_sect_name) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(h_sect_name)
    If Not ValueExists(arr, h_sect_name, h_valu_name, sValue, j, cllHeaderOld) Then GoTo xt
    
    '~~ Transfer all lines down to value except a possible old header
    For i = LBound(arr) To j - 1 - cllHeaderOld.Count
        ArrayAdd arrOut, arr(i)
    Next i
    '~~ Insert the new header
    For Each v In VarItemsAsCollection(h_valu_head)
        ArrayAdd arrOut, Comment(v)
    Next v
    '~~ Transfer all remaining lines
    For i = j To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    
    FileFromString h_file_name, Join(arrOut, vbCrLf)

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Section() As String:        Section = sSection: End Property

Public Property Let Section(ByVal s As String)
' ----------------------------------------------------------------------------
' Note: Only a Section provided as Property becomes the default section name.
'       Any section name provided along with a property or method call is only
'       valid for the call.
' ----------------------------------------------------------------------------
    Const PROC = "Section-Let"
    
    If Not IsValidSectName(s) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidSectName(s)
    sSection = s
    
End Property

Public Property Get Sections(Optional ByVal name_file As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all named sections (pp_sections) in file (pp_file) -
' provided as full name string or as file object - whereby each section is a
' Dictionary with the section name as the key - in ascending
' order - and a Dictionary of the section's values as item with the value name
' as key and the value as item.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim dct         As New Dictionary
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
    '~~ Return all sections
    With dct
        For Each vSection In SectionNames(name_file)
            .Add vSection, ValueNames(name_file, vSection)
        Next vSection
    End With

xt: Set Sections = KeySort(dct)
    Set dct = Nothing
    Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Sections(Optional ByVal name_file As String = vbNullString, _
                                       ByVal s_dct As Dictionary)
' ------------------------------------------------------------------------
' Writes the sections provided as Dictionary (pp_dct) to file (pp_file) -
' provided as full name string or as file object. Existing sections/values
' are overwritten new sections/values are added.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ------------------------------------------------------------------------
    Const PROC = "Sections-Let"
    
    On Error GoTo eh
    Dim vName       As Variant
    Dim vSection    As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    Dim fl          As String
    Dim vValue      As Variant
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    
    For Each vSection In s_dct
        sSection = vSection
        Set dctValues = s_dct(sSection)
        For Each vName In dctValues
            sName = vName
            vValue = dctValues.item(vName)
            Value(name_file:=name_file _
                , name_section:=sSection _
                , name_value:=sName _
                 ) = vValue
        Next vName
    Next vSection
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByRef value_header As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Returns value name's (name_value) value from a section (name_section) in a
' Private Profile file (name_file) - plus its header (it any) as a string of
' header lines delimited with a "||".
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Value-Get"
    
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant
    Dim fl      As String
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidValuName(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
    
    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=name_section _
                                    , lpg_KeyName:=name_value _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=name_file _
                                     )
    Value = Left$(sRetVal, lResult)
    '~~ Return the value header - if any
    value_header = VarItemsAsString(HeaderValue(name_value, name_section, name_file), "||")
    
End Property

Public Property Let Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByRef value_header As Variant = vbNullString, _
                                   ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value (v_value) under a given name (name_value) into a section
' (name_section) in a Private Profile file (name_file).
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
        
    Dim bIsNew  As Boolean
    Dim lChars  As Long
    Dim sValue  As String
    Dim fl      As String
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidValuName(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
    
    Select Case VarType(v_value)
        Case vbBoolean: sValue = VBA.CStr(VBA.CLng(v_value))
        Case Else:      sValue = v_value
    End Select
    
    bIsNew = Not Exists(name_file, name_section, name_value)
    
    lChars = WritePrivateProfileString(lpw_ApplicationName:=name_section _
                                     , lpw_KeyName:=name_value _
                                     , lpw_String:=v_value _
                                     , lpw_FileName:=name_file)
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & name_section & "'" & vbLf & _
               "Value name = '" & name_value & "'" & vbLf & _
               "Value      = '" & CStr(v_value) & "'" & vbLf & _
               "Value file = '" & name_file & "'"
    End If
    
    Header(name_file, name_section, name_value) = value_header
    If bIsNew Then
        FileAsDictMan name_file ' re-write the file in order to maintain ascending value name sequence
    End If
    
End Property

Private Sub ArrayAdd(ByRef a_arr As Variant, _
                     ByVal a_str As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                     
    On Error Resume Next
    ReDim Preserve a_arr(UBound(a_arr) + 1)
    If Err.Number <> 0 Then ReDim a_arr(0)
    a_arr(UBound(a_arr)) = a_str
    
End Sub

Private Function ArrayIsAllocated(ByVal arry As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = UBound(arry) >= LBound(arry)
    On Error GoTo -1
    
End Function

Private Sub BoC(ByVal b_id As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Bnd-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.BoC b_id, b_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.BoC b_id, b_args
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErh Then          ' serves the mTrc/clsTrc when installed and active
    mErh.BoP b_proc, b_args
#ElseIf XcTrc_clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Function Comment(ByVal s As String) As String
    
    Comment = Replace(";" & s, ";;", ";")
    If Left(Comment, 2) <> "; " Then
        Comment = "; " & Right(Comment, Len(Comment) - 1)
    End If
    
End Function

Private Function DelimiterFileHeader(Optional ByVal lMax As Long = 2) As String
    DelimiterFileHeader = "; " & String(lMax, sDelimChar)
End Function

Private Sub EoC(ByVal e_id As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.EoC e_id, e_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.EoC e_id, e_args
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErh = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErh.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrInvalidFileName(ByVal s As String) As String
    ErrInvalidFileName = "The provided file name (" & s & ") is not a valid Private Profile file!" & vbLf & _
                         "Valid is: - a not yet existing file" & vbLf & _
                         "          - an existing but empty file" & vbLf & _
                         "          - an existing file of which the content adheres to Private Profile file content rules." & vbLf & _
                         "            (see " & _
                         "has one been specified explicitely through the FileName property!"
End Function

Private Function ErrInvalidSectName(ByVal s As String) As String
    ErrInvalidSectName = "The provided section name (" & s & ") is invalid! " & _
                         "In a Private Profile files a section names may only " & _
                         "contain the characters A-Z, a-z, 0-9, and _ !"
End Function

Private Function ErrInvalidValuName(ByVal s As String) As String
    ErrInvalidValuName = "The provided value name (" & s & ") is invalid! " & _
                         "In a Private Profile file a value name may only " & _
                         "contain the characters A-Z, a-z, 0-9, and _ whereby the length is >= 1 <= 32 characters!"
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErh = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErh.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else: ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPrivProf." & sProc
End Function

Public Function Exists(ByVal name_file As String, _
              Optional ByRef name_section As String = vbNullString, _
              Optional ByVal name_value As String = vbNullString, _
              Optional ByRef head_value As Collection, _
              Optional ByRef e_valu As String, _
              Optional ByRef head_section As Collection, _
              Optional ByRef head_file As Collection) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided value name, section name, file name exists.
' When the provided (section (name_section
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim bActive As Boolean
    Dim i       As Long
    
    If Not IsValidFileName(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not bActive Then GoTo xt
        
    Select Case True
        Case name_section <> vbNullString And name_value = vbNullString:    Exists = SectionExists(arr, name_section, i, head_section)
        Case name_section <> vbNullString And name_value <> vbNullString:   Exists = ValueExists(arr, name_section, name_value, e_valu, i, head_value, head_section)
    End Select
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub FileAsDict(ByVal f_file As String)
' ----------------------------------------------------------------------------
' The Private Profile file saved as structure of Dictionaries serves as a
' means to keep all items in ascending order and for existence checks.
' --------------------------------------
' ----- Element -----
' Type       Name       Key          Items     Type
' ---------- ---------- ------------ ---------- ----------
' Dictionary dctFiles   file_name    cllFile    Collection
' Collection cllFile                 cllHeader  Collection *)
'                                    dctFile    Dictionary
'                                    cllFooter  Collection *)
' Dictionary dctFile    section_name cllSection Collection
' Collection cllSection              cllHeader  Collection *)
'                                    dctSection Dictionary
' Dictionary dctSection value_name   cllValue   Collection
' Collection cllValue                cllHeader  Collection *)
'                                    value      String
'
' *) cllHeader, cllFooter are a collection of comment lines - possibly empty.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsDict"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim cll         As New Collection
    Dim cllFile     As New Collection
    Dim cllFlHeader As New Collection
    Dim cllFlFooter As New Collection
    Dim cllHeader   As New Collection
    Dim cllHeadSect As Collection
    Dim cllSect     As New Collection
    Dim cllValue    As New Collection
    Dim dctFile     As New Dictionary
    Dim dctSect     As New Dictionary
    Dim dctSects    As New Dictionary
    Dim i           As Long
    Dim s           As String
    Dim sFileStrng  As String
    Dim sName       As String
    Dim sValue      As String
    
    BoP ErrSrc(PROC), f_file
    FileAsDictRem f_file
    arr = Split(FileAsString(f_file), vbCrLf)
    Set cllFlHeader = HeaderFile(f_file)
    Set cllFlFooter = FileFooter(f_file)
    
    For i = cllFlHeader.Count To UBound(arr) - cllFlFooter.Count
        sName = vbNullString
        s = vbNullString
        Select Case True
            Case IsComment(arr(i))
                '~~ Collect any comment except a file header delimiter
                cll.Add arr(i)
            
            Case IsDelimiterFileHeader(arr(i))
                '~~ A file header delimiter indicates the end of a file header
                cll.Add arr(i)
                Set cllFlHeader = cll
                Set cll = New Collection
            
            Case IsSection(arr(i), s)
                '~~ Stow section plus any possible header (may be empty)
                If sSection <> vbNullString Then
                    '~~ This is a subsequent section which indicates that the
                    '~~ previous with its collected values has to be saved
                    FileAsDictSec fs_sect:=sSection _
                                , fs_sect_dct:=dctSect _
                                , fs_sect_header:=cllHeadSect _
                                , fs_file_dct:=dctFile
                End If
                sSection = s
                Set cllHeadSect = cll
                Set cll = New Collection
            
            Case IsValue(arr(i), sName, sValue)
                '~~ Collect value in dctSect with the value name as key
                FileAsDictVal fname_value:=sName _
                            , fv_value:=sValue _
                            , fv_value_header:=cll _
                            , fv_sect_dct:=dctSect
        End Select
    Next i
    
    If dctSect.Count > 0 Then
        FileAsDictSec fs_sect:=sSection _
                    , fs_sect_dct:=dctSect _
                    , fs_sect_header:=cll _
                    , fs_file_dct:=dctFile
    End If
        
    '~~ Stow the file with possible header and footer in the dctFiles Directory
    FileAsDictAdd f_file:=f_file _
               , f_file_header:=cllFlHeader _
               , f_file_dct:=dctFile _
               , f_file_footer:=cllFlFooter
    
xt: Set cll = Nothing
    Set dctSect = Nothing
    Set dctFile = Nothing
    Set cllFlHeader = Nothing
    EoP ErrSrc(PROC), f_file
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub FileAsDictAdd(ByVal f_file As String, _
                          ByVal f_file_header As Collection, _
                          ByVal f_file_dct As Dictionary, _
                          ByVal f_file_footer As Collection)
' ----------------------------------------------------------------------------
' Add a new created file (f_file) as Dictionary to the Directory of files
' dctFiles.
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim cll As New Collection
    
    cll.Add f_file_header
    cll.Add f_file_dct
    cll.Add f_file_footer
    dctFiles.Add f_file, cll
    
    '~~ Additionally stow the file as string, allowing to detect when the file has hanged
    If dctFilesStrng.Exists(f_file) Then
        If FileHasChanged(f_file) Then
            dctFilesStrng.Remove f_file
            dctFilesStrng.Add f_file, FileAsString(f_file)
        End If
    Else
        dctFilesStrng.Add f_file, FileAsString(f_file)
    End If
'    Debug.Print "Saved File: Headers " & f_file_header.Count & ", Sections " & f_file_dct.Count & ", Footers " & f_file_footer.Count
    Set cll = Nothing

End Sub

Private Sub FileAsDictMan(ByVal f_file As String)
' ----------------------------------------------------------------------------
' When the file (f_file) yet not exists in the Dictionary of files (dctFiles)
' or when in exists but it has changed, add it as Dictionary to the Dictionary
' of files dctFiles.
' ----------------------------------------------------------------------------
       
    If Not dctFiles.Exists(f_file) _
    Or FileHasChanged(f_file) Then
        FileAsDict f_file
        FileFromDict f_file ' re-writing the content ensures ascending order
    End If

End Sub

Private Sub FileAsDictRem(ByVal f_file As String)
    If dctFiles.Exists(f_file) Then dctFiles.Remove f_file
    If dctFilesStrng.Exists(f_file) Then dctFilesStrng.Remove f_file
End Sub

Private Sub FileAsDictSec(ByVal fs_sect As String, _
                          ByRef fs_sect_dct As Dictionary, _
                          ByRef fs_file_dct As Dictionary, _
                          ByRef fs_sect_header As Collection)
' ----------------------------------------------------------------------------
' Adds a section (fs_dct_sect) to the file Dictionary (fs_file_dct).
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    If fs_sect_dct.Count = 0 Then           ' empty section
        GoTo xt
    End If
    
    cll.Add fs_sect_header              ' transfer possible comments to the section Collection
    cll.Add KeySort(fs_sect_dct)        ' add the section as the last item to the Collection
    fs_file_dct.Add fs_sect, cll        ' add the Collection with the section name as key to the file Dictionary
    KeySort fs_file_dct
'    Debug.Print "Saved: Section " & fs_sect & " ," & fs_sect_header.Count & " headers, " & fs_sect_dct.Count & " values."
    Set fs_sect_header = New Collection
    Set fs_sect_dct = New Dictionary
    
xt: Set cll = Nothing

End Sub

Private Sub FileAsDictVal(ByVal fname_value As String, _
                          ByVal fv_value As String, _
                          ByRef fv_sect_dct As Dictionary, _
                          ByRef fv_value_header As Collection)
' ----------------------------------------------------------------------------
' This is the first and only, the last one, or a subsequent value name/value
' in a section which needs to be added to the section Directory (fv_sect_dct)
' together with possibly collected comments (fv_comments).
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    cll.Add fv_value_header                 ' add the value header as item to the Collection
    cll.Add fv_value                        ' add the value as item to the Collection
    If Not fv_sect_dct.Exists(fname_value) Then
        fv_sect_dct.Add fname_value, cll ' add the Collection
        KeySort fv_sect_dct
'        Debug.Print "Saved: Value   " & fv_value & " ," & fv_value_header.Count & " headers."
    End If
    Set fv_value_header = New Collection
    Set cll = Nothing
    
End Sub

Private Function FileAsString(Optional ByVal f_file_full_name As String, _
                              Optional ByVal f_append As Boolean = False, _
                              Optional ByRef f_split As String = vbCrLf, _
                              Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    On Error Resume Next
    Open f_file_full_name For Input As #1
    If Err.Number <> 0 Then GoTo xt
    FileAsString = Input$(LOF(1), 1)
    Close #1
    
    Select Case True
        Case InStr(FileAsString, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(FileAsString, vbCr) <> 0:   f_split = vbCr
        Case InStr(FileAsString, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(FileAsString, 1) = VBA.Chr(26) Then
        FileAsString = Left(FileAsString, Len(FileAsString) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    Do While Right(FileAsString, Len(f_split)) = f_split
        FileAsString = Left(FileAsString, Len(FileAsString) - Len(f_split))
        If Len(FileAsString) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        FileAsString = FileAsStringEmptyExcluded(FileAsString)
    End If
    
xt:
End Function

Private Function FileAsStringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileStringTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = s_s
    
End Function

                             
Private Sub FileFromDict(ByVal f_file As String)
' ----------------------------------------------------------------------------
' Replaces the content of a Private Profile file (f_file) with a reorganized
' Directory structure, i.e. with all sections and value names in ascending
' order.
' ----------------------------------------------------------------------------
    Const PROC = "FileFromDict"
    
    On Error GoTo eh
    Dim arrOut()    As String
    Dim vSect       As Variant
    Dim vValue      As Variant
    Dim dctFile     As Dictionary
    Dim cllFile     As Collection
    Dim dctSect     As Dictionary
    Dim cllSect     As Collection
    
    FileAsDictMan f_file ' reloads when changed
    
    Set cllFile = dctFiles(f_file)
    Set dctFile = cllFile(2)
    
    FileFromDictComments arrOut, cllFile(1)                    ' File header
    For Each vSect In dctFile
        FileFromDictComments arrOut, dctFile(vSect)(1)         ' Section header
        ArrayAdd arrOut, "[" & vSect & "]"                     ' Section
        Set dctSect = dctFile(vSect)(2)
        For Each vValue In dctSect
            FileFromDictComments arrOut, dctSect(vValue)(1)    ' Value header
            ArrayAdd arrOut, vValue & "=" & dctSect(vValue)(2) ' Value
        Next vValue
    Next vSect
    
    FileFromDictComments arrOut, cllFile(3)                           ' File footer
'    Debug.Print "Restore File Footer " & cllFile(3).Count & " lines"
    
    FileFromString f_file, Join(arrOut, vbCrLf)
    
    If dctFilesStrng.Exists(f_file) Then
        If FileHasChanged(f_file) Then
            dctFilesStrng.Remove f_file
            dctFilesStrng.Add f_file, FileAsString(f_file)
        End If
    Else
        dctFilesStrng.Add f_file, FileAsString(f_file)
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub FileFromDictComments(ByRef f_arr As Variant, _
                               ByVal f_cll As Collection)
    Dim v As Variant
    
    For Each v In f_cll
        ArrayAdd f_arr, Replace(";" & v, ";;", ";")
    Next v
    
End Sub

Private Sub FileFromString(ByVal f_file_full_name As String, _
                           ByVal f_s As String, _
                  Optional ByVal f_appended As Boolean = False)
' ----------------------------------------------------------------------------
' Writes a string (f_s) with multiple records/lines delimited by a vbCrLf to
' a file (f_file_full_name).
' ----------------------------------------------------------------------------
    
    If f_appended _
    Then Open f_file_full_name For Append As #1 _
    Else Open f_file_full_name For Output As #1
    Print #1, f_s
    Close #1
    
End Sub

Private Function FileHasChanged(ByVal f_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (f_file) differs from the file in dctFilesStrng
' or when the file in dctFilesStrng does not exist.
' ----------------------------------------------------------------------------
    Dim a1  As Variant
    Dim a2  As Variant
    Dim i   As Long
    
    If dctFilesStrng.Exists(f_file) Then
        FileHasChanged = StrComp(dctFilesStrng(f_file), FileAsString(f_file), vbBinaryCompare) <> 0
'        Test/debug only
'        If FileHasChanged Then
'            Debug.Print "Differs"
'            a1 = Split(dctFilesStrng(f_file), vbCrLf)
'            a2 = Split(FileAsString(f_file), vbCrLf)
'            For i = LBound(a1) To Min(UBound(a1), UBound(a2))
'                If a1(i) <> a2(i) Then
'                    Debug.Print a1(i)
'                    Debug.Print a2(i)
'                    Exit For
'                End If
'            Next i
'            If UBound(a1) > UBound(a2) Then
'                For i = UBound(a2) To UBound(a1)
'                    Debug.Print a1(i)
'                Next i
'            ElseIf UBound(a2) > UBound(a1) Then
'                For i = UBound(a1) To UBound(a2)
'                    Debug.Print a2(i)
'                Next i
'            End If
'
'        End If
    Else
        FileHasChanged = True
    End If
    
End Function

Private Function FileStringTrimmed(ByVal s_s As String, _
                          Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
    s = s_s
    '~~ Eliminate any leading empty items
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    '~~ Eliminate any trailing empty items
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    FileStringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Private Function HeaderItems(ByVal i_array As Variant, _
                    Optional ByVal i_sect As String = vbNullString, _
                    Optional ByVal i_value As String = vbNullString, _
                    Optional ByRef i_frst As Long, _
                    Optional ByRef i_last As Long) As Boolean
' ------------------------------------------------------------------------------
' Returns:
' - the top (i_frst) and the bottom index (i_lasr) of a header
'   when a section (i_sect) or value (i_value) is provided that items header,
'   else the file's header
' - TRUE: a header head been localized.
' - FALSE: no header has been detected (-1 is returned for i_frst and i_last).
' ------------------------------------------------------------------------------

    Dim i       As Long
    Dim sSect   As String
    Dim sVal    As String
    
    i_frst = -1
    i_last = -1
    
    If i_sect = vbNullString Then
        If Left(i_array(0), 1) <> ";" Then GoTo xt
        i_frst = 0
        For i = 0 To UBound(i_array)
            If IsDelimiterFileHeader(i_array(i)) Then
                i_last = i
                GoTo xt
            ElseIf IsSection(i_array(i)) Then
                GoTo xt
            End If
        Next i
    End If
    
    If i_sect <> vbNullString Then
        '~~ Get the bottom index of a requested section's header
        '~~ Attention: When it is the very first section with neither a section nor a file header, -1 is returned as i_last.
        For i = LBound(i_array) To UBound(i_array)
            If IsSection(i_array(i), i_sect) Then
                i_last = i - 1
                Exit For
            End If
        Next i
    End If
    
    If i_value <> vbNullString Then
        '~~ Get bottm index of a regquested value's header
        For i = i + 1 To UBound(i_array)
            If IsSection(i_array(i)) Then Exit For ' end of section
            If IsValue(i_array(i), i_value) Then
                sVal = i_value
                i_last = i - 1
                Exit For
            End If
        Next i
    End If
        
    If i_last = -1 Then GoTo xt ' it is the very first section and it has neuther a header nor ther is a file header
    If Not IsComment(i_array(i_last)) Then GoTo xt ' the item (section or value) has no header or the header above is just the file header
    
    '~~ Get top index of the item's header
    i = i_last
    i_frst = i_last
    Do While IsComment(i_array(i)) And i >= LBound(i_array)
        i_frst = i
        i = i - 1
    Loop
    
xt:
End Function

Private Function HeaderOfItem(ByVal h_arr As Variant, _
                              ByRef h_idx As Long) As Collection
    Dim cll As New Collection
    
    If h_idx = LBound(h_arr) Then GoTo xt
    h_idx = h_idx - 1
    Do While IsComment(h_arr(h_idx))
        cll.Add h_arr(h_idx)
        If h_idx = LBound(h_arr) Then Exit Do
        h_idx = h_idx - 1
    Loop

xt: Set HeaderOfItem = RevertOrder(cll)
    Set cll = Nothing
    
End Function

Private Sub HeaderReplace(ByRef h_arr_in As Variant, _
                          ByVal h_frst As Long, _
                          ByVal h_last As Long, _
                          ByVal h_header As Collection, _
                          ByRef h_arr_out As Variant)
' ----------------------------------------------------------------------------
' Returns the input array (h_arr_in) as output (h_arr_out) with an  existing
' header - index from top (h_frst) to bottom (h_last) - eliminated and
' replaced by a provided header (h_header).
' ----------------------------------------------------------------------------
                   
    Dim i As Long
    Dim j As Long: j = -1
    Dim v As Variant
    
    '~~ Transfer top lines if any.
    '~~ In case the section has no header and it is the very first one, no lines will be transferred).
    For i = LBound(h_arr_in) To h_frst - 1
        j = j + 1
        ReDim Preserve h_arr_out(j)
        h_arr_out(j) = h_arr_in(i)
    Next i
        
    '~~ Provide the new header if any
    For Each v In h_header
        j = j + 1
        ReDim Preserve h_arr_out(j)
        h_arr_out(j) = "; " & Trim(Replace(v, ";", vbNullString))
    Next v
        
    '~~ Transfer all the remaining items from h_arr_in to h_arr_out
    For i = h_last + 1 To UBound(h_arr_in)
        j = j + 1
        ReDim Preserve h_arr_out(j)
        h_arr_out(j) = h_arr_in(i)
    Next
        
End Sub

Private Function IsComment(ByVal i_item As String) As Boolean
    IsComment = i_item Like ";*" And Not IsDelimiterFileHeader(i_item)
End Function

Private Function IsDelimiterFileHeader(ByVal i_item As String) As Boolean
    IsDelimiterFileHeader = i_item Like DelimiterFileHeader & "*"
End Function

Private Function IsEmptyFile(ByVal f_file_name As String, _
                    Optional ByRef f_file_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (f_file) is empty, FALSE and the content as array
' otherwise.
' ----------------------------------------------------------------------------
    Dim s As String
    
    If Fso.FileExists(f_file_name) Then
        Open f_file_name For Input As #1
        s = Input$(LOF(1), 1)
        Close #1
        IsEmptyFile = Len(s) = 0
    Else
        IsEmptyFile = True
    End If
    If Not IsEmptyFile Then
        f_file_arry = Split(FileAsString(f_file_name), vbCrLf)
    End If
    
End Function

Private Function IsSection(ByVal i_item As String, _
                  Optional ByRef i_sect_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (i_item) is a Private Profile section and it is
' equal to a provided section name (i_sect_name). When no section name
' (i_sect_name) is provided the function returns TRUE when the item (i_item)
' is a Private Profile Section.
' ------------------------------------------------------------------------------
    Dim sSection As String
    
    If Left(i_item, 1) = "[" And Right(i_item, 1) = "]" Then
        sSection = Replace(Replace(i_item, "[", vbNullString), "]", vbNullString)
        If i_sect_name <> vbNullString Then
            If sSection = i_sect_name Then
                IsSection = True
                
            End If
        Else
            IsSection = True
            i_sect_name = sSection
        End If
    End If
    
End Function

Private Function IsTextFile(ByVal i_file_full_name As String) As Boolean
    
    Dim iFileNumber As Integer
    Dim sContent    As String
    Dim sLine       As String
    Dim i           As Long
    
    '~~ Read file into a string without vbCrLf (which is a non-printable character)
    On Error Resume Next
    sContent = Replace(FileAsString(i_file_full_name), vbCrLf, vbNullString)
    If Err.Number <> 0 Then GoTo xt
    ' Check if the file content contains non-printable characters
    For i = 1 To Len(sContent)
        If Asc(Mid(sContent, i, 1)) < 32 Or Asc(Mid(sContent, i, 1)) > 126 _
        Then GoTo xt ' not an ASCII character
    Next i
    
    IsTextFile = True
xt:

End Function

Private Function IsValidFileActive(ByVal i_file_name As String, _
                          Optional ByRef i_file_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (i_file) is not empty an not contains lines other
' than Comment, Section, and Value and at least one Section and one Value.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileActive"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim i           As Long
    Dim bHasSect    As Boolean
    Dim bHasValu    As Boolean
    
    If Not Fso.FileExists(i_file_name) Then Exit Function
    If IsEmptyFile(i_file_name, arr) Then Exit Function
    
    For i = LBound(arr) To UBound(arr)
        Select Case True
            Case IsSection(arr(i)): bHasSect = True
            Case IsValue(arr(i)):   bHasValu = True
            Case Left(arr(i), 1) = ";"
            Case Else
                '~~ The file has a line which nor conforms with Private Profile file rules
                IsValidFileActive = False
                GoTo xt
        End Select
    Next i
    If bHasSect And bHasValu Then
        IsValidFileActive = True
        i_file_arry = arr
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsValidFileName(ByRef i_file_name As String, _
                        Optional ByRef i_arr As Variant, _
                        Optional ByRef i_valid_active As Boolean, _
                        Optional ByRef i_file_header As Collection = Nothing, _
                        Optional ByRef i_file_footer As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a provided string (i_file_name) is a vaild Private Profile
' file, which is: - The file does not exist (but may be created)
'                 - or the file is empty
'                 - or the file is not empty and has no other lines but
'                   Section, Value, or Comment
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileName"
    
    On Error GoTo eh
    Dim arr As Variant
    Dim cll As Collection
    Dim cllHeader As New Collection
    Dim cllFooter As New Collection
    
    If i_file_name = vbNullString And sFileName <> vbNullString Then
        '~~ The file has already been checkedd as valid
        i_file_name = sFileName
        IsValidFileName = True
        If Fso.FileExists(i_file_name) Then
            i_valid_active = Not IsEmptyFile(i_file_name)
        End If
        GoTo xt
     End If
     
    '~~ An explicitely specified file name is checked for validity and when valid
    '~~ becomes the FileName for all subsequent properties and methods which did not
    '~~ hae a file name specified as argument.
    IsValidFileName = True
    With Fso
        Select Case True
            Case Not .FileExists(i_file_name)
                '~~ The file is valid when it may be created (will immediately be deleted again however)
                On Error Resume Next
                .CreateTextFile i_file_name
                If Err.Number = 0 _
                Then .DeleteFile i_file_name _
                Else IsValidFileName = False
                
            Case IsEmptyFile(i_file_name)
            
            Case IsValidFileActive(i_file_name)
            
            Case Else: IsValidFileName = False
        End Select
    End With
    
xt: If IsValidFileName Then
        i_arr = Split(FileAsString(i_file_name), vbCrLf)
        i_valid_active = True
        If Not i_file_header Is Nothing Then Set i_file_header = HeaderFile(i_file_name)
        If Not i_file_footer Is Nothing Then Set i_file_footer = FileFooter(i_file_name)
    End If
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Function

Private Function IsValidSectName(ByRef i_section As String) As Boolean
    
    If i_section = vbNullString Then
        i_section = sSection
        IsValidSectName = True
    Else
        With Reg
            .Pattern = "^[a-zA-Z0-9_]+$"
            IsValidSectName = .Test(i_section)
        End With
    End If
    
End Function

Private Function IsValidValuName(ByVal i_name As String) As Boolean
    With Reg
        .Pattern = "^[a-zA-Z0-9_]{1,32}$"
        IsValidValuName = .Test(i_name)
    End With
End Function

Private Function IsValue(ByVal i_item As String, _
                Optional ByRef i_valu_name As String = vbNullString, _
                Optional ByRef i_valu As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (i_item) is a Private Profile value name, when a
' value name (i_value_name) is provided, TRUE only when it is the provided
' value name.
' ------------------------------------------------------------------------------
    Dim sName   As String
    Dim sValue  As String
    
    If Left(i_item, 1) = "[" Or Left(i_item, 1) = ";" Then Exit Function
    
    sName = Split(i_item, "=")(0)
    sValue = Replace(i_item, sName & "=", vbNullString)
    If i_valu_name <> vbNullString Then
        If i_valu_name = sName Then
            IsValue = True
            i_valu = sValue
        End If
    Else
        IsValue = True
        i_valu_name = sName
        i_valu = sValue
    End If
    
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, item:=k_dct.item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub NamesRemove(ByVal names_value As Variant, _
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_file As String = vbNullString, _
              Optional ByRef names_removed As Long)
' -----------------------------------------------------------------------------
' PrivateProfile file service. Removes from a PrivateProfile file's (name_file)
' section (name_section) a name (names_value) provided either as a
' comma delimited string, or as one string only.
' When no file (name_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' When the name existed and has been removed, the function returns TRUE.
' -----------------------------------------------------------------------------
    Const PROC = "NamesRemove"
    
    On Error GoTo eh
    Dim v   As Variant
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    
    For Each v In Split(Replace(names_value, " ", vbNullString), ",")
        If ValueNameExists(name_file, name_section, v) Then
            DeletePrivateProfileKey Section:=name_section _
                                  , Key:=v _
                                  , Setting:=0 _
                                  , Name:=name_file
            names_removed = names_removed + 1
        End If
    Next v
    If names_removed > 0 Then Set dctSections = Null
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ValueNameExists(ByVal p_file_name As String, _
                                 ByVal p_sect_name As String, _
                                 ByVal p_valu_name As String, _
                        Optional ByVal p_valu_head As Collection, _
                        Optional ByRef p_valu As String, _
                        Optional ByRef p_sect_coll As Collection, _
                        Optional ByRef p_file_coll As Collection) As Boolean
' ----------------------------------------------------------------------------
' When a named value (p_value_name) exists in a provided section (p_sect_name)
' in a provided Private Profile file (p_file_name) the function returns TRUE and
' the value as Dictionary (p_value_dict).
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameExists"
            
    Dim dctFile As Dictionary
    Dim dctSect As Dictionary
    
    FileAsDictMan p_file_name ' in case the file has changed or the file does not exist as Directory structure
    
    If dctFiles.Exists(p_file_name) Then
        Set p_file_coll = dctFiles(p_file_name)
        Set dctFile = p_file_coll(2)
        If dctFile.Exists(p_sect_name) Then
            Set p_sect_coll = dctFile(p_sect_name)
            Set dctSect = p_sect_coll(2)
            If dctSect.Exists(p_valu_name) Then
                ValueNameExists = True
                Set p_valu_head = dctSect(p_valu_name)(1)
                p_valu = dctSect(p_valu_name)(2)
            End If
            Set dctSect = Nothing
        End If
        Set dctFile = Nothing
    End If

End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-Private-Profile"
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Public Sub Reorg(Optional name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Reorganizes a Private Profile file (name_file) resulting in sections and
' values in ascending order by considering section and value headers.
' When no file is provided it defaults to the file name specified with the
' property FileName.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Reorg"
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    FileAsDictMan name_file
    
End Sub

Private Function RevertOrder(ByVal r_cll As Collection) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection (r_cll) with its items in reverse order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim i   As Long
    
    For i = r_cll.Count To 1 Step -1
        cll.Add r_cll(i)
    Next i
    Set RevertOrder = cll
    
End Function

Private Function SectionExists(ByVal s_arr As Variant, _
                               ByVal s_sect_name As String, _
                      Optional ByRef s_sect_indx As Long, _
                      Optional ByRef s_sect_head As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim i   As Long
    
    If Not ArrayIsAllocated(s_arr) Then GoTo xt
    
    For i = LBound(s_arr) To UBound(s_arr)
        If IsSection(s_arr(i), s_sect_name) Then
            '~~ The line is the section requested
            SectionExists = True
            s_sect_indx = i
            Set s_sect_head = HeaderOfItem(s_arr, i) ' Return the section's header comments (may be none)
            GoTo xt
        End If
    Next i

xt:

End Function

Public Function SectionNames(Optional ByVal name_file As String, _
                             Optional ByRef names_cll As Collection, _
                             Optional ByRef names_dct As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns all section names in a Private Profile file (name_file)
' - As Dictionary
' - As Dictionary (names_ddct)
' - As Collection
' with all Section names in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim arr As Variant
    Dim cll As New Collection
    Dim dct As New Dictionary
    Dim i   As Long
    Dim v   As Variant
    Dim s   As String
    If Not IsValidFileName(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
    With dct
        For i = LBound(arr) To UBound(arr)
            s = vbNullString
            If IsSection(arr(i), s) Then
                .Add s, vbNullString
            End If
        Next i
    End With
    Set names_dct = KeySort(dct)
    For Each v In names_dct
        cll.Add v
    Next v
    
xt: Set names_cll = cll
    Set SectionNames = names_dct
    Set dct = Nothing
    Set cll = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SectionRemove(ByVal name_section As String, _
                Optional ByVal name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Removes a section (name_section) in a Private Profile file (name_file).
' ----------------------------------------------------------------------------
    Const PROC = "SectionRemove"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim arrOut      As Variant
    Dim cll         As New Collection
    Dim cllFile     As Collection
    Dim cllSect     As Collection
    Dim cllSectHead As Collection
    Dim dct         As Dictionary
    Dim i           As Long
    Dim iSect       As Long
    
    If Not IsValidFileName(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    
    If Not SectionExists(s_arr:=arr _
                       , s_sect_name:=name_section _
                       , s_sect_indx:=iSect _
                       , s_sect_head:=cllSectHead) Then GoTo xt
    
    '~~ Transfer all lines down to the section header
    For i = LBound(arr) To iSect - 1 - cllSectHead.Count
        ArrayAdd arrOut, arr(i)
    Next i
    '~~ Skip down to next Section or the beginn of the file footer
    i = iSect + 1
    Do While Not IsSection(arr(i)) And i <= UBound(arr) - FileFooter(name_file).Count
        i = i + 1
    Loop
    '~~ Transfer all remaining lines
    For i = i To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    
    FileFromString name_file, Join(arrOut, vbCrLf)
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SectionsCopy(ByVal name_file_source As String, _
                        ByVal name_file_target As String, _
               Optional ByVal name_sections As Variant = Nothing, _
               Optional ByVal s_merge As Boolean = False)
' ----------------------------------------------------------------------------
' Copies sections provided as names (name_sections) from a source Private
' Profile file (name_file_source) to a target Private Profile file
' (name_file_target).
' ----------------------------------------------------------------------------
    Const PROC = "SectionsCopy"
    
    On Error GoTo eh
    Dim dct   As Dictionary
    Dim vName As Variant
    Dim vSect As Variant
    Dim sSect As String
    Dim sName As String
    Dim sValu As String
    
    If Not IsValidFileName(name_file_source) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file_source)
    If Not IsValidFileName(name_file_target) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file_target)

    For Each vSect In VarItemsAsCollection(name_sections)
        sSect = vSect
        If Not IsValidSectName(sSect) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(sSect)

        If Not s_merge Then SectionRemove sSect, name_file_target
        Set dct = ValueNames(name_file_source, vSect)
        For Each vName In dct
            sName = vName
            sValu = dct(vName)
            Value(sName, sSect, name_file_target) = sValu
        Next vName
     Next vSect

xt: Set dct = Nothing
    FileFromDict name_file_target ' re-write file in order to maintain sections in ascending sequence
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SectionsRemove(Optional ByVal name_sections As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Removes a section in a Private Profile file (name_sections) - provided as
' a comma delimited string, Dictionary section names as key or a Collection of
' section names, in file (name_file).
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "SectionsRemove"
    
    On Error GoTo eh
    Dim v As Variant
    
    If Not IsValidFileName(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
     
    For Each v In VarItemsAsCollection(name_sections)
        SectionRemove name_file, v
    Next v
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ShellRun(ByVal s_string As String, _
                 Optional ByVal s_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, s_string, vbNullString, vbNullString, s_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & s_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function ValueExists(ByVal p_arr As Variant, _
                             ByVal p_sect_name As String, _
                             ByVal p_valu_name As String, _
                    Optional ByRef p_valu As String, _
                    Optional ByRef p_index As Long, _
                    Optional ByRef p_valu_head As Collection, _
                    Optional ByRef p_sect_head As Collection) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ValueExists"
    
    On Error GoTo eh
    Dim i   As Long
    
    If SectionExists(p_arr, p_sect_name, i, p_sect_head) Then
        i = i + 1
        Do While Not IsSection(p_arr(i), p_sect_name)
            If IsValue(p_arr(i), p_valu_name, p_valu) Then
                ValueExists = True
                p_index = i
                Set p_valu_head = HeaderOfItem(p_arr, i) ' Return the Value's header comments (may be none)
                Exit Do
            End If
            If i = UBound(p_arr) Then Exit Do
            i = i + 1
        Loop
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueNameRename(ByVal name_value_old As String, _
                                ByVal name_value_new As String, _
                       Optional ByVal name_sections As Variant = vbNullString, _
                       Optional ByVal name_file As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Renames in a Provate Profile file (name_file) in all provides sections
' (name_sections) an old value name (name_value_old) to a new one
' (name_value_new). When no sections (name_sections) are provided, the old
' value name is renamed in all sections. When no file (name_file) is provided,
' the file defaults to the file name specified explicitely with the property
' FileName.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRename"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim arrOut      As Variant
    Dim dctSections As New Dictionary
    Dim i           As Long
    Dim sValue      As String
    Dim v           As Variant
    
    If name_value_old = vbNullString Or name_value_new = vbNullString Then GoTo xt
    If Not IsValidFileName(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
    If name_sections <> vbNullString _
    Then Set dctSections = VarItemsAsDictionary(name_sections) _
    Else Set dctSections = SectionNames(name_file)
    
    For i = LBound(arr) To UBound(arr)
        sSection = vbNullString
        If IsSection(arr(i), sSection) Then
            ArrayAdd arrOut, arr(i)
            If dctSections.Exists(sSection) Then
                i = i + 1
                Do While Not IsSection(arr(i))
                    If IsValue(arr(i), name_value_old, sValue) Then
                        ArrayAdd arrOut, name_value_new & "=" & sValue
                        ValueNameRename = True
                    Else
                        ArrayAdd arrOut, arr(i)
                    End If
                    If i = UBound(arr) Then Exit Do
                    i = i + 1
                Loop
            End If
        Else
            ArrayAdd arrOut, arr(i)
        End If
    Next i
        
xt: If ValueNameRename Then
        FileFromString name_file, Join(arrOut, vbCrLf)
        FileAsDictMan name_file ' Re-organize the Private Profile file
    End If
    Set dctSections = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueNames(Optional ByVal name_file As String = vbNullString, _
                           Optional ByVal name_sections As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all value names a Private Profile file (name_file)
' with the value name as the key and the value as the item.
' When no section (name_sections) is specified the result is all value names in
' any section, if a section name is provided, only the value names is that
' section. When the file name (name_file) is ommitted it defaults to the name
' specified by the FileName property.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim dct         As New Dictionary
    Dim i           As Long
    Dim sName       As String
    Dim sSection    As String
    Dim sValue      As String
    Dim v           As Variant
    
    If Not IsValidFileName(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_sections) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_sections)
    
    If name_sections = vbNullString Then
        '~~ Collect value names of all sections
        For i = LBound(arr) To UBound(arr)
            sName = vbNullString
            If IsValue(arr(i), sName, sValue) Then
                If Not dct.Exists(sName) Then
                    dct.Add sName, sValue
                End If
            End If
        Next i
    Else
        For Each v In VarItemsAsCollection(name_sections)
            sSection = v
            For i = LBound(arr) To UBound(arr)
                If IsSection(arr(i), sSection) Then
                    i = i + 1
                    Do While Not IsSection(arr(i))
                        sName = vbNullString ' is not requested as a specific value name
                        If IsValue(arr(i), sName, sValue) Then
                            If Not dct.Exists(arr(i)) Then
                                dct.Add sName, sValue
                            End If
                        End If
                        If i = UBound(arr) Then Exit Do Else i = i + 1
                    Loop
                End If
            Next i
        Next v
    End If
    
xt: Set ValueNames = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ValueRemove(ByVal name_value As String, _
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Removes a value name (name_value) from a section (name_section) in a
' Private Profile file (name_file) by considering a possible value header.
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC = "ValueRemove"
    
    Dim arr     As Variant
    Dim arrOut  As Variant
    Dim cll     As New Collection
    Dim iValue  As Long
    Dim i       As Long
    
    If Not IsValidFileName(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidValuName(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
                       
    If Not ValueExists(arr, name_section, name_value, , iValue, cll) Then GoTo xt
    '~~ Transfer all lines down to the value's header
    For i = LBound(arr) To iValue - 1 - cll.Count
        ArrayAdd arrOut, arr(i)
    Next i
    '~~ Transfer all remaining lines by skipping the value to be removed
    For i = iValue + 1 To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    '~~ Re-write Private Profile file
    FileFromString name_file, Join(arrOut, vbCrLf)

xt:
End Sub

Private Sub VarItems(Optional ByVal v_var As Variant = Nothing, _
                     Optional ByRef v_arr As Variant, _
                     Optional ByRef v_cll As Collection = Nothing, _
                     Optional ByRef v_dct As Dictionary = Nothing)
' ----------------------------------------------------------------------------
' Universal/common service providing variant items as Array (v_arr), as
' Collection (v_cll), and as Dictionary with the item as the key. The variant
' items (v_var) may be an array, a Collection of strings, a Dictionary of keys,
' or a string delimited by: vbCrLf, vbLf, ||, |, or a , (comma).
' ----------------------------------------------------------------------------
    Const PROC = "VarItems"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim v       As Variant
    Dim arr     As Variant
    
    Select Case TypeName(v_var)
        Case "String":  If v_var = vbNullString Then GoTo xt
                        Select Case True
                            Case InStr(v_var, vbCrLf) <> 0: sSplit = vbCrLf
                            Case InStr(v_var, vbLf) <> 0:   sSplit = vbLf
                            Case InStr(v_var, "||") <> 0:   sSplit = "||"
                            Case InStr(v_var, "|") <> 0:    sSplit = "|"
                            Case InStr(v_var, ",") <> 0:    sSplit = ","
                        End Select
                        For Each v In Split(v_var, sSplit)
                            ArrayAdd arr, VBA.Trim$(v)
                        Next v
        Case "Collection", "Dictionary", "Array"
                        For Each v In v_var
                            ArrayAdd arr, VBA.Trim$(v)
                        Next v
        Case Else:      Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a String, an Array, a Collecton, nor a Dictionary!"
    End Select
    For Each v In arr
        dct.Add v, vbNullString
        cll.Add v
    Next v
    
xt: v_arr = arr
    Set v_cll = cll
    Set v_dct = KeySort(dct)
    Set cll = Nothing
    Set dct = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function VarItemsAsArray(ByVal v_var As Variant, _
                        Optional ByRef v_cll As Collection = Nothing, _
                        Optional ByRef v_dct As Dictionary = Nothing) As Variant
' ----------------------------------------------------------------------------
' Provides variant items (v_var) as Array (v_arr) and additionally as
' Collection (v_cll) and as Dictionary (v_dct) with the item as the key. The
' variant items (v_var) may be an Array, a Collection of strings, a Dictionary
' of keys, or a string delimited by: vbCrLf, vbLf, ||, |, or a , (comma).
' ----------------------------------------------------------------------------
    Dim arr As Variant
    
    VarItems v_var, arr, v_cll, v_dct
    VarItemsAsArray = arr
        
End Function

Private Function VarItemsAsCollection(ByVal v_var As Variant, _
                             Optional ByRef v_arr As Variant = Nothing, _
                             Optional ByRef v_dct As Dictionary = Nothing) As Collection
' ----------------------------------------------------------------------------
' Provides variant items (v_var) as Collection and additionally a Dictionary
' (v_dct) and array (v_arr). The variant items (v_var) may be an array, a
' Collection of strings, a Dictionary of keys, or a string delimited by:
' vbCrLf, vbLf, ||, |, or a , (comma).
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    VarItems v_var, v_arr, cll, v_dct
    Set VarItemsAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function VarItemsAsDictionary(ByVal v_var As Variant, _
                             Optional ByRef v_arr As Variant, _
                             Optional ByRef v_cll As Collection = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Provides variant items (v_var) as Dictionary with the items as key and
' additionally as Array (v_arr) and Collection (v_cll). The variant items
' (v_var) may be an Array, a Collection of strings, a Dictionary of keys, or a
' string delimited by: vbCrLf, vbLf, ||, |, or a , (comma).
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    VarItems v_var, v_arr, v_cll, dct
    Set VarItemsAsDictionary = dct
    Set dct = Nothing
        
End Function

Private Function VarItemsAsString(ByVal v_var As Variant, _
                         Optional ByVal v_delim As String) As String
' ----------------------------------------------------------------------------
' Provides variant items (v_var) as String delimited with (v_delim), whereby
' the variant items (v_var) may be an Array, a Collection of strings.
' ----------------------------------------------------------------------------
    Dim cll     As Collection
    Dim dct     As Dictionary
    Dim s       As String
    Dim sDelim  As String
    Dim v       As Variant
    
    Select Case TypeName(v_var)
        Case "Collection":  Set cll = v_var
                            For Each v In cll
                                s = s & sDelim & v
                                sDelim = v_delim
                            Next v
        Case "Dictionary":  Set dct = v_var
                            For Each v In dct
                                s = s & sDelim & v
                                sDelim = v_delim
                            Next v
    End Select
    VarItemsAsString = s
    
End Function

