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
' FileName    r/w Specifies/returns a Private Profile file full name
' Footer      r/w Reads/writes a file footer
' Comment      r/w Reads(writes) a file, section, or value comment
' Section     r/w Specifies/returns the current Section name used
'                 for subsequent services which have the section argument
'                 omitted.
' Sections    r   Returns a list of all sections in a Provate Profile fle
' Value       r/w Reads/writes a value from/to a Private Profile File,
'                 optionally reads/writes a value comment
'
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
' VarTrans        Universal service providing variant items - which may be an
'                 Array of items, a Collection of items, a Dictionary of keys,
'                 a TestStream, or a string with items delimited by: vbCrLf,
'                 vbLf, ||, |, or a , (comma)as Array (v_arr) - as Array,
'                 Collection, Dictionary (with the item as the key),
'                 TextStream, or as String delimited by vbCrLf.
'                 Note: When as Array is requested but one of the (v_var)
'                       items is an object, an error is raised and the
'                       function is terminated.
'
' Requires:
' ---------
' Reference to Microsoft Scripting Runtine
' Reference to Microsoft VBScript Regular Expressions 5.5
'
' Uses no other components. Will use optionally mErH, fMsg/mMsg when installed
' and activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
'
' W. Rauschenberger, Berlin Apr 2024
' See also https://github.com/warbe-maker/VBA-Private-Profile.
' ----------------------------------------------------------------------------
Public FSo                      As FileSystemObject

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

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal Name As String) As Long
                 
Private Declare PtrSafe Function apiShellExecute _
                Lib "shell32.dll" Alias "ShellExecuteA" _
               (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private dctPPtwinsAsString      As Dictionary ' file as string
Private dctSections             As Dictionary ' section name as key, item is dctValues
Private dctValues               As Dictionary ' value name as key, item is value
Private Reg                     As RegExp
Private sFileName               As String
Private sSection                As String
Private sDelimCharHeader        As String
Private sDelimCharFooter            As String
Private bBooleanAsTrueFalseString   As Boolean
Private bSectionsSeparated          As Boolean

' ----- Directory/Collections structure of a Private Profile file's twin ------
' ---------- Element ----------|------------------- Implementation ---------------
' Type       Name              | Key          Items                  Type
' ---------- ------------------|------------- ---------------------- -------------
' Dictionary dctPPtwins        | file-name    cllPPtwin              Collection
' Collection cllPPtwin         |              cllPPtwinHeader        Collection *)
'                              |              dctPPtwinSections      Dictionary
'                              |              cllPPtwinFooter        Collection *)
' Dictionary dctPPtwinSections | section-name cllPPtwinSection       Collection
' Collection cllPPtwinSection  |              cllWinSectComments     Collection *)
'                              |              dctPPtwinSection       Dictionary
' Dictionary dctPPtwinSection  | value-name   cllPPtwinValue         Collection
' Collection cllPPtwinValue    |              cllPPtwinValueComments Collection *)
'                              |              value                  String
'
' *) Comment collections may possibly be empty (.Count = 0).

' ------------------------------------------------------------------------
' Private elements of Private Profile file twins structure used throughout
' this component for twin actions (the indentaions depict the structure)
Dim dctPPtwins                                     As Dictionary    ' Directory of Private Profile file's twins with the file name as key
    Dim cllPPtwin                                  As Collection    ' item in dctPPtwins representing a file's twin with the file-name as key
        Dim cllPPtwinHeader                        As Collection    ' cllPPtwin item(1) in cllPPtwin representing the Private Profile file's header
        Dim dctPPtwinSections                      As Dictionary    ' cllPPtwin item(2) in cllPPtwin representing the sections with the section-name as key
            Dim cllPPtwinSection                      As Collection ' cllPPtwin item in dctPPtwinSections representing a section with the section-name as key
                Dim cllPPtwinSectionComments          As Collection ' cllPPtwinSection item(1) in cllPPtwinSection representing the section's comments
                Dim dctPPtwinSection                  As Dictionary ' cllPPtwinSection item(2) in cllPPtwinSection representing a value with the value-name as key
                    Dim cllPPtwinValue             As Collection
                        Dim cllPPtwinValueComments As Collection    ' cllPPtwinValue item(1) in cllPPtwinValue representing the value comments
                        Dim sPPvalue               As String        ' cllPPtwinValue item(2) in cllPPtwinValue representing the value
        Dim cllPPtwinFooter                        As Collection    ' cllPPtwin item(3) in cllPPtwin representing the Private Profile file's footer
' ------------------------------------------------------------------------

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    Set Reg = New RegExp
    Set FSo = New FileSystemObject
    Set dctPPtwins = New Dictionary
    Set dctPPtwinsAsString = New Dictionary
    
    '~~ Initialize the default Private Profile File name
    Select Case Application.Name
        Case "Microsoft Excel": sFileName = ThisWorkbook.Path & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
        Case Else:              Err.Raise AppErr(1), ErrSrc(PROC), "Initializing the FileName failed!" & vbLf & _
                                                                   "The current Application.Name is not Excel " & _
                                                                   "and any other application is yet not supported!"
    End Select
    sDelimCharHeader = "="
    sDelimCharFooter = "-"
    bBooleanAsTrueFalseString = True
    bSectionsSeparated = True
    
End Sub

Private Sub Class_Terminate()
    Set FSo = Nothing
    Set dctPPtwins = Nothing
End Sub

Public Property Let BooleanAsTrueFalseString(ByVal b As Boolean):   bBooleanAsTrueFalseString = b:  End Property

Public Property Get Comments(Optional ByVal name_file As String = vbNullString, _
                           Optional ByVal name_section As String = vbNullString, _
                           Optional ByVal name_value As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Comments-Get"
    
    Dim cll As New Collection
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
                                    
    Select Case True
        Case name_section = vbNullString And name_value = vbNullString:   Set cll = CommentsFile(name_file)
        Case name_section <> vbNullString And name_value = vbNullString:  Set cll = CommentsSection(name_section, name_file)
        Case name_section <> vbNullString And name_value <> vbNullString: Set cll = CommentsValue(name_value, name_section, name_file)
    End Select
    
    Set Comments = cll
    Set cll = Nothing
    
End Property

Public Property Let Comments(Optional ByVal name_file As String = vbNullString, _
                             Optional ByVal name_section As String = vbNullString, _
                             Optional ByVal name_value As String = vbNullString, _
                                      ByVal h_comments As Variant)
' ------------------------------------------------------------------------------
' Adds/replaces:
' - a file header when either only a file name or no arguments at all are
'   provided
' - a section comment when a section but no value name is provided
' - a value comment when a section and a value-name is provided
' The comments (h_comments) may be a series of delimited strings or a Collection
' of strings.
' ------------------------------------------------------------------------------
    Const PROC = "Comments-Let"
    
    On Error GoTo eh
    Dim arr             As Variant
    Dim sInvalidLine    As String
    
    If name_file = vbNullString Then name_file = sFileName
    PPmanage name_file
    If name_section = vbNullString And name_value = vbNullString Then
        If IsEmptyFile(name_file) Then
            '~~ A file header comment may be written even when the file is empty
            CommentsFile(name_file) = h_comments
        Else
            '~~ A non-empty file requires a valid content to write a file header comment
            If IsValidFileContent(name_file, arr, sInvalidLine) _
            Then CommentsFile(name_file) = h_comments _
            Else Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileContent(name_file, sInvalidLine)
        End If
    Else
        '~~ If it is not a file comment which is to be written the file must be an active Private Profile file
        If Not IsValidFileContent(name_file, arr, sInvalidLine) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileContent(name_file, sInvalidLine)
        Select Case True
            Case name_section <> vbNullString And name_value = vbNullString
                If SectionExists(arr, name_section) _
                Then CommentsSection(name_section, name_file) = h_comments
            Case name_section <> vbNullString And name_value <> vbNullString
                If ValueExists(arr, name_section, name_value) _
                Then CommentsValue(name_value, name_section, name_file) = h_comments
        End Select
    End If
    
    PPmanage name_file ' ensure ascending order
                                    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let CommentsFile(Optional ByVal h_file_name As String = vbNullString, _
                                           ByVal h_file_comment As Variant)
' ----------------------------------------------------------------------------
' Replaces any existing file comments with new ones (h_file_comment),
' whereby the new comments (h_file_comment) may be provided as a series
' of strings delimited by vbCrLf, vbLf, "|" or "||", a Collection of
' strings, or a Dictionary with the comments a key.
' ----------------------------------------------------------------------------
    
    Dim cllCommentsOld  As Collection
    Dim cllCommentsNew  As Collection
    Dim arr             As Variant
    Dim arrOut          As Variant
    Dim i               As Long
    Dim v               As Variant
    Dim lLenMax         As Long
    Dim sComment        As String
    
    Set cllCommentsOld = CommentsFile(h_file_name)
    Set cllCommentsNew = VarAsCollection(h_file_comment)
    If cllCommentsOld.Count = 0 And cllCommentsNew.Count = 0 Then Exit Property
        
    '~~ Insert the new file comments excluding a possibly provided comment delimiter
    '~~ and subsequently insert a comment delimiter adjusted to the longest haeder comment
    For Each v In cllCommentsNew
        If Not IsCommentsFileDelimiter(v) Then
            sComment = Comment(v)
            lLenMax = Max(lLenMax, Len(sComment) - 2)
            ArrayAdd arrOut, sComment
        End If
    Next v
    ArrayAdd arrOut, CommentsFileDelimiter(lLenMax)
    
    '~~ Transfer all lines except the old file comment
    If FSo.FileExists(h_file_name) Then
        arr = FileAsArray(h_file_name)
        For i = cllCommentsOld.Count To UBound(arr)
            ArrayAdd arrOut, arr(i)
        Next i
    End If
    
    FileFromString h_file_name, Join(arrOut, vbCrLf)

End Property

Private Property Get CommentsFile(Optional ByVal h_file_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns a Collection with the file comment lines (may be empty). The provided
' (h_file_name) may be an array already with the file's content.
' ----------------------------------------------------------------------------
    
    Dim arr         As Variant
    Dim cll         As New Collection
    Dim i           As Long
    Dim j           As Long
    
    Set CommentsFile = cll
    If Not FSo.FileExists(h_file_name) Then GoTo xt
    arr = FileAsArray(h_file_name)
    For i = UBound(arr) To LBound(arr) Step -1
        If IsCommentsFileDelimiter(arr(i)) Then
            cll.Add arr(i)
            For j = i - 1 To LBound(arr) Step -1
                cll.Add arr(j)
            Next j
            Set CommentsFile = RevertOrder(cll)
            GoTo xt
        End If
    Next i
        
xt: Set cll = Nothing
    
End Property

Private Property Get CommentsSection(Optional ByVal h_sect_name As String, _
                                   Optional ByVal h_file_name As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Returns a Collection with the section's (h_sect_name_name) comment lines
' (possibly empty).The provided (name_file) may be an array already with the
' file's content.
' ----------------------------------------------------------------------------
    
    Dim arr     As Variant
    Dim cll     As New Collection
    Dim i       As Long
    
    arr = FileAsArray(h_file_name)
    If SectionExists(arr, h_sect_name, i, cll) Then
        Set CommentsSection = cll
    End If
    
xt: Set cll = Nothing
    
End Property

Private Property Let CommentsSection(Optional ByVal h_sect_name As String, _
                                     Optional ByVal h_file_name As String = vbNullString, _
                                              ByVal h_sect_comments As Variant)
' ----------------------------------------------------------------------------
' Replaces in file (h_file_name) an old section comment - if any - with the a
' new one (h_comment) - which may be a series of strings delimited by vbCrLf,
' vbLf, "|" or "||", or a Collection of strings.
' ----------------------------------------------------------------------------
    Const PROC = "CommentsSection-Let"
    
    On Error GoTo eh
    Dim cllCommentsOld As Collection
    Dim i            As Long
    Dim k            As Long
    Dim arr          As Variant
    Dim v            As Variant
    Dim arrOut       As Variant
    Dim bSepAdded    As Boolean
    
    arr = FileAsArray(h_file_name)
    If Not SectionExists(arr, h_sect_name, k, cllCommentsOld) Then GoTo xt
    
    '~~ Transfer down to section excluding an old comment
    For i = LBound(arr) To k - 1 - cllCommentsOld.Count
        ArrayAdd arrOut, arr(i)
    Next i
    
    '~~ Insert new comment
    For Each v In VarAsCollection(h_sect_comments)
        If bSectionsSeparated And Not bSepAdded Then
            If Len(Trim(v)) <> 0 Then
                ArrayAdd arrOut, vbNullString
                bSepAdded = True
            End If
        End If
        ArrayAdd arrOut, Comment(v)
    Next v
    
    '~~ Transfer all remaining lines
    For i = k To UBound(arr)
        ArrayAdd arrOut, arr(i)
    Next i
    
    StringAsFile Join(arrOut, vbCrLf), h_file_name
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get CommentsValue(Optional ByVal h_valu_name As String, _
                                 Optional ByVal h_sect_name As String = vbNullString, _
                                 Optional ByVal h_file_name As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection with the value's (h_sect_name) comment lines (may be
' empty). The provided (name_file) may be an array already with the file's
' content.
' ----------------------------------------------------------------------------
    Const PROC = "CommentsValue-Get"
                                  
    On Error GoTo eh
    Dim arr     As Variant
    Dim cll     As New Collection
    Dim i       As Long
    Dim sValue  As String
    
    Set CommentsValue = cll
    arr = FileAsArray(h_file_name)
    If ValueExists(arr, h_sect_name, h_valu_name, sValue, i, cll) Then
        Set CommentsValue = cll
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let CommentsValue(Optional ByVal h_valu_name As String, _
                                   Optional ByVal h_sect_name As String = vbNullString, _
                                   Optional ByVal h_file_name As String = vbNullString, _
                                            ByVal h_valu_comments As Variant)
' ----------------------------------------------------------------------------
' Replaces in an - optionally provided file's - section's (s_section) value
' (h_value_name) a comment with a provided comment (h_comment) which may be
' provided as a series of strings delimited by a | (vertical bar) or as a
' collection of strings.
' ----------------------------------------------------------------------------
    Const PROC = "CommentsValue-Let"
    
    On Error GoTo eh
    Dim arrOut          As Variant
    Dim arr             As Variant
    Dim i               As Long
    Dim j               As Long
    Dim cllCommentsOld  As Collection
    Dim sValue          As String
    Dim v               As Variant
    
    arr = FileAsArray(h_file_name)
    If Not IsValidSectName(h_sect_name) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(h_sect_name)
    If Not ValueExists(arr, h_sect_name, h_valu_name, sValue, j, cllCommentsOld) Then GoTo xt
    
    '~~ Transfer all lines down to value except a possible old comment
    For i = LBound(arr) To j - 1 - cllCommentsOld.Count
        ArrayAdd arrOut, arr(i)
    Next i
    '~~ Insert the new comment
    For Each v In VarAsCollection(h_valu_comments)
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

Public Property Get FileName() As String:               FileName = sFileName:       End Property

Public Property Let FileName(ByVal f_file As String)
' ----------------------------------------------------------------------------
' Note: Only a FileName explicitely provided as Property becomes the default
'       file name for an omitted file argument.
'       Any file name provided as argument along with a property or method
'       call is only valid for the call.
' ----------------------------------------------------------------------------
    Const PROC = "FileName-Let"
    
    Dim sInvalidLine    As String
    Dim arr             As Variant
    
    BoP ErrSrc(PROC)
    Select Case True
        Case Not IsValidFileName(f_file): Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(f_file)
        Case Not IsEmptyFile(f_file) _
         And Not IsValidFileContent(f_file, arr, sInvalidLine): Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidFileContent(f_file, sInvalidLine)
    End Select
    sFileName = f_file
    PPmanage f_file

xt: EoP ErrSrc(PROC)

End Property

Public Property Get Footer(Optional ByVal name_file As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' When the file has no footer an empty Collection is returned. A footer is
' at least one comment line with a footer delimiter lin above.
' ----------------------------------------------------------------------------
    Const PROC = "Footer-Get"
    
    Dim bActive As Boolean
    Dim cll     As New Collection
    Dim arr     As Variant
    
    Set Footer = cll
    If Not IsValidFile(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    Set Footer = PPfooterGet(arr)
    
End Property

Public Property Let Footer(Optional ByVal name_file As String = vbNullString, _
                                    ByVal file_footer As Variant)
' ----------------------------------------------------------------------------
' Replaces an old footer in file (name_file) - if any - with the new footer
' (file_footer) - which may be a series of strings delimited by vbCrLf, vbLf,
' "|" or "||".
' ----------------------------------------------------------------------------
    Const PROC = "Footer-Let"
    
    Dim arr             As Variant
    Dim arrOut          As Variant
    Dim bActive         As Boolean
    Dim cllFooterNew    As Collection
    Dim i               As Long
    Dim v               As Variant
    Dim lLenMax         As Long
    Dim cllFooterOld    As Collection
    
    If Not IsValidFile(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If ArrayIsAllocated(arr) Then
        Set cllFooterOld = PPfooterGet(name_file)
        '~~ Transfer all except the old footer (if any)
        For i = LBound(arr) To UBound(arr) - cllFooterOld.Count
            ArrayAdd arrOut, arr(i)
        Next i
    End If
    
    '~~ Add the file footer in case one has been provided
    Set cllFooterNew = VarAsCollection(file_footer)
    '~~ Get the max footer length
    For Each v In cllFooterNew
        If Not IsDelimiterFileFooter(v) Then lLenMax = Max(lLenMax, Len(v))
    Next v
    '~~ Write the footer delimiter
    ArrayAdd arrOut, DelimiterFileFooter(lLenMax)
    For Each v In cllFooterNew
        If Not IsDelimiterFileFooter(v) Then ArrayAdd arrOut, Comment(v)
    Next v
    
    '~~ Re-write Private Profile File with new/replaced file comment if any
    FileFromString name_file, Join(arrOut, vbCrLf)

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
' Returns a Dictionary with all named sections (pp_sections) in file (pp_file)
' - provided as full name string or as file object - whereby each section is
' a Dictionary with the section name as the key - in ascending order - and a
' Dictionary of the section's values as item with the value name as key and
' the value as item.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim dct         As New Dictionary
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
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
    Dim vValue      As Variant
        
    For Each vSection In s_dct
        sSection = vSection
        Set dctValues = s_dct(sSection)
        For Each vName In dctValues
            sName = vName
            vValue = dctValues.item(vName)
            ValueLet name_file:=name_file _
                   , name_section:=sSection _
                   , name_value:=sName _
                   , v_value:=vValue
        Next vName
    Next vSection
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let SectionsSeparated(ByVal b As Boolean):          bSectionsSeparated = b:         End Property

Public Property Get Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByRef comments_value As Variant = vbNullString, _
                          Optional ByRef comments_section As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Returns value name's (name_value) value from a section (name_section) in a
' Private Profile file (name_file) - plus its comment (it any) as a string of
' comment lines delimited with a "||".
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Value-Get"
    
    Dim lResult As Long
    Dim sRetVal As String
        
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
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
    '~~ Return value/section comments if any
    Set comments_value = VarAsCollection(CommentsValue(name_value, name_section, name_file))
    Set comments_section = VarAsCollection(CommentsSection(name_section, name_file))
    
End Property

Public Property Let Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByRef comments_value As Variant = vbNullString, _
                          Optional ByRef comments_section As Variant, _
                                   ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value (v_value) under a given name (name_value) into a section
' (name_section) in a Private Profile file (name_file).
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
        
    Dim bIsNew      As Boolean
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidValuName(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
    
    bIsNew = Not ValueExists(name_file, name_section, name_value)
    ValueLet name_file, name_section, name_value, v_value
    CommentsValue(name_value, name_section, name_file) = comments_value
    If Not IsMissing(comments_section) Then
        '~~ Attention! Non provided section comments do not affect any current section comments
        CommentsSection(name_section, name_file) = comments_section
    End If
    If bIsNew Then
        PPmanage name_file ' re-write the file in order to maintain ascending value name sequence
    End If
    
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

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

Private Function ArrayAsCollection(ByVal a_array As Variant) As Collection
' ----------------------------------------------------------------------------
' Return an array's (a_array) items as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    With cll
        For Each v In a_array
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function ArrayAsDictionary(ByVal a_array As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    Dim v As Variant
    
    With dct
        For Each v In a_array
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

Private Function ArrayAsFile(ByVal a_array As Variant, _
                   Optional ByRef a_file As Variant = vbNullString, _
                   Optional ByVal a_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes all items of an array (a_arry) to a file (a_file) which might be a
' file object, a file's full name. When no file (a_file) is provided a
' temporary file is returned, else the provided file (a_file) as object.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
      
    If Not ArrayIsAllocated(a_array) Then Exit Function
    
    Select Case True
        Case a_file = vbNullString:     a_file = TempFile
        Case TypeName(a_file) = "File": a_file = a_file.Path
    End Select
    
    If a_file_append _
    Then Open a_file For Append As #1 _
    Else Open a_file For Output As #1
    Print #1, Join(a_array, vbCrLf)
    Close #1
    Set ArrayAsFile = FSo.GetFile(a_file)
    
End Function

Private Function ArrayAsString(ByVal a_array As Variant, _
                     Optional ByVal a_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (a_array) as string with the items delimited (a_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(a_array, a_delim)
End Function

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
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Function CollectionAsArray(ByVal c_coll As Collection) As Variant
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sName   As String
    Dim v       As Variant
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then ArrayAdd arr, sName _
        Else ArrayAdd arr, v
    Next v
    CollectionAsArray = arr
    
End Function

Private Function CollectionAsDictionary(ByVal c_coll As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as Dictionary keys.
' Attention: Because the collection's items are returned as Directory keys,
'            the items will be unified. I e. each item will exist only once.
'            To make this restriction productive, the number of occurrences of
'            each item is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    Dim v As Variant
    
    With dct
        For Each v In c_coll
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

Private Function CollectionAsFile(ByVal c_coll As Collection, _
                        Optional ByRef c_file_name As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns the items of a Collection (c_coll) as records/lines in a file
' (c_file_name), optionally appended (c_file_append).
' Uses: StringAsFile, CollectionAsString.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------

    If c_file_name = vbNullString Then c_file_name = TempFile
    StringAsFile CollectionAsString(c_coll), c_file_name, c_file_append
    Set CollectionAsFile = FSo.GetFile(c_file_name)

End Function

Private Function CollectionAsString(ByVal c_coll As Collection) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    Dim sName   As String
    Dim v       As Variant
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then s = s & sDelim & sName _
        Else s = s & sDelim & v
        sDelim = vbCrLf
    Next v
    CollectionAsString = s

End Function

Private Function Comment(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s) as Provate Profile file comment line by considering
' that an empty line is a valid comment line
' ----------------------------------------------------------------------------
    If Len(Trim(s)) = 0 Then
        Comment = Trim(s)   ' accepts an empty line as a valid quasi comment
    Else
        s = Replace(";" & s, ";;", ";")
        Comment = s         ' ensures, the line begins with a single ;
        If Left(Comment, 2) <> "; " Then
            Comment = "; " & Right(Comment, Len(Comment) - 1)
        End If
    End If
    
End Function

Private Function CommentsFileDelimiter(Optional ByVal lMax As Long = 2) As String
    CommentsFileDelimiter = "; " & String(lMax, sDelimCharHeader)
End Function

Private Sub CommentsFileRestore(ByVal f_file_name As String, _
                                ByVal f_file_comment As Collection)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arrIn   As Variant
    Dim arrOut  As Variant
    Dim i       As Long
    Dim v       As Variant
    
    arrIn = FileAsArray(f_file_name)
    For Each v In f_file_comment
        ArrayAdd arrOut, v
    Next v
    For i = LBound(arrIn) To UBound(arrIn)
        ArrayAdd arrOut, arrIn(i)
    Next i
    StringAsFile Join(arrOut, vbCrLf), f_file_name
    
End Sub

Private Function CommentsFileSave(ByVal f_file_name As String) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arrIn   As Variant
    Dim arrOut  As Variant
    Dim cll     As Collection
    Dim i       As Long
    
    Set cll = CommentsFile(f_file_name)
    If cll.Count = 0 Then GoTo xt
    arrIn = FileAsArray(f_file_name)
    For i = cll.Count To UBound(arrIn)
        ArrayAdd arrOut, arrIn(i)
    Next i
    If Not ArrayIsAllocated(arrOut) Then
        FSo.DeleteFile f_file_name
        FSo.CreateTextFile f_file_name
    Else
        FileFromString f_file_name, Join(arrOut, vbCrLf)
    End If

xt: Set CommentsFileSave = cll
    Set cll = Nothing
    
End Function

Private Function CommentsOfItem(ByVal h_arr As Variant, _
                              ByRef h_idx As Long) As Collection
    Dim cll As New Collection
    
    If h_idx = LBound(h_arr) Then GoTo xt
    h_idx = h_idx - 1
    Do While IsComment(h_arr(h_idx))
        cll.Add h_arr(h_idx)
        If h_idx = LBound(h_arr) Then Exit Do
        h_idx = h_idx - 1
    Loop

xt: Set CommentsOfItem = RevertOrder(cll)
    Set cll = Nothing
    
End Function

Private Function DelimiterFileFooter(Optional ByVal lMax As Long = 2) As String
    DelimiterFileFooter = "; " & String(lMax, sDelimCharFooter)
End Function

'Private Function DictionaryAsArray(ByVal d_dict As Dictionary) As Variant
'' ----------------------------------------------------------------------------
'' Returns a Dictionary's keys as array. In case a key is an object the
'' object's Name property is returned (an error is raised in case the object
'' has no Name property).
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim arr As Variant
'    Dim s   As String
'    Dim v   As Variant
'
'    For Each v In d_dict
'        If IsObject(v, s) _
'        Then ArrayAdd arr, s _
'        Else ArrayAdd arr, v
'    Next v
'    DictionaryAsArray = arr
'
'End Function
'
'Private Function DictionaryAsCollection(ByVal d_dict As Dictionary) As Collection
'' ----------------------------------------------------------------------------
'' Returns a Collection with a Dictionary's (d_dict) keys as items.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim cll As New Collection
'    Dim v   As Variant
'
'    With cll
'        For Each v In d_dict
'            .Add v
'        Next v
'    End With
'    Set DictionaryAsCollection = cll
'    Set cll = Nothing
'
'End Function
'
'Private Function DictionaryAsDictionary(ByVal d_dict As Dictionary) As Dictionary
'' ----------------------------------------------------------------------------
'' Returns the Dictionary (d_dict) with the keys in ascending order.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim dct As New Dictionary
'    Dim v   As Variant
'
'    With dct
'        For Each v In d_dict
'            .Add v, vbNullString
'        Next v
'    End With
'    Set DictionaryAsDictionary = KeySort(dct)
'    Set dct = Nothing
'
'End Function
'
'Private Function DictionaryAsFile(ByVal d_dict As Dictionary, _
'                        Optional ByRef d_file_name As String = vbNullString, _
'                        Optional ByVal d_file_append As Boolean = False) As File
'' ----------------------------------------------------------------------------
'' Returns a Dictionary's (d_dict) keys as file (d_file_name) records/lines,
'' optionally appended (d_file_append). Keys are converted into a string, if an
'' item is an object its Name property is used. An error is raised when the
'' object has no Name property.
'' Uses: StringAsFile, DictionaryAsString.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'
'    If d_file_name = vbNullString Then d_file_name = TempFile
'    StringAsFile DictionaryAsString(d_dict), d_file_name, d_file_append
'    Set DictionaryAsFile = FSo.GetFile(d_file_name)
'
'End Function
'
'Private Function DictionaryAsString(ByVal v_items As Dictionary) As String
'' ----------------------------------------------------------------------------
'' Returns a Dictionary's keys as string with each key delimited by a vbCrLf.
'' Keys are converted into a string, if an item is an object its Name property
'' is used. An error is raised when the object has no Name property.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim s       As String
'    Dim sDelim  As String
'    Dim sName   As String
'    Dim v       As Variant
'
'    For Each v In v_items
'        If IsObject(v, sName) _
'        Then s = s & sDelim & sName _
'        Else s = s & sDelim & v
'        sDelim = vbCrLf
'    Next v
'    DictionaryAsString = s
'
'End Function

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
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrInvalidFileContent(ByVal e_file As String, _
                                       ByVal e_line As String) As String
    ErrInvalidFileContent = "The file """ & e_file & """ has an invalid content ('" & e_line & """!"
End Function

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
#If mErH = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
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
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_value As String = vbNullString, _
              Optional ByRef e_valu As String, _
              Optional ByRef head_value As Collection = Nothing, _
              Optional ByRef head_section As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided value name, section name, file name exists.
' When the provided (section (name_section
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim bActive As Boolean
    Dim i       As Long
    
    If Not IsValidFile(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
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

Private Function FileAsArray(ByVal f_file As Variant, _
                    Optional ByVal f_empty_excluded As Boolean = False) As String()
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsArray"
    
    On Error GoTo eh
    Dim arr() As String
    Dim v     As Variant
    Dim s     As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"
    
    s = FileAsString(f_file, vbCrLf)
    If f_empty_excluded Then
        For Each v In Split(s, vbCrLf)
            If Trim$(v) <> vbNullString Then ArrayAdd arr, v
        Next v
    Else
        arr = Split(s, vbCrLf)
    End If
    FileAsArray = arr
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsCollection(ByVal f_file As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as Collection.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim v   As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"

    With cll
        For Each v In FileAsArray(f_file)
            .Add v
        Next v
    End With
    Set FileAsCollection = cll
    Set cll = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsDictionary(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns a file's records/lines as Dictionary keys.
' Attention: Because the lines become Directory keys, they will become
'            distinct. I. e. each line will exist only once in the Dictionary.
'            To make this restriction productive, the number of occurrences
'            of each line is returned as item.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsDictionary"
    
    On Error GoTo eh
    Dim dct As New Dictionary
    Dim l   As Long
    Dim s   As String
    Dim v   As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"
    
    With dct
        For Each v In FileAsArray(f_file)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsFile(ByVal f_file_in As File, _
                            ByVal f_file_out As String, _
                   Optional ByVal f_append As Boolean = False, _
                   Optional ByVal f_rename As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a file (f_file_in) as a file with another full name.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsFile"
    
    On Error GoTo eh
    With FSo
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsString(f_file_in), f_file_out, True
            Case Else:                          Err.Raise AppErr(2), ErrSrc(PROC), "Rename  a n d  append is not supported!"
        End Select
    Set FileAsFile = .GetFile(f_file_out)
    End With
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsString(ByVal f_file As Variant, _
                     Optional ByRef f_split As String = vbCrLf, _
                     Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a file's (f_file) - provided as full name or object - records/lines
' as a single string with the records/lines delimited (f_split).
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    Dim s   As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), _
                                       "The file '" & f_file & "' provided does not exist!"
    
    Open f_file For Input As #1
    s = Input$(LOF(1), 1)
    Close #1
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(s, vbCr) <> 0:   f_split = vbCr
        Case InStr(s, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s

xt: Exit Function

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

Private Function IsComment(ByVal i_item As String) As Boolean
    IsComment = Len(Trim(i_item)) = 0 ' considers an empty line as a valid (separator) comment
    If Not IsComment Then
        IsComment = i_item Like ";*" And Not IsCommentsFileDelimiter(i_item) And Not IsDelimiterFileFooter(i_item)
    End If
End Function

Private Function IsCommentsFileDelimiter(ByVal i_item As String) As Boolean
    IsCommentsFileDelimiter = i_item Like CommentsFileDelimiter & "*"
End Function

Private Function IsDelimiterFileFooter(ByVal i_item As String) As Boolean
    IsDelimiterFileFooter = i_item Like DelimiterFileFooter & "*"
End Function

Private Function IsEmptyFile(ByVal f_file_name As String, _
                    Optional ByRef f_file_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (f_file) is empty, FALSE and the content as array
' otherwise.
' ----------------------------------------------------------------------------
    Dim s As String
    
    If FSo.FileExists(f_file_name) Then
        Open f_file_name For Input As #1
        s = Input$(LOF(1), 1)
        Close #1
        IsEmptyFile = Len(s) = 0
    Else
        IsEmptyFile = True
    End If
    If Not IsEmptyFile Then
        f_file_arry = FileAsArray(f_file_name)
    End If
    
End Function

Private Function IsObject(ByVal i_var As Variant, _
                          ByRef i_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
' is an object. When the object does not have a Name property an error is
' raised.
' ----------------------------------------------------------------------------
    Const PROC = "IsObject"
    
    If Not VBA.IsObject(i_var) Then Exit Function
    IsObject = True
    On Error Resume Next
    i_name = i_var.Name
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "VarTrans tries to use the Name property of an object when it is to be " & _
         "transferred into a string which is the case when String, Array, or File " & _
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
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

Private Function IsValidFile(ByRef i_file_name As String, _
                    Optional ByRef i_arr As Variant, _
                    Optional ByRef i_valid_content As Boolean, _
                    Optional ByRef i_file_comment As Collection = Nothing, _
                    Optional ByRef i_file_footer As Collection = Nothing, _
                    Optional ByRef i_error As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a provided string (i_file_name) is a vaild Private Profile
' file, which is: - The file does not exist (but may be created)
'                 - or the file is empty
'                 - or the file is not empty and has no other lines but
'                   Section, Value, or Comment
' Attention: This file validity check is restricted for being used by Public
'            methods and properties only!
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFile"
    
    On Error GoTo eh
    Dim sInvalidLine    As String
    
    BoP ErrSrc(PROC)
    If i_file_name = vbNullString And sFileName <> vbNullString Then
        '~~ The file has already been checkedd as valid
        i_file_name = sFileName
        IsValidFile = True
        If FSo.FileExists(i_file_name) Then
            i_valid_content = Not IsEmptyFile(i_file_name, i_arr)
        End If
        GoTo xt
     End If
     
    '~~ An explicitely specified file name is checked for validity and when valid
    '~~ becomes the FileName for all subsequent properties and methods which did not
    '~~ have a file name provided as argument.
    IsValidFile = True
    With FSo
        Select Case True
            Case Not IsValidFileName(i_file_name)
                IsValidFile = False
                i_error = ErrInvalidFileName(i_file_name)
                GoTo xt
            Case Not IsValidFileContent(i_file_name, i_arr, sInvalidLine)
                '~~ The file is valid when it has a valid content (may be empty however)
                IsValidFile = False
                i_error = ErrInvalidFileContent(i_file_name, sInvalidLine)
            Case Else
                i_valid_content = True
                IsValidFile = True
                If Not IsEmptyFile(i_file_name) Then
                    Set i_file_comment = CommentsFile(i_file_name)
                    Set i_file_footer = PPfooterGet(i_arr)
                    '~~ An explicitely specified valid non empty file is immediately
                    '~~ provided as twin when yet unknown or when known but the
                    '~~ specifiec Private Profile file has changed meanwhile.
                    PPmanage i_file_name
                End If
        End Select
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Function

Private Function IsValidFileContent(ByVal i_file_name As String, _
                                    ByRef i_file_arry As Variant, _
                                    ByRef i_line_nvld As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (i_file_name) either does not exist or contains
' only valid Private Profile lines else FALSE is returned and the invalid line
' (i_line_nvld).
' Note: Empty lines are a valid content.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileContent"
    
    On Error GoTo eh
    Dim arr As Variant
    Dim i   As Long
    
    If Not FSo.FileExists(i_file_name) Then
        IsValidFileContent = True ' an empty file is valid because it does not contain any invalid line
        GoTo xt
    End If
    
    arr = FileAsArray(i_file_name)
    For i = LBound(arr) To UBound(arr)
        If Not IsSection(arr(i)) _
        And Not IsValue(arr(i)) _
        And Not Left(arr(i), 1) = ";" _
        And Not Len(Trim(arr(i))) = 0 Then
            IsValidFileContent = False
            i_line_nvld = arr(i)
            GoTo xt
        End If
    Next i
    IsValidFileContent = True
    i_file_arry = arr
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsValidFileName(ByVal i_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file either exists or does not exist but may be
' created.
' ----------------------------------------------------------------------------
    With FSo
        If Not .FileExists(i_file) Then
            '~~ The file is valid when it may be created (will immediately be deleted again however)
            On Error Resume Next
            .CreateTextFile i_file
            If Err.Number = 0 Then
                .DeleteFile i_file
                IsValidFileName = True
            Else
                IsValidFileName = False
                GoTo xt
            End If
        Else
            IsValidFileName = True
        End If
    End With
xt:
End Function

Private Function IsValidSectName(ByRef i_section As String) As Boolean
    
    If i_section = vbNullString Then
        i_section = sSection ' provide the default section
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
' When the item (i_item) is a Private Profile value name
' - and no value-name is provided, the function returns TRUE and the value-name
'   (i_valu_name)
' - and a value-name has been provided, the function only returns TRUE when the
'   item (i_item) is that value-name (i_value_name).
' ------------------------------------------------------------------------------
    Dim sName   As String
    Dim sValue  As String
    
    If i_item = vbNullString Then Exit Function
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
' Returns the items in a Dictionary (k_dct) sorted by key as Dictionary plus
' replacing the provided Dictionary (k_dct) by the sorted one.
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

Private Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Public Sub NamesRemove(ByVal names_value As Variant, _
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_file As String = vbNullString, _
              Optional ByRef names_removed As Long)
' -----------------------------------------------------------------------------
' Removes from a PrivateProfile file (name_file) - which defaults to the
' FileName property provided - in a section (name_section) names provided
' either as a comma delimited string, or as one string only.
' When no file (name_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' When the name existed and has been removed, the function returns TRUE.
' -----------------------------------------------------------------------------
    Const PROC = "NamesRemove"
    
    On Error GoTo eh
    Dim v   As Variant
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    
    For Each v In VarAsArray(names_value)
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

Private Sub PPfileAsPPtwin(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Sets up a Private Profile twin as a structure of Dictionaries and
' Collections which serves as a means to keep the content including its
' comments in ascending order, for existence checks, and for remove section
' and or values.
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwin"
    
    On Error GoTo eh
    Dim arr              As Variant
    Dim cll              As New Collection
    Dim i                As Long
    Dim s                As String
    Dim sName            As String
    Dim sValue           As String
    
    BoP ErrSrc(PROC), p_file
    PPtwinRemove p_file
    arr = FileAsArray(p_file)
    Set cllPPtwinHeader = CommentsFile(p_file)
    Set cllPPtwinFooter = PPfooterGet(p_file)
    Set dctPPtwinSections = Nothing
    Set dctPPtwinSection = Nothing
    
    For i = cllPPtwinHeader.Count To UBound(arr) - cllPPtwinFooter.Count
        sName = vbNullString
        s = vbNullString
        Select Case True
            Case IsComment(arr(i))
                '~~ Collect any comment except a file comment delimiter
                cll.Add arr(i)
            
            Case IsCommentsFileDelimiter(arr(i))
                '~~ A file comment delimiter indicates the end of a file comment
                cll.Add arr(i)
                Set cllPPtwinHeader = cll
                Set cll = New Collection
            
            Case IsSection(arr(i), s)
                '~~ Stow section plus any possible comment (may be empty)
                If sSection <> vbNullString Then
                    '~~ This is a subsequent section which indicates that the
                    '~~ previous with its collected values has to be saved
                    PPfileAsPPtwinSectAdd p_sect:=sSection _
                                     , p_sect_comment:=cllPPtwinSectionComments
                End If
                sSection = s
                Set cllPPtwinSectionComments = cll
                Set cll = New Collection
            
            Case IsValue(arr(i), sName, sValue)
                '~~ Collect value in dctPPtwinSection with the value name as key
                PPfileAsPPtwinValueAdd p_value_name:=sName _
                                     , p_value:=sValue _
                                     , p_value_comments:=cll
        End Select
    Next i
    
    If dctPPtwinSection Is Nothing Then Set dctPPtwinSection = New Dictionary
    If dctPPtwinSection.Count > 0 Then
        PPfileAsPPtwinSectAdd p_sect:=sSection _
                            , p_sect_comment:=cll
    End If
        
    '~~ Stow the file with possible comment and footer in the dctPPtwins Directory
    PPfileAsPPtwinAdd p_file
    
xt: Set cll = Nothing
    Set cllPPtwin = Nothing
    Set cllPPtwinHeader = Nothing
    Set dctPPtwinSections = Nothing
    Set cllPPtwinSection = Nothing
    Set cllPPtwinSectionComments = Nothing
    Set dctPPtwinSection = Nothing
    Set cllPPtwinValue = Nothing
    Set cllPPtwinValueComments = Nothing
    Set cllPPtwinFooter = Nothing
    EoP ErrSrc(PROC), p_file
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPfileAsPPtwinAdd(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Adds a new twin which had been setup from a Private Profile file (p_file) as
' Dictionary/collections structure, to the Directory of twins (dctPPtwins)
' with the file header (cllPPtwinHeader), file sections (dctPPtwinSections),
' and file footer (cllPPtwinFooter).
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinAdd"
    
    On Error GoTo eh
    Dim cll     As New Collection
    
    If cllPPtwinHeader Is Nothing Then Set cllPPtwinHeader = New Collection
    If dctPPtwinSections Is Nothing Then Set dctPPtwinSections = New Dictionary
    If cllPPtwinFooter Is Nothing Then Set cllPPtwinFooter = New Collection
    
    cll.Add cllPPtwinHeader
    PPfileAsPPtwinAddRemoveTopSectionSeparator p_file
    cll.Add dctPPtwinSections
    cll.Add cllPPtwinFooter
    
    If dctPPtwins.Exists(p_file) Then dctPPtwins.Remove p_file
    dctPPtwins.Add p_file, cll
    Set cll = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPfileAsPPtwinAddRemoveTopSectionSeparator(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Remove the top section's speparator line.
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinAddRemoveTopSectionSeparator"
    
    On Error GoTo eh
    Dim sSect   As String
    
    If dctPPtwinSections.Count = 0 Then GoTo xt
    sSect = dctPPtwinSections.Keys()(0)                 ' get the top section's key
    PPtwinSectionGet p_file, sSect
    If cllPPtwinSectionComments(1) = vbNullString Then  ' remove a separator item if any
        cllPPtwinSectionComments.Remove 1
        PPtwinSectionLet p_file, sSect                  ' update/re-write the twin section
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPfileAsPPtwinSectAdd(ByVal p_sect As String, _
                                  ByRef p_sect_comment As Collection)
' ----------------------------------------------------------------------------
' Adds a section (p_dct_sect) to the file Dictionary (p_file_dct).
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinSectAdd"
    
    On Error GoTo eh
    Dim cll As New Collection
    
    If dctPPtwinSections Is Nothing Then Set dctPPtwinSections = New Dictionary
    If dctPPtwinSection Is Nothing Then Set dctPPtwinSection = New Dictionary
    If dctPPtwinSection.Count = 0 Then           ' empty section
        GoTo xt
    End If
    
    cll.Add PPfileAsPPtwinSectComments(p_sect_comment, p_sect)  ' transfer possible comments to the section Collection considering the separator option
    cll.Add KeySort(dctPPtwinSection)                   ' add the section as the last item to the Collection
    dctPPtwinSections.Add p_sect, cll                   ' add the Collection with the section name as key to the file Dictionary
    KeySort dctPPtwinSections
    Set p_sect_comment = New Collection
    Set dctPPtwinSection = New Dictionary
    
xt: Set cll = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PPfileAsPPtwinSectComments(ByVal p_cll As Collection, _
                                            ByVal p_sect As String) As Collection
' ----------------------------------------------------------------------------
' Sets up a sections comments by adding a top separator line when the
' respective option is TRUE.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If bSectionsSeparated Then
        If p_cll.Count = 0 Then
            '~~ No comments at all yet
            cll.Add vbNullString
            Debug.Print "Section separator added for section " & p_sect
        ElseIf Len(Trim(p_cll(1))) <> 0 Then
            '~~ First item is not a sepeartor line
            cll.Add vbNullString
            Debug.Print "Section separator added for section " & p_sect
        End If
    End If
    For Each v In p_cll
        cll.Add v
    Next v
    Set PPfileAsPPtwinSectComments = cll
    Set cll = Nothing
    
End Function

Private Sub PPfileAsPPtwinValueAdd(ByVal p_value_name As String, _
                                   ByVal p_value As String, _
                                   ByRef p_value_comments As Collection)
' ----------------------------------------------------------------------------
' Add the first and only, the last one, or a subsequent value name
' (p_value_name), value (p_value), and possible comments (p_value_comments)
' to a section.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    If dctPPtwinSection Is Nothing Then Set dctPPtwinSection = New Dictionary
    
    cll.Add p_value_comments             ' add the value comment as item to the Collection
    cll.Add p_value                      ' add the value as item to the Collection
    If Not dctPPtwinSection.Exists(p_value_name) Then
        dctPPtwinSection.Add p_value_name, cll ' add the Collection
        KeySort dctPPtwinSection                ' sort added value
    End If
    Set p_value_comments = New Collection
    Set cll = Nothing
    
End Sub

Private Function PPfileHasChanged(ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Private Profile file
' (p_file) has changed.
' ----------------------------------------------------------------------------

   PPfileHasChanged = StrComp(FileAsString(p_file) _
                            , dctPPtwinsAsString(p_file) _
                            , vbBinaryCompare) <> 0
    
End Function

Private Function PPfooterGet(ByVal p_arr As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a Private Profile file's footer comments as Collection which may be
' empty.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim i   As Long
    
    Set PPfooterGet = cll
    If Not IsArray(p_arr) Then p_arr = FileAsArray(p_arr)
    i = UBound(p_arr)
    Do While IsComment(p_arr(i)) And i >= LBound(p_arr)
        '~~ Any bottom comment up to the first non-comment is regarded a file footer
        cll.Add p_arr(i)
        If i = LBound(p_arr) Then Exit Do
        i = i - 1
    Loop
    If IsDelimiterFileFooter(p_arr(i)) Then cll.Add p_arr(i)
    
    If cll.Count > 0 _
    Then Set PPfooterGet = RevertOrder(cll)

xt: Set cll = Nothing
    
End Function

Private Sub PPfooterRestore(ByVal p_file_name As String, _
                            ByVal p_file_footer As Collection)
    CollectionAsFile p_file_footer, p_file_name, True
    
End Sub

Private Function PPfooterSave(ByVal p_file_name As String) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arrIn   As Variant
    Dim arrOut  As Variant
    Dim cll     As Collection
    Dim i       As Long
    
    Set cll = PPfooterGet(FileAsArray(p_file_name))
    arrIn = FileAsArray(p_file_name)
    For i = LBound(arrIn) To UBound(arrIn) - cll.Count
        ArrayAdd arrOut, arrIn(i)
    Next i
    If Not ArrayIsAllocated(arrOut) Then
        FSo.DeleteFile p_file_name
        FSo.CreateTextFile p_file_name
    Else
        FileFromString p_file_name, Join(arrOut, vbCrLf)
    End If
    Set PPfooterSave = cll
    Set cll = Nothing
    
End Function

Private Sub PPmanage(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Manages a Private Profile by:
' - setup a Private Profile twin as a Dictionary structure when a twin doesn't
'   exist
' - re-setup a Private Profile twin in the form of a Dictionary structure when
'   the related Private Profile file has changed
' - re-write when the twin has changed (being an exception though).
' ----------------------------------------------------------------------------
    Const PROC = "PPmanage"
    
    On Error GoTo eh
    Dim bPPtwinHasChanged As Boolean
    Dim bPPfileHasChanged As Boolean
    
    BoP ErrSrc(PROC)
    If Not FSo.FileExists(p_file) Then GoTo xt
    
    If Not PPtwinExists(p_file) Then
        '~~ setup twin from Private Profile file and re-write it reorganized
        PPfileAsPPtwin p_file
        PPtwinAsPPfile p_file
        GoTo xt
    End If
            
    bPPtwinHasChanged = PPtwinHasChanged(p_file)
    bPPfileHasChanged = PPfileHasChanged(p_file)
    Select Case True
        Case Not bPPtwinHasChanged And Not bPPfileHasChanged
            '~~ Nothing (no change) to manage
        Case bPPtwinHasChanged And Not bPPfileHasChanged
           '~~ re-write Private Profile file from the twin
            PPtwinAsPPfile p_file
        Case bPPfileHasChanged And Not bPPtwinHasChanged
            '~~ re-setup the twin to possibly re-organize items in ascending order
            '~~ and replace the Private Profile file with the reorganized twin as string
            PPfileAsPPtwin p_file
            PPtwinAsPPfile p_file
        Case Else
            Err.Raise AppErr(1), _
            ErrSrc(PROC), "The twin has changed and the file has changed too, which is a severe logic error."
    
    End Select

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                          
Private Sub PPtwinAsFile(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Writes a Private Profile's content based on a Private Profile's twin as
' string and saves the string for future change detection.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinAsFile"
    
    On Error GoTo eh
    Dim sPPtwin   As String
    
    BoP ErrSrc(PROC)
    sPPtwin = PPtwinAsString(p_file)
    StringAsFile sPPtwin, p_file
    
    If dctPPtwinsAsString.Exists(p_file) _
    Then dctPPtwinsAsString.Remove p_file
    dctPPtwinsAsString.Add p_file, sPPtwin
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPtwinAsPPfile(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Re-writes a Private Profile file (p_file) from its twin as string and saves
' the string for future change detections.
' ----------------------------------------------------------------------------
    Dim s As String
    
    s = PPtwinAsString(p_file)
    StringAsFile s, p_file
    
    If dctPPtwinsAsString.Exists(p_file) _
    Then dctPPtwinsAsString.Remove p_file
    dctPPtwinsAsString.Add p_file, s
    
End Sub

Private Function PPtwinAsString(ByVal p_file As String) As String
' ----------------------------------------------------------------------------
' Retuns a Private Profile's twin (p_file) transformed into a string.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinAsString"
    
    On Error GoTo eh
    Dim arrOut()      As String
    Dim vPPtwinSect   As Variant
    Dim vPPtwinValue  As Variant
    
    Set cllPPtwin = dctPPtwins(p_file)
    Set dctPPtwinSections = cllPPtwin(2)
    
    PPtwinAsStringComments arrOut, cllPPtwin(1)                                     ' File comment
    For Each vPPtwinSect In dctPPtwinSections
        PPtwinAsStringComments arrOut, dctPPtwinSections(vPPtwinSect)(1)            ' Section comment
        ArrayAdd arrOut, "[" & vPPtwinSect & "]"                                    ' Section
        Set dctPPtwinSection = dctPPtwinSections(vPPtwinSect)(2)
        For Each vPPtwinValue In dctPPtwinSection
            PPtwinAsStringComments arrOut, dctPPtwinSection(vPPtwinValue)(1)           ' Value comment
            ArrayAdd arrOut, vPPtwinValue & "=" & dctPPtwinSection(vPPtwinValue)(2)    ' Value
        Next vPPtwinValue
    Next vPPtwinSect
    
    PPtwinAsStringComments arrOut, cllPPtwin(3)                                      ' File footer
    PPtwinAsString = Join(arrOut, vbCrLf)
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub PPtwinAsStringComments(ByRef p_arr As Variant, _
                           ByVal p_cll As Collection)
    Dim v As Variant
    
    For Each v In p_cll
        '~~ Accepts aan empty line as a valid comment
        If v = vbNullString _
        Then ArrayAdd p_arr, v _
        Else ArrayAdd p_arr, Replace(";" & v, ";;", ";")
    Next v
    
End Sub

Private Function PPtwinExists(ByVal p_file As String) As Boolean
    PPtwinExists = dctPPtwins.Exists(p_file)
End Function

Private Sub PPtwinGet(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Provides a Private Profile twin's Header, Sections, and Footer
' ----------------------------------------------------------------------------
    Set cllPPtwin = dctPPtwins(p_file)
    Set cllPPtwinHeader = cllPPtwin(1)
    Set dctPPtwinSections = cllPPtwin(2)
    Set cllPPtwinFooter = cllPPtwin(3)
End Sub

Private Function PPtwinHasChanged(ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Private Profile's twin (f_file) has changed.
' ----------------------------------------------------------------------------
    PPtwinHasChanged = StrComp(PPtwinAsString(p_file) _
                             , dctPPtwinsAsString(p_file) _
                             , vbBinaryCompare) <> 0
    
End Function

Private Sub PPtwinLet(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Returns/replaces a Private Profile twin with its elements Header, Section,
' and Footer to the twins Directory
' ----------------------------------------------------------------------------
    Dim cllPPtwinNew As New Collection

    cllPPtwinNew.Add cllPPtwinHeader
    cllPPtwinNew.Add dctPPtwinSections
    cllPPtwinNew.Add cllPPtwinFooter
    dctPPtwins.Remove p_file
    dctPPtwins.Add p_file, cllPPtwinNew
    
End Sub

Private Sub PPtwinRemove(ByVal p_file As String)
    If dctPPtwins.Exists(p_file) Then dctPPtwins.Remove p_file
    If dctPPtwinsAsString.Exists(p_file) Then dctPPtwinsAsString.Remove p_file
End Sub

Private Sub PPtwinSectionGet(ByVal p_file As String, _
                             ByVal p_section As String)
' ----------------------------------------------------------------------------
' Provides a Private Profile twin's section elements.
' ----------------------------------------------------------------------------
    Set cllPPtwinSection = dctPPtwinSections(p_section)
    Set cllPPtwinSectionComments = cllPPtwinSection(1)
    Set dctPPtwinSection = cllPPtwinSection(2)

End Sub

Private Sub PPtwinSectionLet(ByVal p_file As String, _
                              ByVal p_section As String)
' ----------------------------------------------------------------------------
' Returns/replaces a Private Profile twin's section elements in the
' dtcPPtwinSections and arranges the sections in ascending order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    cll.Add cllPPtwinSectionComments
    cll.Add dctPPtwinSection
    
    dctPPtwinSections.Remove p_section
    dctPPtwinSections.Add p_section, cll
    
    KeySort dctPPtwinSections

End Sub

Private Sub PPtwinSectionRemove(ByVal p_file As String, _
                                ByVal p_section As String)
' ----------------------------------------------------------------------------
' Removes an existing section (p_section) from a Private Profile file (p_file)
' twin.
' See PPfileAsPPtwin for the implemented structure.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinSectionRemove"
    
    On Error GoTo eh '
    PPtwinGet p_file
    
    If dctPPtwinSections.Exists(p_section) Then
        dctPPtwinSections.Remove p_section
        '~~ Replace twin in twins Dictionary
        PPtwinLet p_file
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPtwinValueRemove(ByVal p_file As String, _
                              ByVal p_section As String, _
                              ByVal p_value_name As String)
' ----------------------------------------------------------------------------
' Removes an existing section (p_section) from a Private Profile file (p_file)
' twin.
' See PPfileAsPPtwin for the implemented structure.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinSectionRemove"
    
    On Error GoTo eh
    
    PPtwinGet p_file
    
    If dctPPtwinSections.Exists(p_section) Then
        PPtwinSectionGet p_file, p_section

        If dctPPtwinSection.Exists(p_value_name) Then
            dctPPtwinSection.Remove p_value_name ' removes the value together with a possible value comment
            '~~ Re-setup section
            PPtwinSectionLet p_file, p_section
            '~~ Replace the twin with the twin with the processed section
            PPtwinLet p_file
        End If
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-Private-Profile"
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

'Public Sub Reorg(Optional name_file As String = vbNullString)
'' ----------------------------------------------------------------------------
'' Reorganizes a Private Profile file (name_file) resulting in sections and
'' values in ascending order by considering section and value headers.
'' When no file is provided it defaults to the file name specified with the
'' property FileName.
'' ----------------------------------------------------------------------------
'    Const PROC  As String = "Reorg"
'
'    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
'    PPmanage name_file
'
'End Sub

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
                      Optional ByRef s_sect_comments As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Called by Exists.
' ----------------------------------------------------------------------------
    Dim i   As Long
    
    If Not ArrayIsAllocated(s_arr) Then GoTo xt
    
    For i = LBound(s_arr) To UBound(s_arr)
        If IsSection(s_arr(i), s_sect_name) Then
            '~~ The line is the section requested
            SectionExists = True
            s_sect_indx = i
            Set s_sect_comments = CommentsOfItem(s_arr, i) ' Return the section's comment comments (may be none)
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
    
    If Not IsValidFile(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
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
' Removes one or more sections (name_section) in a Private Profile file
' (name_file) whereby several sections may be provided as a comma sepearated
' string.
' ----------------------------------------------------------------------------
    Const PROC = "SectionRemove"
    
    On Error GoTo eh
    Dim v As Variant
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    
    For Each v In StringAsArray(name_section)
        PPtwinSectionRemove name_file, v
    Next v
    
    PPmanage name_file
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub SectionsCopy(ByVal name_file_source As String, _
'                        ByVal name_file_target As String, _
'               Optional ByVal name_sections As Variant = Nothing, _
'               Optional ByVal s_merge As Boolean = False)
'' ----------------------------------------------------------------------------
'' Copies sections provided as names (name_sections) from a source Private
'' Profile file (name_file_source) to a target Private Profile file
'' (name_file_target).
'' ----------------------------------------------------------------------------
'    Const PROC = "SectionsCopy"
'
'    On Error GoTo eh
'    Dim dct   As Dictionary
'    Dim vName As Variant
'    Dim vSect As Variant
'    Dim sSect As String
'    Dim sName As String
'    Dim sValu As String
'
'    If Not IsValidFile(name_file_source) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file_source)
'    If Not IsValidFile(name_file_target) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file_target)
'
'    For Each vSect In VarAsCollection(name_sections)
'        sSect = vSect
'        If Not IsValidSectName(sSect) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(sSect)
'
'        If Not s_merge Then SectionRemove sSect, name_file_target
'        Set dct = ValueNames(name_file_source, vSect)
'        For Each vName In dct
'            sName = vName
'            sValu = dct(vName)
'            Value(sName, sSect, name_file_target) = sValu
'        Next vName
'     Next vSect
'
'xt: Set dct = Nothing
'    PPtwinAsFile name_file_target ' re-write file in order to maintain sections in ascending sequence
'    Exit Sub
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

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

Private Function SplitIndctr(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns the split indicator in a string (s) with a given priority, i.e. the
' first one detected matters. When none is found the split indicator defaults
' to vbCrLf.
' ----------------------------------------------------------------------------
    SplitIndctr = vbCrLf ' The dafault
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: SplitIndctr = vbCrLf
        Case InStr(s, "|&|") <> 0:  SplitIndctr = "|&|"
        Case InStr(s, "||") <> 0:  SplitIndctr = "||"
        Case InStr(s, ", ") <> 0:  SplitIndctr = ", "
        Case InStr(s, "; ") <> 0:  SplitIndctr = ", "
        Case InStr(s, ",") <> 0:  SplitIndctr = ","         ' when ", " has been negative
        Case InStr(s, ";") <> 0:  SplitIndctr = ";"         ' when "; " has been negative
    End Select

End Function

Private Function StringAsArray(ByVal s_items As String, _
                      Optional ByVal s_split As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    StringAsArray = Split(s_items, s_split)

End Function

Private Function StringAsCollection(ByVal s_items As String, _
                           Optional ByVal s_split As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    With cll
        For Each v In Split(s_items, s_split)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function StringAsDictionary(ByVal s_items As String, _
                           Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    
    With dct
        For Each v In Split(s_items, s_split)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                l = dct(v) + 1
                .Remove v
                .Add v, l
            End If
        Next v
    End With
    Set StringAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function StringAsFile(ByVal s_strng As String, _
                     Optional ByRef s_file As Variant = vbNullString, _
                     Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Private Function StringAsString(ByVal s_items As String, _
                       Optional ByVal s_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter string replaced with vbCrLf.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)

    StringAsString = Replace(s_items, s_split, vbCrLf)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    FSo.CreateTextFile sTemp

End Function

Private Function ValueExists(ByVal v_source As Variant, _
                             ByVal v_sect_name As String, _
                             ByVal v_valu_name As String, _
                    Optional ByRef v_valu As String, _
                    Optional ByRef v_index As Long, _
                    Optional ByRef v_valu_comments As Collection, _
                    Optional ByRef v_sect_comments As Collection) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a value (v_sect_name, v_valu_name) exists in a source
' (v_source) which might be an array or a string specifiying a files's full
' name. When the value exists its section comments (v_sect_coments) and
' ----------------------------------------------------------------------------
    Const PROC = "ValueExists"
    
    On Error GoTo eh
    Dim i   As Long
    
    If TypeName(v_source) = "String" And Not ArrayIsAllocated(v_source) Then
        '~~ When the source is a file's full name the file is transformed into an array
        v_source = FileAsArray(v_source)
    End If
    If Not ArrayIsAllocated(v_source) Then GoTo xt ' in case the file is empty
    If Not SectionExists(v_source, v_sect_name, i, v_sect_comments) Then GoTo xt
    '~~ When the section exists, continue with the next line
    i = i + 1
    If i <= UBound(v_source) Then
        Do While Not IsSection(v_source(i), v_sect_name)
            '~~ Loop down to the next section or the end of the file when it is the last section
            If IsValue(v_source(i), v_valu_name, v_valu) Then
                ValueExists = True
                v_index = i
                Set v_valu_comments = CommentsOfItem(v_source, i) ' Return the value's comments (may be none)
                Exit Do
            End If
            If i = UBound(v_source) Then Exit Do
            i = i + 1
        Loop
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub ValueLet(ByVal name_file As String, _
                     ByVal name_section As String, _
                     ByVal name_value As String, _
                     ByVal v_value As Variant)
                     
    Dim cllFooter   As New Collection
    Dim cllComments As New Collection
    Dim lChars      As Long

    If VarType(v_value) = vbBoolean Then v_value = VBA.CStr(VBA.CLng(v_value))
    
    Set cllComments = CommentsFileSave(name_file)
    Set cllFooter = PPfooterSave(name_file)
        
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
    If cllComments.Count > 0 Then CommentsFileRestore name_file, cllComments
    If cllFooter.Count > 0 Then PPfooterRestore name_file, cllFooter

End Sub

Private Function ValueNameExists(ByVal p_file_name As String, _
                                 ByVal p_sect_name As String, _
                                 ByVal p_valu_name As String, _
                        Optional ByVal p_valu_comments As Collection, _
                        Optional ByRef p_valu As String, _
                        Optional ByRef p_sect_coll As Collection, _
                        Optional ByRef p_file_coll As Collection) As Boolean
' ----------------------------------------------------------------------------
' When a named value (p_value_name) exists in a provided section (p_sect_name)
' in a provided Private Profile file (p_file_name) the function returns TRUE and
' the value as Dictionary (p_value_dict).
' ----------------------------------------------------------------------------
            
    Dim dctFile As Dictionary
    Dim dctSect As Dictionary
    
    PPmanage p_file_name ' in case the file has changed or the file does not exist as Directory structure
    
    If dctPPtwins.Exists(p_file_name) Then
        Set p_file_coll = dctPPtwins(p_file_name)
        Set dctFile = p_file_coll(2)
        If dctFile.Exists(p_sect_name) Then
            Set p_sect_coll = dctFile(p_sect_name)
            Set dctSect = p_sect_coll(2)
            If dctSect.Exists(p_valu_name) Then
                ValueNameExists = True
                Set p_valu_comments = dctSect(p_valu_name)(1)
                p_valu = dctSect(p_valu_name)(2)
            End If
            Set dctSect = Nothing
        End If
        Set dctFile = Nothing
    End If

End Function

Public Function ValueNameRename(ByVal name_value_old As String, _
                                ByVal name_value_new As String, _
                       Optional ByVal name_sections As Variant = vbNullString, _
                       Optional ByVal name_file As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Renames in a Private Profile file (name_file) a given value-name
' (name_value_old) into a new value-name (name_value_new) in all provided
' sections (name_sections), when none are providedd, in all sections.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRename"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim arrOut      As Variant
    Dim dctSections As New Dictionary
    Dim i           As Long
    Dim sValueName  As String
    
    If name_value_old = vbNullString Or name_value_new = vbNullString Then GoTo xt
    If Not IsValidFile(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
    If name_sections <> vbNullString _
    Then Set dctSections = VarAsDictionary(name_sections) _
    Else Set dctSections = SectionNames(name_file)
    
    For i = LBound(arr) To UBound(arr)
        sSection = vbNullString
        If IsSection(arr(i), sSection) Then
            ArrayAdd arrOut, arr(i)
            If dctSections.Exists(sSection) Then
                '~~ Rename value in section if value-name is in it
                i = i + 1 ' proceed to first value
                sValueName = vbNullString
                Do While IsValue(arr(i), sValueName)
                    '~~ Loop down to next section or end of file when it is the last section
                    If sValueName = name_value_old Then
                        '~~ re-write value with new name
                        ArrayAdd arrOut, name_value_new & "=" & Split(arr(i), "=")(1)
                        ValueNameRename = True
                    Else
                        '~~ write unchanged
                        ArrayAdd arrOut, arr(i)
                    End If
                    If i = UBound(arr) Then Exit Do
                    i = i + 1
                    sValueName = vbNullString
                Loop
                i = i - 1
            End If
            sSection = vbNullString
        Else
            ArrayAdd arrOut, arr(i)
        End If
    Next i
        
xt: If ValueNameRename Then
        '~~ Write the Private Profile file with the renamed value-name(s)
        FileFromString name_file, Join(arrOut, vbCrLf)
        '~~ Re-write the Private Profile file reorganized to ensure value-names in ascending order
        PPmanage name_file
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
' Note: The returned value-names are distinct names! I.e. when a value exists
'       in more than one section it is still one distinct value-name.
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
    
    If Not IsValidFile(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    
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
        GoTo xt
    End If
        
    For Each v In StringAsCollection(name_sections)
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
' Removes one or more value-names (name_value) from a section (name_section)
' in a Private Profile file (name_file), whereby more names may be provided as
' a comma separated string.
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' Note: The value is removed from the twin and the twin is re-written as
'       Private Profile file.
' ----------------------------------------------------------------------------
    Const PROC = "ValueRemove"
    
    Dim v As Variant
      
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidSectName(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidValuName(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
                       
    PPmanage name_file ' ensure an up-to-date twin
    For Each v In StringAsArray(name_value)
        PPtwinValueRemove name_file, name_section, v
    Next v
    PPmanage name_file

xt:
End Sub

Private Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case "File":        VarAsArray = FileAsArray(v_items)
        Case "String":      If IsArray(v_items) _
                            Then VarAsArray = v_items _
                            Else StringAsArray (v_items)
    End Select
    
End Function

Private Function VarAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsCollection = v_items
        Case "File":        Set VarAsCollection = FileAsCollection(v_items)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsCollection = ArrayAsCollection(v_items) _
                            Else Set VarAsCollection = StringAsCollection(v_items)
    End Select
    
End Function

Private Function VarAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsDictionary = CollectionAsDictionary(v_items)
        Case "File":        If Not v_items Is v_items _
                            Then Set VarAsDictionary = FileAsDictionary(v_items)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsDictionary = ArrayAsDictionary(v_items) _
                            Else Set VarAsDictionary = StringAsDictionary(v_items)
    End Select
    
End Function

Private Function VarAsFile(ByVal v_items As Variant, _
                  Optional ByVal v_file As String) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsFile = CollectionAsFile(v_items, v_file)
        Case "File":        If Not v_items Is v_items _
                            Then Set VarAsFile = FileAsFile(v_items, v_file)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsFile = ArrayAsFile(v_items, v_file) _
                            Else Set VarAsFile = StringAsFile(v_items, v_file)
    End Select
    
End Function

Private Function VarAsString(ByVal v_items As Variant) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  VarAsString = CollectionAsString(v_items)
        Case "File":        VarAsString = FileAsString(v_items)
        Case "String":      If IsArray(v_items) _
                            Then VarAsString = ArrayAsString(v_items) _
                            Else VarAsString = StringAsString(v_items)
    End Select
    
End Function

