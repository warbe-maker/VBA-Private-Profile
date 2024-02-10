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
' FileName    r/w - Specifies a Private Profile file full name
'                 - Returns a valid Private Profile file's full name
' Section     r/w - Returns the current specified Section name
'                 - Specifies the Section name valid for all subsequent
'                   methods and properties until another section name is
'                   specified.
' Value       r/w - Reads from a Private Profile File a value with a
'                   provided value-name from a provided section
'                 - Writes to a Private Profile File a value with a provided
'                   value-name into a provided section
' Public methods:
' ---------------
' NamesRemove     Removes provided value names, in a given Private
'                 Properties File, when provided in a specific section,
'                 else in all sections.
' SectionExists   Returns TRUE when a given section exists in a given
'                 Private Profile File.
' SectionNames    Returns a Dictionary of all section names [...] in a
'                 Private Profile File whereby each item is a Dictionary of
'                 value-names as key and the value as item, and the key of
'                 the item is the section name.
' SectionsCopy    Copies sections, provided as comma delimites string of
'                 section names from a soure PrivProfile file into a target
'                 Private Profile file, optionally merged.
' SectionsRemove  Removes the sections provided as a comma delimited string
'                 as argument. Sections not existing are ignored.
' ValueNameExists Returns TRUE when a value-name exists in a provided
'                 Private Profile File's section.
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
Public FSo                      As FileSystemObject

Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-File-System-Objects"
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

Private dctFiles        As Dictionary ' file name as key, item is dctSections
Private dctFilesStrng   As Dictionary ' file as string
Private dctSections     As Dictionary ' section name as key, item is dctValues
Private dctValues       As Dictionary ' value name as key, item is value
Private Reg             As RegExp
Private sFileName       As String
Private sSection        As String

Private Sub Class_Initialize()

    Set Reg = New RegExp
    Set FSo = New FileSystemObject
    Set dctFiles = New Dictionary
    Set dctFilesStrng = New Dictionary
    '~~ Initialize the default Private Profile File name
    sFileName = ThisWorkbook.Path & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
       
End Sub

Private Sub Class_Terminate()
    Set FSo = Nothing
    Set dctFiles = Nothing
End Sub

Public Property Get FileName() As String
    If sFileName = vbNullString _
    Then sFileName = ThisWorkbook.Path & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
    FileName = sFileName
End Property

Public Property Let FileName(ByVal s As String)
    Const PROC = "FileName-Let"
    
    If Not FileNameIsValid(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided file name is not a valid file full name!"
    sFileName = s

End Property

Public Property Get Section() As String:        Section = sSection: End Property

Public Property Let Section(ByVal s As String)
    Const PROC = "Section-Let"
    
    If Not SectionNameIsValid(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Not a vaild section name!"
    sSection = s
    
End Property

Private Property Get Sections(Optional ByVal s_file As String = vbNullString) As Dictionary
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
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim dct         As New Dictionary
    
    If Not FileNameIsValid(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
        
    '~~ Return all sections
    With dct
        For Each vSection In SectionNames(s_file)
            .Add vSection, Values(s_file, vSection)
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

Private Property Let Sections(Optional ByVal s_file As String = vbNullString, _
                                       ByVal s_dct As Dictionary)
' ------------------------------------------------------------------------
' Writes the sections provided as Dictionary (pp_dct) to file (pp_file) -
' provided as full name string or as file object. Existing sections/values
' are overwritten new sections/values are added.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ------------------------------------------------------------------------
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vName       As Variant
    Dim vSection    As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    Dim fl          As String
    Dim vValue      As Variant
    
    If Not FileNameIsValid(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    For Each vSection In s_dct
        sSection = vSection
        Set dctValues = s_dct(sSection)
        For Each vName In dctValues
            sName = vName
            vValue = dctValues.item(vName)
            Value(v_file:=s_file _
                , v_section:=sSection _
                , v_value_name:=sName _
                 ) = vValue
        Next vName
    Next vSection
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Value(Optional ByVal v_value_name As String = vbNullString, _
                          Optional ByVal v_section As String = vbNullString, _
                          Optional ByVal v_file As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Reads a value with a specific name from a section from a file organized:
'[section]
'<value-name>=<value>
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Value-Get"
    
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant
    Dim fl      As String
    
    If Not FileNameIsValid(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not SectionNameIsValid(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not ValueNameIsValid(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
    
    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=v_section _
                                    , lpg_KeyName:=v_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=v_file _
                                     )
    Value = Left$(sRetVal, lResult)
    
End Property

Public Property Let Value(Optional ByVal v_value_name As String = vbNullString, _
                          Optional ByVal v_section As String = vbNullString, _
                          Optional ByVal v_file As String = vbNullString, _
                                   ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value under a given name (v_value_name) into a section (v_section)
' in a file (v_file) organized:
' [section]
' <value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
        
    Dim bIsNew  As Boolean
    Dim lChars  As Long
    Dim sValue  As String
    Dim fl      As String
    
    If Not FileNameIsValid(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not SectionNameIsValid(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not ValueNameIsValid(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
    
    Select Case VarType(v_value)
        Case vbBoolean: sValue = VBA.CStr(VBA.CLng(v_value))
        Case Else:      sValue = v_value
    End Select
    
    bIsNew = Not Exists(v_file, v_section, v_value_name)
    
    lChars = WritePrivateProfileString(lpw_ApplicationName:=v_section _
                                     , lpw_KeyName:=v_value_name _
                                     , lpw_String:=v_value _
                                     , lpw_FileName:=v_file)
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & v_section & "'" & vbLf & _
               "Value name = '" & v_value_name & "'" & vbLf & _
               "Value      = '" & CStr(v_value) & "'" & vbLf & _
               "Value file = '" & v_file & "'"
    End If
    
    If bIsNew Then FileFromDict v_file, Sections(v_file) ' re-write the file in order to maintain asscending sequence
    
End Property

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

Public Function Exists(ByVal e_file As String, _
              Optional ByRef e_section As String = vbNullString, _
              Optional ByVal e_name As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case True
        Case e_section = vbNullString And e_name = vbNullString
            Exists = FSo.FileExists(e_file)
        Case e_section <> vbNullString And e_name = vbNullString
            If FSo.FileExists(e_file) Then
                Exists = SectionExists(e_file, e_section)
            End If
        Case e_section <> vbNullString And e_name <> vbNullString
            If FSo.FileExists(e_file) Then
                If SectionExists(e_file, e_section) Then
                    Exists = ValueNameExists(e_file, e_section, e_name)
                End If
            End If
    End Select
    
End Function

Private Function FileAsDict(ByVal f_file As String) As Dictionary
' ----------------------------------------------------------------------------
' Adds the file (f_file) as Dictionary structure to the dctFiles Dictionary
' with the file-name as the key. An already existing entry is removed first.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsDict"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim dctValues   As Dictionary
    Dim dct         As Dictionary
    Dim i           As Long
    Dim s           As String
    Dim sFileStrng  As String
    Dim sName       As String
    Dim sValue      As String
    
    BoP ErrSrc(PROC), f_file
    sFileStrng = FileAsString(f_file)
    
    arr = Split(sFileStrng, vbCrLf)
    Set dct = New Dictionary
    
    For i = LBound(arr) To UBound(arr)
        If ItemIsSection(arr(i), s) Then
            If Not dctValues Is Nothing Then
                dct.Add sSection, KeySort(dctValues)
                Set dctValues = Nothing
            End If
            sSection = s
            Set dctValues = New Dictionary
        ElseIf ItemIsValue(arr(i), sName, sValue) Then
            If dctValues.Exists(sName) Then
                If dctValues(sName) <> sValue Then
                    dctValues.Remove sName
                    dctValues.Add sName, sValue
                End If
            Else
                dctValues.Add sName, sValue
            End If
        End If
    Next i
    If Not dctValues Is Nothing Then
        If dct.Exists(sSection) Then dct.Remove sSection
        dct.Add sSection, KeySort(dctValues)
        Set dctValues = Nothing
    End If
    
    Set FileAsDict = dct
    Set dct = Nothing

xt: EoP ErrSrc(PROC), f_file
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub FileAsDictRemove(ByVal f_file As String)
    If dctFiles.Exists(f_file) Then dctFiles.Remove f_file
    If dctFilesStrng.Exists(f_file) Then dctFilesStrng.Remove f_file
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
    If Right(FileAsString, Len(f_split)) = f_split Then
        FileAsString = Left(FileAsString, Len(FileAsString) - Len(f_split))
    End If
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

Private Sub FileFromDict(ByVal f_file As String, _
                         ByVal f_from As Dictionary)
' ----------------------------------------------------------------------------
' Replaces the content of a Private Profile file (f_file) with a reorganized
' Directory structure, i.e. with all sections and value names in ascending
' order.
' ----------------------------------------------------------------------------
    Const PROC = "FileFromDict"
    
    On Error GoTo eh
    Dim arr()       As Variant
    Dim dctValues   As Dictionary
    Dim i           As Long
    Dim vSect       As Variant
    Dim vValue      As Variant
    Dim s           As String
    
    i = -1
    For Each vSect In f_from
        i = i + 1
        ReDim Preserve arr(i)
        arr(i) = "[" & vSect & "]"
        Set dctValues = f_from(vSect)
        For Each vValue In dctValues
            i = i + 1
            ReDim Preserve arr(i)
            arr(i) = vValue & "=" & dctValues(vValue)
        Next vValue
    Next vSect
    s = Join(arr, vbCrLf)
    FileFromString f_file, s
    If dctFilesStrng.Exists(f_file) Then dctFilesStrng.Remove f_file
    dctFilesStrng.Add f_file, s
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

    If dctFilesStrng.Exists(f_file) Then
        FileHasChanged = StrComp(dctFilesStrng(f_file), FileAsString(f_file), vbBinaryCompare) <> 0
        If FileHasChanged Then Debug.Print "Differs"
    Else
        FileHasChanged = True
    End If
    
End Function

Public Function FileNameIsValid(ByRef s As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a provided string (s) is a vaild full file name. I.e. it
' contains a : and a \ as the minimum.
' ----------------------------------------------------------------------------
    Const PROC = "FileNameIsValid"
    
    On Error GoTo eh
    If s = vbNullString And sFileName <> vbNullString Then
        s = sFileName
        FileNameIsValid = True
        GoTo xt
     End If
     
    '~~ An explicitely specified file name is checked for validity and when valid
    '~~ becomes the FileName for all subsequent properties and methods which did not
    '~~ hae a file name specified as argument.
    With FSo
        If .FileExists(s) Then
            If IsTextFile(s) Then
                FileNameIsValid = True
            End If
            GoTo xt
        End If
        
        If InStr(s, ":") = 0 Then Exit Function
        If InStr(s, "\") = 0 Then Exit Function
        If .GetExtensionName(s) = vbNullString Then Exit Function
        On Error Resume Next
        .CreateTextFile s
        If Err.Number = 0 Then
            FileNameIsValid = True
        End If
    End With
        
xt: If FileNameIsValid Then
        '~~ FileHasChanged takes into consideration that the file may
        '~~ have changed by an action "outside" this module, e.g. manually.
        If Not dctFiles.Exists(s) Then
            dctFiles.Add s, FileAsDict(s)
            If dctFilesStrng.Exists(s) Then dctFilesStrng.Remove s
            dctFilesStrng.Add s, FileAsString(s)
        ElseIf FileHasChanged(s) Then
            FileAsDictRemove s
            dctFiles.Add s, FileAsDict(sFileName)
            dctFilesStrng.Add s, FileAsString(s)
        End If
        sFileName = s
    End If
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Function ItemIsSection(ByVal i_item As String, _
                               ByRef i_section_name As String) As Boolean

    If Left(i_item, 1) = "[" And Right(i_item, 1) = "]" Then
        ItemIsSection = True
        i_section_name = Mid(i_item, 2, Len(i_item) - 2)
    End If
    
End Function

Private Function ItemIsValue(ByVal i_item As String, _
                             ByRef i_value_name As String, _
                             ByRef i_value As String) As Boolean

    If Left(i_item, 1) <> "[" Then
        ItemIsValue = True
        i_value_name = Split(i_item, "=")(0)
        i_value = Replace(i_item, i_value_name & "=", vbNullString, 1, 1)
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

Public Sub NamesRemove(ByVal n_value_names As Variant, _
              Optional ByVal n_file As String = vbNullString, _
              Optional ByVal n_section As String = vbNullString, _
              Optional ByRef n_removed As Long)
' -----------------------------------------------------------------------------
' PrivateProfile file service. Removes from a PrivateProfile file's (n_file)
' section (n_section) a name (n_value_names) provided either as a
' comma delimited string, or as one string only.
' When no file (n_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' When the name existed and has been removed, the function returns TRUE.
' -----------------------------------------------------------------------------
    Const PROC = "NamesRemove"
    
    On Error GoTo eh
    Dim v   As Variant
    
    If Not FileNameIsValid(n_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not SectionNameIsValid(n_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    
    For Each v In Split(Replace(n_value_names, " ", vbNullString), ",")
        If ValueNameExists(n_file, n_section, v) Then
            DeletePrivateProfileKey Section:=n_section _
                                  , Key:=v _
                                  , Setting:=0 _
                                  , Name:=n_file
            n_removed = n_removed + 1
        End If
    Next v
    If n_removed > 0 Then Set dctSections = Null
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function NameStrings(Optional ByVal a_v As Variant = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the name strings (section names or value names) provided as argument
' (a_v) as Collection whereby the provided argument may be:
' - a Collection of strings
' - a Dictionary of keys
' - a comma delimited string, a
' Dictionary, or a Collection of string items.
' ----------------------------------------------------------------------------
    Const PROC = "NameStrings"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim v   As Variant
    
    Select Case TypeName(a_v)
        Case "Collection"
            Set cll = a_v
        Case "Dictionary"
            For Each v In a_v
                cll.Add v
            Next v
        Case "String"
            For Each v In Split(a_v, ",")
                cll.Add VBA.Trim$(v)
            Next v
        Case "Array"
            For Each v In a_v
                cll.Add VBA.Trim$(v)
            Next v
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a String, an Array, a Collecton, nor a Dictionary!"
    End Select
            
xt: Set NameStrings = cll
    Set cll = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Private Function SectionExists(ByVal s_file As String, _
                               ByVal s_section As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a section (s_section) exists a Private Profile file
' (s_file).
' ----------------------------------------------------------------------------
    
    If Not dctFiles.Exists(s_file) _
    Then dctFiles.Add s_file, FileAsDict(s_file)
    SectionExists = dctFiles(s_file).Exists(s_section)
    
End Function

Private Function SectionNameIsValid(ByRef i_section As String) As Boolean
    
    If i_section = vbNullString Then
        i_section = sSection
        SectionNameIsValid = True
    Else
        With Reg
            .Pattern = "^[a-zA-Z0-9_]+$"
            SectionNameIsValid = .Test(i_section)
            If SectionNameIsValid _
            Then sSection = i_section
        End With
    End If
    
End Function

Public Function SectionNames(Optional ByVal s_file As String) As Dictionary
' ----------------------------------------------------------------------------
' Private Profile File service: Returns a Dictionary with all section names in
' the file (s_file) whereby the file (s_file) may be provided as full name
' string or as a file object. The sections are returned in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim asSections()    As String
    Dim dct             As New Dictionary
    Dim fl              As String
    Dim iLen            As Long
    Dim strBuffer       As String
    Dim v               As Variant
    Dim s               As String
    
    Set SectionNames = dct
    
    If Not FileNameIsValid(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
        
    With dct
        For Each v In Split(FileAsString(s_file), vbCrLf)
            If ItemIsSection(v, s) Then
                .Add s, s
            End If
        Next v
    End With
    
xt: Set SectionNames = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SectionsCopy(ByVal s_source As String, _
                        ByVal s_target As String, _
               Optional ByVal s_sections As Variant = Nothing, _
               Optional ByVal s_merge As Boolean = False)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Copies sections from file (s_source) to file (s_target), when no section
' names (s_sections) are provided all. By default (s_merge) all sections
' are replaced.
' ----------------------------------------------------------------------------
    Const PROC = "PPsectionCopy"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim vName       As Variant
    Dim vSection    As Variant
    
    For Each vSection In NameStrings(s_sections)
        If Not s_merge Then
            '~~ Delete section in target file
            DeletePrivateProfileSection Section:=vSection _
                                      , NoKey:=0 _
                                      , NoSetting:=0 _
                                      , Name:=s_target
        End If
        Set dct = Values(s_source, vSection)
        For Each vName In dct
            WritePrivateProfileString lpw_ApplicationName:=vSection _
                                     , lpw_KeyName:=vName _
                                     , lpw_String:=dct(vName) _
                                     , lpw_FileName:=s_target
        Next vName
     Next vSection

xt: Set dct = Nothing
    FileAsDictRemove s_target
    dctFiles.Add s_target, FileAsDict(s_target)
    dctFilesStrng.Add s_target, FileAsString(s_target)
    '~~ Re-write all sections in order to maintain asscending sequence
    FileFromDict s_target, Sections(s_target)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SectionsRemove(Optional ByVal s_sections As String = vbNullString, _
                          Optional ByVal s_file As String = vbNullString)
' ----------------------------------------------------------------------------
' PrivateProfile file service: Removes a sections (s_sections) - provided as
' a comma delimited string, Dictionary section names as key or a Collection of
' section names, in file (s_file).
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "SectionsRemove"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim sSection    As String
    
    If Not FileNameIsValid(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    For Each v In NameStrings(s_sections)
        sSection = v
        If SectionNameIsValid(sSection) _
        Then DeletePrivateProfileSection Section:=v _
                                       , NoKey:=0 _
                                       , NoSetting:=0 _
                                       , Name:=s_file
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

Public Function Value2(ByVal v_value_name As String, _
              Optional ByVal v_section As String = vbNullString, _
              Optional ByVal v_file As String = vbNullString) As Variant
    
    Const PROC = "Value2"
    Dim dctValues   As Dictionary
    Dim dct         As Dictionary
    
    If Not FileNameIsValid(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not SectionNameIsValid(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not ValueNameIsValid(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
        
    If dctFiles.Exists(v_file) Then
        Set dct = dctFiles(v_file)
        If dct.Exists(v_section) Then
            Set dct = dct(v_section)
            If dct.Exists(v_value_name) Then
                Value2 = dct(v_value_name)
            End If
        End If
    End If
    
End Function

Private Function ValueNameExists(ByVal v_file As String, _
                                 ByVal v_section As String, _
                                 ByVal v_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a named value (v_value_name) exists in a provided section
' (v_section) in a provided Private Profile file.
' ----------------------------------------------------------------------------
    
    ValueNameExists = Values(v_file, v_section).Exists(v_name)

End Function

Private Function ValueNameIsValid(ByVal i_name As String) As Boolean
    With Reg
        .Pattern = "^[a-zA-Z0-9_]{1,32}$"
        ValueNameIsValid = .Test(i_name)
    End With
End Function

Public Function ValueNameRename(ByVal v_name_old As String, _
                                ByVal v_name_new As String, _
                       Optional ByVal v_sections As Variant = vbNullString, _
                       Optional ByVal v_file As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service: Rename a value name (v_name_old) to a new
' name (v_name_new) in all sections (v_sections), provided either as:
' - a Dictionary where the keys are the sections meant
' - a Collection of strings
' - a comma delimited string of section names
' When no sections are provided, the value name is renamed in all sections.
' When no file (v_file) is provided, the file defaults to the file name
' specified either explicitely with the property FileName - which, when a
' vbNullString defaults to: ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat,
' or the file specified with any previous property or method call.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRename"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim vValue      As Variant
    Dim dctSects    As Dictionary
    Dim vSect       As Variant
    Dim sRetVal     As String
    Dim lResult     As Long
    
    If v_name_old = vbNullString Or v_name_new = vbNullString Then GoTo xt
    If Not FileNameIsValid(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
        
    If v_sections = vbNullString _
    Then Set v_sections = NameStrings(SectionNames(v_file)) _
    Else Set v_sections = NameStrings(v_sections)
    
    '~~ Loop through all sections provided
    For Each vSect In v_sections
        If ValueNameExists(v_file, vSect, v_name_old) Then
            
            '~~> Get value of the to-be-renamed value-name
            sRetVal = String(32767, 0)
            lResult = GetPrivateProfileString(lpg_ApplicationName:=vSect _
                                            , lpg_KeyName:=v_name_old _
                                            , lpg_Default:="" _
                                            , lpg_ReturnedString:=sRetVal _
                                            , nSize:=Len(sRetVal) _
                                            , lpg_FileName:=v_file _
                                             )
            vValue = Left$(sRetVal, lResult)
            
            '~~ Delete old value-name
            DeletePrivateProfileKey Section:=vSect _
                                  , Key:=v_name_old _
                                  , Setting:=0 _
                                  , Name:=v_file
            
            '~~ When there was a value, re-write it under the new value-name
            If vValue <> vbNullString _
            Then WritePrivateProfileString lpw_ApplicationName:=vSect _
                                         , lpw_KeyName:=v_name_new _
                                         , lpw_String:=vValue _
                                         , lpw_FileName:=v_file
            ValueNameRename = True
        End If
    Next vSect
    
xt: Set dct = Nothing
    If ValueNameRename Then
        '~~ At least one value name has been re-named
        FileAsDictRemove v_file
        dctFiles.Add v_file, FileAsDict(v_file)
        dctFilesStrng.Add v_file, FileAsString(v_file)
        '~~ Re-write all sections in order to maintain asscending sequence
        FileFromDict v_file, Sections(v_file)
    End If
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueNames(Optional ByVal v_file As String = vbNullString, _
                           Optional ByVal v_section As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile service: Returns a Dictionary of all value names in the
' provided section (v_section), with the value name as key and the value as
' item, in ascending order by value name, in file (v_file) - provided either
' as full file name or as file object.
' When no file is provided it defaults to the file name specified with the
' property FileName - which, when a vbNullString defaults to:
' ThisWorkbook.Path & "\" ThisWorkbook's BaseName & ".dat.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dct         As New Dictionary
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim v           As Variant
    Dim sSection    As String
    Dim sName       As String
    Dim vNames      As Variant
    Dim fl          As String
    
    If Not FileNameIsValid(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not SectionNameIsValid(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    
    Set vNames = NameStrings(v_section)
    If vNames.Count = 0 Then Set vNames = SectionNames(v_file)
    
    For Each v In vNames
        sSection = v
        '~~> Retrieve the names for the provided section
        strBuffer = Space$(32767)
        lResult = GetPrivateProfileString(lpg_ApplicationName:=sSection _
                                        , lpg_KeyName:=vbNullString _
                                        , lpg_Default:=vbNullString _
                                        , lpg_ReturnedString:=strBuffer _
                                        , nSize:=Len(strBuffer) _
                                        , lpg_FileName:=v_file _
                                         )
        sNames = Left$(strBuffer, lResult)
    
        If sNames <> vbNullString Then                                         ' If there were any names
            asNames = Split(sNames, vbNullChar)                      ' have them split into an array
            For i = LBound(asNames) To UBound(asNames)
                sName = asNames(i)
                If Len(sName) <> 0 Then
                    If Not dct.Exists(sName) _
                    Then dct.Add sName, Value(v_file:=fl _
                                            , v_section:=sSection _
                                            , v_value_name:=sName)
                End If
            Next i
        End If
    Next v
        
    Set ValueNames = KeySort(dct)

xt: Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Values(ByVal v_file As String, _
                       ByVal v_section As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns all values in a section (v_section) in a Private Profile file
' (v_file) as Dictionary, with the value name as key and the value as
' item, in asscending order. When the provided section does not exist the
' returned Dictionary is empty.
' Note: The method is public mainly for test purpose.
' ----------------------------------------------------------------------------
    Const PROC = "Values"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dct         As New Dictionary
    Dim fl          As String
    Dim i           As Long
    Dim lResult     As Long
    Dim sName       As String
    Dim sNames      As String
    Dim sRetVal     As String
    Dim strBuffer   As String
    
    Set Values = New Dictionary ' may remain empty though
    If Not SectionExists(v_file, v_section) Then GoTo xt
    
    '~~> Retrieve the names in the provided section
    strBuffer = Space$(32767)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=v_section _
                                    , lpg_KeyName:=vbNullString _
                                    , lpg_Default:=vbNullString _
                                    , lpg_ReturnedString:=strBuffer _
                                    , nSize:=Len(strBuffer) _
                                    , lpg_FileName:=v_file _
                                     )
    sNames = Left$(strBuffer, lResult)
    If sNames = vbNullString Then GoTo xt

    '~~ At least one value name has been returned
    asNames = Split(sNames, vbNullChar)
    With dct
        For i = LBound(asNames) To UBound(asNames)
            sName = asNames(i)
            If Len(sName) <> 0 Then
                If Not .Exists(sName) Then
                    '~~ Get value
                    sRetVal = String(32767, 0)
                    lResult = GetPrivateProfileString(lpg_ApplicationName:=v_section _
                                                    , lpg_KeyName:=sName _
                                                    , lpg_Default:="" _
                                                    , lpg_ReturnedString:=sRetVal _
                                                    , nSize:=Len(sRetVal) _
                                                    , lpg_FileName:=v_file _
                                                     )
                    '~~ Add value to Dictionary
                    .Add sName, Left$(sRetVal, lResult)
                Else
                    Debug.Print "Duplicate Name '" & sName & "' in section '" & v_section & "' ignored!"
                End If
            End If
        Next i
    End With

xt: Set Values = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

