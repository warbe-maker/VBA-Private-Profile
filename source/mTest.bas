Attribute VB_Name = "mTest"
Option Explicit
Public Const NO_OF_TEST_SECTIONS    As Long = 10
Public Const NO_OF_TEST_VALUE_NAMES As Long = 16

Public TestPPf                 As clsPrivProf
Public Test                         As clsTestAid
Public Fso                          As New FileSystemObject

Private Const SECTION_NAME          As String = "Section_"      ' for PrivateProfile services test
Private Const VALUE_NAME_INDIVIDUAL As String = "_Name_"        ' for PrivateProfile services test
Private Const VALUE_NAME            As String = "Value_Name_"   ' for PrivateProfile services test
Private Const VALUE_STRING          As String = "-Value-"       ' for PrivateProfile services test

Private cllTestFiles                As Collection
Private sPrivProfFile               As String
Private sFileName                   As String

Public Function ExpectedTestResultFileName(ByVal e_test_folder As String, _
                                           ByVal e_test_no As String) As String
    ExpectedTestResultFileName = e_test_folder & "\ResultExpected-" & e_test_no & ".dat"

End Function

Public Function ExpectedTestResultFile(ByVal e_test_folder As String, _
                                       ByVal e_test_no As String) As File
     Set ExpectedTestResultFile = Fso.GetFile(ExpectedTestResultFileName(e_test_folder, e_test_no))
     
End Function

Public Property Get FileName() As String:   FileName = sFileName:   End Property

Public Sub Init(Optional ByVal i_default As Boolean = False)
    
    If TestPPf Is Nothing Then Set TestPPf = New clsPrivProf
    If Test Is Nothing Then Set Test = New clsTestAid
    Test.TestedComp = "clsPrivProf"
    sFileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    If Not i_default Then
        TestPPf.FileName = sFileName
    End If
End Sub

Public Property Let FileString(Optional ByVal f_file_full_name As String, _
                               Optional ByVal f_append As Boolean = False, _
                               Optional ByVal f_exclude_empty As Boolean = False, _
                                        ByVal f_s As String)
' ----------------------------------------------------------------------------
' Writes a string (f_s) with multiple records/lines delimited by a vbCrLf to
' a file (f_file_full_name).
' ----------------------------------------------------------------------------
    
    If f_append _
    Then Open f_file_full_name For Append As #1 _
    Else Open f_file_full_name For Output As #1
    Print #1, f_s
    Close #1
        
End Property

Public Function FileAsString(Optional ByVal f_file_full_name As String, _
                             Optional ByVal f_append As Boolean = False, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    Open f_file_full_name For Input As #1
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
    
End Function

Public Function FileAsArray(ByVal f_file_full_name As String) As Variant
    Dim sSplit As String
    FileAsArray = Split(FileAsString(f_file_full_name, , sSplit), sSplit)
End Function

Public Sub FileFromString(ByVal f_file_full_name As String, _
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

Private Property Get FileTemp(Optional ByVal f_path As String = vbNullString, _
                              Optional ByVal f_extension As String = ".tmp") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(TestPPf.Fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest." & sProc
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

Public Property Get PrivProfFile() As File
    Set PrivProfFile = Fso.GetFile(sPrivProfFile)
End Property

Private Sub ArrayAdd(ByRef a_array As Variant, _
                     ByVal a_str As String)
    On Error Resume Next
    ReDim Preserve a_array(UBound(a_array) + 1)
    If Err.Number <> 0 Then ReDim a_array(0)
    a_array(UBound(a_array)) = a_str
    
End Sub

Public Function PrivateProfile_File(Optional ByVal t_sections As Long = NO_OF_TEST_SECTIONS, _
                                    Optional ByVal t_values As Long = NO_OF_TEST_VALUE_NAMES, _
                                    Optional ByVal t_individual_names As Boolean = True) As String
' ----------------------------------------------------------------------------
' Returns the name of a temporary file with n (t_sections) sections, each
' with m (t_values) values all in descending order. Each test file's name is
' saved to a Collection (cllTestFiles) allowing to delete them all at the end
' of the test.
' When t_individual_names is FALSE all sections have the same set of value
' names.
' ----------------------------------------------------------------------------
    Const PROC = "PrivateProfile_File"

    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim sFolder     As String
    Dim arr()       As Variant
    
    sFolder = ThisWorkbook.Path & "\Test"
    If Not TestPPf.Fso.FolderExists(sFolder) Then TestPPf.Fso.CreateFolder (sFolder)
    
    sPrivProfFile = sFolder & "\" & TestPPf.Fso.GetBaseName(ThisWorkbook.Name) & ".dat"
    If TestPPf.Fso.FileExists(sPrivProfFile) Then TestPPf.Fso.DeleteFile sPrivProfFile
    
    For i = t_sections To 1 Step -2
        ArrayAdd arr, "[" & SectionName(i) & "]"
        For j = t_values To 1 Step -2
            If t_individual_names _
            Then ArrayAdd arr, ValueName(i, j) & "=" & ValueString(i, j) _
            Else ArrayAdd arr, ValueName(, j) & "=" & ValueString(i, j)
        Next j
    Next i
    FileFromString sPrivProfFile, Join(arr, vbCrLf)
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sPrivProfFile
    PrivateProfile_File = sPrivProfFile
    
xt: Exit Function

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RemoveTestFiles()

    Dim v As Variant
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    For Each v In cllTestFiles
        If TestPPf.Fso.FileExists(v) Then
            Kill v
        End If
    Next v
    Set cllTestFiles = Nothing
    Set cllTestFiles = New Collection
    
End Sub

Public Function SectionName(ByVal l As Long) As String
    SectionName = SECTION_NAME & Format(l, "00")
End Function

Public Function TempFile() As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "TempFile"

    On Error GoTo eh
    Dim sFileName   As String

    mBasic.BoP ErrSrc(PROC)
    sFileName = FileTemp(f_extension:=".dat")
    TempFile = sFileName

    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sPrivProfFile

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueName(Optional ByVal t_section_name As Long = 0, _
                                   Optional ByVal t_value_name As Long = 0) As String
    If t_section_name <> 0 _
    Then ValueName = SECTION_NAME & Format(t_section_name, "00") & VALUE_NAME_INDIVIDUAL & Format(t_value_name, "00") _
    Else ValueName = VALUE_NAME & Format(t_value_name, "00")
    
End Function

Public Function ValueString(ByVal lS As Long, ByVal lV As Long) As String
    ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Function

