Attribute VB_Name = "mTest"
Option Explicit
Public Const NO_OF_TEST_SECTIONS    As Long = 10
Public Const NO_OF_TEST_VALUE_NAMES As Long = 15

Private Const SECTION_NAME          As String = "Section_"      ' for PrivateProfile services test
Private Const VALUE_NAME_INDIVIDUAL As String = "_Name_"        ' for PrivateProfile services test
Private Const VALUE_NAME            As String = "Value_Name_"   ' for PrivateProfile services test
Private Const VALUE_STRING          As String = "-Value-"       ' for PrivateProfile services test

Private cllTestFiles    As Collection

Public Function FileAsString(Optional ByVal f_file_full_name As String, _
                             Optional ByVal f_append As Boolean = False, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    Open f_file_full_name For Input As #1
    FileAsString = Input$(lOf(1), 1)
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
    sTemp = Replace(PP.FSo.GetTempName, ".tmp", f_extension)
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

Public Function TestProc_PrivateProfile_File(Optional ByVal t_sections As Long = 10, _
                                             Optional ByVal t_values As Long = 15, _
                                             Optional ByVal t_individual_names As Boolean = True) As String
' ----------------------------------------------------------------------------
' Returns the name of a temporary file with n (t_sections) sections, each
' with m (t_values) values all in descending order. Each test file's name is
' saved to a Collection (cllTestFiles) allowing to delete them all at the end
' of the test.
' When t_individual_names is FALSE all sections have the same set of value
' names.
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_PrivateProfile_File"

    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim sFileName   As String
    Dim sFolder     As String
    Dim arr()       As Variant
    Dim k           As Long
    
    mBasic.BoP ErrSrc(PROC)
    sFolder = ThisWorkbook.Path & "\Test"
    If Not PP.FSo.FolderExists(sFolder) Then PP.FSo.CreateFolder (sFolder)
    
    sFileName = sFolder & "\" & PP.FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
    If PP.FSo.FileExists(sFileName) Then PP.FSo.DeleteFile sFileName
    
    ReDim Preserve arr((t_sections * t_values) + t_sections - 1)
    For i = t_sections To 1 Step -1
        arr(k) = "[" & TestProc_SectionName(i) & "]"
        k = k + 1
        For j = t_values To 1 Step -1
            If t_individual_names _
            Then arr(k) = TestProc_ValueName(i, j) & "=" & TestProc_ValueString(i, j) _
            Else arr(k) = TestProc_ValueName(, j) & "=" & TestProc_ValueString(i, j)
            k = k + 1
        Next j
    Next i
    FileFromString sFileName, Join(arr, vbCrLf)
    TestProc_PrivateProfile_File = sFileName
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFileName
    
    PP.FileName = sFileName ' this assignment immediately loadds it as Dictionary

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub TestProc_RemoveTestFiles()

    Dim v As Variant
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    For Each v In cllTestFiles
        If PP.FSo.FileExists(v) Then
            Kill v
        End If
    Next v
    Set cllTestFiles = Nothing
    Set cllTestFiles = New Collection
    
End Sub

Public Function TestProc_SectionName(ByVal l As Long) As String
    TestProc_SectionName = SECTION_NAME & Format(l, "00")
End Function

Public Function TestProc_TempFile() As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_TempFile"

    On Error GoTo eh
    Dim sFileName   As String

    mBasic.BoP ErrSrc(PROC)
    sFileName = FileTemp(f_extension:=".dat")
    TestProc_TempFile = sFileName

    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFileName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function TestProc_ValueName(Optional ByVal t_section_name As Long = 0, _
                                   Optional ByVal t_value_name As Long = 0) As String
    If t_section_name <> 0 _
    Then TestProc_ValueName = SECTION_NAME & Format(t_section_name, "00") & VALUE_NAME_INDIVIDUAL & Format(t_value_name, "00") _
    Else TestProc_ValueName = VALUE_NAME & Format(t_value_name, "00")
    
End Function

Public Function TestProc_ValueString(ByVal lS As Long, ByVal lV As Long) As String
    TestProc_ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Function

