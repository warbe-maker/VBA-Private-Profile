Attribute VB_Name = "mPrivProfTest"
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Public PP As clsPrivProf

Private Property Let Test_Status(ByVal s As String)
    If s <> vbNullString Then
        Application.StatusBar = "Regression test " & ThisWorkbook.Name & " module 'mFso': " & s
    Else
        Application.StatusBar = vbNullString
    End If
End Property

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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPrivProfTest." & sProc
End Function

Public Sub Test_000_Regression()
' ----------------------------------------------------------------------------
' Because all results are asserted there is no manual intervention required.
' When an assertion fails the test procedure will stop, indicating the failed
' assertion. An execution trace is displayed at the end.
' Each test is autonomous. It neither uses any other but the tested Property
' or Method nor does it depend on the result of another test.
'
' Test: P/M Name                Test procedure
'       --- ------------------- ----------------------------------------------
'        P  FileName        r/w Test_010_Property_FileName
'        P  Section         r/w
'        P  Value           r/w Test_170_Method_Value2
'                               Test_220_Property_Value_Reorg
'        M  FileNameIsValid     Test_100_Method_IsValidFileFullName
'        M  NamesRemove         Test_110_.
'        M  RemoveNames         Test_120_.
'        M  SectionExists       Test_130_Methods_xExists
'        M  SectionNames        Test_140_Method_SectionNames
'        M  SectionsCopy        Test_150_Method_SectionsCopy
'        M  SectionsRemove      Test_160_.
'        M  Value2              Test_170_Method_Value2
'        M  ValueNameExists     Test_130_Methods_xExists
'        M  ValueNameRename     Test_180_.
'        M  ValueNames          Test_190_Method_ValueNames
'        M  Values              Test_200_Method_Values
' ----------------------------------------------------------------------------
    Const PROC = "Test_000_Regression"

    On Error GoTo eh
    Dim sTestStatus As String
    
    '~~ Initialization (must be done prior the first BoP!)
    mTrc.FileName = "RegressionTest_clsPrivProf.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProf"
    mTrc.NewFile
    mErh.Regression = True
    Set PP = Nothing: Set PP = New clsPrivProf ' the test runs with the default file name
    
    mBasic.BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_010_Property_FileName
    mPrivProfTest.Test_100_Method_IsValidFileFullName
    mPrivProfTest.Test_130_Methods_xExists
    mPrivProfTest.Test_140_Method_SectionNames
    mPrivProfTest.Test_150_Method_SectionsCopy
    mPrivProfTest.Test_170_Method_Value2
    mPrivProfTest.Test_190_Method_ValueNames
    mPrivProfTest.Test_200_Method_Values
    mPrivProfTest.Test_210_Method_ValueNameRename
    mPrivProfTest.Test_220_Property_Value_Reorg

xt: mTest.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    mErh.Regression = False
    mTrc.Dsply
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_010_Property_FileName()
    Const PROC = "Test_010_Property_FileName"
    
    On Error GoTo eh:
    Dim s   As String
        
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    
    '~~ Test 1: Default name
    s = PP.FileName
    Debug.Assert s = ThisWorkbook.Path & "\" & PP.FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
    
    '~~ Test 2: Specifying an invalid file valid name
    mErh.Asserted AppErr(1)
    PP.FileName = "dat"
    mErh.Asserted
    
    '~~ Test 2: Specifying a file valid name
    PP.FileName = ThisWorkbook.Path & "\Test\" & PP.FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
    
xt: mBasic.EoP ErrSrc(PROC)
    mTest.RemoveTestFiles
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_100_Method_IsValidFileFullName()
    Const PROC = "Test_100_Method_IsValidFileFullName"

    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    Debug.Assert Not PP.FileNameIsValid(ThisWorkbook.FullName)  ' not a text file
    Debug.Assert Not PP.FileNameIsValid("x")                    ' missing :, missing \
    Debug.Assert Not PP.FileNameIsValid("e:x")                  ' missing \
    Debug.Assert Not PP.FileNameIsValid("e:\x")                 ' missing extention
    Debug.Assert PP.FileNameIsValid(PP.FileName)                ' the default file is a valid file name
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub Test_130_Methods_xExists()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_130_Methods_xExists"

    On Error GoTo eh
    Dim sFileName   As String
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
       
    '~~ Section not exists
    Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(100)) = False
    '~~ Section exists
    Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(9)) = True
    '~~ Value-Name exists
    Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(7), mTest.ValueName(7, 3)) = True
    '~~ Value-Name not exists
    Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(7), mTest.ValueName(6, 3)) = False
    
xt: mTest.RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_140_Method_SectionNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_140_Method_SectionNames"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    sFileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    PP.FileName = sFileName
    Set dct = PP.SectionNames()
    Debug.Assert dct.Count = NO_OF_TEST_SECTIONS
    Debug.Assert dct.Keys()(0) = mTest.SectionName(1)
    Debug.Assert dct.Keys()(1) = mTest.SectionName(2)
    Debug.Assert dct.Keys()(2) = mTest.SectionName(3)

xt: mBasic.EoP ErrSrc(PROC)
    mTest.RemoveTestFiles
    Set dct = Nothing
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_150_Method_SectionsCopy()
' ----------------------------------------------------------------------------
' This test relies on successfully tests:
' - Test_140_Method_SectionNames (PP.sectionNames)
' Iplicitely tested are:
' - PP.sections Get and Let
' ----------------------------------------------------------------------------
    Const PROC = "Test_150_Method_SectionsCopy"
    
    On Error GoTo eh
    Dim sSourceFile     As String
    Dim sTargetFile     As String
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    sSourceFile = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES) ' prepare PrivateProfile test file
    sTargetFile = ThisWorkbook.Path & "\Test\CopyTarget.dat"
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    
    BoC "Test 1"
    '~~ Test 1 ------------------------------------
    '~~ Copy a specific section to a new target file
    '~~ Assert before
    Debug.Assert PP.Exists(sSourceFile, mTest.SectionName(5)) = True
    Debug.Assert PP.Exists(sSourceFile, mTest.SectionName(7)) = True
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(5)) = False
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(7)) = False
    
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=mTest.SectionName(5) & "," & mTest.SectionName(7)
    
    '~~ Assert result
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(5)) = True
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(7)) = True
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(5), mTest.ValueName(5, 5))
    EoC "Test 1"
    
    BoC "Test 2"
    '~~ Test 1 ------------------------------------
    PP.Section = mTest.SectionName(3)
    '~~ Copy another specific section to the target file of Test 1a
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=PP.Section
    '~~ Assert result
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(5))
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(5), mTest.ValueName(5, 5))
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(7))
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(7), mTest.ValueName(7, 5))
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(3))
    Debug.Assert PP.Exists(sTargetFile, mTest.SectionName(3), mTest.ValueName(3, 5))
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
    EoC "Test 2"
    
    BoC "Test 3"
    '~~ Test 3 -------------------------------
    '~~ Copy all sections to a new target file (will be re-ordered ascending thereby)
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=PP.SectionNames(sSourceFile) _
                  , s_merge:=False
    '~~ Assert result
    Debug.Assert StrComp(FileAsString(sTargetFile), FileAsString(sSourceFile), vbBinaryCompare) = 0
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
    EoC "Test 3"
            
xt: mTest.RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_170_Method_Value2()
' ----------------------------------------------------------------------------
' This test relies on the Value (Let) service.
' ----------------------------------------------------------------------------
    Const PROC = "Test_170_Method_Value2"
    
    On Error GoTo eh
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
            
    '~~ Test 1: Read non-existing value from a non-existing file
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert PP.Value(v_value_name:="Any" _
                        , v_section:="Any" _
                         ) = vbNullString
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 2: Read existing value
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert mTest.ValueString(3, 2) = PP.Value(v_value_name:=mTest.ValueName(3, 2) _
                                                           , v_section:=mTest.SectionName(3))
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 3: Read non-existing without Lib functions
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert vbNullString = PP.Value2(v_value_name:="x" _
                                        , v_section:=mTest.SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert mTest.ValueString(3, 2) = PP.Value2(v_value_name:=mTest.ValueName(3, 2) _
                                                            , v_section:=mTest.SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    
xt: mTest.RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_190_Method_ValueNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_190_Method_ValueNames"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    
    Set dct = PP.Values(PP.FileName, mTest.SectionName(2))
    mBasic.EoP ErrSrc(PROC)
    Debug.Assert dct.Count = mTest.NO_OF_TEST_VALUE_NAMES
    Debug.Assert dct.Keys()(0) = mTest.ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = mTest.ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = mTest.ValueName(2, 3)
       
xt: mTest.RemoveTestFiles
    Exit Sub

eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_200_Method_Values()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Method_Values"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)

    '~~ Test 1: All values of one section
    Set dct = PP.Values(PP.FileName, mTest.SectionName(2))
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = mTest.NO_OF_TEST_VALUE_NAMES
    Debug.Assert dct.Keys()(0) = mTest.ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = mTest.ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = mTest.ValueName(2, 3)
    Debug.Assert dct.Items()(0) = mTest.ValueString(2, 1)
    Debug.Assert dct.Items()(1) = mTest.ValueString(2, 2)
    Debug.Assert dct.Items()(2) = mTest.ValueString(2, 3)
    
    '~~ Test 2: No section provided
    Debug.Assert PP.Values(PP.FileName, vbNullString).Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert PP.Values(PP.FileName, "xxxxxx").Count = 0

xt: mTest.RemoveTestFiles
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_210_Method_ValueNameRename()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Method_ValueNameRename"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim i           As Long
    Dim sNameOld    As String
    Dim sNameNew    As String
    Dim sSection    As String
    
    mBasic.BoP ErrSrc(PROC)
    
    BoC "Test 1: Rename a value name in a specific section only"
    '~~         Prepare
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    sSection = mTest.SectionName(2)
    sNameOld = mTest.ValueName(2, 3)
    sNameNew = "Renamed_" & mTest.ValueName(2, 3)
    '~~        Rename in a specific section only
    PP.ValueNameRename sNameOld, sNameNew, sSection
    Debug.Assert PP.Exists(PP.FileName, sSection, sNameOld) = False
    Debug.Assert PP.Exists(PP.FileName, sSection, sNameNew) = True
    EoC "Test 1: Rename a value name in a specific section only"
    
    BoC "Test 2: Rename in all sections"
    '~~         Prepare
    Set PP = Nothing: Set PP = New clsPrivProf ' same file with different content !!!
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES, False)
    sNameOld = mTest.ValueName(, 3)
    sNameNew = "Renamed_" & mTest.ValueName(, 3)
    PP.ValueNameRename sNameOld, sNameNew
    For i = 1 To NO_OF_TEST_SECTIONS
        Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(i), sNameOld) = False
        Debug.Assert PP.Exists(PP.FileName, mTest.SectionName(i), sNameNew) = True
    Next i
    EoC "Test 2: Rename in all sections"
    
xt: mTest.RemoveTestFiles
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_220_Property_Value_Reorg()
' ----------------------------------------------------------------------------
' Re-arrange all sections and all names therein
' ----------------------------------------------------------------------------
    Const PROC = "Test_220_Property_Value_Reorg"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    PP.FileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES) ' prepare PrivateProfile test file
    vFile = mTest.FileAsArray(PP.FileName)
    Debug.Assert vFile(UBound(vFile)) = mTest.ValueName(1, 1) & "=" & mTest.ValueString(1, 1)
    
    '~~ Test: Adding a new value reorgs
    PP.Value("New_Value_Name", "New_Section") = "New_Value"
    
    '~~ Assert reorganized result
    vFile = mTest.FileAsArray(PP.FileName)
    Debug.Assert vFile(UBound(vFile)) = mTest.ValueName(NO_OF_TEST_SECTIONS, NO_OF_TEST_VALUE_NAMES) & "=" & mTest.ValueString(NO_OF_TEST_SECTIONS, NO_OF_TEST_VALUE_NAMES)
    
            
xt: mTest.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

