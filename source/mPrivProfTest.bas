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
'                               Test_031_Property_Value_Reorg
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
    Set PP = New clsPrivProf ' the test runs with the default file name
    
    mBasic.BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_010_Property_FileName
    mPrivProfTest.Test_100_Method_IsValidFileFullName
    mPrivProfTest.Test_140_Method_SectionNames
    mPrivProfTest.Test_190_Method_ValueNames
    mPrivProfTest.Test_170_Method_Value2
    mPrivProfTest.Test_200_Method_Values
    mPrivProfTest.Test_130_Methods_xExists
    mPrivProfTest.Test_150_Method_SectionsCopy
    mPrivProfTest.Test_031_Property_Value_Reorg

xt: mTest.TestProc_RemoveTestFiles
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
    Const PROC = " Test_010_Property_FileName"
    
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
    TestProc_RemoveTestFiles
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_100_Method_IsValidFileFullName()
    Const PROC = " Test_100_Method_IsValidFileFullName"

    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProf
    Debug.Assert PP.FileNameIsValid(ThisWorkbook.FullName)
    Debug.Assert Not PP.FileNameIsValid("x")    ' missing :, missing \
    Debug.Assert Not PP.FileNameIsValid("e:x")  ' missing \
    Debug.Assert Not PP.FileNameIsValid("e:\x") ' missing extention
    Debug.Assert PP.FileNameIsValid("e:\x.y")   ' complete with extention
    mBasic.EoP ErrSrc(PROC)

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
    sFileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    PP.FileName = sFileName
    Set dct = PP.SectionNames()
    Debug.Assert dct.Count = NO_OF_TEST_SECTIONS
    Debug.Assert dct.Keys()(0) = mTest.TestProc_SectionName(1)
    Debug.Assert dct.Keys()(1) = mTest.TestProc_SectionName(2)
    Debug.Assert dct.Keys()(2) = mTest.TestProc_SectionName(3)

xt: mBasic.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set dct = Nothing
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
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    
    Set dct = PP.Values(v_section:=TestProc_SectionName(2))
    mBasic.EoP ErrSrc(PROC)
    Debug.Assert dct.Count = mTest.NO_OF_TEST_VALUE_NAMES
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
       
xt: TestProc_RemoveTestFiles
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
    Set PP = New clsPrivProf
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
            
    '~~ Test 1: Read non-existing value from a non-existing file
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert PP.Value(v_value_name:="Any" _
                        , v_section:="Any" _
                         ) = vbNullString
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 2: Read existing value
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert mTest.TestProc_ValueString(3, 2) = PP.Value(v_value_name:=mTest.TestProc_ValueName(3, 2) _
                                                           , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 3: Read non-existing without Lib functions
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert vbNullString = PP.Value2(v_value_name:="x" _
                                        , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert mTest.TestProc_ValueString(3, 2) = PP.Value2(v_value_name:=mTest.TestProc_ValueName(3, 2) _
                                                            , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    
xt: TestProc_RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
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
    Set PP = New clsPrivProf
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)

    '~~ Test 1: All values of one section
    Set dct = PP.Values(v_section:=TestProc_SectionName(2))
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = mTest.NO_OF_TEST_VALUE_NAMES
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
    Debug.Assert dct.Items()(0) = TestProc_ValueString(2, 1)
    Debug.Assert dct.Items()(1) = TestProc_ValueString(2, 2)
    Debug.Assert dct.Items()(2) = TestProc_ValueString(2, 3)
    
    '~~ Test 2: No section provided
    Debug.Assert PP.Values().Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert PP.Values(v_file:=PP.FileName _
                         , v_section:="xxxxxxx").Count = 0

xt: TestProc_RemoveTestFiles
    Set dct = Nothing
    Set PP = Nothing
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
    
    '~~ Test 1: Rename in a specific section only
    '~~         Prepare
    Set PP = Nothing: Set PP = New clsPrivProf
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    sSection = mTest.TestProc_SectionName(2)
    sNameOld = TestProc_ValueName(2, 3)
    sNameNew = "Renamed_" & TestProc_ValueName(2, 3)
    '~~        Rename in a specific section only
    PP.ValueNameRename sNameOld, sNameNew, sSection
    Debug.Assert Not PP.ValueNameExists(sNameOld, sSection)
    Debug.Assert PP.ValueNameExists(sNameNew, sSection)
    
    '~~ Test 2: Rename in all sections
    '~~         Prepare
    Set PP = Nothing: Set PP = New clsPrivProf ' same file with different content !!!
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES, False)
    sNameOld = TestProc_ValueName(, 3)
    sNameNew = "Renamed_" & TestProc_ValueName(, 3)
    PP.ValueNameRename sNameOld, sNameNew
    For i = 1 To NO_OF_TEST_SECTIONS
        Debug.Assert Not PP.ValueNameExists(sNameOld, mTest.TestProc_SectionName(i))
        Debug.Assert PP.ValueNameExists(sNameNew, mTest.TestProc_SectionName(i))
    Next i
    
xt: TestProc_RemoveTestFiles
    Set dct = Nothing
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    Set PP = New clsPrivProf
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
       
    '~~ Section not exists
    Debug.Assert PP.SectionExists(s_file:=PP.FileName _
                                , s_section:=TestProc_SectionName(100) _
                                 ) = False
    '~~ Section exists
    Debug.Assert PP.SectionExists(s_file:=sFileName _
                                , s_section:=TestProc_SectionName(9) _
                                 ) = True
    '~~ Value-Name exists
    Debug.Assert PP.ValueNameExists(v_file:=sFileName _
                                  , v_section:=TestProc_SectionName(7) _
                                  , v_value_name:=TestProc_ValueName(7, 3) _
                                   ) = True
    '~~ Value-Name not exists
    Debug.Assert PP.ValueNameExists(v_file:=sFileName _
                                  , v_section:=TestProc_SectionName(7) _
                                  , v_value_name:=TestProc_ValueName(6, 3) _
                                   ) = False
    
xt: TestProc_RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
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
    Set PP = New clsPrivProf
    sSourceFile = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES) ' prepare PrivateProfile test file
    sTargetFile = ThisWorkbook.Path & "\Test\CopyTarget.dat"
    
    '~~ Test 1a ------------------------------------
    '~~ Copy a specific section to a new target file
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(5), sSourceFile)
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(7), sSourceFile)
    Debug.Assert Not PP.SectionExists(mTest.TestProc_SectionName(5), sTargetFile)
    Debug.Assert Not PP.SectionExists(mTest.TestProc_SectionName(7), sTargetFile)
    
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=mTest.TestProc_SectionName(5) & "," & mTest.TestProc_SectionName(7)
    
    '~~ Assert result
    PP.FileName = sTargetFile
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(5))
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(7))
    Debug.Assert PP.ValueNameExists(TestProc_ValueName(5, 5), mTest.TestProc_SectionName(5))
    
    '~~ Test 1b ------------------------------------
    PP.Section = mTest.TestProc_SectionName(3)
    '~~ Copy another specific section to the target file of Test 1a
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=PP.Section
    '~~ Assert result
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(5))
    Debug.Assert PP.ValueNameExists(TestProc_ValueName(5, 5), mTest.TestProc_SectionName(5))
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(7))
    Debug.Assert PP.ValueNameExists(TestProc_ValueName(7, 5), mTest.TestProc_SectionName(7))
    Debug.Assert PP.SectionExists(mTest.TestProc_SectionName(3))
    Debug.Assert PP.ValueNameExists(TestProc_ValueName(3, 5), mTest.TestProc_SectionName(3))
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
    
    '~~ Test 3 -------------------------------
    '~~ Copy all sections to a new target file (will be re-ordered ascending thereby)
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=PP.SectionNames(sSourceFile) _
                  , s_merge:=False
    '~~ Assert result
    Debug.Assert FileAsString(sTargetFile) = FileAsString(sSourceFile)
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
            
xt: TestProc_RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_031_Property_Value_Reorg()
' ----------------------------------------------------------------------------
' Rearrange all sections and all names therein
' ----------------------------------------------------------------------------
    Const PROC = "Test_031_Property_Value_Reorg"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProf
    PP.FileName = mTest.TestProc_PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES) ' prepare PrivateProfile test file
    vFile = mTest.FileAsArray(PP.FileName)
    Debug.Assert vFile(UBound(vFile)) = TestProc_ValueName(1, 1) & "=" & TestProc_ValueString(1, 1)
    
    '~~ Test: Adding a new value reorgs
    PP.Value("New_Value_Name", "New_Section") = "New_Value"
    
    '~~ Assert reorganized result
    vFile = mTest.FileAsArray(PP.FileName)
    Debug.Assert vFile(UBound(vFile)) = TestProc_ValueName(NO_OF_TEST_SECTIONS, NO_OF_TEST_VALUE_NAMES) & "=" & TestProc_ValueString(NO_OF_TEST_SECTIONS, NO_OF_TEST_VALUE_NAMES)
    
            
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub




