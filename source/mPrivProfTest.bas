Attribute VB_Name = "mPrivProfTest"
' ----------------------------------------------------------------
' Standard Module mPrivProvTest: Test of all services provided by
' ============================== the clsPrivProf class module.
'
' Uses:
' -----
' clsTestAid      Common services supporting test including
'                 regression testing.
' clsTestPrivProf Services supporting tests of methods and
'                 properties of the class module clsPrivProf.
' mTrc            Execution trace of tests.
' ----------------------------------------------------------------
Public PP       As clsPrivProf
Public Test     As New clsTestPrivProf
Public TestAid  As New clsTestAid

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
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

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
'        P  FileName        r/w Test_100_Property_FileName
'        P  Section         r/w
'        P  Value           r/w Test_120_Property_Value
'        M  IsValidFileName     Test_100_Method_IsValidFileFullName
'        M  NamesRemove         Test_110_.
'        M  RemoveNames         Test_320_.
'        M  PPdctSectionExists       Test_110_Method_Exists
'        M  SectionNames        Test_300_Method_SectionNames
'        M  SectionsCopy        Test_700_Method_SectionsCopy
'        M  SectionsRemove      Test_360_.
'        M  Exists              Test_110_Method_Exists
'        M  ValueNameRename     Test_380_.
'        M  ValueNames          Test_400_Method_ValueNames
'        M  Reorg               Test_500_Method_Reorg
' ----------------------------------------------------------------------------
    Const PROC = "Test_000_Regression"

    On Error GoTo eh
    Dim sTestStatus As String
    
    '~~ Initialization (must be done prior the first BoP!)
    Set Test = New clsTestPrivProf
    mTrc.FileFullName = TestAid.TestFolder & "\Regression.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProf"
    mTrc.NewFile
    mErH.Regression = True
    
    TestAid.ModeRegression = True
    
    mBasic.BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_100_Property_FileName
    mPrivProfTest.Test_110_Method_Exists
    mPrivProfTest.Test_120_Property_Value
    mPrivProfTest.Test_200_Property_Header
    mPrivProfTest.Test_300_Method_SectionNames
    mPrivProfTest.Test_400_Method_ValueNames
    mPrivProfTest.Test_410_Method_ValueNameRename
    mPrivProfTest.Test_500_Method_Reorg
    mPrivProfTest.Test_600_Method_Remove
    mPrivProfTest.Test_700_Method_SectionsCopy
    mPrivProfTest.Test_800_Lifecycle
    TestAid.DsplySummary
    
xt: mBasic.EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Set Test = Nothing
    Set TestAid = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_001_TestAid()
    
    Dim Fso         As New FileSystemObject
    Dim sFileResult As String
        
    With New clsTestAid
        .TestNumber = "001-1"
        .TestedComp = "clsTestAid"  ' remains the default for all subsequent tests
        .TestedProc = "Result, .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result conforms with Result expected"""
        .ResultExpected = True
        .Result = True
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "001-2"
        .TestedProc = "Result, .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result not conforms with Result expected"
        .Result = False
        .ResultExpected = True
        Debug.Assert .ResultAsExpected
        .DsplySummary ' will be bypassed since ModeRegression = False
        ' ======================================================================
        
        .TestNumber = "001-3"
        .TestedProc = "Result, .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result and .ResultExpected are files which differ"
        
        '~~ Prepare test Result file
        sFileResult = ThisWorkbook.Path & "\Test\TestResult.txt"
        TestAid.FileFromString sFileResult, "Result"
        .Result = Fso.GetFile(sFileResult)
        '~~ Prepare test ResultExpected file
        TestAid.FileFromString Test.ExpectedTestResultFileName(.TestFolder, .TestNumber), "Result expected"
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected ' This will trigger the display the difference if "False"
        .DsplySummary ' will be bypassed since ModeRegression = False
        ' ======================================================================
        
        .ModeRegression = True
        .TestNumber = "001-4"
        .TestedProc = "Result, .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Test result is Failed but passed on and displayed as summary"
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        .Result = Fso.GetFile(sFileResult)
        
        '~~ The below assertion will trigger the display the difference if "False".
        '~~ However, since ModeRegression is TRUE this assertion will be bypassed
        '~~ and the test result will be collected instead, finally displayed with .DsplySummary
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        .DsplySummary
    End With

End Sub

Public Sub Prepare(Optional ByVal p_default As Boolean = False)
    
    Test.PrivateProfile_File lNoOfTestSections, lNoOfTestValues
    Set PP = Nothing
    Set PP = New clsPrivProf
    If Not p_default Then
        PP.FileName = Test.PrivProfFileFullName
    End If

End Sub

Public Sub Test_100_Property_FileName()
    Const PROC = "Test_100_Property_FileName"
    
    On Error GoTo eh:
        
    mBasic.BoP ErrSrc(PROC)
    Prepare p_default:=True
    With TestAid
        .TestedProc = "FileName Get"
        .TestedType = "Property"
        .TestNumber = "100-1"
        .TestDscrpt = "Default name"
        .BoTP
        .Result = PP.FileName
        .EoTP
        .ResultExpected = ThisWorkbook.Path & "\" & .Fso.GetBaseName(ThisWorkbook.Name) & ".dat"
        Debug.Assert .ResultAsExpected
        
        ' ======================================================================
        PP.FileName = Test.PrivProfFileFullName ' continue with specific test file
        .TestedProc = "FileName Let"
        .TestedType = "Property"
        .TestNumber = "100-2"
        .TestDscrpt = "Specifying a file valid name"
        .BoTP
        .Result = PP.FileName
        .EoTP
        .ResultExpected = ThisWorkbook.Path & "\Test\" & PP.Fso.GetBaseName(ThisWorkbook.Name) & ".dat"
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestedProc = "FileName Let"
        .TestedType = "Property"
        .TestNumber = "100-3"
        .TestDscrpt = "Specifying an invalid file valid name"
        .AssertedErrors AppErr(1)
        .ResultExpected = AppErr(1)
        .BoTP
        PP.FileName = "dat"
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Test.RemoveTestFiles
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_110_Method_Exists()
    Const PROC = "Test_110_Methods_xExists"

    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    Prepare
       
    With TestAid
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestNumber = "110-1"
        .TestDscrpt = "Section not exists"
        .ResultExpected = False
        .BoTP
        .Result = PP.Exists(PP.FileName, Test.SectionName(7))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Section exists"
        .ResultExpected = True
        .BoTP
        .Result = PP.Exists(PP.FileName, Test.SectionName(8))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name exists"
        .ResultExpected = True
        .BoTP
        .Result = PP.Exists(PP.FileName, Test.SectionName(6), Test.ValueName(6, 4))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        .TestNumber = "110-4"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name not exists"
        .ResultExpected = False
        .BoTP
        .Result = PP.Exists(PP.FileName, Test.SectionName(6), Test.ValueName(6, 3))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
    
xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_120_Property_Value()
' ----------------------------------------------------------------------------
' This test relies on the Value (Let) service.
' ----------------------------------------------------------------------------
    Const PROC = "Test_120_Property_Value"
    
    On Error GoTo eh
    Dim cyValue     As Currency: cyValue = 12345.6789
    
    mBasic.BoP ErrSrc(PROC)
    Prepare
    
    With TestAid
        .TestNumber = "120-1"
        .TestedProc = "Value Get"
        .TestedType = "Property"
        .TestDscrpt = "Read non-existing value from a non-existing file"
        .ResultExpected = vbNullString
        .BoTP
        .Result = PP.Value(name_value:="Any" _
                              , name_section:="Any")
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "120-2"
        .TestedProc = "Value Get"
        .TestedType = "Property"
        .TestDscrpt = "Read existing value"
        .ResultExpected = Test.ValueString(2, 4)
        .BoTP
        .Result = PP.Value(name_value:=Test.ValueName(2, 4) _
                              , name_section:=Test.SectionName(2))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "120-3"
        .TestedProc = "Value Let"
        .TestedType = "Property"
        .TestDscrpt = "Write changed value"
        .BoTP
        PP.Value(name_value:=Test.ValueName(4, 2) _
                    , name_section:=Test.SectionName(4)) = "Changed value"
        .EoTP
        .ResultExpected = "Changed value"
        .Result = PP.Value(name_value:=Test.ValueName(4, 2) _
                              , name_section:=Test.SectionName(4))
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "120-4"
        .TestedProc = "Value Let"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in existing section"
        .ResultExpected = "New value, existing section"
        .BoTP
        PP.Value(Test.ValueName(2, 17) _
                    , Test.SectionName(2)) = "New value, existing section"
        .EoTP
        .Result = PP.Value(name_value:=Test.ValueName(2, 17) _
                              , name_section:=Test.SectionName(2))
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "120-5"
        .TestedProc = "Value Let"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in new section"
        .ResultExpected = "New value, new section"
        .BoTP
        PP.Value(name_value:=Test.ValueName(11, 1) _
                    , name_section:=Test.SectionName(11)) = "New value, new section"
        .EoTP
        .Result = PP.Value(name_value:=Test.ValueName(11, 1) _
                              , name_section:=Test.SectionName(11))
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
    
xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_200_Property_Header()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Property_Header"

    On Error GoTo eh
    Dim sHeader As String
    Dim sResult As String
    Dim sValue  As String
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "200-1"
        .TestedProc = "Header Let"
        .TestedType = "Property"
        .TestDscrpt = "File header write"
        .BoTP
        PP.Header() = "File Header Line 1 (the delimiter below is adjusted to the longest header)" & vbCrLf & _
                      "File Header Line 2"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        
        ' ======================================================================
        .TestNumber = "200-2"
        .TestedProc = "Header Get"
        .TestedType = "Property"
        .TestDscrpt = "File header read"
        .BoTP
        .Result = PP.Header()
        .EoTP
        .ResultExpected = .AsCollection("; File Header Line 1 (the delimiter below is adjusted to the longest header)", _
                                        "; File Header Line 2", _
                                        "; ==========================================================================")
        Debug.Assert .ResultAsExpected
        
        ' ======================================================================
        .TestNumber = "200-3"
        .TestedProc = "Header Let"
        .TestedType = "Property"
        .TestDscrpt = "Write section header"
        .BoTP
        PP.Header(, Test.SectionName(6)) = "Header Section 06 Line 1" & vbCrLf & _
                                                 "Header Section 06 Line 2"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        
        ' ======================================================================
        .TestNumber = "200-4"
        .TestedProc = "Header Get"
        .TestedType = "Property"
        .TestDscrpt = "Read section header"
        .BoTP
        .Result = PP.Header(, Test.SectionName(6))
        .EoTP
        .ResultExpected = .AsCollection("; Header Section 06 Line 1", "; Header Section 06 Line 2")
        Debug.Assert .ResultAsExpected
        
        ' =====================================================================
        .TestNumber = "200-5"
        .TestedProc = "Header Let"
        .TestedType = "Property"
        .TestDscrpt = "Write value header"
        .BoTP
        PP.Header(, Test.SectionName(6), Test.ValueName(6, 2)) = "Header Section 06 Value 02 Line 1" & vbCrLf & _
                                                                        "Header Section 06 Value 02 Line 2"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "200-6"
        .TestedProc = "Header Get"
        .TestedType = "Property"
        .TestDscrpt = "Read value header"
        .ResultExpected = .AsCollection("; Header Section 06 Value 02 Line 1", "; Header Section 06 Value 02 Line 2")
        .BoTP
        .Result = PP.Header(, Test.SectionName(6), Test.ValueName(6, 2))
        .EoTP
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        .TestNumber = "200-7"
        .TestedProc = "Value Let"
        .TestedType = "Property"
        .TestDscrpt = "Write a new value including a value header"
        .BoTP
        PP.Value(Test.ValueName(12, 1), Test.SectionName(12), , "The new value's header line 1||The new value's header line 2!") = "New value"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    
        .TestNumber = "200-8"
        .TestedProc = "Value Get"
        .TestedType = "Property"
        .TestDscrpt = "Read a value together with its header"
        .BoTP
        sValue = PP.Value(Test.ValueName(12, 1), Test.SectionName(12), , sHeader)
        .EoTP
        .Result = PP.VarItems(sHeader, enAsTextFile, TestAid.TempFile)
        .ResultExpected = PP.VarItems("; The new value's header line 1||; The new value's header line 2!", enAsTextFile, TestAid.TempFile)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With

xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_300_Method_SectionNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_300_Method_SectionNames"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "300-1"
        .TestedProc = "SectionNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all section names in a Dictionary"
        .ResultExpected = 5
        Set dct = PP.SectionNames()
        .Result = dct.Count
        Debug.Assert .ResultAsExpected
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Test.RemoveTestFiles
    Set dct = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_400_Method_ValueNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_400_Method_ValueNames"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "400-1"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all value names of all sections in a Dictionary"
        .ResultExpected = 40
        Set dct = PP.ValueNames()
        .Result = dct.Count
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    
        .TestNumber = "400-2"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all value names of a certain section in a Dictionary"
        .ResultExpected = 8
        .Result = PP.ValueNames(, Test.SectionName(6)).Count
        Debug.Assert .ResultAsExpected
    
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_410_Method_ValueNameRename()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_400_Method_ValueNameRename"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "410-1"
        .TestedProc = "ValueNameRename"
        .TestedType = "Method"
        .TestDscrpt = "Rename a value name in each section."
        PP.ValueNameRename Test.ValueName(2, 2), "Renamed_" & Test.ValueName(2, 2)
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
    
xt: Test.RemoveTestFiles
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_500_Method_Reorg()
' ----------------------------------------------------------------------------
' Re-arrange all sections and all names therein
' ----------------------------------------------------------------------------
    Const PROC = "Test_500_Method_Reorg"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "500-1"
        .TestedProc = "Reorg"
        .TestedType = "Method"
        .TestDscrpt = "Reorganizes a Private Profile file considering headers and a file footer."
        PP.Reorg
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
            
xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_600_Method_Remove()
' ----------------------------------------------------------------------------
' The test relies on: - Header value
' ----------------------------------------------------------------------------
    Const PROC = "Test_700_Method_SectionsCopy"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
    
    With TestAid
        PP.Header(, Test.SectionName(6), Test.ValueName(6, 4)) = "Header value 06-04"
        .TestNumber = "600-1"
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes a value from a section including its header comment(s)."
        PP.ValueRemove Test.ValueName(6, 4), Test.SectionName(6)
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
        
        PP.Header(, Test.SectionName(6)) = "Header section 06"
        .TestNumber = "600-2"
        .TestedProc = "SectionRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes a section including its header comment(s)."
        PP.SectionRemove Test.SectionName(6)
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With

xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_700_Method_SectionsCopy()
' ----------------------------------------------------------------------------
' This test relies on: - method SectionNames (Test_300_Method_SectionNames),
'                      - method SectionRemove (Test_600_Method_Remove)
' The test implicitely tests the property Sections Get/Let.
' ----------------------------------------------------------------------------
    Const PROC = "Test_700_Method_SectionsCopy"
    
    On Error GoTo eh
    Dim sSourceFile     As String
    Dim sTargetFile     As String
    
    mBasic.BoP ErrSrc(PROC)
    Prepare ' Test preparation
    sTargetFile = ThisWorkbook.Path & "\Test\CopyTarget.dat"
    
    With TestAid
        If .Fso.FileExists(sTargetFile) Then .Fso.DeleteFile sTargetFile
        .TestNumber = "700-1"
        .TestedProc = "SectionsCopy"
        .TestedType = "Method"
        .TestDscrpt = "Copies two sections from a soure to a traget Private Profile file."
        .BoTP
        PP.SectionsCopy name_file_source:=sSourceFile _
                           , name_file_target:=sTargetFile _
                           , name_sections:=Test.SectionName(6) & "," & Test.SectionName(2)
        .EoTP
        .Result = .Fso.GetFile(sTargetFile)
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    
        .TestNumber = "700-2"
        .TestedProc = "SectionsCopy"
        .TestedType = "Method"
        .TestDscrpt = "Copies an additional sections from a soure to a traget Private Profile file."
        .BoTP
        PP.SectionsCopy name_file_source:=sSourceFile _
                           , name_file_target:=sTargetFile _
                           , name_sections:=Test.SectionName(4)
        .EoTP
        .Result = .Fso.GetFile(sTargetFile)
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    End With
                
xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_800_Lifecycle()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = "Test_800_Lifecycle"
    
    On Error GoTo eh
   
    mBasic.BoP ErrSrc(PROC)
    
    With TestAid
        Prepare ' Test preparation
        .Fso.DeleteFile Test.PrivProfFileFullName
        .TestNumber = "800-1"
        .TestedProc = "Header, Footer"
        .TestedType = "Method"
        .TestDscrpt = "Writes a file header and footer into an empty file."
        .BoTP
        PP.Header() = "File Header Line 1 (the delimiter below is adjusted to the longest header)" & vbCrLf & _
                      "File Header Line 2"
        PP.Footer() = "File Footer Line 1 (the delimiter below is adjusted to the longest header)" & vbCrLf & _
                      "File Footer Line 2"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    
        .TestNumber = "800-2"
        .TestedProc = "Value Let"
        .TestedType = "Propety"
        .TestDscrpt = "Writes a file header and footer into an empty file."
        .BoTP
        PP.Value("Value_Name", "Section_Name") = "Value"
        .EoTP
        .Result = Test.PrivProfFile
        .ResultExpected = Test.ExpectedTestResultFile(.TestFolder, .TestNumber)
        Debug.Assert .ResultAsExpected
        ' ======================================================================
    
    End With

xt: Test.RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
