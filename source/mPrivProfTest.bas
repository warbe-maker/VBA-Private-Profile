Attribute VB_Name = "mPrivProfTest"
Option Explicit
#Const mTrc = 1
' ----------------------------------------------------------------
' Standard Module mPrivProvTest: Test of all services provided by
' ============================== the clsPrivProf class module.
' Usually each test is autonomous and preferrably uses no or only
' tested other Properties/Methods.
'
' Uses:
' - clsTestAid      Common services supporting test including
'                   regression testing.
' - clsPrivProfTests Services supporting tests of methods and
'                   properties of the class module clsPrivProf.
' - mTrc            Execution trace of tests.
'
' W. Rauschenberger, Berlin May 2024
' See also https://github.com/warbe-maker/VBA-Private-Profile.
' ----------------------------------------------------------------
Public PrivProf         As clsPrivProf
Public PrivProfTests    As New clsPrivProfTests
Public TestAid            As clsTestAid
Private cllExpctd       As Collection
Private FSo             As New FileSystemObject

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

Public Sub Prepare(Optional ByVal p_no As Long = 1, _
                   Optional ByVal p_init As Boolean = True)
' ----------------------------------------------------------------------------
' Prepares for a new test:
' 1. A test Private Profile file considering a nmber (p_no)
' 2. A new clsPrivProf class instance
' 3. By default, initializes the FileName property (p_init)
' Note: By default a file ....1.dat (p_no) is setup from scratch, other
' numbers (p_no) may just copy a backup.
' ----------------------------------------------------------------------------
    Const PROC = "Prepare"
    
    Dim sFile As String
    On Error GoTo eh
    If TestAid Is Nothing Then Set TestAid = New clsTestAid
    Set PrivProf = Nothing
    Set PrivProf = New clsPrivProf
    PrivProfTests.ProvideTestPrivProf p_no
    If p_init And p_no <> 0 Then PrivProf.FileName = PrivProfTests.PrivProfFileFullName

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_000_Regression()
' ----------------------------------------------------------------------------
' Please note: All results are programmatically asserted and thus there is no
' manual intervention during this test. In case an assertion fails the test
' procedure will  n o t  stop but keep a record of the failed assertion.
'
' An execution trace is displayed at the end.
' ----------------------------------------------------------------------------
    Const PROC = "Test_000_Regression"

    On Error GoTo eh
    Dim sTestStatus     As String
    Dim bModeRegression As Boolean
    
    '~~ Initialization (must be done prior the first BoP!)
    Set PrivProfTests = New clsPrivProfTests
    mTrc.FileFullName = TestAid.TestFolder & "\Regression.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProf"
    mTrc.NewFile
    bModeRegression = True
    mErH.Regression = bModeRegression
    TestAid.ModeRegression = bModeRegression
    TestAid.TestFilesRemove "Result_" ' remove any files resulting from individual tests
    
    BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_001_TestAid
    mPrivProfTest.Test_100_Property_FileName
    mPrivProfTest.Test_110_Method_Exists
    mPrivProfTest.Test_120_Property_Value
    mPrivProfTest.Test_200_Property_Comments
    mPrivProfTest.Test_300_Method_SectionNames
    mPrivProfTest.Test_400_Method_ValueNames
    mPrivProfTest.Test_410_Method_ValueNameRename
    mPrivProfTest.Test_500_Method_Remove
    mPrivProfTest.Test_600_Lifecycle
    mPrivProfTest.Test_700_HskpngNames
    TestAid.DsplySummary
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Set TestAid = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_001_TestAid()
' ----------------------------------------------------------------------------
' Test of the means (clsTestAid) used by all tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_001_TestAid"
    
    On Error GoTo eh
    Dim sFileResult     As String
    Dim sFileExpected   As String
    
    BoP ErrSrc(PROC)
    Prepare 1, False
    With TestAid
        .ModeRegression = mErH.Regression
        .TestNumber = "001-1"
        .TestedComp = "clsPrivProf"
        .TestDscrpt = "Initialize with a new PP-file"
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName

        
        .TestNumber = "001-2"
        .TestedComp = "clsTestAid"  ' remains the default for all subsequent tests
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result is the result expected"
        .ResultExpected = True
        .BoTP
        .Result = True
        .EoTP
        ' ======================================================================
        
        .TestNumber = "001-3"
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result is  F a i l e d  because the result/expected boolean differs"
        .ResultExpected = True
        .BoTP
        .Result = False
        .EoTP
        ' ======================================================================
        
        .TestNumber = "001-4"
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result is  F a i l e d  because the result/expected files differ"
        
        '~~ Prepare test Result file
        sFileResult = ThisWorkbook.Path & "\Test\TestResult.txt"
        .ResultExpected = .StringAsFile("AAAAAA", sFileExpected)
        .TestFile = sFileExpected
        '~~ Prepare test ResultExpected file
        .BoTP
        .Result = .StringAsFile("BBBBBBBB", sFileResult)
        .EoTP
        .TestFile = sFileResult
        ' ======================================================================
        
        .ModeRegression = True
        .TestNumber = "001-5"
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .TestDscrpt = "Result is  F a i l e d  because result/expected files differ"
        .ResultExpected = .StringAsFile("ResultExpected", sFileExpected)
        .TestFile = sFileExpected
        .BoTP
        .Result = .StringAsFile("Result", sFileResult)
        .EoTP
        .TestFile = sFileResult
        ' ======================================================================
        If Not mErH.Regression Then .DsplySummary
    End With

xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_100_Property_FileName()
    Const PROC = "Test_100_Property_FileName"
    
    On Error GoTo eh:
        
    BoP ErrSrc(PROC)
    Prepare 1, False ' The FileName property is provided in the test
    With TestAid
        .TestNumber = "100-1"
        .TestedComp = "clsPrivProf"
        .TestedProc = "FileName_Let"
        .TestedType = "Property"
        .TestDscrpt = "Initialize PP-file"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "100-2"
        .TestedProc = "Let FileName"
        .TestedType = "Property"
        .TestDscrpt = "Specifying a file valid name"
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName ' continue with current test file
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_110_Method_Exists()
    Const PROC = "Test_110_Methods_xExists"

    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    '~~ Test preparation
    Prepare
       
    With TestAid
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestNumber = "110-1"
        .TestDscrpt = "Section not exists"
        .ResultExpected = False
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(7))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Section exists"
        .ResultExpected = True
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(8))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name exists"
        .ResultExpected = True
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-4"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name not exists"
        .ResultExpected = False
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 3))
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
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
    
    BoP ErrSrc(PROC)
    Prepare
    
    With TestAid
        .TestNumber = "120-1"
        .TestedProc = "Get Value"
        .TestedType = "Property"
        .TestDscrpt = "Read non-existing value from a non-existing file"
        .ResultExpected = vbNullString
        .BoTP
        .Result = PrivProf.Value(name_value:="Any" _
                               , name_section:="Any")
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-2"
        .TestedProc = "Get Value"
        .TestedType = "Property"
        .TestDscrpt = "Read existing value"
        .ResultExpected = PrivProfTests.ValueString(2, 4)
        .BoTP
        .Result = PrivProf.Value(name_value:=PrivProfTests.ValueName(2, 4) _
                               , name_section:=PrivProfTests.SectionName(2))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-3"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write changed value"
        .ResultExpected = "Changed value"
        .BoTP
        PrivProf.Value(name_value:=PrivProfTests.ValueName(4, 2) _
                     , name_section:=PrivProfTests.SectionName(4)) = "Changed value"
        .Result = PrivProf.Value(name_value:=PrivProfTests.ValueName(4, 2) _
                               , name_section:=PrivProfTests.SectionName(4))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-4"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in existing section"
        .ResultExpected = "New value, existing section"
        .BoTP
        PrivProf.Value(PrivProfTests.ValueName(2, 17) _
                    , PrivProfTests.SectionName(2)) = "New value, existing section"
        .Result = PrivProf.Value(name_value:=PrivProfTests.ValueName(2, 17) _
                              , name_section:=PrivProfTests.SectionName(2))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-5"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in new section"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11)) = "New value, new section"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-6"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Change value plus the value and the section comments"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11) _
                      ) = "Changed new value, new section"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-7"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Changed again"
        .BoTP
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11) _
                      ) = "Changed again new value, new section"
        .Result = PrivProfTests.PrivProfFile
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_200_Property_Comments()
' ----------------------------------------------------------------------------
' Let/Get comments (FileHeader, FileFooter, SectionComment, ValueComment)
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Property_Comments"

    On Error GoTo eh
    Dim sHeader         As String
    Dim sResult         As String
    Dim sValue          As String
    Dim sCommentValue   As String
    Dim sCommentSect    As String
    Dim sFileHeader     As String
    Dim sFileFooter     As String
    Dim cllResultExpectd As Collection
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "200-1"
        .TestedProc = "FileHeader-Let"
        .TestedType = "Property"
        .TestDscrpt = "Write a file header"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        '~~ Note: For the missing file name the property FileName is used
        '~~ and the missing section- and value-name indicate a file header
        PrivProf.FileHeader() = "File Comment Line 1 (the comments delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Comment Line 2"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================

        .TestNumber = "200-3"
        .TestedProc = "FileHeader-Get"
         .TestedType = "Property"
        .TestDscrpt = "File comment read"
        Set cllResultExpectd = New Collection
        cllResultExpectd.Add "; File Comment Line 1 (the comments delimiter below is adjusted to the longest comment)"
        cllResultExpectd.Add "; File Comment Line 2"
        cllResultExpectd.Add "; ====================================================================================="
        .ResultExpected = cllResultExpectd
        .BoTP
        .Result = TestAid.StringAsCollection(PrivProf.FileHeader())
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-4"
        .TestedProc = "SectionComment-Let"
        .TestedType = "Property"
        .TestDscrpt = "Write section comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) = "Comment Section 06 Line 1" _
                                                     & vbCrLf & "Comment Section 06 Line 2"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-5"
        .TestedProc = "SectionComment-Get"
        .TestedType = "Property"
        .TestDscrpt = "Read section comment"
        .ResultExpected = "Comment Section 06 Line 1" _
               & vbCrLf & "Comment Section 06 Line 2"
        .BoTP
        .Result = PrivProf.SectionComment(PrivProfTests.SectionName(6))
        .EoTP
        ' =====================================================================
        
        .TestNumber = "200-6"
        .TestedProc = "ValueComment-Let"
        .TestedType = "Property"
        .TestDscrpt = "Write value comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6)) = "Comment Section 06 Value 02 Line 1" _
                                                                                  & vbCrLf & "Comment Section 06 Value 02 Line 2"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-7"
        .TestedProc = "Value-Comment-Get"
        .TestedType = "Property"
        .TestDscrpt = "Read value comment"
        .ResultExpected = "Comment Section 06 Value 02 Line 1" _
               & vbCrLf & "Comment Section 06 Value 02 Line 2"
        .BoTP
        .Result = PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6))
        .EoTP
        ' ======================================================================
            
    End With

xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    If Not mErH.Regression Then mTrc.Dsply
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
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "300-1"
        .TestedProc = "SectionNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all section names in a Dictionary"
        .ResultExpected = 5
        .BoTP
        .Result = PrivProf.SectionNames().Count
        .EoTP
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
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
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "400-1"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all value names of all sections in a Dictionary"
        .ResultExpected = 40
        .BoTP
        Set dct = PrivProf.ValueNames()
        .Result = dct.Count
        .EoTP
        ' ======================================================================
    
        .TestNumber = "400-2"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all value names of a certain section in a Dictionary"
        .ResultExpected = 8
        .BoTP
        .Result = PrivProf.ValueNames(, PrivProfTests.SectionName(6)).Count
        .EoTP
        ' ======================================================================
      
        .TestNumber = "400-3"
        Prepare 2
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all value names of all sections in a Dictionary"
        .BoTP
        .Result = PrivProf.ValueNames().Count
        .ResultExpected = 6
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
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
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .TestNumber = "410-1"
        .TestedProc = "ValueNameRename"
        .TestedType = "Method"
        .TestDscrpt = "Rename a value name in each section."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueNameRename PrivProfTests.ValueName(2, 2), "Renamed_" & PrivProfTests.ValueName(2, 2)
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_500_Method_Remove()
' ----------------------------------------------------------------------------
' The test relies on: - Comment value
' ----------------------------------------------------------------------------
    Const PROC = "Test_500_Method_Remove"
    
    On Error GoTo eh
    Dim sFile As String
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
    
    With TestAid
        PrivProf.ValueComment(PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4)) = "Comment value 06-04"
        .TestNumber = "500-1"
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Remove a value from a section including its comments."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueRemove PrivProfTests.ValueName(6, 4), PrivProfTests.SectionName(6)
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
        
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) = "Comment section 06"
        .TestNumber = "500-2"
        .TestedProc = "SectionRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes a section including its comments."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.SectionRemove PrivProfTests.SectionName(6)
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
    
        .TestNumber = "500-3"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Remove 2 names in 2 sections."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueRemove name_value:="Last_Modified_AtDateTime,Last_Modified_InWbkFullName", name_section:="clsLog,clsQ"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
    
        .TestNumber = "500-4"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes all values in one section which removes the section."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueRemove name_value:="ExportFileExtention" & _
                                         ",Last_Modified_AtDateTime" & _
                                         ",Last_Modified_InWbkFullName" & _
                                         ",Last_Modified_InWbkName" & _
                                         ",LastModExpFileFullNameOrigin" _
                           , name_section:="clsQ"
        
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
       
        .TestNumber = "500-5"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Remove all values in all sections - file is removed."
        sFile = PrivProfTests.PrivProfFile
        .ResultExpected = False
        .BoTP
        PrivProf.ValueRemove name_value:="ExportFileExtention" & _
                                         ",Last_Modified_AtDateTime" & _
                                         ",Last_Modified_InWbkFullName" & _
                                         ",Last_Modified_InWbkName" & _
                                         ",LastModExpFileFullNameOrigin" & _
                                         ",DoneNamesHskpng"
        
        .Result = FSo.FileExists(sFile) ' is False
        .EoTP
        ' ======================================================================
    End With

xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_600_Lifecycle()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = "Test_600_Lifecycle"
    
    On Error GoTo eh
   
    Prepare 0, False
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestNumber = "600-1"
        '~~ Begin with a non existing file.
        '~~ Note 1: Since there is no file without at least one section with at least one value,
        '~~         live starts with a value in a section in a yet not existing Private Profile file
        '~~ Note 2: When a header/footer is specified, the strings may preferrably delimited by a vbCrLf
        '~~         in order not to conflict withe any used character.
        If FSo.FileExists(PrivProfTests.PrivProfFileFullName) Then .FSo.DeleteFile PrivProfTests.PrivProfFileFullName
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .TestedProc = "Value-Let"
        .TestedType = "Property"
        .TestDscrpt = "Writes a new file with 1 section, 1 value and  a file header/footer."
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .BoTP
        PrivProf.Value(name_value:="Any-Value-Name" _
                    , name_section:="Any-Section-Name" _
                    , name_file:=PrivProfTests.PrivProfFileFullName _
                      ) = "Any-Value"
        .Result = PrivProfTests.PrivProfFile
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .EoTP
        ' ======================================================================
        
        .TestNumber = "600-2"
        '~~ Beginning a non existing file with writing header and/or footer raises an error
        '~~ (display ignored with Regression True)
        '~~ Note: This is consequent since ther is no Private Profile file without
        '~~       at least on section with one value
        If FSo.FileExists(PrivProfTests.PrivProfFileFullName) Then FSo.DeleteFile PrivProfTests.PrivProfFileFullName
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .TestedProc = "Let Comment/Footer"
        .TestedType = "Method"
        .TestDscrpt = "Writes a footer and a header into an empty file."
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .ResultExpected = True
        .BoTP
        mErH.Asserted AppErr(1) ' effective only when mErH.Regression = True
        PrivProf.FileFooter() = "File Footer Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Footer Line 2"
        PrivProf.FileHeader() = "File Comment Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Comment Line 2"
        mErH.Asserted ' reset to none
        .Result = Not FSo.FileExists(PrivProfTests.PrivProfFileFullName)
        .EoTP
        ' ======================================================================
        
        .TestNumber = "600-3"
        '~~ Removing the only value in the only section ends with no file
        '~~ Note: This is consequent since there is no Private Profile file without
        '~~       at least on section with one value
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes the only value in the only section"
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .ResultExpected = True
        .BoTP
        mErH.Asserted AppErr(1) ' effective only when mErH.Regression = True
        PrivProf.ValueRemove name_value:="Any-Value-Name" _
                           , name_section:="Any-Section-Name"
        mErH.Asserted ' reset to none
        .Result = Not FSo.FileExists(PrivProfTests.PrivProfFileFullName)
        .EoTP
        ' ======================================================================
        
    
    End With

xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_700_HskpngNames()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = " Test_700_HskpngNames"
    
    On Error GoTo eh
   
    BoP ErrSrc(PROC)
    
    Prepare 2   ' uses a ready for test file copied from a backup
    With TestAid
        .TestNumber = "700-1"
        .TestedProc = "HouskeepingNames"
        .TestedType = "Method"
        .TestDscrpt = "One value-name change in two sections."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.HskpngNames PrivProf.FileName, "clsLog:clsQ:Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
                
        .TestNumber = "700-2"
        Prepare 2   ' uses a ready for test file copied from a backup
        .TestedProc = "HouskeepingNames"
        .TestedType = "Method"
        .TestDscrpt = "Two value-names changes in all sections"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.HskpngNames PrivProf.FileName, "Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime" _
                                              , "LastModExpFileFullNameOrigin>Last_Modified_ExpFileFullNameOrigin"
        .Result = PrivProfTests.PrivProfFile
        .EoTP
        ' ======================================================================
                
    End With

xt: EoP ErrSrc(PROC)
    TestAid.TestFilesRemove
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

