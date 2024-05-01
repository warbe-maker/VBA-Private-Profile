Attribute VB_Name = "mPrivProfTest"
' ----------------------------------------------------------------
' Standard Module mPrivProvTest: Test of all services provided by
' ============================== the clsPrivProf class module.
' Usually each test is autonomous and preferrably uses no or only
' tested other Properties/Methods.
'
' Uses:
' - clsTestAid      Common services supporting test including
'                   regression testing.
' - clsPrivProfTest Services supporting tests of methods and
'                   properties of the class module clsPrivProf.
' - mTrc            Execution trace of tests.
'
' W. Rauschenberger, Berlin Apr 2024
' See also https://github.com/warbe-maker/VBA-Private-Profile.
' ----------------------------------------------------------------
Public PrivProf     As clsPrivProf
Public PrivProfTest As New clsPrivProfTest

Private cllExpctd   As Collection
Private FSo         As New FileSystemObject

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

Public Sub Prepare(Optional ByVal p_init_filename As Boolean = True)
    Const PROC = "Prepare"
    
    On Error GoTo eh
    If Tests Is Nothing Then Set Tests = New clsTestAid
    mTest.PrivateProfile_File ' uses Test.TestFolder
    Set PrivProf = Nothing
    Set PrivProf = New clsPrivProf
    If p_init_filename Then
        PrivProf.FileName = PrivProfTest.PrivProfFileFullName
    End If

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
    Set PrivProfTest = New clsPrivProfTest
    mTrc.FileFullName = Tests.TestFolder & "\Regression.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProf"
    mTrc.NewFile
    bModeRegression = True
    mErH.Regression = bModeRegression
    Tests.ModeRegression = bModeRegression
    Tests.TestFilesRemove "Result_" ' remove any files resulting from individual tests
    
    BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_001_Tests
    mPrivProfTest.Test_100_Property_FileName
    mPrivProfTest.Test_110_Method_Exists
    mPrivProfTest.Test_120_Property_Value
    mPrivProfTest.Test_200_Property_Comments
    mPrivProfTest.Test_300_Method_SectionNames
    mPrivProfTest.Test_400_Method_ValueNames
    mPrivProfTest.Test_410_Method_ValueNameRename
    mPrivProfTest.Test_600_Method_Remove
'    mPrivProfTest.Test_700_Method_SectionsCopy
    mPrivProfTest.Test_800_Lifecycle
    Tests.DsplySummary
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Set Test = Nothing
    Set Tests = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_001_Tests()
' ----------------------------------------------------------------------------
' Test of the means (clsTestAid) used by all tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_001_Tests"
    
    On Error GoTo eh
    Dim sFileResult     As String
    Dim sFileExpected   As String
    
    BoP ErrSrc(PROC)
    Prepare False
    With Tests
        .ModeRegression = mErH.Regression
        .TestNumber = "001-1"
        .TestedComp = "clsPrivProf"
        .TestDscrpt = "Initialize with a new PP-file"
        PrivProf.FileName = PrivProfTest.PrivProfFileFullName

        
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
    Tests.TestFilesRemove
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
    Prepare False
    With Tests
        .ModeRegression = mErH.Regression
        .TestNumber = "100-1"
        .TestedComp = "clsPrivProf"
        .TestedProc = "FileName_Let"
        .TestedType = "Property"
        .TestDscrpt = "Initialize PP-file"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.FileName = PrivProfTest.PrivProfFileFullName
        .Result = mTest.PrivProfFile
        .EoTP
        ' ======================================================================
        .TestNumber = "100-2"
        .TestedProc = "Let FileName"
        .TestedType = "Property"
        .TestDscrpt = "Specifying a file valid name"
        PrivProf.FileName = PrivProfTest.PrivProfFileFullName ' continue with specific test file
        .ResultExpected = PrivProfTest.PrivateProfile_File
        .BoTP
        .Result = PrivProf.FileName
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
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
       
    With Tests
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestNumber = "110-1"
        .TestDscrpt = "Section not exists"
        .ResultExpected = False
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, mTest.SectionName(7))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Section exists"
        .ResultExpected = True
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, mTest.SectionName(8))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name exists"
        .ResultExpected = True
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, mTest.SectionName(6), mTest.ValueName(6, 4))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "110-4"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestDscrpt = "Value-Name not exists"
        .ResultExpected = False
        .BoTP
        .Result = PrivProf.Exists(PrivProf.FileName, mTest.SectionName(6), mTest.ValueName(6, 3))
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
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
    
    With Tests
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
        .ResultExpected = mTest.ValueString(2, 4)
        .BoTP
        .Result = PrivProf.Value(name_value:=mTest.ValueName(2, 4) _
                              , name_section:=mTest.SectionName(2))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-3"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write changed value"
        .ResultExpected = "Changed value"
        .BoTP
        PrivProf.Value(name_value:=mTest.ValueName(4, 2) _
                    , name_section:=mTest.SectionName(4)) = "Changed value"
        .Result = PrivProf.Value(name_value:=mTest.ValueName(4, 2) _
                              , name_section:=mTest.SectionName(4))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-4"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in existing section"
        .ResultExpected = "New value, existing section"
        .BoTP
        PrivProf.Value(mTest.ValueName(2, 17) _
                    , mTest.SectionName(2)) = "New value, existing section"
        .Result = PrivProf.Value(name_value:=mTest.ValueName(2, 17) _
                              , name_section:=mTest.SectionName(2))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-5"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write new value in new section"
        .ResultExpected = "New value, new section"
        .BoTP
        PrivProf.Value(name_value:=mTest.ValueName(11, 1) _
                    , name_section:=mTest.SectionName(11)) = "New value, new section"
        .Result = PrivProf.Value(name_value:=mTest.ValueName(11, 1) _
                              , name_section:=mTest.SectionName(11))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-6"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Change value and value plus section comments"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Value(name_value:=mTest.ValueName(11, 1) _
                     , name_section:=mTest.SectionName(11) _
                     , comments_value:="Value comment" _
                     , comments_section:="Section comment (by the way)" _
                      ) = "Changed new value, new section"
        .Result = mTest.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "120-7"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Changes value and value comments but not the section comments"
        .BoTP
        PrivProf.Value(name_value:=mTest.ValueName(11, 1) _
                     , name_section:=mTest.SectionName(11) _
                     , comments_value:="Value comment changed" _
                      ) = "Changed again new value, new section"
        .Result = mTest.PrivProfFile
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_200_Property_Comments()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Property_Comments"

    On Error GoTo eh
    Dim sHeader     As String
    Dim sResult     As String
    Dim sValue      As String
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With Tests
        .TestNumber = "200-1"
        .TestedProc = "Let Comment"
        .TestedType = "Property"
        .TestDscrpt = "Write file comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        '~~ Note: For the missing file name the property FileName is used
        '~~ and the missing section- and value-name indicate a file comment
        PrivProf.Comments() = "File Comment Line 1 (the comments delimiter below is adjusted to the longest comment)" & vbCrLf & _
                              "File Comment Line 2"
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-3"
        .TestedProc = "Get Comment"
         .TestedType = "Property"
        .TestDscrpt = "File comment read"
        Set cllResultExpectd = New Collection
        cllResultExpectd.Add "; File Comment Line 1 (the comments delimiter below is adjusted to the longest comment)"
        cllResultExpectd.Add "; File Comment Line 2"
        cllResultExpectd.Add "; ====================================================================================="
        .ResultExpected = cllResultExpectd
        .BoTP
        .Result = PrivProf.Comments()
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-4"
        .TestedProc = "Let Comment"
        .TestedType = "Property"
        .TestDscrpt = "Write section comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Comments(, mTest.SectionName(6)) = "Comment Section 06 Line 1" & vbCrLf & _
                                                    "Comment Section 06 Line 2"
        .Result = mTest.PrivProfFile
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-5"
        .TestedProc = "Get Comment"
        .TestedType = "Property"
        .TestDscrpt = "Read section comment"
        .ResultExpected = ",; Comment Section 06 Line 1,; Comment Section 06 Line 2"
        .BoTP
        .Result = PrivProf.Comments(, mTest.SectionName(6))
        .EoTP
        ' =====================================================================
        
        .TestNumber = "200-6"
        .TestedProc = "Let Comment"
        .TestedType = "Property"
        .TestDscrpt = "Write value comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Comments(, mTest.SectionName(6), mTest.ValueName(6, 2)) = "Comment Section 06 Value 02 Line 1" & vbCrLf & _
                                                                        "Comment Section 06 Value 02 Line 2"
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-7"
        .TestedProc = "Get Comment"
        .TestedType = "Property"
        .TestDscrpt = "Read value comment"
        .ResultExpected = "; Comment Section 06 Value 02 Line 1,; Comment Section 06 Value 02 Line 2"
        .BoTP
        .Result = PrivProf.Comments(, mTest.SectionName(6), mTest.ValueName(6, 2))
        .EoTP
        ' ======================================================================
        
        .TestNumber = "200-8"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        .TestDscrpt = "Write a new value including a value comment"
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Value(mTest.ValueName(12, 1), mTest.SectionName(12), , "The new value's comment line 1,The new value's comment line 2!") = "New value"
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
    
        .TestNumber = "200-9"
        .TestedProc = "Get Value"
        .TestedType = "Property"
        .TestDscrpt = "Read a value-comment along with a value"
        .ResultExpected = "; The new value's comment line 1|&|; The new value's comment line 2!"
        .BoTP
        sValue = PrivProf.Value(mTest.ValueName(12, 1), mTest.SectionName(12), , sComment)
        .Result = sComment
        .EoTP
        ' ======================================================================
    End With

xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
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
       
    With Tests
        .TestNumber = "300-1"
        .TestedProc = "SectionNames"
        .TestedType = "Function"
        .TestDscrpt = "Get all section names in a Dictionary"
        .ResultExpected = 5
        .BoTP
        Set dct = PrivProf.SectionNames()
        .Result = dct.Count
        .EoTP
    End With
    
xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
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
       
    With Tests
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
        .Result = PrivProf.ValueNames(, mTest.SectionName(6)).Count
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
    Dim dct         As Dictionary
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With Tests
        .TestNumber = "410-1"
        .TestedProc = "ValueNameRename"
        .TestedType = "Method"
        .TestDscrpt = "Rename a value name in each section."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueNameRename mTest.ValueName(2, 2), "Renamed_" & mTest.ValueName(2, 2)
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    Set dct = Nothing
    Tests.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_600_Method_Remove()
' ----------------------------------------------------------------------------
' The test relies on: - Comment value
' ----------------------------------------------------------------------------
    Const PROC = "Test_600_Method_Remove"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
    
    With Tests
        PrivProf.Comments(, mTest.SectionName(6), mTest.ValueName(6, 4)) = "Comment value 06-04"
        .TestNumber = "600-1"
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .TestDscrpt = "Remove a value from a section including its comments."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.ValueRemove mTest.ValueName(6, 4), mTest.SectionName(6)
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
        
        PrivProf.Comments(, mTest.SectionName(6)) = "Comment section 06"
        .TestNumber = "600-2"
        .TestedProc = "SectionRemove"
        .TestedType = "Method"
        .TestDscrpt = "Removes a section including its comments."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.SectionRemove mTest.SectionName(6)
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
    
    End With

xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub Test_700_Method_SectionsCopy()
'' ----------------------------------------------------------------------------
'' This test relies on: - method SectionNames (Test_300_Method_SectionNames),
''                      - method SectionRemove (Test_600_Method_Remove)
'' The test implicitely tests the property Sections Get/Let.
'' ----------------------------------------------------------------------------
'    Const PROC = "Test_700_Method_SectionsCopy"
'
'    On Error GoTo eh
'    Dim sSourceFile     As String
'    Dim sTargetFile     As String
'
'    BoP ErrSrc(PROC)
'    Prepare ' Test preparation
'
'    With Tests
'        sTargetFile = .TestFolder & "\CopyTarget.dat"
'        sSourceFile = mTest.PrivateProfile_File
'        If .FSo.FileExists(sTargetFile) Then .FSo.DeleteFile sTargetFile
'        .TestNumber = "700-1"
'        .TestedProc = "SectionsCopy"
'        .TestedType = "Method"
'        .TestDscrpt = "Copies two sections from a soure to a traget Private Profile file."
'        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
'        .BoTP
'        PrivProf.SectionsCopy name_file_source:=sSourceFile _
'                           , name_file_target:=sTargetFile _
'                           , name_sections:=mTest.SectionName(6) & "," & mTest.SectionName(2)
'        .Result = .FSo.GetFile(sTargetFile)
'        .EoTP
'        ' ======================================================================
'
'        .TestNumber = "700-2"
'        .TestedProc = "SectionsCopy"
'        .TestedType = "Method"
'        .TestDscrpt = "Copies an additional sections from a soure to a traget Private Profile file."
'        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
'        .BoTP
'        PrivProf.SectionsCopy name_file_source:=sSourceFile _
'                           , name_file_target:=sTargetFile _
'                           , name_sections:=mTest.SectionName(4)
'        .Result = .FSo.GetFile(sTargetFile)
'        .EoTP
'        ' ======================================================================
'    End With
'
'xt: EoP ErrSrc(PROC)
'    Tests.TestFilesRemove
'    Exit Sub
'
'eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Public Sub Test_800_Lifecycle()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = "Test_800_Lifecycle"
    
    On Error GoTo eh
   
    Prepare
    BoP ErrSrc(PROC)
    
    With Tests
        Prepare ' Test preparation
        .FSo.DeleteFile PrivProfTest.PrivProfFileFullName
        .TestNumber = "800-1"
        .TestedProc = "Let Comment/Footer"
        .TestedType = "Method"
        .TestDscrpt = "Writes a file comment and footer into an empty file."
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .BoTP
        PrivProf.Comments() = "File Comment Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                              "File Comment Line 2"
        PrivProf.Footer() = "File Footer Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                            "File Footer Line 2"
        .Result = FSo.GetFile(PrivProf.FileName)
        .EoTP
        ' ======================================================================
        
        .TestNumber = "800-2"
        .TestedProc = "Let Comment/Footer"
        .TestedType = "Method"
        .TestDscrpt = "Writes a file comment and footer into an empty file."
        .BoTP
        PrivProf.Value("AnyValueName", "AnySectionName") = "AnyValue"
        .Result = FSo.GetFile(PrivProf.FileName)
        .ResultExpected = .ExpectedTestResultFile(.TestNumber)
        .EoTP
        ' ======================================================================
        
    End With

xt: EoP ErrSrc(PROC)
    Tests.TestFilesRemove
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

