VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestAid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' Class Module clsTestAid: Common test supporting Properties and Methods,
' ======================== including Regression-Test support. The module
' supports a "regression test mode" (Property "ModeRegression") which
' suppresses a Debug.Assert interruption of a test sequence in case of a
' "failed" test result, but just collects it and provide a DsplySummary
' method for reporting all in a "Regression.Summary.log" file.
' In non "regression test mode", when the result and the result expected
' difffers WinMerge is used to display the difference (an error is
' displayed when not installed).
'
' Public properties:
' ------------------
' TestDscrpt           Let A string describing the purpose of a test.
' TestNumber         Let A string identifying a specific test within the test
'                    procedure
' TestedProc         The name of the tested procedure (Sub, Function,
'                    Property, Method).
' Result         Let Variant expression a the result of a test, maybe a file
'                    object, any other object or any other data type
' ResultExpected Let The expected result of a test, a file object, any other
'                    object or any other data type
' Public methods:
' ---------------
' BoC/EoC           Effective only when the execution trace module (mTrc or
'                   clsTrc) is installed, indicates the begin of a code
'                   sequence.
' BoP/EoP           Effective only when the execution trace module (mTrc or
'                   clsTrc) is installed, indicates the begin/end of a
'                   procedure.
' DsplySummary      Displays a RegressionTest.log file provided ModeRegression = True.
' ResultAsExpected  Returns True or False depending result provided once the Result
'                   and the ResultExpected had been assigned.
'
' Requires:
' - Refernce to "Microsoft Scripting Runtime"
' - mTrc/clsTrc     Only when an execution trace is desired
'
' W. Rauschenberger, Berlin Feb 2024
' ----------------------------------------------------------------------------
Public Fso                     As New FileSystemObject

Private Const RESULT_PASSED = " Passed "
Private Const RESULT_FAILED = " F a i l e d ! "
Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&


Private bModeRegression         As Boolean
Private dctResults              As New Dictionary
Private sTestDscrpt             As String
Private sTestNumber             As String
Private sResult                 As String ' " P a s s e d !" or " F a i l e d !"
Private sRegressionSummaryLog   As String
Private sTestedComp             As String
Private sTestedProc             As String
Private sTestResult             As String
Private vResult                 As Variant
Private vResultExpected         As Variant
Private sTestFolder             As String
Private sTestedType             As String
Private lMaxLenTestProc         As Long

Private Sub Class_Initialize()
    
    sTestFolder = ThisWorkbook.Path & "\Test"
    If Not Fso.FolderExists(sTestFolder) _
    Then Fso.CreateFolder ThisWorkbook.Path & "\Test"
    sRegressionSummaryLog = ThisWorkbook.Path & "\Test\Regression.Summary.log"
    If Fso.FileExists(sRegressionSummaryLog) Then
        Fso.DeleteFile sRegressionSummaryLog
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

Public Property Let ModeRegression(ByVal b As Boolean)
    
    bModeRegression = b
#If mErh Then
    mErh.Regression = b
#End If

End Property

Public Property Get Result() As Variant
    If IsObject(vResult) Then Set Result = vResult Else Result = vResult
End Property

Public Property Let Result(ByVal r_var As Variant)

    If IsObject(r_var) Then Set vResult = r_var Else vResult = r_var
    '~~ When also the expected test result is assigned the test result is provided
    If Not IsEmpty(vResultExpected) Then
        ProvideFinalTestResult
    End If
    
End Property

Public Property Get ResultExpected() As Variant:    ResultExpected = vResultExpected:   End Property

Public Property Let ResultExpected(ByVal v As Variant)

    If IsObject(v) Then Set vResultExpected = v Else vResultExpected = v
    '~~ When also vResult has been assigned, the test result is provided
    If Not IsEmpty(vResult) Then
        ProvideFinalTestResult
    End If
    
End Property

Public Property Let TestDscrpt(ByVal s As String):  sTestDscrpt = s:                    End Property

Public Property Let TestedComp(ByVal s As String)
    sTestedComp = s
End Property

Public Property Let TestedProc(ByVal s As String)
    sTestedProc = s
    lMaxLenTestProc = Max(lMaxLenTestProc, Len(s))
End Property

Public Property Get TestedType() As String:         TestedType = sTestedType:           End Property

Public Property Let TestedType(ByVal s As String):  sTestedType = s:                    End Property

Public Property Get TestFolder() As String:         TestFolder = sTestFolder:           End Property

Public Property Let TestFolder(ByVal s As String)
    sTestFolder = s
    If Not Fso.FolderExists(s) Then Fso.CreateFolder s
End Property

Public Property Get TestId() As String
    Dim s As String
    Select Case Left(sTestedType, 1)
        Case "P":   s = "Property"
        Case "F":   s = "Function"
        Case "S":   s = "Sub     "
        Case "M":   s = "Method  "
    End Select
    TestId = sTestNumber & " " & s & " " & sTestedComp & "." & sTestedProc & ": " & sTestDscrpt

End Property

Public Property Get TestNumber() As String:         TestNumber = sTestNumber:           End Property

Public Property Let TestNumber(ByVal s As String)
    If sTestNumber <> s Then
        ResetTest ' new number indicates new test
        sTestNumber = s
    End If
    
End Property

Public Property Get TestResult() As String: TestResult = sTestResult:   End Property

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

Private Function AppIsInstalled(ByVal sApp As String) As Boolean
    
    Dim i As Long: i = 1
    Do Until VBA.Left$(Environ$(i), 5) = "Path=":   i = i + 1:  Loop
    AppIsInstalled = InStr(Environ$(i), sApp) <> 0

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

Public Function AsCollection(ParamArray c_arr() As Variant) As Collection
    Dim cll As New Collection
    Dim i   As Long
    
    If ArrayIsAllocated(c_arr) Then
        For i = LBound(c_arr) To UBound(c_arr)
            cll.Add c_arr(i)
        Next i
    End If
    Set AsCollection = cll
    Set cll = Nothing
    
End Function

Public Function AsFile(ParamArray c_arr() As Variant) As File
    Const PROC = "AsFile"
    
    On Error GoTo eh
    Dim cll As Collection
    Dim i   As Long
    Dim v   As Variant
    Dim arr As Variant
    Dim s   As String
    
    If ArrayIsAllocated(c_arr) Then
        If TypeName(c_arr(0)) = "Collection" Then
            Set cll = c_arr(0)
            For Each v In cll
                ArrayAdd arr, v
            Next v
        Else
            For i = LBound(c_arr) To UBound(c_arr)
                ArrayAdd arr, c_arr(i)
            Next i
        End If
    End If
    s = TempFile
    If ArrayIsAllocated(arr) Then
        FileFromString s, Join(arr, vbCrLf)
    Else
        Fso.CreateTextFile s
    End If
    Set AsFile = Fso.GetFile(s)
    
xt: Exit Function
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub AssertedErrors(ParamArray a_app_errs() As Variant)
' ----------------------------------------------------------------------------
' When mErH is not installed (Cond. Comp. Arg. `mErh=1`) this is ignored.
' ----------------------------------------------------------------------------
#If mErh Then
    Const PROC = "AssertedErrors"

    On Error GoTo xt
    If ArrayIsAllocated(a_app_errs) Then
        Select Case UBound(a_app_errs)
            Case 0: mErh.Asserted a_app_errs(0)
            Case 1: mErh.Asserted a_app_errs(0), a_app_errs(1)
            Case 2: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2)
            Case 3: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3)
            Case 4: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4)
            Case 5: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5)
            Case 6: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6)
            Case 7: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6), a_app_errs(7)
            Case 8: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6), a_app_errs(7), a_app_errs(8)
            Case 9: mErh.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(6), a_app_errs(6), a_app_errs(7), a_app_errs(8), a_app_errs(9)
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "Max number of assertable errors (10) exceeded!"
        End Select
    End If

xt: Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
#End If
End Sub


Public Sub BoTP()
' ------------------------------------------------------------------------------
' Begin of Testesd Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErh Then          ' serves the mTrc/clsTrc when installed and active
    mErh.BoP Me.TestId
#ElseIf XcTrc_clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP Me.TestId
#ElseIf XcTrc_mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Function CollAsString(ByVal c As Collection) As String
    Dim v As Variant
    For Each v In c
        CollAsString = CollAsString & v
    Next v
End Function

Public Sub DsplySummary()
    Const PROC = "DsplySummary"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lMax        As Long
    Dim lMaxDesc    As Long
    Dim s           As String
    Dim arr         As Variant
    Dim cll         As Collection
    Dim sType       As String
    
    If Not bModeRegression Then Exit Sub ' is bypassed except ModeRegression = True
    
    s = "Summary Regression-Test " & Format(Now(), "YYYY.MM.DD hh.mm.ss")
    If Len(FileAsString(sRegressionSummaryLog)) > 1 Then
        arr = Split(FileAsString(sRegressionSummaryLog), vbCrLf)
        FileString(sRegressionSummaryLog, True) = String(Len(arr(UBound(arr) - 1)), "=")
    End If
    FileString(sRegressionSummaryLog, True) = s
    FileString(sRegressionSummaryLog, True) = String(Len(s), "-")
    For Each v In dctResults
        Set cll = dctResults(v)
'       1 sTestNumber
'       2 sTestedType
'       3 sTestedComp
'       4 sTestedProc
'       5 sTestDscrpt
'       6 sTestResult

        lMax = Max(lMax, Len(cll(3) & "." & cll(4)))
        lMaxDesc = Max(lMaxDesc, Len(cll(5)))
    Next v
    
    For Each v In dctResults
        Set cll = dctResults(v)
        
        Select Case Left(cll(2), 1)
            Case "P":   sType = "Property"
            Case "F":   sType = "Function"
            Case "S":   sType = "Sub     "
            Case "M":   sType = "Method  "
        End Select
        
        s = cll(3) & "." & cll(4)
        
        FileString(sRegressionSummaryLog, True) = cll(1) & " " & _
                                        sType & " " & _
                                        s & " " & _
                                        String(lMax - Len(s), ".") & _
                                        " " & _
                                        cll(5) & _
                                        String(lMaxDesc - Len(cll(5)) + 5, ".") & _
                                        ": " & cll(6)
    Next v
    ShellRun sRegressionSummaryLog
    
xt: Exit Sub
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub EoTP()
' ------------------------------------------------------------------------------
' End of Tested Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErh = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErh.EoP Me.TestId, Me.TestResult
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP Me.TestId, Me.TestResult
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP Me.TestId, Me.TestResult
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
    ErrSrc = "clsTestAid." & sProc
End Function

Private Function FileAsString(Optional ByVal f_file As Variant) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file) - may be a full name string or a file
' object - as a single string.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    On Error Resume Next
    Dim fl  As File
    
    If Not TypeName(f_file) = "File" _
    And Not TypeName(f_file) = "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a file object nor a file's full name string!"
    
    If TypeName(f_file) = "File" Then
        Set fl = f_file
        Open fl.Path For Input As #1
    ElseIf TypeName(f_file) = "String" Then
        If Fso.FileExists(f_file) _
        Then Open f_file For Input As #1 _
        Else Err.Raise AppErr(2), ErrSrc(PROC), "The provided string is not an exoisting file's name!"
    End If
    
    FileAsString = Input$(LOF(1), 1)
    Close #1
        
xt: Exit Function
    
eh: Select Case mErh.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileDiffDsply(ByVal f_fl_left As Variant, _
                               ByVal f_fl_1eft_title As String, _
                               ByVal f_fl_right As Variant, _
                               ByVal f_fl_right_title As String) As Boolean
' ----------------------------------------------------------------------------
' Displays the differences between export file 1 and 2 by means of WinMerge!
' Note: CompMan ignores any differences caused by empty code lines or case.
'       When a difference is displayed it is thus not because of this kind of
'       differneces but of others. Unfortunately it depends on the installed
'       WinMerge's set option wether or not these kind of differences are
'       displayed.
' ----------------------------------------------------------------------------
    Const PROC              As String = "ExpFilesDiffDisplay"
    
    On Error GoTo eh
    Dim waitOnReturn        As Boolean: waitOnReturn = True
    Dim windowStyle         As Integer: windowStyle = 1
    Dim sCommand            As String
    Dim wshShell            As Object
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="WinMerge is used to display the difference between the result " & _
                                "and the expected result when both are files - but WinMerge is " & _
                                "not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not Fso.FileExists(f_fl_left) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & f_fl_left & """ does not exist!"
    
    If Not Fso.FileExists(f_fl_right) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & f_fl_right & """ does not exist!"
            
    '~~ Prepare command line
    sCommand = "WinMergeU " & _
               """" & f_fl_left & """" & " " & _
               """" & f_fl_right & """" & _
               " /e " & _
               " /dl " & DQUOTE & f_fl_1eft_title & DQUOTE & _
               " /dr " & DQUOTE & f_fl_right_title & DQUOTE & """"

    '~~ Execute command line
    Set wshShell = CreateObject("WScript.Shell")
    FileDiffDsply = wshShell.Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileFromCollection(ByVal f_coll As Collection) As File
    Dim arr As Variant
    Dim v   As Variant
    Dim s   As String
    
    For Each v In f_coll
        ArrayAdd arr, v
    Next v
    s = Fso.GetTempName
    FileFromString s, Join(arr, vbCrLf)
    Set FileFromCollection = Fso.GetFile(s)
    
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

Private Sub ProvideFinalTestResult()
' ----------------------------------------------------------------------------
' Returns " P a s s e d !" when the provided Result (vResult) is equal the provided
' ResultExpected (vResultExpected), whereby both may be a File object, any
' other object or any other date type. This function should reset the Varaints
' vResult and vResultExpected to Empty
' ----------------------------------------------------------------------------
    Const PROC = "ProvideFinalTestResult"
    
    On Error GoTo eh
    Dim flResultExpected    As File
    Dim flResult            As File
    Dim cll                 As Collection
    
    If IsObject(vResult) And Not IsObject(vResultExpected) _
    Or Not IsObject(vResult) And IsObject(vResultExpected) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Result and ResultExpected are not compatibel! Either of the two is an object but the other isn't."

    If sTestedProc = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The name of the tested procedure (Sub, Function, or Property) is still unknown!" & "||" & _
                                            "Once a Result and an expected result has been provided, the final test result is provided. " & _
                                            "However, this requires the name of the test procedure in charge."
    
    sTestResult = RESULT_FAILED
    Select Case True
        Case TypeName(vResult) = "File" And TypeName(vResultExpected) = "File"
            Set flResultExpected = vResultExpected
            Set flResult = vResult
            If StrComp(FileAsString(flResult.Path), FileAsString(flResultExpected), vbBinaryCompare) = 0 _
            Then sTestResult = RESULT_PASSED ' Files are equal
    
        Case TypeName(vResult) = "Collection" And TypeName(vResultExpected) = "Collection"
            If CollComp(vResult, vResultExpected) = 0 _
            Then sTestResult = RESULT_PASSED ' Files are equal

        Case IsObject(vResult) And IsObject(vResultExpected)
            If vResult Is vResultExpected _
            Then sTestResult = RESULT_PASSED
            
        Case Else
            If vResult = vResultExpected _
            Then sTestResult = RESULT_PASSED ' values are equal
    End Select
    
    If bModeRegression Then
        '~~ Save test result and return TRUE
        Set cll = New Collection
        With cll
            .Add sTestNumber
            .Add sTestedType
            .Add sTestedComp
            .Add sTestedProc
            .Add sTestDscrpt
            .Add sTestResult
        End With
        dctResults.Add TestId, cll
        Set cll = Nothing
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function CollComp(ByVal c_1 As Collection, _
                          ByVal c_2 As Collection) As Long
    Dim i
    CollComp = 0
    If c_1.Count <> c_2.Count Then
        CollComp = -1
        Exit Function
    End If
        
    For i = 1 To c_1.Count
        Select Case True
            Case IsObject(c_1(i)) And IsObject(c_2(i))
                If Not c_1(i) Is c_2(i) Then
                    CollComp = -0
                    Exit Function
                End If
            Case Else
                If c_1(i) <> c_2(i) Then
                    CollComp = -1
                    Exit Function
                End If
        End Select
    Next i
    
End Function

Private Sub ResetTest()

    vResult = Empty
    vResultExpected = Empty
    
End Sub

Public Function ResultAsExpected() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided test result is " P a s s e d !" and FALSE
' otherwise.
' Note: Obtaining the test result as expected True or False indicates the end
'       of the test and thus resets it which prepares it for the next one.
' ----------------------------------------------------------------------------
    Const PROC = "ResultAsExpected"
    
    On Error GoTo eh
    Dim flLeft  As File
    Dim flRight As File
    
    If bModeRegression Then
        ResultAsExpected = True
        GoTo xt
    End If
    ResultAsExpected = sTestResult = RESULT_PASSED
    If sTestResult = RESULT_FAILED Then
        Select Case True
            Case TypeName(vResult) = "File" And TypeName(vResultExpected) = "File"
                Set flLeft = vResult
                Set flRight = vResultExpected
                FileDiffDsply flLeft.Path, "R e s u l t  of test " & Me.TestId, flRight.Path, "R e s u l t   e x p e c t e d !"
            Case TypeName(vResult) = "Collection" And TypeName(vResultExpected) = "Collection"
                Set flLeft = Me.AsFile(vResult)
                Set flRight = Me.AsFile(vResultExpected)
                FileDiffDsply flLeft.Path, "R e s u l t  of test " & Me.TestId, flRight.Path, "R e s u l t   e x p e c t e d !"
                With Fso
                    .DeleteFile flLeft.Path
                    .DeleteFile flRight.Path
                End With
        End Select
    End If
    
    '~~ Once the result as expected (True or False) has been obtained, both are re-set to empty
xt: ResetTest
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ShellRun(ByVal sr_string As String, _
                 Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
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
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".tmp") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted the path is the CurDir path, else in at the provided one.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(Fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = Environ("TMP")
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    
End Function

Private Function WinMergeIsInstalled() As Boolean
    WinMergeIsInstalled = AppIsInstalled("WinMerge")
End Function

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

Private Sub Class_Terminate()
    Set dctResults = Nothing
End Sub
