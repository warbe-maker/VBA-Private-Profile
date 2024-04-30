Public Function VarItems(ByVal v_items As Variant, _
                         ByVal v_items_as As enVarItems, _
                Optional ByVal v_file_name As String = vbNullString, _
                Optional ByVal v_file_append As Boolean = False, _
                Optional ByVal v_items_empty_excluded As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Universal items conversion service. Variant items (v_items) - which
' may be an Array of items, a Collection of items, a Dictionary of keys, a
' TextStream file, or a string with items delimited by: vbCrLf, vbLf, ||, |,
' or a , (comma)as Array (v_arr) - are provided as Array, Collection,
' Dictionary (with the item as the key), TextStream file, or as a String
' with the items delimited by a vbCrLf.
' Note: When the provision "as Array" (v_items_as) is requested and an item is
'       an object which does not have a Name property, an error is raised and
'       the function is terminated.
' ----------------------------------------------------------------------------
    Const PROC = "VarItems"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim v       As Variant
    Dim arr     As Variant
    Dim s       As String
    Dim sDelim  As String
    
    '~~ Save all items into a Collection
    Select Case TypeName(v_items)
        Case "String"
            If v_items = vbNullString Then GoTo xt
            Select Case True
                Case InStr(v_items, vbCrLf) <> 0: sSplit = vbCrLf
                Case InStr(v_items, vbLf) <> 0:   sSplit = vbLf
                Case InStr(v_items, "||") <> 0:   sSplit = "||"
                Case InStr(v_items, "|") <> 0:    sSplit = "|"
                Case InStr(v_items, ",") <> 0:    sSplit = ","
            End Select
            For Each v In Split(v_items, sSplit)
                cll.Add VBA.Trim$(v)
            Next v
        Case "Collection", "Dictionary", "Array"
            For Each v In v_items
                cll.Add VBA.Trim$(v)
            Next v
        Case "File"
            s = FileAsString(v_items)
            For Each v In Split(s, vbCrLf)
                cll.Add v
            Next v
        Case Else:      Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a String, an Array, a Collecton, a Dictionary, nor a TextFile!"
    End Select
            
    '~~ Prepare the output items in the requested form
xt: Select Case v_items_as
        Case enAsCollection
            Set VarItems = cll
        Case enAsArray
            For Each v In cll
                If IsObject(v) Then
                    On Error Resume Next
                    s = v.Name
                    If Err.Number <> 0 _
                    Then Err.Raise AppErr(1), ErrSrc(PROC), "At least one of the variant items is an object which cannot be provided in an Array since it does not provide a Name property!"
                Else
                    s = v
                End If
                ArrayAdd arr, v
            Next v
            VarItems = arr
        Case enAsDictionary
            For Each v In cll
                dct.Add v, vbNullString
            Next v
            Set VarItems = KeySort(dct)
        Case enAsString
            For Each v In cll
                s = s & sDelim & v
                sDelim = vbCrLf
            Next v
            VarItems = s
        Case enAsTextFile
            For Each v In cll
                If Not v_items_empty_excluded _
                Then s = s & sDelim & v _
                Else If VBA.Trim$(v) <> vbNullString Then s = s & sDelim & v
                If s <> vbNullString Then sDelim = vbCrLf
            Next v
            If v_file_append _
            Then Open v_file_name For Append As #1 _
            Else Open v_file_name For Output As #1
            Print #1, s
            Close #1
        Case Else
            Err.Raise AppErr(2), ErrSrc(PROC), _
            "The variant items are not requested to be returned neither as String, Array, Collection, Directory, nor as TextStream file!"
    End Select
    
    Set cll = Nothing
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function