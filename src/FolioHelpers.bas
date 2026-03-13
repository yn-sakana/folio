Attribute VB_Name = "FolioHelpers"
Option Explicit

' ============================================================================
' JSON Parser
' ============================================================================

Public Function ParseJson(ByVal text As String) As Object
    Dim p As Long: p = 1
    SkipWS text, p
    If p > Len(text) Then Set ParseJson = NewDict(): Exit Function
    Dim ch As String: ch = Mid$(text, p, 1)
    If ch = "{" Then
        Set ParseJson = ParseObj(text, p)
    ElseIf ch = "[" Then
        Set ParseJson = ParseArr(text, p)
    Else
        Set ParseJson = NewDict()
    End If
End Function

Private Function ParseObj(ByRef s As String, ByRef p As Long) As Object
    Dim d As Object: Set d = NewDict()
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "}" Then p = p + 1: Set ParseObj = d: Exit Function
    End If
    Do
        SkipWS s, p
        Dim k As String: k = ParseStr(s, p)
        SkipWS s, p
        If p <= Len(s) Then p = p + 1 ' skip :
        Dim v As Variant: ParseVal s, p, v
        d.Add k, v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "}" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseObj = d
End Function

Private Function ParseArr(ByRef s As String, ByRef p As Long) As Object
    Dim c As New Collection
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "]" Then p = p + 1: Set ParseArr = c: Exit Function
    End If
    Do
        SkipWS s, p
        Dim v As Variant: ParseVal s, p, v
        c.Add v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "]" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseArr = c
End Function

Private Sub ParseVal(ByRef s As String, ByRef p As Long, ByRef result As Variant)
    SkipWS s, p
    If p > Len(s) Then result = Null: Exit Sub
    Dim ch As String: ch = Mid$(s, p, 1)
    Select Case ch
        Case "{":  Set result = ParseObj(s, p)
        Case "[":  Set result = ParseArr(s, p)
        Case """": result = ParseStr(s, p)
        Case "t":  result = True: p = p + 4
        Case "f":  result = False: p = p + 5
        Case "n":  result = Null: p = p + 4
        Case Else: result = ParseNum(s, p)
    End Select
End Sub

Private Function ParseStr(ByRef s As String, ByRef p As Long) As String
    p = p + 1
    Dim buf As String, start As Long: start = p
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch = """" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            ParseStr = buf: Exit Function
        ElseIf ch = "\" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            If p <= Len(s) Then
                Dim esc As String: esc = Mid$(s, p, 1)
                Select Case esc
                    Case """", "\", "/": buf = buf & esc
                    Case "n": buf = buf & vbLf
                    Case "r": buf = buf & vbCr
                    Case "t": buf = buf & vbTab
                    Case "u"
                        If p + 4 <= Len(s) Then
                            On Error Resume Next
                            buf = buf & ChrW$(CLng("&H" & Mid$(s, p + 1, 4)))
                            On Error GoTo 0
                            p = p + 4
                        End If
                End Select
                p = p + 1: start = p
            End If
        Else
            p = p + 1
        End If
    Loop
    ParseStr = buf & Mid$(s, start, p - start)
End Function

Private Function ParseNum(ByRef s As String, ByRef p As Long) As Double
    Dim start As Long: start = p
    If p <= Len(s) Then If Mid$(s, p, 1) = "-" Then p = p + 1
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch Like "[0-9.eE+-]" Then p = p + 1 Else Exit Do
    Loop
    On Error Resume Next
    ParseNum = CDbl(Mid$(s, start, p - start))
    On Error GoTo 0
End Function

Private Sub SkipWS(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbLf, vbCr: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

' ============================================================================
' JSON Serializer
' ============================================================================

Public Function ToJson(ByVal v As Variant, Optional ind As Long = -1) As String
    If IsObject(v) Then
        If v Is Nothing Then ToJson = "null": Exit Function
        Dim obj As Object: Set obj = v
        If TypeName(obj) = "Dictionary" Then ToJson = DictToJson(obj, ind): Exit Function
        If TypeName(obj) = "Collection" Then ToJson = CollToJson(obj, ind): Exit Function
        ToJson = "null"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ToJson = "null"
    ElseIf VarType(v) = vbString Then
        ToJson = """" & JsonEscape(CStr(v)) & """"
    ElseIf VarType(v) = vbBoolean Then
        ToJson = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        ToJson = CStr(v)
    Else
        ToJson = """" & JsonEscape(CStr(v)) & """"
    End If
End Function

Private Function DictToJson(d As Object, ind As Long) As String
    If d.Count = 0 Then DictToJson = "{}": Exit Function
    Dim keys() As Variant: keys = d.keys
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(d.Count - 1)
    Dim i As Long
    For i = 0 To d.Count - 1
        Dim val As Variant
        If IsObject(d(keys(i))) Then Set val = d(keys(i)) Else val = d(keys(i))
        parts(i) = sp & """" & JsonEscape(CStr(keys(i))) & """:" & IIf(ind >= 0, " ", "") & ToJson(val, ind2)
    Next i
    DictToJson = "{" & nl & Join(parts, "," & nl) & nl & csp & "}"
End Function

Private Function CollToJson(c As Object, ind As Long) As String
    If c.Count = 0 Then CollToJson = "[]": Exit Function
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        Dim val As Variant
        If IsObject(c(i)) Then Set val = c(i) Else val = c(i)
        parts(i - 1) = sp & ToJson(val, ind2)
    Next i
    CollToJson = "[" & nl & Join(parts, "," & nl) & nl & csp & "]"
End Function

Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

' ============================================================================
' Dictionary Helpers
' ============================================================================

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Public Function DictStr(d As Object, key As String, Optional def As String = "") As String
    DictStr = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictStr = CStr(d(key))
End Function

Public Function DictObj(d As Object, key As String) As Object
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If Not IsObject(d(key)) Then Exit Function
    Set DictObj = d(key)
End Function

Public Function DictBool(d As Object, key As String, Optional def As Boolean = False) As Boolean
    DictBool = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictBool = CBool(d(key))
End Function

Public Function DictLng(d As Object, key As String, Optional def As Long = 0) As Long
    DictLng = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictLng = CLng(d(key))
End Function

Public Sub DictPut(d As Object, key As String, val As Variant)
    If d.Exists(key) Then d.Remove key
    d.Add key, val
End Sub

' ============================================================================
' File System
' ============================================================================

Public Function ReadTextFile(path As String) As String
    On Error GoTo ErrOut
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile path
    ReadTextFile = stm.ReadText
    stm.Close
    Exit Function
ErrOut:
    ReadTextFile = ""
End Function

Public Sub WriteTextFile(path As String, content As String)
    On Error GoTo ErrOut
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8"
    stm.Open: stm.WriteText content
    stm.Position = 0: stm.Type = 1: stm.Position = 3
    Dim out As Object: Set out = CreateObject("ADODB.Stream")
    out.Type = 1: out.Open
    stm.CopyTo out
    out.SaveToFile path, 2
    out.Close: stm.Close
    Exit Sub
ErrOut:
End Sub

Public Sub EnsureFolder(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureFolder parent
    End If
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Public Function FileExists(path As String) As Boolean
    FileExists = Len(Dir$(path)) > 0
End Function

Public Function FolderExists(path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(path)
End Function

' ============================================================================
' String Utilities
' ============================================================================

Public Function SafeName(ByVal text As String) As String
    text = Trim$(text)
    If Len(text) = 0 Then text = "blank"
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        text = Replace(text, CStr(bad), "_")
    Next bad
    If Len(text) > 80 Then text = Left$(text, 80)
    SafeName = text
End Function

Public Function FormatFieldValue(val As Variant, Optional fieldType As String = "text") As String
    If IsNull(val) Or IsEmpty(val) Then FormatFieldValue = "": Exit Function
    If IsObject(val) Then
        If TypeName(val) = "Collection" Then
            Dim parts() As String: ReDim parts(val.Count - 1)
            Dim i As Long
            For i = 1 To val.Count: parts(i - 1) = CStr(val(i)): Next i
            FormatFieldValue = Join(parts, "; "): Exit Function
        End If
        FormatFieldValue = "": Exit Function
    End If
    Dim s As String: s = CStr(val)
    Select Case fieldType
        Case "date"
            If VarType(val) = vbDate Then
                FormatFieldValue = Format$(CDate(val), "yyyy/mm/dd"): Exit Function
            End If
        Case "currency"
            If IsNumeric(s) Then
                Dim n As Double: n = CDbl(s)
                If n = Int(n) Then
                    FormatFieldValue = Format$(CLng(n), "#,0")
                Else
                    FormatFieldValue = Format$(n, "#,0.##")
                End If
                Exit Function
            End If
    End Select
    FormatFieldValue = s
End Function

Public Function GetFieldLabel(fieldName As String) As String
    GetFieldLabel = Replace(fieldName, "_", " ")
End Function

Public Function GetFieldGroup(fieldName As String) As String
    Dim pos As Long: pos = InStr(fieldName, "_")
    If pos > 1 And pos < Len(fieldName) Then GetFieldGroup = Left$(fieldName, pos - 1)
End Function

Public Function GetFieldShortName(fieldName As String) As String
    Dim pos As Long: pos = InStr(fieldName, "_")
    If pos > 1 And pos < Len(fieldName) Then
        GetFieldShortName = Mid$(fieldName, pos + 1)
    Else
        GetFieldShortName = fieldName
    End If
End Function

Public Function CountFieldGroups(fields As Collection) As Long
    Dim groups As Object: Set groups = NewDict()
    Dim i As Long
    For i = 1 To fields.Count
        Dim g As String: g = GetFieldGroup(CStr(fields(i)))
        If Len(g) > 0 And Not groups.Exists(g) Then groups.Add g, True
    Next i
    CountFieldGroups = groups.Count
End Function
