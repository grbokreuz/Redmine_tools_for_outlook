VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INVALID_JSON      As Long = 1
Const INVALID_OBJECT    As Long = 2
Const INVALID_ARRAY     As Long = 3
Const INVALID_BOOLEAN   As Long = 4
Const INVALID_NULL      As Long = 5
Const INVALID_KEY       As Long = 6

'
'   parse string and create JSON object (Dictionary or Collection in VB)
'
Public Function parse(ByRef str As String) As Object
'    If debug_ Then DebugPrintFile "JSONLIB parse  called ( " & str & " )"
    Dim index As Long
    index = 1
    
    On Error Resume Next

    Call skipChar(str, index)
    Select Case Mid(str, index, 1)
    Case "{"
        Set parse = parseObject(str, index)
    Case "["
        Set parse = parseArray(str, index)
    End Select

End Function

'
'   parse collection of key/value (Dictionary in VB)
'
Private Function parseObject(ByRef str As String, ByRef index As Long) As Object
'    If debug_ Then DebugPrintFile "JSONLIB parseObject called ( " & index & " )"
    Dim a As Integer
    Set parseObject = New Dictionary
    
    If debug_ Then DebugPrintFile "index is " & index & " on (" & str & " )"
    
    ' "{"
    Call skipChar(str, index)
    If Mid(str, index, 1) <> "{" Then Err.Raise vbObjectError + INVALID_OBJECT, Description:="char " & index & " : " & Mid(str, index)
    index = index + 1
    
    Do
    
        Call skipChar(str, index)
        If "}" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        End If
        
        Dim key As String
        
        ' add key/value pair
        parseObject.Add key:=parseKey(str, index), Item:=parseValue(str, index)
        a = 1
    Loop

End Function

'
'   parse list (Collection in VB)
'
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection
'    If debug_ Then DebugPrintFile "JSONLIB parseArray called ( " & index & " )"
    Set parseArray = New Collection
    
    ' "["
    Call skipChar(str, index)
    If Mid(str, index, 1) <> "[" Then Err.Raise vbObjectError + INVALID_ARRAY, Description:="char " & index & " : " + Mid(str, index)
    index = index + 1
    
    Do
        
        Call skipChar(str, index)
        If "]" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        End If
        
        ' add value
        parseArray.Add parseValue(str, index)
        
    Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef index As Long)
 '   If debug_ Then DebugPrintFile "JSONLIB parseValue called ( " & index & " )"
    Call skipChar(str, index)
    
    Select Case Mid(str, index, 1)
    Case "{"
        Set parseValue = parseObject(str, index)
    Case "["
        Set parseValue = parseArray(str, index)
    Case """", "'"
        parseValue = parseString(str, index)
        If debug_ Then DebugPrintFile "this val is " & parseValue
    Case "t", "f"
        parseValue = parseBoolean(str, index)
    Case "n"
        parseValue = parseNull(str, index)
        If debug_ Then DebugPrintFile "this val is " & parseValue
    Case Else
        parseValue = parseNumber(str, index)
        If debug_ Then DebugPrintFile "this val is " & parseValue
    End Select

End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String
'    If debug_ Then DebugPrintFile "JSONLIB parseString called ( " & index & " )"
    Dim quote   As String
    Dim char    As String
    Dim code    As String

If index > 512000 Then
MsgBox "i cant read redmine data. parse failed on [" & Mid(parseString, -10, 10) & "]"
End If
    Call skipChar(str, index)
    quote = Mid(str, index, 1)
    index = index + 1
    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
        Select Case (char)
        Case "\"
            index = index + 1
            char = Mid(str, index, 1)
            Select Case (char)
            Case """", "\", "/"
                parseString = parseString & char
                index = index + 1
            Case "b"
                parseString = parseString & vbBack
                index = index + 1
            Case "f"
                parseString = parseString & vbFormFeed
                index = index + 1
            Case "n"
                parseString = parseString & vbNewLine
                index = index + 1
            Case "r"
                parseString = parseString & vbCr
                index = index + 1
            Case "t"
                parseString = parseString & vbTab
                index = index + 1
            Case "u"
                index = index + 1
                code = Mid(str, index, 4)
                parseString = parseString & ChrW(val("&h" + code))
                index = index + 4
            End Select
        Case quote
            index = index + 1
            Exit Function
        Case Else
            parseString = parseString & char
            index = index + 1
        End Select
    Loop

End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)
'    If debug_ Then DebugPrintFile "JSONLIB parseNumber called ( " & index & " )"
    Dim value   As String
    Dim char    As String
    
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
        If InStr("+-0123456789.eE", char) Then
            value = value & char
            index = index + 1
        Else
            If InStr(value, ".") Or InStr(value, "e") Or InStr(value, "E") Then
                parseNumber = CDbl(value)
            Else
                parseNumber = CInt(value)
            End If
            Exit Function
        End If
    Loop


End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean
'    If debug_ Then DebugPrintFile "JSONLIB parseBoolean called ( " & index & " )"
    Call skipChar(str, index)
    If Mid(str, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf Mid(str, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    Else
        Err.Raise vbObjectError + INVALID_BOOLEAN, Description:="char " & index & " : " & Mid(str, index)
    End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)
'    If debug_ Then DebugPrintFile "JSONLIB parseNull called ( " & index & " )"
    Call skipChar(str, index)
    If Mid(str, index, 4) = "null" Then
        parseNull = Null
        index = index + 4
    Else
        Err.Raise vbObjectError + INVALID_NULL, Description:="char " & index & " : " & Mid(str, index)
    End If

End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String
 '   If debug_ Then DebugPrintFile "JSONLIB parseKey called ( " & index & " )"
    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim char    As String
    
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
 '       DebugPrintFile char
        Select Case (char)
        Case """"
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
                Call skipChar(str, index)
                If Mid(str, index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & index & " : " & parseKey
                End If
            End If
        Case "'"
            squote = Not squote
            index = index + 1
            If Not squote Then
                Call skipChar(str, index)
                If Mid(str, index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & index & " : " & parseKey
                End If
            End If
        Case ":"
            If Not dquote And Not squote Then
                index = index + 1
                Exit Do
            End If
        Case Else
'            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", char) Then
            If InStr(vbCrLf & vbCr & vbLf & vbTab, char) Then '' why space ? added
            Else
                parseKey = parseKey & char
                If debug_ Then DebugPrintFile "this key is " & parseKey
            End If
            index = index + 1
        End Select
    Loop

End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef index As Long)
'    If debug_ Then DebugPrintFile "JSONLIB skipChar called ( " & index & " )"
    While index > 0 And index <= Len(str) And InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Mid(str, index, 1))
        index = index + 1
    Wend

End Sub

Public Function toString(ByRef obj As Variant) As String
'    If debug_ Then DebugPrintFile "JSONLIB toString called"
    Select Case VarType(obj)
        Case vbNull
            toString = "null"
        Case vbDate
            toString = """" & CStr(obj) & """"
        Case vbString
            toString = """" & encode(obj) & """"
        Case vbObject
            Dim bFI, i
            bFI = True
            If TypeName(obj) = "Dictionary" Then
                toString = toString & "{"
                Dim Keys
                Keys = obj.Keys
                For i = 0 To obj.Count - 1
                    If bFI Then bFI = False Else toString = toString & ","
                    Dim key
                    key = Keys(i)
                    toString = toString & """" & key & """:" & toString(obj(key))
                Next i
                toString = toString & "}"
            ElseIf TypeName(obj) = "Collection" Then
                toString = toString & "["
                Dim value
                For Each value In obj
                    If bFI Then bFI = False Else toString = toString & ","
                    toString = toString & toString(value)
                Next value
                toString = toString & "]"
            End If
        Case vbBoolean
            If obj Then toString = "true" Else toString = "false"
        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            toString = multiArray(obj, 1, "", sEB)
        Case Else
            toString = Replace(obj, ",", ".")
    End Select

End Function

Private Function encode(str) As String
'    If debug_ Then DebugPrintFile "JSONLIB encode called"
    Dim i, j, aL1, aL2, C, p

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(str)
        p = True
        C = Mid(str, i, 1)
        For j = 0 To 7
            If C = Chr(aL1(j)) Then
                encode = encode & "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next

        If p Then
            Dim a
            a = AscW(C)
            If a > 31 And a < 127 Then
                encode = encode & C
            ElseIf a > -1 Or a < 65535 Then
                encode = encode & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
            End If
        End If
    Next
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
'    If debug_ Then DebugPrintFile "JSONLIB multiArray called"
    Dim iDU, iDL, i ' Integer DimensionUBound, Integer DimensionLBound
    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)
    
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid(sPB1, i, 1)
        Next
'        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
        multiArray = multiArray & toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        multiArray = multiArray & "["
        For i = iDL To iDU
            multiArray = multiArray & multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then multiArray = multiArray & ","
        Next
        multiArray = multiArray & "]"
        sPT = Left(sPT, iBC - 2)
    End If
    Err.Clear
End Function


