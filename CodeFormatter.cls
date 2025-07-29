VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CODE FORMATTER CLASS MODULE
' Original Author: Tibor Fellegi (GitHub: fellegitt, Gmail: same username – derive address as needed)
' Last updated: 2025-07-29
'
' MIT License
'
' Copyright (c) 2025 Tibor Fellegi (fellegitt)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
Option Explicit
 
Private Enum Enum_InsertSpace
    Above
    Below
    Both
End Enum
 
Private Enum Enum_QuotePosition
    FirstOccurence
    LastOccurence
End Enum
 
Private Type Type_VbaKeywords
    SpaceAbove As Variant
    SpaceBelow As Variant
    SpaceBoth As Variant
    Indent As Variant
    Outdent As Variant
    OutdentAndIndent As Variant
    IsProcedureStart As Variant
    IsProcedureEnd As Variant
End Type
 
Private Type Type_Fields
    Keywords As Type_VbaKeywords
    SpacingKeywords As Object
    OriginalCode As Collection
    FormattedCode As Collection
    Indentation As Long
    IsMultiline As Boolean
End Type
 
Private this As Type_Fields
 
'Class Initialization
Private Sub Class_Initialize()
 
    InitializeKeywords
    InitializeSpacingDictionary
 
End Sub
 
Private Sub InitializeKeywords()
 
    With this.Keywords
        .SpaceAbove = Array("'", "With ", "Public Enum", "Private Enum", "Type ", "Public Type ", _
            "Private Type ")
        .SpaceBelow = Array("End With", "End Type", "End Enum")
        .SpaceBoth = Array("On Error Resume Next", "On Error GoTo", "Resume", "Resume Next", "Stop ", _
            "Sub", "Public Sub", "Enum ", "Private Sub", "Function ", "Public Function", "Private Function", _
            "End Sub", "End Property", "End Function", "Property ", "Public Property ", "Private Property ")
        .Indent = Array("If", "#If", "For", "With", "Do", "Select Case", _
            "Enum", "Private Enum", "Public Enum", "Type", "Public Type", "Private Type", _
            "Property", "Public Property", "Private Property", "While")
        .Outdent = Array("End If", "#End If", "Next", "End With", "Loop", "Loop Until", "Loop While", _
            "End Select", "End Type", "End Enum", "End Property", "Wend")
        .OutdentAndIndent = Array("Else", "ElseIf", "Case")
        .IsProcedureStart = Array("Sub", "Private Sub", "Public Sub", "Function", "Property", "Private Property", "Public Property", "Private Function", "Public Function")
        .IsProcedureEnd = Array("End Sub", "End Function", "End Property")
    End With
 
End Sub
 
'Public Scope
Public Sub FormatProject()
 
    Dim component As Object, vbComp As Object
 
    Set vbComp = ThisWorkbook.VBProject
    For Each component In vbComp.VBComponents
        FormatModule component.CodeModule
    Next
    Set vbComp = Nothing
    Debug.Print "Project was formatted successfully."
 
End Sub
 
Public Sub FormatModule(Optional ByRef module As Variant)
 
    Dim vbComp As Object
 
    If IsMissing(module) = True Then
        Set vbComp = Application.VBE.ActiveCodePane.CodeModule
    ElseIf TypeName(module) = "String" Then
        Set vbComp = ThisWorkbook.VBProject.VBComponents(module).CodeModule
    Else
        Set vbComp = module
    End If
 
    With this
        Set .OriginalCode = New Collection
        Set .FormattedCode = New Collection
    End With
 
    If this.SpacingKeywords Is Nothing Then
        InitializeKeywords
    End If
    ProcessModule vbComp
    Set vbComp = Nothing
 
End Sub
 
'Private Scope
Private Sub ProcessModule(ByRef vbComp As Object)
 
    If vbComp.countoflines = 0 Then
        Exit Sub
    End If
    LinesToCollection Split(vbComp.lines(1, vbComp.countoflines), vbCrLf)
    ProcessLines
    vbComp.DeleteLines 1, vbComp.countoflines
    vbComp.AddFromString Join(CollectionToArray(this.FormattedCode), vbCrLf)
 
End Sub
 
Private Sub PreprocessCode()
 
    Dim i As Long
 
    For i = this.OriginalCode.count To 1 Step -1
        SingleLineIfToMultiline i
        ColonDelimitedLinesToMultiline i
    Next i
 
End Sub
 
Private Sub ProcessLines()
 
    Dim i As Long
 
    this.Indentation = 0
    PreprocessCode
    For i = 1 To this.OriginalCode.count
        HandleSpacing this.FormattedCode, ApplyIndentation(CheckForMultiLine(i))
    Next i
    OrganizeDeclarations
    PostprocessCode
 
End Sub
 
Private Sub PostprocessCode()
 
    FormatDeclarations
    RemoveDoubleBlankLines
    RemoveSpacesAfterComments
 
End Sub
 
Private Sub LinesToCollection(ByVal lines As Variant)
 
    Dim i As Long
 
    Set this.OriginalCode = New Collection
    For i = LBound(lines) To UBound(lines)
        lines(i) = CleanLine(lines(i))
        If LenB(lines(i)) > 0 Then
            this.OriginalCode.Add lines(i)
        End If
    Next i
 
End Sub
 
Private Function CollectionToArray(ByVal col As Collection) As Variant
 
    Dim i As Long
    Dim arr() As String
 
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i
    CollectionToArray = arr
 
End Function
 
Private Function SingleLineIfToMultiline(i As Long)
 
    Dim thenPos As Long
    Dim conditionPart As String, actionPart As String, currentLine As String
 
    currentLine = this.OriginalCode(i)
    If Left$(currentLine, 3) = "If " Then
        thenPos = InStr(1, currentLine, " Then ", vbTextCompare)
        If thenPos > 0 And Len(currentLine) > thenPos + 5 And InStrRev(currentLine, "'") < thenPos Then
            conditionPart = Left$(currentLine, thenPos + 4)
            actionPart = Mid$(currentLine, thenPos + 6)
 
            With this.OriginalCode
                .Remove i
                .Add "End If", before:=i
                .Add String$(4, " ") & Trim$(actionPart), before:=i
                .Add Trim$(conditionPart), before:=i
            End With
 
        End If
    End If
 
End Function
 
Private Function ColonDelimitedLinesToMultiline(i As Long)
 
    Dim parts As Collection
    Dim inQuotes As Boolean
    Dim segment As String, currentLine As String
    Dim j As Long, nextColon As Long, nextQuote As Long, pos As Long, startPos As Long
 
    currentLine = this.OriginalCode(i)
    Set parts = New Collection
    startPos = 1
    pos = 1
    inQuotes = False
    Do While pos <= Len(currentLine)
        nextColon = InStr(pos, currentLine, ":")
        nextQuote = InStr(pos, currentLine, """")
        If nextQuote > 0 And (nextQuote < nextColon Or nextColon = 0) Then
            inQuotes = Not inQuotes
            pos = nextQuote + 1
        ElseIf nextColon > 0 Then
            If Not inQuotes Then
                If Mid(currentLine, nextColon, 2) <> ":=" And _
                    Not IsLineCommented(currentLine) And _
                    Not IsInQuotes(currentLine, ":", , nextColon) And _
                    Not IsInComment(currentLine, ":", , nextColon) Then
                    segment = Trim(Mid(currentLine, startPos, nextColon - startPos))
                    If segment <> "" Then
                        parts.Add segment
                    End If
                    startPos = nextColon + 1
                End If
            End If
            pos = nextColon + 1
        Else
            Exit Do
        End If
    Loop
    If startPos <= Len(currentLine) Then
        segment = Trim(Mid(currentLine, startPos))
        If segment <> "" Then
            parts.Add segment
        End If
    End If
    If parts.count > 1 Then
 
        With this.OriginalCode
            .Remove i
            For j = parts.count To 1 Step -1
                .Add parts(j), before:=i
            Next j
        End With
 
    End If
 
End Function
 
Private Function IsLineCommented(line As String) As Boolean
 
    IsLineCommented = Left(CleanLine(line), 1) = "'"
 
End Function
 
Private Function IsInComment(ByVal str As String, ByVal Target As String, _
    Optional ByVal position As Enum_QuotePosition = FirstOccurence, _
    Optional ByVal positionIndex As Long = 0) As Boolean
 
    Dim commentPos As Long, targetPos As Long
 
    Select Case position
        Case FirstOccurence
            targetPos = InStr(1, str, Target)
        Case LastOccurence
            targetPos = InStrRev(str, Target)
        Case Else
            targetPos = IIf(positionIndex > 0, positionIndex, InStr(1, str, Target))
    End Select
    commentPos = InStr(1, str, "'")
    If targetPos = 0 Then
        IsInComment = False
    Else
        IsInComment = (commentPos > 0 And targetPos > commentPos)
    End If
 
End Function
 
Private Function IsCommentLine(ByVal str As String) As Boolean
 
    IsCommentLine = Left(Replace(Trim(str), vbTab, vbNullString), 1) = "'"
 
End Function
 
Private Function IsInQuotes(ByVal str As String, _
    ByVal Target As String, _
    Optional ByVal position As Enum_QuotePosition = FirstOccurence, _
    Optional ByVal positionIndex As Long = 0) As Boolean
 
    Dim inQuotes As Boolean
    Dim quoteStart As Long, quoteEnd As Long, startPos As Long, endPos As Long
 
    Select Case position
        Case FirstOccurence
            startPos = InStr(1, str, Target)
        Case LastOccurence
            startPos = InStrRev(str, Target)
        Case Else
            startPos = IIf(positionIndex > 0, positionIndex, InStr(1, str, Target))
    End Select
    If startPos > 0 Then
        endPos = startPos + Len(Target) - 1
    Else
        IsInQuotes = False
        Exit Function
    End If
    quoteStart = InStr(1, str, """")
    Do While quoteStart > 0
        quoteEnd = InStr(quoteStart + 1, str, """")
        If quoteEnd > 0 Then
            If startPos >= quoteStart And endPos <= quoteEnd Then
                IsInQuotes = True
                Exit Function
            End If
            quoteStart = InStr(quoteEnd + 1, str, """")
        Else
            Exit Do
        End If
    Loop
    IsInQuotes = False
 
End Function
 
Private Function CheckForMultiLine(ByRef i As Long) As String
 
    Dim isContinuation As Boolean
    Dim currentLine As String
 
    currentLine = this.OriginalCode(i)
    If Right(currentLine, 1) = "_" Then
        this.IsMultiline = True
        Do
            i = i + 1
            currentLine = currentLine & vbCrLf & ApplyIndentation(this.OriginalCode(i))
            isContinuation = (Right(this.OriginalCode(i), 1) = "_")
        Loop While isContinuation
    End If
    If Not isContinuation Then
        this.IsMultiline = False
    End If
    CheckForMultiLine = currentLine
 
End Function
 
Private Function CleanLine(ByVal line As String) As String
 
    CleanLine = Replace(Trim(line), vbTab, "")
 
End Function
 
Private Sub FormatDeclarations()
 
    Dim currentLine As String, lastDeclarationType As String
    Dim i As Long
 
    For i = 1 To this.FormattedCode.count
        currentLine = LCase(CleanLine(this.FormattedCode(i)))
        Select Case True
            Case StartsWith(currentLine, "dim")
                lastDeclarationType = "Dim"
            Case StartsWith(currentLine, "const")
                lastDeclarationType = "Const"
            Case StartsWith(currentLine, "static")
                lastDeclarationType = "Static"
            Case lastDeclarationType <> ""
                this.FormattedCode.Add " ", before:=i
                lastDeclarationType = ""
                i = i + 1
        End Select
    Next i
 
End Sub
 
Private Function StartsWith(ByVal text As String, ByVal prefix As String) As Boolean
 
    StartsWith = (Left$(text, Len(prefix)) = LCase$(prefix))
 
End Function
 
Private Sub RemoveDoubleBlankLines()
 
    Dim i As Long
 
    For i = this.FormattedCode.count To 2 Step -1
        If CleanLine(this.FormattedCode(i)) = "" And CleanLine(this.FormattedCode(i - 1)) = "" Then
            this.FormattedCode.Remove i
        End If
    Next i
 
End Sub
 
Private Sub RemoveSpacesAfterComments()
 
    Dim i As Long
 
    For i = this.FormattedCode.count To 2 Step -1
        If CleanLine(this.FormattedCode(i)) = "" And Left(CleanLine(this.FormattedCode(i - 1)), 1) = "'" Then
            this.FormattedCode.Remove i
        End If
    Next i
 
End Sub
 
Private Function ApplyIndentation(ByVal line As String) As String
 
    Static selectCase As Boolean
 
    Dim space As String
 
    If this.IsMultiline Then
        this.Indentation = this.Indentation + 1
    End If
    If InStr(1, line, "End Select") = 1 Then
        this.Indentation = Application.max(this.Indentation - 1, 0)
        selectCase = False
    End If
    Select Case True
        Case Indent(line)
            space = String(this.Indentation, vbTab)
            this.Indentation = this.Indentation + 1
        Case Outdent(line)
            this.Indentation = Application.max(this.Indentation - 1, 0)
            space = String(this.Indentation, vbTab)
        Case OutdentAndIndent(line)
            this.Indentation = Application.max(this.Indentation - 1, 0)
            space = String(this.Indentation, vbTab)
            ApplyIndentation = space & line
            this.Indentation = this.Indentation + 1
            Exit Function
        Case Else
            space = String(this.Indentation, vbTab)
    End Select
    ApplyIndentation = space & line
    If InStr(1, line, "Select Case") = 1 Then
        selectCase = True
        this.Indentation = this.Indentation + 1
    End If
    If this.IsMultiline Then
        this.Indentation = this.Indentation - 1
    End If
 
End Function
 
Private Sub InitializeSpacingDictionary()
 
    If this.SpacingKeywords Is Nothing Then
        Set this.SpacingKeywords = CreateObject("Scripting.Dictionary")
 
        With this.Keywords
            AddKeywordsToDictionary .SpaceAbove, Above
            AddKeywordsToDictionary .SpaceBelow, Below
            AddKeywordsToDictionary .SpaceBoth, Both
        End With
 
    End If
 
End Sub
 
Private Sub AddKeywordsToDictionary(ByVal Keywords As Variant, ByVal position As Enum_InsertSpace)
 
    Dim kw As Variant
 
    For Each kw In Keywords
        If Not this.SpacingKeywords.Exists(kw) Then
            this.SpacingKeywords.Add kw, position
        End If
    Next kw
 
End Sub
 
Private Sub HandleSpacing(ByRef procLines As Collection, ByVal line As String)
 
    Dim i As Long
    Dim insertAbove As Boolean, insertBelow As Boolean
    Dim kw As Variant
 
    insertAbove = False
    insertBelow = False
    For Each kw In this.SpacingKeywords.keys
        If Left(CleanLine(line), Len(kw)) = kw Then
            Select Case this.SpacingKeywords(kw)
                Case Enum_InsertSpace.Above
                    insertAbove = True
                Case Enum_InsertSpace.Below
                    insertBelow = True
                Case Enum_InsertSpace.Both
                    insertAbove = True
                    insertBelow = True
            End Select
            Exit For
        End If
    Next kw
    If insertAbove And procLines.count > 1 Then
        If IsCommentLine(CleanLine(line)) Then
            For i = procLines.count To 2 Step -1
                If Not IsCommentLine(CleanLine(procLines(i))) Then
                    procLines.Add " ", after:=i
                    Exit For
                End If
            Next i
        Else
            procLines.Add " "
        End If
    End If
    procLines.Add line
    If insertBelow Then
        procLines.Add " "
    End If
 
End Sub
 
Private Function ContainsKeyword(ByVal ln As String, ByVal Keywords As Variant) As Boolean
 
    Dim keywordLength As Long
    Dim kw As Variant
 
    ln = CleanLine(ln)
    If TypeName(Keywords) = "Variant()" Then
        For Each kw In Keywords
            keywordLength = Len(kw)
            If StrComp(Left(ln, keywordLength), kw, vbTextCompare) = 0 Then
                If Len(ln) = keywordLength Or Mid(ln, keywordLength + 1, 1) = " " Then
                    ContainsKeyword = True
                    Exit Function
                End If
            End If
        Next kw
    ElseIf TypeName(Keywords) = "String" Then
        keywordLength = Len(Keywords)
        If StrComp(Left(ln, keywordLength), Keywords, vbTextCompare) = 0 Then
            If Len(ln) = keywordLength Or Mid(ln, keywordLength + 1, 1) = " " Then
                ContainsKeyword = True
                Exit Function
            End If
        End If
    End If
    ContainsKeyword = False
 
End Function
 
Private Function Indent(ByVal ln As String) As Boolean
 
    Indent = ContainsKeyword(ln, this.Keywords.Indent) Or IsProcedureStart(ln)
 
End Function
 
Private Function Outdent(ByVal ln As String) As Boolean
 
    Outdent = ContainsKeyword(ln, this.Keywords.Outdent) Or IsProcedureEnd(ln)
 
End Function
 
Private Function OutdentAndIndent(ByVal ln As String) As Boolean
 
    OutdentAndIndent = ContainsKeyword(ln, this.Keywords.OutdentAndIndent)
 
End Function
 
Private Function IsProcedureStart(ByVal ln As String) As Boolean
 
    IsProcedureStart = ContainsKeyword(ln, this.Keywords.IsProcedureStart)
 
End Function
 
Private Function IsProcedureEnd(ByVal ln As String) As Boolean
 
    IsProcedureEnd = ContainsKeyword(ln, this.Keywords.IsProcedureEnd)
 
End Function
 
Private Sub OrganizeDeclarations()
 
    Dim declarations As Object
    Dim i As Long
 
    For i = this.FormattedCode.count To 1 Step -1
        If IsProcedureEnd(this.FormattedCode(i)) Then
            Set declarations = CreateObject("Scripting.Dictionary")
            Do Until IsProcedureStart(this.FormattedCode(i))
                HandleDeclaration declarations, i
                i = i - 1
            Loop
            If Not declarations Is Nothing Then
                WriteDeclarations declarations, i
                Set declarations = Nothing
            End If
        End If
    Next i
 
End Sub
 
Private Sub HandleDeclaration(declarations As Object, currentLineIndex As Long)
 
    Dim arr As Variant
    Dim category As String, line As String
    Dim i As Long
 
    line = this.FormattedCode(currentLineIndex)
    category = DeclarationType(line)
    If category <> vbNullString Then
        line = NormalizeSpaces(CleanLine(Split(line, "'")(0)))
        arr = SplitRespectingQuotes(line, ",")
        For i = 0 To UBound(arr)
            HandleVariable declarations, category, arr(i)
        Next i
        this.FormattedCode.Remove (currentLineIndex)
    End If
 
End Sub
 
Private Sub HandleVariable(ByRef dict As Object, category As String, ByVal variable As String)
 
    Dim words As Variant
    Dim ObjectIsNew As Boolean
    Dim asPos As Long, j As Long
    Dim newWords() As String, firstWord As String, declarationPart As String, variableName As String, variableType As String, constantPart As String, parts As String, variableNameWithBounds As String
 
    constantPart = ""
 
    If InStr(1, variable, "=") > 0 Then
        constantPart = "=" & SplitRespectingQuotes(variable, "=")(1)
 
    End If
    parts = SplitRespectingQuotes(variable, "=")(0)
    declarationPart = Trim(parts)
    asPos = InStr(1, LCase(declarationPart), " as ")
    If asPos > 0 Then
        variableNameWithBounds = Trim(Left(declarationPart, asPos - 1))
        variableType = Trim(Mid(declarationPart, asPos + 4))
        If LCase(variableType) Like "new *" Then
            ObjectIsNew = True
            variableType = Trim(Mid(variableType, 5))
        Else
            ObjectIsNew = False
        End If
        variableName = variableNameWithBounds
        words = Split(variableName, " ")
        firstWord = LCase(words(0))
        If firstWord = "dim" Or firstWord = "const" Or firstWord = "static" Then
            If UBound(words) >= 1 Then
                ReDim newWords(0 To UBound(words) - 1)
                For j = 1 To UBound(words)
                    newWords(j - 1) = words(j)
                Next j
                variableName = Join(newWords, " ")
            Else
                variableName = ""
            End If
        End If
    Else
        variableType = "Variant"
        ObjectIsNew = False
        words = Split(declarationPart, " ")
        firstWord = LCase(words(0))
        If firstWord = "dim" Or firstWord = "const" Or firstWord = "static" Then
            If UBound(words) >= 1 Then
                variableName = words(1)
            Else
                variableName = ""
            End If
        Else
            variableName = words(0)
        End If
    End If
    If Not dict.Exists(category) Then
        dict.Add category, CreateObject("Scripting.Dictionary")
    End If
    If Not dict(category).Exists(variableType) Then
        dict(category).Add variableType, New Collection
    End If
    dict(category)(variableType).Add variableName & "|" & constantPart & "|" & ObjectIsNew
 
End Sub
 
Private Function SplitRespectingQuotes(ByVal inputStr As String, ByVal delimiter As String) As Variant
 
    Dim count As Long, i As Long
    Dim part As String, currentChar As String, char As String, temp As String, Result() As String
    Dim inQuotes As Boolean
 
    ReDim Result(0 To 0)
    part = ""
    inQuotes = False
    count = 0
    For i = 1 To Len(inputStr)
        currentChar = Mid(inputStr, i, 1)
        If currentChar = """" Then
            inQuotes = Not inQuotes
            part = part & currentChar
        ElseIf currentChar = delimiter And Not inQuotes Then
            ReDim Preserve Result(0 To count)
            Result(count) = part
            part = ""
            count = count + 1
        Else
            part = part & currentChar
        End If
    Next i
    ReDim Preserve Result(0 To count)
    Result(count) = part
    SplitRespectingQuotes = Result
 
End Function
 
Private Sub WriteDeclarations(ByRef declarations As Object, ByVal x As Long)
 
    Dim l As Long, k As Long, j As Long, i As Long
    Dim variableTypeName As String, categoryName As String, declarationString As String
    Dim variable As Variant
    Dim variables As Collection
    Dim categories As Object
 
    i = x + 1
    If declarations.count > 0 Then
        For j = 0 To declarations.count - 1
            Set categories = declarations(declarations.keys()(j))
            categoryName = declarations.keys()(j)
            SortArray categories.keys()
            For k = 0 To categories.count - 1
                variableTypeName = categories.keys()(k)
                Set variables = categories(variableTypeName)
                declarationString = vbTab & categoryName & " "
                For l = 1 To variables.count
                    variable = Split(variables(l), "|")
                    declarationString = declarationString & variable(0) & " As " & _
                        IIf(variable(2) = "True", "New ", vbNullString) & variableTypeName & variable(1) & IIf(l < variables.count, ",", vbNullString)
                Next l
                this.FormattedCode.Add declarationString, after:=i
                i = i + 1
            Next k
            this.FormattedCode.Add " ", after:=i
            i = i + 1
        Next j
    End If
 
End Sub
 
Private Function DeclarationType(line As String) As String
 
    Dim keyword As Variant, Keywords As Variant
 
    Keywords = Array("Dim", "Const", "Static")
    For Each keyword In Keywords
        If ContainsKeyword(line, keyword) Then
            DeclarationType = keyword
            Exit Function
        End If
    Next keyword
    DeclarationType = vbNullString
 
End Function
 
Private Function NormalizeSpaces(text As String) As String
 
    Dim regex As Object
 
    Set regex = CreateObject("VBScript.RegExp")
 
    With regex
        .Pattern = "\s+"
        .Global = True
    End With
 
    NormalizeSpaces = regex.Replace(text, " ")
 
End Function
 
Private Sub SortArray(ByRef arr As Variant, Optional ByVal descending As Boolean = False, _
    Optional ByVal low As Long = -1, Optional ByVal high As Long = -1)
 
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
 
    If low = -1 Then
        low = LBound(arr)
    End If
    If high = -1 Then
        high = UBound(arr)
    End If
    If low >= high Then
        Exit Sub
    End If
    i = low
    j = high
    pivot = arr((low + high) \ 2)
    Do While i <= j
        If descending Then
            Do While arr(i) > pivot
                i = i + 1
            Loop
            Do While arr(j) < pivot
                j = j - 1
            Loop
        Else
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
        End If
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then
        SortArray arr, descending, low, j
    End If
    If i < high Then
        SortArray arr, descending, i, high
    End If
 
End Sub
 



