Attribute VB_Name = "modLambdaLoader"
Option Explicit

' Silent loader for Lambda definitions.
' Source resolution: try the local file first; if it is missing, fall back to
' the raw text of LambdaDefinitions.txt on GitHub.
' Expected format:
' /* metadata comments */
' function.name = LAMBDA(...);
'
' Behavior:
' - No file picker
' - No prompts
' - No log worksheet
' - Silently exits if neither source is available or no definitions are found
' - Existing same-name workbook names are overwritten

' Local source path. Preferred when present. Change here if the file moves.
Private Const LAMBDA_SOURCE_PATH As String = "D:\05_Coding test\excel_lambda\LambdaDefinitions.txt"

' Fallback source URL (used only when the local file is missing).
Private Const LAMBDA_SOURCE_URL As String = _
    "https://raw.githubusercontent.com/siriusun/excel_lambda/main/LambdaDefinitions.txt"

Public Sub LoadLambdaDefinitions()

    Dim txt As String
    Dim defs As Collection
    Dim item As Variant
    Dim wb As Workbook
    Dim lambdaName As String
    Dim lambdaFormula As String

    On Error GoTo CleanExit

    Set wb = ActiveWorkbook

    ' Prefer the local copy; fall back to GitHub if it is not found.
    If Dir(LAMBDA_SOURCE_PATH) <> "" Then
        txt = ReadTextFileUtf8(LAMBDA_SOURCE_PATH)
    Else
        txt = FetchTextFromUrl(LAMBDA_SOURCE_URL)
    End If

    If Len(txt) = 0 Then GoTo CleanExit

    txt = HtmlDecodeBasic(txt)
    txt = RemoveBlockComments(txt)

    Set defs = ParseLambdaDefinitions(txt)
    If defs Is Nothing Then GoTo CleanExit
    If defs.Count = 0 Then GoTo CleanExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Round 1: create placeholder names first.
    ' This helps when one Lambda references another Lambda that appears later in the file.
    For Each item In defs
        lambdaName = CStr(item(0))

        On Error Resume Next
        wb.Names(lambdaName).Delete
        wb.Names.Add Name:=lambdaName, RefersTo:="=""Loading..."""
        On Error GoTo CleanExit
    Next item

    ' Round 2: set real formulas.
    For Each item In defs
        lambdaName = CStr(item(0))
        lambdaFormula = CStr(item(1))

        Call SetDefinedNameFormula(wb, lambdaName, lambdaFormula)
    Next item

CleanExit:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Function ReadTextFileUtf8(ByVal filePath As String) As String

    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadTextFileUtf8 = .ReadText
        .Close
    End With

End Function

' Synchronous GET via late-bound MSXML2.XMLHTTP.
' Returns the response body decoded as UTF-8. Empty string on any failure
' (network error, non-200 status, empty body) so callers can silently skip.
Private Function FetchTextFromUrl(ByVal url As String) As String

    Dim http As Object
    Dim body As String

    On Error GoTo Failed

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.setRequestHeader "Accept", "text/plain"
    http.send

    If http.Status <> 200 Then GoTo Failed
    If LenB(http.responseText) = 0 Then GoTo Failed

    FetchTextFromUrl = http.responseText
    Exit Function

Failed:
    FetchTextFromUrl = ""

End Function

Private Function HtmlDecodeBasic(ByVal txt As String) As String

    txt = Replace(txt, "&lt;", "<")
    txt = Replace(txt, "&gt;", ">")
    txt = Replace(txt, "&amp;", "&")
    txt = Replace(txt, "&quot;", """")
    txt = Replace(txt, "&#39;", "'")

    HtmlDecodeBasic = txt

End Function

Private Function RemoveBlockComments(ByVal txt As String) As String

    Dim result As String
    Dim startPos As Long
    Dim openPos As Long
    Dim closePos As Long

    startPos = 1

    Do
        openPos = InStr(startPos, txt, "/*")

        If openPos = 0 Then
            result = result & Mid$(txt, startPos)
            Exit Do
        End If

        result = result & Mid$(txt, startPos, openPos - startPos)

        closePos = InStr(openPos + 2, txt, "*/")
        If closePos = 0 Then Exit Do   ' Unterminated comment: drop the remainder.

        startPos = closePos + 2
    Loop

    RemoveBlockComments = result

End Function

Private Function ParseLambdaDefinitions(ByVal txt As String) As Collection

    Dim defs As New Collection
    Dim i As Long
    Dim n As Long
    Dim startName As Long
    Dim lambdaName As String
    Dim formulaBody As String
    Dim formulaStart As Long
    Dim ch As String
    Dim depth As Long
    Dim inString As Boolean

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    n = Len(txt)
    i = 1

    Do While i <= n

        Do While i <= n And IsWhiteSpace(Mid$(txt, i, 1))
            i = i + 1
        Loop

        If i > n Then Exit Do

        startName = i

        Do While i <= n
            ch = Mid$(txt, i, 1)
            If ch = "=" Or IsWhiteSpace(ch) Then Exit Do
            i = i + 1
        Loop

        lambdaName = Trim$(Mid$(txt, startName, i - startName))

        Do While i <= n And IsWhiteSpace(Mid$(txt, i, 1))
            i = i + 1
        Loop

        If i > n Or Mid$(txt, i, 1) <> "=" Then
            i = i + 1
            GoTo ContinueLoop
        End If

        i = i + 1

        Do While i <= n And IsWhiteSpace(Mid$(txt, i, 1))
            i = i + 1
        Loop

        formulaStart = i
        depth = 0
        inString = False

        Do While i <= n

            ch = Mid$(txt, i, 1)

            If ch = """" Then
                If inString And i < n And Mid$(txt, i + 1, 1) = """" Then
                    i = i + 2
                    GoTo ContinueFormulaLoop
                Else
                    inString = Not inString
                End If
            End If

            If Not inString Then
                If ch = "(" Then
                    depth = depth + 1
                ElseIf ch = ")" Then
                    If depth > 0 Then depth = depth - 1
                ElseIf ch = ";" And depth = 0 Then
                    Exit Do
                End If
            End If

            i = i + 1

ContinueFormulaLoop:
        Loop

        formulaBody = Trim$(Mid$(txt, formulaStart, i - formulaStart))

        If lambdaName <> "" And formulaBody <> "" Then
            If Left$(formulaBody, 1) <> "=" Then formulaBody = "=" & formulaBody
            defs.Add Array(lambdaName, formulaBody)
        End If

        If i <= n And Mid$(txt, i, 1) = ";" Then i = i + 1

ContinueLoop:
    Loop

    Set ParseLambdaDefinitions = defs

End Function

Private Function SetDefinedNameFormula(ByVal wb As Workbook, ByVal lambdaName As String, ByVal formulaText As String) As Boolean

    Dim nm As Name

    On Error GoTo Failed

    Set nm = wb.Names(lambdaName)

    On Error Resume Next
    CallByName nm, "RefersTo2", VbLet, formulaText

    If Err.Number <> 0 Then
        Err.Clear
        nm.RefersTo = formulaText
    End If

    On Error GoTo Failed

    SetDefinedNameFormula = True
    Exit Function

Failed:
    SetDefinedNameFormula = False

End Function

Private Function IsWhiteSpace(ByVal ch As String) As Boolean

    IsWhiteSpace = ch = " " Or ch = vbTab Or ch = vbLf Or ch = vbCr

End Function
