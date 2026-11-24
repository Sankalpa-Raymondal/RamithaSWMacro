Attribute VB_Name = "FillTitleBlock1"
' FillTitleBlock - SOLIDWORKS drawing macro
'
' Fills in the title block (bottom right table) of every sheet from a template file.
'   - With  REMOVE PROJECTION = yes  in the template it also clears the "THIRD ANGLE" text and
'     the projection symbol from the Projection Method box.
'   - Template: TitleBlockTemplate.txt, one line per box:  Label = value
'     The value can be fixed text and/or fields filled in from SOLIDWORKS:
'     {partname} {material} {mass} {date} {sheet} {sheets} {scale} {sheetsize} {filename}
'   - The template is looked for in the drawing's folder, then next to this macro;
'     if there is none, a file dialog asks for it.
'   - For each label it finds the "Label:" heading in the title block and fills the box
'     next to it (just right of the heading, or just below it).
'   - Only EMPTY boxes are filled, unless the template says  OVERWRITE = yes.
' The drawing is not saved.

Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type

Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Const TITLE As String = "Fill Title Block"
Const TEMPLATE_NAME As String = "TitleBlockTemplate.txt"

Dim swApp As Object, swModel As Object
Dim tLabel() As String, tValue() As String, nT As Long, overwrite As Boolean, noProjection As Boolean
Dim fPart As String, fMat As String, fMass As String, fFile As String
Dim curScale As String, curPaper As String

Sub main()
    On Error GoTo failed
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Open a drawing first.", vbExclamation, TITLE: Exit Sub
    End If
    If swModel.GetType <> 3 Then
        MsgBox "The active document is not a drawing.", vbExclamation, TITLE: Exit Sub
    End If

    ' ---- template ----
    Dim tpl As String
    tpl = FindTemplate()
    If tpl = "" Then Exit Sub
    If Not ReadTemplate(tpl) Then
        MsgBox "No ""Label = value"" lines found in" & vbCrLf & tpl, vbExclamation, TITLE: Exit Sub
    End If

    ' ---- information from the model shown in the drawing ----
    ReadModelInfo

    ' ---- every sheet ----
    Dim names As Variant, startSheet As String, i As Long, report As String, filled As Long
    names = swModel.GetSheetNames
    startSheet = swModel.GetCurrentSheet.GetName
    For i = LBound(names) To UBound(names)
        swModel.ActivateSheet names(i)
        report = report & vbCrLf & names(i) & ":" & FillSheet(i - LBound(names) + 1, UBound(names) - LBound(names) + 1, filled)
    Next i
    swModel.ActivateSheet startSheet
    swModel.ForceRebuild3 False
    swModel.GraphicsRedraw2

    MsgBox filled & " box(es) filled from " & Mid(tpl, InStrRev(tpl, "\") + 1) & IIf(overwrite, " (overwrite on)", "") & ":" & _
           report & vbCrLf & vbCrLf & "Save the drawing to keep the changes.", vbInformation, TITLE
    Exit Sub
failed:
    Dim errText As String
    errText = Err.Description & " (" & Err.Number & ")"
    On Error Resume Next
    swModel.EditSheet
    MsgBox "Stopped by an error: " & errText, vbExclamation, TITLE
End Sub

' ---------------------------------------------------------------------------------------
Private Function FindTemplate() As String
    Dim p As String
    p = swModel.GetPathName
    If p <> "" Then
        p = Left(p, InStrRev(p, "\")) & TEMPLATE_NAME
        If Dir(p) <> "" Then FindTemplate = p: Exit Function
    End If
    p = swApp.GetCurrentMacroPathName
    If p <> "" Then
        p = Left(p, InStrRev(p, "\")) & TEMPLATE_NAME
        If Dir(p) <> "" Then FindTemplate = p: Exit Function
    End If
    FindTemplate = PickFile("Title block template (*.txt)", "*.txt", "Select the title block template")
End Function

Private Function ReadTemplate(path As String) As Boolean
    Dim f As Integer, ln As String, k As Long, key As String, val As String
    nT = 0: overwrite = False: noProjection = False
    ReDim tLabel(0): ReDim tValue(0)
    Dim txt As String, lines As Variant, li As Long
    f = FreeFile
    Open path For Binary Access Read As #f            ' whole file; works with Windows or Unix line ends
    txt = Space(LOF(f))
    Get #f, , txt
    Close #f
    If Left(txt, 3) = Chr(239) & Chr(187) & Chr(191) Then txt = Mid(txt, 4)   ' UTF-8 marker
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For li = LBound(lines) To UBound(lines)
        ln = Trim(Replace(lines(li), vbTab, " "))
        k = InStr(ln, "=")
        If ln <> "" And Left(ln, 1) <> ";" And k > 1 Then
            key = Trim(Left(ln, k - 1)): val = Trim(Mid(ln, k + 1))
            If Right(key, 1) = ":" Then key = Trim(Left(key, Len(key) - 1))
            If UCase(key) = "OVERWRITE" Then
                overwrite = (LCase(val) = "yes" Or LCase(val) = "true" Or val = "1")
            ElseIf UCase(key) = "REMOVE PROJECTION" Then
                noProjection = (LCase(val) = "yes" Or LCase(val) = "true" Or val = "1")
            Else
                ReDim Preserve tLabel(nT): ReDim Preserve tValue(nT)
                tLabel(nT) = key: tValue(nT) = val
                nT = nT + 1
            End If
        End If
    Next li
    ReadTemplate = (nT > 0)
End Function

Private Sub ReadModelInfo()
    Dim v As Object, mdl As Object, cfg As String, db As String, mp As Object
    fFile = Mid(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") + 1)
    If InStrRev(fFile, ".") > 0 Then fFile = Left(fFile, InStrRev(fFile, ".") - 1)
    Set v = swModel.GetFirstView.GetNextView
    Do While Not v Is Nothing
        Set mdl = v.ReferencedDocument
        If Not mdl Is Nothing Then Exit Do
        Set v = v.GetNextView
    Loop
    If mdl Is Nothing Then Exit Sub
    fPart = Mid(mdl.GetPathName, InStrRev(mdl.GetPathName, "\") + 1)
    If InStrRev(fPart, ".") > 0 Then fPart = Left(fPart, InStrRev(fPart, ".") - 1)
    cfg = v.ReferencedConfiguration
    On Error Resume Next
    If mdl.GetType = 1 Then fMat = mdl.GetMaterialPropertyName2(cfg, db)    ' parts only
    Set mp = mdl.Extension.CreateMassProperty
    If Not mp Is Nothing Then fMass = Format(mp.Mass, "0.00")
    On Error GoTo 0
End Sub

Private Function Resolve(s As String, sheetNo As Long, sheetCount As Long) As String
    Dim r As String
    r = s
    r = Replace(r, "{partname}", fPart, , , vbTextCompare)
    r = Replace(r, "{material}", fMat, , , vbTextCompare)
    r = Replace(r, "{mass}", fMass, , , vbTextCompare)
    r = Replace(r, "{date}", Format(Date, "yyyy\/mm\/dd"), , , vbTextCompare)
    r = Replace(r, "{sheet}", Format(sheetNo, "00"), , , vbTextCompare)
    r = Replace(r, "{sheets}", Format(sheetCount, "00"), , , vbTextCompare)
    r = Replace(r, "{scale}", curScale, , , vbTextCompare)
    r = Replace(r, "{sheetsize}", curPaper, , , vbTextCompare)
    r = Replace(r, "{filename}", fFile, , , vbTextCompare)
    Resolve = r
End Function

Private Function PaperName(code As Variant) As String
    Select Case code
        Case 0, 1: PaperName = "A"
        Case 2: PaperName = "B"
        Case 3: PaperName = "C"
        Case 4: PaperName = "D"
        Case 5: PaperName = "E"
        Case 6, 7: PaperName = "A4"
        Case 8: PaperName = "A3"
        Case 9: PaperName = "A2"
        Case 10: PaperName = "A1"
        Case 11: PaperName = "A0"
        Case Else: PaperName = ""
    End Select
End Function

' ---------------------------------------------------------------------------------------
' Current sheet: fill each template label's box. Returns a report; adds to filled.
Private Function FillSheet(sheetNo As Long, sheetCount As Long, filled As Long) As String
    Dim sp As Variant, sw_ As Double, sh_ As Double
    sp = swModel.GetCurrentSheet.GetProperties2           ' read before EditTemplate: no current sheet while editing the format
    sw_ = sp(5) * 1000: sh_ = sp(6) * 1000
    curScale = CStr(sp(2)) & ":" & CStr(sp(3))
    curPaper = PaperName(sp(0))

    swModel.EditTemplate                                 ' title block notes live in the sheet format
    If noProjection Then
        If RemoveProjection(sw_) Then FillSheet = vbCrLf & "  Projection Method:  THIRD/FIRST ANGLE removed"
    End If

    Dim nts() As Object, tx() As Double, ty() As Double, tt() As String, isLab() As Boolean, cnt As Long
    ReDim nts(500): ReDim tx(500): ReDim ty(500): ReDim tt(500): ReDim isLab(500)
    Dim curNote As Object, p As Variant, s As String
    Set curNote = swModel.GetFirstView.GetFirstNote
    Do While Not curNote Is Nothing And cnt <= 500
        p = curNote.GetTextPoint2
        s = Trim(Replace(curNote.GetText, vbLf, ""))
        If p(0) * 1000 > 10 And p(0) * 1000 < sw_ - 10 And p(1) * 1000 > 10 And p(1) * 1000 < sh_ - 10 Then
            Set nts(cnt) = curNote
            tx(cnt) = p(0) * 1000: ty(cnt) = p(1) * 1000: tt(cnt) = s
            isLab(cnt) = (Len(s) > 1 And Right(s, 1) = ":" And InStr(s, vbCr) = 0)
            cnt = cnt + 1
        End If
        Set curNote = curNote.GetNext
    Loop

    Dim t As Long, i As Long, j As Long, best As Long, d As Double, bestD As Double, rep As String, val As String
    For t = 0 To nT - 1
        best = -2                                        ' -2: label not found
        For i = 0 To cnt - 1
            If isLab(i) Then
                If LCase(Trim(Left(tt(i), Len(tt(i)) - 1))) = LCase(tLabel(t)) Then
                    best = -1: bestD = 1E+99
                    For j = 0 To cnt - 1                 ' box just right of the heading
                        If IsBox(j, isLab, tt) Then
                            If tx(j) >= tx(i) + 2 And tx(j) <= tx(i) + 40 And Abs(ty(j) - ty(i)) <= 1.5 Then
                                d = tx(j) - tx(i)
                                If d < bestD Then bestD = d: best = j
                            End If
                        End If
                    Next j
                    If best < 0 Then                     ' else just below it
                        For j = 0 To cnt - 1
                            If IsBox(j, isLab, tt) Then
                                If tx(j) >= tx(i) - 10 And tx(j) <= tx(i) + 25 And ty(j) >= ty(i) - 10 And ty(j) <= ty(i) - 1.5 Then
                                    d = Sqr((tx(j) - tx(i)) ^ 2 + (ty(j) - ty(i)) ^ 2)
                                    If d < bestD Then bestD = d: best = j
                                End If
                            End If
                        Next j
                    End If
                    Exit For
                End If
            End If
        Next i

        If best = -2 Then
            rep = rep & vbCrLf & "  " & tLabel(t) & ":  no such heading in the title block"
        ElseIf best = -1 Then
            rep = rep & vbCrLf & "  " & tLabel(t) & ":  no box found next to the heading"
        Else
            val = Resolve(tValue(t), sheetNo, sheetCount)
            If tt(best) <> "" And Not overwrite Then
                rep = rep & vbCrLf & "  " & tLabel(t) & ":  kept """ & tt(best) & """"
                If FitToBox(nts(best), BoxRight(best, nts, ty, cnt, sw_)) Then rep = rep & " (made smaller to fit)"
            ElseIf val = "" Then
                rep = rep & vbCrLf & "  " & tLabel(t) & ":  no value (field empty in SOLIDWORKS)"
            Else
                nts(best).SetText val
                filled = filled + 1
                rep = rep & vbCrLf & "  " & tLabel(t) & ":  " & val
                If FitToBox(nts(best), BoxRight(best, nts, ty, cnt, sw_)) Then rep = rep & " (made smaller to fit)"
            End If
        End If
    Next t

    swModel.EditSheet
    FillSheet = FillSheet & rep
End Function

' a note that can be a value box: not a heading, one line, not a unit printed next to the box
Private Function IsBox(j As Long, isLab() As Boolean, tt() As String) As Boolean
    If isLab(j) Or InStr(tt(j), vbCr) > 0 Then Exit Function
    Select Case LCase(tt(j))
        Case "kg", "g", "mm", "m", "in", "lb"
            Exit Function
    End Select
    IsBox = True
End Function

' (sheet format open) "Projection Method:  THIRD ANGLE" -> "Projection Method:", and delete the
' projection symbol (lines/circles) drawn inside that box. True if anything was removed.
Private Function RemoveProjection(sw_ As Double) As Boolean
    Dim n As Object, lab As Object, s As String, k As Long, e As Variant, q As Variant
    Dim x0 As Double, x1 As Double, y1 As Double, p As Variant
    Set n = swModel.GetFirstView.GetFirstNote
    Do While Not n Is Nothing
        s = n.GetText
        If InStr(1, s, "Projection Method", vbTextCompare) > 0 Then
            Set lab = n
            If InStr(1, s, "THIRD ANGLE", vbTextCompare) > 0 Or InStr(1, s, "FIRST ANGLE", vbTextCompare) > 0 Then
                n.SetText "Projection Method:"
                RemoveProjection = True
            End If
        End If
        Set n = n.GetNext
    Loop
    If lab Is Nothing Then Exit Function

    ' the box: from the label's left edge to the next note on its right, below the label
    e = lab.GetExtent
    x0 = e(0) * 1000 - 3: y1 = e(1) * 1000: x1 = sw_ - 10
    p = lab.GetTextPoint2
    Set n = swModel.GetFirstView.GetFirstNote
    Do While Not n Is Nothing
        e = n.GetExtent: q = n.GetTextPoint2
        If Abs(q(1) - p(1)) * 1000 <= 4 And e(0) * 1000 > x0 + 6 And e(0) * 1000 < x1 Then x1 = e(0) * 1000
        Set n = n.GetNext
    Loop
    x1 = x1 - 1

    Dim sk As Object, segs As Variant, seg As Object, pts(3) As Double, np As Long, inside As Boolean, picked As Long, i As Long
    Set sk = swModel.GetActiveSketch2
    If sk Is Nothing Then Exit Function
    segs = sk.GetSketchSegments
    If IsEmpty(segs) Then Exit Function
    swModel.ClearSelection2 True
    For k = LBound(segs) To UBound(segs)
        Set seg = segs(k)
        np = 0
        Select Case seg.GetType
            Case 0                                       ' line
                pts(0) = seg.GetStartPoint2.X: pts(1) = seg.GetStartPoint2.Y
                pts(2) = seg.GetEndPoint2.X: pts(3) = seg.GetEndPoint2.Y: np = 2
            Case 1                                       ' arc / circle
                pts(0) = seg.GetCenterPoint2.X: pts(1) = seg.GetCenterPoint2.Y: np = 1
        End Select
        If np > 0 Then
            inside = True
            For i = 0 To np - 1
                If pts(2 * i) * 1000 <= x0 Or pts(2 * i) * 1000 >= x1 Or pts(2 * i + 1) * 1000 <= 2 Or pts(2 * i + 1) * 1000 >= y1 Then inside = False
            Next i
            If inside Then
                seg.Select4 picked > 0, Nothing
                picked = picked + 1
            End If
        End If
    Next k
    If picked > 0 Then
        swModel.EditDelete
        RemoveProjection = True
    End If
    swModel.ClearSelection2 True
End Function

' right edge (mm) of box b: just before the nearest note to its right on the same row
' (the next box or its heading, by its left edge), else the sheet border
Private Function BoxRight(b As Long, nts() As Object, ty() As Double, cnt As Long, sw_ As Double) As Double
    Dim k As Long, r As Double, e As Variant, x0 As Double, xk As Double
    r = sw_ - 10
    e = nts(b).GetExtent: x0 = e(0) * 1000
    For k = 0 To cnt - 1
        If k <> b And Abs(ty(k) - ty(b)) <= 4 Then
            e = nts(k).GetExtent: xk = e(0) * 1000
            If xk > x0 + 2 And xk < r Then r = xk
        End If
    Next k
    BoxRight = r - 1.5                                   ' keep clear of the box line
End Function

' make the note's text smaller if it runs past rightX (mm). True if it was changed.
Private Function FitToBox(boxNote As Object, rightX As Double) As Boolean
    Dim e As Variant, x0 As Double, x1 As Double, ann As Object, tf As Object, h As Double, f As Double
    e = boxNote.GetExtent
    If IsEmpty(e) Then Exit Function
    x0 = e(0) * 1000: x1 = e(3) * 1000
    If x1 <= rightX Or x1 - x0 < 0.5 Or rightX <= x0 Then Exit Function
    Set ann = boxNote.GetAnnotation
    Set tf = ann.GetTextFormat(0)
    h = tf.CharHeight
    f = (rightX - x0) / (x1 - x0) * 0.97
    If h * f < 0.001 Then f = 0.001 / h                  ' never below 1 mm
    tf.CharHeight = h * f
    ann.SetTextFormat 0, False, tf
    FitToBox = True
End Function

' ---------------------------------------------------------------------------------------
Private Function PickFile(filterName As String, pattern As String, dlgTitle As String) As String
    Dim ofn As OPENFILENAME
    ofn.lStructSize = LenB(ofn)
    On Error Resume Next
    ofn.hwndOwner = swApp.Frame.GetHWndx64
    On Error GoTo 0
    ofn.lpstrFilter = filterName & vbNullChar & pattern & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = String(4096, vbNullChar)
    ofn.nMaxFile = Len(ofn.lpstrFile)
    ofn.lpstrTitle = dlgTitle
    ofn.flags = &H80000 Or &H1000 Or &H800 Or &H4        ' explorer, must exist, hide read-only
    If GetOpenFileName(ofn) = 0 Then Exit Function
    PickFile = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
End Function
