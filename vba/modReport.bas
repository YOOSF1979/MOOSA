Attribute VB_Name = "modReport"
Option Compare Database
Option Explicit

' ===========================================================
'  ﬁ—Ì— ›Õ’ Ê„⁄«·Ã… FSD 75-FPD-FRM-21 ·„Ã„Ê⁄… Ê«Õœ… (Word)
' «·≈’œ«— 1 - 22 ‰Ê›„»— 2026
'
' «·ﬁ«·»: <„Ã·œ «·»—‰«„Ã>\assets\report_template.docx  (Ìı»‰Ï »‹ src\new\make_template.ps1)
' «·‰« Ã: <„Ã·œ «·»—‰«„Ã>\«· ﬁ«—Ì—\<—ﬁ„ «·„—Ã⁄>-„<—ﬁ„ «·„Ã„Ê⁄…>.docx
'
' - «·⁄·«„«  ´...ª ›Ì «·ﬁ«·»  ı” »œ·. ÃœÊ· «·‰ «∆Ã Ìı⁄—Û› »‰’ ´‰Ê⁄ «·„⁄«·Ã…ª ·« »—ﬁ„Â.
' - √”ÿ— «·«‰ ﬁ«· ·«  ŸÂ— ›Ì «· ﬁ—Ì—.
' - «·√⁄„œ…: „ | ‰Ê⁄ | ⁄œœ | ‰Ê⁄ «·„⁄«·Ã… | ≈ÌÃ«»Ì… | ”·»Ì… |  ⁄–—
'     ≈ÌÃ«»Ì… = ⁄œœ «·¬À«—. ”·»Ì… Ê ⁄–— = ⁄·«„…° √Ê «·⁄œœ ≈–« «Œ ·ÿ  «·‰ «∆Ã ›Ì «·⁄Ì‰….
' -  «—ÌŒ «·›Õ’ Ê«· ’œÌ—: ≈‰ ﬂ«‰« ›«—€Ì‰ Ìı„·¬‰ » «—ÌŒ «·ÌÊ„.
' - Word »—»ÿ „ √Œ— (CreateObject): ·« „—Ã⁄ „ﬂ »…° ›·« „‘«ﬂ· ≈’œ«—« .
' ===========================================================

Private Const MIN_ROWS As Long = 6          ' √ﬁ· ⁄œœ √”ÿ— »Ì«‰«  ›Ì «·ÃœÊ·° ﬂ«·‰„Ê–Ã «·Ê—ﬁÌ
Private Const PHOTO_WIDTH_CM As Double = 8.5
Private Const SIG_HEIGHT_CM As Double = 1.4

' “— ´«· ﬁ—Ì—ª ›Ì ‘«‘… «·⁄Ì‰« 
Public Function MakeReport(frm As Form) As Boolean
    Dim batchID As Long
    If Not SaveRecord(frm) Then Exit Function
    batchID = CLng(Nz(frm.OpenArgs, 0))
    If batchID = 0 Then Exit Function
    MakeReport = (Len(BuildReport(batchID, True)) > 0)
End Function

' Ì»‰Ì «· ﬁ—Ì— ÊÌ⁄Ìœ „”«— «·„·›° √Ê "" ≈–« —ı›÷. showWord = False ··«Œ »«— «·¬·Ì.
Public Function BuildReport(ByVal batchID As Long, Optional ByVal showWord As Boolean = True) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object, wd As Object, doc As Object
    Dim refNo As String, tplPath As String, outDir As String, outFile As String
    Dim examDate As Date, exportDate As Date
    Dim examinerID As Long, reviewerID As Long, approverID As Long

    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ---------- «·»Ì«‰«  Ê«· Õﬁﬁ ----------
    Set rs = db.OpenRecordset( _
        "SELECT c.RefNo, c.ForensicNo, c.CaseDate, a.AgencyName, b.BatchNo, b.ReceivedDate, b.ExamDate, b.ExportDate, " & _
        "b.ExaminerID, b.ReviewerID, b.ApproverID " & _
        "FROM (tblBatches AS b INNER JOIN tblCases AS c ON b.CaseID = c.CaseID) " & _
        "LEFT JOIN tblAgencies AS a ON c.AgencyID = a.AgencyID WHERE b.BatchID=" & batchID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    If DbFirst(db, "SELECT Count(*) FROM tblSamples WHERE BatchID=" & batchID & " AND LineType=1") = 0 Then
        rs.Close
        Msg "·«  ÊÃœ ⁄Ì‰«  „⁄«·Ã… ›Ì Â–Â «·„Ã„Ê⁄…. √”ÿ— «·«‰ ﬁ«· ·«  ŸÂ— ›Ì «· ﬁ—Ì—.", vbExclamation
        Exit Function
    End If
    If IsNull(rs!reviewerID) Or IsNull(rs!approverID) Then
        rs.Close
        Msg "«Œ — «·ﬁ«∆„ »«·„—«Ã⁄… Ê«·„⁄ „œ ··„Ã„Ê⁄… ›Ì «·‘«‘… «·—∆Ì”Ì… √Ê·«.", vbExclamation
        Exit Function
    End If

    refNo = rs!refNo
    examinerID = rs!examinerID: reviewerID = rs!reviewerID: approverID = rs!approverID
    examDate = Nz(rs!examDate, Date)
    exportDate = Nz(rs!exportDate, Date)
    If exportDate < examDate Then exportDate = examDate

    tplPath = CurrentProject.path & "\assets\report_template.docx"
    If Not fso.FileExists(tplPath) Then
        rs.Close
        Msg "ﬁ«·» «· ﬁ—Ì— €Ì— „ÊÃÊœ:" & vbCrLf & tplPath, vbCritical
        Exit Function
    End If
    outDir = CurrentProject.path & "\«· ﬁ«—Ì—"
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    outFile = outDir & "\" & refNo & "-„" & rs!BatchNo & ".docx"
    If fso.FileExists(outFile) Then
        If Not Ask("«· ﬁ—Ì— „ÊÃÊœ „”»ﬁ«:" & vbCrLf & outFile & vbCrLf & "«” »œ«·Â »‰”Œ… ÃœÌœ…ø") Then
            rs.Close
            If showWord Then OpenInWord outFile
            BuildReport = outFile
            Exit Function
        End If
        If Not TryDelete(fso, outFile) Then
            rs.Close
            Msg "«· ﬁ—Ì— „› ÊÕ ›Ì Word. √€·ﬁÂ À„ √⁄œ «·„Õ«Ê·….", vbExclamation
            Exit Function
        End If
    End If

    '  «—ÌŒ «·›Õ’ Ê«· ’œÌ— ÌıÀ»Û¯ «‰ ›Ì «·„Ã„Ê⁄…
    db.Execute "UPDATE tblBatches SET ExamDate=" & SqlDate(examDate) & ", ExportDate=" & SqlDate(exportDate) & _
               " WHERE BatchID=" & batchID, dbFailOnError

    ' ---------- Word ----------
    On Error GoTo EH
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    Set doc = wd.Documents.Add(tplPath)

    RepAll doc, "´RefNoª", refNo
    RepAll doc, "´Agencyª", Nz(rs!AgencyName, "")
    ' —ﬁ„ «·√œ·…: «·Œ«‰… ›ÌÂ« «·—ﬁ„ ›ﬁÿ° Ê«· ﬁ—Ì— Ìﬂ »Â 312312/2026-Ÿ œ«Œ· «·›ﬁ—… «·⁄—»Ì…
    RepAll doc, "´ForensicNoª", ForensicText(rs!ForensicNo, rs!caseDate)
    ' «· Ê«—ÌŒ  ıﬂ » „»«‘—… »« Ã«Â „‰ «·Ì”«— Õ Ï ·«  ‰ﬁ·»: 2026/11/19
    PutText wd, doc.Content, "´ReceivedDateª", Format(rs!ReceivedDate, "yyyy/mm/dd")
    PutText wd, doc.Content, "´ExamDateª", Format(examDate, "yyyy/mm/dd")
    PutText wd, doc.Content, "´ExportDateª", Format(exportDate, "yyyy/mm/dd")
    RepAll doc, "´ExaminerNameª", PersonLine(db, examinerID)
    RepAll doc, "´ReviewerNameª", PersonLine(db, reviewerID)
    RepAll doc, "´ApproverNameª", PersonLine(db, approverID)
    rs.Close

    FillResults wd, doc, db, batchID
    ' «· Ê«ﬁÌ⁄: SignaturePath √Ê assets\signatures\<«”„ «·œŒÊ·> (modApp.SignatureFile)
    Dim sigMissing As String
    sigMissing = PlaceSignature(wd, doc.Content, "´SigExaminerª", examinerID, SIG_HEIGHT_CM) & _
                 PlaceSignature(wd, doc.Content, "´SigReviewerª", reviewerID, SIG_HEIGHT_CM) & _
                 PlaceSignature(wd, doc.Content, "´SigApproverª", approverID, SIG_HEIGHT_CM)
    PutSignature wd, doc.Sections(1).Footers(1).Range, "´SigFooterª", SignatureFile(examinerID), 0.9
    FillPhotos wd, doc, db, fso, batchID, refNo

    doc.SaveAs2 outFile, 16                     ' wdFormatXMLDocument
    If showWord Then
        wd.Visible = True
        doc.Activate
    Else
        doc.Close 0
        wd.Quit 0
    End If
    BuildReport = outFile
    If Len(sigMissing) > 0 Then
        Msg "√ı‰‘∆ «· ﬁ—Ì—° ·ﬂ‰ ·«  ÊÃœ ’Ê—…  ÊﬁÌ⁄ ·‹:" & vbCrLf & sigMissing & _
            "√÷›Â« „‰ ‘«‘… ´«·„” Œœ„Ê‰ª (“— ´«Œ Ì«—ª »Ã«‰» ﬂ· „ÊŸ›) À„ √⁄œ ≈‰‘«¡ «· ﬁ—Ì—.", vbExclamation
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "Œÿ√ " & Err.Number & " √À‰«¡ ≈‰‘«¡ «· ﬁ—Ì—: " & Err.Description
    If Not doc Is Nothing Then doc.Close 0
    If Not wd Is Nothing Then wd.Quit 0
    Msg errText, vbCritical
End Function

' ===========================================================
' ÃœÊ· «·‰ «∆Ã
' ===========================================================
Private Sub FillResults(wd As Object, doc As Object, db As DAO.Database, ByVal batchID As Long)
    Dim tbl As Object, rs As DAO.Recordset
    Dim nSamples As Long, needRows As Long, haveRows As Long, r As Long, k As Long
    Dim cellRng As Object

    Set tbl = FindTable(doc, "‰Ê⁄ «·„⁄«·Ã…")
    If tbl Is Nothing Then Err.Raise vbObjectError + 10, , "ÃœÊ· «·‰ «∆Ã €Ì— „ÊÃÊœ ›Ì «·ﬁ«·» (Ìı⁄—› »⁄»«—… ´‰Ê⁄ «·„⁄«·Ã…ª)."

    Set rs = db.OpenRecordset( _
        "SELECT Barcode, Description, ItemCount, MethodCode, PosCount, NegCount, UnkCount, TraceCount " & _
        "FROM tblSamples WHERE BatchID=" & batchID & " AND LineType=1 ORDER BY SampleNo", dbOpenSnapshot)
    rs.MoveLast: nSamples = rs.RecordCount: rs.MoveFirst

    ' ”ÿ—« «·—√” À«» «‰∫ «·»Ì«‰«  „‰ «·”ÿ— 3. ﬂ· ⁄Ì‰… ”ÿ—«‰: «·»«—ﬂÊœ À„ «·»Ì«‰« 
    needRows = 2 * nSamples
    If needRows < MIN_ROWS Then needRows = MIN_ROWS
    haveRows = tbl.rows.Count - 2
    If needRows > haveRows Then
        tbl.Cell(tbl.rows.Count, 1).Range.Select
        wd.Selection.InsertRowsBelow needRows - haveRows
    End If

    ' Œÿ „ÊÕœ ·√”ÿ— «·»Ì«‰«  («·√—ﬁ«„ Ê«·⁄—»Ì…  √Œ– ÕÃ„ ´Biª ›Ì «·›ﬁ—«  „‰ «·Ì„Ì‰)
    For r = 3 To tbl.rows.Count
        For k = 1 To 7
            With tbl.Cell(r, k).Range.Font
                .Size = 12: .SizeBi = 12
            End With
        Next
    Next

    r = 3
    k = 0
    Do Until rs.EOF
        k = k + 1
        ' ”ÿ— «·»«—ﬂÊœ
        Set cellRng = tbl.Cell(r, 2).Range
        cellRng.text = "»«—ﬂÊœ —ﬁ„ (" & LtrText(rs!Barcode) & ")"
        cellRng.Font.Bold = True: cellRng.Font.BoldBi = True
        cellRng.Font.Underline = 1
        ' ”ÿ— «·»Ì«‰« 
        SetText tbl, r + 1, 1, CStr(k)
        SetText tbl, r + 1, 2, Nz(rs!Description, "")
        SetText tbl, r + 1, 3, CStr(rs!ItemCount)
        SetText tbl, r + 1, 4, Nz(rs!MethodCode, "")
        If rs!PosCount > 0 Then SetText tbl, r + 1, 5, CStr(Nz(rs!TraceCount, 0))
        SetMark tbl, r + 1, 6, rs!NegCount, rs!ItemCount
        SetMark tbl, r + 1, 7, rs!UnkCount, rs!ItemCount
        r = r + 2
        rs.MoveNext
    Loop
    rs.Close
End Sub

' ⁄·«„… ≈–« ﬂ«‰  «·‰ ÌÃ… ·ﬂ· «·⁄Ì‰…° Ê«·⁄œœ ≈–« «Œ ·ÿ  «·‰ «∆Ã
Private Sub SetMark(tbl As Object, ByVal r As Long, ByVal c As Long, ByVal n As Long, ByVal total As Long)
    Dim rng As Object
    If n <= 0 Then Exit Sub
    If n = total Then
        Set rng = tbl.Cell(r, c).Range
        rng.text = ChrW(&H2713)
        rng.Font.Name = "Segoe UI Symbol"
        rng.Font.Size = 14
    Else
        SetText tbl, r, c, CStr(n)
    End If
End Sub

Private Sub SetText(tbl As Object, ByVal r As Long, ByVal c As Long, ByVal s As String)
    tbl.Cell(r, c).Range.text = s
End Sub

' √Ê· ÃœÊ· ÌÕ ÊÌ «·‰’ «·„⁄ÿÏ
Private Function FindTable(doc As Object, ByVal label As String) As Object
    Dim t As Object
    For Each t In doc.Tables
        If InStr(1, t.Range.text, label) > 0 Then
            Set FindTable = t
            Exit Function
        End If
    Next
End Function

' ===========================================================
' ’›Õ«  «·’Ê—: ’Ê— «·⁄Ì‰«  À„ ’Ê— ¬À«—Â«° ’Ê— «‰ ›Ì ﬂ· ”ÿ— „⁄ «· ”„Ì…
' ===========================================================
Private Sub FillPhotos(wd As Object, doc As Object, db As DAO.Database, fso As Object, ByVal batchID As Long, ByVal refNo As String)
    Dim rs As DAO.Recordset
    Dim paths As New Collection, captions As New Collection
    Dim head As Object, spot As Object, tbl As Object, shp As Object, rng As Object
    Dim i As Long, rows As Long, r As Long, c As Long

    ' «·⁄Ì‰« 
    Set rs = db.OpenRecordset("SELECT Barcode, PhotoPath FROM tblSamples WHERE BatchID=" & batchID & _
                              " AND LineType=1 AND PhotoPath Is Not Null ORDER BY SampleNo", dbOpenSnapshot)
    Do Until rs.EOF
        If fso.FileExists(rs!PhotoPath) Then paths.Add CStr(rs!PhotoPath): captions.Add "«·⁄Ì‰… " & rs!Barcode
        rs.MoveNext
    Loop
    rs.Close
    ' «·¬À«—
    Set rs = db.OpenRecordset("SELECT t.TraceCode, t.PhotoPath FROM tblTraces AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID " & _
                              "WHERE s.BatchID=" & batchID & " AND s.LineType=1 AND t.PhotoPath Is Not Null " & _
                              "ORDER BY s.SampleNo, t.TraceNo", dbOpenSnapshot)
    Do Until rs.EOF
        If fso.FileExists(rs!PhotoPath) Then paths.Add CStr(rs!PhotoPath): captions.Add "«·√À— " & LtrText(rs!TraceCode)
        rs.MoveNext
    Loop
    rs.Close

    Set head = FindIn(doc.Content, "´PhotosHeadª")
    Set spot = FindIn(doc.Content, "´Photosª")
    If head Is Nothing Or spot Is Nothing Then Exit Sub

    If paths.Count = 0 Then
        ' ·« ’Ê—: ÌıÕ–› ﬁ”„ «·’Ê— ﬂ·Â° „‰ ‰Â«Ì… „« ﬁ»·Â ≈·Ï ¬Œ— «·„” ‰œ
        Set rng = doc.Range(head.Paragraphs(1).Range.Start - 1, doc.Content.End)
        rng.Delete
        Exit Sub
    End If

    head.text = "—ﬁ„ «·„—Ã⁄: " & refNo
    spot.text = ""
    rows = (paths.Count + 1) \ 2
    Set tbl = doc.Tables.Add(spot, rows * 2, 2)
    tbl.TableDirection = 0
    tbl.Borders.Enable = 0
    For i = 1 To paths.Count
        r = ((i - 1) \ 2) * 2 + 1
        c = ((i - 1) Mod 2) + 1
        Set rng = tbl.Cell(r, c).Range
        rng.ParagraphFormat.Alignment = 1
        Set shp = doc.InlineShapes.AddPicture(paths(i), False, True, rng)
        shp.LockAspectRatio = True
        shp.Width = wd.CentimetersToPoints(PHOTO_WIDTH_CM)
        If shp.Height > wd.CentimetersToPoints(10) Then shp.Height = wd.CentimetersToPoints(10)
        Set rng = tbl.Cell(r + 1, c).Range
        rng.text = captions(i)
        rng.Font.Size = 11: rng.Font.SizeBi = 11
        rng.ParagraphFormat.Alignment = 1
        rng.ParagraphFormat.SpaceAfter = 10
    Next
End Sub

' ===========================================================
' √œÊ« 
' ===========================================================
' Ì” »œ· ›Ì ﬂ· √Ã“«¡ «·„” ‰œ («·„ ‰ Ê«·—√” Ê«· –ÌÌ·)
Private Sub RepAll(doc As Object, ByVal findText As String, ByVal repl As String)
    Dim story As Object
    For Each story In doc.StoryRanges
        Do
            With story.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .text = findText
                .Replacement.text = repl
                .Forward = True
                .Wrap = 1
                .MatchWildcards = False
                .Execute Replace:=2
            End With
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next
End Sub

Private Function FindIn(scope As Object, ByVal txt As String) As Object
    Dim r As Object
    Set r = scope.Duplicate
    With r.Find
        .ClearFormatting
        .text = txt
        .Forward = True
        .Wrap = 0
        .MatchWildcards = False
        If .Execute Then Set FindIn = r
    End With
End Function

' Ìﬂ » «·‰’ „ﬂ«‰ «·⁄·«„… ÊÌÃ⁄· « Ã«ÂÂ „‰ «·Ì”«— (·· Ê«—ÌŒ Ê«·√—ﬁ«„ –«  «·›Ê«’·)
' LtrRun „ «Õ ⁄·Ï Selection ›ﬁÿ («·Œÿ√ 438 ⁄·Ï Range)
Private Sub PutText(wd As Object, scope As Object, ByVal mark As String, ByVal txt As String)
    Dim r As Object
    Set r = FindIn(scope, mark)
    If r Is Nothing Then Exit Sub
    r.text = txt
    r.Select
    wd.Selection.LtrRun
End Sub

' Ì÷⁄ ’Ê—… «· ÊﬁÌ⁄ „ﬂ«‰ «·⁄·«„…° √Ê Ì„”Õ «·⁄·«„… ≈–« ·«  ÊﬁÌ⁄
Private Sub PutSignature(wd As Object, scope As Object, ByVal mark As String, ByVal imgPath As String, ByVal heightCm As Double)
    Dim r As Object, shp As Object
    Set r = FindIn(scope, mark)
    If r Is Nothing Then Exit Sub
    r.text = ""
    If Len(imgPath) = 0 Then Exit Sub
    Set shp = r.Document.InlineShapes.AddPicture(imgPath, False, True, r)
    shp.LockAspectRatio = True
    shp.Height = wd.CentimetersToPoints(heightCm)
End Sub

' «· Ê«—ÌŒ Ê«·—„Ê“ œ«Œ· ›ﬁ—… ⁄—»Ì…  ‰ﬁ·» (2026/11/19  ŸÂ— 19/11/2026):
'  ıÕ«ÿ »⁄·«„ Ì ´ ÷„Ì‰ „‰ «·Ì”«—ª Õ Ï  »ﬁÏ ﬂ„« ﬂı » 
Private Function LtrText(ByVal s As String) As String
    LtrText = ChrW(&H202A) & s & ChrW(&H202C)
End Function

' ´«·— »…/ «·«”„ª ﬂ„« ›Ì «·‰„Ê–Ã
Private Function PersonLine(db As DAO.Database, ByVal userID As Long) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT RankTitle, FullName FROM tblUsers WHERE UserID=" & userID, dbOpenSnapshot)
    If Not rs.EOF Then
        If Len(Nz(rs!RankTitle, "")) > 0 Then PersonLine = rs!RankTitle & "/ " & rs!FullName Else PersonLine = rs!FullName
    End If
    rs.Close
End Function

' Ì÷⁄  ÊﬁÌ⁄ «·„” Œœ„ „ﬂ«‰ «·⁄·«„…∫ Ì⁄Ìœ «”„Â (”ÿ—«) ≈–« ·„  ÊÃœ ’Ê—…  ÊﬁÌ⁄Â
Private Function PlaceSignature(wd As Object, scope As Object, ByVal mark As String, ByVal userID As Long, ByVal heightCm As Double) As String
    Dim p As String
    p = SignatureFile(userID)
    PutSignature wd, scope, mark, p, heightCm
    If Len(p) = 0 Then PlaceSignature = "- " & Nz(DLookup("FullName", "tblUsers", "UserID=" & userID), "") & vbCrLf
End Function

Private Function DbFirst(db As DAO.Database, ByVal sql As String) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then DbFirst = rs.Fields(0).Value
    rs.Close
End Function

Private Function SqlDate(ByVal d As Date) As String
    SqlDate = "#" & Format(d, "yyyy-mm-dd") & "#"
End Function

Private Function TryDelete(fso As Object, ByVal path As String) As Boolean
    On Error GoTo Locked
    fso.DeleteFile path, True
    TryDelete = True
    Exit Function
Locked:
End Function

Private Sub OpenInWord(ByVal path As String)
    Dim wd As Object
    Set wd = CreateObject("Word.Application")
    wd.Documents.Open path
    wd.Visible = True
End Sub
