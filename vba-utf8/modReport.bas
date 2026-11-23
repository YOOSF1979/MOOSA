Attribute VB_Name = "modReport"
Option Compare Database
Option Explicit

' ===========================================================
' تقرير فحص ومعالجة FSD 75-FPD-FRM-21 لمجموعة واحدة (Word)
' الإصدار 1 - 22 نوفمبر 2026
'
' القالب: <مجلد البرنامج>\assets\report_template.docx  (يُبنى بـ src\new\make_template.ps1)
' الناتج: <مجلد البرنامج>\التقارير\<رقم المرجع>-م<رقم المجموعة>.docx
'
' - العلامات «...» في القالب تُستبدل. جدول النتائج يُعرَف بنص «نوع المعالجة» لا برقمه.
' - أسطر الانتقال لا تظهر في التقرير.
' - الأعمدة: م | نوع | عدد | نوع المعالجة | إيجابية | سلبية | تعذر
'     إيجابية = عدد الآثار. سلبية وتعذر = علامة، أو العدد إذا اختلطت النتائج في العينة.
' - تاريخ الفحص والتصدير: إن كانا فارغين يُملآن بتاريخ اليوم.
' - Word بربط متأخر (CreateObject): لا مرجع مكتبة، فلا مشاكل إصدارات.
' ===========================================================

Private Const MIN_ROWS As Long = 6          ' أقل عدد أسطر بيانات في الجدول، كالنموذج الورقي
Private Const PHOTO_WIDTH_CM As Double = 8.5
Private Const SIG_HEIGHT_CM As Double = 1.4

' زر «التقرير» في شاشة العينات
Public Function MakeReport(frm As Form) As Boolean
    Dim batchID As Long
    If Not SaveRecord(frm) Then Exit Function
    batchID = CLng(Nz(frm.OpenArgs, 0))
    If batchID = 0 Then Exit Function
    MakeReport = (Len(BuildReport(batchID, True)) > 0)
End Function

' يبني التقرير ويعيد مسار الملف، أو "" إذا رُفض. showWord = False للاختبار الآلي.
Public Function BuildReport(ByVal batchID As Long, Optional ByVal showWord As Boolean = True) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object, wd As Object, doc As Object
    Dim refNo As String, tplPath As String, outDir As String, outFile As String
    Dim examDate As Date, exportDate As Date
    Dim examinerID As Long, reviewerID As Long, approverID As Long

    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ---------- البيانات والتحقق ----------
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
        Msg "لا توجد عينات معالجة في هذه المجموعة. أسطر الانتقال لا تظهر في التقرير.", vbExclamation
        Exit Function
    End If
    If IsNull(rs!reviewerID) Or IsNull(rs!approverID) Then
        rs.Close
        Msg "اختر القائم بالمراجعة والمعتمد للمجموعة في الشاشة الرئيسية أولاً.", vbExclamation
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
        Msg "قالب التقرير غير موجود:" & vbCrLf & tplPath, vbCritical
        Exit Function
    End If
    outDir = CurrentProject.path & "\التقارير"
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    outFile = outDir & "\" & refNo & "-م" & rs!BatchNo & ".docx"
    If fso.FileExists(outFile) Then
        If Not Ask("التقرير موجود مسبقاً:" & vbCrLf & outFile & vbCrLf & "استبداله بنسخة جديدة؟") Then
            rs.Close
            If showWord Then OpenInWord outFile
            BuildReport = outFile
            Exit Function
        End If
        If Not TryDelete(fso, outFile) Then
            rs.Close
            Msg "التقرير مفتوح في Word. أغلقه ثم أعد المحاولة.", vbExclamation
            Exit Function
        End If
    End If

    ' تاريخ الفحص والتصدير يُثبَّتان في المجموعة
    db.Execute "UPDATE tblBatches SET ExamDate=" & SqlDate(examDate) & ", ExportDate=" & SqlDate(exportDate) & _
               " WHERE BatchID=" & batchID, dbFailOnError

    ' ---------- Word ----------
    On Error GoTo EH
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    Set doc = wd.Documents.Add(tplPath)

    RepAll doc, "«RefNo»", refNo
    RepAll doc, "«Agency»", Nz(rs!AgencyName, "")
    ' رقم الأدلة: الخانة فيها الرقم فقط، والتقرير يكتبه 312312/2026-ظ داخل الفقرة العربية
    RepAll doc, "«ForensicNo»", ForensicText(rs!ForensicNo, rs!caseDate)
    ' التواريخ تُكتب مباشرة باتجاه من اليسار حتى لا تنقلب: 2026/11/19
    PutText wd, doc.Content, "«ReceivedDate»", Format(rs!ReceivedDate, "yyyy/mm/dd")
    PutText wd, doc.Content, "«ExamDate»", Format(examDate, "yyyy/mm/dd")
    PutText wd, doc.Content, "«ExportDate»", Format(exportDate, "yyyy/mm/dd")
    RepAll doc, "«ExaminerName»", PersonLine(db, examinerID)
    RepAll doc, "«ReviewerName»", PersonLine(db, reviewerID)
    RepAll doc, "«ApproverName»", PersonLine(db, approverID)
    rs.Close

    FillResults wd, doc, db, batchID
    ' التواقيع: SignaturePath أو assets\signatures\<اسم الدخول> (modApp.SignatureFile)
    Dim sigMissing As String
    sigMissing = PlaceSignature(wd, doc.Content, "«SigExaminer»", examinerID, SIG_HEIGHT_CM) & _
                 PlaceSignature(wd, doc.Content, "«SigReviewer»", reviewerID, SIG_HEIGHT_CM) & _
                 PlaceSignature(wd, doc.Content, "«SigApprover»", approverID, SIG_HEIGHT_CM)
    PutSignature wd, doc.Sections(1).Footers(1).Range, "«SigFooter»", SignatureFile(examinerID), 0.9
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
        Msg "أُنشئ التقرير، لكن لا توجد صورة توقيع لـ:" & vbCrLf & sigMissing & _
            "أضفها من شاشة «المستخدمون» (زر «اختيار» بجانب كل موظف) ثم أعد إنشاء التقرير.", vbExclamation
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ " & Err.Number & " أثناء إنشاء التقرير: " & Err.Description
    If Not doc Is Nothing Then doc.Close 0
    If Not wd Is Nothing Then wd.Quit 0
    Msg errText, vbCritical
End Function

' ===========================================================
' جدول النتائج
' ===========================================================
Private Sub FillResults(wd As Object, doc As Object, db As DAO.Database, ByVal batchID As Long)
    Dim tbl As Object, rs As DAO.Recordset
    Dim nSamples As Long, needRows As Long, haveRows As Long, r As Long, k As Long
    Dim cellRng As Object

    Set tbl = FindTable(doc, "نوع المعالجة")
    If tbl Is Nothing Then Err.Raise vbObjectError + 10, , "جدول النتائج غير موجود في القالب (يُعرف بعبارة «نوع المعالجة»)."

    Set rs = db.OpenRecordset( _
        "SELECT Barcode, Description, ItemCount, MethodCode, PosCount, NegCount, UnkCount, TraceCount " & _
        "FROM tblSamples WHERE BatchID=" & batchID & " AND LineType=1 ORDER BY SampleNo", dbOpenSnapshot)
    rs.MoveLast: nSamples = rs.RecordCount: rs.MoveFirst

    ' سطرا الرأس ثابتان؛ البيانات من السطر 3. كل عينة سطران: الباركود ثم البيانات
    needRows = 2 * nSamples
    If needRows < MIN_ROWS Then needRows = MIN_ROWS
    haveRows = tbl.rows.Count - 2
    If needRows > haveRows Then
        tbl.Cell(tbl.rows.Count, 1).Range.Select
        wd.Selection.InsertRowsBelow needRows - haveRows
    End If

    ' خط موحد لأسطر البيانات (الأرقام والعربية تأخذ حجم «Bi» في الفقرات من اليمين)
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
        ' سطر الباركود
        Set cellRng = tbl.Cell(r, 2).Range
        cellRng.text = "باركود رقم (" & LtrText(rs!Barcode) & ")"
        cellRng.Font.Bold = True: cellRng.Font.BoldBi = True
        cellRng.Font.Underline = 1
        ' سطر البيانات
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

' علامة إذا كانت النتيجة لكل العينة، والعدد إذا اختلطت النتائج
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

' أول جدول يحتوي النص المعطى
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
' صفحات الصور: صور العينات ثم صور آثارها، صورتان في كل سطر مع التسمية
' ===========================================================
Private Sub FillPhotos(wd As Object, doc As Object, db As DAO.Database, fso As Object, ByVal batchID As Long, ByVal refNo As String)
    Dim rs As DAO.Recordset
    Dim paths As New Collection, captions As New Collection
    Dim head As Object, spot As Object, tbl As Object, shp As Object, rng As Object
    Dim i As Long, rows As Long, r As Long, c As Long

    ' العينات
    Set rs = db.OpenRecordset("SELECT Barcode, PhotoPath FROM tblSamples WHERE BatchID=" & batchID & _
                              " AND LineType=1 AND PhotoPath Is Not Null ORDER BY SampleNo", dbOpenSnapshot)
    Do Until rs.EOF
        If fso.FileExists(rs!PhotoPath) Then paths.Add CStr(rs!PhotoPath): captions.Add "العينة " & rs!Barcode
        rs.MoveNext
    Loop
    rs.Close
    ' الآثار
    Set rs = db.OpenRecordset("SELECT t.TraceCode, t.PhotoPath FROM tblTraces AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID " & _
                              "WHERE s.BatchID=" & batchID & " AND s.LineType=1 AND t.PhotoPath Is Not Null " & _
                              "ORDER BY s.SampleNo, t.TraceNo", dbOpenSnapshot)
    Do Until rs.EOF
        If fso.FileExists(rs!PhotoPath) Then paths.Add CStr(rs!PhotoPath): captions.Add "الأثر " & LtrText(rs!TraceCode)
        rs.MoveNext
    Loop
    rs.Close

    Set head = FindIn(doc.Content, "«PhotosHead»")
    Set spot = FindIn(doc.Content, "«Photos»")
    If head Is Nothing Or spot Is Nothing Then Exit Sub

    If paths.Count = 0 Then
        ' لا صور: يُحذف قسم الصور كله، من نهاية ما قبله إلى آخر المستند
        Set rng = doc.Range(head.Paragraphs(1).Range.Start - 1, doc.Content.End)
        rng.Delete
        Exit Sub
    End If

    head.text = "رقم المرجع: " & refNo
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
' أدوات
' ===========================================================
' يستبدل في كل أجزاء المستند (المتن والرأس والتذييل)
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

' يكتب النص مكان العلامة ويجعل اتجاهه من اليسار (للتواريخ والأرقام ذات الفواصل)
' LtrRun متاح على Selection فقط (الخطأ 438 على Range)
Private Sub PutText(wd As Object, scope As Object, ByVal mark As String, ByVal txt As String)
    Dim r As Object
    Set r = FindIn(scope, mark)
    If r Is Nothing Then Exit Sub
    r.text = txt
    r.Select
    wd.Selection.LtrRun
End Sub

' يضع صورة التوقيع مكان العلامة، أو يمسح العلامة إذا لا توقيع
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

' التواريخ والرموز داخل فقرة عربية تنقلب (2026/11/19 تظهر 19/11/2026):
' تُحاط بعلامتي «تضمين من اليسار» حتى تبقى كما كُتبت
Private Function LtrText(ByVal s As String) As String
    LtrText = ChrW(&H202A) & s & ChrW(&H202C)
End Function

' «الرتبة/ الاسم» كما في النموذج
Private Function PersonLine(db As DAO.Database, ByVal userID As Long) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT RankTitle, FullName FROM tblUsers WHERE UserID=" & userID, dbOpenSnapshot)
    If Not rs.EOF Then
        If Len(Nz(rs!RankTitle, "")) > 0 Then PersonLine = rs!RankTitle & "/ " & rs!FullName Else PersonLine = rs!FullName
    End If
    rs.Close
End Function

' يضع توقيع المستخدم مكان العلامة؛ يعيد اسمه (سطراً) إذا لم توجد صورة توقيعه
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
