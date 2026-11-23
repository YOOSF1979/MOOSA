Attribute VB_Name = "modDashboard"
Option Compare Database
Option Explicit

' ===========================================================
' لوحة المؤشرات (frmDashboard)
' الإصدار 1 - 23 نوفمبر 2026
'
' - الفترة: السنة + الربع أو الشهر (الشهر يغلب الربع). الأرقام بتاريخ ورود المجموعة،
'   وسجل التشغيل بتاريخ المعالجة، والمخزون: الرصيد الحالي + المصروف في الفترة.
' - 12 مؤشراً في بطاقات (lblK1..lblK12) و9 رسوم (imgC1..imgC9).
' - الرسوم تُرسم في Excel مخفي (ربط متأخر) وتُصدَّر PNG إلى %TEMP%\dashboard،
'   لأن رسوم Access الحديثة لا تُنشأ آلياً. Excel الأحدث في الإصدار نفسه من Office.
' ===========================================================

Private Const C_BLUE As Long = 12611584       ' RGB(0,112,192)
Private Const C_LIGHT As Long = 13998939      ' RGB(91,155,213)
Private Const C_GREEN As Long = 5939246       ' RGB(46,160,90)
Private Const C_RED As Long = 4210892         ' RGB(204,64,64)
Private Const C_AMBER As Long = 3243501       ' RGB(237,125,49)
Private Const C_GREY As Long = 10855845       ' RGB(165,165,165)
Private Const C_NAVY As Long = 6567967        ' RGB(31,56,100)
Private Const C_TEAL As Long = 7249936        ' RGB(16,160,110) أخضر البرنامج

Private mFrom As Date, mTo As Date, mYear As Long

' ---------- فتح اللوحة ----------
Public Function OpenDashboard() As Boolean
    DoCmd.OpenForm "frmDashboard"
    OpenDashboard = True
End Function

Public Function DashOpen(frm As Form) As Boolean
    Dim rs As DAO.Recordset, years As String, y As Long
    y = Year(Date)
    years = CStr(y)
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT Year(ReceivedDate) FROM tblBatches WHERE Year(ReceivedDate)<>" & y & _
                                     " ORDER BY 1 DESC", dbOpenSnapshot)
    Do Until rs.EOF
        years = years & ";" & rs.Fields(0).Value
        rs.MoveNext
    Loop
    rs.Close
    frm!cboYear.RowSource = years
    frm!cboYear = y
    frm!cboQuarter = 0
    frm!cboMonth = 0
    DashOpen = DashRefresh(frm)
End Function

' ---------- الفترة من الفلاتر ----------
Public Function DashPeriod(ByVal yr As Long, ByVal qtr As Long, ByVal mon As Long) As String
    mYear = yr
    If mon >= 1 And mon <= 12 Then
        mFrom = DateSerial(yr, mon, 1): mTo = DateSerial(yr, mon + 1, 0)
        DashPeriod = "شهر " & mon & " / " & yr
    ElseIf qtr >= 1 And qtr <= 4 Then
        mFrom = DateSerial(yr, (qtr - 1) * 3 + 1, 1): mTo = DateSerial(yr, qtr * 3 + 1, 0)
        DashPeriod = "الربع " & qtr & " / " & yr
    Else
        mFrom = DateSerial(yr, 1, 1): mTo = DateSerial(yr, 12, 31)
        DashPeriod = "سنة " & yr
    End If
End Function

' ---------- التحديث ----------
Public Function DashRefresh(frm As Form) As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim inB As String, pos As Double, neg As Double, unk As Double, tot As Double
    Dim files As Variant, i As Long

    If Not IsNumeric(frm!cboYear) Then Msg "اختر السنة.", vbExclamation: Exit Function
    On Error GoTo EH
    frm!lblPeriod.Caption = "الفترة: " & DashPeriod(CLng(frm!cboYear), CLng(Nz(frm!cboQuarter, 0)), CLng(Nz(frm!cboMonth, 0)))
    Set db = CurrentDb
    inB = "ReceivedDate Between " & SqlD(mFrom) & " And " & SqlD(mTo)

    ' ---- المؤشرات ----
    SetK frm, 1, DbOne(db, "SELECT Count(*) FROM (SELECT DISTINCT CaseID FROM tblBatches WHERE " & inB & ")")
    SetK frm, 2, DbOne(db, "SELECT Count(*) FROM tblBatches WHERE " & inB)
    Set rs = db.OpenRecordset("SELECT Sum(Samples), Sum(PosCount), Sum(NegCount), Sum(UnkCount), Sum(ProcTraces)+Sum(VisitTraces), " & _
                              "Sum(Tests), Sum(VisitCases) FROM qStatBase WHERE " & inB, dbOpenSnapshot)
    tot = Nz(rs.Fields(0), 0): pos = Nz(rs.Fields(1), 0): neg = Nz(rs.Fields(2), 0): unk = Nz(rs.Fields(3), 0)
    SetK frm, 3, tot
    SetK frm, 4, pos
    SetK frm, 5, neg
    SetK frm, 6, unk
    SetK frm, 7, Nz(rs.Fields(4), 0)
    SetK frm, 8, Nz(rs.Fields(5), 0)
    SetK frm, 9, Nz(rs.Fields(6), 0)
    rs.Close
    ' If لا IIf: IIf يحسب الطرفين، فيقسم على صفر حين لا عينات (Overflow)
    If tot > 0 Then frm!lblK10.Caption = Format(pos / tot, "0.0%") Else frm!lblK10.Caption = "-"
    SetK frm, 11, DbOne(db, "SELECT Count(*) FROM tblRuns WHERE RunDate Between " & SqlD(mFrom) & " And " & SqlD(mTo))
    SetK frm, 12, DbOne(db, "SELECT Count(*) FROM qStockBalance WHERE MinStock Is Not Null AND Balance<=MinStock")

    ' ---- الرسوم ----
    DoCmd.Hourglass True
    files = MakeCharts(db, pos, neg, unk)
    DoCmd.Hourglass False
    For i = 1 To 9
        If Len(files(i)) > 0 Then frm("imgC" & i).Picture = files(i)
    Next
    DashRefresh = True
    Exit Function
EH:
    DoCmd.Hourglass False
    Msg "تعذر تحديث لوحة المؤشرات: " & Err.Number & " - " & Err.Description, vbExclamation
End Function

Private Sub SetK(frm As Form, ByVal k As Long, ByVal v As Variant)
    frm("lblK" & k).Caption = Format(Nz(v, 0), "#,##0")
End Sub

' ===========================================================
' الرسوم في Excel
' ===========================================================
Public Function MakeCharts(db As DAO.Database, ByVal pos As Double, ByVal neg As Double, ByVal unk As Double) As Variant
    Dim xl As Object, wb As Object, ws As Object, fso As Object
    Dim outDir As String, f(1 To 9) As String, tag As String
    Dim inB As String, inR As String, inM As String, rs As DAO.Recordset

    Set fso = CreateObject("Scripting.FileSystemObject")
    outDir = Environ$("TEMP") & "\dashboard"
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    tag = Format(Now, "hhnnss")                   ' اسم جديد في كل تحديث حتى لا تُعرض صورة قديمة
    inB = "b.ReceivedDate Between " & SqlD(mFrom) & " And " & SqlD(mTo)
    inR = "r.RunDate Between " & SqlD(mFrom) & " And " & SqlD(mTo)
    inM = "s.MoveDate Between " & SqlD(mFrom) & " And " & SqlD(mTo)

    On Error GoTo EH
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.DisplayRightToLeft = True

    ' 1) نتائج العينات (دائري مجوف)
    ws.Range("A1:B4").Value = ArrTable(Array("النتيجة", "العدد"), Array("إيجابي", pos), Array("سلبي", neg), Array("لم يتبين", unk))
    f(1) = OneChart(ws, ws.Range("A1:B4"), -4120, "نتائج العينات (بالأوراق)", Array(C_GREEN, C_RED, C_AMBER), True, outDir & "\c1_" & tag & ".png")

    ' 2) حالة المجموعات: صدر تقريرها / قيد العمل
    Set rs = db.OpenRecordset("SELECT IIf(ExportDate Is Null,'قيد العمل','صدر التقرير') AS St, Count(*) FROM tblBatches AS b WHERE " & inB & _
                              " GROUP BY IIf(ExportDate Is Null,'قيد العمل','صدر التقرير')", dbOpenSnapshot)
    f(2) = RsChart(ws, rs, 10, Array("الحالة", "المجموعات"), 5, "حالة القضايا (المجموعات)", Array(C_GREEN, C_AMBER), True, outDir & "\c2_" & tag & ".png")

    ' 3) حالة الآثار
    Set rs = db.OpenRecordset("SELECT st.StatusName, Count(*) FROM ((tblTraces AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) " & _
                              "INNER JOIN tblBatches AS b ON s.BatchID = b.BatchID) INNER JOIN tblTraceStatus AS st ON t.StatusID = st.StatusID " & _
                              "WHERE " & inB & " GROUP BY st.StatusID, st.StatusName ORDER BY st.StatusID", dbOpenSnapshot)
    f(3) = RsChart(ws, rs, 20, Array("الحالة", "الآثار"), 57, "حالة الآثار", Array(C_TEAL), False, outDir & "\c3_" & tag & ".png")

    ' 4) الإنتاج الشهري خلال السنة: العينات (أعمدة) والمجموعات (خط)
    Set rs = db.OpenRecordset("SELECT Mo, Sum(Samples), Count(*) FROM qStatBase WHERE Yr=" & mYear & " GROUP BY Mo ORDER BY Mo", dbOpenSnapshot)
    f(4) = MonthlyChart(ws, rs, 30, outDir & "\c4_" & tag & ".png")

    ' 5) الموظفون: العينات والإيجابي والفحوص
    Set rs = db.OpenRecordset("SELECT FullName, Sum(Samples), Sum(PosCount), Sum(Tests) FROM qStatBase AS b WHERE " & inB & _
                              " GROUP BY FullName ORDER BY Sum(Samples)", dbOpenSnapshot)
    f(5) = RsChart(ws, rs, 50, Array("الموظف", "العينات", "الإيجابي", "الفحوص"), 57, "إحصائيات الموظفين", Array(C_BLUE, C_GREEN, C_AMBER), False, outDir & "\c5_" & tag & ".png")

    ' 6) أكثر الجهات (عدد القضايا)
    Set rs = db.OpenRecordset("SELECT TOP 10 IIf(a.AgencyName Is Null,'(بلا جهة)',a.AgencyName) AS Ag, Count(*) AS n FROM " & _
                              "(SELECT DISTINCT c.CaseID, c.AgencyID FROM tblBatches AS b INNER JOIN tblCases AS c ON b.CaseID = c.CaseID WHERE " & inB & ") AS x " & _
                              "LEFT JOIN tblAgencies AS a ON x.AgencyID = a.AgencyID GROUP BY IIf(a.AgencyName Is Null,'(بلا جهة)',a.AgencyName) ORDER BY Count(*)", dbOpenSnapshot)
    f(6) = RsChart(ws, rs, 70, Array("الجهة", "القضايا"), 57, "القضايا حسب المراكز والجهات", Array(C_NAVY), False, outDir & "\c6_" & tag & ".png")

    ' 7) الإجراءات: أنواع المعالجة (بالأوراق) + الانتقال (بالأسطر)
    Set rs = db.OpenRecordset("SELECT IIf(s.LineType=2,'انتقال',s.MethodCode) AS PrName, Sum(IIf(s.LineType=2,1,s.ItemCount)) FROM tblSamples AS s " & _
                              "INNER JOIN tblBatches AS b ON s.BatchID = b.BatchID WHERE " & inB & _
                              " GROUP BY IIf(s.LineType=2,'انتقال',s.MethodCode) ORDER BY 1", dbOpenSnapshot)
    f(7) = RsChart(ws, rs, 90, Array("الإجراء", "العدد"), 51, "الإجراءات: أنواع المعالجة والانتقال", Array(C_TEAL), False, outDir & "\c7_" & tag & ".png")

    ' 8) المخزون: الرصيد الحالي والمصروف في الفترة
    Set rs = db.OpenRecordset("SELECT m.MaterialName, q.Balance, IIf(x.Used Is Null,0,x.Used) FROM (tblMaterials AS m INNER JOIN qStockBalance AS q " & _
                              "ON m.MaterialID = q.MaterialID) LEFT JOIN (SELECT s.MaterialID, Sum(s.Qty) AS Used FROM tblStockMoves AS s " & _
                              "WHERE s.MoveType=2 AND " & inM & " GROUP BY s.MaterialID) AS x ON m.MaterialID = x.MaterialID " & _
                              "WHERE m.IsActive=True AND (q.Balance<>0 OR x.Used Is Not Null OR m.MinStock Is Not Null) ORDER BY m.MaterialName", dbOpenSnapshot)
    f(8) = RsChart(ws, rs, 110, Array("المادة", "الرصيد الحالي", "المصروف في الفترة"), 57, "المخزون", Array(C_TEAL, C_RED), False, outDir & "\c8_" & tag & ".png")

    ' 9) سجل التشغيل: التشغيلات والعينات لكل جهاز
    Set rs = db.OpenRecordset("SELECT d.DeviceName, Count(*) AS Runs, Sum(IIf(x.Items Is Null,0,x.Items)) FROM (tblRuns AS r INNER JOIN tblDevices AS d " & _
                              "ON r.DeviceID = d.DeviceID) LEFT JOIN (SELECT rs.RunID, Sum(s.ItemCount) AS Items FROM tblRunSamples AS rs " & _
                              "INNER JOIN tblSamples AS s ON rs.SampleID = s.SampleID GROUP BY rs.RunID) AS x ON r.RunID = x.RunID " & _
                              "WHERE " & inR & " GROUP BY d.DeviceName ORDER BY d.DeviceName", dbOpenSnapshot)
    f(9) = RsChart(ws, rs, 130, Array("الجهاز", "التشغيلات", "العينات"), 51, "سجل تشغيل الأجهزة", Array(C_BLUE, C_TEAL), False, outDir & "\c9_" & tag & ".png")

    wb.Close False
    xl.Quit
    MakeCharts = f
    Exit Function
EH:
    Dim errText As String
    errText = "تعذر رسم المخططات (Excel): " & Err.Description
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Msg errText, vbExclamation
    MakeCharts = f
End Function

' جدول بيانات من سجل DAO: صف عناوين ثم القيم؛ إن كان فارغاً فسطر «لا توجد بيانات»
Private Function RsChart(ws As Object, rs As DAO.Recordset, ByVal row0 As Long, heads As Variant, ByVal chartType As Long, _
                         ByVal title As String, colors As Variant, ByVal byPoint As Boolean, ByVal outFile As String) As String
    Dim n As Long, i As Long, cols As Long
    cols = UBound(heads) + 1
    For i = 0 To cols - 1
        ws.Cells(row0, i + 1).Value = heads(i)
    Next
    If rs.EOF Then
        ws.Cells(row0 + 1, 1).Value = "لا توجد بيانات"
        For i = 2 To cols: ws.Cells(row0 + 1, i).Value = 0: Next
        n = 1
    Else
        n = ws.Cells(row0 + 1, 1).CopyFromRecordset(rs)
    End If
    rs.Close
    RsChart = OneChart(ws, ws.Range(ws.Cells(row0, 1), ws.Cells(row0 + n, cols)), chartType, title, colors, byPoint, outFile)
End Function

Private Function OneChart(ws As Object, src As Object, ByVal chartType As Long, ByVal title As String, colors As Variant, _
                          ByVal byPoint As Boolean, ByVal outFile As String) As String
    Dim co As Object, ch As Object, i As Long
    Set co = ws.ChartObjects.Add(300, 10, 440, 270)
    Set ch = co.Chart
    ch.chartType = chartType
    ch.SetSourceData src, 2                       ' xlColumns
    ch.HasTitle = True
    ch.ChartTitle.text = title
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 15
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C_NAVY
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 11
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    If byPoint Then
        ' دائري: لون لكل جزء، والتسميات بالقيمة والنسبة
        For i = 1 To ch.SeriesCollection(1).Points.Count
            ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = colors((i - 1) Mod (UBound(colors) + 1))
        Next
        ch.SeriesCollection(1).HasDataLabels = True
        ch.SeriesCollection(1).DataLabels.ShowPercentage = True
        ch.SeriesCollection(1).DataLabels.Font.Size = 12
        ch.HasLegend = True
        ch.Legend.Position = -4107                ' أسفل
    Else
        For i = 1 To ch.SeriesCollection.Count
            ch.SeriesCollection(i).Format.Fill.ForeColor.RGB = colors((i - 1) Mod (UBound(colors) + 1))
            ch.SeriesCollection(i).HasDataLabels = True
            ch.SeriesCollection(i).DataLabels.Font.Size = 11
        Next
        ch.HasLegend = (ch.SeriesCollection.Count > 1)
        If ch.HasLegend Then ch.Legend.Position = -4107
        ch.ChartGroups(1).GapWidth = 60
    End If
    ch.Export outFile, "PNG"
    co.Delete
    OneChart = outFile
End Function

' الإنتاج الشهري: الأشهر 1..12 كلها، العينات أعمدة والمجموعات خط على محور ثانوي
Private Function MonthlyChart(ws As Object, rs As DAO.Recordset, ByVal row0 As Long, ByVal outFile As String) As String
    Dim m As Long, co As Object, ch As Object
    ws.Cells(row0, 1).Value = "الشهر": ws.Cells(row0, 2).Value = "العينات": ws.Cells(row0, 3).Value = "المجموعات"
    For m = 1 To 12
        ws.Cells(row0 + m, 1).Value = Choose(m, "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر")
        ws.Cells(row0 + m, 2).Value = 0
        ws.Cells(row0 + m, 3).Value = 0
    Next
    Do Until rs.EOF
        ws.Cells(row0 + rs.Fields(0).Value, 2).Value = Nz(rs.Fields(1).Value, 0)
        ws.Cells(row0 + rs.Fields(0).Value, 3).Value = Nz(rs.Fields(2).Value, 0)
        rs.MoveNext
    Loop
    rs.Close
    Set co = ws.ChartObjects.Add(300, 10, 440, 270)
    Set ch = co.Chart
    ch.chartType = 51
    ch.SetSourceData ws.Range(ws.Cells(row0, 1), ws.Cells(row0 + 12, 3)), 2
    ch.SeriesCollection(2).chartType = 65         ' xlLineMarkers
    ch.SeriesCollection(2).AxisGroup = 2
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = C_BLUE
    ch.SeriesCollection(2).Format.Line.ForeColor.RGB = C_TEAL
    ch.SeriesCollection(2).Format.Line.Weight = 2.5
    ch.HasTitle = True
    ch.ChartTitle.text = "الإنتاج الشهري " & mYear
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 15
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C_NAVY
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 11
    ch.ChartArea.Format.Line.Visible = False
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.ChartGroups(1).GapWidth = 60
    ch.Export outFile, "PNG"
    co.Delete
    MonthlyChart = outFile
End Function

' مصفوفة ثنائية من صفوف
Private Function ArrTable(ParamArray rows() As Variant) As Variant
    Dim a() As Variant, r As Long, c As Long
    ReDim a(1 To UBound(rows) + 1, 1 To UBound(rows(0)) + 1)
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            a(r + 1, c + 1) = rows(r)(c)
        Next
    Next
    ArrTable = a
End Function

Private Function DbOne(db As DAO.Database, ByVal sql As String) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then DbOne = rs.Fields(0).Value
    rs.Close
End Function

Private Function SqlD(ByVal d As Date) As String
    SqlD = "#" & Format(d, "yyyy-mm-dd") & "#"
End Function
