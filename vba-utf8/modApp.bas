Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

' ===========================================================
' برنامج قسم الإظهار الكيميائي - منطق التشغيل والشاشات
' الإصدار 2 - 22 نوفمبر 2026 (مستوى المجموعات)
'
' النماذج لا تحتوي كوداً. كل حدث في النموذج يستدعي دالة هنا
' مباشرة من خاصية الحدث، مثل:  =SampleAfterUpdate([Form])
'
' الحلقة:
'   القضية (tblCases) -> المجموعات (tblBatches: كل استلام بموظفه وتواريخه وتقريره)
'     -> العينات (tblSamples) -> الآثار (tblTraces)
'   الآثار تُنشأ تلقائياً بعدد الآثار الإيجابية (SyncTraces)، ولا يُضيفها الموظف بيده.
' ===========================================================

Public gSilent As Boolean        ' للتشغيل الآلي: الرسائل تُجمع في gLog بدل النوافذ
Public gLog As String

Private Const APP_NAME As String = "قسم الإظهار الكيميائي"
Private Const PROC_TRACE_DESC As String = "أثر معالجة رقم "
Private Const VISIT_TRACE_DESC As String = "أثر مصور رقم "

' ===========================================================
' الرسائل
' ===========================================================
Public Sub Msg(ByVal text As String, Optional ByVal style As VbMsgBoxStyle = vbInformation)
    If gSilent Then
        gLog = gLog & text & vbCrLf
    Else
        MsgBox text, style + vbMsgBoxRight + vbMsgBoxRtlReading, APP_NAME
    End If
End Sub

' في الوضع الصامت الجواب «لا» دائماً
Public Function Ask(ByVal text As String) As Boolean
    If gSilent Then
        gLog = gLog & "[سؤال، الجواب لا] " & text & vbCrLf
    Else
        Ask = (MsgBox(text, vbYesNo + vbQuestion + vbDefaultButton2 + vbMsgBoxRight + vbMsgBoxRtlReading, APP_NAME) = vbYes)
    End If
End Function

' للاختبار الآلي فقط: يفعّل/يوقف الوضع الصامت ويعيد الرسائل المجمّعة ثم يفرّغها
Public Function SetSilent(ByVal silentOn As Boolean) As String
    gSilent = silentOn
    SetSilent = gLog
    gLog = ""
End Function

' للاختبار الآلي: أول قيمة من استعلام يقرأ [TempVars]! (DAO لا يحلّها وحده، فتُملأ المعلمات بـ Eval)
Public Function QueryValue(ByVal sql As String) As Variant
    Dim qd As DAO.QueryDef, prm As DAO.Parameter, rs As DAO.Recordset
    Set qd = CurrentDb.CreateQueryDef("", sql)
    For Each prm In qd.Parameters
        prm.Value = Eval(prm.Name)
    Next
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then QueryValue = rs.Fields(0).Value
    rs.Close
End Function

Public Function NotYet(ByVal featureName As String) As Boolean
    Msg featureName & ": سيُبنى في مرحلة لاحقة.", vbInformation
End Function

' ===========================================================
' الدخول: اسم مستخدم الجهاز مقابل tblUsers
' ===========================================================
Public Function SignIn(Optional ByVal winUser As String = "") As Boolean
    Dim rs As DAO.Recordset
    If Len(winUser) = 0 Then winUser = Environ$("USERNAME")
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT UserID, FullName FROM tblUsers WHERE WinUser=" & Q(winUser) & " AND IsActive=True", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Msg "اسم المستخدم «" & winUser & "» غير مسجل في البرنامج." & vbCrLf & _
            "راجع مسؤول البرنامج لإضافتك.", vbCritical
        Exit Function
    End If
    TempVars.Add "UserID", CLng(rs!userID.Value)
    TempVars.Add "UserName", CStr(rs!FullName.Value)
    TempVars.Add "WinUser", winUser
    rs.Close
    SignIn = True
End Function

' ===========================================================
' الشاشة الرئيسية (القضية)
' ===========================================================
Public Function MainLoad(frm As Form) As Boolean
    If Not SignIn() Then
        If Not gSilent Then DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    frm!lblWelcome.Caption = "مرحباً، " & TempVars("UserName")
    MainLoad = True
End Function

Public Function CaseNew(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    frm.SetFocus
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    frm!refNo.SetFocus
    CaseNew = True
End Function

Public Function CaseFind(frm As Form) As Boolean
    Dim caseID As Variant
    caseID = frm!cboFind.Value
    frm!cboFind.Value = Null
    If IsNull(caseID) Then Exit Function
    If Not SaveRecord(frm) Then Exit Function
    frm.Recordset.FindFirst "CaseID=" & caseID
    If frm.Recordset.NoMatch Then Msg "لم يُعثر على القضية.", vbExclamation Else CaseFind = True
End Function

' يحفظ السجل الحالي؛ يعيد False ويعرض السبب إذا رفضته قواعد الجدول
Public Function SaveRecord(frm As Form) As Boolean
    On Error GoTo EH
    If frm.Dirty Then frm.Dirty = False
    SaveRecord = True
    Exit Function
EH:
    Msg "لم يُحفظ السجل:" & vbCrLf & Err.Description, vbExclamation
End Function

' القضية يجب أن تكون محفوظة قبل الانتقال لعيناتها
Private Function CaseReady(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If frm.NewRecord Or IsNull(frm!caseID) Then
        Msg "أدخل بيانات القضية واحفظها أولاً.", vbExclamation
        Exit Function
    End If
    CaseReady = True
End Function

' ===========================================================
' المجموعات: كل استلام (موظف + مجموعة عينات) له تواريخه وتقريره وسجل تشغيله
' ===========================================================
' استلام جديد: المستخدم هو الفاحص، تاريخ الورود اليوم، ثم تُفتح عيناته مباشرة
Public Function BatchNew(frm As Form) As Boolean
    Dim db As DAO.Database
    Dim batchID As Long, nextNo As Long
    If Not CaseReady(frm) Then Exit Function
    Set db = CurrentDb
    nextNo = Nz(DbVal(db, "SELECT Max(BatchNo) FROM tblBatches WHERE CaseID=" & frm!caseID), 0) + 1
    db.Execute "INSERT INTO tblBatches (CaseID, BatchNo, ReceivedDate, ExaminerID, CreatedBy) VALUES (" & _
               frm!caseID & ", " & nextNo & ", Date(), " & TempVars("UserID") & ", " & Q(Nz(TempVars("WinUser"), "")) & ")", dbFailOnError
    batchID = DbVal(db, "SELECT @@IDENTITY")
    frm!sfBatches.Requery
    DoCmd.OpenForm "frmSamples", , , "BatchID=" & batchID, , , CStr(batchID)
    BatchNew = True
End Function

' زر «فتح» في سطر المجموعة
Public Function BatchOpen(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If IsNull(frm!batchID) Then Exit Function
    DoCmd.OpenForm "frmSamples", , , "BatchID=" & frm!batchID, , , CStr(frm!batchID)
    BatchOpen = True
End Function

' scope: 0 = آثار القضية من الرئيسية، 1 = آثار عينة واحدة، 2 = آثار المجموعة من شاشة العينات
Public Function OpenTraces(frm As Form, ByVal scope As Integer) As Boolean
    Dim where As String
    Select Case scope
        Case 0
            If Not CaseReady(frm) Then Exit Function
            where = "CaseID=" & frm!caseID
        Case 1
            If Not SaveRecord(frm) Then Exit Function
            If frm.NewRecord Or IsNull(frm!sampleID) Then
                Msg "احفظ سطر العينة أولاً.", vbExclamation
                Exit Function
            End If
            where = "SampleID=" & frm!sampleID
        Case Else
            If Not SaveRecord(frm) Then Exit Function
            where = "BatchID=" & Nz(frm.OpenArgs, 0)
    End Select

    If DCount("*", "qTracesEdit", where) = 0 Then
        Msg "لا توجد آثار." & vbCrLf & _
            "الآثار تُنشأ تلقائياً عند اختيار «إيجابية» وتحديد عدد الآثار في شاشة العينات.", vbInformation
        Exit Function
    End If
    DoCmd.OpenForm "frmTraces", , , where
    OpenTraces = True
End Function

' سجل التشغيل يُفتح من شاشة العينات ومعه رقم المجموعة
Public Function OpenRuns(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If CurrentProject.AllForms("frmRuns").IsLoaded Then DoCmd.Close acForm, "frmRuns"
    DoCmd.OpenForm "frmRuns", , , , , , CStr(Nz(frm.OpenArgs, 0))
    OpenRuns = True
End Function

Public Function CloseMe(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    DoCmd.Close acForm, frm.Name
    CloseMe = True
End Function

' ===========================================================
' شاشة العينات
' ===========================================================
Public Function SamplesOpen(frm As Form) As Boolean
    Dim batchID As Long
    Dim rs As DAO.Recordset

    batchID = CLng(Nz(frm.OpenArgs, 0))
    If batchID = 0 Then
        Msg "افتح شاشة العينات من مجموعات القضية في الشاشة الرئيسية.", vbExclamation
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT c.RefNo, c.ForensicNo, b.BatchNo, b.ReceivedDate, u.FullName " & _
        "FROM (tblBatches AS b INNER JOIN tblCases AS c ON b.CaseID = c.CaseID) " & _
        "INNER JOIN tblUsers AS u ON b.ExaminerID = u.UserID WHERE b.BatchID=" & batchID, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!lblCase.Caption = "رقم المرجع: " & rs!refNo & "     الأدلة الجنائية: " & Nz(rs!ForensicNo, "-") & _
                              "     المجموعة " & rs!BatchNo & ": " & rs!FullName & " - " & Format(rs!ReceivedDate, "yyyy/mm/dd")
    End If
    rs.Close

    ' القيم الافتراضية للسطر الجديد
    frm!batchID.DefaultValue = batchID
    frm!MethodCode.DefaultValue = """" & Nz(DLookup("MethodCode", "tblMethods", "IsDefault=True"), "") & """"
    frm!TransferTargetID.DefaultValue = Nz(DLookup("TargetID", "tblTransferTargets", "IsDefault=True"), "Null")
    SamplesOpen = True
End Function

' رقم السطر التالي في المجموعة (الترقيم يبدأ من 1 في كل تقرير)
Public Function SampleBeforeInsert(frm As Form) As Boolean
    frm!SampleNo = Nz(DMax("SampleNo", "tblSamples", "BatchID=" & frm!batchID), 0) + 1
    SampleBeforeInsert = True
End Function

' عينة (1) أو انتقال (2)
Public Function SampleTypeChanged(frm As Form) As Boolean
    ' الانتقال: حقول العينة تُفرَّغ وتُقفل (تنسيق شرطي)، ولا يبقى إلا عدد الآثار المرفوعة
    If Nz(frm!LineType, 1) = 2 Then
        frm!Barcode = Null
        frm!Description = Null
        frm!ItemCount = Null
        frm!LiftLocation = Null
        frm!MethodCode = Null
        frm!PosCount = Null
        frm!NegCount = Null
        frm!UnkCount = Null
        frm!PhotoPath = Null
        If IsNull(frm!TraceCount) Then frm!TraceCount = 0
        If IsNull(frm!VisitDateTime) Then frm!VisitDateTime = Now()
        frm!TraceCount.SetFocus
    Else
        frm!VisitDateTime = Null
        frm!TraceCount = Null
        If IsNull(frm!ItemCount) Then frm!ItemCount = 1
        If IsNull(frm!MethodCode) Then frm!MethodCode = DLookup("MethodCode", "tblMethods", "IsDefault=True")
    End If
    SampleTypeChanged = True
End Function

' العدد = إيجابي + سلبي + لم يتبين: «لم يتبين» يُحسب وحده من الباقي
' ولا آثار بلا إيجابي
Public Function SampleCountsChanged(frm As Form) As Boolean
    Dim rest As Long
    If Not IsNull(frm!ItemCount) And Not IsNull(frm!PosCount) And Not IsNull(frm!NegCount) Then
        rest = frm!ItemCount - frm!PosCount - frm!NegCount
        If rest >= 0 Then frm!UnkCount = rest
    End If
    If Nz(frm!PosCount, 0) = 0 Then frm!TraceCount = Null
    SampleCountsChanged = True
End Function

Public Function SampleAfterUpdate(frm As Form) As Boolean
    SampleAfterUpdate = SyncTraces(CLng(frm!sampleID))
    If CurrentProject.AllForms("frmTraces").IsLoaded Then Forms("frmTraces").Requery
    StaffAfterSampleChange CLng(frm!batchID)
End Function

Public Function SampleAfterDelete(frm As Form) As Boolean
    StaffAfterSampleChange CLng(Nz(frm.OpenArgs, 0))
    SampleAfterDelete = True
End Function

' بعد أي تغيير في العينات: غير المشتركة تُحدَّث تلقائياً، والمشتركة تُنبَّه إذا اختل التطابق
Private Sub StaffAfterSampleChange(ByVal batchID As Long)
    If batchID = 0 Then Exit Sub
    If Not SyncBatchStaff(batchID) Then
        Msg "هذه مجموعة مشتركة وتغيّرت عيناتها:" & vbCrLf & BatchStaffDiff(batchID) & vbCrLf & _
            "عدّل التوزيع من زر «توزيع العمل».", vbExclamation
    End If
End Sub

' ===========================================================
' توزيع العمل (tblBatchStaff) - مثل سطر Tests القديم، لكل موظف في المجموعة
'   - غير مشتركة (سطر واحد أو لا سطر): سطر الفاحص = مجاميع العينات، تلقائياً
'   - مشتركة (أكثر من سطر): يوزّعها الموظف بيده، والمجموع يجب أن يطابق العينات
'   - الأعداد بالأوراق/القطع: عينة من 100 ورقة = 100 (إيجابي + سلبي + لم يتبين)
'   - الفحوص لا تُخزَّن: لكل تشغيلة دخلتها العينة = (إيجابي + سلبي) فحصاً، تُنسب لموظف التشغيلة
'     (qRunWork.TestItems). فحوص «لم يتبين» لا تدخل في المجموع، كما في Tests القديم.
' ===========================================================
' مجاميع المجموعة من العينات: إيجابي، سلبي، لم يتبين، آثار المعالجة، آثار الانتقال
Private Function BatchTotals(db As DAO.Database, ByVal batchID As Long) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT Sum(IIf(LineType=1,Nz(PosCount,0),0)), Sum(IIf(LineType=1,Nz(NegCount,0),0)), " & _
        "Sum(IIf(LineType=1,Nz(UnkCount,0),0)), Sum(IIf(LineType=1,Nz(TraceCount,0),0)), " & _
        "Sum(IIf(LineType=2,Nz(TraceCount,0),0)) FROM tblSamples WHERE BatchID=" & batchID, dbOpenSnapshot)
    BatchTotals = Array(Nz(rs.Fields(0), 0), Nz(rs.Fields(1), 0), Nz(rs.Fields(2), 0), Nz(rs.Fields(3), 0), Nz(rs.Fields(4), 0))
    rs.Close
End Function

Private Function StaffTotals(db As DAO.Database, ByVal batchID As Long) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT Sum(PosCount), Sum(NegCount), Sum(UnkCount), Sum(ProcTraces), Sum(VisitTraces) " & _
        "FROM tblBatchStaff WHERE BatchID=" & batchID, dbOpenSnapshot)
    StaffTotals = Array(Nz(rs.Fields(0), 0), Nz(rs.Fields(1), 0), Nz(rs.Fields(2), 0), Nz(rs.Fields(3), 0), Nz(rs.Fields(4), 0))
    rs.Close
End Function

' يعيد False إذا كانت المجموعة مشتركة ومجموعها لا يطابق العينات
Public Function SyncBatchStaff(ByVal batchID As Long) As Boolean
    Dim db As DAO.Database
    Dim t As Variant
    Dim rows As Long, examinerID As Long

    DBEngine.Idle dbRefreshCache
    Set db = CurrentDb
    t = BatchTotals(db, batchID)
    rows = DbVal(db, "SELECT Count(*) FROM tblBatchStaff WHERE BatchID=" & batchID)
    If rows = 0 Then
        examinerID = DbVal(db, "SELECT ExaminerID FROM tblBatches WHERE BatchID=" & batchID)
        db.Execute "INSERT INTO tblBatchStaff (BatchID, UserID, PosCount, NegCount, UnkCount, ProcTraces, VisitTraces) VALUES (" & _
                   batchID & ", " & examinerID & ", " & t(0) & ", " & t(1) & ", " & t(2) & ", " & t(3) & ", " & t(4) & ")", dbFailOnError
    ElseIf rows = 1 Then
        db.Execute "UPDATE tblBatchStaff SET PosCount=" & t(0) & ", NegCount=" & t(1) & ", UnkCount=" & t(2) & _
                   ", ProcTraces=" & t(3) & ", VisitTraces=" & t(4) & " WHERE BatchID=" & batchID, dbFailOnError
    Else
        SyncBatchStaff = (Len(BatchStaffDiff(batchID)) = 0)
        Exit Function
    End If
    SyncBatchStaff = True
End Function

' نص الفروق بين التوزيع والعينات، أو "" إذا تطابقا
Public Function BatchStaffDiff(ByVal batchID As Long) As String
    Dim db As DAO.Database
    Dim t As Variant, d As Variant, names As Variant
    Dim i As Long, s As String
    Set db = CurrentDb
    t = BatchTotals(db, batchID)
    d = StaffTotals(db, batchID)
    names = Array("الإيجابية", "السلبية", "لم يتبين", "آثار المعالجة", "آثار الانتقال")
    For i = 0 To 4
        If t(i) <> d(i) Then s = s & "- " & names(i) & ": في العينات " & t(i) & "، الموزَّع " & d(i) & vbCrLf
    Next
    BatchStaffDiff = s
End Function

Public Function OpenStaff(frm As Form) As Boolean
    Dim batchID As Long
    If Not SaveRecord(frm) Then Exit Function
    batchID = CLng(Nz(frm.OpenArgs, 0))
    If batchID = 0 Then Exit Function
    SyncBatchStaff batchID
    DoCmd.OpenForm "frmBatchStaff", , , "BatchID=" & batchID, , , CStr(batchID)
    OpenStaff = True
End Function

Public Function StaffOpen(frm As Form) As Boolean
    frm!batchID.DefaultValue = CLng(Nz(frm.OpenArgs, 0))
    StaffOpen = StaffRefresh(frm)
End Function

Public Function StaffRefresh(frm As Form) As Boolean
    Dim db As DAO.Database
    Dim batchID As Long, t As Variant, d As Variant, diff As String
    batchID = CLng(Nz(frm.OpenArgs, 0))
    Set db = CurrentDb
    t = BatchTotals(db, batchID)
    d = StaffTotals(db, batchID)
    diff = BatchStaffDiff(batchID)
    frm!lblTotals.Caption = "العينات:   إيجابية " & t(0) & "   سلبية " & t(1) & "   لم يتبين " & t(2) & _
                            "   آثار المعالجة " & t(3) & "   آثار الانتقال " & t(4) & vbCrLf & _
                            "الموزَّع:  إيجابية " & d(0) & "   سلبية " & d(1) & "   لم يتبين " & d(2) & _
                            "   آثار المعالجة " & d(3) & "   آثار الانتقال " & d(4) & _
                            IIf(Len(diff) = 0, "     (مطابق)", "     (غير مطابق)")
    frm!lblTotals.ForeColor = IIf(Len(diff) = 0, RGB(16, 128, 88), RGB(192, 0, 0))
    StaffRefresh = True
End Function

' الإغلاق فقط إذا طابق التوزيع العينات
Public Function StaffClose(frm As Form) As Boolean
    Dim diff As String
    If Not SaveRecord(frm) Then Exit Function
    diff = BatchStaffDiff(CLng(Nz(frm.OpenArgs, 0)))
    If Len(diff) > 0 Then
        Msg "التوزيع لا يطابق عينات المجموعة:" & vbCrLf & diff, vbExclamation
        StaffRefresh frm
        Exit Function
    End If
    DoCmd.Close acForm, frm.Name
    StaffClose = True
End Function

' ===========================================================
' إنشاء الآثار تلقائياً
'   - العينة الإيجابية: عدد أسطر الآثار = عدد الآثار
'   - سطر الانتقال: عدد أسطر الآثار = عدد الآثار المصورة، وتبدأ من حالة 2
'   - الإنقاص لا يحذف أثراً مسجلاً في منصة أثر
'   - رمز الأثر = باركود العينة-رقم الأثر، مثل 900000101-01
'     الانتقال بلا باركود: رقم المرجع-رقم المجموعة V رقم السطر-رقم الأثر، مثل 226000000001-1V2-01
' ===========================================================
Public Function SyncTraces(ByVal sampleID As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim want As Long, have As Long, extra As Long, removable As Long, nextNo As Long
    Dim baseCode As String, descPrefix As String, firstStatus As Long, liftDefault As String

    DBEngine.Idle dbRefreshCache        ' قراءة آخر ما حُفظ، لا نسخة المحرك المؤقتة
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT s.SampleNo, s.LineType, s.Barcode, s.Description, s.PosCount, s.TraceCount, b.BatchNo, c.RefNo " & _
        "FROM (tblSamples AS s INNER JOIN tblBatches AS b ON s.BatchID = b.BatchID) " & _
        "INNER JOIN tblCases AS c ON b.CaseID = c.CaseID WHERE s.SampleID=" & sampleID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    If rs!LineType = 2 Then
        want = Nz(rs!TraceCount, 0)
        baseCode = rs!refNo & "-" & rs!BatchNo & "V" & rs!SampleNo
        descPrefix = VISIT_TRACE_DESC
        firstStatus = 2
    Else
        If Nz(rs!PosCount, 0) > 0 Then want = Nz(rs!TraceCount, 0)
        baseCode = Nz(rs!Barcode, "")
        descPrefix = PROC_TRACE_DESC
        firstStatus = 1
        ' مكان رفع الأثر الافتراضي: من على وصف العينة
        If Not IsNull(rs!Description) Then liftDefault = "من على " & rs!Description
    End If
    rs.Close

    ' الآثار التي لم يُحدَّد مكان رفعها تأخذ الافتراضي (المعدَّل يدوياً يبقى)
    If Len(liftDefault) > 0 Then
        db.Execute "UPDATE tblTraces SET LiftLocation = " & Q(liftDefault) & _
                   " WHERE SampleID=" & sampleID & " AND LiftLocation Is Null", dbFailOnError
    End If

    ' تحديث الرموز إذا تغيّر الباركود (الوصف الافتراضي يتبع الرمز، والوصف المعدّل يبقى)
    db.Execute "UPDATE tblTraces SET Description = " & Q(descPrefix & baseCode & "-") & " & Format(TraceNo,'00') " & _
               "WHERE SampleID=" & sampleID & " AND Description = " & Q(descPrefix) & " & TraceCode", dbFailOnError
    db.Execute "UPDATE tblTraces SET TraceCode = " & Q(baseCode & "-") & " & Format(TraceNo,'00') " & _
               "WHERE SampleID=" & sampleID, dbFailOnError

    ' العدّ عبر db نفسه لا DCount/DMax: دوال النطاق تعمل في جلسة منفصلة
    have = DbVal(db, "SELECT Count(*) FROM tblTraces WHERE SampleID=" & sampleID)

    ' زيادة
    Do While have < want
        nextNo = Nz(DbVal(db, "SELECT Max(TraceNo) FROM tblTraces WHERE SampleID=" & sampleID), 0) + 1
        db.Execute "INSERT INTO tblTraces (SampleID, TraceNo, TraceCode, Description, LiftLocation, StatusID) VALUES (" & _
                   sampleID & ", " & nextNo & ", " & Q(baseCode & "-" & Format(nextNo, "00")) & ", " & _
                   Q(descPrefix & baseCode & "-" & Format(nextNo, "00")) & ", " & _
                   IIf(Len(liftDefault) > 0, Q(liftDefault), "Null") & ", " & firstStatus & ")", dbFailOnError
        have = have + 1
    Loop

    ' إنقاص: من آخر الأرقام، وفقط الآثار غير المسجلة في منصة أثر
    If have > want Then
        extra = have - want
        removable = DbVal(db, "SELECT Count(*) FROM tblTraces WHERE SampleID=" & sampleID & " AND AtharBarcode Is Null")
        If removable < extra Then
            Msg "عدد الآثار أقل من الآثار المسجلة في منصة أثر لهذه العينة." & vbCrLf & _
                "لم يُحذف أي أثر. عدّل العدد أو احذف باركود أثر من شاشة الآثار أولاً.", vbExclamation
            Exit Function
        End If
        If Not Ask("سيتم حذف " & extra & " من آثار هذه العينة (الأرقام الأخيرة). متابعة؟") Then
            Msg "لم يُحذف أي أثر. عدد الآثار في العينة لا يطابق أسطر الآثار الآن.", vbExclamation
            Exit Function
        End If
        db.Execute "DELETE FROM tblTraces WHERE TraceID IN (SELECT TOP " & extra & " TraceID FROM tblTraces " & _
                   "WHERE SampleID=" & sampleID & " AND AtharBarcode Is Null ORDER BY TraceNo DESC)", dbFailOnError
    End If
    SyncTraces = True
End Function

' ===========================================================
' شاشة الآثار (الحالة تتغير بالإجراء، لا يختارها الموظف)
' ===========================================================
Public Function TraceBarcodeChanged(frm As Form) As Boolean
    Dim v As String
    v = Trim$(Nz(frm!AtharBarcode, ""))
    If Len(v) > 0 Then
        frm!AtharBarcode = v
        If frm!StatusID < 3 Then
            frm!StatusID = 3
            frm!AtharRegDate = Now()
        End If
    Else
        frm!AtharBarcode = Null
        frm!AtharRegDate = Null
        If frm!StatusID = 3 Then frm!StatusID = IIf(IsNull(frm!PhotoPath) And frm!LineType = 1, 1, 2)
    End If
    TraceBarcodeChanged = True
End Function

Public Function TracePhotoChanged(frm As Form) As Boolean
    If Not IsNull(frm!PhotoPath) And frm!StatusID = 1 Then frm!StatusID = 2
    TracePhotoChanged = True
End Function

' ===========================================================
' اختيار ملف صورة
' ===========================================================
Public Function PickFile(frm As Form, ByVal fieldName As String) As Boolean
    Dim fd As Object
    Set fd = Application.FileDialog(3)      ' msoFileDialogFilePicker
    fd.title = "اختر الصورة"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "الصور", "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff"
    If Len(Nz(TempVars("LastFolder"), "")) > 0 Then fd.InitialFileName = TempVars("LastFolder")
    If fd.Show = -1 Then
        frm(fieldName) = fd.SelectedItems(1)
        TempVars.Add "LastFolder", Left$(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), "\"))
        If fieldName = "PhotoPath" And frm.Name = "frmTraces" Then TracePhotoChanged frm
        PickFile = True
    End If
End Function

' ===========================================================
' سجل تشغيل الأجهزة (LOGBOOK)
' ===========================================================
Public Function RunsOpen(frm As Form) As Boolean
    Dim rs As DAO.Recordset
    ' الجهاز الافتراضي = آخر جهاز استخدمه الموظف
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 DeviceID FROM tblRuns WHERE UserID=" & Nz(TempVars("UserID"), 0) & " ORDER BY RunID DESC", dbOpenSnapshot)
    If Not rs.EOF Then frm!DeviceID.DefaultValue = rs!DeviceID
    rs.Close
    ' المجموعة التي فُتح منها السجل (لزر «كل عينات المجموعة»)
    TempVars.Add "RunBatchID", CLng(Nz(frm.OpenArgs, 0))
    frm.SetFocus
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    RunsOpen = True
End Function

' يضيف عينات المجموعة (المعالجة فقط، غير المضافة) إلى التشغيلة الحالية
Public Function RunAddBatch(frm As Form) As Boolean
    Dim db As DAO.Database
    Dim batchID As Long, before As Long, added As Long

    batchID = Nz(TempVars("RunBatchID"), 0)
    If batchID = 0 Then
        Msg "افتح سجل التشغيل من شاشة عينات المجموعة.", vbExclamation
        Exit Function
    End If
    If IsNull(frm!DeviceID) Then
        Msg "اختر الجهاز أولاً.", vbExclamation
        Exit Function
    End If
    If Not SaveRecord(frm) Then Exit Function

    Set db = CurrentDb
    ' العدد بمجموع «العدد» لا بعدد الأسطر
    before = RunItems(db, frm!runID)
    db.Execute "INSERT INTO tblRunSamples (RunID, SampleID) SELECT " & frm!runID & ", SampleID FROM tblSamples " & _
               "WHERE BatchID=" & batchID & " AND LineType=1 AND SampleID NOT IN " & _
               "(SELECT SampleID FROM tblRunSamples WHERE RunID=" & frm!runID & ")", dbFailOnError
    added = RunItems(db, frm!runID) - before
    frm!sfRunSamples.Requery
    frm!txtCount.Requery
    Msg "أضيفت عينات عددها " & added & " إلى التشغيلة.", vbInformation
    RunAddBatch = True
End Function

Public Function RunNew(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    frm.SetFocus
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    RunNew = True
End Function

' which: "start" أو "end"
Public Function RunNow(frm As Form, ByVal which As String) As Boolean
    If which = "start" Then
        frm!StartTime = Now()
        RunSetEnd frm
    Else
        frm!EndTime = Now()
    End If
    RunNow = SaveRecord(frm)
End Function

' اختيار الجهاز: البدء الآن (إن لم يُحدَّد)، والانتهاء = البدء + مدة دورة الجهاز
Public Function RunDeviceChanged(frm As Form) As Boolean
    If IsNull(frm!StartTime) Then frm!StartTime = Now()
    RunDeviceChanged = RunSetEnd(frm)
End Function

' وقت الانتهاء = وقت البدء + مدة الجهاز (tblDevices.DurationMin)
Public Function RunSetEnd(frm As Form) As Boolean
    Dim mins As Long
    If IsNull(frm!DeviceID) Or IsNull(frm!StartTime) Then Exit Function
    mins = Nz(DLookup("DurationMin", "tblDevices", "DeviceID=" & frm!DeviceID), 0)
    If mins > 0 Then frm!EndTime = DateAdd("n", mins, frm!StartTime)
    RunSetEnd = True
End Function

' مسح باركود عينة يضيفها للتشغيلة
Public Function RunScan(frm As Form) As Boolean
    Dim code As String
    Dim sampleID As Variant

    code = Trim$(Nz(frm!txtScan, ""))
    frm!txtScan = Null
    If Len(code) = 0 Then Exit Function
    If IsNull(frm!DeviceID) Then
        Msg "اختر الجهاز أولاً.", vbExclamation
        Exit Function
    End If
    If Not SaveRecord(frm) Then Exit Function

    sampleID = DLookup("SampleID", "tblSamples", "Barcode=" & Q(code))
    If IsNull(sampleID) Then
        Msg "لا توجد عينة بالباركود " & code & "." & vbCrLf & "أدخل العينة في شاشة العينات أولاً.", vbExclamation
    ElseIf DCount("*", "tblRunSamples", "RunID=" & frm!runID & " AND SampleID=" & sampleID) > 0 Then
        Msg "العينة " & code & " مضافة لهذه التشغيلة مسبقاً.", vbExclamation
    Else
        CurrentDb.Execute "INSERT INTO tblRunSamples (RunID, SampleID) VALUES (" & frm!runID & ", " & sampleID & ")", dbFailOnError
        RunScan = True
    End If
    frm!sfRunSamples.Requery
    frm!txtCount.Requery
    frm!DeviceID.SetFocus           ' للعودة إلى خانة المسح بعد AfterUpdate
    frm!txtScan.SetFocus
End Function

Public Function RunSampleRemove(frm As Form) As Boolean
    If IsNull(frm!sampleID) Then Exit Function
    CurrentDb.Execute "DELETE FROM tblRunSamples WHERE RunID=" & frm!runID & " AND SampleID=" & frm!sampleID, dbFailOnError
    frm.Requery
    frm.Parent!txtCount.Requery
    RunSampleRemove = True
End Function

' ===========================================================
' الإحصائيات (أزرار تذييل الرئيسية، كما في البرنامج القديم)
'   الفترة من شريط التذييل: السنة، من شهر، إلى شهر -> TempVars تقرؤها الاستعلامات
'   الفترة حسب تاريخ ورود المجموعة. الأعداد بالأوراق، والفحوص من سجل التشغيل.
' ===========================================================
Public Function SetStatPeriod(ByVal yr As Long, ByVal m1 As Long, ByVal m2 As Long) As Boolean
    If yr < 2000 Or yr > 2100 Or m1 < 1 Or m1 > 12 Or m2 < 1 Or m2 > 12 Or m1 > m2 Then
        Msg "الفترة غير صحيحة: السنة 4 أرقام، والشهر من 1 إلى 12، و«من شهر» لا يتجاوز «إلى شهر».", vbExclamation
        Exit Function
    End If
    TempVars.Add "StatYear", yr
    TempVars.Add "StatFrom", m1
    TempVars.Add "StatTo", m2
    ' جدول الجهات × الأشهر لا يقبل متغيرات: يُعاد بناؤه بالسنة
    CurrentDb.QueryDefs("احصائية القضايا حسب الجهات").sql = _
        "TRANSFORM Count(c.CaseID) AS n SELECT IIf(a.AgencyName Is Null,'(بلا جهة)',a.AgencyName) AS [الجهة], Count(c.CaseID) AS [الإجمالي] " & _
        "FROM tblCases AS c LEFT JOIN tblAgencies AS a ON c.AgencyID = a.AgencyID WHERE Year(c.CaseDate)=" & yr & " " & _
        "GROUP BY IIf(a.AgencyName Is Null,'(بلا جهة)',a.AgencyName) " & _
        "PIVOT Format(c.CaseDate,'mm') IN ('01','02','03','04','05','06','07','08','09','10','11','12');"
    SetStatPeriod = True
End Function

Public Function OpenStat(frm As Form, ByVal queryName As String) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If Not IsNumeric(frm!txtStatYear) Or Not IsNumeric(frm!txtStatFrom) Or Not IsNumeric(frm!txtStatTo) Then
        Msg "اكتب السنة والشهرين أرقاماً.", vbExclamation
        Exit Function
    End If
    If Not SetStatPeriod(CLng(frm!txtStatYear), CLng(frm!txtStatFrom), CLng(frm!txtStatTo)) Then Exit Function
    DoCmd.OpenQuery queryName, acViewNormal, acReadOnly
    OpenStat = True
End Function

' ===========================================================
' المخزون: الرصيد = الإضافات - المصروفات (tblStockMoves)
'   الصرف لا يتجاوز الرصيد، وحذف إضافة لا يجعل الرصيد سالباً.
' ===========================================================
Public Function OpenStock() As Boolean
    DoCmd.OpenForm "frmStock"
    OpenStock = True
End Function

Public Function StockBalance(ByVal materialID As Long) As Double
    StockBalance = Nz(DbVal(CurrentDb, "SELECT Sum(IIf(MoveType=1,Qty,-Qty)) FROM tblStockMoves WHERE MaterialID=" & materialID), 0)
End Function

Public Function StockOpen(frm As Form) As Boolean
    Dim rs As DAO.Recordset
    frm!txtDate = Date
    frm!cboMoveType = 2                         ' الصرف أكثر الحركات
    ' الموقع الافتراضي = آخر موقع استعمله الموظف
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 LocationID FROM tblStockMoves WHERE UserID=" & Nz(TempVars("UserID"), 0) & _
                                     " AND LocationID Is Not Null ORDER BY MoveID DESC", dbOpenSnapshot)
    If Not rs.EOF Then frm!cboLocation = rs!LocationID
    rs.Close
    StockOpen = True
End Function

' اختيار مادة من قائمة الأرصدة يملأ خانة المادة ويعرض حركاتها
Public Function StockPick(frm As Form) As Boolean
    On Error GoTo NoParent                      ' أثناء تحميل النموذج الأب لا يكون جاهزاً بعد
    If IsNull(frm!materialID) Then Exit Function
    frm.Parent!cboMaterial = frm!materialID
    StockPick = True
    Exit Function
NoParent:
End Function

Public Function StockSave(frm As Form) As Boolean
    Dim qty As Double, bal As Double, moveType As Long

    If IsNull(frm!cboMaterial) Then Msg "اختر المادة.", vbExclamation: Exit Function
    If IsNull(frm!cboMoveType) Then Msg "اختر الحركة: إضافة أو صرف.", vbExclamation: Exit Function
    If Not IsNumeric(frm!txtQty) Then Msg "اكتب الكمية رقماً.", vbExclamation: Exit Function
    qty = CDbl(frm!txtQty)
    If qty <= 0 Then Msg "الكمية أكبر من صفر.", vbExclamation: Exit Function
    If Not IsDate(frm!txtDate) Then Msg "اكتب التاريخ.", vbExclamation: Exit Function
    If Not IsNull(frm!txtExpiry) And Not IsDate(frm!txtExpiry) Then Msg "تاريخ الانتهاء غير صحيح.", vbExclamation: Exit Function

    moveType = frm!cboMoveType
    If moveType = 2 Then
        bal = StockBalance(frm!cboMaterial)
        If qty > bal Then
            Msg "لا يمكن صرف " & qty & ": الرصيد الحالي " & bal & ".", vbExclamation
            Exit Function
        End If
    End If

    CurrentDb.Execute "INSERT INTO tblStockMoves (MaterialID, MoveType, Qty, MoveDate, ExpiryDate, LotNo, LocationID, UserID, Notes) VALUES (" & _
        frm!cboMaterial & ", " & moveType & ", " & Str(qty) & ", " & SqlDateV(frm!txtDate) & ", " & SqlDateV(frm!txtExpiry) & ", " & _
        SqlTextV(frm!txtLot) & ", " & IIf(IsNull(frm!cboLocation), "Null", frm!cboLocation) & ", " & Nz(TempVars("UserID"), "Null") & ", " & _
        SqlTextV(frm!txtNotes) & ")", dbFailOnError

    frm!lblStatus.Caption = "سُجّلت: " & IIf(moveType = 1, "إضافة ", "صرف ") & qty & " - الرصيد الآن " & StockBalance(frm!cboMaterial)
    frm!txtQty = Null: frm!txtExpiry = Null: frm!txtLot = Null: frm!txtNotes = Null
    frm!sfBalance.Requery
    frm!sfMoves.Requery
    StockSave = True
End Function

Public Function StockMoveDelete(frm As Form) As Boolean
    Dim qty As Double, matID As Long
    If IsNull(frm!MoveID) Then Exit Function
    matID = frm!materialID
    qty = DbVal(CurrentDb, "SELECT Qty FROM tblStockMoves WHERE MoveID=" & frm!MoveID)
    If frm!moveType = 1 And StockBalance(matID) - qty < 0 Then
        Msg "لا يمكن حذف هذه الإضافة: صُرف منها، والرصيد سيصبح سالباً.", vbExclamation
        Exit Function
    End If
    If Not Ask("حذف هذه الحركة؟") Then Exit Function
    CurrentDb.Execute "DELETE FROM tblStockMoves WHERE MoveID=" & frm!MoveID, dbFailOnError
    frm.Requery
    frm.Parent!sfBalance.Requery
    StockMoveDelete = True
End Function

Public Function OpenMaterials() As Boolean
    DoCmd.OpenForm "frmMaterials"
    OpenMaterials = True
End Function

Public Function MaterialsClose(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    DoCmd.Close acForm, frm.Name
    If CurrentProject.AllForms("frmStock").IsLoaded Then
        Forms("frmStock")!sfBalance.Requery
        Forms("frmStock")!cboMaterial.Requery
    End If
    MaterialsClose = True
End Function

Private Function SqlDateV(ByVal v As Variant) As String
    If IsNull(v) Or Not IsDate(v) Then SqlDateV = "Null" Else SqlDateV = "#" & Format(CDate(v), "yyyy-mm-dd") & "#"
End Function

Private Function SqlTextV(ByVal v As Variant) As String
    If Len(Trim$(Nz(v, ""))) = 0 Then SqlTextV = "Null" Else SqlTextV = Q(Trim$(v))
End Function

' مجموع «العدد» لعينات التشغيلة
Private Function RunItems(db As DAO.Database, ByVal runID As Long) As Long
    RunItems = Nz(DbVal(db, "SELECT Sum(s.ItemCount) FROM tblRunSamples AS rs INNER JOIN tblSamples AS s " & _
                            "ON rs.SampleID = s.SampleID WHERE rs.RunID=" & runID), 0)
End Function

' ===========================================================
' رقم الأدلة الجنائية: الخانة تحفظ الرقم فقط (312312)
'   الصيغة الكاملة 312312/2026-ظ تُركَّب في التقرير وحده (ForensicText)
'   إذا لُصقت الصيغة الكاملة يؤخذ الرقم الذي قبل «/»
' ===========================================================
Public Function ForensicChanged(frm As Form) As Boolean
    Dim v As String, d As String, i As Long
    v = Trim$(Nz(frm!ForensicNo, ""))
    If Len(v) = 0 Then Exit Function
    If InStr(v, "/") > 0 Then v = Left$(v, InStr(v, "/") - 1)
    For i = 1 To Len(v)
        If Mid$(v, i, 1) Like "[0-9]" Then d = d & Mid$(v, i, 1)
    Next i
    If d <> Trim$(Nz(frm!ForensicNo, "")) Then frm!ForensicNo = d
    ForensicChanged = True
End Function

' الصيغة الرسمية للتقرير: الرقم/سنة تاريخ القضية-ظ
Public Function ForensicText(ByVal no As Variant, ByVal caseDate As Variant) As String
    If Len(Trim$(Nz(no, ""))) = 0 Then Exit Function
    If InStr(no, "/") > 0 Then ForensicText = no: Exit Function      ' بيانات قديمة بالصيغة الكاملة
    ForensicText = Trim$(no) & "/" & Year(Nz(caseDate, Date)) & "-ظ"
End Function

' ===========================================================
' المستخدمون وتواقيعهم (مسؤول البرنامج فقط)
'   صورة التوقيع تُنسخ إلى <مجلد البرنامج>\assets\signatures\<اسم الدخول>.<الامتداد>
'   ويُحفظ مسارها نسبياً، فتعمل على كل جهاز يُنقل إليه البرنامج مع مجلد assets
' ===========================================================
Public Function OpenUsers() As Boolean
    If Not Nz(DLookup("IsAdmin", "tblUsers", "UserID=" & Nz(TempVars("UserID"), 0)), False) Then
        Msg "شاشة المستخدمين لمسؤول البرنامج فقط.", vbExclamation
        Exit Function
    End If
    DoCmd.OpenForm "frmUsers"
    OpenUsers = True
End Function

Public Function PickSignature(frm As Form) As Boolean
    Dim fd As Object, fso As Object
    Dim src As String, ext As String, dstDir As String, dst As String

    If Not SaveRecord(frm) Then Exit Function
    If IsNull(frm!winUser) Then
        Msg "اكتب اسم الدخول للمستخدم أولاً.", vbExclamation
        Exit Function
    End If
    Set fd = Application.FileDialog(3)
    fd.title = "صورة توقيع " & Nz(frm!FullName, frm!winUser)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "الصور", "*.png;*.jpg;*.jpeg;*.bmp"
    If fd.Show <> -1 Then Exit Function

    src = fd.SelectedItems(1)
    ext = LCase$(Mid$(src, InStrRev(src, ".")))
    Set fso = CreateObject("Scripting.FileSystemObject")
    dstDir = CurrentProject.path & "\assets"
    If Not fso.FolderExists(dstDir) Then fso.CreateFolder dstDir
    dstDir = dstDir & "\signatures"
    If Not fso.FolderExists(dstDir) Then fso.CreateFolder dstDir
    dst = dstDir & "\" & frm!winUser & ext
    fso.CopyFile src, dst, True
    frm!SignaturePath = "assets\signatures\" & frm!winUser & ext
    PickSignature = SaveRecord(frm)
End Function

' مسار مكتوب يدوياً: «نسخ كمسار» في Windows يضيف علامتي تنصيص، وقد يُلصق أكثر من مسار.
' يُؤخذ أول سطر بلا علامات تنصيص ولا مسافات زائدة.
Public Function CleanPath(ByVal s As String) As String
    Dim cut As Long
    s = Replace(s, vbCr, vbLf)
    cut = InStr(1, s, vbLf)
    Do While cut = 1                         ' أسطر فارغة في البداية
        s = Mid$(s, 2)
        cut = InStr(1, s, vbLf)
    Loop
    If cut > 0 Then s = Left$(s, cut - 1)
    CleanPath = Trim$(Replace(s, """", ""))
End Function

' المسار الكامل لصورة التوقيع، أو "" إذا لا توجد:
'   1) SignaturePath (نسبي إلى مجلد البرنامج أو كامل)  2) assets\signatures\<اسم الدخول>.png/.jpg
Public Function SignatureFile(ByVal userID As Long) As String
    Dim fso As Object, rs As DAO.Recordset
    Dim p As String, win As String, ext As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rs = CurrentDb.OpenRecordset("SELECT WinUser, SignaturePath FROM tblUsers WHERE UserID=" & userID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    p = CleanPath(Nz(rs!SignaturePath, ""))
    win = Nz(rs!winUser, "")
    rs.Close
    If Len(p) > 0 Then
        If Not (p Like "?:\*" Or p Like "\\*") Then p = CurrentProject.path & "\" & p
        If fso.FileExists(p) Then SignatureFile = p: Exit Function
    End If
    For Each ext In Array(".png", ".jpg", ".jpeg", ".bmp")
        p = CurrentProject.path & "\assets\signatures\" & win & ext
        If fso.FileExists(p) Then SignatureFile = p: Exit Function
    Next
End Function

' ===========================================================
' أدوات
' ===========================================================
Private Function DbVal(db As DAO.Database, ByVal sql As String) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    DbVal = rs.Fields(0).Value
    rs.Close
End Function

Public Function Q(ByVal s As String) As String
    Q = "'" & Replace(s, "'", "''") & "'"
End Function
