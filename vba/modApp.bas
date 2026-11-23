Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

' ===========================================================
' »—‰«„Ã ﬁ”„ «·≈ŸÂ«— «·ﬂÌ„Ì«∆Ì - „‰ÿﬁ «· ‘€Ì· Ê«·‘«‘« 
' «·≈’œ«— 2 - 22 ‰Ê›„»— 2026 („” ÊÏ «·„Ã„Ê⁄« )
'
' «·‰„«–Ã ·«  Õ ÊÌ ﬂÊœ«. ﬂ· ÕœÀ ›Ì «·‰„Ê–Ã Ì” œ⁄Ì œ«·… Â‰«
' „»«‘—… „‰ Œ«’Ì… «·ÕœÀ° „À·:  =SampleAfterUpdate([Form])
'
' «·Õ·ﬁ…:
'   «·ﬁ÷Ì… (tblCases) -> «·„Ã„Ê⁄«  (tblBatches: ﬂ· «” ·«„ »„ÊŸ›Â Ê Ê«—ÌŒÂ Ê ﬁ—Ì—Â)
'     -> «·⁄Ì‰«  (tblSamples) -> «·¬À«— (tblTraces)
'   «·¬À«—  ı‰‘√  ·ﬁ«∆Ì« »⁄œœ «·¬À«— «·≈ÌÃ«»Ì… (SyncTraces)° Ê·« Ìı÷Ì›Â« «·„ÊŸ› »ÌœÂ.
' ===========================================================

Public gSilent As Boolean        ' ·· ‘€Ì· «·¬·Ì: «·—”«∆·  ıÃ„⁄ ›Ì gLog »œ· «·‰Ê«›–
Public gLog As String

Private Const APP_NAME As String = "ﬁ”„ «·≈ŸÂ«— «·ﬂÌ„Ì«∆Ì"
Private Const PROC_TRACE_DESC As String = "√À— „⁄«·Ã… —ﬁ„ "
Private Const VISIT_TRACE_DESC As String = "√À— „’Ê— —ﬁ„ "

' ===========================================================
' «·—”«∆·
' ===========================================================
Public Sub Msg(ByVal text As String, Optional ByVal style As VbMsgBoxStyle = vbInformation)
    If gSilent Then
        gLog = gLog & text & vbCrLf
    Else
        MsgBox text, style + vbMsgBoxRight + vbMsgBoxRtlReading, APP_NAME
    End If
End Sub

' ›Ì «·Ê÷⁄ «·’«„  «·ÃÊ«» ´·«ª œ«∆„«
Public Function Ask(ByVal text As String) As Boolean
    If gSilent Then
        gLog = gLog & "[”ƒ«·° «·ÃÊ«» ·«] " & text & vbCrLf
    Else
        Ask = (MsgBox(text, vbYesNo + vbQuestion + vbDefaultButton2 + vbMsgBoxRight + vbMsgBoxRtlReading, APP_NAME) = vbYes)
    End If
End Function

' ··«Œ »«— «·¬·Ì ›ﬁÿ: Ì›⁄¯·/ÌÊﬁ› «·Ê÷⁄ «·’«„  ÊÌ⁄Ìœ «·—”«∆· «·„Ã„¯⁄… À„ Ì›—¯€Â«
Public Function SetSilent(ByVal silentOn As Boolean) As String
    gSilent = silentOn
    SetSilent = gLog
    gLog = ""
End Function

' ··«Œ »«— «·¬·Ì: √Ê· ﬁÌ„… „‰ «” ⁄·«„ Ìﬁ—√ [TempVars]! (DAO ·« ÌÕ·¯Â« ÊÕœÂ° › ı„·√ «·„⁄·„«  »‹ Eval)
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
    Msg featureName & ": ”Ìı»‰Ï ›Ì „—Õ·… ·«Õﬁ….", vbInformation
End Function

' ===========================================================
' «·œŒÊ·: «”„ „” Œœ„ «·ÃÂ«“ „ﬁ«»· tblUsers
' ===========================================================
Public Function SignIn(Optional ByVal winUser As String = "") As Boolean
    Dim rs As DAO.Recordset
    If Len(winUser) = 0 Then winUser = Environ$("USERNAME")
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT UserID, FullName FROM tblUsers WHERE WinUser=" & Q(winUser) & " AND IsActive=True", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Msg "«”„ «·„” Œœ„ ´" & winUser & "ª €Ì— „”Ã· ›Ì «·»—‰«„Ã." & vbCrLf & _
            "—«Ã⁄ „”ƒÊ· «·»—‰«„Ã ·≈÷«› ﬂ.", vbCritical
        Exit Function
    End If
    TempVars.Add "UserID", CLng(rs!userID.Value)
    TempVars.Add "UserName", CStr(rs!FullName.Value)
    TempVars.Add "WinUser", winUser
    rs.Close
    SignIn = True
End Function

' ===========================================================
' «·‘«‘… «·—∆Ì”Ì… («·ﬁ÷Ì…)
' ===========================================================
Public Function MainLoad(frm As Form) As Boolean
    If Not SignIn() Then
        If Not gSilent Then DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    frm!lblWelcome.Caption = "„—Õ»«° " & TempVars("UserName")
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
    If frm.Recordset.NoMatch Then Msg "·„ Ìı⁄À— ⁄·Ï «·ﬁ÷Ì….", vbExclamation Else CaseFind = True
End Function

' ÌÕ›Ÿ «·”Ã· «·Õ«·Ì∫ Ì⁄Ìœ False ÊÌ⁄—÷ «·”»» ≈–« —›÷ Â ﬁÊ«⁄œ «·ÃœÊ·
Public Function SaveRecord(frm As Form) As Boolean
    On Error GoTo EH
    If frm.Dirty Then frm.Dirty = False
    SaveRecord = True
    Exit Function
EH:
    Msg "·„ ÌıÕ›Ÿ «·”Ã·:" & vbCrLf & Err.Description, vbExclamation
End Function

' «·ﬁ÷Ì… ÌÃ» √‰  ﬂÊ‰ „Õ›ÊŸ… ﬁ»· «·«‰ ﬁ«· ·⁄Ì‰« Â«
Private Function CaseReady(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If frm.NewRecord Or IsNull(frm!caseID) Then
        Msg "√œŒ· »Ì«‰«  «·ﬁ÷Ì… Ê«Õ›ŸÂ« √Ê·«.", vbExclamation
        Exit Function
    End If
    CaseReady = True
End Function

' ===========================================================
' «·„Ã„Ê⁄« : ﬂ· «” ·«„ („ÊŸ› + „Ã„Ê⁄… ⁄Ì‰« ) ·Â  Ê«—ÌŒÂ Ê ﬁ—Ì—Â Ê”Ã·  ‘€Ì·Â
' ===========================================================
' «” ·«„ ÃœÌœ: «·„” Œœ„ ÂÊ «·›«Õ’°  «—ÌŒ «·Ê—Êœ «·ÌÊ„° À„  ı› Õ ⁄Ì‰« Â „»«‘—…
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

' “— ´› Õª ›Ì ”ÿ— «·„Ã„Ê⁄…
Public Function BatchOpen(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If IsNull(frm!batchID) Then Exit Function
    DoCmd.OpenForm "frmSamples", , , "BatchID=" & frm!batchID, , , CStr(frm!batchID)
    BatchOpen = True
End Function

' scope: 0 = ¬À«— «·ﬁ÷Ì… „‰ «·—∆Ì”Ì…° 1 = ¬À«— ⁄Ì‰… Ê«Õœ…° 2 = ¬À«— «·„Ã„Ê⁄… „‰ ‘«‘… «·⁄Ì‰« 
Public Function OpenTraces(frm As Form, ByVal scope As Integer) As Boolean
    Dim where As String
    Select Case scope
        Case 0
            If Not CaseReady(frm) Then Exit Function
            where = "CaseID=" & frm!caseID
        Case 1
            If Not SaveRecord(frm) Then Exit Function
            If frm.NewRecord Or IsNull(frm!sampleID) Then
                Msg "«Õ›Ÿ ”ÿ— «·⁄Ì‰… √Ê·«.", vbExclamation
                Exit Function
            End If
            where = "SampleID=" & frm!sampleID
        Case Else
            If Not SaveRecord(frm) Then Exit Function
            where = "BatchID=" & Nz(frm.OpenArgs, 0)
    End Select

    If DCount("*", "qTracesEdit", where) = 0 Then
        Msg "·«  ÊÃœ ¬À«—." & vbCrLf & _
            "«·¬À«—  ı‰‘√  ·ﬁ«∆Ì« ⁄‰œ «Œ Ì«— ´≈ÌÃ«»Ì…ª Ê ÕœÌœ ⁄œœ «·¬À«— ›Ì ‘«‘… «·⁄Ì‰« .", vbInformation
        Exit Function
    End If
    DoCmd.OpenForm "frmTraces", , , where
    OpenTraces = True
End Function

' ”Ã· «· ‘€Ì· Ìı› Õ „‰ ‘«‘… «·⁄Ì‰«  Ê„⁄Â —ﬁ„ «·„Ã„Ê⁄…
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
' ‘«‘… «·⁄Ì‰« 
' ===========================================================
Public Function SamplesOpen(frm As Form) As Boolean
    Dim batchID As Long
    Dim rs As DAO.Recordset

    batchID = CLng(Nz(frm.OpenArgs, 0))
    If batchID = 0 Then
        Msg "«› Õ ‘«‘… «·⁄Ì‰«  „‰ „Ã„Ê⁄«  «·ﬁ÷Ì… ›Ì «·‘«‘… «·—∆Ì”Ì….", vbExclamation
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT c.RefNo, c.ForensicNo, b.BatchNo, b.ReceivedDate, u.FullName " & _
        "FROM (tblBatches AS b INNER JOIN tblCases AS c ON b.CaseID = c.CaseID) " & _
        "INNER JOIN tblUsers AS u ON b.ExaminerID = u.UserID WHERE b.BatchID=" & batchID, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!lblCase.Caption = "—ﬁ„ «·„—Ã⁄: " & rs!refNo & "     «·√œ·… «·Ã‰«∆Ì…: " & Nz(rs!ForensicNo, "-") & _
                              "     «·„Ã„Ê⁄… " & rs!BatchNo & ": " & rs!FullName & " - " & Format(rs!ReceivedDate, "yyyy/mm/dd")
    End If
    rs.Close

    ' «·ﬁÌ„ «·«› —«÷Ì… ··”ÿ— «·ÃœÌœ
    frm!batchID.DefaultValue = batchID
    frm!MethodCode.DefaultValue = """" & Nz(DLookup("MethodCode", "tblMethods", "IsDefault=True"), "") & """"
    frm!TransferTargetID.DefaultValue = Nz(DLookup("TargetID", "tblTransferTargets", "IsDefault=True"), "Null")
    SamplesOpen = True
End Function

' —ﬁ„ «·”ÿ— «· «·Ì ›Ì «·„Ã„Ê⁄… («· —ﬁÌ„ Ì»œ√ „‰ 1 ›Ì ﬂ·  ﬁ—Ì—)
Public Function SampleBeforeInsert(frm As Form) As Boolean
    frm!SampleNo = Nz(DMax("SampleNo", "tblSamples", "BatchID=" & frm!batchID), 0) + 1
    SampleBeforeInsert = True
End Function

' ⁄Ì‰… (1) √Ê «‰ ﬁ«· (2)
Public Function SampleTypeChanged(frm As Form) As Boolean
    ' «·«‰ ﬁ«·: ÕﬁÊ· «·⁄Ì‰…  ı›—Û¯€ Ê ıﬁ›· ( ‰”Ìﬁ ‘—ÿÌ)° Ê·« Ì»ﬁÏ ≈·« ⁄œœ «·¬À«— «·„—›Ê⁄…
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

' «·⁄œœ = ≈ÌÃ«»Ì + ”·»Ì + ·„ Ì »Ì‰: ´·„ Ì »Ì‰ª ÌıÕ”» ÊÕœÂ „‰ «·»«ﬁÌ
' Ê·« ¬À«— »·« ≈ÌÃ«»Ì
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

' »⁄œ √Ì  €ÌÌ— ›Ì «·⁄Ì‰« : €Ì— «·„‘ —ﬂ…  ıÕœÛ¯À  ·ﬁ«∆Ì«° Ê«·„‘ —ﬂ…  ı‰»Û¯Â ≈–« «Œ · «· ÿ«»ﬁ
Private Sub StaffAfterSampleChange(ByVal batchID As Long)
    If batchID = 0 Then Exit Sub
    If Not SyncBatchStaff(batchID) Then
        Msg "Â–Â „Ã„Ê⁄… „‘ —ﬂ… Ê €Ì¯—  ⁄Ì‰« Â«:" & vbCrLf & BatchStaffDiff(batchID) & vbCrLf & _
            "⁄œ¯· «· Ê“Ì⁄ „‰ “— ´ Ê“Ì⁄ «·⁄„·ª.", vbExclamation
    End If
End Sub

' ===========================================================
'  Ê“Ì⁄ «·⁄„· (tblBatchStaff) - „À· ”ÿ— Tests «·ﬁœÌ„° ·ﬂ· „ÊŸ› ›Ì «·„Ã„Ê⁄…
'   - €Ì— „‘ —ﬂ… (”ÿ— Ê«Õœ √Ê ·« ”ÿ—): ”ÿ— «·›«Õ’ = „Ã«„Ì⁄ «·⁄Ì‰« °  ·ﬁ«∆Ì«
'   - „‘ —ﬂ… (√ﬂÀ— „‰ ”ÿ—): ÌÊ“¯⁄Â« «·„ÊŸ› »ÌœÂ° Ê«·„Ã„Ê⁄ ÌÃ» √‰ Ìÿ«»ﬁ «·⁄Ì‰« 
'   - «·√⁄œ«œ »«·√Ê—«ﬁ/«·ﬁÿ⁄: ⁄Ì‰… „‰ 100 Ê—ﬁ… = 100 (≈ÌÃ«»Ì + ”·»Ì + ·„ Ì »Ì‰)
'   - «·›ÕÊ’ ·«  ıŒ“Û¯‰: ·ﬂ·  ‘€Ì·… œŒ· Â« «·⁄Ì‰… = (≈ÌÃ«»Ì + ”·»Ì) ›Õ’«°  ı‰”» ·„ÊŸ› «· ‘€Ì·…
'     (qRunWork.TestItems). ›ÕÊ’ ´·„ Ì »Ì‰ª ·«  œŒ· ›Ì «·„Ã„Ê⁄° ﬂ„« ›Ì Tests «·ﬁœÌ„.
' ===========================================================
' „Ã«„Ì⁄ «·„Ã„Ê⁄… „‰ «·⁄Ì‰« : ≈ÌÃ«»Ì° ”·»Ì° ·„ Ì »Ì‰° ¬À«— «·„⁄«·Ã…° ¬À«— «·«‰ ﬁ«·
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

' Ì⁄Ìœ False ≈–« ﬂ«‰  «·„Ã„Ê⁄… „‘ —ﬂ… Ê„Ã„Ê⁄Â« ·« Ìÿ«»ﬁ «·⁄Ì‰« 
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

' ‰’ «·›—Êﬁ »Ì‰ «· Ê“Ì⁄ Ê«·⁄Ì‰« ° √Ê "" ≈–«  ÿ«»ﬁ«
Public Function BatchStaffDiff(ByVal batchID As Long) As String
    Dim db As DAO.Database
    Dim t As Variant, d As Variant, names As Variant
    Dim i As Long, s As String
    Set db = CurrentDb
    t = BatchTotals(db, batchID)
    d = StaffTotals(db, batchID)
    names = Array("«·≈ÌÃ«»Ì…", "«·”·»Ì…", "·„ Ì »Ì‰", "¬À«— «·„⁄«·Ã…", "¬À«— «·«‰ ﬁ«·")
    For i = 0 To 4
        If t(i) <> d(i) Then s = s & "- " & names(i) & ": ›Ì «·⁄Ì‰«  " & t(i) & "° «·„Ê“Û¯⁄ " & d(i) & vbCrLf
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
    frm!lblTotals.Caption = "«·⁄Ì‰« :   ≈ÌÃ«»Ì… " & t(0) & "   ”·»Ì… " & t(1) & "   ·„ Ì »Ì‰ " & t(2) & _
                            "   ¬À«— «·„⁄«·Ã… " & t(3) & "   ¬À«— «·«‰ ﬁ«· " & t(4) & vbCrLf & _
                            "«·„Ê“Û¯⁄:  ≈ÌÃ«»Ì… " & d(0) & "   ”·»Ì… " & d(1) & "   ·„ Ì »Ì‰ " & d(2) & _
                            "   ¬À«— «·„⁄«·Ã… " & d(3) & "   ¬À«— «·«‰ ﬁ«· " & d(4) & _
                            IIf(Len(diff) = 0, "     („ÿ«»ﬁ)", "     (€Ì— „ÿ«»ﬁ)")
    frm!lblTotals.ForeColor = IIf(Len(diff) = 0, RGB(16, 128, 88), RGB(192, 0, 0))
    StaffRefresh = True
End Function

' «·≈€·«ﬁ ›ﬁÿ ≈–« ÿ«»ﬁ «· Ê“Ì⁄ «·⁄Ì‰« 
Public Function StaffClose(frm As Form) As Boolean
    Dim diff As String
    If Not SaveRecord(frm) Then Exit Function
    diff = BatchStaffDiff(CLng(Nz(frm.OpenArgs, 0)))
    If Len(diff) > 0 Then
        Msg "«· Ê“Ì⁄ ·« Ìÿ«»ﬁ ⁄Ì‰«  «·„Ã„Ê⁄…:" & vbCrLf & diff, vbExclamation
        StaffRefresh frm
        Exit Function
    End If
    DoCmd.Close acForm, frm.Name
    StaffClose = True
End Function

' ===========================================================
' ≈‰‘«¡ «·¬À«—  ·ﬁ«∆Ì«
'   - «·⁄Ì‰… «·≈ÌÃ«»Ì…: ⁄œœ √”ÿ— «·¬À«— = ⁄œœ «·¬À«—
'   - ”ÿ— «·«‰ ﬁ«·: ⁄œœ √”ÿ— «·¬À«— = ⁄œœ «·¬À«— «·„’Ê—…° Ê »œ√ „‰ Õ«·… 2
'   - «·≈‰ﬁ«’ ·« ÌÕ–› √À—« „”Ã·« ›Ì „‰’… √À—
'   - —„“ «·√À— = »«—ﬂÊœ «·⁄Ì‰…-—ﬁ„ «·√À—° „À· 900000101-01
'     «·«‰ ﬁ«· »·« »«—ﬂÊœ: —ﬁ„ «·„—Ã⁄-—ﬁ„ «·„Ã„Ê⁄… V —ﬁ„ «·”ÿ—-—ﬁ„ «·√À—° „À· 226000000001-1V2-01
' ===========================================================
Public Function SyncTraces(ByVal sampleID As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim want As Long, have As Long, extra As Long, removable As Long, nextNo As Long
    Dim baseCode As String, descPrefix As String, firstStatus As Long, liftDefault As String

    DBEngine.Idle dbRefreshCache        ' ﬁ—«¡… ¬Œ— „« Õı›Ÿ° ·« ‰”Œ… «·„Õ—ﬂ «·„ƒﬁ …
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
        ' „ﬂ«‰ —›⁄ «·√À— «·«› —«÷Ì: „‰ ⁄·Ï Ê’› «·⁄Ì‰…
        If Not IsNull(rs!Description) Then liftDefault = "„‰ ⁄·Ï " & rs!Description
    End If
    rs.Close

    ' «·¬À«— «· Ì ·„ ÌıÕœÛ¯œ „ﬂ«‰ —›⁄Â«  √Œ– «·«› —«÷Ì («·„⁄œÛ¯· ÌœÊÌ« Ì»ﬁÏ)
    If Len(liftDefault) > 0 Then
        db.Execute "UPDATE tblTraces SET LiftLocation = " & Q(liftDefault) & _
                   " WHERE SampleID=" & sampleID & " AND LiftLocation Is Null", dbFailOnError
    End If

    '  ÕœÌÀ «·—„Ê“ ≈–«  €Ì¯— «·»«—ﬂÊœ («·Ê’› «·«› —«÷Ì Ì »⁄ «·—„“° Ê«·Ê’› «·„⁄œ¯· Ì»ﬁÏ)
    db.Execute "UPDATE tblTraces SET Description = " & Q(descPrefix & baseCode & "-") & " & Format(TraceNo,'00') " & _
               "WHERE SampleID=" & sampleID & " AND Description = " & Q(descPrefix) & " & TraceCode", dbFailOnError
    db.Execute "UPDATE tblTraces SET TraceCode = " & Q(baseCode & "-") & " & Format(TraceNo,'00') " & _
               "WHERE SampleID=" & sampleID, dbFailOnError

    ' «·⁄œ¯ ⁄»— db ‰›”Â ·« DCount/DMax: œÊ«· «·‰ÿ«ﬁ  ⁄„· ›Ì Ã·”… „‰›’·…
    have = DbVal(db, "SELECT Count(*) FROM tblTraces WHERE SampleID=" & sampleID)

    ' “Ì«œ…
    Do While have < want
        nextNo = Nz(DbVal(db, "SELECT Max(TraceNo) FROM tblTraces WHERE SampleID=" & sampleID), 0) + 1
        db.Execute "INSERT INTO tblTraces (SampleID, TraceNo, TraceCode, Description, LiftLocation, StatusID) VALUES (" & _
                   sampleID & ", " & nextNo & ", " & Q(baseCode & "-" & Format(nextNo, "00")) & ", " & _
                   Q(descPrefix & baseCode & "-" & Format(nextNo, "00")) & ", " & _
                   IIf(Len(liftDefault) > 0, Q(liftDefault), "Null") & ", " & firstStatus & ")", dbFailOnError
        have = have + 1
    Loop

    ' ≈‰ﬁ«’: „‰ ¬Œ— «·√—ﬁ«„° Ê›ﬁÿ «·¬À«— €Ì— «·„”Ã·… ›Ì „‰’… √À—
    If have > want Then
        extra = have - want
        removable = DbVal(db, "SELECT Count(*) FROM tblTraces WHERE SampleID=" & sampleID & " AND AtharBarcode Is Null")
        If removable < extra Then
            Msg "⁄œœ «·¬À«— √ﬁ· „‰ «·¬À«— «·„”Ã·… ›Ì „‰’… √À— ·Â–Â «·⁄Ì‰…." & vbCrLf & _
                "·„ ÌıÕ–› √Ì √À—. ⁄œ¯· «·⁄œœ √Ê «Õ–› »«—ﬂÊœ √À— „‰ ‘«‘… «·¬À«— √Ê·«.", vbExclamation
            Exit Function
        End If
        If Not Ask("”Ì „ Õ–› " & extra & " „‰ ¬À«— Â–Â «·⁄Ì‰… («·√—ﬁ«„ «·√ŒÌ—…). „ «»⁄…ø") Then
            Msg "·„ ÌıÕ–› √Ì √À—. ⁄œœ «·¬À«— ›Ì «·⁄Ì‰… ·« Ìÿ«»ﬁ √”ÿ— «·¬À«— «·¬‰.", vbExclamation
            Exit Function
        End If
        db.Execute "DELETE FROM tblTraces WHERE TraceID IN (SELECT TOP " & extra & " TraceID FROM tblTraces " & _
                   "WHERE SampleID=" & sampleID & " AND AtharBarcode Is Null ORDER BY TraceNo DESC)", dbFailOnError
    End If
    SyncTraces = True
End Function

' ===========================================================
' ‘«‘… «·¬À«— («·Õ«·…   €Ì— »«·≈Ã—«¡° ·« ÌŒ «—Â« «·„ÊŸ›)
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
' «Œ Ì«— „·› ’Ê—…
' ===========================================================
Public Function PickFile(frm As Form, ByVal fieldName As String) As Boolean
    Dim fd As Object
    Set fd = Application.FileDialog(3)      ' msoFileDialogFilePicker
    fd.title = "«Œ — «·’Ê—…"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "«·’Ê—", "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff"
    If Len(Nz(TempVars("LastFolder"), "")) > 0 Then fd.InitialFileName = TempVars("LastFolder")
    If fd.Show = -1 Then
        frm(fieldName) = fd.SelectedItems(1)
        TempVars.Add "LastFolder", Left$(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), "\"))
        If fieldName = "PhotoPath" And frm.Name = "frmTraces" Then TracePhotoChanged frm
        PickFile = True
    End If
End Function

' ===========================================================
' ”Ã·  ‘€Ì· «·√ÃÂ“… (LOGBOOK)
' ===========================================================
Public Function RunsOpen(frm As Form) As Boolean
    Dim rs As DAO.Recordset
    ' «·ÃÂ«“ «·«› —«÷Ì = ¬Œ— ÃÂ«“ «” Œœ„Â «·„ÊŸ›
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 DeviceID FROM tblRuns WHERE UserID=" & Nz(TempVars("UserID"), 0) & " ORDER BY RunID DESC", dbOpenSnapshot)
    If Not rs.EOF Then frm!DeviceID.DefaultValue = rs!DeviceID
    rs.Close
    ' «·„Ã„Ê⁄… «· Ì ›ı Õ „‰Â« «·”Ã· (·“— ´ﬂ· ⁄Ì‰«  «·„Ã„Ê⁄…ª)
    TempVars.Add "RunBatchID", CLng(Nz(frm.OpenArgs, 0))
    frm.SetFocus
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    RunsOpen = True
End Function

' Ì÷Ì› ⁄Ì‰«  «·„Ã„Ê⁄… («·„⁄«·Ã… ›ﬁÿ° €Ì— «·„÷«›…) ≈·Ï «· ‘€Ì·… «·Õ«·Ì…
Public Function RunAddBatch(frm As Form) As Boolean
    Dim db As DAO.Database
    Dim batchID As Long, before As Long, added As Long

    batchID = Nz(TempVars("RunBatchID"), 0)
    If batchID = 0 Then
        Msg "«› Õ ”Ã· «· ‘€Ì· „‰ ‘«‘… ⁄Ì‰«  «·„Ã„Ê⁄….", vbExclamation
        Exit Function
    End If
    If IsNull(frm!DeviceID) Then
        Msg "«Œ — «·ÃÂ«“ √Ê·«.", vbExclamation
        Exit Function
    End If
    If Not SaveRecord(frm) Then Exit Function

    Set db = CurrentDb
    ' «·⁄œœ »„Ã„Ê⁄ ´«·⁄œœª ·« »⁄œœ «·√”ÿ—
    before = RunItems(db, frm!runID)
    db.Execute "INSERT INTO tblRunSamples (RunID, SampleID) SELECT " & frm!runID & ", SampleID FROM tblSamples " & _
               "WHERE BatchID=" & batchID & " AND LineType=1 AND SampleID NOT IN " & _
               "(SELECT SampleID FROM tblRunSamples WHERE RunID=" & frm!runID & ")", dbFailOnError
    added = RunItems(db, frm!runID) - before
    frm!sfRunSamples.Requery
    frm!txtCount.Requery
    Msg "√÷Ì›  ⁄Ì‰«  ⁄œœÂ« " & added & " ≈·Ï «· ‘€Ì·….", vbInformation
    RunAddBatch = True
End Function

Public Function RunNew(frm As Form) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    frm.SetFocus
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    RunNew = True
End Function

' which: "start" √Ê "end"
Public Function RunNow(frm As Form, ByVal which As String) As Boolean
    If which = "start" Then
        frm!StartTime = Now()
        RunSetEnd frm
    Else
        frm!EndTime = Now()
    End If
    RunNow = SaveRecord(frm)
End Function

' «Œ Ì«— «·ÃÂ«“: «·»œ¡ «·¬‰ (≈‰ ·„ ÌıÕœÛ¯œ)° Ê«·«‰ Â«¡ = «·»œ¡ + „œ… œÊ—… «·ÃÂ«“
Public Function RunDeviceChanged(frm As Form) As Boolean
    If IsNull(frm!StartTime) Then frm!StartTime = Now()
    RunDeviceChanged = RunSetEnd(frm)
End Function

' Êﬁ  «·«‰ Â«¡ = Êﬁ  «·»œ¡ + „œ… «·ÃÂ«“ (tblDevices.DurationMin)
Public Function RunSetEnd(frm As Form) As Boolean
    Dim mins As Long
    If IsNull(frm!DeviceID) Or IsNull(frm!StartTime) Then Exit Function
    mins = Nz(DLookup("DurationMin", "tblDevices", "DeviceID=" & frm!DeviceID), 0)
    If mins > 0 Then frm!EndTime = DateAdd("n", mins, frm!StartTime)
    RunSetEnd = True
End Function

' „”Õ »«—ﬂÊœ ⁄Ì‰… Ì÷Ì›Â« ·· ‘€Ì·…
Public Function RunScan(frm As Form) As Boolean
    Dim code As String
    Dim sampleID As Variant

    code = Trim$(Nz(frm!txtScan, ""))
    frm!txtScan = Null
    If Len(code) = 0 Then Exit Function
    If IsNull(frm!DeviceID) Then
        Msg "«Œ — «·ÃÂ«“ √Ê·«.", vbExclamation
        Exit Function
    End If
    If Not SaveRecord(frm) Then Exit Function

    sampleID = DLookup("SampleID", "tblSamples", "Barcode=" & Q(code))
    If IsNull(sampleID) Then
        Msg "·«  ÊÃœ ⁄Ì‰… »«·»«—ﬂÊœ " & code & "." & vbCrLf & "√œŒ· «·⁄Ì‰… ›Ì ‘«‘… «·⁄Ì‰«  √Ê·«.", vbExclamation
    ElseIf DCount("*", "tblRunSamples", "RunID=" & frm!runID & " AND SampleID=" & sampleID) > 0 Then
        Msg "«·⁄Ì‰… " & code & " „÷«›… ·Â–Â «· ‘€Ì·… „”»ﬁ«.", vbExclamation
    Else
        CurrentDb.Execute "INSERT INTO tblRunSamples (RunID, SampleID) VALUES (" & frm!runID & ", " & sampleID & ")", dbFailOnError
        RunScan = True
    End If
    frm!sfRunSamples.Requery
    frm!txtCount.Requery
    frm!DeviceID.SetFocus           ' ··⁄Êœ… ≈·Ï Œ«‰… «·„”Õ »⁄œ AfterUpdate
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
' «·≈Õ’«∆Ì«  (√“—«—  –ÌÌ· «·—∆Ì”Ì…° ﬂ„« ›Ì «·»—‰«„Ã «·ﬁœÌ„)
'   «·› —… „‰ ‘—Ìÿ «· –ÌÌ·: «·”‰…° „‰ ‘Â—° ≈·Ï ‘Â— -> TempVars  ﬁ—ƒÂ« «·«” ⁄·«„« 
'   «·› —… Õ”»  «—ÌŒ Ê—Êœ «·„Ã„Ê⁄…. «·√⁄œ«œ »«·√Ê—«ﬁ° Ê«·›ÕÊ’ „‰ ”Ã· «· ‘€Ì·.
' ===========================================================
Public Function SetStatPeriod(ByVal yr As Long, ByVal m1 As Long, ByVal m2 As Long) As Boolean
    If yr < 2000 Or yr > 2100 Or m1 < 1 Or m1 > 12 Or m2 < 1 Or m2 > 12 Or m1 > m2 Then
        Msg "«·› —… €Ì— ’ÕÌÕ…: «·”‰… 4 √—ﬁ«„° Ê«·‘Â— „‰ 1 ≈·Ï 12° Ê´„‰ ‘Â—ª ·« Ì Ã«Ê“ ´≈·Ï ‘Â—ª.", vbExclamation
        Exit Function
    End If
    TempVars.Add "StatYear", yr
    TempVars.Add "StatFrom", m1
    TempVars.Add "StatTo", m2
    ' ÃœÊ· «·ÃÂ«  ◊ «·√‘Â— ·« Ìﬁ»· „ €Ì—« : Ìı⁄«œ »‰«ƒÂ »«·”‰…
    CurrentDb.QueryDefs("«Õ’«∆Ì… «·ﬁ÷«Ì« Õ”» «·ÃÂ« ").sql = _
        "TRANSFORM Count(c.CaseID) AS n SELECT IIf(a.AgencyName Is Null,'(»·« ÃÂ…)',a.AgencyName) AS [«·ÃÂ…], Count(c.CaseID) AS [«·≈Ã„«·Ì] " & _
        "FROM tblCases AS c LEFT JOIN tblAgencies AS a ON c.AgencyID = a.AgencyID WHERE Year(c.CaseDate)=" & yr & " " & _
        "GROUP BY IIf(a.AgencyName Is Null,'(»·« ÃÂ…)',a.AgencyName) " & _
        "PIVOT Format(c.CaseDate,'mm') IN ('01','02','03','04','05','06','07','08','09','10','11','12');"
    SetStatPeriod = True
End Function

Public Function OpenStat(frm As Form, ByVal queryName As String) As Boolean
    If Not SaveRecord(frm) Then Exit Function
    If Not IsNumeric(frm!txtStatYear) Or Not IsNumeric(frm!txtStatFrom) Or Not IsNumeric(frm!txtStatTo) Then
        Msg "«ﬂ » «·”‰… Ê«·‘Â—Ì‰ √—ﬁ«„«.", vbExclamation
        Exit Function
    End If
    If Not SetStatPeriod(CLng(frm!txtStatYear), CLng(frm!txtStatFrom), CLng(frm!txtStatTo)) Then Exit Function
    DoCmd.OpenQuery queryName, acViewNormal, acReadOnly
    OpenStat = True
End Function

' ===========================================================
' «·„Œ“Ê‰: «·—’Ìœ = «·≈÷«›«  - «·„’—Ê›«  (tblStockMoves)
'   «·’—› ·« Ì Ã«Ê“ «·—’Ìœ° ÊÕ–› ≈÷«›… ·« ÌÃ⁄· «·—’Ìœ ”«·»«.
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
    frm!cboMoveType = 2                         ' «·’—› √ﬂÀ— «·Õ—ﬂ« 
    ' «·„Êﬁ⁄ «·«› —«÷Ì = ¬Œ— „Êﬁ⁄ «” ⁄„·Â «·„ÊŸ›
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 LocationID FROM tblStockMoves WHERE UserID=" & Nz(TempVars("UserID"), 0) & _
                                     " AND LocationID Is Not Null ORDER BY MoveID DESC", dbOpenSnapshot)
    If Not rs.EOF Then frm!cboLocation = rs!LocationID
    rs.Close
    StockOpen = True
End Function

' «Œ Ì«— „«œ… „‰ ﬁ«∆„… «·√—’œ… Ì„·√ Œ«‰… «·„«œ… ÊÌ⁄—÷ Õ—ﬂ« Â«
Public Function StockPick(frm As Form) As Boolean
    On Error GoTo NoParent                      ' √À‰«¡  Õ„Ì· «·‰„Ê–Ã «·√» ·« ÌﬂÊ‰ Ã«Â“« »⁄œ
    If IsNull(frm!materialID) Then Exit Function
    frm.Parent!cboMaterial = frm!materialID
    StockPick = True
    Exit Function
NoParent:
End Function

Public Function StockSave(frm As Form) As Boolean
    Dim qty As Double, bal As Double, moveType As Long

    If IsNull(frm!cboMaterial) Then Msg "«Œ — «·„«œ….", vbExclamation: Exit Function
    If IsNull(frm!cboMoveType) Then Msg "«Œ — «·Õ—ﬂ…: ≈÷«›… √Ê ’—›.", vbExclamation: Exit Function
    If Not IsNumeric(frm!txtQty) Then Msg "«ﬂ » «·ﬂ„Ì… —ﬁ„«.", vbExclamation: Exit Function
    qty = CDbl(frm!txtQty)
    If qty <= 0 Then Msg "«·ﬂ„Ì… √ﬂ»— „‰ ’›—.", vbExclamation: Exit Function
    If Not IsDate(frm!txtDate) Then Msg "«ﬂ » «· «—ÌŒ.", vbExclamation: Exit Function
    If Not IsNull(frm!txtExpiry) And Not IsDate(frm!txtExpiry) Then Msg " «—ÌŒ «·«‰ Â«¡ €Ì— ’ÕÌÕ.", vbExclamation: Exit Function

    moveType = frm!cboMoveType
    If moveType = 2 Then
        bal = StockBalance(frm!cboMaterial)
        If qty > bal Then
            Msg "·« Ì„ﬂ‰ ’—› " & qty & ": «·—’Ìœ «·Õ«·Ì " & bal & ".", vbExclamation
            Exit Function
        End If
    End If

    CurrentDb.Execute "INSERT INTO tblStockMoves (MaterialID, MoveType, Qty, MoveDate, ExpiryDate, LotNo, LocationID, UserID, Notes) VALUES (" & _
        frm!cboMaterial & ", " & moveType & ", " & Str(qty) & ", " & SqlDateV(frm!txtDate) & ", " & SqlDateV(frm!txtExpiry) & ", " & _
        SqlTextV(frm!txtLot) & ", " & IIf(IsNull(frm!cboLocation), "Null", frm!cboLocation) & ", " & Nz(TempVars("UserID"), "Null") & ", " & _
        SqlTextV(frm!txtNotes) & ")", dbFailOnError

    frm!lblStatus.Caption = "”ıÃ¯· : " & IIf(moveType = 1, "≈÷«›… ", "’—› ") & qty & " - «·—’Ìœ «·¬‰ " & StockBalance(frm!cboMaterial)
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
        Msg "·« Ì„ﬂ‰ Õ–› Â–Â «·≈÷«›…: ’ı—› „‰Â«° Ê«·—’Ìœ ”Ì’»Õ ”«·»«.", vbExclamation
        Exit Function
    End If
    If Not Ask("Õ–› Â–Â «·Õ—ﬂ…ø") Then Exit Function
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

' „Ã„Ê⁄ ´«·⁄œœª ·⁄Ì‰«  «· ‘€Ì·…
Private Function RunItems(db As DAO.Database, ByVal runID As Long) As Long
    RunItems = Nz(DbVal(db, "SELECT Sum(s.ItemCount) FROM tblRunSamples AS rs INNER JOIN tblSamples AS s " & _
                            "ON rs.SampleID = s.SampleID WHERE rs.RunID=" & runID), 0)
End Function

' ===========================================================
' —ﬁ„ «·√œ·… «·Ã‰«∆Ì…: «·Œ«‰…  Õ›Ÿ «·—ﬁ„ ›ﬁÿ (312312)
'   «·’Ì€… «·ﬂ«„·… 312312/2026-Ÿ  ı—ﬂÛ¯» ›Ì «· ﬁ—Ì— ÊÕœÂ (ForensicText)
'   ≈–« ·ı’ﬁ  «·’Ì€… «·ﬂ«„·… ÌƒŒ– «·—ﬁ„ «·–Ì ﬁ»· ´/ª
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

' «·’Ì€… «·—”„Ì… ·· ﬁ—Ì—: «·—ﬁ„/”‰…  «—ÌŒ «·ﬁ÷Ì…-Ÿ
Public Function ForensicText(ByVal no As Variant, ByVal caseDate As Variant) As String
    If Len(Trim$(Nz(no, ""))) = 0 Then Exit Function
    If InStr(no, "/") > 0 Then ForensicText = no: Exit Function      ' »Ì«‰«  ﬁœÌ„… »«·’Ì€… «·ﬂ«„·…
    ForensicText = Trim$(no) & "/" & Year(Nz(caseDate, Date)) & "-Ÿ"
End Function

' ===========================================================
' «·„” Œœ„Ê‰ Ê Ê«ﬁÌ⁄Â„ („”ƒÊ· «·»—‰«„Ã ›ﬁÿ)
'   ’Ê—… «· ÊﬁÌ⁄  ı‰”Œ ≈·Ï <„Ã·œ «·»—‰«„Ã>\assets\signatures\<«”„ «·œŒÊ·>.<«·«„ œ«œ>
'   ÊÌıÕ›Ÿ „”«—Â« ‰”»Ì«° › ⁄„· ⁄·Ï ﬂ· ÃÂ«“ Ìı‰ﬁ· ≈·ÌÂ «·»—‰«„Ã „⁄ „Ã·œ assets
' ===========================================================
Public Function OpenUsers() As Boolean
    If Not Nz(DLookup("IsAdmin", "tblUsers", "UserID=" & Nz(TempVars("UserID"), 0)), False) Then
        Msg "‘«‘… «·„” Œœ„Ì‰ ·„”ƒÊ· «·»—‰«„Ã ›ﬁÿ.", vbExclamation
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
        Msg "«ﬂ » «”„ «·œŒÊ· ··„” Œœ„ √Ê·«.", vbExclamation
        Exit Function
    End If
    Set fd = Application.FileDialog(3)
    fd.title = "’Ê—…  ÊﬁÌ⁄ " & Nz(frm!FullName, frm!winUser)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "«·’Ê—", "*.png;*.jpg;*.jpeg;*.bmp"
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

' „”«— „ﬂ Ê» ÌœÊÌ«: ´‰”Œ ﬂ„”«—ª ›Ì Windows Ì÷Ì› ⁄·«„ Ì  ‰’Ì’° Êﬁœ Ìı·’ﬁ √ﬂÀ— „‰ „”«—.
' ÌıƒŒ– √Ê· ”ÿ— »·« ⁄·«„«   ‰’Ì’ Ê·« „”«›«  “«∆œ….
Public Function CleanPath(ByVal s As String) As String
    Dim cut As Long
    s = Replace(s, vbCr, vbLf)
    cut = InStr(1, s, vbLf)
    Do While cut = 1                         ' √”ÿ— ›«—€… ›Ì «·»œ«Ì…
        s = Mid$(s, 2)
        cut = InStr(1, s, vbLf)
    Loop
    If cut > 0 Then s = Left$(s, cut - 1)
    CleanPath = Trim$(Replace(s, """", ""))
End Function

' «·„”«— «·ﬂ«„· ·’Ê—… «· ÊﬁÌ⁄° √Ê "" ≈–« ·«  ÊÃœ:
'   1) SignaturePath (‰”»Ì ≈·Ï „Ã·œ «·»—‰«„Ã √Ê ﬂ«„·)  2) assets\signatures\<«”„ «·œŒÊ·>.png/.jpg
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
' √œÊ« 
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
