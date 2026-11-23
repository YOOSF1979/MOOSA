Attribute VB_Name = "modBuild"
Option Compare Database
Option Explicit

' ===========================================================
' ÈäÇÁ ÈÑäÇãÌ ŞÓã ÇáÅÙåÇÑ ÇáßíãíÇÆí İí ŞÇÚÏÉ ÈíÇäÇÊ İÇÑÛÉ
'
' ÇáØÑíŞÉ:
'   1. ÃäÔÆ ŞÇÚÏÉ ÈíÇäÇÊ ÃßÓÓ İÇÑÛÉ ÌÏíÏÉ æÇÍİÙåÇ.
'   2. Alt+F11 Ëã File > Import File æÇÎÊÑ åĞÇ Çáãáİ (modBuild.bas).
'   3. ÖÚ ÇáãÄÔÑ ÏÇÎá BuildApp æÇÖÛØ F5.
'   4. ÇÎÊÑ ãáİ ÇáÈÑäÇãÌ ÇáŞÏíã ÚäÏãÇ íõØáÈ (ÊõäÓÎ ãäå ÇáäãÇĞÌ ææÍÏÇÊ ÇáßæÏ).
'   5. ÈÚÏ ÇáÇäÊåÇÁ: Debug > Compile¡ Ëã ÇÍĞİ åĞå ÇáæÍÏÉ Åä ÔÆÊ.
'
' íäÔÆ: 19 ÌÏæáÇğ¡ 35 İåÑÓÇğ¡ 21 ÚáÇŞÉ¡ 109 ÓÌáÇğ¡ 21 ÇÓÊÚáÇãÇğ.
' ===========================================================

Private db As DAO.Database
Private mStage As String

Public Sub BuildApp()
    Dim src As String
    Set db = CurrentDb
    If TableExists("tblUsers") Then
        MsgBox "ŞÇÚÏÉ ÇáÈíÇäÇÊ åĞå İíåÇ ÇáÌÏÇæá ãÓÈŞÇğ. ÔÛøá ÇáÈäÇÁ İí ŞÇÚÏÉ ÈíÇäÇÊ İÇÑÛÉ ÌÏíÏÉ.", vbExclamation
        Exit Sub
    End If
    src = PickSource()
    If src = "" Then
        If MsgBox("áã ÊÎÊÑ ãáİ ÇáÈÑäÇãÌ ÇáŞÏíã¡ İáä ÊõäÓÎ ÇáäãÇĞÌ ææÍÏÇÊ ÇáßæÏ." & vbCrLf & "ãÊÇÈÚÉ ÇáÈäÇÁ ÈÏæäåÇ¿", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    DoCmd.Hourglass True
    mStage = "ÌÏæá tblUsers": T_tblUsers
    mStage = "ÌÏæá tblAgencies": T_tblAgencies
    mStage = "ÌÏæá tblIncidentTypes": T_tblIncidentTypes
    mStage = "ÌÏæá tblResults": T_tblResults
    mStage = "ÌÏæá tblLineTypes": T_tblLineTypes
    mStage = "ÌÏæá tblMethods": T_tblMethods
    mStage = "ÌÏæá tblDevices": T_tblDevices
    mStage = "ÌÏæá tblTransferTargets": T_tblTransferTargets
    mStage = "ÌÏæá tblTraceStatus": T_tblTraceStatus
    mStage = "ÌÏæá tblCases": T_tblCases
    mStage = "ÌÏæá tblBatches": T_tblBatches
    mStage = "ÌÏæá tblSamples": T_tblSamples
    mStage = "ÌÏæá tblBatchStaff": T_tblBatchStaff
    mStage = "ÌÏæá tblTraces": T_tblTraces
    mStage = "ÌÏæá tblRuns": T_tblRuns
    mStage = "ÌÏæá tblRunSamples": T_tblRunSamples
    mStage = "ÌÏæá tblStorageLocations": T_tblStorageLocations
    mStage = "ÌÏæá tblMaterials": T_tblMaterials
    mStage = "ÌÏæá tblStockMoves": T_tblStockMoves
    mStage = "ÈíÇäÇÊ tblUsers": D_tblUsers
    mStage = "ÈíÇäÇÊ tblAgencies": D_tblAgencies
    mStage = "ÈíÇäÇÊ tblIncidentTypes": D_tblIncidentTypes
    mStage = "ÈíÇäÇÊ tblResults": D_tblResults
    mStage = "ÈíÇäÇÊ tblLineTypes": D_tblLineTypes
    mStage = "ÈíÇäÇÊ tblMethods": D_tblMethods
    mStage = "ÈíÇäÇÊ tblDevices": D_tblDevices
    mStage = "ÈíÇäÇÊ tblTransferTargets": D_tblTransferTargets
    mStage = "ÈíÇäÇÊ tblTraceStatus": D_tblTraceStatus
    mStage = "ÈíÇäÇÊ tblStorageLocations": D_tblStorageLocations
    mStage = "ÈíÇäÇÊ tblMaterials": D_tblMaterials
    mStage = "ÇáÚáÇŞÇÊ": BuildRelations
    mStage = "ÇáÇÓÊÚáÇãÇÊ": BuildQueries
    If src <> "" Then ImportForms src
    mStage = "ÎÕÇÆÕ ÇáÈÏÁ": SetStartup
    DoCmd.Hourglass False
    MsgBox "Êã ÈäÇÁ ÇáÌÏÇæá æÇáÚáÇŞÇÊ æÇáÇÓÊÚáÇãÇÊ æÇáäãÇĞÌ." & vbCrLf & vbCrLf & _
        "ÇÖÛØ ãæÇİŞ áäÓÎ æÍÏÇÊ ÇáßæÏ. ÈÚÏåÇ: Debug > Compile¡ Ëã ÃÛáŞ Çáãáİ æÇİÊÍå ãä ÌÏíÏ.", vbInformation
    ' æÍÏÇÊ ÇáßæÏ ÂÎÑ ÎØæÉ¡ áÃä ÅÖÇİÉ æÍÏÉ ÃËäÇÁ ÊÔÛíá ÇáßæÏ ŞÏ ÊÚíÏ ÖÈØ ÇáãÔÑæÚ
    mStage = "äÓÎ æÍÏÇÊ ÇáßæÏ"
    If src <> "" Then ImportModules src
    Exit Sub
Fail:
    DoCmd.Hourglass False
    MsgBox "ÊæŞİ ÇáÈäÇÁ İí: " & mStage & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub T_tblUsers()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblUsers")
    t.Fields.Append NewField(t, "UserID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "WinUser", 10, 50, False, True, "", "", "")
    t.Fields.Append NewField(t, "FullName", 10, 100, False, True, "", "", "")
    t.Fields.Append NewField(t, "RankTitle", 10, 50, False, False, "", "", "")
    t.Fields.Append NewField(t, "IsExaminer", 1, 0, False, False, "True", "", "")
    t.Fields.Append NewField(t, "IsReviewer", 1, 0, False, False, "False", "", "")
    t.Fields.Append NewField(t, "IsApprover", 1, 0, False, False, "False", "", "")
    t.Fields.Append NewField(t, "IsAdmin", 1, 0, False, False, "False", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    t.Fields.Append NewField(t, "SignaturePath", 10, 255, False, False, "", "", "")
    AddIndex t, "pkUsers", True, True, False, "UserID"
    AddIndex t, "uxUsersWin", False, True, False, "WinUser"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblUsers").Fields("UserID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblUsers").Fields("WinUser"), "Caption", "ÇÓã ÇáãÓÊÎÏã İí ÇáÌåÇÒ"
    SetProp db.TableDefs("tblUsers").Fields("FullName"), "Caption", "ÇáÇÓã"
    SetProp db.TableDefs("tblUsers").Fields("RankTitle"), "Caption", "ÇáÑÊÈÉ æÇáÕİÉ"
    SetProp db.TableDefs("tblUsers").Fields("IsExaminer"), "Caption", "İÇÍÕ"
    SetProp db.TableDefs("tblUsers").Fields("IsReviewer"), "Caption", "ãÑÇÌöÚ"
    SetProp db.TableDefs("tblUsers").Fields("IsApprover"), "Caption", "ãÚÊãÏ"
    SetProp db.TableDefs("tblUsers").Fields("IsAdmin"), "Caption", "ãÓÄæá ÇáÈÑäÇãÌ"
    SetProp db.TableDefs("tblUsers").Fields("IsActive"), "Caption", "İÚÇá"
    SetProp db.TableDefs("tblUsers").Fields("SignaturePath"), "Caption", "ãÓÇÑ ÇáÊæŞíÚ"
End Sub

Private Sub T_tblAgencies()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblAgencies")
    t.Fields.Append NewField(t, "AgencyID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "AgencyName", 10, 150, False, True, "", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkAgencies", True, True, False, "AgencyID"
    AddIndex t, "uxAgencyName", False, True, False, "AgencyName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblAgencies").Fields("AgencyID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblAgencies").Fields("AgencyName"), "Caption", "ÇáÌåÉ"
    SetProp db.TableDefs("tblAgencies").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblIncidentTypes()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblIncidentTypes")
    t.Fields.Append NewField(t, "IncidentTypeID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "IncidentTypeName", 10, 100, False, True, "", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkIncidentTypes", True, True, False, "IncidentTypeID"
    AddIndex t, "uxIncidentTypeName", False, True, False, "IncidentTypeName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblIncidentTypes").Fields("IncidentTypeID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblIncidentTypes").Fields("IncidentTypeName"), "Caption", "äæÚ ÇáÍÇÏË"
    SetProp db.TableDefs("tblIncidentTypes").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblResults()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblResults")
    t.Fields.Append NewField(t, "ResultID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "ResultName", 10, 30, False, True, "", "", "")
    t.Fields.Append NewField(t, "ReportLabel", 10, 30, False, False, "", "", "")
    AddIndex t, "pkResults", True, True, False, "ResultID"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblResults").Fields("ResultID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblResults").Fields("ResultName"), "Caption", "ÇáäÊíÌÉ"
    SetProp db.TableDefs("tblResults").Fields("ReportLabel"), "Caption", "İí ÇáÊŞÑíÑ"
End Sub

Private Sub T_tblLineTypes()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblLineTypes")
    t.Fields.Append NewField(t, "LineTypeID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "LineTypeName", 10, 30, False, True, "", "", "")
    AddIndex t, "pkLineTypes", True, True, False, "LineTypeID"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblLineTypes").Fields("LineTypeID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblLineTypes").Fields("LineTypeName"), "Caption", "ÇáäæÚ"
End Sub

Private Sub T_tblMethods()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblMethods")
    t.Fields.Append NewField(t, "MethodCode", 10, 30, False, True, "", "", "")
    t.Fields.Append NewField(t, "MethodName", 10, 100, False, False, "", "", "")
    t.Fields.Append NewField(t, "IsDefault", 1, 0, False, False, "False", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkMethods", True, True, False, "MethodCode"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblMethods").Fields("MethodCode"), "Caption", "äæÚ ÇáãÚÇáÌÉ"
    SetProp db.TableDefs("tblMethods").Fields("MethodName"), "Caption", "ÇáæÕİ"
    SetProp db.TableDefs("tblMethods").Fields("IsDefault"), "Caption", "ÇİÊÑÇÖí"
    SetProp db.TableDefs("tblMethods").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblDevices()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblDevices")
    t.Fields.Append NewField(t, "DeviceID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "DeviceName", 10, 50, False, True, "", "", "")
    t.Fields.Append NewField(t, "DurationMin", 3, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkDevices", True, True, False, "DeviceID"
    AddIndex t, "uxDeviceName", False, True, False, "DeviceName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblDevices").Fields("DeviceID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblDevices").Fields("DeviceName"), "Caption", "ÇáÌåÇÒ"
    SetProp db.TableDefs("tblDevices").Fields("DurationMin"), "Caption", "ãÏÉ ÇáÏæÑÉ (ÏŞíŞÉ)"
    SetProp db.TableDefs("tblDevices").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblTransferTargets()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblTransferTargets")
    t.Fields.Append NewField(t, "TargetID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "TargetName", 10, 100, False, True, "", "", "")
    t.Fields.Append NewField(t, "IsDefault", 1, 0, False, False, "False", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkTransferTargets", True, True, False, "TargetID"
    AddIndex t, "uxTargetName", False, True, False, "TargetName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblTransferTargets").Fields("TargetID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblTransferTargets").Fields("TargetName"), "Caption", "ÌåÉ ÇáÊÍæíá"
    SetProp db.TableDefs("tblTransferTargets").Fields("IsDefault"), "Caption", "ÇİÊÑÇÖí"
    SetProp db.TableDefs("tblTransferTargets").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblTraceStatus()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblTraceStatus")
    t.Fields.Append NewField(t, "StatusID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "StatusName", 10, 50, False, True, "", "", "")
    AddIndex t, "pkTraceStatus", True, True, False, "StatusID"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblTraceStatus").Fields("StatusID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblTraceStatus").Fields("StatusName"), "Caption", "ÇáÍÇáÉ"
End Sub

Private Sub T_tblCases()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblCases")
    t.Fields.Append NewField(t, "CaseID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "RefNo", 10, 20, False, True, "", "Not Like ""*[!0-9]*""", "ÑŞã ÇáãÑÌÚ ÃÑŞÇã İŞØ.")
    t.Fields.Append NewField(t, "ForensicNo", 10, 30, False, False, "", "", "")
    t.Fields.Append NewField(t, "AgencyID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "IncidentTypeID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "CaseDate", 8, 0, False, True, "=Date()", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedBy", 10, 50, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedAt", 8, 0, False, False, "=Now()", "", "")
    AddIndex t, "pkCases", True, True, False, "CaseID"
    AddIndex t, "uxCaseRef", False, True, False, "RefNo"
    t.ValidationRule = "[ForensicNo] Is Null Or [ForensicNo] Not Like ""*[!0-9]*"""
    t.ValidationText = "ÑŞã ÇáÃÏáÉ ÇáÌäÇÆíÉ: ÇáÑŞã İŞØ¡ ãËá 312312 (ÇáÓäÉ æ«Ù» ÊõÖÇİ İí ÇáÊŞÑíÑ)."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblCases").Fields("CaseID"), "Caption", "ÇáÑŞã ÇáÏÇÎáí"
    SetProp db.TableDefs("tblCases").Fields("RefNo"), "Caption", "ÑŞã ÇáãÑÌÚ"
    SetProp db.TableDefs("tblCases").Fields("ForensicNo"), "Caption", "ÑŞã ÇáÃÏáÉ ÇáÌäÇÆíÉ"
    SetProp db.TableDefs("tblCases").Fields("AgencyID"), "Caption", "ÇáÌåÉ"
    SetProp db.TableDefs("tblCases").Fields("IncidentTypeID"), "Caption", "äæÚ ÇáÍÇÏË"
    SetProp db.TableDefs("tblCases").Fields("CaseDate"), "Caption", "ÊÇÑíÎ ÇáŞÖíÉ"
    SetProp db.TableDefs("tblCases").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
    SetProp db.TableDefs("tblCases").Fields("CreatedBy"), "Caption", "ÃÖíİ ÈæÇÓØÉ"
    SetProp db.TableDefs("tblCases").Fields("CreatedAt"), "Caption", "ÊÇÑíÎ ÇáÅÖÇİÉ"
End Sub

Private Sub T_tblBatches()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblBatches")
    t.Fields.Append NewField(t, "BatchID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "CaseID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "BatchNo", 3, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "ReceivedDate", 8, 0, False, True, "=Date()", "", "")
    t.Fields.Append NewField(t, "ExamDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "ExportDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "ExaminerID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "ReviewerID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "ApproverID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedBy", 10, 50, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedAt", 8, 0, False, False, "=Now()", "", "")
    AddIndex t, "ixBatchExaminer", False, False, False, "ExaminerID"
    AddIndex t, "ixBatchReceived", False, False, False, "ReceivedDate"
    AddIndex t, "pkBatches", True, True, False, "BatchID"
    AddIndex t, "uxBatchNo", False, True, False, "CaseID,BatchNo"
    t.ValidationRule = "([ExamDate] Is Null Or [ExamDate]>=[ReceivedDate]) And ([ExportDate] Is Null Or [ExamDate] Is Null Or [ExportDate]>=[ExamDate])"
    t.ValidationText = "ÊÇÑíÎ ÇáİÍÕ áÇ íÓÈŞ ÊÇÑíÎ ÇáæÑæÏ¡ æÊÇÑíÎ ÇáÊÕÏíÑ áÇ íÓÈŞ ÊÇÑíÎ ÇáİÍÕ."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblBatches").Fields("BatchID"), "Caption", "ÇáÑŞã ÇáÏÇÎáí"
    SetProp db.TableDefs("tblBatches").Fields("CaseID"), "Caption", "ÇáŞÖíÉ"
    SetProp db.TableDefs("tblBatches").Fields("BatchNo"), "Caption", "ÑŞã ÇáãÌãæÚÉ"
    SetProp db.TableDefs("tblBatches").Fields("ReceivedDate"), "Caption", "ÊÇÑíÎ ÇáæÑæÏ"
    SetProp db.TableDefs("tblBatches").Fields("ExamDate"), "Caption", "ÊÇÑíÎ ÇáİÍÕ"
    SetProp db.TableDefs("tblBatches").Fields("ExportDate"), "Caption", "ÊÇÑíÎ ÇáÊÕÏíÑ"
    SetProp db.TableDefs("tblBatches").Fields("ExaminerID"), "Caption", "ÇáŞÇÆã ÈÇáİÍÕ"
    SetProp db.TableDefs("tblBatches").Fields("ReviewerID"), "Caption", "ÇáŞÇÆã ÈÇáãÑÇÌÚÉ"
    SetProp db.TableDefs("tblBatches").Fields("ApproverID"), "Caption", "ÇáãÚÊãÏ"
    SetProp db.TableDefs("tblBatches").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
    SetProp db.TableDefs("tblBatches").Fields("CreatedBy"), "Caption", "ÃÖíİ ÈæÇÓØÉ"
    SetProp db.TableDefs("tblBatches").Fields("CreatedAt"), "Caption", "ÊÇÑíÎ ÇáÅÖÇİÉ"
End Sub

Private Sub T_tblSamples()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblSamples")
    t.Fields.Append NewField(t, "SampleID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "BatchID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "SampleNo", 3, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "LineType", 4, 0, False, True, "1", "", "")
    t.Fields.Append NewField(t, "Barcode", 10, 20, False, False, "", "", "")
    t.Fields.Append NewField(t, "Description", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "ItemCount", 3, 0, False, False, "1", "", "")
    t.Fields.Append NewField(t, "LiftLocation", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "VisitDateTime", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "MethodCode", 10, 30, False, False, "", "", "")
    t.Fields.Append NewField(t, "PosCount", 3, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "NegCount", 3, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "UnkCount", 3, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "TraceCount", 3, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "PhotoPath", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "TransferTargetID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "StoredDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "RetrievedDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "TransferredDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedBy", 10, 50, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedAt", 8, 0, False, False, "=Now()", "", "")
    AddIndex t, "pkSamples", True, True, False, "SampleID"
    AddIndex t, "uxSampleBarcode", False, True, True, "Barcode"
    AddIndex t, "uxSampleNo", False, True, False, "BatchID,SampleNo"
    t.ValidationRule = "([LineType] In (1,2)) And (([LineType]=1 And [Barcode] Is Not Null And [Description] Is Not Null And [ItemCount] Is Not Null And [LiftLocation] Is Not" & _
        " Null And [MethodCode] Is Not Null And [PosCount] Is Not Null And [NegCount] Is Not Null And [UnkCount] Is Not Null And [ItemCount]>=1 And [PosCount]>" & _
        "=0 And [NegCount]>=0 And [UnkCount]>=0 And [PosCount]+[NegCount]+[UnkCount]=[ItemCount] And (([PosCount]>0 And [TraceCount] Is Not Null And [TraceCoun" & _
        "t]>0) Or ([PosCount]=0 And ([TraceCount] Is Null Or [TraceCount]=0)))) Or ([LineType]=2 And [Barcode] Is Null And [MethodCode] Is Null And [PosCount] " & _
        "Is Null And [NegCount] Is Null And [UnkCount] Is Null And [TraceCount] Is Not Null And [TraceCount]>=0))"
    t.ValidationText = "ÇáÚíäÉ: ÇáÈÇÑßæÏ æÇáæÕİ æÇáÚÏÏ æãßÇä ÇáÑİÚ æäæÚ ÇáãÚÇáÌÉ æÇáÅíÌÇÈí æÇáÓáÈí æáã íÊÈíä ÅÌÈÇÑíÉ¡ æãÌãæÚ ÇáÅíÌÇÈí æÇáÓáÈí æáã íÊÈíä íÓÇæí ÇáÚÏÏ. ÇáÅíÌÇÈí " & _
        "íÍÊÇÌ ÚÏÏ ÇáÂËÇÑ (ÃßÈÑ ãä ÕİÑ). ÇáÇäÊŞÇá: áÇ íõßÊÈ İíå ÅáÇ ÚÏÏ ÇáÂËÇÑ ÇáãÑİæÚÉ (ÕİÑ Ãæ ÃßËÑ)."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblSamples").Fields("SampleID"), "Caption", "ÇáÑŞã ÇáÏÇÎáí"
    SetProp db.TableDefs("tblSamples").Fields("BatchID"), "Caption", "ÇáãÌãæÚÉ"
    SetProp db.TableDefs("tblSamples").Fields("SampleNo"), "Caption", "ã"
    SetProp db.TableDefs("tblSamples").Fields("LineType"), "Caption", "ÇáäæÚ"
    SetProp db.TableDefs("tblSamples").Fields("Barcode"), "Caption", "ÈÇÑßæÏ ÇáÚíäÉ"
    SetProp db.TableDefs("tblSamples").Fields("Description"), "Caption", "æÕİ ÇáÚíäÉ"
    SetProp db.TableDefs("tblSamples").Fields("ItemCount"), "Caption", "ÇáÚÏÏ"
    SetProp db.TableDefs("tblSamples").Fields("LiftLocation"), "Caption", "ãßÇä ÇáÑİÚ"
    SetProp db.TableDefs("tblSamples").Fields("VisitDateTime"), "Caption", "ÊÇÑíÎ ææŞÊ ÇáÇäÊŞÇá"
    SetProp db.TableDefs("tblSamples").Fields("MethodCode"), "Caption", "äæÚ ÇáãÚÇáÌÉ"
    SetProp db.TableDefs("tblSamples").Fields("PosCount"), "Caption", "ÅíÌÇÈíÉ"
    SetProp db.TableDefs("tblSamples").Fields("NegCount"), "Caption", "ÓáÈíÉ"
    SetProp db.TableDefs("tblSamples").Fields("UnkCount"), "Caption", "áã íÊÈíä"
    SetProp db.TableDefs("tblSamples").Fields("TraceCount"), "Caption", "ÚÏÏ ÇáÂËÇÑ"
    SetProp db.TableDefs("tblSamples").Fields("PhotoPath"), "Caption", "ÕæÑÉ ÇáÚíäÉ"
    SetProp db.TableDefs("tblSamples").Fields("TransferTargetID"), "Caption", "ÌåÉ ÇáÊÍæíá"
    SetProp db.TableDefs("tblSamples").Fields("StoredDate"), "Caption", "ÊÇÑíÎ ÇáÍİÙ İí ÇáÎÒÇäÉ"
    SetProp db.TableDefs("tblSamples").Fields("RetrievedDate"), "Caption", "ÊÇÑíÎ ÇáÇÓÊÑÌÇÚ ãä ÇáÎÒÇäÉ"
    SetProp db.TableDefs("tblSamples").Fields("TransferredDate"), "Caption", "ÊÇÑíÎ ÇáÊÍæíá"
    SetProp db.TableDefs("tblSamples").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
    SetProp db.TableDefs("tblSamples").Fields("CreatedBy"), "Caption", "ÃÖíİ ÈæÇÓØÉ"
    SetProp db.TableDefs("tblSamples").Fields("CreatedAt"), "Caption", "ÊÇÑíÎ ÇáÅÖÇİÉ"
End Sub

Private Sub T_tblBatchStaff()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblBatchStaff")
    t.Fields.Append NewField(t, "BatchID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "UserID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "PosCount", 3, 0, False, True, "0", "", "")
    t.Fields.Append NewField(t, "NegCount", 3, 0, False, True, "0", "", "")
    t.Fields.Append NewField(t, "UnkCount", 3, 0, False, True, "0", "", "")
    t.Fields.Append NewField(t, "ProcTraces", 3, 0, False, True, "0", "", "")
    t.Fields.Append NewField(t, "VisitTraces", 3, 0, False, True, "0", "", "")
    AddIndex t, "pkBatchStaff", True, True, False, "BatchID,UserID"
    t.ValidationRule = "[PosCount]>=0 And [NegCount]>=0 And [UnkCount]>=0 And [ProcTraces]>=0 And [VisitTraces]>=0"
    t.ValidationText = "ÇáÃÚÏÇÏ ÕİÑ Ãæ ÃßËÑ."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblBatchStaff").Fields("BatchID"), "Caption", "ÇáãÌãæÚÉ"
    SetProp db.TableDefs("tblBatchStaff").Fields("UserID"), "Caption", "ÇáãæÙİ"
    SetProp db.TableDefs("tblBatchStaff").Fields("PosCount"), "Caption", "ÅíÌÇÈíÉ"
    SetProp db.TableDefs("tblBatchStaff").Fields("NegCount"), "Caption", "ÓáÈíÉ"
    SetProp db.TableDefs("tblBatchStaff").Fields("UnkCount"), "Caption", "áã íÊÈíä"
    SetProp db.TableDefs("tblBatchStaff").Fields("ProcTraces"), "Caption", "ÂËÇÑ ÇáãÚÇáÌÉ"
    SetProp db.TableDefs("tblBatchStaff").Fields("VisitTraces"), "Caption", "ÂËÇÑ ÇáÇäÊŞÇá"
End Sub

Private Sub T_tblTraces()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblTraces")
    t.Fields.Append NewField(t, "TraceID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "SampleID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "TraceNo", 3, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "TraceCode", 10, 40, False, True, "", "", "")
    t.Fields.Append NewField(t, "AtharBarcode", 10, 30, False, False, "", "", "")
    t.Fields.Append NewField(t, "LiftLocation", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "Description", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "PhotoPath", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "StatusID", 4, 0, False, True, "1", "", "")
    t.Fields.Append NewField(t, "AtharRegDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedAt", 8, 0, False, False, "=Now()", "", "")
    AddIndex t, "pkTraces", True, True, False, "TraceID"
    AddIndex t, "uxTraceAthar", False, True, True, "AtharBarcode"
    AddIndex t, "uxTraceCode", False, True, False, "TraceCode"
    AddIndex t, "uxTraceNo", False, True, False, "SampleID,TraceNo"
    t.ValidationRule = "[StatusID]<3 Or [AtharBarcode] Is Not Null"
    t.ValidationText = "ÇáÃËÑ ÇáãÓÌá İí ãäÕÉ ÃËÑ íÍÊÇÌ ÈÇÑßæÏ ÇáÃËÑ."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblTraces").Fields("TraceID"), "Caption", "ÇáÑŞã ÇáÏÇÎáí"
    SetProp db.TableDefs("tblTraces").Fields("SampleID"), "Caption", "ÇáÚíäÉ"
    SetProp db.TableDefs("tblTraces").Fields("TraceNo"), "Caption", "ÑŞã ÇáÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("TraceCode"), "Caption", "ÑãÒ ÇáÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("AtharBarcode"), "Caption", "ÈÇÑßæÏ ãäÕÉ ÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("LiftLocation"), "Caption", "ãßÇä ÑİÚ ÇáÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("Description"), "Caption", "æÕİ ÇáÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("PhotoPath"), "Caption", "ÕæÑÉ ÇáÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("StatusID"), "Caption", "ÇáÍÇáÉ"
    SetProp db.TableDefs("tblTraces").Fields("AtharRegDate"), "Caption", "ÊÇÑíÎ ÇáÊÓÌíá İí ÃËÑ"
    SetProp db.TableDefs("tblTraces").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
    SetProp db.TableDefs("tblTraces").Fields("CreatedAt"), "Caption", "ÊÇÑíÎ ÇáÅÖÇİÉ"
End Sub

Private Sub T_tblRuns()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblRuns")
    t.Fields.Append NewField(t, "RunID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "DeviceID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "RunDate", 8, 0, False, True, "=Date()", "", "")
    t.Fields.Append NewField(t, "StartTime", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "EndTime", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "UserID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    t.Fields.Append NewField(t, "CreatedAt", 8, 0, False, False, "=Now()", "", "")
    AddIndex t, "pkRuns", True, True, False, "RunID"
    t.ValidationRule = "[EndTime] Is Null Or [StartTime] Is Null Or [EndTime]>=[StartTime]"
    t.ValidationText = "æŞÊ ÇáÇäÊåÇÁ áÇ íÓÈŞ æŞÊ ÇáÈÏÁ."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblRuns").Fields("RunID"), "Caption", "ÑŞã ÇáÊÔÛíáÉ"
    SetProp db.TableDefs("tblRuns").Fields("DeviceID"), "Caption", "ÇáÌåÇÒ"
    SetProp db.TableDefs("tblRuns").Fields("RunDate"), "Caption", "ÊÇÑíÎ ÇáãÚÇáÌÉ"
    SetProp db.TableDefs("tblRuns").Fields("StartTime"), "Caption", "æŞÊ ÇáÈÏÁ"
    SetProp db.TableDefs("tblRuns").Fields("EndTime"), "Caption", "æŞÊ ÇáÇäÊåÇÁ"
    SetProp db.TableDefs("tblRuns").Fields("UserID"), "Caption", "ÇáãæÙİ"
    SetProp db.TableDefs("tblRuns").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
    SetProp db.TableDefs("tblRuns").Fields("CreatedAt"), "Caption", "ÊÇÑíÎ ÇáÅÖÇİÉ"
End Sub

Private Sub T_tblRunSamples()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblRunSamples")
    t.Fields.Append NewField(t, "RunID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "SampleID", 4, 0, False, True, "", "", "")
    AddIndex t, "pkRunSamples", True, True, False, "RunID,SampleID"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblRunSamples").Fields("RunID"), "Caption", "ÇáÊÔÛíáÉ"
    SetProp db.TableDefs("tblRunSamples").Fields("SampleID"), "Caption", "ÇáÚíäÉ"
End Sub

Private Sub T_tblStorageLocations()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblStorageLocations")
    t.Fields.Append NewField(t, "LocationID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "LocationName", 10, 100, False, True, "", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkLocations", True, True, False, "LocationID"
    AddIndex t, "uxLocationName", False, True, False, "LocationName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblStorageLocations").Fields("LocationID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblStorageLocations").Fields("LocationName"), "Caption", "ãæŞÚ ÇáÊÎÒíä"
    SetProp db.TableDefs("tblStorageLocations").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblMaterials()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblMaterials")
    t.Fields.Append NewField(t, "MaterialID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "MaterialName", 10, 100, False, True, "", "", "")
    t.Fields.Append NewField(t, "Unit", 10, 30, False, False, "", "", "")
    t.Fields.Append NewField(t, "MinStock", 7, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "IsActive", 1, 0, False, False, "True", "", "")
    AddIndex t, "pkMaterials", True, True, False, "MaterialID"
    AddIndex t, "uxMaterialName", False, True, False, "MaterialName"
    db.TableDefs.Append t
    SetProp db.TableDefs("tblMaterials").Fields("MaterialID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblMaterials").Fields("MaterialName"), "Caption", "ÇáãÇÏÉ"
    SetProp db.TableDefs("tblMaterials").Fields("Unit"), "Caption", "ÇáæÍÏÉ"
    SetProp db.TableDefs("tblMaterials").Fields("MinStock"), "Caption", "ÇáÍÏ ÇáÃÏäì"
    SetProp db.TableDefs("tblMaterials").Fields("IsActive"), "Caption", "İÚÇá"
End Sub

Private Sub T_tblStockMoves()
    Dim t As DAO.TableDef
    Set t = db.CreateTableDef("tblStockMoves")
    t.Fields.Append NewField(t, "MoveID", 4, 0, True, False, "", "", "")
    t.Fields.Append NewField(t, "MaterialID", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "MoveType", 4, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "Qty", 7, 0, False, True, "", "", "")
    t.Fields.Append NewField(t, "MoveDate", 8, 0, False, True, "=Date()", "", "")
    t.Fields.Append NewField(t, "ExpiryDate", 8, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "LotNo", 10, 50, False, False, "", "", "")
    t.Fields.Append NewField(t, "LocationID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "UserID", 4, 0, False, False, "", "", "")
    t.Fields.Append NewField(t, "Notes", 10, 255, False, False, "", "", "")
    AddIndex t, "pkStockMoves", True, True, False, "MoveID"
    t.ValidationRule = "[MoveType] In (1,2) And [Qty]>0"
    t.ValidationText = "äæÚ ÇáÍÑßÉ ÅÖÇİÉ (1) Ãæ ÕÑİ (2)¡ æÇáßãíÉ ÃßÈÑ ãä ÕİÑ."
    db.TableDefs.Append t
    SetProp db.TableDefs("tblStockMoves").Fields("MoveID"), "Caption", "ÇáÑŞã"
    SetProp db.TableDefs("tblStockMoves").Fields("MaterialID"), "Caption", "ÇáãÇÏÉ"
    SetProp db.TableDefs("tblStockMoves").Fields("MoveType"), "Caption", "ÇáÍÑßÉ"
    SetProp db.TableDefs("tblStockMoves").Fields("Qty"), "Caption", "ÇáßãíÉ"
    SetProp db.TableDefs("tblStockMoves").Fields("MoveDate"), "Caption", "ÇáÊÇÑíÎ"
    SetProp db.TableDefs("tblStockMoves").Fields("ExpiryDate"), "Caption", "ÊÇÑíÎ ÇáÇäÊåÇÁ"
    SetProp db.TableDefs("tblStockMoves").Fields("LotNo"), "Caption", "ÑŞã ÇáÏİÚÉ"
    SetProp db.TableDefs("tblStockMoves").Fields("LocationID"), "Caption", "ÇáãæŞÚ"
    SetProp db.TableDefs("tblStockMoves").Fields("UserID"), "Caption", "ÇáãæÙİ"
    SetProp db.TableDefs("tblStockMoves").Fields("Notes"), "Caption", "ãáÇÍÙÇÊ"
End Sub

Private Sub D_tblUsers()
    X "INSERT INTO [tblUsers] ([UserID],[WinUser],[FullName],[RankTitle],[IsExaminer],[IsReviewer],[IsApprover],[IsAdmin],[IsActive],[SignaturePath]) VALUES " & _
        "(1,'ÈæÓáØÇä','íæÓİ ãæÓì ãÍãÏ','äŞíÈ',True,True,True,True,True,'assets\signatures\ÈæÓáØÇä.bmp')"
    X "INSERT INTO [tblUsers] ([UserID],[WinUser],[FullName],[RankTitle],[IsExaminer],[IsReviewer],[IsApprover],[IsAdmin],[IsActive],[SignaturePath]) VALUES " & _
        "(2,'TEST1','ÓíÏ ÇáåÇÔãí','ÇáÚÑíİ Çæá',True,True,True,True,True,Null)"
    X "INSERT INTO [tblUsers] ([UserID],[WinUser],[FullName],[RankTitle],[IsExaminer],[IsReviewer],[IsApprover],[IsAdmin],[IsActive],[SignaturePath]) VALUES " & _
        "(3,'TEST_R1','ÓíÏ ÇáåÇÔãí','ÇáÚÑíİ Çæá',True,True,True,True,True,'assets\signatures\TEST_R1.bmp')"
    X "INSERT INTO [tblUsers] ([UserID],[WinUser],[FullName],[RankTitle],[IsExaminer],[IsReviewer],[IsApprover],[IsAdmin],[IsActive],[SignaturePath]) VALUES " & _
        "(4,'TEST_R2','ãÑÇÌöÚ æåãí 2','ãÑÇÌöÚ',True,True,True,True,True,'assets\signatures\TEST_R2.bmp')"
    X "INSERT INTO [tblUsers] ([UserID],[WinUser],[FullName],[RankTitle],[IsExaminer],[IsReviewer],[IsApprover],[IsAdmin],[IsActive],[SignaturePath]) VALUES " & _
        "(5,'TEST_A1','ãÍãÏ ËÇäí','ÇáÑÇÆÏ ÎÈíÑ',False,False,True,False,True,'assets\signatures\TEST_A1.bmp')"
End Sub

Private Sub D_tblAgencies()
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (1,'ãÑßÒ ÔÑØÉ ÈÑÏÈí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (2,'ãÑßÒ ÔÑØÉ ÇáÈÑÔÇÁ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (3,'ãÑßÒ ÔÑØÉ äÇíİ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (4,'ãÑßÒ ÔÑØÉ äÏÇáÔÈÇ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (5,'ãÑßÒ ÔÑØÉ ÇáãæÇäÆ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (6,'ãÑßÒ ÔÑØÉ ÌÈá Úáí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (7,'ãÑßÒ ÔÑØÉ ÇáŞÕíÕ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (8,'ãÑßÒ ÔÑØÉ ÇáãÑŞÈÇÊ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (9,'ãÑßÒ ÔÑØÉ ÇáÑİÇÚÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (10,'ãÑßÒ ÔÑØÉ ÇáÑÇÔÏíÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (11,'ãÑßÒ ÔÑØÉ ÍÊÇ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (12,'ãÑßÒ ÔÑØÉ áåÈÇÈ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (13,'ãÑßÒ ÔÑØÉ ÇáİŞÚ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (14,'ãÑßÒ ÔÑØÉ ÇáÎæÇäíÌ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (15,'ÇáÅÏÇÑÉ ÇáÚÇãÉ ááÊÍÑíÇÊ æÇáãÈÇÍË ÇáÌäÇÆíÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (16,'ÇáÅÏÇÑÉ ÇáÚÇãÉ ááãÎÏÑÇÊ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (17,'ÇáÅÏÇÑÉ ÇáÚÇãÉ áÃãä ÇáãØÇÑÇÊ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (18,'ÅÏÇÑÉ ÇáØÈ ÇáÔÑÚí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (19,'ÇáÅÏÇÑÉ ÇáÚÇãÉ áÃãä ÇáåíÆÇÊ æÇáãäÔÆÇÊ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (20,'ÔÑØÉ ÚÌãÇä',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (21,'ÔÑØÉ ÇáÔÇÑŞÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (22,'ÔÑØÉ Ãã ÇáŞíæíä',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (23,'ÔÑØÉ ÑÃÓ ÇáÎíãÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (24,'ÔÑØÉ ÇáİÌíÑÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (25,'äíÇÈÉ ÏÈí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (26,'äíÇÈÉ ÃÈæ ÙÈí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (27,'äíÇÈÉ ÚÌãÇä',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (28,'äíÇÈÉ Ãã ÇáŞíæíä',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (29,'äíÇÈÉ ÇáÔÇÑŞÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (30,'äíÇÈÉ ÇáİÌíÑÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (31,'äíÇÈÉ ÑÃÓ ÇáÎíãÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (32,'ÌãÇÑß ÏÈí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (33,'ÇáãßÊÈ ÇáÊäÙíãí ÇáŞíÇÏÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (34,'ãÍßãÉ ÑÃÓ ÇáÎíãÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (35,'ãÍßãÉ ÏÈí',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (36,'ãÍßãÉ ÚÌãÇä',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (37,'ãÍßãÉ ÇáÔÇÑŞÉ',True)"
    X "INSERT INTO [tblAgencies] ([AgencyID],[AgencyName],[IsActive]) VALUES (38,'ÇáÅÏÇÑÉ ÇáÚÇãÉ áÃãä ÇáÏæáÉ',True)"
End Sub

Private Sub D_tblIncidentTypes()
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (1,'ÇáÇÊáÇİ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (2,'ÇáÇÍÊíÇá',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (3,'ÇáÇÓÊíáÇÁ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (4,'ÇáÇÓÍáÉ æÇáãÊİÌÑÇÊ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (5,'ÇáÇÔÊÈÇå',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (6,'ÇáÇÚÊÏÇÁ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (7,'ÇáÇäÊÍÇÑ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (8,'ÇáÊÒæíÑ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (9,'ÇáÊåÏíÏ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (10,'ÇáÍÑíŞ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (11,'ÇáÎØİ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (12,'ÇáÓÑŞÇÊ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (13,'ÇáÓŞæØ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (14,'ÇáØÚä',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (15,'ÇáÚËæÑ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (16,'ÇáŞÊá',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (17,'ÇáåÑæÈ ãä ÑÌÇá ÇáÃãä',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (18,'ÇáæİÇÉ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (19,'ÇäÊÍÇá',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (20,'ÇäÊåÇß ÍÑãÉ ãáß ÇáÛíÑ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (21,'ÃÎÑì',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (22,'ÎíÇäÉ ÇáÃãÇäÉ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (23,'İÍÕ ÌËÉ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (24,'İÍæÕÇÊ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (25,'ŞÖÇíÇ ÇáÎãæÑ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (26,'ŞÖÇíÇ ÇáãÎÏÑÇÊ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (27,'ŞÖÇíÇ ÌäÓíÉ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (28,'ŞÖÇíÇ ãÑæÑíÉ',True)"
    X "INSERT INTO [tblIncidentTypes] ([IncidentTypeID],[IncidentTypeName],[IsActive]) VALUES (29,'ãÒÇæáÉ ãåäÉ ÈÏæä ÊÑÎíÕ',True)"
End Sub

Private Sub D_tblResults()
    X "INSERT INTO [tblResults] ([ResultID],[ResultName],[ReportLabel]) VALUES (1,'ÅíÌÇÈíÉ','ÅíÌÇÈíÉ')"
    X "INSERT INTO [tblResults] ([ResultID],[ResultName],[ReportLabel]) VALUES (2,'ÓáÈíÉ','ÓáÈíÉ')"
    X "INSERT INTO [tblResults] ([ResultID],[ResultName],[ReportLabel]) VALUES (3,'áã íÊÈíä','ÊÚĞÑ')"
End Sub

Private Sub D_tblLineTypes()
    X "INSERT INTO [tblLineTypes] ([LineTypeID],[LineTypeName]) VALUES (1,'ÚíäÉ')"
    X "INSERT INTO [tblLineTypes] ([LineTypeID],[LineTypeName]) VALUES (2,'ÇäÊŞÇá')"
End Sub

Private Sub D_tblMethods()
    X "INSERT INTO [tblMethods] ([MethodCode],[MethodName],[IsDefault],[IsActive]) VALUES ('FSD 75-FPD-SM-01','íÍÏÏ áÇÍŞÇğ',False,True)"
    X "INSERT INTO [tblMethods] ([MethodCode],[MethodName],[IsDefault],[IsActive]) VALUES ('FSD 75-FPD-SM-02','íÍÏÏ áÇÍŞÇğ',True,True)"
    X "INSERT INTO [tblMethods] ([MethodCode],[MethodName],[IsDefault],[IsActive]) VALUES ('FSD 75-FPD-SM-03','íÍÏÏ áÇÍŞÇğ',False,True)"
    X "INSERT INTO [tblMethods] ([MethodCode],[MethodName],[IsDefault],[IsActive]) VALUES ('FSD 75-FPD-SM-04','íÍÏÏ áÇÍŞÇğ',False,True)"
End Sub

Private Sub D_tblDevices()
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (1,'MVC1000',50,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (2,'MVC5000',85,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (3,'MVC5000b',85,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (4,'MVClite',35,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (5,'CARON1',7,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (6,'CARON2',15,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (7,'VMD',10,True)"
    X "INSERT INTO [tblDevices] ([DeviceID],[DeviceName],[DurationMin],[IsActive]) VALUES (8,'EVISCAN',20,True)"
End Sub

Private Sub D_tblTransferTargets()
    X "INSERT INTO [tblTransferTargets] ([TargetID],[TargetName],[IsDefault],[IsActive]) VALUES (1,'ÇáÔÄæä ÇáİäíÉ',True,True)"
    X "INSERT INTO [tblTransferTargets] ([TargetID],[TargetName],[IsDefault],[IsActive]) VALUES (2,'ŞÓã ãŞÇÑäÉ ÇáÈÕãÇÊ',False,True)"
End Sub

Private Sub D_tblTraceStatus()
    X "INSERT INTO [tblTraceStatus] ([StatusID],[StatusName]) VALUES (1,'Êã ÇáÅÙåÇÑ')"
    X "INSERT INTO [tblTraceStatus] ([StatusID],[StatusName]) VALUES (2,'Êã ÇáÊÕæíÑ æÇáÊæËíŞ')"
    X "INSERT INTO [tblTraceStatus] ([StatusID],[StatusName]) VALUES (3,'ãÓÌá İí ãäÕÉ ÃËÑ')"
End Sub

Private Sub D_tblStorageLocations()
    X "INSERT INTO [tblStorageLocations] ([LocationID],[LocationName],[IsActive]) VALUES (1,'ÇáÊÕæíÑ 1',True)"
    X "INSERT INTO [tblStorageLocations] ([LocationID],[LocationName],[IsActive]) VALUES (2,'ãÎÊÈÑ ÓæÈÑ Ìáæ',True)"
    X "INSERT INTO [tblStorageLocations] ([LocationID],[LocationName],[IsActive]) VALUES (3,'ãÎÊÈÑ İÖ ÇáÇÍÑÇÒ',True)"
    X "INSERT INTO [tblStorageLocations] ([LocationID],[LocationName],[IsActive]) VALUES (4,'ãÎÊÈÑ äÇíä åíÏÑíä',True)"
    X "INSERT INTO [tblStorageLocations] ([LocationID],[LocationName],[IsActive]) VALUES (5,'ãÓÊæÏÚ ÇáÑÆíÓí ááÇÙåÇÑ',True)"
End Sub

Private Sub D_tblMaterials()
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (1,'äÇíä åíÏÑíä',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (2,'ÇãíÏæÈáÇß',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (3,'ÈæÏÑå ÚÇÏíÉ',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (4,'ÈæÏÑÉ ããÛäØÉ',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (5,'ÇÓÊíßí ÓÇíÊ ÇÓæÏ',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (6,'ÇÓÊíßí ÓÇíÊ ÇÈíÖ',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (7,'ãíËÇäæá',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (8,'ÇÓíÊæä',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (9,'äÇíä åíÏÑíä i609',Null,Null,True)"
    X "INSERT INTO [tblMaterials] ([MaterialID],[MaterialName],[Unit],[MinStock],[IsActive]) VALUES (10,'äÇíä åíÏÑíä t609',Null,Null,True)"
End Sub

Private Sub BuildRelations()
    AddRelation "fkCaseAgency", "tblAgencies", "tblCases", 0, "AgencyID=AgencyID"
    AddRelation "fkCaseIncident", "tblIncidentTypes", "tblCases", 0, "IncidentTypeID=IncidentTypeID"
    AddRelation "fkBatchCase", "tblCases", "tblBatches", 0, "CaseID=CaseID"
    AddRelation "fkBatchExaminer", "tblUsers", "tblBatches", 0, "UserID=ExaminerID"
    AddRelation "fkBatchReviewer", "tblUsers", "tblBatches", 0, "UserID=ReviewerID"
    AddRelation "fkBatchApprover", "tblUsers", "tblBatches", 0, "UserID=ApproverID"
    AddRelation "fkSampleBatch", "tblBatches", "tblSamples", 0, "BatchID=BatchID"
    AddRelation "fkSampleLineType", "tblLineTypes", "tblSamples", 0, "LineTypeID=LineType"
    AddRelation "fkSampleMethod", "tblMethods", "tblSamples", 256, "MethodCode=MethodCode"
    AddRelation "fkSampleTarget", "tblTransferTargets", "tblSamples", 0, "TargetID=TransferTargetID"
    AddRelation "fkStaffBatch", "tblBatches", "tblBatchStaff", 4096, "BatchID=BatchID"
    AddRelation "fkStaffUser", "tblUsers", "tblBatchStaff", 0, "UserID=UserID"
    AddRelation "fkTraceSample", "tblSamples", "tblTraces", 0, "SampleID=SampleID"
    AddRelation "fkTraceStatus", "tblTraceStatus", "tblTraces", 0, "StatusID=StatusID"
    AddRelation "fkRunDevice", "tblDevices", "tblRuns", 0, "DeviceID=DeviceID"
    AddRelation "fkRunUser", "tblUsers", "tblRuns", 0, "UserID=UserID"
    AddRelation "fkRunSampleRun", "tblRuns", "tblRunSamples", 4096, "RunID=RunID"
    AddRelation "fkRunSampleSample", "tblSamples", "tblRunSamples", 0, "SampleID=SampleID"
    AddRelation "fkMoveMaterial", "tblMaterials", "tblStockMoves", 0, "MaterialID=MaterialID"
    AddRelation "fkMoveLocation", "tblStorageLocations", "tblStockMoves", 0, "LocationID=LocationID"
    AddRelation "fkMoveUser", "tblUsers", "tblStockMoves", 0, "UserID=UserID"
End Sub

Private Sub BuildQueries()
    mStage = "ÇáÇÓÊÚáÇã qTracesEdit"
    db.CreateQueryDef "qTracesEdit", "SELECT t.TraceID, t.SampleID, t.TraceNo, t.TraceCode, t.AtharBarcode, t.LiftLocation, t.Description, t.PhotoPath, t.StatusID, t.AtharRegDate, s.BatchI" & _
        "D, b.CaseID, b.BatchNo, s.SampleNo, s.LineType, s.Description AS SampleDesc FROM (tblTraces AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID" & _
        ") INNER JOIN tblBatches AS b ON s.BatchID = b.BatchID ORDER BY b.BatchNo, s.SampleNo, t.TraceNo;"
    mStage = "ÇáÇÓÊÚáÇã qRunSamples"
    db.CreateQueryDef "qRunSamples", "SELECT rs.RunID, rs.SampleID, s.Barcode, s.Description, s.ItemCount, c.RefNo FROM ((tblRunSamples AS rs INNER JOIN tblSamples AS s ON rs.SampleID = s." & _
        "SampleID) INNER JOIN tblBatches AS b ON s.BatchID = b.BatchID) INNER JOIN tblCases AS c ON b.CaseID = c.CaseID ORDER BY s.Barcode;"
    mStage = "ÇáÇÓÊÚáÇã qRunWork"
    db.CreateQueryDef "qRunWork", "SELECT rs.RunID, rs.SampleID, r.UserID, r.RunDate, s.BatchID, IIf(s.PosCount Is Null,0,s.PosCount) AS PosItems, IIf(s.NegCount Is Null,0,s.NegCount) A" & _
        "S NegItems, IIf(s.PosCount Is Null,0,s.PosCount)+IIf(s.NegCount Is Null,0,s.NegCount) AS TestItems FROM (tblRunSamples AS rs INNER JOIN tblRuns AS r O" & _
        "N rs.RunID = r.RunID) INNER JOIN tblSamples AS s ON rs.SampleID = s.SampleID;"
    mStage = "ÇáÇÓÊÚáÇã qRunWorkSum"
    db.CreateQueryDef "qRunWorkSum", "SELECT BatchID, UserID, Sum(TestItems) AS Tests, Sum(PosItems) AS PosTests, Sum(NegItems) AS NegTests FROM qRunWork GROUP BY BatchID, UserID;"
    mStage = "ÇáÇÓÊÚáÇã qBatchFacts"
    db.CreateQueryDef "qBatchFacts", "SELECT b.BatchID, b.CaseID, b.BatchNo, b.ReceivedDate, b.ReviewerID, Sum(IIf(s.LineType=2,1,0)) AS VisitLines FROM tblBatches AS b LEFT JOIN tblSample" & _
        "s AS s ON b.BatchID = s.BatchID GROUP BY b.BatchID, b.CaseID, b.BatchNo, b.ReceivedDate, b.ReviewerID;"
    mStage = "ÇáÇÓÊÚáÇã qStaffCount"
    db.CreateQueryDef "qStaffCount", "SELECT BatchID, Count(*) AS StaffCount FROM tblBatchStaff GROUP BY BatchID;"
    mStage = "ÇáÇÓÊÚáÇã qStatBase"
    db.CreateQueryDef "qStatBase", "SELECT st.UserID, u.FullName, f.BatchID, f.CaseID, c.RefNo, a.AgencyName, f.BatchNo, f.ReceivedDate, Year(f.ReceivedDate) AS Yr, Month(f.ReceivedDate)" & _
        " AS Mo, DatePart(""q"",f.ReceivedDate) AS Qt, st.PosCount+st.NegCount+st.UnkCount AS Samples, st.PosCount, st.NegCount, st.UnkCount, st.ProcTraces, st.V" & _
        "isitTraces, IIf(st.PosCount+st.NegCount+st.UnkCount>0,1,0) AS ProcCases, IIf(f.VisitLines>0 And (st.VisitTraces>0 Or sc.StaffCount=1),1,0) AS VisitCas" & _
        "es, IIf(sc.StaffCount>1,1,0) AS SharedFlag, IIf(w.Tests Is Null,0,w.Tests) AS Tests, IIf(w.PosTests Is Null,0,w.PosTests) AS PosTests, IIf(w.NegTests " & _
        "Is Null,0,w.NegTests) AS NegTests FROM (((((tblBatchStaff AS st INNER JOIN qBatchFacts AS f ON st.BatchID = f.BatchID) INNER JOIN tblCases AS c ON f.C" & _
        "aseID = c.CaseID) INNER JOIN tblUsers AS u ON st.UserID = u.UserID) INNER JOIN qStaffCount AS sc ON st.BatchID = sc.BatchID) LEFT JOIN tblAgencies AS " & _
        "a ON c.AgencyID = a.AgencyID) LEFT JOIN qRunWorkSum AS w ON st.BatchID = w.BatchID AND st.UserID = w.UserID;"
    mStage = "ÇáÇÓÊÚáÇã ÇáÇÍÕÇÆíÉ ÇáÔåÑíÉ"
    db.CreateQueryDef "ÇáÇÍÕÇÆíÉ ÇáÔåÑíÉ", "SELECT FullName & '' AS [ÇáŞÇÆã ÈÇáİÍÕ], Sum(VisitCases) AS [ÚÏÏ ŞÖÇíÇ ÇáÇäÊŞÇá], Sum(VisitTraces) AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Sum(ProcCases) AS [ÚÏÏ ŞÖÇí" & _
        "Ç ÇáãÚÇáÌÉ], Sum(Samples) AS [ÚÏÏ ÇáÚíäÇÊ], Sum(PosCount) AS [ÇáÅíÌÇÈí], Sum(NegCount) AS [ÇáÓáÈí], Sum(UnkCount) AS [áã íÊÈíä], Sum(ProcTraces) AS [Ú" & _
        "ÏÏ ÇáÂËÇÑ], Count(*) AS [ÇáãäÌÒ ãä ÇáŞÖÇíÇ], Sum(Tests) AS [ÚÏÏ ÇáİÍæÕ], Sum(PosTests) AS [ÇáİÍÕ ÇáÅíÌÇÈí], Sum(NegTests) AS [ÇáİÍÕ ÇáÓáÈí] FROM qStat" & _
        "Base WHERE [Yr]=[TempVars]![StatYear] AND [Mo]=[TempVars]![StatFrom] GROUP BY FullName ORDER BY FullName;"
    mStage = "ÇáÇÓÊÚáÇã ÇÍÕÇÆíÉ ÇáãæÙİ"
    db.CreateQueryDef "ÇÍÕÇÆíÉ ÇáãæÙİ", "SELECT FullName & '' AS [ÇáŞÇÆã ÈÇáİÍÕ], Sum(VisitCases) AS [ÚÏÏ ŞÖÇíÇ ÇáÇäÊŞÇá], Sum(VisitTraces) AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Sum(ProcCases) AS [ÚÏÏ ŞÖÇí" & _
        "Ç ÇáãÚÇáÌÉ], Sum(Samples) AS [ÚÏÏ ÇáÚíäÇÊ], Sum(PosCount) AS [ÇáÅíÌÇÈí], Sum(NegCount) AS [ÇáÓáÈí], Sum(UnkCount) AS [áã íÊÈíä], Sum(ProcTraces) AS [Ú" & _
        "ÏÏ ÇáÂËÇÑ], Count(*) AS [ÇáãäÌÒ ãä ÇáŞÖÇíÇ], Sum(Tests) AS [ÚÏÏ ÇáİÍæÕ], Sum(PosTests) AS [ÇáİÍÕ ÇáÅíÌÇÈí], Sum(NegTests) AS [ÇáİÍÕ ÇáÓáÈí] FROM qStat" & _
        "Base WHERE [Yr]=[TempVars]![StatYear] AND [Mo] Between [TempVars]![StatFrom] And [TempVars]![StatTo] GROUP BY FullName ORDER BY FullName;"
    mStage = "ÇáÇÓÊÚáÇã ÇÍÕÇÆíÉ ÇáÑÈÚ"
    db.CreateQueryDef "ÇÍÕÇÆíÉ ÇáÑÈÚ", "SELECT FullName & '' AS [ÇáŞÇÆã ÈÇáİÍÕ], Qt AS [ÇáÑÈÚ], Sum(VisitCases) AS [ÚÏÏ ŞÖÇíÇ ÇáÇäÊŞÇá], Sum(VisitTraces) AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Sum(ProcCase" & _
        "s) AS [ÚÏÏ ŞÖÇíÇ ÇáãÚÇáÌÉ], Sum(Samples) AS [ÚÏÏ ÇáÚíäÇÊ], Sum(PosCount) AS [ÇáÅíÌÇÈí], Sum(NegCount) AS [ÇáÓáÈí], Sum(UnkCount) AS [áã íÊÈíä], Sum(Pr" & _
        "ocTraces) AS [ÚÏÏ ÇáÂËÇÑ], Count(*) AS [ÇáãäÌÒ ãä ÇáŞÖÇíÇ], Sum(Tests) AS [ÚÏÏ ÇáİÍæÕ], Sum(PosTests) AS [ÇáİÍÕ ÇáÅíÌÇÈí], Sum(NegTests) AS [ÇáİÍÕ ÇáÓ" & _
        "áÈí] FROM qStatBase WHERE [Yr]=[TempVars]![StatYear] GROUP BY FullName, Qt ORDER BY FullName, Qt;"
    mStage = "ÇáÇÓÊÚáÇã qReviewsYear"
    db.CreateQueryDef "qReviewsYear", "SELECT ReviewerID, Count(*) AS Reviews FROM tblBatches WHERE ReviewerID Is Not Null AND Year(ReceivedDate)=[TempVars]![StatYear] GROUP BY ReviewerID;"
    mStage = "ÇáÇÓÊÚáÇã qStatYearSums"
    db.CreateQueryDef "qStatYearSums", "SELECT UserID, FullName & '' AS [ÇáŞÇÆã ÈÇáİÍÕ], Sum(VisitCases) AS [ÚÏÏ ŞÖÇíÇ ÇáÇäÊŞÇá], Sum(VisitTraces) AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Sum(ProcCases) AS [" & _
        "ÚÏÏ ŞÖÇíÇ ÇáãÚÇáÌÉ], Sum(Samples) AS [ÚÏÏ ÇáÚíäÇÊ], Sum(PosCount) AS [ÇáÅíÌÇÈí], Sum(NegCount) AS [ÇáÓáÈí], Sum(UnkCount) AS [áã íÊÈíä], Sum(ProcTrace" & _
        "s) AS [ÚÏÏ ÇáÂËÇÑ], Count(*) AS [ÇáãäÌÒ ãä ÇáŞÖÇíÇ], Sum(Tests) AS [ÚÏÏ ÇáİÍæÕ], Sum(PosTests) AS [ÇáİÍÕ ÇáÅíÌÇÈí], Sum(NegTests) AS [ÇáİÍÕ ÇáÓáÈí] FR" & _
        "OM qStatBase WHERE [Yr]=[TempVars]![StatYear] GROUP BY UserID, FullName;"
    mStage = "ÇáÇÓÊÚáÇã ÇáÇÍÕÇÆíÉ ÇáÓäæíÉ"
    db.CreateQueryDef "ÇáÇÍÕÇÆíÉ ÇáÓäæíÉ", "SELECT y.*, IIf(r.Reviews Is Null,0,r.Reviews) AS [ãÑÇÌÚÉ] FROM qStatYearSums AS y LEFT JOIN qReviewsYear AS r ON y.UserID = r.ReviewerID ORDER BY y.[" & _
        "ÇáŞÇÆã ÈÇáİÍÕ];"
    mStage = "ÇáÇÓÊÚáÇã ÇÍÕÇÆíÉ ÇáİÍæÕÇÊ"
    db.CreateQueryDef "ÇÍÕÇÆíÉ ÇáİÍæÕÇÊ", "SELECT FullName & '' AS [ÇáŞÇÆã ÈÇáİÍÕ], RefNo & '' AS [ÑŞã ÇáãÑÌÚ], BatchNo+0 AS [ÇáãÌãæÚÉ], DateValue(ReceivedDate) AS [ÊÇÑíÎ ÇáæÑæÏ], AgencyName & " & _
        "'' AS [ÇáÌåÉ], Samples+0 AS [ÚÏÏ ÇáÚíäÇÊ], PosCount+0 AS [ÇáÅíÌÇÈí], NegCount+0 AS [ÇáÓáÈí], UnkCount+0 AS [áã íÊÈíä], ProcTraces+0 AS [ÚÏÏ ÇáÂËÇÑ], V" & _
        "isitTraces+0 AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Tests+0 AS [ÚÏÏ ÇáİÍæÕ], PosTests+0 AS [ÇáİÍÕ ÇáÅíÌÇÈí], NegTests+0 AS [ÇáİÍÕ ÇáÓáÈí], IIf(SharedFlag=1,""äÚã"","""")" & _
        " AS [ãÔÊÑßÉ] FROM qStatBase WHERE [Yr]=[TempVars]![StatYear] AND [Mo] Between [TempVars]![StatFrom] And [TempVars]![StatTo] ORDER BY ReceivedDate, Ref" & _
        "No, FullName;"
    mStage = "ÇáÇÓÊÚáÇã ÇáÇÍÕÇÆíÉ ÇáÚÇãÉ"
    db.CreateQueryDef "ÇáÇÍÕÇÆíÉ ÇáÚÇãÉ", "SELECT Sum(VisitCases) AS [ÚÏÏ ŞÖÇíÇ ÇáÇäÊŞÇá], Sum(VisitTraces) AS [ÚÏÏ ÂËÇÑ ÇáÇäÊŞÇá], Sum(ProcCases) AS [ÚÏÏ ŞÖÇíÇ ÇáãÚÇáÌÉ], Sum(Samples) AS [ÚÏÏ " & _
        "ÇáÚíäÇÊ], Sum(PosCount) AS [ÇáÅíÌÇÈí], Sum(NegCount) AS [ÇáÓáÈí], Sum(UnkCount) AS [áã íÊÈíä], Sum(ProcTraces) AS [ÚÏÏ ÇáÂËÇÑ], Count(*) AS [ÇáãäÌÒ ãä" & _
        " ÇáŞÖÇíÇ], Sum(Tests) AS [ÚÏÏ ÇáİÍæÕ], Sum(PosTests) AS [ÇáİÍÕ ÇáÅíÌÇÈí], Sum(NegTests) AS [ÇáİÍÕ ÇáÓáÈí] FROM qStatBase WHERE [Yr]=[TempVars]![StatYe" & _
        "ar] AND [Mo] Between [TempVars]![StatFrom] And [TempVars]![StatTo];"
    mStage = "ÇáÇÓÊÚáÇã qReviewCases"
    db.CreateQueryDef "qReviewCases", "SELECT ReviewerID, CaseID, Count(*) AS Batches FROM tblBatches WHERE ReviewerID Is Not Null AND Year(ReceivedDate)=[TempVars]![StatYear] AND Month(Rec" & _
        "eivedDate) Between [TempVars]![StatFrom] And [TempVars]![StatTo] GROUP BY ReviewerID, CaseID;"
    mStage = "ÇáÇÓÊÚáÇã ÇÍÕÇÆíÉ ÇáãÑÇÌÚíä"
    db.CreateQueryDef "ÇÍÕÇÆíÉ ÇáãÑÇÌÚíä", "SELECT u.FullName & '' AS [ÇáŞÇÆã ÈÇáãÑÇÌÚÉ], Count(*) AS [ÚÏÏ ÇáŞÖÇíÇ], Sum(q.Batches) AS [ÚÏÏ ÇáãÌãæÚÇÊ] FROM qReviewCases AS q INNER JOIN tblUsers " & _
        "AS u ON q.ReviewerID = u.UserID GROUP BY u.FullName ORDER BY u.FullName;"
    mStage = "ÇáÇÓÊÚáÇã ÇÍÕÇÆíÉ ÇáŞÖÇíÇ ÍÓÈ ÇáÌåÇÊ"
    db.CreateQueryDef "ÇÍÕÇÆíÉ ÇáŞÖÇíÇ ÍÓÈ ÇáÌåÇÊ", "TRANSFORM Count(c.CaseID) AS n SELECT IIf(a.AgencyName Is Null,'(ÈáÇ ÌåÉ)',a.AgencyName) AS [ÇáÌåÉ], Count(c.CaseID) AS [ÇáÅÌãÇáí] FROM tblCases AS c " & _
        "LEFT JOIN tblAgencies AS a ON c.AgencyID = a.AgencyID WHERE Year(c.CaseDate)=2026 GROUP BY IIf(a.AgencyName Is Null,'(ÈáÇ ÌåÉ)',a.AgencyName) PIVOT Fo" & _
        "rmat(c.CaseDate,'mm') IN ('01','02','03','04','05','06','07','08','09','10','11','12');"
    mStage = "ÇáÇÓÊÚáÇã qStockBalance"
    db.CreateQueryDef "qStockBalance", "SELECT m.MaterialID, m.MaterialName & '' AS [ÇáãÇÏÉ], m.Unit & '' AS [ÇáæÍÏÉ], m.MinStock, IIf(Sum(IIf(s.MoveType=1,s.Qty,IIf(s.MoveType=2,-s.Qty,0)))" & _
        " Is Null,0,Sum(IIf(s.MoveType=1,s.Qty,IIf(s.MoveType=2,-s.Qty,0)))) AS Balance, Min(IIf(s.MoveType=1 And s.ExpiryDate>=Date(),s.ExpiryDate,Null)) AS N" & _
        "earestExpiry, Max(s.MoveDate) AS LastMove FROM tblMaterials AS m LEFT JOIN tblStockMoves AS s ON m.MaterialID = s.MaterialID WHERE m.IsActive=True GRO" & _
        "UP BY m.MaterialID, m.MaterialName, m.Unit, m.MinStock ORDER BY m.MaterialName;"
    mStage = "ÇáÇÓÊÚáÇã qStockMoves"
    db.CreateQueryDef "qStockMoves", "SELECT s.MoveID, s.MaterialID, s.MoveDate, IIf(s.MoveType=1,'ÅÖÇİÉ','ÕÑİ') AS MoveName, s.MoveType, s.Qty, s.LotNo, s.ExpiryDate, l.LocationName, u.Fu" & _
        "llName, s.Notes FROM (tblStockMoves AS s LEFT JOIN tblStorageLocations AS l ON s.LocationID = l.LocationID) LEFT JOIN tblUsers AS u ON s.UserID = u.Us" & _
        "erID ORDER BY s.MoveDate DESC, s.MoveID DESC;"
    mStage = "ÇáÇÓÊÚáÇã qCaseFind"
    db.CreateQueryDef "qCaseFind", "SELECT CaseID, RefNo, ForensicNo, Format(CaseDate,'yyyy/mm/dd') AS Dated FROM tblCases ORDER BY CaseID DESC;"
End Sub

Private Sub ImportForms(ByVal src As String)
    Dim v As Variant
    For Each v In Array("frmBatchesSub", "frmMain", "frmSamples", "frmTraces", "frmRunSamplesSub", "frmRuns", "frmBatchStaff", "frmStockBalanceSub", "frmStockMovesSub", "frmStock", "frmDashboard", "frmMaterials", "frmUsers")
        mStage = "äÓÎ ÇáäãæĞÌ " & v
        DoCmd.TransferDatabase acImport, "Microsoft Access", src, acForm, CStr(v), CStr(v)
    Next
End Sub

Private Sub ImportModules(ByVal src As String)
    Dim v As Variant
    For Each v In Array("modApp", "modReport", "modDashboard")
        DoCmd.TransferDatabase acImport, "Microsoft Access", src, acModule, CStr(v), CStr(v)
    Next
End Sub

Private Sub SetStartup()
    SetProp CurrentDb, "AppTitle", "ŞÓã ÇáÅÙåÇÑ ÇáßíãíÇÆí"
    If FormExists("frmMain") Then SetProp CurrentDb, "StartupForm", "frmMain"
    Application.RefreshTitleBar
End Sub

' ---------- ÃÏæÇÊ ãÓÇÚÏÉ ----------

Private Function NewField(t As DAO.TableDef, ByVal fName As String, ByVal fType As Integer, ByVal fSize As Long, _
        ByVal autoNum As Boolean, ByVal req As Boolean, ByVal defVal As String, _
        ByVal valRule As String, ByVal valText As String) As DAO.Field
    Dim f As DAO.Field
    If fSize > 0 Then
        Set f = t.CreateField(fName, fType, fSize)
    Else
        Set f = t.CreateField(fName, fType)
    End If
    If autoNum Then f.Attributes = dbFixedField + dbAutoIncrField
    If req Then f.Required = True
    If Len(defVal) > 0 Then f.DefaultValue = defVal
    If Len(valRule) > 0 Then f.ValidationRule = valRule
    If Len(valText) > 0 Then f.ValidationText = valText
    Set NewField = f
End Function

Private Sub AddIndex(t As DAO.TableDef, ByVal iName As String, ByVal isPrimary As Boolean, _
        ByVal isUnique As Boolean, ByVal ignoreNulls As Boolean, ByVal fieldList As String)
    Dim i As DAO.Index, v As Variant
    Set i = t.CreateIndex(iName)
    i.Primary = isPrimary
    i.Unique = isUnique
    i.IgnoreNulls = ignoreNulls
    For Each v In Split(fieldList, ",")
        i.Fields.Append i.CreateField(CStr(v))
    Next
    t.Indexes.Append i
End Sub

Private Sub AddRelation(ByVal rName As String, ByVal pkTable As String, ByVal fkTable As String, _
        ByVal attrs As Long, ByVal pairs As String)
    Dim r As DAO.Relation, f As DAO.Field, v As Variant
    Set r = db.CreateRelation(rName, pkTable, fkTable, attrs)
    For Each v In Split(pairs, ";")
        Set f = r.CreateField(Split(v, "=")(0))
        f.ForeignName = Split(v, "=")(1)
        r.Fields.Append f
    Next
    db.Relations.Append r
End Sub

Private Sub X(ByVal sql As String)
    db.Execute sql, dbFailOnError
End Sub

Private Sub SetProp(obj As Object, ByVal pName As String, ByVal pValue As String)
    On Error Resume Next
    obj.Properties(pName) = pValue
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        obj.Properties.Append obj.CreateProperty(pName, dbText, pValue)
    End If
End Sub

Private Function TableExists(ByVal tName As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = tName Then TableExists = True: Exit Function
    Next
End Function

Private Function FormExists(ByVal fName As String) As Boolean
    Dim o As AccessObject
    For Each o In CurrentProject.AllForms
        If o.Name = fName Then FormExists = True: Exit Function
    Next
End Function

Private Function PickSource() As String
    With Application.FileDialog(3)      ' msoFileDialogFilePicker
        .Title = "ÇÎÊÑ ãáİ ÇáÈÑäÇãÌ ÇáŞÏíã (áäÓÎ ÇáäãÇĞÌ ææÍÏÇÊ ÇáßæÏ)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Access", "*.accdb; *.mdb"
        If .Show Then PickSource = .SelectedItems(1)
    End With
    If PickSource = CurrentDb.Name Then PickSource = ""
End Function
