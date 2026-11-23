<#
  يبني ملف أكسس للبرنامج من الملفات النصية في مجلد src:
    src\app.json        الجداول والحقول والفهارس والعلاقات والبيانات الأساسية والاستعلامات
    src\modules\*.bas   وحدات الكود (UTF-8)
    src\forms\*.txt     النماذج (تُنشأ أول مرة من ملفك الحالي عبر -SourceDb)

  الاستخدام:
    powershell -ExecutionPolicy Bypass -File Build-App.ps1
    powershell -ExecutionPolicy Bypass -File Build-App.ps1 -SourceDb "D:\القديم.accdb"
    powershell -ExecutionPolicy Bypass -File Build-App.ps1 -Output "D:\برنامج.accdb" -Force

  يحتاج Microsoft Access مثبتاً على الجهاز.
#>
param(
    [string]$Output = (Join-Path $PSScriptRoot 'build\الإظهار الكيميائي.accdb'),
    [string]$SourceDb,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'src'
$spec = Get-Content -Raw -Encoding UTF8 (Join-Path $src 'app.json') | ConvertFrom-Json

function Has($obj, $name) { return $null -ne $obj.PSObject.Properties[$name] }
function Step($msg) { Write-Host "-> $msg" -ForegroundColor Cyan }

# ---------- التحقق من المسارات ----------
$Output = [IO.Path]::GetFullPath($Output)
if (Test-Path $Output) {
    if (-not $Force) { throw "الملف موجود: $Output  (استخدم -Force للاستبدال)" }
    Remove-Item $Output
}
New-Item -ItemType Directory -Force -Path (Split-Path $Output) | Out-Null
if ($SourceDb) {
    $SourceDb = [IO.Path]::GetFullPath($SourceDb)
    if (-not (Test-Path $SourceDb)) { throw "الملف المصدر غير موجود: $SourceDb" }
}

$acc = New-Object -ComObject Access.Application
try {
    $acc.Visible = $false
    Step "إنشاء قاعدة البيانات: $Output"
    $acc.NewCurrentDatabase($Output)
    $db = $acc.CurrentDb()

    # ---------- الجداول ----------
    foreach ($t in $spec.tables) {
        Step "جدول $($t.name)"
        $tdf = $db.CreateTableDef($t.name)
        foreach ($f in $t.fields) {
            $size = if (Has $f 'size') { [int]$f.size } else { 0 }
            $fld = if ($size) { $tdf.CreateField($f.name, [int]$f.type, $size) } else { $tdf.CreateField($f.name, [int]$f.type) }
            if ((Has $f 'autoNumber') -and $f.autoNumber) { $fld.Attributes = 17 }   # dbFixedField + dbAutoIncrField
            if ((Has $f 'required') -and $f.required) { $fld.Required = $true }
            if (Has $f 'defaultValue')   { $fld.DefaultValue   = $f.defaultValue }
            if (Has $f 'validationRule') { $fld.ValidationRule = $f.validationRule }
            if (Has $f 'validationText') { $fld.ValidationText = $f.validationText }
            $tdf.Fields.Append($fld)
        }
        foreach ($i in $t.indexes) {
            $idx = $tdf.CreateIndex($i.name)
            $idx.Primary = [bool]$i.primary
            $idx.Unique = [bool]$i.unique
            $idx.IgnoreNulls = [bool]$i.ignoreNulls
            foreach ($c in $i.fields) {
                $desc = $c.EndsWith(' DESC')
                $ifld = $idx.CreateField($c.Replace(' DESC', ''))
                if ($desc) { $ifld.Attributes = 1 }
                $idx.Fields.Append($ifld)
            }
            $tdf.Indexes.Append($idx)
        }
        if (Has $t 'validationRule') {
            $tdf.ValidationRule = $t.validationRule
            $tdf.ValidationText = $t.validationText
        }
        $db.TableDefs.Append($tdf)

        # الخصائص التي لا تُضاف إلا بعد حفظ الجدول (التسمية التوضيحية والتنسيق والوصف)
        $tdf = $db.TableDefs.Item($t.name)
        foreach ($f in $t.fields) {
            $fld = $tdf.Fields.Item($f.name)
            foreach ($p in @(@('caption', 'Caption'), @('format', 'Format'), @('description', 'Description'))) {
                if (Has $f $p[0]) { $fld.Properties.Append($fld.CreateProperty($p[1], 10, $f.($p[0]))) }
            }
        }
    }

    # ---------- البيانات الأساسية ----------
    foreach ($t in $spec.tables) {
        if ($t.rows.Count -gt 0) {
            Step "بيانات $($t.name) ($($t.rows.Count) سجل)"
            foreach ($sql in $t.rows) { $db.Execute($sql, 128) }   # dbFailOnError
        }
    }

    # ---------- العلاقات ----------
    Step "العلاقات ($($spec.relations.Count))"
    foreach ($r in $spec.relations) {
        $rel = $db.CreateRelation($r.name, $r.table, $r.foreignTable, [int]$r.attributes)
        foreach ($pair in $r.fields) {
            $rf = $rel.CreateField($pair[0])
            $rf.ForeignName = $pair[1]
            $rel.Fields.Append($rf)
        }
        $db.Relations.Append($rel)
    }

    # ---------- الاستعلامات ----------
    Step "الاستعلامات ($($spec.queries.Count))"
    foreach ($q in $spec.queries) { $db.CreateQueryDef($q.name, $q.sql) | Out-Null }

    # ---------- وحدات الكود ----------
    $ansi = [Text.Encoding]::GetEncoding(1256)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('moosa_build_' + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    foreach ($m in $spec.modules) {
        Step "وحدة $m"
        $text = [IO.File]::ReadAllText((Join-Path $src "modules\$m.bas"), [Text.Encoding]::UTF8)
        $text = ($text -replace "`r?`n", "`r`n")
        $file = Join-Path $tmp "$m.bas"
        [IO.File]::WriteAllText($file, $text, $ansi)
        try {
            $acc.LoadFromText(5, $m, $file)          # acModule
        } catch {
            $acc.VBE.ActiveVBProject.VBComponents.Import($file) | Out-Null
            $acc.DoCmd.Save(5, $m)
        }
    }
    Remove-Item -Recurse -Force $tmp

    # ---------- النماذج ----------
    $formsDir = Join-Path $src 'forms'
    $haveText = (Test-Path $formsDir) -and (Get-ChildItem $formsDir -Filter *.txt -ErrorAction SilentlyContinue)
    if ($haveText) {
        foreach ($frm in $spec.forms) {
            $file = Join-Path $formsDir "$frm.txt"
            if (Test-Path $file) { Step "نموذج $frm"; $acc.LoadFromText(2, $frm, $file) }   # acForm
            else { Write-Warning "النموذج غير موجود في src\forms: $frm" }
        }
    } elseif ($SourceDb) {
        New-Item -ItemType Directory -Force -Path $formsDir | Out-Null
        foreach ($frm in $spec.forms) {
            Step "استيراد النموذج $frm من الملف المصدر"
            $acc.DoCmd.TransferDatabase(0, 'Microsoft Access', $SourceDb, 2, $frm, $frm)   # acImport, acForm
            $acc.SaveAsText(2, $frm, (Join-Path $formsDir "$frm.txt"))
        }
        Write-Host "   حُفظت النماذج نصياً في src\forms، والبناء القادم لا يحتاج -SourceDb." -ForegroundColor Green
    } else {
        Write-Warning "لا توجد نماذج: شغّل السكربت مع -SourceDb ومسار ملفك الحالي لاستيراد النماذج ($($spec.forms.Count) نموذج)."
    }

    # ---------- خصائص قاعدة البيانات ----------
    Step "خصائص بدء التشغيل"
    $props = @{ AppTitle = $spec.dbprops.AppTitle }
    $allForms = $acc.CurrentProject.AllForms
    for ($n = 0; $n -lt $allForms.Count; $n++) {
        if ($allForms.Item($n).Name -eq $spec.dbprops.StartupForm) { $props.StartupForm = $spec.dbprops.StartupForm }
    }
    foreach ($k in $props.Keys) {
        try { $db.Properties.Item($k).Value = $props[$k] }
        catch { $db.Properties.Append($db.CreateProperty($k, 10, $props[$k])) }
    }

    $db = $null
    $acc.CloseCurrentDatabase()
    Write-Host "`nتم. الملف: $Output" -ForegroundColor Green
    Write-Host "افتحه، ثم من محرر VBA (Alt+F11) اختر Debug > Compile للتأكد من الكود." -ForegroundColor Green
}
finally {
    try { $acc.Quit(2) } catch {}   # acQuitSaveNone
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($acc)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
