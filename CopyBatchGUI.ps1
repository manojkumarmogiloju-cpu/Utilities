# WinForms GUI for batch Copy/Move while preserving structure (PowerShell 5.1 compatible)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Helpers ----------
function Ensure-Dir([string]$dir) {
    if (-not [System.IO.Directory]::Exists($dir)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
    }
}

function Get-RelativePath {
    param([string]$Root, [string]$FullPath)
    try {
        $root = (Resolve-Path $Root).Path
        if (-not $root.EndsWith('\')) { $root += '\' }
        $uriRoot = New-Object System.Uri($root)
        $uriFull = New-Object System.Uri((Resolve-Path $FullPath).Path)
        return $uriRoot.MakeRelativeUri($uriFull).ToString().Replace('/','\')
    } catch {
        $r = [System.IO.Path]::GetFullPath((Resolve-Path $Root).Path).TrimEnd('\')
        $p = [System.IO.Path]::GetFullPath((Resolve-Path $FullPath).Path)
        if ($p.StartsWith($r,[System.StringComparison]::InvariantCultureIgnoreCase)) {
            return $p.Substring($r.Length).TrimStart('\')
        }
        return Split-Path $p -Leaf
    }
}

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Copy/Move (preserve structure)"
$form.StartPosition = "CenterScreen"
$form.ClientSize = New-Object System.Drawing.Size(860, 560)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source root:"
$lblSource.Location = New-Object System.Drawing.Point(15, 18)
$lblSource.AutoSize = $true
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(120, 15)
$txtSource.Size = New-Object System.Drawing.Size(640, 24)
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Browse..."
$btnSource.Location = New-Object System.Drawing.Point(770, 14)
$btnSource.Size = New-Object System.Drawing.Size(75, 26)
$btnSource.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtSource.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnSource)

$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination root:"
$lblDest.Location = New-Object System.Drawing.Point(15, 54)
$lblDest.AutoSize = $true
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(120, 51)
$txtDest.Size = New-Object System.Drawing.Size(640, 24)
$form.Controls.Add($txtDest)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Browse..."
$btnDest.Location = New-Object System.Drawing.Point(770, 50)
$btnDest.Size = New-Object System.Drawing.Size(75, 26)
$btnDest.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtDest.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnDest)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Action"
$grp.Location = New-Object System.Drawing.Point(18, 88)
$grp.Size = New-Object System.Drawing.Size(250, 60)
$form.Controls.Add($grp)

$optCopy = New-Object System.Windows.Forms.RadioButton
$optCopy.Text = "Copy"
$optCopy.Location = New-Object System.Drawing.Point(15, 25)
$optCopy.Checked = $true
$grp.Controls.Add($optCopy)

$optMove = New-Object System.Windows.Forms.RadioButton
$optMove.Text = "Move (free source space)"
$optMove.Location = New-Object System.Drawing.Point(75, 25)
$grp.Controls.Add($optMove)

$lblBatch = New-Object System.Windows.Forms.Label
$lblBatch.Text = "Files per batch:"
$lblBatch.Location = New-Object System.Drawing.Point(285, 108)
$lblBatch.AutoSize = $true
$form.Controls.Add($lblBatch)

$nudBatch = New-Object System.Windows.Forms.NumericUpDown
$nudBatch.Minimum = 1
$nudBatch.Maximum = 1000000
$nudBatch.Value = 5000
$nudBatch.Location = New-Object System.Drawing.Point(380, 105)
$nudBatch.Size = New-Object System.Drawing.Size(100, 24)
$form.Controls.Add($nudBatch)

$chkDry = New-Object System.Windows.Forms.CheckBox
$chkDry.Text = "Dry run (no changes)"
$chkDry.Location = New-Object System.Drawing.Point(500, 106)
$chkDry.AutoSize = $true
$form.Controls.Add($chkDry)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Optional file filter (e.g. *.pdf;*.docx):"
$lblFilter.Location = New-Object System.Drawing.Point(15, 160)
$lblFilter.AutoSize = $true
$form.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(260, 157)
$txtFilter.Size = New-Object System.Drawing.Size(505, 24)
$txtFilter.Text = "*"
$form.Controls.Add($txtFilter)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "CSV log (optional):"
$lblLog.Location = New-Object System.Drawing.Point(15, 195)
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(120, 192)
$txtLog.Size = New-Object System.Drawing.Size(640, 24)
$form.Controls.Add($txtLog)

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = "Choose..."
$btnLog.Location = New-Object System.Drawing.Point(770, 191)
$btnLog.Size = New-Object System.Drawing.Size(75, 26)
$btnLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtLog.Text = $dlg.FileName }
})
$form.Controls.Add($btnLog)

$btnEstimate = New-Object System.Windows.Forms.Button
$btnEstimate.Text = "Estimate count"
$btnEstimate.Location = New-Object System.Drawing.Point(18, 232)
$btnEstimate.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnEstimate)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Process NEXT batch"
$btnNext.Location = New-Object System.Drawing.Point(150, 232)
$btnNext.Size = New-Object System.Drawing.Size(160, 30)
$form.Controls.Add($btnNext)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(320, 232)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(18, 275)
$txtOut.Multiline = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.ReadOnly = $true
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.Size = New-Object System.Drawing.Size(827, 265)
$form.Controls.Add($txtOut)

function Append-Out([string]$line){ $txtOut.AppendText($line + [Environment]::NewLine) }

# ---------- Background worker (thread-safe updates) ----------
$bw = New-Object System.ComponentModel.BackgroundWorker
$bw.WorkerReportsProgress = $true
$bw.WorkerSupportsCancellation = $true

$bw.add_DoWork({
    param($sender, $e)
    $state     = $e.Argument
    $srcRoot   = $state.Source
    $dstRoot   = $state.Dest
    $batchSize = [int]$state.Batch
    $doMove    = [bool]$state.Move
    $dry       = [bool]$state.DryRun
    $patterns  = $state.Patterns
    $logPath   = $state.LogPath

    $rows = New-Object System.Collections.Generic.List[object]
    $processed = 0
    $errors = 0

    foreach ($pat in $patterns) {
        foreach ($f in [System.IO.Directory]::EnumerateFiles($srcRoot, $pat, [System.IO.SearchOption]::AllDirectories)) {
            if ($bw.CancellationPending) { $e.Cancel = $true; break }
            if ($processed -ge $batchSize) { break }

            $rel = Get-RelativePath -Root $srcRoot -FullPath $f
            $target = Join-Path $dstRoot $rel
            $targetDir = Split-Path $target -Parent
            try { Ensure-Dir $targetDir } catch {}

            $act = if ($doMove) { "Move" } else { "Copy" }
            $msg = "OK"

            try {
                if (-not $dry) {
                    if ($doMove) {
                        [System.IO.File]::Move($f, $target, $false)
                    } else {
                        [System.IO.File]::Copy($f, $target, $true)
                    }
                }
                $rows.Add([pscustomobject]@{
                    Timestamp = (Get-Date)
                    Action    = $act
                    Status    = "Done"
                    Source    = $f
                    Target    = $target
                    Message   = $msg
                }) | Out-Null
                $processed++
                $sender.ReportProgress(0, ("[{0}] {1} -> {2}" -f $act, $f, $target))
            } catch {
                $errors++
                $rows.Add([pscustomobject]@{
                    Timestamp = (Get-Date)
                    Action    = $act
                    Status    = "Error"
                    Source    = $f
                    Target    = $target
                    Message   = $_.Exception.Message
                }) | Out-Null
                $sender.ReportProgress(0, ("ERROR: {0} -> {1} | {2}" -f $f, $target, $_.Exception.Message))
            }
        }
        if ($processed -ge $batchSize) { break }
        if ($e.Cancel) { break }
    }

    $e.Result = @{ Processed = $processed; Errors = $errors; Rows = $rows; LogPath = $logPath }
})

$bw.add_ProgressChanged({
    param($sender, $e)
    Append-Out ($e.UserState)
})

$bw.add_RunWorkerCompleted({
    param($sender, $e)
    $btnNext.Enabled = $true
    $btnCancel.Enabled = $false

    if ($e.Cancelled) {
        Append-Out "Batch cancelled."
        return
    }
    if ($e.Error -ne $null) {
        Append-Out ("Worker failed: {0}" -f $e.Error.Message)
        return
    }
    $res = $e.Result
    if ($res.LogPath) {
        try { $res.Rows | Export-Csv -Path $res.LogPath -NoTypeInformation -Force }
        catch { Append-Out ("Log write failed: {0}" -f $_.Exception.Message) }
    }
    Append-Out ("Batch complete. Files: {0}, Errors: {1}" -f $res.Processed, $res.Errors)
})

function Validate-Roots {
    $src = $txtSource.Text.Trim()
    $dst = $txtDest.Text.Trim()
    if (-not (Test-Path -LiteralPath $src)) { [void][System.Windows.Forms.MessageBox]::Show("Source root not found."); return $null }
    if (-not (Test-Path -LiteralPath $dst)) { try { Ensure-Dir $dst } catch { [void][System.Windows.Forms.MessageBox]::Show("Cannot create destination root."); return $null } }
    return @{ Source=$src; Dest=$dst }
}

# Buttons
$btnEstimate.Add_Click({
    $roots = Validate-Roots
    if (-not $roots) { return }
    $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { "*" } else { $_ } })
    $count = 0
    foreach ($pat in $patterns) {
        foreach ($f in [System.IO.Directory]::EnumerateFiles($roots.Source, $pat, [System.IO.SearchOption]::AllDirectories)) {
            $count++
            if ($count % 100000 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        }
    }
    Append-Out ("Estimated files matching filter: {0}" -f $count)
})

$btnNext.Add_Click({
    $roots = Validate-Roots
    if (-not $roots) { return }
    $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { "*" } else { $_ } })
    $state = @{
        Source   = $roots.Source
        Dest     = $roots.Dest
        Batch    = [int]$nudBatch.Value
        Move     = $optMove.Checked
        DryRun   = $chkDry.Checked
        Patterns = $patterns
        LogPath  = $txtLog.Text.Trim()
    }
    $btnNext.Enabled = $false
    $btnCancel.Enabled = $true
    $bw.RunWorkerAsync($state) | Out-Null
})

$btnCancel.Add_Click({
    if ($bw.IsBusy) { $bw.CancelAsync() }
})

# Show
[void]$form.ShowDialog()
