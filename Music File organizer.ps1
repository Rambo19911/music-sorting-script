###############################################################################
# Jauns, vienkāršots skripts: Mūzikas failu organizators
# Versija: 1.0
# Autors: ChatGPT
###############################################################################

# Pārliecināmies, ka nepieciešamās .NET bibliotēkas ir pieejamas
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

# Pārbaudām, vai sistēmā ir pieejamas Windows Forms (ja skripts laikus apstājas, tad Windows Forms nav pieejams)
try {
    $testForm = New-Object System.Windows.Forms.Form -ErrorAction Stop
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testForm) | Out-Null
} catch {
    Write-Host "Šis skripts var darboties tikai uz Windows ar pieejamu GUI (Windows Forms)."
    return
}

# Funkcija pārvieto nosaukto failu vai mapi uz atkritni
function Move-ToRecycleBin {
    param([String]$Path)
    if (Test-Path $Path) {
        $item = Get-Item $Path
        if ($item.PSIsContainer) {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($Path, 'OnlyErrorDialogs', 'SendToRecycleBin')
        } else {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Path, 'OnlyErrorDialogs', 'SendToRecycleBin', 'ThrowIfMissing')
        }
    }
}

# Globālie mainīgie, lai varētu apturēt ilgus procesus
$global:StopRequested = $false
$global:CurrentJob = $null

###############################################################################
# Palīgfunkcijas metadatu nolasīšanai 
# (cenšas atrast Artist un Title ar Shell.API + WMP, ja nav, tad izmanto faila nosaukumu)
###############################################################################
function Get-Artist {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) { return "Unknown" }

    $artist = $null
    try {
        # Izmēģinām Shell API
        $shell   = New-Object -ComObject Shell.Application
        $folder  = Split-Path $FilePath -Parent
        $file    = Split-Path $FilePath -Leaf
        $fldObj  = $shell.Namespace($folder)
        $filObj  = $fldObj.ParseName($file)

        # Drošības pēc pārbaudām vairākus metadatu ID
        $artistIndices = 13, 20, 275
        foreach ($idx in $artistIndices) {
            $val = $fldObj.GetDetailsOf($filObj, $idx)
            if ($val) {
                $artist = $val
                break
            }
        }
        # Objektu atbrīvošana
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($filObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fldObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
    catch { }

    # Ja Shell API neko nedeva, mēģinām WMP
    if (-not $artist) {
        try {
            $wmp   = New-Object -ComObject WMPlayer.OCX
            $media = $wmp.newMedia($FilePath)
            if ($media) {
                $tmp = $media.getItemInfo("Artist")
                if ($tmp) { $artist = $tmp }
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($media) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wmp) | Out-Null
        }
        catch { }
    }

    if (-not $artist) {
        # pēdējais variants - mēģinām nolasīt daļu no faila nosaukuma
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        if ($baseName -match "^(.*?)\s*-") { $artist = $matches[1].Trim() }
        if (-not $artist) { $artist = "Unknown" }
    }

    # Attīram bīstamos rakstzīmju un atgriežam
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) {
        $artist = $artist -replace [regex]::Escape($c), "_"
    }
    $artist = $artist.Trim()
    if (-not $artist) { $artist = "Unknown" }
    return $artist
}

function Get-Title {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) { return "Unknown" }

    $title = $null
    try {
        # Shell API
        $shell   = New-Object -ComObject Shell.Application
        $folder  = Split-Path $FilePath -Parent
        $file    = Split-Path $FilePath -Leaf
        $fldObj  = $shell.Namespace($folder)
        $filObj  = $fldObj.ParseName($file)

        # Dažādi indeksi (21 = Title, 2 = Name, ...). Palaižam meklējumus.
        $titleIndices = 21, 2
        foreach ($idx in $titleIndices) {
            $val = $fldObj.GetDetailsOf($filObj, $idx)
            if ($val) {
                $title = $val
                break
            }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($filObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fldObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
    catch { }

    if (-not $title) {
        try {
            $wmp = New-Object -ComObject WMPlayer.OCX
            $media = $wmp.newMedia($FilePath)
            if ($media) {
                $tmp = $media.getItemInfo("Title")
                if ($tmp) { $title = $tmp }
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($media) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wmp)   | Out-Null
        } catch { }
    }

    if (-not $title) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $title = $baseName
    }

    # Aizstājam nederīgos simbolus
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) {
        $title = $title -replace [regex]::Escape($c), "_"
    }
    $title = $title.Trim()
    if (-not $title) { $title = "Unknown" }
    return $title
}

###############################################################################
# Formas izveidošana
###############################################################################
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Mūzikas failu organizators"
$Form.Size = New-Object System.Drawing.Size(600, 400)
$Form.StartPosition = "CenterScreen"
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$Form.AutoSize = $true
$Form.AutoSizeMode = "GrowAndShrink"

# Pogas un kontroles
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Izvēlēties mapi"
$btnSelectFolder.Location = New-Object System.Drawing.Point(20, 30)
$btnSelectFolder.Size = New-Object System.Drawing.Size(130, 35)
$Form.Controls.Add($btnSelectFolder)

$lblSelectedFolder = New-Object System.Windows.Forms.Label
$lblSelectedFolder.Text = "Nav izvēlēta mape"
$lblSelectedFolder.Location = New-Object System.Drawing.Point(160, 30)
$lblSelectedFolder.Size = New-Object System.Drawing.Size(400, 35)
$lblSelectedFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Form.Controls.Add($lblSelectedFolder)

# Radio pogas
$GroupBox = New-Object System.Windows.Forms.GroupBox
$GroupBox.Text = "Izvēlieties darbību"
$GroupBox.Location = New-Object System.Drawing.Point(20, 80)
$GroupBox.Size = New-Object System.Drawing.Size(540, 140)
$Form.Controls.Add($GroupBox)

$rbOrganize = New-Object System.Windows.Forms.RadioButton
$rbOrganize.Text = "Sakārtot pēc mākslinieka"
$rbOrganize.Location = New-Object System.Drawing.Point(10, 20)
$rbOrganize.Checked = $true
$GroupBox.Controls.Add($rbOrganize)

$rbRename = New-Object System.Windows.Forms.RadioButton
$rbRename.Text = "Pārdēvēt (Mākslinieks - Dziesma)"
$rbRename.Location = New-Object System.Drawing.Point(10, 45)
$GroupBox.Controls.Add($rbRename)

$rbDuplicates = New-Object System.Windows.Forms.RadioButton
$rbDuplicates.Text = "Atrast dublikātus un dzēst"
$rbDuplicates.Location = New-Object System.Drawing.Point(10, 70)
$GroupBox.Controls.Add($rbDuplicates)

$rbEmptyFolders = New-Object System.Windows.Forms.RadioButton
$rbEmptyFolders.Text = "Atrast un dzēst tukšās mapes"
$rbEmptyFolders.Location = New-Object System.Drawing.Point(10, 95)
$GroupBox.Controls.Add($rbEmptyFolders)

# START/STOP poga
$btnStartStop = New-Object System.Windows.Forms.Button
$btnStartStop.Text = "SĀKT"
$btnStartStop.Location = New-Object System.Drawing.Point(20, 230)
$btnStartStop.Size = New-Object System.Drawing.Size(130, 40)
$btnStartStop.Enabled = $false
$btnStartStop.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80) # zaļgana
$btnStartStop.ForeColor = [System.Drawing.Color]::White
$Form.Controls.Add($btnStartStop)

# Loga (teksta) lauks
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 280)
$txtLog.Size = New-Object System.Drawing.Size(540, 70)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::WhiteSmoke
$Form.Controls.Add($txtLog)

# Palīgfunkcija loga paziņojumiem
function Write-Log {
    param([string]$message, [string]$color = "Black")
    $txtLog.SelectionColor = [System.Drawing.Color]::FromName($color)
    $txtLog.AppendText("$message`r`n")
    $txtLog.ScrollToCaret()
}

# Pārtraukuma (Stop) apstrāde
function Stop-CurrentJob {
    $global:StopRequested = $true
    if ($global:CurrentJob -and ($global:CurrentJob.State -eq "Running")) {
        Stop-Job $global:CurrentJob -Force
    }
    Write-Log "Darbība apturēta." "Red"
}

###############################################################################
# Funkcijas darbībām
###############################################################################
function Organize-MusicByArtist {
    param($FolderPath)

    Write-Log "Sāk organizēt failus pēc mākslinieka..."
    $files = Get-ChildItem -Path $FolderPath -Recurse -File | 
             Where-Object { $_.Extension -match '^(?i)\.(mp3|flac|wav|aac|m4a|ogg|wma)$'}

    $count = $files.Count
    Write-Log "Atrasti $count mūzikas faili."

    $scriptBlock = {
        param($filesArg, $basePath)
        $i = 0
        $total = $filesArg.Count

        foreach ($f in $filesArg) {
            if ($global:StopRequested) { break }
            $i++
            Write-Progress -Activity "Organizē failus" -Status $f.Name -PercentComplete (($i/$total)*100)
            
            $artist = Get-Artist $f.FullName
            $destination = Join-Path $basePath $artist
            if (-not (Test-Path $destination)) { New-Item $destination -ItemType Directory | Out-Null }

            $targetFile = Join-Path $destination $f.Name
            try {
                Move-Item $f.FullName -Destination $targetFile -Force
                Write-Output "Pārvietots: $($f.Name) -> $artist|Green"
            } catch {
                Write-Output "Kļūda pārvietojot $($f.Name): $($_.Exception.Message)|Red"
            }
        }
    }

    $global:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $files, $FolderPath

    while ($global:CurrentJob.State -eq "Running") {
        Start-Sleep -Milliseconds 300
        $out = Receive-Job $global:CurrentJob -Keep
        foreach ($line in $out) {
            $msg,$col = $line -split '\|',2
            Write-Log $msg $col
        }
        [System.Windows.Forms.Application]::DoEvents()
        if ($global:StopRequested) {
            Stop-Job $global:CurrentJob -Force
            return
        }
    }

    $final = Receive-Job $global:CurrentJob -ErrorAction SilentlyContinue
    foreach ($line in $final) {
        $msg,$col = $line -split '\|',2
        Write-Log $msg $col
    }
    Remove-Job $global:CurrentJob
    $global:CurrentJob = $null
    Write-Log "Organizēšana pabeigta!" "Blue"
}

function Rename-FilesByMetadata {
    param($FolderPath)

    Write-Log "Sāk pārdēvēšanu..."
    $files = Get-ChildItem -Path $FolderPath -Recurse -File | 
             Where-Object { $_.Extension -match '^(?i)\.(mp3|flac|wav|aac|m4a|ogg|wma)$'}

    $count = $files.Count
    Write-Log "Atrasti $count mūzikas faili."

    $scriptBlock = {
        param($filesArg)
        $i = 0
        $total = $filesArg.Count

        foreach ($f in $filesArg) {
            if ($global:StopRequested) { break }
            $i++
            Write-Progress -Activity "Pārdēvē failus" -Status $f.Name -PercentComplete (($i/$total)*100)

            $artist = Get-Artist $f.FullName
            $title  = Get-Title  $f.FullName
            $newName = "$artist - $title$($f.Extension)"
            $newPath = Join-Path $f.DirectoryName $newName
            
            if ($f.FullName -ne $newPath) {
                try {
                    Rename-Item $f.FullName -NewName $newName -Force
                    Write-Output "Pārdēvēts: $($f.Name) -> $newName|Green"
                } catch {
                    Write-Output "Kļūda pārdēvējot $($f.Name): $($_.Exception.Message)|Red"
                }
            }
        }
    }

    $global:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $files

    while ($global:CurrentJob.State -eq "Running") {
        Start-Sleep -Milliseconds 300
        $out = Receive-Job $global:CurrentJob -Keep
        foreach ($line in $out) {
            $msg,$col = $line -split '\|',2
            Write-Log $msg $col
        }
        [System.Windows.Forms.Application]::DoEvents()
        if ($global:StopRequested) {
            Stop-Job $global:CurrentJob -Force
            return
        }
    }

    $final = Receive-Job $global:CurrentJob -ErrorAction SilentlyContinue
    foreach ($line in $final) {
        $msg,$col = $line -split '\|',2
        Write-Log $msg $col
    }
    Remove-Job $global:CurrentJob
    $global:CurrentJob = $null
    Write-Log "Pārdēvēšana pabeigta!" "Blue"
}

function Find-DuplicatesAndDelete {
    param($FolderPath)

    Write-Log "Sāk dublikātu meklēšanu..."
    $files = Get-ChildItem -Path $FolderPath -Recurse -File | 
             Where-Object { $_.Extension -match '^(?i)\.(mp3|flac|wav|aac|m4a|ogg|wma)$'}

    $count = $files.Count
    Write-Log "Atrasti $count mūzikas faili."

    if ($count -eq 0) {
        Write-Log "Nav failu, ko pārbaudīt." "Red"
        return
    }

    $scriptBlock = {
        param($filesArg)
        $hashTable = @{}
        $i = 0
        $total = $filesArg.Count
        
        foreach ($fileObj in $filesArg) {
            if ($global:StopRequested) { break }
            $i++
            Write-Progress -Activity "Rēķina kontrolsummas" -Status $fileObj.Name -PercentComplete (($i/$total)*100)
            if (Test-Path $fileObj.FullName) {
                $fileHash = (Get-FileHash -Algorithm MD5 -Path $fileObj.FullName).Hash
                if ($hashTable.ContainsKey($fileHash)) {
                    $hashTable[$fileHash] += $fileObj
                } else {
                    $hashTable[$fileHash] = @($fileObj)
                }
            }
        }
        return $hashTable
    }

    $global:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $files

    while ($global:CurrentJob.State -eq "Running") {
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.Application]::DoEvents()
        if ($global:StopRequested) {
            Stop-Job $global:CurrentJob -Force
            return
        }
    }

    $hashTableResult = Receive-Job $global:CurrentJob -ErrorAction SilentlyContinue
    Remove-Job $global:CurrentJob
    $global:CurrentJob = $null

    if (-not $hashTableResult) {
        Write-Log "Kontrolsumma nav aprēķināta, iespējams, lietotājs apturēja darbu."
        return
    }

    $duplicateGroups = $hashTableResult.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    if (-not $duplicateGroups) {
        Write-Log "Dublikāti nav atrasti." "Blue"
        return
    }

    # Parāda jaunu formu ar dublikātiem, lai lietotājs var izvēlēties, kuri jādzēš
    $dupForm = New-Object System.Windows.Forms.Form
    $dupForm.Text = "Atrastie dublikāti"
    $dupForm.Size = New-Object System.Drawing.Size(800, 500)
    $dupForm.StartPosition = "CenterScreen"

    $lblDups = New-Object System.Windows.Forms.Label
    $lblDups.Text = "Lūdzu atķeksējiet failus, kurus pārvietot uz atkritni."
    $lblDups.Location = New-Object System.Drawing.Point(20, 20)
    $lblDups.Size = New-Object System.Drawing.Size(760, 30)
    $dupForm.Controls.Add($lblDups)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(20, 60)
    $panel.Size = New-Object System.Drawing.Size(760, 350)
    $panel.AutoScroll = $true
    $dupForm.Controls.Add($panel)

    $deleteBtn = New-Object System.Windows.Forms.Button
    $deleteBtn.Text = "Dzēst atķeksētos"
    $deleteBtn.Location = New-Object System.Drawing.Point(600, 420)
    $deleteBtn.Size = New-Object System.Drawing.Size(180, 40)
    $deleteBtn.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $deleteBtn.ForeColor = [System.Drawing.Color]::White
    $dupForm.Controls.Add($deleteBtn)

    $checkboxes = @{}
    $yPos = 10
    $groupIndex = 1
    foreach ($grp in $duplicateGroups) {
        $filesInGroup = $grp.Value
        # Katram grupam - groupbox
        $gBox = New-Object System.Windows.Forms.GroupBox
        $gBox.Text = "Dublikātu grupa #$groupIndex"
        $gBox.Location = New-Object System.Drawing.Point(10, $yPos)
        $gBox.Size = New-Object System.Drawing.Size(700, 20 + ($filesInGroup.Count * 25))
        $panel.Controls.Add($gBox)

        $subY = 20
        $fIndex = 1
        foreach ($fItem in $filesInGroup) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = $fItem.FullName
            $cb.Location = New-Object System.Drawing.Point(10, $subY)
            $cb.Size = New-Object System.Drawing.Size(680, 20)

            # Pēc noklusējuma *atķeksējam* visus, izņemot pirmo (pieņemot, ka pirmo atstājam)
            if ($fIndex -gt 1) { $cb.Checked = $true }
            $fIndex++

            $gBox.Controls.Add($cb)
            $checkboxes[$cb] = $fItem
            $subY += 25
        }

        $yPos += $gBox.Height + 10
        $groupIndex++
    }

    $deleteBtn.Add_Click({
        $selFiles = $checkboxes.Keys | Where-Object { $_.Checked } | ForEach-Object { $checkboxes[$_] }
        if ($selFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nav atķeksēts neviens fails!", "Informācija",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Vai tiešām dzēst "+$selFiles.Count+" failus?", 
            "Apstiprinājums",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -eq "Yes") {
            foreach ($sf in $selFiles) {
                if (Test-Path $sf.FullName) {
                    try {
                        Move-ToRecycleBin -Path $sf.FullName
                        Write-Log "Dublikāts izdzēsts: $($sf.FullName)" "Green"
                    } catch {
                        Write-Log "Kļūda dzēšot $($sf.FullName): $($_.Exception.Message)" "Red"
                    }
                }
            }
            [System.Windows.Forms.MessageBox]::Show("$($selFiles.Count) fail(-s) pārvietoti uz atkritni.", 
                "Pabeigts",
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $dupForm.Close()
        }
    })

    $dupForm.ShowDialog()
    Write-Log "Dublikātu apstrāde pabeigta." "Blue"
}

function Remove-EmptyFolders {
    param($FolderPath)

    Write-Log "Meklē tukšās mapes..."
    $folders = Get-ChildItem $FolderPath -Recurse -Directory | Sort-Object FullName -Descending
    $count = $folders.Count
    Write-Log "Atrastās mapes: $count"

    $scriptBlock = {
        param($foldersArg)
        $deletedCount = 0
        $processed = 0
        $total = $foldersArg.Count

        foreach ($fld in $foldersArg) {
            if ($global:StopRequested) { break }
            $processed++
            Write-Progress -Activity "Pārbauda mapes" -Status $fld.FullName -PercentComplete (($processed/$total)*100)
            $filesInside = Get-ChildItem -Path $fld.FullName -Recurse -File -ErrorAction SilentlyContinue
            if ($filesInside.Count -eq 0) {
                try {
                    Move-ToRecycleBin -Path $fld.FullName
                    $deletedCount++
                    Write-Output "Izdzēsta tukša mape: $($fld.FullName)|Blue"
                } catch {
                    Write-Output "Kļūda dzēšot mapi $($fld.FullName): $($_.Exception.Message)|Red"
                }
            }
        }
        Write-Output "Tukšo mapju dzēšana pabeigta. Kopā izdzēsts: $deletedCount|Green"
    }

    $global:CurrentJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $folders

    while ($global:CurrentJob.State -eq "Running") {
        Start-Sleep -Milliseconds 300
        $out = Receive-Job $global:CurrentJob -Keep
        foreach ($line in $out) {
            $msg,$col = $line -split '\|',2
            Write-Log $msg $col
        }
        [System.Windows.Forms.Application]::DoEvents()
        if ($global:StopRequested) {
            Stop-Job $global:CurrentJob -Force
            return
        }
    }

    $final = Receive-Job $global:CurrentJob -ErrorAction SilentlyContinue
    foreach ($line in $final) {
        $msg,$col = $line -split '\|',2
        Write-Log $msg $col
    }
    Remove-Job $global:CurrentJob
    $global:CurrentJob = $null
    Write-Log "Tukšās mapes (ja bija) izdzēstas." "Blue"
}

###############################################################################
# Notikumi
###############################################################################
$btnSelectFolder.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Izvēlieties mapi ar mūzikas failiem"
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lblSelectedFolder.Text = $fbd.SelectedPath
        $btnStartStop.Enabled = $true
    }
})

$btnStartStop.Add_Click({
    if ($btnStartStop.Text -eq "SĀKT") {
        if (-not (Test-Path $lblSelectedFolder.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Nav izvēlēta derīga mape!", "Kļūda",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return
        }

        $txtLog.Clear()
        $global:StopRequested = $false
        $btnStartStop.Text = "APTURĒT"
        $btnStartStop.BackColor = [System.Drawing.Color]::FromArgb(244, 67, 54) # sarkanīga

        $folderPath = $lblSelectedFolder.Text

        if ($rbOrganize.Checked) {
            Organize-MusicByArtist $folderPath
        } elseif ($rbRename.Checked) {
            Rename-FilesByMetadata $folderPath
        } elseif ($rbDuplicates.Checked) {
            Find-DuplicatesAndDelete $folderPath
        } elseif ($rbEmptyFolders.Checked) {
            Remove-EmptyFolders $folderPath
        }

        # Pēc pabeigšanas
        $btnStartStop.Text = "SĀKT"
        $btnStartStop.BackColor = [System.Drawing.Color]::FromArgb(76,175,80)
    }
    else {
        # APTURĒT
        Stop-CurrentJob
        $btnStartStop.Text = "SĀKT"
        $btnStartStop.BackColor = [System.Drawing.Color]::FromArgb(76,175,80)
    }
})

[System.Windows.Forms.Application]::Run($Form)

# Pēc formas aizvēršanas notīrām, ja ir kāds job
Get-Job | Remove-Job -Force
