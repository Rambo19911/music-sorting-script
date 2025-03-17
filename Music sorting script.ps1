Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Valodu tulkojumu haštabula
$translations = @{
    "lv" = @{
        Title = "Mūzikas Failu Organizators"
        Description = "Organizē mūzikas failus, pārdēvē tos, izveido atskaņošanas sarakstus un sakārto kolekciju."
        SelectFolder = "Izvēlēties mapi"
        Start = "SĀKT"
        NoFolderSelected = "Nav izvēlēta mape"
        ActionGroup = "Izvēlieties darbību"
        OrganizeByArtist = "Sakārtot failus pēc pirmā mākslinieka"
        FindEmptyFolders = "Atrast un dzēst tukšās mapes"
        FindDuplicates = "Atrast dublikātus"
        FindVideoFiles = "Atrast un dzēst video failus"
        FindImageFiles = "Atrast un dzēst attēlu failus"
        RenameFiles = "Pārdēvēt failus pēc metadatiem (Mākslinieks - Dziesma)"
        GeneratePlaylist = "Izveidot atskaņošanas sarakstu (.m3u)"
        TestMetadata = "Testēt metadatus"
        Ready = "Gatavs darbam"
        InfoTooltips = @{
            OrganizeByArtist = "Šī funkcija sakārto mūzikas failus mapēs, balstoties uz pirmo mākslinieku metadatos."
            FindEmptyFolders = "Atrast un dzēš tukšas mapes izvēlētajā direktorijā."
            FindDuplicates = "Atrast un piedāvā dzēst dublikātus, pamatojoties uz failu kontrolsummām."
            FindVideoFiles = "Atrast un piedāvā dzēst video failus (piem., .mp4, .avi) no mūzikas mapes."
            FindImageFiles = "Atrast un piedāvā dzēst attēlu failus (piem., .jpg, .png) no mūzikas mapes."
            RenameFiles = "Pārdēvēt mūzikas failus formātā 'Mākslinieks - Dziesma', izmantojot metadatus."
            GeneratePlaylist = "Izveidot .m3u atskaņošanas sarakstu ar visiem mūzikas failiem no mapes."
        }
    }
    "en" = @{
        Title = "Music File Organizer"
        Description = "Organize music files, rename them, create playlists, and sort your collection."
        SelectFolder = "Select Folder"
        Start = "START"
        NoFolderSelected = "No folder selected"
        ActionGroup = "Select Action"
        OrganizeByArtist = "Organize files by first artist"
        FindEmptyFolders = "Find and delete empty folders"
        FindDuplicates = "Find duplicates"
        FindVideoFiles = "Find and delete video files"
        FindImageFiles = "Find and delete image files"
        RenameFiles = "Rename files by metadata (Artist - Song)"
        GeneratePlaylist = "Generate playlist (.m3u)"
        TestMetadata = "Test metadata"
        Ready = "Ready to work"
        InfoTooltips = @{
            OrganizeByArtist = "This function organizes music files into folders based on the first artist in metadata."
            FindEmptyFolders = "Finds and deletes empty folders in the selected directory."
            FindDuplicates = "Finds and offers to delete duplicates based on file checksums."
            FindVideoFiles = "Finds and offers to delete video files (e.g., .mp4, .avi) from the music folder."
            FindImageFiles = "Finds and offers to delete image files (e.g., .jpg, .png) from the music folder."
            RenameFiles = "Renames music files to the format 'Artist - Song' using metadata."
            GeneratePlaylist = "Creates a .m3u playlist with all music files from the folder."
        }
    }
    "lt" = @{
        Title = "Muzikos Failų Organizatorius"
        Description = "Tvarkykite muzikos failus, pervadinkite juos, sukurkite grojaraščius ir rūšiuokite kolekciją."
        SelectFolder = "Pasirinkti aplanką"
        Start = "PRADĖTI"
        NoFolderSelected = "Nėra pasirinkto aplanko"
        ActionGroup = "Pasirinkite veiksmą"
        OrganizeByArtist = "Rūšiuoti failus pagal pirmąjį atlikėją"
        FindEmptyFolders = "Rasti ir ištrinti tuščius aplankus"
        FindDuplicates = "Rasti dublikatus"
        FindVideoFiles = "Rasti ir ištrinti vaizdo failus"
        FindImageFiles = "Rasti ir ištrinti vaizdo failus"
        RenameFiles = "Pervadinti failus pagal metaduomenis (Atlikėjas - Daina)"
        GeneratePlaylist = "Sukurti grojaraštį (.m3u)"
        TestMetadata = "Testuoti metaduomenis"
        Ready = "Pasiruošęs darbui"
        InfoTooltips = @{
            OrganizeByArtist = "Ši funkcija rūšiuoja muzikos failus į aplankus pagal pirmąjį atlikėją metaduomenyse."
            FindEmptyFolders = "Randa ir ištrina tuščius aplankus pasirinktame kataloge."
            FindDuplicates = "Randa ir siūlo ištrinti dublikatus, remiantis failų kontrolės sumomis."
            FindVideoFiles = "Randa ir siūlo ištrinti vaizdo failus (pvz., .mp4, .avi) iš muzikos aplanko."
            FindImageFiles = "Randa ir siūlo ištrinti vaizdo failus (pvz., .jpg, .png) iš muzikos aplanko."
            RenameFiles = "Pervadina muzikos failus į formatą 'Atlikėjas - Daina', naudodama metaduomenis."
            GeneratePlaylist = "Sukuria .m3u grojaraštį su visais muzikos failiem iš aplanko."
        }
    }
    "ee" = @{
        Title = "Muusikafailide Korraldaja"
        Description = "Korralda muusikafailid, nimeta ne ümber, loo esitusloendeid ja sorteeri oma kogumik."
        SelectFolder = "Vali kaust"
        Start = "ALUSTA"
        NoFolderSelected = "Kausta pole valitud"
        ActionGroup = "Vali tegevus"
        OrganizeByArtist = "Korralda failid esimese artisti järgi"
        FindEmptyFolders = "Leia ja kustuta tühjad kaustad"
        FindDuplicates = "Leia duplikaadid"
        FindVideoFiles = "Leia ja kustuta videofailid"
        FindImageFiles = "Leia ja kustuta pildifailid"
        RenameFiles = "Nimeta failid ümber metateabe järgi (Artist - Laul)"
        GeneratePlaylist = "Loo esitusloend (.m3u)"
        TestMetadata = "Testi metateavet"
        Ready = "Valmis tööks"
        InfoTooltips = @{
            OrganizeByArtist = "See funktsioon korraldab muusikafailid kaustadesse esimese artisti järgi metateabe alusel."
            FindEmptyFolders = "Leiab ja kustutab tühjad kaustad valitud kataloogis."
            FindDuplicates = "Leiab ja pakub kustutada duplikaate failide kontrollsummade alusel."
            FindVideoFiles = "Leiab ja pakub kustutada videofailid (nt .mp4, .avi) muusikakaustast."
            FindImageFiles = "Leiab ja pakub kustutada pildifailid (nt .jpg, .png) muusikakaustast."
            RenameFiles = "Nimeta muusikafailid ümber formaati 'Artist - Laul', kasutades metateavet."
            GeneratePlaylist = "Loob .m3u esitusloendi kõigi muusikafailidega kaustast."
        }
    }
}

# Galvenā forma
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Size = New-Object System.Drawing.Size(700, 700)
$mainForm.StartPosition = "CenterScreen"
$mainForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Valodu izvēles izveidošana (labajā pusē)
$languageComboBox = New-Object System.Windows.Forms.ComboBox
$languageComboBox.Location = New-Object System.Drawing.Point(560, 20)
$languageComboBox.Size = New-Object System.Drawing.Size(100, 25)
$languageComboBox.Items.AddRange(@("lv", "en", "lt", "ee"))
$languageComboBox.SelectedItem = "lv"
$mainForm.Controls.Add($languageComboBox)

# BuyMeACoffee poga (kreisajā pusē, vienā līmenī ar valodas izvēlni)
$buyMeACoffeeButton = New-Object System.Windows.Forms.Button
$buyMeACoffeeButton.Size = New-Object System.Drawing.Size(174, 48)
$buyMeACoffeeButton.Location = New-Object System.Drawing.Point(20, 20) # Pretējā pusē valodas izvēlnei
$buyMeACoffeeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buyMeACoffeeButton.FlatAppearance.BorderSize = 0

# Funkcija attēla ielādei no interneta
function Load-ImageFromUrl {
    param ([string]$url)
    try {
        $webClient = New-Object System.Net.WebClient
        $imageBytes = $webClient.DownloadData($url)
        $memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
        $image = [System.Drawing.Image]::FromStream($memoryStream)
        return $image
    }
    catch {
        Write-Host "Kļūda ielādējot attēlu: $_"
        return $null
    }
}

# Ielādē BuyMeACoffee attēlu un pievieno to pogai
$coffeeImageUrl = "https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png"
$coffeeImage = Load-ImageFromUrl -url $coffeeImageUrl
if ($coffeeImage -ne $null) {
    $buyMeACoffeeButton.BackgroundImage = $coffeeImage
    $buyMeACoffeeButton.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $buyMeACoffeeButton.Add_Click({
        Start-Process "https://buymeacoffee.com/latvietis"
    })
    $mainForm.Controls.Add($buyMeACoffeeButton)
}
else {
    Write-Host "Neizdevās ielādēt BuyMeACoffee attēlu."
}

# Funkcija valodu maiņai
function UpdateLanguage {
    $selectedLang = $languageComboBox.SelectedItem
    $mainForm.Text = $translations[$selectedLang].Title
    $descriptionLabel.Text = $translations[$selectedLang].Description
    $selectFolderButton.Text = $translations[$selectedLang].SelectFolder
    $startButton.Text = $translations[$selectedLang].Start
    $selectedFolderLabel.Text = $translations[$selectedLang].NoFolderSelected
    $actionGroupBox.Text = $translations[$selectedLang].ActionGroup
    $organizeByArtistRadio.Text = $translations[$selectedLang].OrganizeByArtist
    $findEmptyFoldersRadio.Text = $translations[$selectedLang].FindEmptyFolders
    $findDuplicatesRadio.Text = $translations[$selectedLang].FindDuplicates
    $findVideoFilesRadio.Text = $translations[$selectedLang].FindVideoFiles
    $findImageFilesRadio.Text = $translations[$selectedLang].FindImageFiles
    $renameFilesRadio.Text = $translations[$selectedLang].RenameFiles
    $generatePlaylistRadio.Text = $translations[$selectedLang].GeneratePlaylist
    $testMetadataButton.Text = $translations[$selectedLang].TestMetadata
    $progressLabel.Text = $translations[$selectedLang].Ready

    # Atjauninām tooltip tekstus
    $organizeByArtistToolTip.SetToolTip($organizeByArtistInfo, $translations[$selectedLang].InfoTooltips.OrganizeByArtist)
    $findEmptyFoldersToolTip.SetToolTip($findEmptyFoldersInfo, $translations[$selectedLang].InfoTooltips.FindEmptyFolders)
    $findDuplicatesToolTip.SetToolTip($findDuplicatesInfo, $translations[$selectedLang].InfoTooltips.FindDuplicates)
    $findVideoFilesToolTip.SetToolTip($findVideoFilesInfo, $translations[$selectedLang].InfoTooltips.FindVideoFiles)
    $findImageFilesToolTip.SetToolTip($findImageFilesInfo, $translations[$selectedLang].InfoTooltips.FindImageFiles)
    $renameFilesToolTip.SetToolTip($renameFilesInfo, $translations[$selectedLang].InfoTooltips.RenameFiles)
    $generatePlaylistToolTip.SetToolTip($generatePlaylistInfo, $translations[$selectedLang].InfoTooltips.GeneratePlaylist)
}

$languageComboBox.Add_SelectedIndexChanged({ UpdateLanguage })

# Virsraksta izveidošana
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 118, 210)
$titleLabel.Location = New-Object System.Drawing.Point(20, 70)
$titleLabel.Size = New-Object System.Drawing.Size(660, 40)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$mainForm.Controls.Add($titleLabel)

# Apraksta izveidošana
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 110)
$descriptionLabel.Size = New-Object System.Drawing.Size(660, 30)
$descriptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$mainForm.Controls.Add($descriptionLabel)

# Mapes izvēles pogas izveidošana
$selectFolderButton = New-Object System.Windows.Forms.Button
$selectFolderButton.Location = New-Object System.Drawing.Point(20, 150)
$selectFolderButton.Size = New-Object System.Drawing.Size(150, 40)
$selectFolderButton.BackColor = [System.Drawing.Color]::FromArgb(25, 118, 210)
$selectFolderButton.ForeColor = [System.Drawing.Color]::White
$selectFolderButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$mainForm.Controls.Add($selectFolderButton)

# Izvēlētās mapes ceļa izveidošana
$selectedFolderLabel = New-Object System.Windows.Forms.Label
$selectedFolderLabel.Location = New-Object System.Drawing.Point(180, 150)
$selectedFolderLabel.Size = New-Object System.Drawing.Size(340, 40)
$selectedFolderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$selectedFolderLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$selectedFolderLabel.BackColor = [System.Drawing.Color]::White
$mainForm.Controls.Add($selectedFolderLabel)

# Sākt pogas izveidošana
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(530, 150)
$startButton.Size = New-Object System.Drawing.Size(150, 40)
$startButton.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$startButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$startButton.Enabled = $false
$mainForm.Controls.Add($startButton)

# Darbību grupas izveidošana
$actionGroupBox = New-Object System.Windows.Forms.GroupBox
$actionGroupBox.Location = New-Object System.Drawing.Point(20, 210)
$actionGroupBox.Size = New-Object System.Drawing.Size(660, 240)
$mainForm.Controls.Add($actionGroupBox)

# Radio pogas ar informācijas ikonām
$organizeByArtistRadio = New-Object System.Windows.Forms.RadioButton
$organizeByArtistRadio.Location = New-Object System.Drawing.Point(20, 30)
$organizeByArtistRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($organizeByArtistRadio)

$organizeByArtistInfo = New-Object System.Windows.Forms.Button
$organizeByArtistInfo.Text = "i"
$organizeByArtistInfo.Location = New-Object System.Drawing.Point(580, 30)
$organizeByArtistInfo.Size = New-Object System.Drawing.Size(30, 30)
$organizeByArtistInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$organizeByArtistInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($organizeByArtistInfo)

$findEmptyFoldersRadio = New-Object System.Windows.Forms.RadioButton
$findEmptyFoldersRadio.Location = New-Object System.Drawing.Point(20, 60)
$findEmptyFoldersRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($findEmptyFoldersRadio)

$findEmptyFoldersInfo = New-Object System.Windows.Forms.Button
$findEmptyFoldersInfo.Text = "i"
$findEmptyFoldersInfo.Location = New-Object System.Drawing.Point(580, 60)
$findEmptyFoldersInfo.Size = New-Object System.Drawing.Size(30, 30)
$findEmptyFoldersInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$findEmptyFoldersInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($findEmptyFoldersInfo)

$findDuplicatesRadio = New-Object System.Windows.Forms.RadioButton
$findDuplicatesRadio.Location = New-Object System.Drawing.Point(20, 90)
$findDuplicatesRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($findDuplicatesRadio)

$findDuplicatesInfo = New-Object System.Windows.Forms.Button
$findDuplicatesInfo.Text = "i"
$findDuplicatesInfo.Location = New-Object System.Drawing.Point(580, 90)
$findDuplicatesInfo.Size = New-Object System.Drawing.Size(30, 30)
$findDuplicatesInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$findDuplicatesInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($findDuplicatesInfo)

$findVideoFilesRadio = New-Object System.Windows.Forms.RadioButton
$findVideoFilesRadio.Location = New-Object System.Drawing.Point(20, 120)
$findVideoFilesRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($findVideoFilesRadio)

$findVideoFilesInfo = New-Object System.Windows.Forms.Button
$findVideoFilesInfo.Text = "i"
$findVideoFilesInfo.Location = New-Object System.Drawing.Point(580, 120)
$findVideoFilesInfo.Size = New-Object System.Drawing.Size(30, 30)
$findVideoFilesInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$findVideoFilesInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($findVideoFilesInfo)

$findImageFilesRadio = New-Object System.Windows.Forms.RadioButton
$findImageFilesRadio.Location = New-Object System.Drawing.Point(20, 150)
$findImageFilesRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($findImageFilesRadio)

$findImageFilesInfo = New-Object System.Windows.Forms.Button
$findImageFilesInfo.Text = "i"
$findImageFilesInfo.Location = New-Object System.Drawing.Point(580, 150)
$findImageFilesInfo.Size = New-Object System.Drawing.Size(30, 30)
$findImageFilesInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$findImageFilesInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($findImageFilesInfo)

$renameFilesRadio = New-Object System.Windows.Forms.RadioButton
$renameFilesRadio.Location = New-Object System.Drawing.Point(20, 180)
$renameFilesRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($renameFilesRadio)

$renameFilesInfo = New-Object System.Windows.Forms.Button
$renameFilesInfo.Text = "i"
$renameFilesInfo.Location = New-Object System.Drawing.Point(580, 180)
$renameFilesInfo.Size = New-Object System.Drawing.Size(30, 30)
$renameFilesInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$renameFilesInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($renameFilesInfo)

$generatePlaylistRadio = New-Object System.Windows.Forms.RadioButton
$generatePlaylistRadio.Location = New-Object System.Drawing.Point(20, 210)
$generatePlaylistRadio.Size = New-Object System.Drawing.Size(550, 30)
$actionGroupBox.Controls.Add($generatePlaylistRadio)

$generatePlaylistInfo = New-Object System.Windows.Forms.Button
$generatePlaylistInfo.Text = "i"
$generatePlaylistInfo.Location = New-Object System.Drawing.Point(580, 210)
$generatePlaylistInfo.Size = New-Object System.Drawing.Size(30, 30)
$generatePlaylistInfo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$generatePlaylistInfo.ForeColor = [System.Drawing.Color]::White
$actionGroupBox.Controls.Add($generatePlaylistInfo)

# ToolTip kontrole
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 100
$toolTip.AutoPopDelay = 5000

# Pievienojam tooltip katrai informācijas pogai
$organizeByArtistToolTip = $toolTip
$findEmptyFoldersToolTip = $toolTip
$findDuplicatesToolTip = $toolTip
$findVideoFilesToolTip = $toolTip
$findImageFilesToolTip = $toolTip
$renameFilesToolTip = $toolTip
$generatePlaylistToolTip = $toolTip

# Testēšanas poga
$testMetadataButton = New-Object System.Windows.Forms.Button
$testMetadataButton.Location = New-Object System.Drawing.Point(20, 460)
$testMetadataButton.Size = New-Object System.Drawing.Size(150, 30)
$testMetadataButton.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
$testMetadataButton.ForeColor = [System.Drawing.Color]::Black
$testMetadataButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$mainForm.Controls.Add($testMetadataButton)

# Progresa indikators
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 500)
$progressBar.Size = New-Object System.Drawing.Size(660, 30)
$progressBar.Style = "Continuous"
$progressBar.Value = 0
$mainForm.Controls.Add($progressBar)

# Progresa apraksts
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(20, 535)
$progressLabel.Size = New-Object System.Drawing.Size(660, 20)
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$mainForm.Controls.Add($progressLabel)

# Rezultātu žurnāls
$logTextBox = New-Object System.Windows.Forms.RichTextBox
$logTextBox.Location = New-Object System.Drawing.Point(20, 560)
$logTextBox.Size = New-Object System.Drawing.Size(660, 120)
$logTextBox.ReadOnly = $true
$logTextBox.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$mainForm.Controls.Add($logTextBox)

# Funkcijas
function Log {
    param ([string]$message, [string]$color = "Black")
    $logTextBox.SelectionColor = [System.Drawing.Color]::FromName($color)
    $logTextBox.AppendText("$message`r`n")
    $logTextBox.ScrollToCaret()
    $logTextBox.Refresh()
}

function UpdateProgress {
    param ([int]$value, [string]$status)
    $progressBar.Value = $value
    $progressLabel.Text = "$status - $value%"
    $progressBar.Refresh()
    $progressLabel.Refresh()
}

function GetArtistName {
    param ([string]$filePath)
    try {
        $artist = $null
        $fileExt = [System.IO.Path]::GetExtension($filePath).ToLower()
        $musicExts = @(".mp3", ".flac", ".wav", ".aac", ".m4a", ".ogg", ".wma")
        
        if ($musicExts -notcontains $fileExt) {
            Log "Fails nav atbalstīts mūzikas formāts $filePath" "Yellow"
            return "Nezināms"
        }
        
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = Split-Path $filePath -Parent
            $file = Split-Path $filePath -Leaf
            $folderObj = $shell.Namespace($folder)
            $fileObj = $folderObj.ParseName($file)
            
            $artistIndices = @(
                @{Index = 13; Name = "Artist"}, @{Index = 20; Name = "Contributing artists"}, 
                @{Index = 275; Name = "Album artist"}, @{Index = 16; Name = "Album artist older"}
            )
            
            foreach ($artIdx in $artistIndices) {
                $propValue = $folderObj.GetDetailsOf($fileObj, $artIdx.Index)
                if (-not [string]::IsNullOrEmpty($propValue)) {
                    $artist = $propValue
                    Log "Atrasts mākslinieks indeksā $($artIdx.Index) ($($artIdx.Name)) = $artist" "Green"
                    break
                }
            }
            
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fileObj) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folderObj) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        catch {
            Log "Kļūda piekļūstot Shell API $($_.Exception.Message)" "Red"
        }
        
        if ([string]::IsNullOrEmpty($artist)) {
            try {
                $wmp = New-Object -ComObject WMPlayer.OCX
                $wmp.settings.autoStart = $false
                $media = $wmp.newMedia($filePath)
                if ($media.getItemInfo("Artist")) {
                    $artist = $media.getItemInfo("Artist")
                    Log "Atrasts mākslinieks no WMP = $artist" "Green"
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($media) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wmp) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            catch {
                Log "Kļūda izmantojot WMP API $($_.Exception.Message)" "Red"
            }
        }
        
        if ([string]::IsNullOrEmpty($artist)) {
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
            if ($fileName -match "^(.*?)\s*[-\.]") {
                $artist = $matches[1].Trim()
                Log "Atrasts mākslinieks no faila nosaukuma = $artist" "Yellow"
            }
        }
        
        if ([string]::IsNullOrEmpty($artist)) {
            return "Nezināms"
        }
        
        if ($artist -match "(.*?)\s*[;,&]") {
            $artist = $matches[1].Trim()
        }
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        $invalidChars | ForEach-Object { $artist = $artist.Replace($_, "_") }
        $artist = $artist -replace '[\\/:*?"<>|]', '_' -replace '[\s\.]+$', ''
        
        if ([string]::IsNullOrWhiteSpace($artist)) {
            return "Nezināms"
        }
        return $artist
    }
    catch {
        Log "Kļūda nolasot mākslinieka datus $filePath $($_.Exception.Message)" "Red"
        return "Nezināms"
    }
}

function GetSongTitle {
    param ([string]$filePath)
    try {
        $title = $null
        $shell = New-Object -ComObject Shell.Application
        $folder = Split-Path $filePath -Parent
        $file = Split-Path $filePath -Leaf
        $folderObj = $shell.Namespace($folder)
        $fileObj = $folderObj.ParseName($file)
        
        $titleIndices = @(
            @{Index = 21; Name = "Title"}, @{Index = 2; Name = "Name"}
        )
        
        foreach ($idx in $titleIndices) {
            $propValue = $folderObj.GetDetailsOf($fileObj, $idx.Index)
            if (-not [string]::IsNullOrEmpty($propValue)) {
                $title = $propValue
                Log "Atrasts dziesmas nosaukums indeksā $($idx.Index) ($($idx.Name)) = $title" "Green"
                break
            }
        }
        
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fileObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folderObj) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        [System.GC]::Collect()
        
        if ([string]::IsNullOrEmpty($title)) {
            $title = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
            Log "Izmantots faila nosaukums kā dziesmas nosaukums = $title" "Yellow"
        }
        
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        $invalidChars | ForEach-Object { $title = $title.Replace($_, "_") }
        $title = $title -replace '[\\/:*?"<>|]', '_' -replace '[\s\.]+$', ''
        
        return $title
    }
    catch {
        Log "Kļūda nolasot dziesmas nosaukumu $filePath $($_.Exception.Message)" "Red"
        return "Nezināms"
    }
}

function OrganizeByArtist {
    param ([string]$folderPath)
    UpdateProgress 0 "Tiek analizēta mape"
    Log "Sākam failu analīzi..." "Blue"
    
    $musicFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object {
        $_.Extension -match "\.(mp3|flac|wav|aac|m4a|ogg|wma)$"
    }
    
    $totalFiles = $musicFiles.Count
    Log "Atrasti $totalFiles mūzikas faili" "Green"
    
    if ($totalFiles -eq 0) {
        Log "Nav atrasts neviens mūzikas fails" "Red"
        UpdateProgress 100 "Pabeigts"
        return
    }
    
    $processedFiles = 0
    foreach ($file in $musicFiles) {
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100)
        UpdateProgress $progress "Tiek apstrādāti faili"
        
        $artist = GetArtistName $file.FullName
        $targetFolder = Join-Path $folderPath $artist
        
        if (-not (Test-Path $targetFolder)) {
            New-Item -Path $targetFolder -ItemType Directory | Out-Null
            Log "Izveidota mape $artist" "Blue"
        }
        
        $targetPath = Join-Path $targetFolder $file.Name
        try {
            Move-Item -Path $file.FullName -Destination $targetPath -Force
            Log "Pārvietots $file.Name uz mapi $artist" "Green"
        }
        catch {
            Log "Kļūda pārvietojot failu $file.Name $($_.Exception.Message)" "Red"
        }
    }
    
    UpdateProgress 100 "Pabeigts"
    Log "Failu kārtošana pabeigta!" "Green"
}

function FindEmptyFolders {
    param ([string]$folderPath)
    UpdateProgress 0 "Tiek analizēta mape"
    Log "Sākam tukšo mapju meklēšanu..." "Blue"
    
    $allFolders = Get-ChildItem -Path $folderPath -Directory -Recurse | Sort-Object FullName -Descending
    $totalFolders = $allFolders.Count
    Log "Atrastas $totalFolders mapes" "Green"
    
    $emptyFolders = 0
    foreach ($folder in $allFolders) {
        if ((Get-ChildItem -Path $folder.FullName -Recurse -File).Count -eq 0) {
            try {
                Remove-Item -Path $folder.FullName -Force
                Log "Dzēsta tukšā mape $folder.FullName" "Blue"
                $emptyFolders++
            }
            catch {
                Log "Kļūda dzēšot mapi $folder.FullName $($_.Exception.Message)" "Red"
            }
        }
    }
    
    Log "Dzēstas $emptyFolders tukšās mapes" "Green"
    UpdateProgress 100 "Pabeigts"
}

function FindDuplicates {
    param ([string]$folderPath)
    UpdateProgress 0 "Tiek analizēta mape"
    Log "Sākam dublikātu meklēšanu..." "Blue"
    
    $musicFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object {
        $_.Extension -match "\.(mp3|flac|wav|aac|m4a|ogg|wma)$"
    }
    
    $totalFiles = $musicFiles.Count
    Log "Atrasti $totalFiles mūzikas faili" "Green"
    
    $fileHashes = @{}
    $processedFiles = 0
    foreach ($file in $musicFiles) {
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100)
        UpdateProgress $progress "Tiek aprēķinātas kontrolsummas"
        
        $hash = (Get-FileHash -Path $file.FullName -Algorithm MD5).Hash
        if ($fileHashes.ContainsKey($hash)) {
            $fileHashes[$hash] += $file
        }
        else {
            $fileHashes[$hash] = @($file)
        }
    }
    
    $duplicates = $fileHashes.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    $totalDuplicateFiles = ($duplicates | ForEach-Object { $_.Value.Count - 1 } | Measure-Object -Sum).Sum
    Log "Atrasti $totalDuplicateFiles dublēti faili" "Green"
    
    if ($totalDuplicateFiles -eq 0) {
        Log "Nav atrasts neviens dublikāts" "Yellow"
        UpdateProgress 100 "Pabeigts"
        return
    }
    
    $duplicateForm = New-Object System.Windows.Forms.Form
    $duplicateForm.Text = "Atrasti dublikāti"
    $duplicateForm.Size = New-Object System.Drawing.Size(900, 600)
    $duplicateForm.StartPosition = "CenterScreen"
    $duplicateForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    $duplicateLabel = New-Object System.Windows.Forms.Label
    $duplicateLabel.Text = "Atrasti $totalDuplicateFiles dublēti faili. Atzīmējiet dzēšamos:"
    $duplicateLabel.Location = New-Object System.Drawing.Point(20, 20)
    $duplicateLabel.Size = New-Object System.Drawing.Size(860, 30)
    $duplicateForm.Controls.Add($duplicateLabel)
    
    $duplicatePanel = New-Object System.Windows.Forms.Panel
    $duplicatePanel.Location = New-Object System.Drawing.Point(20, 60)
    $duplicatePanel.Size = New-Object System.Drawing.Size(860, 440)
    $duplicatePanel.AutoScroll = $true
    $duplicateForm.Controls.Add($duplicatePanel)
    
    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Text = "Dzēst atzīmētos"
    $deleteButton.Location = New-Object System.Drawing.Point(700, 510)
    $deleteButton.Size = New-Object System.Drawing.Size(180, 40)
    $deleteButton.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $deleteButton.ForeColor = [System.Drawing.Color]::White
    $duplicateForm.Controls.Add($deleteButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Atcelt"
    $cancelButton.Location = New-Object System.Drawing.Point(510, 510)
    $cancelButton.Size = New-Object System.Drawing.Size(180, 40)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $duplicateForm.Controls.Add($cancelButton)
    
    $yPos = 10
    $checkBoxes = @{}
    $i = 1
    foreach ($duplicate in $duplicates) {
        $files = $duplicate.Value
        $groupBox = New-Object System.Windows.Forms.GroupBox
        $groupBox.Text = "Dublētu grupa #$i"
        $groupBox.Location = New-Object System.Drawing.Point(10, $yPos)
        $groupBox.Size = New-Object System.Drawing.Size(820, 30 + $files.Count * 30)
        $duplicatePanel.Controls.Add($groupBox)
        
        $fileYPos = 20
        foreach ($file in $files) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = $file.FullName
            $checkBox.Location = New-Object System.Drawing.Point(10, $fileYPos)
            $checkBox.Size = New-Object System.Drawing.Size(800, 25)
            $groupBox.Controls.Add($checkBox)
            if ($fileYPos -gt 20) { $checkBox.Checked = $true }
            $checkBoxes[$checkBox] = $file
            $fileYPos += 30
        }
        $yPos += $groupBox.Height + 10
        $i++
    }
    
    $cancelButton.Add_Click({ $duplicateForm.Close() })
    $deleteButton.Add_Click({
        $selectedFiles = $checkBoxes.GetEnumerator() | Where-Object { $_.Key.Checked } | ForEach-Object { $_.Value }
        $totalSelected = $selectedFiles.Count
        
        if ($totalSelected -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nav atzīmēts neviens fails", "Informācija", "OK", "Information")
            return
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show("Vai tiešām dzēst $totalSelected failus?", "Apstiprinājums", "YesNo", "Warning")
        if ($result -eq "Yes") {
            $deletedCount = 0
            foreach ($file in $selectedFiles) {
                try {
                    Remove-Item -Path $file.FullName -Force
                    Log "Dzēsts dublikāts $file.FullName" "Green"
                    $deletedCount++
                }
                catch {
                    Log "Kļūda dzēšot $file.FullName $($_.Exception.Message)" "Red"
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Dzēsti $deletedCount no $totalSelected failiem", "Informācija", "OK", "Information")
            $duplicateForm.Close()
        }
    })
    
    UpdateProgress 100 "Pabeigts"
    $duplicateForm.ShowDialog()
}

function FindMediaFiles {
    param ([string]$folderPath, [string]$mediaType)
    UpdateProgress 0 "Tiek analizēta mape"
    
    if ($mediaType -eq "Video") {
        Log "Sākam video failu meklēšanu..." "Blue"
        $extensions = @(".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm")
        $filesLabel = "video faili"
    }
    else {
        Log "Sākam attēlu failu meklēšanu..." "Blue"
        $extensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp")
        $filesLabel = "attēli"
    }
    
    $mediaFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object { $extensions -contains $_.Extension.ToLower() }
    $totalFiles = $mediaFiles.Count
    Log "Atrasti $totalFiles $filesLabel" "Green"
    
    if ($totalFiles -eq 0) {
        Log "Nav atrasts neviens $mediaType fails" "Yellow"
        UpdateProgress 100 "Pabeigts"
        return
    }
    
    $mediaForm = New-Object System.Windows.Forms.Form
    $mediaForm.Text = "Atrasti $filesLabel"
    $mediaForm.Size = New-Object System.Drawing.Size(900, 600)
    $mediaForm.StartPosition = "CenterScreen"
    $mediaForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    $mediaLabel = New-Object System.Windows.Forms.Label
    $mediaLabel.Text = "Atrasti $totalFiles $filesLabel. Atzīmējiet dzēšamos:"
    $mediaLabel.Location = New-Object System.Drawing.Point(20, 20)
    $mediaLabel.Size = New-Object System.Drawing.Size(860, 30)
    $mediaForm.Controls.Add($mediaLabel)
    
    $selectAllCheckBox = New-Object System.Windows.Forms.CheckBox
    $selectAllCheckBox.Text = "Atzīmēt visus"
    $selectAllCheckBox.Location = New-Object System.Drawing.Point(20, 50)
    $selectAllCheckBox.Size = New-Object System.Drawing.Size(860, 30)
    $mediaForm.Controls.Add($selectAllCheckBox)
    
    $mediaPanel = New-Object System.Windows.Forms.Panel
    $mediaPanel.Location = New-Object System.Drawing.Point(20, 80)
    $mediaPanel.Size = New-Object System.Drawing.Size(860, 420)
    $mediaPanel.AutoScroll = $true
    $mediaForm.Controls.Add($mediaPanel)
    
    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Text = "Dzēst atzīmētos"
    $deleteButton.Location = New-Object System.Drawing.Point(700, 510)
    $deleteButton.Size = New-Object System.Drawing.Size(180, 40)
    $deleteButton.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $deleteButton.ForeColor = [System.Drawing.Color]::White
    $mediaForm.Controls.Add($deleteButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Atcelt"
    $cancelButton.Location = New-Object System.Drawing.Point(510, 510)
    $cancelButton.Size = New-Object System.Drawing.Size(180, 40)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $mediaForm.Controls.Add($cancelButton)
    
    $yPos = 10
    $checkBoxes = @{}
    foreach ($file in $mediaFiles) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = $file.FullName
        $checkBox.Location = New-Object System.Drawing.Point(10, $yPos)
        $checkBox.Size = New-Object System.Drawing.Size(820, 25)
        $mediaPanel.Controls.Add($checkBox)
        $checkBoxes[$checkBox] = $file
        $yPos += 30
    }
    
    $selectAllCheckBox.Add_CheckedChanged({ foreach ($cb in $checkBoxes.Keys) { $cb.Checked = $selectAllCheckBox.Checked } })
    $cancelButton.Add_Click({ $mediaForm.Close() })
    $deleteButton.Add_Click({
        $selectedFiles = $checkBoxes.GetEnumerator() | Where-Object { $_.Key.Checked } | ForEach-Object { $_.Value }
        $totalSelected = $selectedFiles.Count
        
        if ($totalSelected -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nav atzīmēts neviens fails", "Informācija", "OK", "Information")
            return
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show("Vai tiešām dzēst $totalSelected failus?", "Apstiprinājums", "YesNo", "Warning")
        if ($result -eq "Yes") {
            $deletedCount = 0
            foreach ($file in $selectedFiles) {
                try {
                    Remove-Item -Path $file.FullName -Force
                    Log "Dzēsts fails $file.FullName" "Green"
                    $deletedCount++
                }
                catch {
                    Log "Kļūda dzēšot $file.FullName $($_.Exception.Message)" "Red"
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Dzēsti $deletedCount no $totalSelected failiem", "Informācija", "OK", "Information")
            $mediaForm.Close()
        }
    })
    
    UpdateProgress 100 "Pabeigts"
    $mediaForm.ShowDialog()
}

function RenameFilesByMetadata {
    param ([string]$folderPath)
    UpdateProgress 0 "Tiek analizēta mape"
    Log "Sākam failu pārdēvēšanu..." "Blue"
    
    $musicFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object {
        $_.Extension -match "\.(mp3|flac|wav|aac|m4a|ogg|wma)$"
    }
    
    $totalFiles = $musicFiles.Count
    Log "Atrasti $totalFiles mūzikas faili" "Green"
    
    if ($totalFiles -eq 0) {
        Log "Nav atrasts neviens mūzikas fails" "Red"
        UpdateProgress 100 "Pabeigts"
        return
    }
    
    $processedFiles = 0
    foreach ($file in $musicFiles) {
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100)
        UpdateProgress $progress "Tiek pārdēvēti faili"
        
        $artist = GetArtistName $file.FullName
        $title = GetSongTitle $file.FullName
        $newName = "$artist - $title$($file.Extension)"
        $newPath = Join-Path $file.DirectoryName $newName
        
        try {
            Rename-Item -Path $file.FullName -NewName $newPath -Force
            Log "Pārdēvēts $file.Name uz $newName" "Green"
        }
        catch {
            Log "Kļūda pārdēvējot $file.Name $($_.Exception.Message)" "Red"
        }
    }
    
    UpdateProgress 100 "Pabeigts"
    Log "Failu pārdēvēšana pabeigta!" "Green"
}

function GeneratePlaylist {
    param ([string]$folderPath)
    UpdateProgress 0 "Tiek analizēta mape"
    Log "Sākam atskaņošanas saraksta izveidi..." "Blue"
    
    $musicFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object {
        $_.Extension -match "\.(mp3|flac|wav|aac|m4a|ogg|wma)$"
    }
    
    $totalFiles = $musicFiles.Count
    Log "Atrasti $totalFiles mūzikas faili" "Green"
    
    if ($totalFiles -eq 0) {
        Log "Nav atrasts neviens mūzikas fails" "Red"
        UpdateProgress 100 "Pabeigts"
        return
    }
    
    $playlistPath = Join-Path $folderPath "MusicPlaylist.m3u"
    "#EXTM3U" | Out-File -FilePath $playlistPath -Encoding UTF8
    
    $processedFiles = 0
    foreach ($file in $musicFiles) {
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100)
        UpdateProgress $progress "Tiek veidots atskaņošanas saraksts"
        
        $artist = GetArtistName $file.FullName
        $title = GetSongTitle $file.FullName
        $relativePath = $file.FullName.Substring($folderPath.Length + 1)
        
        "#EXTINF:-1,$artist - $title" | Out-File -FilePath $playlistPath -Append -Encoding UTF8
        $relativePath | Out-File -FilePath $playlistPath -Append -Encoding UTF8
        Log "Pievienots sarakstam: $relativePath" "Green"
    }
    
    UpdateProgress 100 "Pabeigts"
    Log "Atskaņošanas saraksts izveidots: $playlistPath" "Green"
    [System.Windows.Forms.MessageBox]::Show("Atskaņošanas saraksts izveidots: $playlistPath", "Informācija", "OK", "Information")
}

# Notikumu apstrāde
$selectFolderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Izvēlieties mapi ar mūzikas failiem"
    $folderBrowser.RootFolder = "MyComputer"
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $selectedFolderLabel.Text = $folderBrowser.SelectedPath
        $startButton.Enabled = $true
    }
})

$startButton.Add_Click({
    $folderPath = $selectedFolderLabel.Text
    if (-not (Test-Path $folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Izvēlētā mape neeksistē", "Kļūda", "OK", "Error")
        return
    }
    
    $logTextBox.Clear()
    if ($organizeByArtistRadio.Checked) { OrganizeByArtist $folderPath }
    elseif ($findEmptyFoldersRadio.Checked) { FindEmptyFolders $folderPath }
    elseif ($findDuplicatesRadio.Checked) { FindDuplicates $folderPath }
    elseif ($findVideoFilesRadio.Checked) { FindMediaFiles $folderPath "Video" }
    elseif ($findImageFilesRadio.Checked) { FindMediaFiles $folderPath "Image" }
    elseif ($renameFilesRadio.Checked) { RenameFilesByMetadata $folderPath }
    elseif ($generatePlaylistRadio.Checked) { GeneratePlaylist $folderPath }
})

$testMetadataButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Mūzikas faili|*.mp3;*.flac;*.wav;*.aac;*.m4a;*.ogg;*.wma|Visi faili|*.*"
    $openFileDialog.Title = "Izvēlieties failu metadatu testēšanai"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $artist = GetArtistName $openFileDialog.FileName
        $title = GetSongTitle $openFileDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Mākslinieks = $artist`nDziesma = $title", "Metadatu tests", "OK", "Information")
    }
})

# Informācijas pogu notikumi
$organizeByArtistInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.OrganizeByArtist, "Informācija", "OK", "Information")
})
$findEmptyFoldersInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.FindEmptyFolders, "Informācija", "OK", "Information")
})
$findDuplicatesInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.FindDuplicates, "Informācija", "OK", "Information")
})
$findVideoFilesInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.FindVideoFiles, "Informācija", "OK", "Information")
})
$findImageFilesInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.FindImageFiles, "Informācija", "OK", "Information")
})
$renameFilesInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.RenameFiles, "Informācija", "OK", "Information")
})
$generatePlaylistInfo.Add_Click({
    $selectedLang = $languageComboBox.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($translations[$selectedLang].InfoTooltips.GeneratePlaylist, "Informācija", "OK", "Information")
})

# Sākotnējā valodas iestatīšana
UpdateLanguage

# Pārbaudām, vai forma ir inicializēta pirms izsaukuma
if ($null -eq $mainForm) {
    Write-Host "Kļūda: Galvenā forma nav inicializēta!"
} else {
    [System.Windows.Forms.Application]::Run($mainForm)
}