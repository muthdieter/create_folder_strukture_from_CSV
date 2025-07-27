Clear-Host

$ScriptName = "create_folder_strukture_from_CSV"
$scriptVersion = "V_1_0_1"
$scriptGitHub = "https://github.com/muthdieter"
$scriptDate = "7.2025"

mode 300

Write-Host ""
Write-Host "             ____  __  __"
Write-Host "            |  _ \|  \/  |"
Write-Host "            | | | | |\/| |"
Write-Host "            | |_| | |  | |"
Write-Host "            |____/|_|  |_|"
Write-Host "   "
Write-Host ""
Write-Host "       $scriptGitHub " -ForegroundColor magenta
Write-Host ""
Write-Host "       $ScriptName   " -ForegroundColor Green
write-Host "       $scriptVersion" -ForegroundColor Green
write-host "       $scriptDate   " -ForegroundColor Green
Write-Host ""
Write-Host ""
Write-Host ""
Pause

Add-Type -AssemblyName System.Windows.Forms

function Select-File {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select a folder structure file (.csv, .xlsx, .txt)"
    $dialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|Excel Files (*.xlsx)|*.xlsx"
    if ($dialog.ShowDialog() -eq "OK") {
        return $dialog.FileName
    }
    return $null
}

function Select-Folder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the base folder (local or network)"
    if ($dialog.ShowDialog() -eq "OK") {
        return $dialog.SelectedPath
    }
    return $null
}

# Step 1: Select source file
$inputFile = Select-File
if (-not $inputFile) {
    Write-Host "❌ Cancelled file selection." -ForegroundColor Red
    exit 1
}

# Step 2: Select target folder
$basePath = Select-Folder
if (-not $basePath) {
    Write-Host "❌ Cancelled folder selection." -ForegroundColor Red
    exit 1
}

# Step 3: Load file content
if ($inputFile -like "*.csv" -or $inputFile -like "*.txt") {
    $data = Import-Csv -Path $inputFile -Encoding Default
}
elseif ($inputFile -like "*.xlsx") {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel
    $data = Import-Excel -Path $inputFile
}
else {
    Write-Host "❌ Unsupported file type." -ForegroundColor Red
    exit 1
}

# Step 4: Collect folder paths
$foldersToCreate = @()

foreach ($row in $data) {
    $path = $basePath

    foreach ($col in $row.PSObject.Properties.Name) {
        $segment = $row.$col
        if ($segment -and $segment.Trim() -ne "") {
            $path = Join-Path $path $segment.Trim()
            if (-not (Test-Path $path) -and -not ($foldersToCreate -contains $path)) {
                $foldersToCreate += $path
            }
        }
    }
}

# Step 5: Preview
Write-Host "`n📋 The following folders will be created:`n" -ForegroundColor Cyan
$foldersToCreate | ForEach-Object { Write-Host " - $_" }

# Step 6: Confirm
$response = Read-Host "`n❓ Proceed with creating these folders? (Y/N)"
if ($response -notmatch '^[Yy]') {
    Write-Host "❌ Operation cancelled." -ForegroundColor Yellow
    exit
}

# Step 7: Create folders
foreach ($folder in $foldersToCreate) {
    try {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Write-Host "✅ Created: $folder"
    } catch {
        Write-Host "❌ Failed: $folder -- $_" -ForegroundColor Red
    }
}

Write-Host "`n🎉 DONE: All folders processed." -ForegroundColor Green
