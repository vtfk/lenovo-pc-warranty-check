﻿Param (
    [String]$FilePath,
    [Int32]$SerialsPerRequest = 100,
    [String]$ConfigPath = "./lenovo-serial-config.json"
)

# DEBUG
$FilePath = "./pc kun serial.csv"

$Config = @{}

if (Test-Path $ConfigPath) {
    Write-Host "Getting config from `"$ConfigPath`""
    $Config = Get-Content -Path $ConfigPath | ConvertFrom-Json
    $Config.ClientID = $Config.ClientID | ConvertTo-SecureString
} else {
    Write-Host "### First-time setup ###"
    
    $Config = @{
        ClientID = Read-Host -Prompt "Please input Lenovo Client-ID: " -AsSecureString
        ApiUri = "https://supportapi.lenovo.com/v2.5/product"
    }

    $ConfigToFile = $Config
    $ConfigToFile.ClientID = $ConfigToFile.ClientID | ConvertFrom-SecureString
    
    Write-Host "Saving config to `"$ConfigPath`""
    $Config | ConvertTo-Json | Set-Content -Path $ConfigPath
}

Write-Host
Write-Host "### Input file path ###"
if (!$FilePath) {
    try {
        $FileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $FileDialog.Title = "Select a Excel document with Lenovo serial numbers..."
        $FileDialog.Filter = "SpreadSheet (*.xlsx)|*.xlsx"
        $FileDialog.InitialDirectory = "$PSScriptRoot"
        $FileDialog.ShowDialog()
        Write-Host "Selected: $($FileDialog.FileName)"
        $FilePath = $FileDialog.FileName
    }
    catch {
        $FilePath = ""
    }
}

if (!$FilePath) {
    Write-Host "Failed to show the open file dialog, please input the full path to the Excel or CSV file."
    Write-Host "Tip: You can drag the file onto this terminal window."
    $FilePath = Read-Host -Prompt "Excel/CSV file path: "
    $FilePath = $FilePath -replace "^['`"``]|['`"``]$", ""
    Write-Host "FP: $FilePath"
}

if (!$FilePath) {
    Write-Host "Failed to get path, unknown error, exiting.."
    exit 1
}

Write-Host
Write-Host "### Import serials ###"
try {
    if ($FilePath -match "`.xlsx$") {
        Write-Host "Importing Excel sheet..."
        $FileExt = "xlsx"
        $Computers = Import-Excel -Path $FilePath
    } elseif ($FilePath -match "`.csv$") {
        Write-Host "Importing CSV file..."
        $FileExt = "csv"
        $Computers = Import-CSV -Path $FilePath
    } else {
        Write-Host "Unknown file extension on file `"$FilePath`"!"
        Write-Host "Valid file extensions are `".xlsx`", `".csv`""
        exit 1
    }
}
catch {
    Write-Error "Failed to import file `"$FilePath`", exiting...`n"
    exit 1
}


try {
    Write-Host "Exracting serial numbers from column `"Serial Number`"..."
    # TODO: Report if Serial Number is not found
    $Serials = $Computers | Select-Object -ExpandProperty "Serial Number" -ErrorAction Stop
    Write-Host "Found $(@($Serials).Length) serial numbers"
}
catch {
    Write-Error "Couldn't find the `"Serial Number`" column, exiting...`n"
    exit 1
}

Write-Host
Write-Host "### Request warranty information ###"
$Headers = @{
    "ClientID" = $Config.ClientID | ConvertFrom-SecureString -AsPlainText -ErrorAction Stop
    "Content-Type" = "application/x-www-form-urlencoded"
}

$TotalIterations = [Math]::Ceiling($Serials.length / $SerialsPerRequest)
$TotalTimeTaken = 0

Write-Host "Requesting information from Lenovo APIs, this can take some time..."
$Warranties = For ($i = 0; $i -lt $TotalIterations -and $i -lt 10; $i++) {
    $StartIndex = $i * $SerialsPerRequest
    $EndIndex = $i * $SerialsPerRequest + $SerialsPerRequest - 1

    Write-Progress `
        -PercentComplete ($i / $TotalIterations * 100) `
        -Activity "Getting warranty information for serials: $($StartIndex + 1) - $($EndIndex + 1)" `
        -Status "Avg. response time: $([Math]::Round($TotalTimeTaken / ($i + 1)))ms" `
        -SecondsRemaining ([Math]::Round((($TotalTimeTaken / ($i + 1)) / 1000) * ($TotalIterations - $i)))
    
    $Body = @{
        "Serial" = $Serials[$StartIndex..$EndIndex] -join ","
    }

    $RequestStart = Get-Date

    $Response = Invoke-RestMethod -Method POST -Headers $Headers -Uri "$($Config.APIUri)" -Body $Body -ErrorAction Stop
    
    $RequestTime = $(Get-Date) - $RequestStart
    $RequestMs = [Math]::Round($RequestTime.TotalMilliseconds)
    $TotalTimeTaken += $RequestMs

    $Response 
}
Write-Host "Done! All requests where completed in $($TotalTimeTaken)ms."

$LocalCopyFileName = "$PSScriptRoot\lenovo-data-$(Get-Date -f "yyyy-mm-dd_hh-mm-ss").csv"
Write-Host "Saving a local copy of request data to `"$LocalCopyFileName`""
$Warranties | Export-Csv "$LocalCopyFileName"

#$Warranties = Import-Csv "lenovo-data-2020-30-02_10-30-18.csv"

$Formatted = ""
$Formatted = $Warranties | ForEach-Object {
    $IDSplit = $_.ID -split "/"
    $Warranty = $_.Warranty | Where-Object ID -eq "UCN"  #TODO: UCN may not exist EX. PF0IMTGK

    [PSCustomObject]@{
        ID             = $_.ID
        Serial         = $IDSplit[5]
        Name           = $IDSplit[2]
        Model          = $IDSplit[4]
        Manufacturer   = "Lenovo"
        #WarrantyStart = $Warranty.Start
        WarrantyEnd    = $Warranty.End
        #Released      = $_.Released
        #Purchased     = $_.Purchased
    }
}

Write-Host
Write-Host "### Save information ###"

$FileName = $FilePath.Split("/")[-1].Split(".")[0]
$NewFilePath = "$PSScriptRoot/$FileName-updated.$FileExt"

Write-Host "Saving file to: `"$NewFilePath`""
if (Test-Path "$NewFilePath") {
    $OverwritePrompt = Read-Host -Prompt "File exists, do you want to overwrite it? (y/N)"
    if ($OverwritePrompt -notmatch "y|yes") {
        exit
    }
}

if ($FileExt -eq "xlsx") {
    $Formatted | Export-Excel -ClearSheet -Path "$NewFilePath"
} elseif ($FileExt -eq "csv") {
    $Formatted | Export-Csv -Force -Path "$NewFilePath"
}
