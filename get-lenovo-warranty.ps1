﻿Param (
    [String]$FilePath,
    [Int32]$SerialsPerRequest = 100,
    [Int32]$MaxAttempts = 3,
    [String]$ConfigPath = ".\lenovo-serial-config.json",
    [String]$TempFilePath = ".\lenovo-serials.tmp.csv",
    [String]$InvalidSerialsPath = ".\lenovo-serials.invalid.csv"
)

try {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
}
catch {}

Write-Host "###################################################"
Write-Host "##         Warranty check for Lenovo PCs         ##"
Write-Host "## https://github.com/vtfk/lenovo-warranty-check ##"
Write-Host "###################################################"
Write-Host

Write-Host "Checking for the required modules..."
$ModulesRequired = @("ImportExcel", "Join-Object")
$ModulesInstalled = Get-InstalledModule -Name $ModulesRequired -ErrorAction Ignore

if ($ModulesInstalled.length -lt 2) {
    Write-Host "Installing modules:"
    Write-Host " - $($ModulesRequired -join "`n - ")"
    Write-Host
    Write-Host "Please accept the following questions by typing `"A`". It might take a few seconds.."
    Write-Host

    Install-Module -Name $ModulesRequired -Scope CurrentUser -ErrorAction Stop
}

Function SecureStringToString($Value)
{
    [System.IntPtr] $Bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($value);
    try
    {
        [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($Bstr);
    }
    finally
    {
        [System.Runtime.InteropServices.Marshal]::FreeBSTR($Bstr);
    }
}

$Config = @{}

if (Test-Path $ConfigPath) {
    Write-Host "Getting config from `"$ConfigPath`""
    $Config = Get-Content -Path $ConfigPath | ConvertFrom-Json
    $Config.ClientID = $Config.ClientID | ConvertTo-SecureString
} else {
    Write-Host "### First-time setup ###"
    
    $Config = @{
        ClientID = Read-Host -Prompt "Please input Lenovo Client-ID" -AsSecureString
        ApiUri = "https://supportapi.lenovo.com/v2.5/product"
    }

    if ($Config.ClientID.length -lt 1) {
        Write-Host "Please input a valid key, exiting..."
        Pause
        exit 1
    }

    $ConfigToFile = $Config.PSObject.Copy()
    $ConfigToFile.ClientID = $ConfigToFile.ClientID | ConvertFrom-SecureString
    
    Write-Host "Saving config to `"$ConfigPath`""
    $ConfigToFile | ConvertTo-Json | Set-Content -Path $ConfigPath
}

if (!$FilePath) {
    Write-Host
    Write-Host "### Input file path ###"
    try {
        $FileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $FileDialog.Title = "Select a Excel document with Lenovo serial numbers..."
        $FileDialog.Filter = "Excel/CSV (*.xlsx, *.csv)|*.xlsx;*.csv|SpreadSheet (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv"
        $FileDialog.InitialDirectory = "$PWD"
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
    $FilePath = Read-Host -Prompt "Excel/CSV file path"
    $FilePath = $FilePath -replace "^['`"``]|['`"``]$", ""
}

if (!$FilePath) {
    Write-Host "Failed to get path, unknown error, exiting.."
    Pause
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
        Pause
        exit 1
    }
    Write-Host "Found $(@($Computers).length) computers in file"
    $Computers = $Computers | Where-Object -FilterScript { $_.Model -eq "Lenovo" -or !$_.Model }
    Write-Host "Filtered to $(@($Computers).length) Lenovo computers"
}
catch {
    Write-Error "Failed to import file `"$FilePath`", exiting...`n"
    Pause
    exit 1
}

try {
    Write-Host "Exracting serial numbers from column `"Serial Number`"..."
    $Serials = $Computers | Select-Object -ExpandProperty "Serial Number" -ErrorAction Stop
}
catch {
    Write-Error "Couldn't find the `"Serial Number`" column, exiting...`n"
    Pause
    exit 1
}

$TestedSerials = @()
if (Test-Path $TempFilePath) {
    try {
        $TestedSerials = Import-Csv -Path $TempFilePath
        Write-Host "Found unfinished job containing $($TestedSerials.length) checked serials.."
        $Serials = ($Serials | Where-Object -FilterScript { $TestedSerials."Serial Number" -notcontains $_ })
        Write-Host "Getting info for $($Serials.length) unchecked serials"
    }
    catch {
        Write-Host "Failed to import unfinished job, ignoring.."
    }
}

Write-Host
Write-Host "### Request warranty information ###"

$Headers = @{
    "ClientID" = SecureStringToString($Config.ClientID)
    "Content-Type" = "application/x-www-form-urlencoded"
}

$TotalIterations = [Math]::Ceiling($Serials.length / $SerialsPerRequest)
$TotalTimeTaken = 0
$FailedAttempts = 0

Write-Host "Requesting information from Lenovo's API, this can take some time..."
Write-Host -NoNewLine "`r|--------------------| 0% | Time left: N/A | Serials: 1 - 100 | Avg. response time: N/A"
$Warranties = For ($i = 0; $i -lt $TotalIterations; $i++) {
    try {
        if (($MaxAttempts -ne 0) -and ($FailedAttempts -ge $MaxAttempts)) {
            Write-Host "The last $FailedAttempts requests has failed, retrying in 30 seconds..."
            Start-Sleep -Seconds 30
        }
        $StartIndex = $i * $SerialsPerRequest
        $EndIndex = $i * $SerialsPerRequest + $SerialsPerRequest - 1

        $RequestStart = Get-Date

        $Body = @{
            "Serial" = $Serials[$StartIndex..$EndIndex] -join ","
        }

        $Response = Invoke-RestMethod -Method POST -Headers $Headers -Uri "$($Config.APIUri)" -Body $Body -ErrorAction Stop

        $FormattedResponse = $Response | ForEach-Object {
            $IDSplit = $_.ID -split "/"

            $LastWarranty = $null
            $_.Warranty | ForEach-Object {
                if ($LastWarranty -eq $null) {
                    $LastWarranty = $_
                } elseif ($LastWarranty.End -lt $_.End) {
                    $LastWarranty = $_
                }
            }

            if (@($IDSplit).length -eq 6) {
                $PCSerial = $IDSplit[-1]
                $PCName = $IDSplit[2]
                $PCModel = $IDSplit[4]
            } else {
                $PCSerial = $IDSplit[-1]
                $PCName = "N/A"
                $PCModel = "Legacy Machine"
            }

            [PSCustomObject]@{
                ID                  = $_.ID
                "Serial Number"     = $PCSerial
                Name                = $PCName
                Model               = $PCModel
                Manufacturer        = "Lenovo"
                "Warranty ID"          = $LastWarranty.ID
                "Warranty Name"        = $LastWarranty.Name
                "Warranty Start"      = $Warranty.Start
                "Warranty End"         = $LastWarranty.End
            }
        }

        $FormattedResponse | Export-CSV -NoTypeInformation -Append -Path $TempFilePath

        $FormattedResponse
        $FailedAttempts = 0
    } catch {
        $FailedAttempts++
        Write-Host
        Write-Host "ERROR: Failed to get serials between $($StartIndex + 1) - $($EndIndex + 1)"
        Write-Host "Error message:"
        Write-Error ($Error[0])
    } finally {
        $RequestTime = $(Get-Date) - $RequestStart
        $RequestMs = [Math]::Round($RequestTime.TotalMilliseconds)
        $TotalTimeTaken += $RequestMs
        
        $Progress = $i / $TotalIterations
        $LengthOfBar = 20
        $ProgressBarString = (
            "|" + 
            ("#" * ($LengthOfBar * $Progress)) + 
            ("-" * ($LengthOfBar * (1 - $Progress))) +
            "| " + [Math]::Round($Progress * 100) + "%" + " | " +
            "Time left: $([Math]::Round((($TotalTimeTaken / ($i + 1)) / 1000) * ($TotalIterations - $i)))s | " +
            "Serials: $($StartIndex + 1) - $($EndIndex + 1) | " +
            "Avg. response time: $([Math]::Round($TotalTimeTaken / ($i + 1)))ms"
        )
        if ($Progress -ge 1) {
            $ProgressBarString = $ProgressBarString + "`n"
        }

        Write-Host -NoNewline ("`r" + $ProgressBarString)
    }
}
Write-Host
Write-Host "Done! All requests where completed in $($TotalTimeTaken)ms."

if ($TestedSerials.length -gt 0) {
    Write-Host "Merging unfinished job with this job.."
    $Warranties = $TestedSerials + $Warranties
}

if (@($Serials).length -gt @($Warranties).length) {
    Write-Host
    Write-Host "### Some Invalid Serials ###"
    $InvalidSerials = ($Serials | Where-Object -FilterScript { $Warranties."Serial Number" -notcontains $_ })
    $InvalidSerials = foreach ($InvalidSerial in $InvalidSerials) {
        [PSCustomObject]@{
            "Serial Number" = $InvalidSerial
        }
    }
    Write-Host "Some serials ($(@($InvalidSerials).length)) were invalid/not found and was not returned from the API."
    Write-Host "They are listed in `"$InvalidSerialsPath`""
    $InvalidSerials | Export-Csv -NoTypeInformation -Append -Path $InvalidSerialsPath
}

$MergedWarranties = [PSCustomObject]@{}
try {
    Write-Host "### Merging data ###"
    $MergedWarranties = Join-Object -Left $Computers `
        -Right $Warranties `
        -LeftJoinProperty "Serial Number" `
        -RightJoinProperty "Serial Number" `
        -Type "AllInLeft" `
        -ExcludeLeftProperties "Model", "Manufacturer" `
        -RightProperties "Name", "Model", "Manufacturer", "Warranty End"
}
catch {
    Write-Host "Failed while merging data.."
    Write-Host "Rerun the script or find the unmerged data in `"$TempFilePath`""
    Write-Host "Error message:"
    Write-Error $Error[0]
    Pause
    exit 1
}

Write-Host
Write-Host "### Save information ###" 
$FileName = $FilePath.Replace("/", "\").Split("\")[-1].Split(".")[0] 
$NewFilePath = "$PWD\$FileName-updated.$FileExt"

Write-Host "Saving file to: `"$NewFilePath`""
if (Test-Path "$NewFilePath") {
    $OverwritePrompt = Read-Host -Prompt "File exists, do you want to overwrite it? (y/N)"
    if ($OverwritePrompt -notmatch "y|yes") {
        Pause
        exit
    }
}

if ($FileExt -eq "xlsx") {
    $MergedWarranties | Export-Excel -ClearSheet -Path "$NewFilePath"
} elseif ($FileExt -eq "csv") {
    $MergedWarranties | Export-Csv -NoTypeInformation -Force -Path "$NewFilePath"
}

Remove-Item -Path $TempFilePath
Pause