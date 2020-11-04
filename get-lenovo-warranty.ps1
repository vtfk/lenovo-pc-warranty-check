Param (
    [String]$FilePath,
    [Int32]$SerialsPerRequest = 100,
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
        $FileDialog.Filter = "SpreadSheet (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv"
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
}
catch {
    Write-Error "Failed to import file `"$FilePath`", exiting...`n"
    Pause
    exit 1
}

try {
    Write-Host "Exracting serial numbers from column `"Serial Number`"..."
    $Serials = $Computers | Select-Object -ExpandProperty "Serial Number" -ErrorAction Stop
    Write-Host "Found $(@($Serials).Length) serial numbers in file"
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

Write-Host "Requesting information from Lenovo APIs, this can take some time..."
$Warranties = For ($i = 0; $i -lt $TotalIterations -and $i -lt 5; $i++) {
    try {
        $StartIndex = $i * $SerialsPerRequest
        $EndIndex = $i * $SerialsPerRequest + $SerialsPerRequest - 1
        
        $RequestStart = Get-Date

        Write-Progress `
            -PercentComplete ($i / $TotalIterations * 100) `
            -Activity "Getting warranty information for serials: $($StartIndex + 1) - $($EndIndex + 1)" `
            -Status "Avg. response time: $([Math]::Round($TotalTimeTaken / ($i + 1)))ms" `
            -SecondsRemaining ([Math]::Round((($TotalTimeTaken / ($i + 1)) / 1000) * ($TotalIterations - $i)))
        
        $Body = @{
            "Serial" = $Serials[$StartIndex..$EndIndex] -join ","
        }

        $Response = Invoke-RestMethod -Method POST -Headers $Headers -Uri "$($Config.APIUri)" -Body $Body -ErrorAction Stop

        $FormattedResponse = $Response | ForEach-Object {
            $IDSplit = $_.ID -split "/"
            $Warranty = $_.Warranty | Where-Object ID -eq "UCN"  #TODO: UCN may not exist EX. PF0IMTGK
        
            [PSCustomObject]@{
                ID                  = $_.ID
                "Serial Number"     = $IDSplit[-1]
                Name                = $IDSplit[2]
                Model               = $IDSplit[4]
                Manufacturer        = "Lenovo"
                #WarrantyStart      = $Warranty.Start
                WarrantyEnd         = $Warranty.End
                #Released           = $_.Released
                #Purchased          = $_.Purchased
            }
        }

        $FormattedResponse | Export-CSV -Append -Path $TempFilePath

        $FormattedResponse
    } catch {
        Write-Host "ERROR: Failed to get serials between $($StartIndex + 1) - $($EndIndex + 1)"
        Write-Host "Error message:"
        Write-Error ($Error[0])
    } finally {
        $RequestTime = $(Get-Date) - $RequestStart
        $RequestMs = [Math]::Round($RequestTime.TotalMilliseconds)
        $TotalTimeTaken += $RequestMs
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
    Write-Host "Some serials were invalid/not found and was not returned from the API."
    Write-Host "They are listed in `"$InvalidSerialsPath`""
    $InvalidSerials = ($Serials | Where-Object -FilterScript { $Warranties."Serial Number" -notcontains $_ })
    $InvalidSerials = foreach ($InvalidSerial in $InvalidSerials) {
        [PSCustomObject]@{
            "Serial Number" = $InvalidSerial
        }
    }
    $InvalidSerials | Export-Csv -Append -Path $InvalidSerialsPath
}

Write-Host
Write-Host "### Save information ###"
$FileName = $FilePath.Split("/")[-1].Split(".")[0]
$NewFilePath = "$PWD\$FileName-updated.$FileExt"

Write-Host "Saving file to: `"$NewFilePath`""
if (Test-Path "$NewFilePath") {
    $OverwritePrompt = Read-Host -Prompt "File exists, do you want to overwrite it? (y/N)"
    if ($OverwritePrompt -notmatch "y|yes") {
        exit
    }
}

if ($FileExt -eq "xlsx") {
    $Warranties | Export-Excel -ClearSheet -Path "$NewFilePath"
} elseif ($FileExt -eq "csv") {
    $Warranties | Export-Csv -Force -Path "$NewFilePath"
}

Remove-Item -Path $TempFilePath
