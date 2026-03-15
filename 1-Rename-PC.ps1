#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Renames the computer to the BIOS Asset Tag, checking all known WMI locations.
#>

Write-Host "=== Rename PC to BIOS Asset Tag ===" -ForegroundColor Cyan
Write-Host ""

# ── QUERY ALL KNOWN LOCATIONS ─────────────────────────────────────────────────
Write-Host "Scanning all WMI sources for Asset Tag..." -ForegroundColor Yellow
Write-Host ""

$candidates = [ordered]@{}

# Source 1
try {
    $val = (Get-WmiObject -Class Win32_SystemEnclosure).SMBIOSAssetTag
    $candidates["Win32_SystemEnclosure.SMBIOSAssetTag"] = $val
} catch { $candidates["Win32_SystemEnclosure.SMBIOSAssetTag"] = "(query failed)" }

# Source 2
try {
    $val = (Get-WmiObject -Class Win32_ComputerSystemProduct).IdentifyingNumber
    $candidates["Win32_ComputerSystemProduct.IdentifyingNumber"] = $val
} catch { $candidates["Win32_ComputerSystemProduct.IdentifyingNumber"] = "(query failed)" }

# Source 3
try {
    $val = (Get-WmiObject -Class Win32_BIOS).SerialNumber
    $candidates["Win32_BIOS.SerialNumber"] = $val
} catch { $candidates["Win32_BIOS.SerialNumber"] = "(query failed)" }

# Print all found values
foreach ($key in $candidates.Keys) {
    $val = $candidates[$key]
    $color = if ([string]::IsNullOrWhiteSpace($val) -or $val -match "No Asset Tag|Default string|To be filled|Not Specified|\(query failed\)") { "Red" } else { "Green" }
    Write-Host "  $key" -ForegroundColor White
    Write-Host "    Value: '$val'" -ForegroundColor $color
    Write-Host ""
}

# ── FIND THE BEST CANDIDATE ───────────────────────────────────────────────────
$invalidPatterns = "No Asset Tag|Default string|To be filled|Not Specified|^\s*$"
$assetTag = $null

foreach ($key in $candidates.Keys) {
    $val = $candidates[$key]
    if (-not [string]::IsNullOrWhiteSpace($val) -and $val -notmatch $invalidPatterns -and $val -ne "(query failed)") {
        $assetTag = $val.Trim()
        Write-Host "Using value from: $key" -ForegroundColor Cyan
        Write-Host "Asset Tag: '$assetTag'" -ForegroundColor Cyan
        Write-Host ""
        break
    }
}

# ── RENAME OR REPORT ──────────────────────────────────────────────────────────
if ($null -eq $assetTag) {
    Write-Host "ERROR: No valid Asset Tag found in any WMI source." -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible reasons:" -ForegroundColor Yellow
    Write-Host "  - The Asset Tag field in BIOS has not been saved/applied properly" -ForegroundColor White
    Write-Host "  - The value is stored in a non-standard SMBIOS field on this model" -ForegroundColor White
    Write-Host "  - WMI repository may need to be rebuilt (run: winmgmt /resetrepository)" -ForegroundColor White
    Write-Host ""
    Write-Host "All raw values found above -- verify which one matches your BIOS entry." -ForegroundColor Yellow
} else {
    $currentName = $env:COMPUTERNAME
    Write-Host "Current computer name: $currentName" -ForegroundColor White
    Write-Host "Target computer name:  $assetTag" -ForegroundColor White
    Write-Host ""

    if ($currentName -eq $assetTag) {
        Write-Host "Computer is already named '$assetTag'. Nothing to do." -ForegroundColor Green
    } else {
        try {
            Rename-Computer -NewName $assetTag -Force -ErrorAction Stop
            Write-Host "SUCCESS: Computer will be renamed to '$assetTag' on next reboot." -ForegroundColor Green
        } catch {
            Write-Host "ERROR: Failed to rename computer." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "Done. Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
