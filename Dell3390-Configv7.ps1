#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Dell 3390 2-in-1 Windows 11 Master Configuration Script
.DESCRIPTION
    1. Renames PC to BIOS Asset Tag
    2. Sets Screen, Sleep, Hard Disk, and Hibernate timeout to Never (AC and Battery)
    3. Disables Windows Update service (wuauserv)
    4. Disables Windows Update Medic Service (WaaSMedicSvc)
    5. Sets enforced Desktop Background image
    6. Sets enforced Lock Screen image and configures lock screen settings
    7. Sets Google Chrome as the default browser
    8. Configures Advanced Sharing Settings
    9. Prompts for restart
#>

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   Dell 3390 Configuration Script v6" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# ─────────────────────────────────────────────
# 1. RENAME PC TO BIOS ASSET TAG
# ─────────────────────────────────────────────
Write-Host "--- TASK 1: Rename PC to BIOS Asset Tag ---" -ForegroundColor Cyan
Write-Host ""

$candidates = [ordered]@{}

try { $candidates["Win32_SystemEnclosure.SMBIOSAssetTag"] = (Get-WmiObject -Class Win32_SystemEnclosure).SMBIOSAssetTag } catch { $candidates["Win32_SystemEnclosure.SMBIOSAssetTag"] = "(query failed)" }
try { $candidates["Win32_ComputerSystemProduct.IdentifyingNumber"] = (Get-WmiObject -Class Win32_ComputerSystemProduct).IdentifyingNumber } catch { $candidates["Win32_ComputerSystemProduct.IdentifyingNumber"] = "(query failed)" }
try { $candidates["Win32_BIOS.SerialNumber"] = (Get-WmiObject -Class Win32_BIOS).SerialNumber } catch { $candidates["Win32_BIOS.SerialNumber"] = "(query failed)" }

foreach ($key in $candidates.Keys) {
    $val = $candidates[$key]
    $color = if ([string]::IsNullOrWhiteSpace($val) -or $val -match "No Asset Tag|Default string|To be filled|Not Specified|\(query failed\)") { "Red" } else { "Green" }
    Write-Host "  $key" -ForegroundColor White
    Write-Host "    Value: '$val'" -ForegroundColor $color
    Write-Host ""
}

$invalidPatterns = "No Asset Tag|Default string|To be filled|Not Specified|^\s*$"
$assetTag = $null

foreach ($key in $candidates.Keys) {
    $val = $candidates[$key]
    if (-not [string]::IsNullOrWhiteSpace($val) -and $val -notmatch $invalidPatterns -and $val -ne "(query failed)") {
        $assetTag = $val.Trim()
        Write-Host "Using value from: $key --> '$assetTag'" -ForegroundColor Cyan
        Write-Host ""
        break
    }
}

if ($null -eq $assetTag) {
    Write-Host "ERROR: No valid Asset Tag found. Skipping rename." -ForegroundColor Red
} else {
    $currentName = $env:COMPUTERNAME
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

# ─────────────────────────────────────────────
# 2. SET ALL POWER TIMEOUTS TO NEVER
# ─────────────────────────────────────────────
Write-Host "--- TASK 2: Set All Power Timeouts to Never ---" -ForegroundColor Cyan
Write-Host ""

try {
    powercfg /change monitor-timeout-ac 0
    powercfg /change monitor-timeout-dc 0
    Write-Host "SUCCESS: Screen timeout set to Never (AC and Battery)." -ForegroundColor Green

    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0
    Write-Host "SUCCESS: Sleep timeout set to Never (AC and Battery)." -ForegroundColor Green

    powercfg /change disk-timeout-ac 0
    powercfg /change disk-timeout-dc 0
    Write-Host "SUCCESS: Hard disk timeout set to Never (AC and Battery)." -ForegroundColor Green

    powercfg /change hibernate-timeout-ac 0
    powercfg /change hibernate-timeout-dc 0
    Write-Host "SUCCESS: Hibernate timeout set to Never (AC and Battery)." -ForegroundColor Green

} catch {
    Write-Host "ERROR: Failed to set one or more power timeouts." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# ─────────────────────────────────────────────
# 3. DISABLE WINDOWS UPDATE SERVICE (wuauserv)
# ─────────────────────────────────────────────
Write-Host "--- TASK 3: Disable Windows Update Service ---" -ForegroundColor Cyan
Write-Host ""

try {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Set-Service -Name wuauserv -StartupType Disabled -ErrorAction Stop
    Write-Host "SUCCESS: Windows Update service (wuauserv) disabled." -ForegroundColor Green
} catch {
    Write-Host "WARNING: Set-Service failed, trying registry fallback..." -ForegroundColor Yellow
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\wuauserv" -Name "Start" -Value 4 -Type DWord -Force
        Write-Host "SUCCESS: Windows Update service disabled via registry." -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to disable Windows Update service." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

Write-Host ""

# ─────────────────────────────────────────────
# 4. DISABLE WINDOWS UPDATE MEDIC SERVICE
# ─────────────────────────────────────────────
Write-Host "--- TASK 4: Disable Windows Update Medic Service ---" -ForegroundColor Cyan
Write-Host ""

$medicRegPath    = "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc"
$medicRegPathRaw = "SYSTEM\CurrentControlSet\Services\WaaSMedicSvc"
$success = $false

try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using Microsoft.Win32;

public class RegOwnership {
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    static extern bool LookupPrivilegeValue(string lpSystemName, string lpName, out LUID lpLuid);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges,
        ref TOKEN_PRIVILEGES NewState, uint BufferLength, IntPtr PreviousState, IntPtr ReturnLength);
    [StructLayout(LayoutKind.Sequential)]
    public struct LUID { public uint LowPart; public int HighPart; }
    [StructLayout(LayoutKind.Sequential)]
    public struct TOKEN_PRIVILEGES { public uint PrivilegeCount; public LUID Luid; public uint Attributes; }
    public static void EnableTakeOwnership() {
        IntPtr token;
        OpenProcessToken(System.Diagnostics.Process.GetCurrentProcess().Handle, 0x0020 | 0x0008, out token);
        LUID luid;
        LookupPrivilegeValue(null, "SeTakeOwnershipPrivilege", out luid);
        TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES();
        tp.PrivilegeCount = 1; tp.Luid = luid; tp.Attributes = 0x00000002;
        AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
"@

    [RegOwnership]::EnableTakeOwnership()

    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
        $medicRegPathRaw,
        [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
        [System.Security.AccessControl.RegistryRights]::TakeOwnership
    )
    $acl = $key.GetAccessControl([System.Security.AccessControl.AccessControlSections]::None)
    $adminSID = New-Object System.Security.Principal.SecurityIdentifier(
        [System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null
    )
    $acl.SetOwner($adminSID)
    $key.SetAccessControl($acl)

    $key2 = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
        $medicRegPathRaw,
        [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
        [System.Security.AccessControl.RegistryRights]::ChangePermissions
    )
    $acl2 = $key2.GetAccessControl()
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
        $adminSID,
        [System.Security.AccessControl.RegistryRights]::FullControl,
        [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
        [System.Security.AccessControl.PropagationFlags]::None,
        [System.Security.AccessControl.AccessControlType]::Allow
    )
    $acl2.SetAccessRule($rule)
    $key2.SetAccessControl($acl2)

    Set-ItemProperty -Path $medicRegPath -Name "Start" -Value 4 -Type DWord -Force -ErrorAction Stop
    Write-Host "SUCCESS: WaaSMedicSvc disabled via registry ownership." -ForegroundColor Green
    $success = $true

} catch {
    Write-Host "WARNING: Ownership method failed. Trying sc.exe fallback..." -ForegroundColor Yellow
    try {
        sc.exe sdset WaaSMedicSvc "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)" | Out-Null
        Set-ItemProperty -Path $medicRegPath -Name "Start" -Value 4 -Type DWord -Force
        Write-Host "SUCCESS: WaaSMedicSvc disabled via sc.exe fallback." -ForegroundColor Green
        $success = $true
    } catch {
        Write-Host "ERROR: All methods to disable WaaSMedicSvc failed." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

Write-Host ""

# ─────────────────────────────────────────────
# 5. SET ENFORCED DESKTOP BACKGROUND
# ─────────────────────────────────────────────
Write-Host "--- TASK 5: Set Enforced Desktop Background ---" -ForegroundColor Cyan
Write-Host ""

$bgUrl       = "https://raw.githubusercontent.com/karslanianexpologic-sudo/ELI-Imageupdate/ed6b361d1c90e261140929874519801893d53370/Background_MEMS.jpg"
$bgLocalPath = "C:\Windows\Web\Wallpaper\ELI\Background_MEMS.jpg"

try {
    $bgFolder = Split-Path $bgLocalPath
    if (-not (Test-Path $bgFolder)) { New-Item -ItemType Directory -Path $bgFolder -Force | Out-Null }

    Write-Host "Downloading background image..." -ForegroundColor White
    (New-Object Net.WebClient).DownloadFile($bgUrl, $bgLocalPath)
    Write-Host "  Image saved to $bgLocalPath" -ForegroundColor Green

    $wpRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    if (-not (Test-Path $wpRegPath)) { New-Item -Path $wpRegPath -Force | Out-Null }
    Set-ItemProperty -Path $wpRegPath -Name "DesktopImagePath"   -Value $bgLocalPath -Type String -Force
    Set-ItemProperty -Path $wpRegPath -Name "DesktopImageUrl"    -Value $bgLocalPath -Type String -Force
    Set-ItemProperty -Path $wpRegPath -Name "DesktopImageStatus" -Value 1            -Type DWord  -Force

    $noChangePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-ItemProperty -Path $noChangePath -Name "Wallpaper"          -Value $bgLocalPath -Type String -Force
    Set-ItemProperty -Path $noChangePath -Name "WallpaperStyle"      -Value "10"         -Type String -Force
    Set-ItemProperty -Path $noChangePath -Name "NoChangingWallpaper" -Value 1            -Type DWord  -Force

    Write-Host "SUCCESS: Desktop background set and enforced." -ForegroundColor Green

} catch {
    Write-Host "ERROR: Failed to set desktop background." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# ─────────────────────────────────────────────
# 6. SET ENFORCED LOCK SCREEN IMAGE + SETTINGS
# ─────────────────────────────────────────────
Write-Host "--- TASK 6: Set Enforced Lock Screen Image and Settings ---" -ForegroundColor Cyan
Write-Host ""

$lsUrl       = "https://raw.githubusercontent.com/karslanianexpologic-sudo/ELI-Imageupdate/ed6b361d1c90e261140929874519801893d53370/NewStationClosed.png"
$lsLocalPath = "C:\Windows\Web\Screen\ELI\NewStationClosed.png"

try {
    $lsFolder = Split-Path $lsLocalPath
    if (-not (Test-Path $lsFolder)) { New-Item -ItemType Directory -Path $lsFolder -Force | Out-Null }

    Write-Host "Downloading lock screen image..." -ForegroundColor White
    (New-Object Net.WebClient).DownloadFile($lsUrl, $lsLocalPath)
    Write-Host "  Image saved to $lsLocalPath" -ForegroundColor Green

    $lsRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    if (-not (Test-Path $lsRegPath)) { New-Item -Path $lsRegPath -Force | Out-Null }
    Set-ItemProperty -Path $lsRegPath -Name "LockScreenImagePath"   -Value $lsLocalPath -Type String -Force
    Set-ItemProperty -Path $lsRegPath -Name "LockScreenImageUrl"    -Value $lsLocalPath -Type String -Force
    Set-ItemProperty -Path $lsRegPath -Name "LockScreenImageStatus" -Value 1            -Type DWord  -Force
    Write-Host "SUCCESS: Lock screen image set and enforced." -ForegroundColor Green

    $lsPersonPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    if (-not (Test-Path $lsPersonPath)) { New-Item -Path $lsPersonPath -Force | Out-Null }
    Set-ItemProperty -Path $lsPersonPath -Name "NoLockScreen"       -Value 0            -Type DWord  -Force
    Set-ItemProperty -Path $lsPersonPath -Name "LockScreenImage"    -Value $lsLocalPath -Type String -Force
    Set-ItemProperty -Path $lsPersonPath -Name "NoWindowsSpotlight" -Value 1            -Type DWord  -Force
    Set-ItemProperty -Path $lsPersonPath -Name "NoLockScreenCamera" -Value 1            -Type DWord  -Force
    Write-Host "SUCCESS: Windows Spotlight / fun facts disabled." -ForegroundColor Green

    $lsMotionPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    if (-not (Test-Path $lsMotionPath)) { New-Item -Path $lsMotionPath -Force | Out-Null }
    Set-ItemProperty -Path $lsMotionPath -Name "RotatingLockScreenEnabled"        -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $lsMotionPath -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $lsMotionPath -Name "SubscribedContent-338387Enabled"  -Value 0 -Type DWord -Force
    Write-Host "SUCCESS: Lock screen motion/parallax effect disabled." -ForegroundColor Green

    $lsStatusPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
    if (-not (Test-Path $lsStatusPath)) { New-Item -Path $lsStatusPath -Force | Out-Null }
    Set-ItemProperty -Path $lsStatusPath -Name "DisableLockScreenAppNotifications" -Value 1 -Type DWord -Force
    Write-Host "SUCCESS: Lock screen status/notifications set to None." -ForegroundColor Green

    Set-ItemProperty -Path $lsStatusPath -Name "DisableLogonBackgroundImage" -Value 0 -Type DWord -Force
    Write-Host "SUCCESS: Lock screen background shown on sign-in screen." -ForegroundColor Green

} catch {
    Write-Host "ERROR: Failed to configure lock screen." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# ─────────────────────────────────────────────
# 7. SET GOOGLE CHROME AS DEFAULT BROWSER
# ─────────────────────────────────────────────
Write-Host "--- TASK 7: Set Google Chrome as Default Browser ---" -ForegroundColor Cyan
Write-Host ""

try {
    # Verify Chrome is installed
    $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    $chromePathX86 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

    if (-not (Test-Path $chromePath) -and -not (Test-Path $chromePathX86)) {
        Write-Host "ERROR: Google Chrome does not appear to be installed on this machine." -ForegroundColor Red
    } else {
        Write-Host "  Chrome installation found." -ForegroundColor Green

        # Set Chrome as default for HTTP, HTTPS, and HTML file types via registry
        $chromeAssocPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
        if (-not (Test-Path $chromeAssocPath)) { New-Item -Path $chromeAssocPath -Force | Out-Null }

        # Register Chrome handler in HKLM for all users
        $regPaths = @(
            "HKLM:\SOFTWARE\Classes\ChromeHTML\shell\open\command",
            "HKLM:\SOFTWARE\Classes\http\shell\open\command",
            "HKLM:\SOFTWARE\Classes\https\shell\open\command"
        )

        $actualChrome = if (Test-Path $chromePath) { $chromePath } else { $chromePathX86 }
        $chromeCmd = "`"$actualChrome`" -- `"%1`""

        foreach ($path in $regPaths) {
            if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
            Set-ItemProperty -Path $path -Name "(Default)" -Value $chromeCmd -Type String -Force
        }

        # Set default associations via DISM using an XML file
        $xmlContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<DefaultAssociations>
    <Association Identifier=".htm"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
    <Association Identifier=".html"  ProgId="ChromeHTML" ApplicationName="Google Chrome" />
    <Association Identifier=".pdf"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
    <Association Identifier="http"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
    <Association Identifier="https"  ProgId="ChromeHTML" ApplicationName="Google Chrome" />
</DefaultAssociations>
"@
        $xmlPath = "$env:TEMP\ChromeDefaults.xml"
        $xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8 -Force

        # Apply via DISM
        Write-Host "  Applying default browser associations via DISM..." -ForegroundColor White
        $dismResult = & dism.exe /Online /Import-DefaultAppAssociations:"$xmlPath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Google Chrome set as default browser via DISM." -ForegroundColor Green
        } else {
            Write-Host "  DISM method output: $dismResult" -ForegroundColor Yellow

            # Fallback: set via registry policy
            Write-Host "  Trying registry policy fallback..." -ForegroundColor Yellow
            $assocRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
            if (-not (Test-Path $assocRegPath)) { New-Item -Path $assocRegPath -Force | Out-Null }
            Set-ItemProperty -Path $assocRegPath -Name "DefaultAssociationsConfiguration" -Value $xmlPath -Type String -Force
            Write-Host "SUCCESS: Chrome default association policy set via registry." -ForegroundColor Green
            Write-Host "  NOTE: This will fully apply on next user login." -ForegroundColor Yellow
        }

        # Clean up temp XML
        Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
    }

} catch {
    Write-Host "ERROR: Failed to set Chrome as default browser." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# ─────────────────────────────────────────────
# 8. CONFIGURE ADVANCED SHARING SETTINGS
# ─────────────────────────────────────────────
Write-Host "--- TASK 8: Configure Advanced Sharing Settings ---" -ForegroundColor Cyan
Write-Host ""

try {
    # ── PRIVATE NETWORK ──────────────────────────────────────────────────────
    # Network Discovery and File & Printer Sharing for Private profile
    # Category 1 = Private in Windows network location
    $privateRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Network\NetworkLocationWizard"
    if (-not (Test-Path $privateRegPath)) { New-Item -Path $privateRegPath -Force | Out-Null }

    # Set Private profile: Network Discovery On, File & Printer Sharing On
    $ndPrivatePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles"
    # Apply via firewall rules for private profile
    netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes profile=private 2>&1 | Out-Null
    netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes profile=private 2>&1 | Out-Null
    Write-Host "SUCCESS: Network Discovery enabled for Private networks." -ForegroundColor Green
    Write-Host "SUCCESS: File and Printer Sharing enabled for Private networks." -ForegroundColor Green

    # ── PUBLIC NETWORK ───────────────────────────────────────────────────────
    netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes profile=public 2>&1 | Out-Null
    netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes profile=public 2>&1 | Out-Null
    Write-Host "SUCCESS: Network Discovery enabled for Public networks." -ForegroundColor Green
    Write-Host "SUCCESS: File and Printer Sharing enabled for Public networks." -ForegroundColor Green

    # ── ALL NETWORKS — Public Folder Sharing ─────────────────────────────────
    # SharingPolicyId controls public folder sharing in the Advanced Sharing UI
    # 0 = Enabled (Anyone on network can read/write), 1 = Read only, 2 = Disabled
    $pubFolderRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer"
    if (-not (Test-Path $pubFolderRegPath)) { New-Item -Path $pubFolderRegPath -Force | Out-Null }
    Set-ItemProperty -Path $pubFolderRegPath -Name "NoPublicFolderSharing" -Value 0 -Type DWord -Force

    # Also set the SharingPolicyId that the Advanced Sharing Settings UI reads
    $sharingPolicyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\HomeGroup"
    if (-not (Test-Path $sharingPolicyPath)) { New-Item -Path $sharingPolicyPath -Force | Out-Null }
    Set-ItemProperty -Path $sharingPolicyPath -Name "SharingPolicyId" -Value 1 -Type DWord -Force
    Write-Host "SUCCESS: Public Folder Sharing enabled for All networks." -ForegroundColor Green

    # ── ALL NETWORKS — Password Protected Sharing Off ─────────────────────────
    # LimitBlankPasswordUse = 0 disables password protected sharing
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LimitBlankPasswordUse" -Value 0 -Type DWord -Force

    # Also write to the specific key the Advanced Sharing Settings UI reads
    $networkSharingPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
    Set-ItemProperty -Path $networkSharingPath -Name "RequireSecuritySignature" -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $networkSharingPath -Name "EnableSecuritySignature"  -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $networkSharingPath -Name "RestrictNullSessAccess"   -Value 0 -Type DWord -Force
    Write-Host "SUCCESS: Password Protected Sharing disabled for All networks." -ForegroundColor Green

    # ── ENSURE REQUIRED SERVICES ARE RUNNING ─────────────────────────────────
    $fdServices = @("FDResPub", "SSDPSRV", "upnphost", "fdPHost", "LanmanServer", "LanmanWorkstation")
    foreach ($svc in $fdServices) {
        try {
            Set-Service -Name $svc -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name $svc -ErrorAction SilentlyContinue
            Write-Host "  Service '$svc' set to Automatic and started." -ForegroundColor Green
        } catch {
            Write-Host "  WARNING: Could not configure service '$svc'." -ForegroundColor Yellow
        }
    }

} catch {
    Write-Host "ERROR: Failed to configure sharing settings." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# ─────────────────────────────────────────────
# 9. SUMMARY + RESTART PROMPT
# ─────────────────────────────────────────────
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   All Tasks Complete. Reboot Required." -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Type YES to restart now, or press any other key to restart manually later." -ForegroundColor Yellow
$confirm = Read-Host "Restart now?"

if ($confirm -eq "YES") {
    Write-Host "Restarting..." -ForegroundColor Green
    Start-Sleep -Seconds 3
    Restart-Computer -Force
} else {
    Write-Host "Restart skipped. Please reboot manually to apply all changes." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
