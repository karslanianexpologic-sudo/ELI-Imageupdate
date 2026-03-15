#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Dell 3390 2-in-1 Windows 11 Master Configuration Script
.DESCRIPTION
    1. Renames PC to BIOS Asset Tag
    2. Sets screen timeout to Never (AC and Battery)
    3. Disables Windows Update service (wuauserv)
    4. Disables Windows Update Medic Service (WaaSMedicSvc)
    5. Prompts for restart
#>

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   Dell 3390 Configuration Script" -ForegroundColor Cyan
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
# 2. SET SCREEN TIMEOUT TO NEVER
# ─────────────────────────────────────────────
Write-Host "--- TASK 2: Set Screen Timeout to Never ---" -ForegroundColor Cyan
Write-Host ""

try {
    powercfg /change monitor-timeout-ac 0
    powercfg /change monitor-timeout-dc 0
    Write-Host "SUCCESS: Screen timeout set to Never for AC and Battery." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to set screen timeout." -ForegroundColor Red
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
# 5. SUMMARY + RESTART PROMPT
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
