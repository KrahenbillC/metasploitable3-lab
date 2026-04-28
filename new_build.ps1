$ErrorActionPreference = "Stop"

function Find-CommandPath($commandName, $fallbackPaths) {
    $cmd = Get-Command $commandName -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    foreach ($path in $fallbackPaths) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

$virtualBox = Find-CommandPath "VBoxManage" @(
    "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
)

$packer = Find-CommandPath "packer" @(
    "C:\Program Files\Packer\packer.exe",
    "C:\Users\$env:USERNAME\Documents\packer\packer.exe"
)

$vagrant = Find-CommandPath "vagrant" @(
    "C:\Program Files\Vagrant\bin\vagrant.exe",
    "C:\Program Files (x86)\Vagrant\bin\vagrant.exe"
)

function Show-Header($text) {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host $text -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
}

function Pause-Step {
    Write-Host ""
    Read-Host "Press Enter to continue"
}

function Require-Tool($toolPath, $friendlyName) {
    if (-not $toolPath) {
        throw "$friendlyName is not installed or not accessible."
    }
    Write-Host "[OK] $friendlyName found at: $toolPath" -ForegroundColor Green
}

function Confirm-Action($message) {
    $response = Read-Host "$message (y/n)"
    return ($response -eq "y")
}

function Remove-PathIfExists($path) {
    if (Test-Path $path) {
        Write-Host "[INFO] Removing $path" -ForegroundColor Yellow
        Remove-Item -Recurse -Force $path -ErrorAction SilentlyContinue
        Write-Host "[OK] Removed $path" -ForegroundColor Green
    }
}

function Ensure-VagrantPlugin($pluginName) {
    $plugins = & $vagrant plugin list
    if ($plugins -match "(?m)^$([regex]::Escape($pluginName))\s") {
        Write-Host "[OK] Vagrant plugin '$pluginName' found." -ForegroundColor Green
    } else {
        Write-Host "[INFO] Installing Vagrant plugin '$pluginName'..." -ForegroundColor Yellow
        & $vagrant plugin install $pluginName
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to install Vagrant plugin: $pluginName"
        }
        Write-Host "[OK] Vagrant plugin '$pluginName' installed." -ForegroundColor Green
    }
}

function Ensure-PackerPlugin($pluginRef) {
    $pluginName = ($pluginRef -split "/")[-1]
    $installed = & $packer plugins installed 2>$null | Out-String

    if ($installed -match [regex]::Escape($pluginName)) {
        Write-Host "[OK] Packer plugin '$pluginName' found." -ForegroundColor Green
    } else {
        Write-Host "[INFO] Installing Packer plugin '$pluginName'..." -ForegroundColor Yellow
        & $packer plugins install $pluginRef
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to install Packer plugin: $pluginRef"
        }
        Write-Host "[OK] Packer plugin '$pluginName' installed." -ForegroundColor Green
    }
}

function Get-BoxVersion($templatePath) {
    $content = Get-Content $templatePath -Raw
    if ($content -match '"box_version"\s*:\s*"([0-9]+\.[0-9]+\.[0-9]+)"') {
        return $matches[1]
    }
    throw "Could not find box_version in $templatePath"
}

function Get-VMName($osShort) {
    switch ($osShort) {
        "win2k8" { return "metasploitable3-win2k8" }
        "ub1404" { return "metasploitable3-ub1404" }
        default  { return "metasploitable3-$osShort" }
    }
}

function Get-VMFolderCandidates($vmName) {
    return @(
        "F:\$vmName",
        "C:\$vmName",
        "C:\metasploitable3-workspace\$vmName"
    )
}

function Cleanup-BuildArtifacts {
    Show-Header "Cleanup Build Artifacts"

    $paths = @(
        ".\output-virtualbox-iso",
        ".\.vagrant"
    )

    foreach ($path in $paths) {
        Remove-PathIfExists $path
    }

    Write-Host "[OK] Local build artifacts cleaned." -ForegroundColor Green
}

function Remove-VirtualBoxVM($vmName) {
    $vms = & $virtualBox list vms | Out-String
    if ($vms -notmatch [regex]::Escape($vmName)) {
        Write-Host "[INFO] VM not found in VirtualBox: $vmName" -ForegroundColor DarkYellow
        return
    }

    Write-Host "[INFO] Removing VirtualBox VM: $vmName" -ForegroundColor Yellow

    $info = & $virtualBox showvminfo $vmName --machinereadable 2>$null | Out-String
    if ($info -match 'VMState="running"') {
        Write-Host "[INFO] VM is running. Powering off first..." -ForegroundColor Yellow
        & $virtualBox controlvm $vmName poweroff | Out-Null
        Start-Sleep -Seconds 5
    }

    & $virtualBox unregistervm $vmName --delete 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[WARN] '--delete' failed. Retrying unregister only..." -ForegroundColor DarkYellow
        & $virtualBox unregistervm $vmName 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to unregister VirtualBox VM: $vmName"
        }
    }

    Write-Host "[OK] VirtualBox VM removed/unregistered: $vmName" -ForegroundColor Green
}

function Remove-VagrantBox($boxName) {
    $boxes = & $vagrant box list | Out-String
    if ($boxes -match [regex]::Escape($boxName)) {
        Write-Host "[INFO] Removing Vagrant box: $boxName" -ForegroundColor Yellow
        & $vagrant box remove $boxName --force
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to remove Vagrant box: $boxName"
        }
        Write-Host "[OK] Removed Vagrant box: $boxName" -ForegroundColor Green
    } else {
        Write-Host "[INFO] Vagrant box not found: $boxName" -ForegroundColor DarkYellow
    }
}

function Full-Reset {
    Show-Header "Full Reset"

    if (-not (Confirm-Action "This will delete build artifacts, VirtualBox VMs, and Vagrant boxes. Continue")) {
        Write-Host "[INFO] Full reset canceled." -ForegroundColor DarkYellow
        return
    }

    Cleanup-BuildArtifacts

    foreach ($vmName in @("metasploitable3-win2k8", "metasploitable3-ub1404")) {
        Remove-VirtualBoxVM $vmName
        foreach ($path in (Get-VMFolderCandidates $vmName)) {
            Remove-PathIfExists $path
        }
    }

    Remove-VagrantBox "rapid7/metasploitable3-win2k8"
    Remove-VagrantBox "rapid7/metasploitable3-ub1404"

    Write-Host "[OK] Full reset complete." -ForegroundColor Green
}


function Test-VBoxVMExists($vmName) {
    if (-not $virtualBox) { return $false }
    try {
        $vms = & $virtualBox list vms 2>$null | Out-String
        return ($vms -match [regex]::Escape($vmName))
    }
    catch {
        return $false
    }
}

function Show-BuildStage {
    param(
        [Parameter(Mandatory=$true)][string]$Stage,
        [Parameter(Mandatory=$true)][string]$Message
    )

    Write-Host ""
    Write-Host "--------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "[$Stage] $Message" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------" -ForegroundColor DarkCyan
}


function Show-ISODownloadPatienceBanner {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "              WINDOWS SERVER 2008 R2 ISO DOWNLOAD" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Packer may download the Windows Server 2008 R2 ISO from Microsoft." -ForegroundColor Yellow
    Write-Host "This can take 5-20+ minutes depending on internet speed." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "During the download, Packer may appear paused at lines like:" -ForegroundColor Yellow
    Write-Host "  virtualbox-iso: Retrieving ISO" -ForegroundColor Gray
    Write-Host "  virtualbox-iso: Trying https://download.microsoft.com/..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "This is normal. Do not close this PowerShell window." -ForegroundColor Yellow
    Write-Host "A status reminder will appear periodically while Packer is still running." -ForegroundColor Yellow
    Write-Host ""
}

function Show-StudentFinalizationBanner($vmName) {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "              STUDENT MANUAL FINALIZATION" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "The VM has been created/preserved in VirtualBox." -ForegroundColor Green
    Write-Host ""
    Write-Host "VM name:" -ForegroundColor Cyan
    Write-Host "  $vmName" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Complete these steps in VirtualBox:"
    Write-Host "  1. Observe that Windows Server 2008 R2 has completed setup"
    Write-Host "  2. Install VirtualBox Guest Additions manually if needed"
    Write-Host "  3. Disable UAC manually if needed"
    Write-Host "  4. Set the VM network adapter to Host-Only Adapter"
    Write-Host "  5. Reboot the VM after those changes"
    Write-Host ""
    Write-Host "When the VM is ready:"
    Write-Host "  - Return to this PowerShell window"
    Write-Host "  - Press ENTER to return to the main menu"
    Write-Host "  - Type 0 at the main menu to exit"
    Write-Host ""
}

function Invoke-PackerWithStudentRelease {
    param(
        [Parameter(Mandatory=$true)][string]$TemplatePath,
        [Parameter(Mandatory=$true)][string]$VmName
    )

    $arguments = @(
        "build",
        "-force",
        "-on-error=abort",
        "--only=virtualbox-iso",
        $TemplatePath
    )

    Show-ISODownloadPatienceBanner

    Write-Host "[*] Starting Packer build..." -ForegroundColor Yellow
    Write-Host "[*] Packer output will appear below." -ForegroundColor Yellow
    Write-Host "[*] If the screen pauses while the ISO downloads, keep waiting. This is normal." -ForegroundColor Yellow
    Write-Host "[*] If Packer waits on SSH/WinRM after the VM is visibly complete, finish the VM manually." -ForegroundColor Yellow
    Write-Host "[*] Then press ENTER in this window to release the menu without deleting the VM." -ForegroundColor Yellow
    Write-Host ""

    $proc = Start-Process -FilePath $packer -ArgumentList $arguments -NoNewWindow -PassThru
    $studentReleased = $false
    $exitCode = 999
    $buildStart = Get-Date
    $lastReminder = Get-Date

    while (-not $proc.HasExited) {
        Start-Sleep -Seconds 2

        $now = Get-Date
        if (($now - $lastReminder).TotalSeconds -ge 30) {
            $elapsed = New-TimeSpan -Start $buildStart -End $now
            Write-Host ""
            Write-Host ("[Still working] Packer is downloading/building. Elapsed: {0}h {1}m {2}s. Please wait." -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds) -ForegroundColor DarkYellow
            Write-Host "[Tip] If the VM is visibly complete but Packer is stuck waiting on SSH/WinRM, press ENTER here to preserve the VM and return to the menu." -ForegroundColor DarkYellow
            Write-Host ""
            $lastReminder = $now
        }

        try {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq "Enter") {
                    $studentReleased = $true
                    Write-Host ""
                    Write-Host "[INFO] Student manual release requested. Preserving VM and stopping Packer wait..." -ForegroundColor Yellow

                    try {
                        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                    }
                    catch {}

                    Start-Sleep -Seconds 2
                    break
                }
            }
        }
        catch {
            # Some hosts do not expose KeyAvailable. Continue waiting for Packer.
        }
    }

    if ($proc.HasExited) {
        $exitCode = $proc.ExitCode
    }
    elseif ($studentReleased) {
        $exitCode = 130
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        StudentReleased = $studentReleased
    }
}

function Install-Box($osFull, $osShort) {
    $templatePath = ".\packer\templates\$osFull.json"
    if (-not (Test-Path $templatePath)) {
        throw "Template not found: $templatePath"
    }

    $boxVersion = Get-BoxVersion $templatePath
    $boxPath = "packer\builds\$($osFull)_virtualbox_$boxVersion.box"
    $vmName = Get-VMName $osShort

    Show-Header "Building $vmName"

    if ($osShort -ne "win2k8") {
        Write-Host "[INFO] Non-Windows-2008 builds use the original Vagrant box workflow." -ForegroundColor Yellow
        Write-Host "Template: $templatePath"
        Write-Host "Expected box file: $boxPath"
        Write-Host ""

        if (Test-Path $boxPath) {
            Write-Host "[OK] Existing box found. Skipping Packer build." -ForegroundColor Green
        } else {
            Cleanup-BuildArtifacts
            Remove-VirtualBoxVM $vmName
            foreach ($path in (Get-VMFolderCandidates $vmName)) {
                Remove-PathIfExists $path
            }

            Write-Host "[INFO] Starting Packer build..." -ForegroundColor Yellow
            & $packer build -force --only=virtualbox-iso $templatePath
            if ($LASTEXITCODE -ne 0) {
                throw "Error building the Vagrant box using Packer."
            }
            Write-Host "[OK] Box successfully built by Packer." -ForegroundColor Green
        }

        return
    }

    Show-BuildStage -Stage "Stage 1 of 5" -Message "Checking build files and environment"
    Require-Tool $virtualBox "VirtualBox"
    Require-Tool $packer "Packer"
    Write-Host "Template: $templatePath"
    Write-Host "VM Name : $vmName"
    Write-Host ""

    Show-BuildStage -Stage "Stage 2 of 5" -Message "Preparing the VirtualBox build VM"
    Write-Host "This option preserves completed VMs and avoids deleting usable student builds." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "If an older VM with the same name exists, remove it only when you intentionally want a fresh rebuild." -ForegroundColor Yellow
    $existing = Test-VBoxVMExists $vmName
    if ($existing) {
        Write-Host "[INFO] Existing VM detected: $vmName" -ForegroundColor Yellow
        $rebuild = Read-Host "Type REBUILD to delete it and start fresh, or press ENTER to keep it"
        if ($rebuild -eq "REBUILD") {
            Remove-VirtualBoxVM $vmName
            foreach ($path in (Get-VMFolderCandidates $vmName)) {
                Remove-PathIfExists $path
            }
        } else {
            Write-Host "[OK] Existing VM preserved. Skipping destructive cleanup." -ForegroundColor Green
            Show-StudentFinalizationBanner $vmName
            Read-Host "Press ENTER to return to the main menu" | Out-Null
            return
        }
    } else {
        Cleanup-BuildArtifacts
    }

    Show-BuildStage -Stage "Stage 3 of 5" -Message "Installing Windows Server 2008 R2 and creating the VM"
    Write-Host "The VM will appear in VirtualBox during this stage." -ForegroundColor Yellow
    Write-Host "Do not close VirtualBox while the build is running." -ForegroundColor Yellow
    Write-Host ""

    $buildStart = Get-Date
    $result = Invoke-PackerWithStudentRelease -TemplatePath $templatePath -VmName $vmName
    $elapsed = New-TimeSpan -Start $buildStart -End (Get-Date)

    Show-BuildStage -Stage "Stage 4 of 5" -Message "Student manual completion"
    $vmDetected = Test-VBoxVMExists $vmName

    if ($vmDetected) {
        Show-StudentFinalizationBanner $vmName
        Read-Host "Press ENTER to return to the main menu" | Out-Null
    } else {
        Write-Host "[WARNING] The VM was not detected by name in VirtualBox." -ForegroundColor Yellow
        Write-Host "Check VirtualBox manually before attempting a rebuild." -ForegroundColor Yellow
        Read-Host "Press ENTER to return to the main menu" | Out-Null
    }

    Show-BuildStage -Stage "Stage 5 of 5" -Message "Build summary"
    Write-Host "Elapsed time: $($elapsed.Hours)h $($elapsed.Minutes)m $($elapsed.Seconds)s"
    Write-Host "Packer exit code: $($result.ExitCode)"
    Write-Host "Student manual release: $($result.StudentReleased)"
    Write-Host ""

    if ($vmDetected) {
        Write-Host "[SUCCESS] VM created/preserved in VirtualBox: $vmName" -ForegroundColor Green
        Write-Host "[INFO] No automatic delete/reset was performed." -ForegroundColor Cyan
    } else {
        Write-Host "[WARNING] Build finished, but the VM was not confirmed." -ForegroundColor Yellow
    }
}

function Start-VM($vmShortName) {
    Show-Header "Starting $vmShortName"
    & $vagrant up $vmShortName
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to start VM: $vmShortName"
    }
    Write-Host "[OK] $vmShortName started." -ForegroundColor Green
}

function Stop-VBoxVMIfRunning($vmName) {
    $info = & $virtualBox showvminfo $vmName --machinereadable 2>$null | Out-String
    if (-not $info) {
        throw "VirtualBox VM '$vmName' was not found."
    }

    if ($info -match 'VMState="running"') {
        Write-Host "[INFO] VM is running. Powering off..." -ForegroundColor Yellow
        & $virtualBox controlvm $vmName poweroff | Out-Null
        Start-Sleep -Seconds 5
    }
}

function Export-OVA($vmShortName) {
    $vmName = Get-VMName $vmShortName

    Show-Header "Exporting $vmName to OVA"

    $outputDir = ".\exports"
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmm"
    $ovaPath = Join-Path $outputDir "$vmName-$timestamp.ova"

    Write-Host "[INFO] Export path: $ovaPath" -ForegroundColor Yellow
    Stop-VBoxVMIfRunning $vmName

    & $virtualBox export $vmName --output $ovaPath
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to export $vmName to OVA."
    }

    Write-Host "[OK] Export complete: $ovaPath" -ForegroundColor Green
}

function Show-EnvironmentInfo {
    Show-Header "Environment Check"

    Require-Tool $virtualBox "VirtualBox"
    Require-Tool $packer "Packer"
    Require-Tool $vagrant "Vagrant"

    Write-Host ""
    Write-Host "VirtualBox version: $(& $virtualBox -v)"
    Write-Host "Packer version: $(& $packer version | Select-Object -First 1)"
    Write-Host "Vagrant version: $(& $vagrant -v)"

    Write-Host ""
    Ensure-VagrantPlugin "vagrant-reload"
    Ensure-PackerPlugin "github.com/hashicorp/virtualbox"

    Write-Host ""
    Write-Host "[OK] All requirements found." -ForegroundColor Green
}

function Show-Menu {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host " Metasploitable3 Windows 2008 Student Builder" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host " Build the VM, complete final steps in VirtualBox,"
    Write-Host " then return here and type 0 to exit."
    Write-Host ""
    Write-Host "[1] Check environment"
    Write-Host "[2] Build Windows 2008 VM"
    Write-Host "[3] Build Ubuntu 1404 box"
    Write-Host "[4] Build both"
    Write-Host "[5] Start win2k8 VM"
    Write-Host "[6] Start ub1404 VM"
    Write-Host "[7] Export win2k8 to OVA"
    Write-Host "[8] Export ub1404 to OVA"
    Write-Host "[9] Cleanup build artifacts only"
    Write-Host "[10] Full reset / delete VMs"
    Write-Host "[0] Exit"
    Write-Host ""
}

function Run-Interactive {
    $exitRequested = $false

    do {
        try {
            Show-Menu
            $choice = (Read-Host "Select an option").Trim()

            switch ($choice) {
                "1"  { Show-EnvironmentInfo; Pause-Step }
                "2"  { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8" }
                "3"  { Show-EnvironmentInfo; Install-Box -osFull "ubuntu_1404" -osShort "ub1404"; Pause-Step }
                "4"  { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8"; Install-Box -osFull "ubuntu_1404" -osShort "ub1404"; Pause-Step }
                "5"  { Start-VM "win2k8"; Pause-Step }
                "6"  { Start-VM "ub1404"; Pause-Step }
                "7"  { Export-OVA "win2k8"; Pause-Step }
                "8"  { Export-OVA "ub1404"; Pause-Step }
                "9"  { Cleanup-BuildArtifacts; Pause-Step }
                "10" { Full-Reset; Pause-Step }
                "0"  {
                    $exitRequested = $true
                    Write-Host ""
                    Write-Host "Exiting Metasploitable3 Student Builder." -ForegroundColor Cyan
                }
                default {
                    Write-Host "[ERROR] Invalid selection." -ForegroundColor Red
                    Pause-Step
                }
            }
        }
        catch {
            Write-Host ""
            Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Pause-Step
        }
    }
    while (-not $exitRequested)
}

if ($args.Length -gt 0) {
    try {
        switch ($args[0].ToLower()) {
            "windows2008"   { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8" }
            "ubuntu1404"    { Show-EnvironmentInfo; Install-Box -osFull "ubuntu_1404" -osShort "ub1404" }
            "both"          { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8"; Install-Box -osFull "ubuntu_1404" -osShort "ub1404" }
            "start-win2k8"  { Start-VM "win2k8" }
            "start-ub1404"  { Start-VM "ub1404" }
            "export-win2k8" { Export-OVA "win2k8" }
            "export-ub1404" { Export-OVA "ub1404" }
            "cleanup"       { Cleanup-BuildArtifacts }
            "reset"         { Full-Reset }
            default {
                Write-Host "Unknown argument."
                Write-Host "Valid options: windows2008, ubuntu1404, both, start-win2k8, start-ub1404, export-win2k8, export-ub1404, cleanup, reset"
            }
        }
    }
    catch {
        Write-Host ""
        Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Run-Interactive
}