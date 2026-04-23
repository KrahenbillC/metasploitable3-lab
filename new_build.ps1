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

function Install-Box($osFull, $osShort) {
    $templatePath = ".\packer\templates\$osFull.json"
    if (-not (Test-Path $templatePath)) {
        throw "Template not found: $templatePath"
    }

    $boxVersion = Get-BoxVersion $templatePath
    $boxPath = "packer\builds\$($osFull)_virtualbox_$boxVersion.box"
    $vmName = Get-VMName $osShort

    Show-Header "Building $vmName"
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

    Write-Host ""
    Write-Host "[INFO] Checking whether Vagrant box is already added..." -ForegroundColor Yellow
    $boxList = & $vagrant box list | Out-String

    if ($boxList -match [regex]::Escape("rapid7/metasploitable3-$osShort")) {
        Write-Host "[OK] rapid7/metasploitable3-$osShort already exists in Vagrant." -ForegroundColor Green
    } else {
        Write-Host "[INFO] Adding box to Vagrant..." -ForegroundColor Yellow
        & $vagrant box add $boxPath --name "rapid7/metasploitable3-$osShort"
        if ($LASTEXITCODE -ne 0) {
            throw "Error adding metasploitable3-$osShort box to Vagrant."
        }
        Write-Host "[OK] rapid7/metasploitable3-$osShort box successfully added." -ForegroundColor Green
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
    Show-Header "Metasploitable3 Interactive Builder"
    Write-Host "1. Check environment"
    Write-Host "2. Build Windows 2008 box"
    Write-Host "3. Build Ubuntu 1404 box"
    Write-Host "4. Build both boxes"
    Write-Host "5. Start win2k8 VM"
    Write-Host "6. Start ub1404 VM"
    Write-Host "7. Export win2k8 to OVA"
    Write-Host "8. Export ub1404 to OVA"
    Write-Host "9. Cleanup build artifacts"
    Write-Host "10. Full reset"
    Write-Host "11. Exit"
    Write-Host ""
}

function Run-Interactive {
    do {
        try {
            Show-Menu
            $choice = Read-Host "Choose an option"

            switch ($choice) {
                "1"  { Show-EnvironmentInfo; Pause-Step }
                "2"  { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8"; Pause-Step }
                "3"  { Show-EnvironmentInfo; Install-Box -osFull "ubuntu_1404" -osShort "ub1404"; Pause-Step }
                "4"  { Show-EnvironmentInfo; Install-Box -osFull "windows_2008_r2" -osShort "win2k8"; Install-Box -osFull "ubuntu_1404" -osShort "ub1404"; Pause-Step }
                "5"  { Start-VM "win2k8"; Pause-Step }
                "6"  { Start-VM "ub1404"; Pause-Step }
                "7"  { Export-OVA "win2k8"; Pause-Step }
                "8"  { Export-OVA "ub1404"; Pause-Step }
                "9"  { Cleanup-BuildArtifacts; Pause-Step }
                "10" { Full-Reset; Pause-Step }
                "11" { Write-Host "Exiting." }
                default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Step }
            }
        }
        catch {
            Write-Host ""
            Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Pause-Step
        }
    }
    while ($choice -ne "11")
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