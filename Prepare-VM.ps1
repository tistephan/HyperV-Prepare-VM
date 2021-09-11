<#
    .SYNOPSIS
        Erstellt standardisierte VMs, basierend auf einer automatisierten Windows 10-Installation.

    .DESCRIPTION
        Liest alle .iso-Dateien (Windows-Installationsmedium) in einem Verzeichis ein, konfiguert eine Hyper-V-VM und automatisiert die Windows-Installation.

        Systemanforderungen:
            Microsoft SQL Server-Instanz - für Testzwecke die SQL Server 2019 Developer Edition
            Microsoft Hyper-V-Rolle aktiviert
            Microsoft Windows Assessment and Deployment Kit, Komponente Bereitstellungstools auf dem Hyper-V-Host installieren (https://docs.microsoft.com/de-de/windows-hardware/get-started/adk-install)
            PowerShell-Module PSSQLite und SQLServer vorhanden
            Programme und Konfigurationsdateien der Datei Tools.zip extrahiert

        Einrichtung:
            Verzeichnisse festlegen - Empfehlung: die Verzeichnisstruktur beibehalten!
            Windows 10-Images bereitstellen, z. B. durch Download mit Fido (https://github.com/pbatard/Fido)
            Virtuellen Switch mit der Bezeichnung INTERN erstellen (Typ: Internes Netzwerk)
            PowerShell-Modul PSSQLite installieren: https://www.powershellgallery.com/packages/PSSQLite/1.1.0
            PowerShell-Modul SQLServer installieren: https://docs.microsoft.com/de-de/sql/powershell/download-sql-server-ps-module?view=sql-server-ver15
            Initialize-HyperV-Host() einmalig ausführen
            Methode New-VirtualMachine(): Hardware-Konfiguration der VM festlegen
            Autounattended.xml: einen gültigen Windows 10-Product Key in Element <ProductKey> eintragen
            Tools.zip entpacken
            config.ini: sofern automatisierte USB-Stick-Emulation zwischen Hyper-V-Host und VM gewünscht (https://www.virtualhere.com)
            SetupComplete.cmd, Finish-VM.ps1: Pfad anpassen, wenn nicht in C:\_INSTALL installiert
            Das Administrator-Kennwort der VM lautet: Pa$$W0rd

        Hinweise zur Windows-Installation:
            Windows 10 Version 1511: Setup bleibt bei Auswahl der Edition stehen und muss manuell ausgewählt werden
            Windows 10 Version 1809: Setup nicht zu automatisieren

    .EXAMPLE
        Prepare-VM.ps1

    .LINK
        https://github.com/tistephan/HyperV-Prepare-VM
#>
function Initialize-HyperV-Host() {
    New-Item -ItemType Directory $iso_source_directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $iso_expanded_directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $iso_target_directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $vm_path -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $mounted_wimfile -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $artifacts_root_directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

    # Dateien der PowerShell-Module freischalten
    $ps_modules=("C:\Program Files\WindowsPowerShell\Modules\PSSQLite","C:\Program Files\WindowsPowerShell\Modules\SQLServer")
    foreach ($ps in $ps_modules) {
        $pssqlite_files = Get-ChildItem $ps -File -Recurse -Force
        foreach ($p in $pssqlite_files) {
            Unblock-File $p.FullName
        }
    }

    Install-Module -Name PSSQLite -Force
    Install-Module -Name SQLServer -Force
}
function New-CustomISOFile() {
    Write-host -NoNewline "   Installationsmedium erstellen..."

    # Verzeichnis für Dateien pro Windows-Version erstellen
    $artifacts_directory = "$($artifacts_root_directory)\$($w10_version_name)"

    New-Item -ItemType Directory $artifacts_directory -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType Directory $target_directoryname -ErrorAction SilentlyContinue | Out-Null

    $mount_isofile = Mount-DiskImage -ImagePath "$($_.FullName)" -Confirm:$false
    $mount_isofile_driveLetter = ($mount_isofile | Get-Volume).DriveLetter

    Start-Process -FilePath "robocopy.exe" -Wait -ArgumentList "$($mount_isofile_driveLetter):\  $($target_directoryname) /s /e /z /COPY:DT" # Mit Zeitstempel-Übernahme von Dateien
    Dismount-DiskImage -ImagePath "$($_.FullName)" -Confirm:$false | Out-Null

    # Autounattended.xml kopieren und anpassen: Bezeichnung (Computername) der zu erstellenden VM nach Windows-Version setzen
    Copy-Item "$($host_tools_path)\Autounattend.xml" "$($target_directoryname)"
    $autounattended_content = Get-Content -Path "$($target_directoryname)\Autounattend.xml"
    $autounattended_newcontent = $autounattended_content -replace "<ComputerName>W10<", "<ComputerName>$($w10_version_name)<"
    $autounattended_newcontent = $autounattended_newcontent -replace "<Label>Windows</Label>", "<Label>$($w10_version_name)_OS</Label>"
    
    $autounattended_newcontent | Set-Content -Path "$($target_directoryname)\Autounattend.xml"

    # Von einer Windows 10-ISO können mehrere Editionen ausgewählt werden. Um die Edition "Pro" (enthält BitLocker) auszuwählen, sind verschiende Indexe auszuwählen
    $pro_edition_index_no = get-windowsimage -imagepath "$($iso_expanded_directory)\$($w10_version_name)\sources\install.wim" | Where-Object { $_.ImageName -eq "Windows 10 Pro" } | Select-Object -ExpandProperty ImageIndex

    New-Item -ItemType Directory $mounted_wimfile -ErrorAction SilentlyContinue | Out-Null # Vorsorglich neu erstellen
    Set-ItemProperty -Path "$($iso_expanded_directory)\$($w10_version_name)\sources\install.wim" -Name IsReadOnly -Value $false # Attribut "Schreibgeschützt" von install.wim entfernen, ein Seiteneffekt von Robocopy per ISO-Datei

    Start-Process "dism.exe" -Wait -ArgumentList "/Cleanup-Mountpoints" | Out-Null
    Mount-WindowsImage -ImagePath "$($iso_expanded_directory)\$($w10_version_name)\sources\install.wim" -Index $pro_edition_index_no -Path $mounted_wimfile | Out-Null
    
    Start-Sleep -Seconds 20 # Vorsorglich warten, um Fehlermeldungen zu vermeiden
    New-Item -ItemType Directory "$($mounted_wimfile)\_INSTALL" | Out-Null
    
    Start-Process -FilePath "robocopy.exe" -Wait -ArgumentList """$($vm_tools_path)""  ""$($mounted_wimfile)\_INSTALL"" /s /e /z /COPY:DT" # Dateien für Szenarien auf VM kopieren - mit Zeitstempel-Übernahme von Dateien 
    New-Item -ItemType Directory "$($mounted_wimfile)\Windows\Setup\Scripts" | Out-Null
    Copy-Item "$($host_tools_path)\SetupComplete.cmd" "$($mounted_wimfile)\Windows\Setup\Scripts\SetupComplete.cmd"

    Start-Sleep -Seconds 30 # Vorsorglich warten, um Fehlermeldungen zu vermeiden
    # install.wim mit dism (!) wieder Dis-Mounten
    start-process "dism.exe" -wait -ArgumentList "/Unmount-Image /commit /mountdir:$($env:SystemDrive):\mount" | Out-Null

    Start-Sleep -Seconds 20 # Vorsorglich warten und danach aufräumen und Verzeichnis neu erstellen
    start-process "dism.exe" -wait -ArgumentList "/Cleanup-Mountpoints"  | Out-Null

    Remove-Item -Path $mounted_wimfile -Force -Recurse
    New-Item -ItemType Directory $mounted_wimfile -ErrorAction SilentlyContinue | Out-Null

    # "Press any key to..." abschalten, indem .bin/.efi-Dateien umbenannt werden
    Rename-Item -Path "$($iso_expanded_directory)\$($w10_version_name)\efi\microsoft\boot\cdboot.efi" -NewName "cdboot-prompt_NOT_NEEDED.efi"
    Rename-Item -Path "$($iso_expanded_directory)\$($w10_version_name)\efi\microsoft\boot\cdboot_noprompt.efi" -NewName "cdboot.efi"
    Rename-Item -Path "$($iso_expanded_directory)\$($w10_version_name)\efi\microsoft\boot\efisys.bin" -NewName "efisys_prompt_NOT_NEEDED.bin"
    Rename-Item -Path "$($iso_expanded_directory)\$($w10_version_name)\efi\microsoft\boot\efisys_noprompt.bin" -NewName "efisys.bin"

    # ISO-Datei zur automatisierten Installation von W10 erstellen (enthält autounattended.xml und angepasste install.wim)
    $vm_w10_iso_file = "$($iso_target_directory)\$($w10_version_name)_custom.iso" 
    
    Start-Process "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\Oscdimg.exe" -Wait -ArgumentList "-l$($w10_version_name) -m -o -u2 -udfver102 -bootdata:2#p0,e,b$($iso_expanded_directory)\$($w10_version_name)\boot\etfsboot.com#pEF,e,b$($iso_expanded_directory)\$($w10_version_name)\efi\microsoft\boot\efisys.bin $($iso_expanded_directory)\$($w10_version_name) $($vm_w10_iso_file)" | Out-Null

    # Entpackte Dateien der W10-ISO-Datei aus Speicherplatzgründen löschen
    Remove-Item "$($iso_expanded_directory)" -Recurse -Force
    
    Write-Host "fertig."
}

function New-VirtualMachine() {
    # Verwendung von TPM vorbereiten
    Write-Host -NoNewline "   VM erstellen..."
    $owner = Get-HgsGuardian UntrustedGuardian
    $kp = New-HgsKeyProtector -Owner $owner -AllowUntrustedRoot

    [uint64]$os_vhd_size = [bigint]20GB # Standardmäßige Festplattengröße für das Betriebssystem
    New-VM -Name "$($w10_version_name)" -Path "$($vm_path)" -MemoryStartupBytes 8192MB -NewVHDPath  "$($vm_path)\$($w10_version_name)\$($w10_version_name)_OS.vhdx" -NewVHDSizeBytes $($os_vhd_size) -Generation 2 -SwitchName INTERN | Out-Null

    Set-VMMemory "$($w10_version_name)" -DynamicMemoryEnabled $false # Keinen dynamischen RAM
    Set-VMKeyProtector -VMName "$($w10_version_name)" -KeyProtector $kp.RawData
    Enable-VMTPM -VMName "$($w10_version_name)"
    Set-VMProcessor -VMName "$($w10_version_name)" -Maximum 100 -Count 4 # Performance: am besten alle verfügbaren CPU-Kerne zuweisen

    # Festplattendatei zum Speichern von Process Monitor und Memory-Dumps
    New-VHD -Path "$($vm_path)\$($w10_version_name)\$($w10_version_name)_temp.vhdx" -SizeBytes 100GB | Mount-VHD -Passthru | Initialize-Disk -Passthru -PartitionStyle GPT | New-Partition -UseMaximumSize -DriveLetter Q | Format-Volume -FileSystem NTFS -Confirm:$($false) -Force -NewFileSystemLabel "Temp_$($w10_version_name)" | Out-Null
    Dismount-VHD -Path "$($vm_path)\$($w10_version_name)\$($w10_version_name)_temp.vhdx" # Nur in separatem Befehl möglich
    Add-VMHardDiskDrive -VMName "$($w10_version_name)" -Path "$($vm_path)\$($w10_version_name)\$($w10_version_name)_temp.vhdx"

    $vm_w10_iso_file = "$($iso_target_directory)\$($w10_version_name)_custom.iso" # Hier vorsorglich neu zusammensetzen; Variable wird aus der ISO-Erstellung nicht übernommen

    Add-VMDvdDrive -VMName "$($w10_version_name)" -ControllerNumber 0 -ControllerLocation 3 -Path $vm_w10_iso_file
    Set-VMFirmware -VMName "$($w10_version_name)" -BootOrder $(Get-VMDvdDrive -VMName "$($w10_version_name)"), $(Get-VMHardDiskDrive -VMName "$($w10_version_name)" | Where-Object { $_.ControllerLocation -eq 0 }) # Boot-Reihenfolge: DVD, Festplatte
    Set-VM -Name "$($w10_version_name)" -AutomaticCheckpointsEnabled $false
    Write-Host "fertig."

    # Windows 10 installieren
    Write-Host -NoNewline "   Windows 10 installieren..."
    Start-VM -VMName "$($w10_version_name)" # Einschalten für W10-Installation

    # Warten, bis die Installation fertig ist (= VM ist heruntergefahren durch SetupComplete.cmd)
    for (;;) {
        Start-Sleep -Seconds 5 # Vorsorglich warten
        If ((Get-VM "$($w10_version_name)").State -eq "Off") {
            break
        }
    }
    Write-Host "fertig."
}

function Initialize-VirtualMachine() {
    # DVD entfernen, da später bei den Szenarien das BitLocker-Setup aus Sicherheitsgründen nicht startet
    Get-VMDvdDrive -VMName "$($w10_version_name)" | Set-VMDvdDrive -Path $null

    # Snapshot von der Grundinstallation erstellen
    Write-Host -NoNewline "   Snapshot ""Grundinstallation"" erstellen..."
    Start-Sleep -Seconds 5 # Vorsorglich warten
    Checkpoint-VM -Name "$($w10_version_name)" -Snapshotname "Grundinstallation"
    Write-Host "fertig."
}

Clear-Host

$Global:ProgressPreference = 'SilentlyContinue' # Alle Fortschrittsbalken unterdrücken, z. B. bei Mount-WindowsImage
$host_tools_path = "C:\_INSTALL\Tools"
$vm_tools_path = "C:\_INSTALL\VM"
$vm_path = "D:\HyperV"
$iso_source_directory = "D:\W10_ISO"            # Original Windows 10-ISO-Dateien
$iso_target_directory = "D:\W10_ISO_customized" # Individualisierte Windows 10-ISO-Dateien
$iso_expanded_directory = "D:\W10_ISO_expanded" # Zwischenspeicher für individualisierte Windows 10-ISO-Dateien
$artifacts_root_directory="D:\W10_Artifacts"    # Speicherort der Artefakte
$mounted_wimfile = "$($env:SystemDrive)\mount" # install.wim anpassen: hier die Skript-Dateien zur Durchführung der Szenarien innerhalb der VM hinkopieren
$vm_w10_iso_file = "" # Hier global definieren, wird erst später benötigt
 
Get-ChildItem "$($iso_source_directory)" -Filter *.iso | Sort-Object -Property Name | # -Descending | 
foreach-object  {

    $target_directoryname = ("$($iso_expanded_directory)\$($_.Name)").Replace(".iso","")
    $w10_version_name = ($target_directoryname).Replace("$($iso_expanded_directory)\","")

    # Vorsorglich neue VM nur erstellen, wenn diese noch nicht existiert!
    if (!(Get-VM -VMName "$($w10_version_name)" -ErrorAction SilentlyContinue)) {
        Write-Host ""
        Write-Host "=> Einrichtung von VM ""$($w10_version_name)"" gestartet: $((Get-Date -Format "dd.MM.yyyy HH:mm:ss"))"
        Write-Host ""

        New-CustomISOFile
        New-VirtualMachine
        Initialize-VirtualMachine
        
        Write-Host ""
        Write-Host "=> Einrichtung von VM ""$($w10_version_name)"" beendet: $((Get-Date -Format "dd.MM.yyyy HH:mm:ss"))"
        Write-Host ""
    }
}
