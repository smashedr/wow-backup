$ErrorActionPreference = "Stop"
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function ExitScript() {
    Read-Host -Prompt "Press <enter> or close this window to exit"
    exit
}

function GetWowDir() {
    $wow_dir_path = Join-Path -Path ${Env:ProgramFiles(x86)} -ChildPath "World of Warcraft"

    if (Test-Path $wow_dir_path) {
        $wow_dir = Get-Item $wow_dir_path
        return $wow_dir
    } else {
        $wow_folder = New-Object System.Windows.Forms.FolderBrowserDialog
        $wow_folder.Description = "Select your World of Warcraft installation directory."
        $wow_folder.rootfolder = "MyComputer"
        if ($wow_folder.ShowDialog() -eq "OK") {
            $wow_dir = Get-Item $wow_folder.SelectedPath
            return $wow_dir
        } else {
            return $null
        }
    }
}

function PerformBackup() {
    $wow_dir_path = Join-Path -Path ${Env:ProgramFiles(x86)} -ChildPath "World of Warcraft"

    if (Test-Path $wow_dir_path) {
        $wow_dir = Get-Item $wow_dir_path
        Write-Output "Found WoW installation directory: $wow_dir"
    } else {
        $wow_folder = New-Object System.Windows.Forms.FolderBrowserDialog
        $wow_folder.Description = "Select your World of Warcraft installation directory."
        $wow_folder.rootfolder = "MyComputer"
        if ($wow_folder.ShowDialog() -eq "OK") {
            $wow_folder_path += $wow_folder.SelectedPath
        } else {
            Write-Output "No WoW directory selected, please try again."
            ExitScript
        }

        $wow_dir = Get-Item $wow_folder_path
        Write-Output "User selected WoW installation directory: $wow_dir"
    }

    $user_dir = @(Get-ChildItem -Directory -Path $wow_dir -Filter "_*_" | `
        Out-GridView -Title 'Chose a single version of WoW to backup...' -PassThru)

    if (-Not $user_dir) {
        Write-Output "No WoW version found/selected, please try again."
        ExitScript
    }

    $backup_dir_path = Join-Path -Path $wow_dir -ChildPath $user_dir.Name

    Write-Output "User selected directory: $backup_dir_path"

    $backup_folder = New-Object System.Windows.Forms.FolderBrowserDialog
    $backup_folder.Description = "Select folder where backup file will be placed."
    $backup_folder.rootfolder = "MyComputer"
    if ($backup_folder.ShowDialog() -eq "OK") {
        $backup_folder_name += $backup_folder.SelectedPath
    } else {
        Write-Output "No backup folder selected, please try again."
        ExitScript
    }

    $backup_file_name = $user_dir.Name + $(get-date -f MM_dd_yyyy) + "-" + $(get-date -f HH_mm_ss) + ".zip"
    $backup_file_path = Join-Path -Path $backup_folder_name -ChildPath $backup_file_name

    $wtf_dir_path = Join-Path -Path $backup_dir_path -ChildPath "WTF"
    $interface_dir_path = Join-Path -Path $backup_dir_path -ChildPath "Interface"

    if (!(Test-Path $wtf_dir_path) -or !(Test-Path $interface_dir_path)) {
        Write-Output "Error, could not find WTF or Interface directory in: $backup_dir_path"
        ExitScript
    }

    Write-Output @(""
    "Will backup WoW with the following settings:"
    "--------------------------------------------------------------------------------"
    "Source WoW Directory:        $backup_dir_path"
    "WTF Folder:                  $wtf_dir_path"
    "Interface Folder:            $interface_dir_path"
    "Target Backup Directory:     $backup_folder_name"
    "Backup File Name:            $backup_file_name"
    "--------------------------------------------------------------------------------"
    "")

    Write-Output "Proceed with backup?"
    Read-Host -Prompt "Press <enter> to proceed"
    Write-Output "Compressing backup archive now, please wait..."

    try {
        $compress = @{
            LiteralPath = $wtf_dir_path, $interface_dir_path
            CompressionLevel = "Optimal"
            DestinationPath = $backup_file_path
        }
        Compress-Archive @compress
    } catch {
        Write-Output "Error creating backup file. Maybe a perms issue? $backup_file_path"
        ExitScript
    }

    Write-Output @(
    ""
    "Success! WoW Bakcup has completed."
    ""
    "Backup File:     $backup_file_path"
    ""
    )
    Read-Host -Prompt "Press <enter> or close this window to exit"
}

function PerformRestore() {
    Write-Output "Please select a backup archive to restore..."

    $backup_archive = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $backup_archive.AddExtension = $true
    $backup_archive.Filter = 'ZIP Files (*.zip)|*.zip|All Files|*.*'
    $backup_archive.Multiselect = $false
    $backup_archive.FilterIndex = 0
    $backup_archive.RestoreDirectory = $true
    $backup_archive.Title = 'Select backup file to restore'

    if ($backup_archive.ShowDialog() -eq 'OK') {
        $backup_archive_path = $backup_archive.FileName
        Write-Output "Using backup archive: $backup_archive_path"
    } else {
        Write-Output "No backup archive chosen for restore, please try again."
        ExitScript
    }

    $wow_dir = GetWowDir
    if (-Not $wow_dir) {
        Write-Output "No WoW installation directory selected, try again..."
        ExitScript
    }
    Write-Output "Using World of Warcraft directory: $wow_dir"

    $backup_file = Get-Item $backup_archive_path
    $split = $backup_file.Name -split "_"
    $wow_version = $split[1]
    $wow_version_path = Join-Path -Path $wow_dir -ChildPath "_${wow_version}_"

    if (!(Test-Path $wow_version_path)) {
        Write-Output "Error, could not determine WoW version from backup file name: $backup_file.Name"
        Write-Output "While verifying WoW version directory: $wow_version_path"
        Write-Output "This is probably due to backup archive being renamed or not created with this tool."
        Write-Output "Restore is currently not possible, please complete the restore manually..."
        ExitScript
    }

    Write-Output "Using World of Warcraft version directory: $wow_version_path"

    $wtf_dir_path = Join-Path -Path $wow_version_path -ChildPath "WTF"
    $interface_dir_path = Join-Path -Path $wow_version_path -ChildPath "Interface"

    if (!(Test-Path $wtf_dir_path) -or !(Test-Path $interface_dir_path)) {
        Write-Output "Error, could not find WTF or Interface directory in: $backup_dir_path"
        ExitScript
    }

    Write-Output "Using World of Warcraft WTF directory: $wtf_dir_path"
    Write-Output "Using World of Warcraft Interface directory: $interface_dir_path"

    Write-Output @(""
    "Will restore WoW with the following settings:"
    "--------------------------------------------------------------------------------"
    "Backup File Path:                 $backup_file"
    "WoW Directory:                    $wow_version_path"
    "WTF Folder (will replace):        $wtf_dir_path"
    "Interface Folder (will replace):  $interface_dir_path"
    "--------------------------------------------------------------------------------"
    ""
    "Please review the abote details! Note the folders being removed/restored."
    "")

    do {
        $action = Read-Host 'To proceed with the restore type "proceed"'
    } while ("proceed" -ne $action.Trim('"').ToLower())

    $temp_dir = New-Item -Force -Path $env:TEMP -Name "wow-restore" -ItemType "directory"

    if ($temp_dir.GetFiles() -or $temp_dir.GetDirectories()) {
        Write-Output "Removing and re-creating existing temp directory: $temp_dir"
        $temp_dir.Delete($True)
        $temp_dir = New-Item -Force -Path $env:TEMP -Name "wow-restore" -ItemType "directory"
    }

    try {
        Write-Output "Extracting backup archive to temp directory: $temp_dir"
        Expand-Archive -literalpath $backup_file -destinationpath $temp_dir.FullName
    } catch {
        Write-Output "Error extracting backup archive to: $temp_dir"
        $temp_dir.Delete($True)
        ExitScript
    }

    $wtf_backup_path = Join-Path -Path $wow_version_path -ChildPath "WTF"
    $interface_backup_path = Join-Path -Path $wow_version_path -ChildPath "Interface"
    if (!(Test-Path $wtf_backup_path) -or !(Test-Path $interface_backup_path)) {
        Write-Output "Error, could not find WTF or Interface directory in backup archive:"
        Get-ChildItem $temp_dir
        $temp_dir.Delete($True)
        ExitScript
    }

    try {
        Write-Output "Removing WTF directory: $wtf_dir_path"
        Remove-Item -Recurse -Force $wtf_dir_path
        Write-Output "Removing Interface directory: $interface_dir_path"
        Remove-Item -Recurse -Force $interface_dir_path
    } catch {
        Write-Output "Fatal Error removing WTF or Interface directory."
        Write-Output "Manual resolution required! Temp directory not deleted: $temp_dir"
        ExitScript
    }

    try {
        Write-Output "Restoring backup from:"
        Get-ChildItem $temp_dir
        Write-Output "`nTo destination: $wow_version_path"
        Move-Item -Path "$temp_dir\*" -Destination $wow_version_path
    } catch {
        Write-Output "Fatal Error copying backup files into live directory."
        Write-Output "Manual resolution required! Temp directory not deleted: $temp_dir"
        ExitScript
    }

    Write-Output "Cleaning up temp directory: $temp_dir"
    $temp_dir.Delete($True)

    Write-Output "`nSuccess! WoW Restore has completed.`n"
    Read-Host -Prompt "Press <enter> or close this window to exit"
}

Write-Host @"

Welcome to the wow-back and restore utility.

1.  Backup   -  Create a new backup file.
2.  Restore  -  Restore WoW from a backup file.

"@

do {
    $action = Read-Host "Please enter 1 or 2 to proceed"
} while (1..2 -NotContains $action)

if ($action -eq 1) {
    PerformBackup
} elseif ($action -eq 2) {
    PerformRestore
}
