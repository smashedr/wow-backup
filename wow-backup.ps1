$ErrorActionPreference = "Stop"
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function ExitScript() {
    Read-Host -Prompt "Press <enter> to exit or close this window"
    exit
}

$wow_dir_path = Join-Path -Path ${Env:ProgramFiles(x86)} -ChildPath "World of Warcraft"

if (Test-Path $wow_dir_path)
{
    $wow_dir = Get-Item $wow_dir_path
    Write-Output "Found WoW installation directory: $wow_dir"
} else {
    $wow_folder = New-Object System.Windows.Forms.FolderBrowserDialog
    $wow_folder.Description = "Select your World of Warcraft installation directory..."
    $wow_folder.rootfolder = "MyComputer"

    if($wow_folder.ShowDialog() -eq "OK")
    {
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
$backup_folder.Description = "Select backup Folder where archives will be stored..."
$backup_folder.rootfolder = "MyComputer"
if($backup_folder.ShowDialog() -eq "OK")
{
    $backup_folder_name += $backup_folder.SelectedPath
} else {
    Write-Output "No backup folder selected, please try again."
    ExitScript
}

$backup_file_name = $user_dir.Name + $(get-date -f MM_dd_yyyy) + "-" + $(get-date -f HH_mm_ss) + ".zip"
$backup_file_path = Join-Path -Path $backup_folder_name -ChildPath $backup_file_name

Write-Output @"

Will backup wow with the following settings:
--------------------------------------------------------------------------------
Source WoW Directory:        $backup_dir_path
Target Backup Directory:     $backup_folder_name
Backup File Name:            $backup_file_name
--------------------------------------------------------------------------------

"@

Read-Host -Prompt "Proceed with backup? Press <enter> to proceed"

$interface_dir_path = Join-Path -Path $backup_dir_path -ChildPath "Interface"
$wtf_dir_path = Join-Path -Path $backup_dir_path -ChildPath "WTF"

Write-Output "Compressing backup archive now, please wait..."

$compress = @{
    LiteralPath = $interface_dir_path, $wtf_dir_path
    CompressionLevel = "Optimal"
    DestinationPath = $backup_file_path
}
Compress-Archive @compress

Write-Output @"

Success! WoW Bakcup has completed.

Backup File:     $backup_file_path

"@
Read-Host -Prompt "Press <enter> to exit or close this window"
