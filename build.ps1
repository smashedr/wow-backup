$source_file_name = "wow-backup.ps1"
$output_file_name = "WoW-Backup.exe"
$application_name = "WoW Backup"
$author_name = "Shane"

$ErrorActionPreference = "Stop"

if (-Not $args[0]) {
    Write-Output "Specify the version number in format x.x.x.x for argument one."
    exit
}
$version = $args[0]

if (!(Test-Path ".\build\")) {
    $current_dir = Get-Location
    New-Item -Force -Path $current_dir.Path -Name "build" -ItemType "directory"
}
if (Test-Path ".\build\$output_file_name") {
    Remove-Item -Force ".\build\$output_file_name"
}

.\libs\ps2exe.ps1 `
    -inputFile ".\$source_file_name" `
    -outputFile ".\build\$output_file_name" `
    -iconFile ".\assets\icon.ico" `
    -title $application_name `
    -description $application_name `
    -product $application_name `
    -company $author_name `
    -copyright $author_name `
    -version $version
