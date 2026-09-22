@echo off
setlocal
echo ========================================================
echo Building FootnoteConverter.oxt for LibreOffice
echo ========================================================

powershell -NoProfile -Command ^
  "Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem; " ^
  "$out = 'FootnoteConverter.oxt'; if (Test-Path $out) { Remove-Item -Force $out }; " ^
  "$fs = [System.IO.File]::Create((Resolve-Path .).Path + '\' + $out); " ^
  "$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create); " ^
  "@('footnoteconverter.py', 'Addons.xcu', 'Accelerators.xcu', 'ProtocolHandler.xcu', 'description.xml', 'README.md', 'LICENSE') | ForEach-Object { if (Test-Path $_) { [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Resolve-Path $_).Path, $_) | Out-Null } }; " ^
  "if (Test-Path 'META-INF/manifest.xml') { [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Resolve-Path 'META-INF/manifest.xml').Path, 'META-INF/manifest.xml') | Out-Null }; " ^
  "Get-ChildItem -Path 'icons' -Filter '*.png' | ForEach-Object { [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, ('icons/' + $_.Name)) | Out-Null }; " ^
  "$zip.Dispose(); $fs.Dispose(); Write-Output '[SUCCESS] FootnoteConverter.oxt created successfully!'"

echo.
pause
