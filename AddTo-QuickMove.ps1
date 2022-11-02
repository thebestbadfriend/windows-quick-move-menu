param(
  [Parameter(Mandatory)]
  [string]$folderPath
)

$fileQuickMoveMenuBase = "HKLM:\SOFTWARE\Classes\*\shell\Quick Move"
$directoryQuickMoveMenuBase = "HKLM:\SOFTWARE\Classes\Directory\shell\Quick Move"


$fileMenuBase = 'HKLM:\SOFTWARE\Classes\*\shell\Quick Move'
$directoryMenuBase = 'HKLM:\SOFTWARE\Classes\Directory\shell\Quick Move'
$folderName = Split-Path -Path $folderPath -Leaf

$commandStoreBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
$commandStoreQuickMoveEntry = "$commandStoreBase\thebestbadfriend.QuickMoveTo_changetherestofthistomakeituniqueperfolder" # need to generate this uniquely to each folder, are spaces allowed in command store entries?

$quickMoveScript = "C:\\Toolbox\\Coding\\PowerShell\\Scripts\\QuickMove\\QuickMove-FileOrFolder.ps1"


if(Test-Path -LiteralPath "$fileMenuBase\Move To $folderName") {
  Add-Type -AssemblyName PresentationFramework
  $regKeyExistsResponse = [System.Windows.MessageBox]::Show("A folder named $folderName already exists in the Quick Move menu. Do you want to replace it with this folder?", "Folder Name Already In Menu", "YesNo", "Exclamation")

  if($regKeyExistsResponse -eq "Yes"){
    Set-ItemProperty -LiteralPath "$fileMenuBase\Move To $folderName\command" -Name '(Default)' -Value "powershell.exe $quickMoveScript '%1' $folderPath"
    Set-ItemProperty -LiteralPath "$directoryMenuBase\Move To $folderName\command" -Name '(Default)' -Value "powershell.exe $quickMoveScript '%1' $folderPath"
  }
  elseif($regKeyExistsResponse -eq "No") {
    [System.Windows.MessageBox]::Show("No changes made.", "Operation Cancelled", "OK", "Information")
  }
}
else {
  New-Item -Path "$fileMenuBase\Move To $folderName"
  New-Item -Path "$fileMenuBase\Move To $folderName\command" -Value "powershell.exe $quickMoveScript '%1' $folderPath"
  New-Item -Path "$directoryMenuBase\Move To $folderName"
  New-Item -Path "$directoryMenuBase\Move To $folderName\command" -Value "powershell.exe $quickMoveScript '%1' $folderPath"
}

# TODO
# make this put items in a fly-out menu rather than putting them directly in the right-click menu
