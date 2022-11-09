param(
  [Parameter(Mandatory)]
  [string]$folderPath
)
#region regkeys
$fileQuickMoveMenu = "HKCU:\SOFTWARE\Classes\*\shell\Quick Move Menu"
$fileQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "*\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$fileQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove"
$fileQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

$directoryQuickMoveMenu = "HKCU:\SOFTWARE\Classes\Directory\shell\Quick Move Menu"
$directoryQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "Directory\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$directoryQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove"
$directoryQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove\shell"
#endregion

$folderName = Split-Path -Path $folderPath -Leaf

$quickMoveScript = "$($(Split-Path $MyInvocation.MyCommand.Source).Replace("\","\\"))QuickMove-FileOrFolder.ps1"


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
