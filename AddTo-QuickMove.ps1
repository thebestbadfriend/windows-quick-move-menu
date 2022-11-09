param(
  [Parameter(Mandatory)]
  [string]$folderPath
)

"addto-quickmove.ps1 tried to run" > C:\Toolbox\quickmovelog.txt

#region regkeys
$fileQuickMoveMenu = "HKCU:\SOFTWARE\Classes\*\shell\Quick Move Menu"

$fileQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove"
$fileQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

$directoryQuickMoveMenu = "HKCU:\SOFTWARE\Classes\Directory\shell\Quick Move Menu"

$directoryQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove"
$directoryQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove\shell"
#endregion

$folderName = Split-Path -Path $folderPath -Leaf

$quickMoveScript = "$($(Split-Path $MyInvocation.MyCommand.Source).Replace("\","\\"))\\QuickMove-FileOrFolder.ps1"

function CreateRegistryItem ([string] $keyPath, [hashtable] $properties = $null) {
  if(Test-Path -LiteralPath $keyPath){
    Remove-Item -LiteralPath $keyPath -Recurse
  }
  
  New-Item -Path $keyPath
  
  if ( $properties -ne $null) {
    foreach ($key in $properties.Keys){
      if ($key -eq "Default") {
        Set-ItemProperty -LiteralPath $keyPath -Name "(Default)"  -Value $($properties.$key)
      }
      else {
        Set-ItemProperty -LiteralPath $keyPath -Name $key  -Value $($properties.$key)
      }
    }
  }
}


if(Test-Path -LiteralPath "$fileQuickMoveMenuShellexShell\Move To $folderName") {
  Add-Type -AssemblyName PresentationFramework
  $regKeyExistsResponse = [System.Windows.MessageBox]::Show("A folder named $folderName already exists in the Quick Move menu. Do you want to replace it with this folder?", "Folder Name Already In Menu", "YesNo", "Exclamation")

  if($regKeyExistsResponse -eq "Yes"){
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}

    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}
  }
  elseif($regKeyExistsResponse -eq "No") {
    [System.Windows.MessageBox]::Show("No changes made.", "Operation Cancelled", "OK", "Information")
  }
}
else {
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}

    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}
}

# TODO
# make this put items in a fly-out menu rather than putting them directly in the right-click menu
