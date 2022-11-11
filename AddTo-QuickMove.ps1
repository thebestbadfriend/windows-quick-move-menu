param(
  [Parameter(Mandatory)]
  [string]$folderPath
)

$fileQuickMoveMenu = "HKCU:\SOFTWARE\Classes\*\shell\Quick Move Menu"
$fileQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "*\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$directoryQuickMoveMenu = "HKCU:\SOFTWARE\Classes\Directory\shell\Quick Move Menu"
$directoryQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "Directory\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$fileQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove\shell"
$directoryQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

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

function CreateFileAndDirectoryMenuRegistryItems {
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$fileQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}

    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName" -properties @{"Default" = "Move To $folderName"}
    CreateRegistryItem -keyPath "$directoryQuickMoveMenuShellexShell\Move To $folderName\command" -properties @{"Default" = "powershell.exe & '$quickMoveScript' '%1' $folderPath"}
}

# main flow
if(-not (Test-Path -LiteralPath $fileQuickMoveMenu)) {
  New-Item -Path $fileQuickMoveMenu
  foreach ($key in $fileQuickMoveMenuProperties.Keys){
    if ($key -eq "Default") {
      Set-ItemProperty -LiteralPath $fileQuickMoveMenu -Name "(Default)"  -Value $($fileQuickMoveMenuProperties.$key)
    }
    else {
      Set-ItemProperty -LiteralPath $fileQuickMoveMenu -Name $key  -Value $($fileQuickMoveMenuProperties.$key)
    }
  }
}

if(-not (Test-Path -LiteralPath $directoryQuickMoveMenu)) {
  New-Item -Path $directoryQuickMoveMenu
  foreach ($key in $directoryQuickMoveMenuProperties.Keys){
    if ($key -eq "Default") {
      Set-ItemProperty -LiteralPath $directoryQuickMoveMenu -Name "(Default)"  -Value $($directoryQuickMoveMenuProperties.$key)
    }
    else {
      Set-ItemProperty -LiteralPath $directoryQuickMoveMenu -Name $key  -Value $($directoryQuickMoveMenuProperties.$key)
    }
  }
}

if(Test-Path -LiteralPath "$fileQuickMoveMenuShellexShell\Move To $folderName") {
  Add-Type -AssemblyName PresentationFramework
  $regKeyExistsResponse = [System.Windows.MessageBox]::Show("A folder named $folderName already exists in the Quick Move menu. Do you want to replace it with this folder?", "Folder Name Already In Menu", "YesNo", "Exclamation")

  if($regKeyExistsResponse -eq "Yes"){
    CreateFileAndDirectoryMenuRegistryItems
  }
  elseif($regKeyExistsResponse -eq "No") {
    [System.Windows.MessageBox]::Show("No changes made.", "Operation Cancelled", "OK", "Information")
  }
}
else {
  CreateFileAndDirectoryMenuRegistryItems
}