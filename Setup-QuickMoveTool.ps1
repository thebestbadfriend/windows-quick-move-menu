#region constants
$installDirectory = "C:\Program Files (x86)\QuickMoveTool"
$setupScriptLocation = Split-Path $MyInvocation.MyCommand.Source

# Quick Move menus
$quickMoveMenusMUIVerb = "Quick Move"

# $fileQuickMoveMenu = "HKCU:\SOFTWARE\Classes\*\shell\Quick Move Menu"
# $fileQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "*\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$fileQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove"
$fileQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

# $directoryQuickMoveMenu = "HKCU:\SOFTWARE\Classes\Directory\shell\Quick Move Menu"
# $directoryQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "Directory\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$directoryQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove"
$directoryQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

$directoryBackgroundQuickMoveMenu = "HKCU:\SOFTWARE\Classes\Directory\background\shell\Quick Move Menu"
$directoryBackgroundQuickMoveMenuProperties = @{ExtendedSubCommandsKey = "Directory\\background\\shellex\\ContextMenuHandlers\\Menu_QuickMove"; MUIVerb = $quickMoveMenusMUIVerb}

$directoryBackgroundQuickMoveMenuShellexBase = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove"
$directoryBackgroundQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

$addToQuickMoveBaseKey = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove\shell\AddToQuickMove"
$addToQuickMoveBaseProperties = @{Default = "Add This Folder To QuickMove"}
$addToQuickMoveCommandKey = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove\shell\AddToQuickMove\command"
$addToQuickMoveCommandProperties = @{Default = "powershell.exe & '$($installDirectory.Replace("\","\\"))\\AddTo-QuickMove.ps1' %V"}

$removeFromQuickMoveBaseKey = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove\shell\RemoveFromQuickMove"
$removeFromQuickMoveBaseProperties = @{Default = "Remove This Folder From QuickMove"}
$removeFromQuickMoveCommandKey = "HKCU:\SOFTWARE\Classes\Directory\background\shellex\ContextMenuHandlers\Menu_QuickMove\shell\RemoveFromQuickMove\command"
$removeFromQuickMoveCommandProperties = @{Default = "powershell.exe & '$($installDirectory.Replace("\","\\"))\\RemoveFrom-QuickMove.ps1' %V"}
#endregion

#region menu setup function definitions
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

#endregion

#region other function definitions
function PutFilesWhereTheyGo {
  "$setupScriptLocation"

  Set-Location -LiteralPath $setupScriptLocation

  if(-not (Test-Path -LiteralPath $installDirectory)) {
    New-Item -ItemType Directory $installDirectory
  }
  
  foreach($item in Get-ChildItem -Exclude *Setup*,*test*) {
    Copy-Item $item $installDirectory -Recurse -Force
  }
}
#endregion

#region main flow
PutFilesWhereTheyGo

# file quick move menu
# CreateRegistryItem -keyPath $fileQuickMoveMenu -properties $fileQuickMoveMenuProperties
CreateRegistryItem -keyPath $fileQuickMoveMenuShellexBase
CreateRegistryItem -keyPath $fileQuickMoveMenuShellexShell

# directory quick move menu
# CreateRegistryItem -keyPath $directoryQuickMoveMenu -properties $directoryQuickMoveMenuProperties
CreateRegistryItem -keyPath $directoryQuickMoveMenuShellexBase
CreateRegistryItem -keyPath $directoryQuickMoveMenuShellexShell

# directory background quick move menu
CreateRegistryItem -keyPath $directoryBackgroundQuickMoveMenu -properties $directoryBackgroundQuickMoveMenuProperties
CreateRegistryItem -keyPath $directoryBackgroundQuickMoveMenuShellexBase
CreateRegistryItem -keyPath $directoryBackgroundQuickMoveMenuShellexShell

# directory background quick move menu options
CreateRegistryItem -keyPath $addToQuickMoveBaseKey -properties $addToQuickMoveBaseProperties
CreateRegistryItem -keyPath $addToQuickMoveCommandKey -properties $addToQuickMoveCommandProperties
CreateRegistryItem -keyPath $removeFromQuickMoveBaseKey -properties $removeFromQuickMoveBaseProperties
CreateRegistryItem -keyPath $removeFromQuickMoveCommandKey -properties $removeFromQuickMoveCommandProperties

#endregion