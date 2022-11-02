#region constants
$installDirectory = "C:\Program Files (x86)\QuickMoveTool"

# Quick Move menus
$quickMoveMenusMUIVerb = "Quick Move Menu"

$fileQuickMoveMenuBase = "HKLM:\SOFTWARE\Classes\*\shell\Quick Move"
$directoryQuickMoveMenuBase = "HKLM:\SOFTWARE\Classes\Directory\shell\Quick Move"

$directoryBackgroundQuickMoveMenuBase = "HKLM:\SOFTWARE\Classes\Directory\background\shell\Quick Move"
$directoryBackgroundQuickMoveSubCommands = "thebestbadfriend.AddThisFolderToQuickMove;thebestbadfriend.RemoveThisFolderFromQuickMove"


# Initial command store items
$commandStoreBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
$commandStoreAddThisFolderToQuickMoveMenu = "$commandStoreBase\thebestbadfriend.AddThisFolderToQuickMove"
$commandStoreRemoveThisFolderFromQuickMoveMenu = "$commandStoreBase\thebestbadfriend.RemoveThisFolderFromQuickMove"
#endregion

#region main flow
PutFilesWhereTheyGo
SetUpDirectoryBackground
SetUpFileQuickMove
SetUpFolderQuickMove
SetUpAddToQuickMoveCommand
SetUpRemoveFromQuickMoveCommand
#endregion

#region menu setup function definitions
function SetUpDirectoryBackground {
  if(-not (Test-Path -LiteralPath $directoryBackgroundQuickMoveMenuBase)){
    New-Item -Path $directoryBackgroundQuickMoveMenuBase
  }
  
  Set-ItemProperty -LiteralPath $directoryBackgroundQuickMoveMenuBase -Name "MUIVerb" -Value $quickMoveMenusMUIVerb
  Set-ItemProperty -LiteralPath $directoryBackgroundQuickMoveMenuBase -Name "SubCommands" -Value $directoryBackgroundQuickMoveSubCommands
} # Done, not tested

function SetUpFileQuickMove {
  if(-not (Test-Path -LiteralPath $fileQuickMoveMenuBase)){
    New-Item -Path $fileQuickMoveMenuBase
  }
  
  Set-ItemProperty -LiteralPath $fileQuickMoveMenuBase -Name "MUIVerb" -Value $quickMoveMenusMUIVerb
  Set-ItemProperty -LiteralPath $fileQuickMoveMenuBase -Name "SubCommands" -Value ""
}

function SetUpFolderQuickMove {
  if(-not (Test-Path -LiteralPath $directoryQuickMoveMenuBase)){
    New-Item -Path $directoryQuickMoveMenuBase
  }

  Set-ItemProperty -LiteralPath $directoryQuickMoveMenuBase -Name "MUIVerb" -Value $quickMoveMenusMUIVerb
  Set-ItemProperty -LiteralPath $directoryQuickMoveMenuBase -Name "SubCommands" -Value ""
} # Done, not tested
#endregion

#region command store setup function definitions
function SetUpAddToQuickMoveCommand {
  if(-not (Test-Path -LiteralPath $commandStoreAddThisFolderToQuickMoveMenu)){
    New-Item -Path $commandStoreAddThisFolderToQuickMoveMenu
  }

  if(-not (Test-Path -LiteralPath "$commandStoreAddThisFolderToQuickMoveMenu\command")){
    New-Item -Path "$commandStoreAddThisFolderToQuickMoveMenu\command"
  }
  
  Set-ItemProperty -LiteralPath $commandStoreAddThisFolderToQuickMoveMenu -Name "(Default)" -Value "Add This Folder To Quick-Move"
  Set-ItemProperty -LiteralPath "$commandStoreAddThisFolderToQuickMoveMenu\command" -Name "(Default)" -Value "$installDirectory\run-addto-quickmove.bat %V"
}

function SetUpRemoveFromQuickMoveCommand {
  if(-not (Test-Path -LiteralPath $commandStoreRemoveThisFolderFromQuickMoveMenu)){
    New-Item -Path $commandStoreRemoveThisFolderFromQuickMoveMenu
  }

  if(-not (Test-Path -LiteralPath "$commandStoreRemoveThisFolderFromQuickMoveMenu\command")){
    New-Item -Path "$commandStoreRemoveThisFolderFromQuickMoveMenu\command"
  }
  
  Set-ItemProperty -LiteralPath $commandStoreRemoveThisFolderFromQuickMoveMenu -Name "(Default)" -Value "Remove This Folder From Quick-Move"
  Set-ItemProperty -LiteralPath "$commandStoreRemoveThisFolderFromQuickMoveMenu\command" -Name "(Default)" -Value "$installDirectory\run-removefrom-quickmove.bat %V"
}

#endregion

#region other function definitions
function PutFilesWhereTheyGo {
  if(-not (Test-Path $installDirectory)) {
    New-Item -ItemType Directory $installDirectory
  }
  
  foreach($item in Get-ChildItem -Exclude *Setup*,*test*) {
    Copy-Item $item $installDirectory -Recurse -Force
  }
}
#endregion