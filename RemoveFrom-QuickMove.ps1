param(
  [Parameter(Mandatory)]
  [string]$folderPath
)

$fileQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\Menu_QuickMove\shell"
$directoryQuickMoveMenuShellexShell = "HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\Menu_QuickMove\shell"

$folderName = Split-Path -Path $folderPath -Leaf
$regKey = "Move To $folderName"

Add-Type -AssemblyName PresentationFramework

if(Test-Path -LiteralPath "$fileQuickMoveMenuShellexShell\$regKey") {
  $commandSplit = (((Get-ItemProperty -LiteralPath "$directoryQuickMoveMenuShellexShell\$regKey\command").'(Default)') -split '''%1''').trim()
  $commandFolderPath = ($commandSplit[-1])#.replace('\','\\')
  if($commandFolderPath -eq "$folderPath") {
    Remove-Item -LiteralPath "$fileQuickMoveMenuShellexShell\$regKey" -Recurse
    Remove-Item -LiteralPath "$directoryQuickMoveMenuShellexShell\$regKey" -Recurse
  }
  else {
    $otherFolderSameNameMessage = "This folder:
    $folderPath

does not exist in the quick move menu, but a different folder with the same name:
    $commandFolderPath

does exist. Would you like to remove that folder from the quick move menu?
    "


    $otherFolderSameNameResponse = [System.Windows.MessageBox]::Show($otherFolderSameNameMessage, "Different Folder", "YesNo", "Exclamation")

    if($otherFolderSameNameResponse -eq "Yes"){
      Remove-Item -LiteralPath "$fileQuickMoveMenuShellexShell\$regKey" -Recurse
      Remove-Item -LiteralPath "$directoryQuickMoveMenuShellexShell\$regKey" -Recurse
    }
    elseif($otherFolderSameNameResponse -eq "No") {
      [System.Windows.MessageBox]::Show("No changes made.", "Operation Cancelled", "OK", "Information")
    }
  }
}
else {
  [System.Windows.MessageBox]::Show("No folder with this name exists in the quick move menu.", "Operation Cancelled", "OK", "Information")
}

# TODO
# update this to remove the fly out menu items rather than the direct right click ones