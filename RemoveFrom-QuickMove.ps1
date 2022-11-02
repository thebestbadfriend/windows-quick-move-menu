param(
  [Parameter(Mandatory)]
  [string]$folderPath
)

$fileMenuRegPath = 'HKLM:\Software\Classes\*\shell\Quick-Move Menu'
$directoryMenuRegPath = 'HKLM:\Software\Classes\Directory\shell\Quick Move Menu'
$folderName = Split-Path -Path $folderPath -Leaf
$regKey = "Move To $folderName"

$commandStoreBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

Add-Type -AssemblyName PresentationFramework

if(Test-Path -LiteralPath "$fileMenuRegPath\$regKey") {
  $commandSplit = (((Get-ItemProperty -LiteralPath "HKLM:\Software\Classes\*\shell\Move To Toolbox\command").'(default)') -split '''%1''').trim()
  $commandFolderPath = ($commandSplit[-1])#.replace('\','\\')
  if($commandFolderPath -eq "$folderPath") {
    Remove-Item -LiteralPath "$fileMenuRegPath\$regKey" -Recurse
    Remove-Item -LiteralPath "$directoryMenuRegPath\$regKey" -Recurse
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
      Remove-Item -LiteralPath "$fileMenuRegPath\$regKey" -Recurse
      Remove-Item -LiteralPath "$directoryMenuRegPath\$regKey" -Recurse
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
#   (after the change is made so that they are added that way to begin with)