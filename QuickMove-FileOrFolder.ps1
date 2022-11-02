param(
  [parameter(Mandatory)]
  [string]$source,

  [parameter(Mandatory)]
  [string]$dest
)

$fileName = Split-Path -Path $source -Leaf

$fileExists = Test-Path -Path "$dest\$fileName"

if($fileExists) {
  Add-Type -AssemblyName PresentationFramework
  $fileExistsResponse = [System.Windows.MessageBox]::Show("A file named $fileName already exists at $dest. Do you want to replace it with this file?", "File Already Exists", "YesNo", "Exclamation")

  if ($fileExistsResponse -eq "Yes") {
    Move-Item -Path $source -Destination $dest -Force
  }
  elseif ($fileExistsResponse -eq "No") {
    [System.Windows.MessageBox]::Show("No changes made.", "Operation Cancelled", "OK", "Information")
  }
}
else {
  Move-Item -Path $source -Destination $dest
}

# TODO
# check whether the destination directory still exists before attempting to move the file
# if it does, move the file normally
# if it does not, show a message to that effect and prompt for whether the user would like
#   to delete the nonexistent folder from the quick move menu