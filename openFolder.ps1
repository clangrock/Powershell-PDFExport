function Get-SHDOpenFolderDialog {
    [cmdletbinding()]
    param (
        [string]$Title = "Please Select A folder"
    )
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

    $FolderBrowser.Description = $Title

    $result = $FolderBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
    if ($result -eq [Windows.Forms.DialogResult]::OK){
        $folderPath = $FolderBrowser.SelectedPath
        $FolderBrowser.dispose()
    }
    else {
        $FolderBrowser.dispose()
        exit
    }
    
    return $folderPath
}

#$Result = Get-SHDOpenFolderDialog -Title "Select the root folder for the Table of Contents"
#$Result  

