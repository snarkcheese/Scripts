param(
    [parameter(Mandatory)]
    [string]$TargetUPN,

    [parameter(Mandatory)]
    [string]$SharepointAdministratorUPN,

    [parameter(Mandatory)]
    [string]$DownloadFolder
)
$DownloadPath = "\\Server\Backups\User Backups\$DownloadFolder"
If (!(Test-Path -LiteralPath $DownloadPath)) {
    $null = New-Item -ItemType Directory -Path $DownloadPath
}

#Connect to admin site
$SPAdminConnection = Connect-PnPOnline 'https://tenant-admin.sharepoint.com' -Interactive -ReturnConnection -ErrorAction Stop

#Set self as admin on site
$OneDriveSiteURL = (Get-PnPUserProfileProperty -Account $TargetUPN -Connection $SPAdminConnection).PersonalUrl -replace '\/$', ''
Set-PnPSite -Identity $OneDriveSiteURL -Owners $SharepointAdministratorUPN -Connection $SPAdminConnection

#Connect to OneDrive site
$OnedriveConnection = Connect-PnPOnline $OneDriveSiteURL -Interactive -ReturnConnection -ErrorAction Stop 
$Web = Get-PnPWeb -Connection $OnedriveConnection

#Get the "Documents" library where all OneDrive files are stored
$List = Get-PnPList -Identity "Documents" -Connection $OnedriveConnection
Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Gathering Files" -PercentComplete ((1/6)*100) -Status "Step 1 of 6"
#Get all Items from the Library - with progress bar
$global:counter = 0
$ListItems = Get-PnPListItem -List $List -PageSize 500 -Fields ID -Connection $OnedriveConnection -ScriptBlock {
    Param($items)
    #Exit if no items
    if ($null -eq $items -or $items.count -eq 0) {
        return
    }
    $global:counter += $items.Count
    Write-Progress -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting Items from OneDrive:" -Status "Processing Items $global:Counter to $($List.ItemCount)" -id 1 -ParentId 0
}
Write-Progress -Activity "Completed Retrieving Files and Folders from OneDrive!" -Completed -id 1 -ParentId 0

Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Creating folder structure" -PercentComplete ((2/6)*100) -Status "Step 2 of 6"
#Get all Subfolders of the library
$SubFolders = $ListItems | Where-Object { $_.FileSystemObjectType -eq "Folder" -and $_.FieldValues.FileLeafRef -ne "Forms" }
$SubFolders | ForEach-Object {
    #Ensure All Folders in the Local Path
    $LocalFolder = $DownloadPath + ($_.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)) -replace "/", "\"
    #Create Local Folder, if it doesn't exist
    If (!(Test-Path -LiteralPath $LocalFolder)) {
        $null = New-Item -ItemType Directory -Path $LocalFolder
    }
}

#Get all Files from the folder
$FilesColl = $ListItems | Where-Object { $_.FileSystemObjectType -eq "File" }

#Exit if no files
if ($null -eq $FilesColl -or $filescoll.count -eq 0) {
    return
}
Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Downloading files" -PercentComplete ((3/6)*100) -Status "Step 3 of 6"

if ($FilesColl.count -gt 5) {
    $pool = [runspacefactory]::CreateRunspacePool(1, 5)
    $pool.Open()
    $pooled = $true
}

$Downloads = [System.Collections.Generic.List[object]]::new()
#Iterate through each file and download
foreach ($file in $FilesColl) {
    $FileDownloadPath = ($DownloadPath + ($file.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)) -replace "/", "\").Replace($file.FieldValues.FileLeafRef, '')
    $job = {
        param(
            $File,
            $FileDownloadPath,
            $Connection
        )
        if (!(test-path -literalPath ($FileDownloadPath + $file.FieldValues.FileLeafRef) -PathType Leaf)) {
            Get-PnPFile -ServerRelativeUrl $File.FieldValues.FileRef -Path $FileDownloadPath -FileName $file.FieldValues.FileLeafRef -AsFile -Connection $Connection -Force
        }
    }
    $shell = [powershell]::Create()
    $null = $shell.AddScript($job)
    $null = $shell.AddArgument($File)
    $null = $shell.AddArgument($FileDownloadPath)
    $null = $shell.AddArgument($OnedriveConnection)
    if ($pooled) {
        $shell.RunspacePool = $pool
        $handle = $shell.BeginInvoke()
        $Downloads.add([pscustomobject]@{
            "Handle" = $handle
            "Shell" = $shell
        })
    }
    else {
        $shell.Invoke()
    }
}

#Wait for download
if ($pooled) {
    Do {
        $completed = ($downloads | Where-Object {$_.handle.iscompleted -eq $true}).Count
        Write-Progress -PercentComplete ($completed / ($FilesColl.Count) * 100) -Activity "Downloading Items from OneDrive:" -Status "Processing item $completed of $($List.ItemCount)" -id 2 -ParentId 0
        Start-sleep -Milliseconds 300
    } Until ($completed -eq $FilesColl.Count)
    Write-Progress -Activity "Completed Downloading Files from OneDrive!" -Completed -id 2 -ParentId 0

    #Close powershell instances
    foreach ($d in $downloads){
        $null = $d.shell.EndInvoke($d.handle)
        $d.Shell.Dispose()
    }

    #Close threads
    $pool.Close()
    $pool.Dispose()
}

#Close connections
Remove-Variable "SPAdminConnection"
Remove-Variable "OnedriveConnection"

Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Verifying download" -PercentComplete ((4/6)*100) -Status "Step 4 of 6"
#Verify Download
$FailedDownloads = foreach ($file in $FilesColl) {
    $FileDownloadPath = ($DownloadPath + ($file.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)) -replace "/", "\").Replace($file.FieldValues.FileLeafRef, '')
    $FilePath = $FileDownloadPath + $file.FieldValues.FileLeafRef
    if (!(test-path -literalPath $FilePath -PathType Leaf)) {
        $file.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)
    }
}
#Log failed downloads
Set-Content -LiteralPath "$DownloadPath\FailedDownloads.txt" -Value $FailedDownloads

Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Compressing download" -PercentComplete ((5/6)*100) -Status "Step 5 of 6"
#Zip
$7ZipArgs = @(
    "a", 
    "-bb", 
    "-v2g", 
    "`"$DownloadPath\Onedrive.7z`"", 
    "`"$DownloadPath\Documents`""
)
$ProcessParams = @{
    "FilePath" = "c:\Program Files\7-Zip\7z.exe"
    "ArgumentList" = $7ZipArgs
    "RedirectStandardOutput" = "$BackupPath\FileTransfer.txt"
    "RedirectStandardError" = "$BackupPath\FileTransferErrors.txt"
}
# Start zip
Start-Process @ProcessParams -NoNewWindow -Wait

start-sleep -Seconds 5
Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -CurrentOperation "Removing source files" -percentcomplete ((2/6)*100) -Status "Step 6 of 6"
#Remove source files
Remove-item -literalpath "$DownloadPath\Documents" -Recurse -Force
Write-Progress -Activity "Processing Onedrive for $TargetUPN" -id 0 -Completed
