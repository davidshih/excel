# PowerShell script to share folders on SharePoint
# Run this script after uploading folders to SharePoint

$siteUrl = Read-Host 'Enter SharePoint site URL'
$baseFolder = Read-Host 'Enter base folder path on SharePoint'

Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Share folder for Jane Smith
$folderPath = Join-Path $baseFolder 'Jane Smith'
$userEmail = Read-Host 'Enter email for Jane Smith'
Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'
Write-Host 'Shared folder for Jane Smith with Edit permissions'

# Share folder for Mike Wilson
$folderPath = Join-Path $baseFolder 'Mike Wilson'
$userEmail = Read-Host 'Enter email for Mike Wilson'
Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'
Write-Host 'Shared folder for Mike Wilson with Edit permissions'

# Share folder for Bob Johnson
$folderPath = Join-Path $baseFolder 'Bob Johnson'
$userEmail = Read-Host 'Enter email for Bob Johnson'
Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'
Write-Host 'Shared folder for Bob Johnson with Edit permissions'

# Share folder for Alice Chen
$folderPath = Join-Path $baseFolder 'Alice Chen'
$userEmail = Read-Host 'Enter email for Alice Chen'
Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'
Write-Host 'Shared folder for Alice Chen with Edit permissions'

# Share folder for John Doe
$folderPath = Join-Path $baseFolder 'John Doe'
$userEmail = Read-Host 'Enter email for John Doe'
Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'
Write-Host 'Shared folder for John Doe with Edit permissions'

