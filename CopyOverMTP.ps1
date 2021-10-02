# Windows Powershell Script to move a set of files (based on a filter) from a folder
# on a MTP device (e.g. Android phone) to a folder on a computer, using the Windows Shell.
# By Daiyan Yingyu, 19 March 2018, https://blog.daiyanyingyu.uk/2018/03/20/powershell-mtp/
# based on the (non-working) script found here:
# https://www.pstips.net/access-file-system-against-mtp-connection.html
# as referenced here:
# https://powershell.org/forums/topic/powershell-mtp-connections/
# Further modifications have been added by Sajid Alamgir to mitigate some NullException Errors in original script.
# https://github.com/sajidalamgir/
# This Powershell script is provided 'as-is', without any express or implied warranty.
# In no event will the author be held liable for any damages arising from the use of this script.
#
# Again, please note that used 'as-is' this script will COPY files from you phone:
#
# If you want to move files instead, you can replace the CopyHere function call with "MoveHere" instead.
# But once again, I can't take any responsibility for the use, or misuse, of this script.</em>
#
[CmdletBinding()]
param(

        [Parameter(Mandatory=$true)][string]$phoneName,
        [Parameter(Mandatory=$true)][string]$sourceFolder,
        [Parameter(Mandatory=$true)][string]$targetFolder        
)

function Get-ShellProxy
{
    if( -not $global:ShellProxy)
    {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}
 
function Get-Phone
{
    param($phoneName)
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.Items() | where { $_.name -match $phoneName }
    return $phone
}
 
function Get-SubFolder
{
    param($phoneName,[string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $phone = Get-Phone -phoneName $phoneName
        foreach ($pathPart in $pathParts)
    {
        if ($pathPart)
        {
           $phone = $phone.GetFolder.items() | where { $_.Name -match $pathPart }
        }
    }
    return $phone
}

$filter='(.jpg)|(.mp4)$'
#Add any kind of Filters which are required here
$phoneFolderPath = $sourceFolder
$destinationFolderPath = $targetFolder
# Optionally add additional sub-folders to the destination path, such as one based on date
$folder = Get-SubFolder -phoneName $phoneName -path $phoneFolderPath
$items = @( $folder.GetFolder.items() | where { $_.Name -match $filter } )
if ($items)
{
    $totalItems = $items.count
    if ($totalItems -gt 0)
    {
        # If destination path doesn't exist, create it only if we have some items to move
        if (-not (test-path $destinationFolderPath) )
        {
            $created = new-item -itemtype directory -path $destinationFolderPath
        }
 
        Write-Verbose "Processing Path : $phoneName\$phoneFolderPath"
        Write-Verbose "Copying to : $destinationFolderPath"
 
        $shell = Get-ShellProxy
        $destinationFolder = $shell.Namespace($destinationFolderPath).self
        $count = 0;
        foreach ($item in $items)
        {
            $fileName = $item.Name
 
            ++$count
            $percent = [int](($count * 100) / $totalItems)
            Write-Progress -Activity "Processing Files in $phoneName\$phoneFolderPath" `
                -status "Processing File ${count} / ${totalItems} (${percent}%)" `
                -CurrentOperation $fileName `
                -PercentComplete $percent

            # Check the target file doesn't exist:
            $targetFilePath = join-path -path $destinationFolderPath -childPath $fileName
            if (test-path -path $targetFilePath)
            {
                write-error "Destination file exists - file not moved:`n`t$targetFilePath"
            }
            else
            {
                $destinationFolder.GetFolder.CopyHere($item)
                if (test-path -path $targetFilePath)
                {
                    # Optionally do something with the file, such as modify the name (e.g. removed phone-added prefix, etc.)
                }
                else
                {
                    write-error "Failed to move file to destination:`n`t$targetFilePath"
                }
            }
        }
    }
}