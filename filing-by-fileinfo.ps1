# Ref : https://github.com/chrisdee/Scripts/blob/master/PowerShell/Working/files/SortAndMoveFilesToDateFolders.ps1
# Ref : https://powershellmagazine.com/2015/04/13/pstip-use-shell-application-to-display-extended-file-attributes/
# Ref : https://geekeefy.wordpress.com/2016/10/15/powershell-get-mp3mp4-files-metadata-and-how-to-use-it-to-make-you-life-easy/

# Get the files which should be moved, without folders
$sourcePath = 'C:\Users\XXX\OneDrive\Pictures\SourcePath'

# Target Filder where files should be moved to. The script will automatically create a folder for the year and month.
$targetPath = 'C:\Users\XXX\OneDrive\Pictures\TargetPath'

# Suppose only 1 folder with all files there, otherwise, use -Resurse
# $files = Get-ChildItem 'D:\Temp' -Recurse | where {!$_.PsIsContainer}
$files = Get-ChildItem $sourcePath -File | where {!$_.PsIsContainer}

# List Files which will be proceed
# $files

# Define for file extened info
$Shell = New-Object -ComObject shell.application

foreach ($file in $files)
{

    $Folder = $Shell.NameSpace($file.DirectoryName)
    $File = $Folder.ParseName($file.Name)

    # Find the available property
    # 208 = Video's "Media created"
    # 12 = Photo's "Date taken"
    # 3 = File's "Date modified"
    # 4 = File's "Date created"
    $Property = $Folder.GetDetailsOf($File,208)
    if (-not $Property) {
        $Property = $Folder.GetDetailsOf($File,12)
        if (-not $Property) {
            $Property = $Folder.GetDetailsOf($File,3)
            if (-not $Property) {
                $Property = $Folder.GetDetailsOf($File,4)
            }
        }
    }
    # Get date in the required format as a string
    $RawDate = ($Property -Replace "[^\w /:]")
    $DateTime = [DateTime]::Parse($RawDate)
    #$DateTaken = $DateTime.ToString("yyyyMMdd_HHmm")

    $year = $DateTime.ToString("yyyy")
    $month = $DateTime.ToString("MM")

    # Set Directory Path
    $Directory = $targetPath + "\" + $year + $month
    # Create directory if it doesn't exsist
    if (!(Test-Path $Directory))
        {
        New-Item $directory -type directory | Out-Null
        "Created Folder " + $Directory
        }

    "Filing " + $Directory + "\" + $file.Name
    # Move File to new location
    $file | Move-Item -Destination $Directory
    # Copy File to new location
    # $file | Copy-Item -Destination $Directory
}
