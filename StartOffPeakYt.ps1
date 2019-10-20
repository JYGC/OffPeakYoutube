C:\Python34\python.exe C:\Users\JunYing\OneDrive\Programming\OffPeakYtFetcher\OffPeakYt.py

$ShortVideoDir = "C:\Users\JunYing\Desktop\dwhelper"
$LongVideoDir = "C:\Users\JunYing\Desktop\ToWatch"

$LongVideoLimit = [System.DateTime]"00:18:00"

$TargetDir = "C:\Users\JunYing\OffPeakYt\work"

Get-ChildItem $TargetDir | ForEach-Object {
    $FileName = $_.Name
    $LengthColumn = 27
    $objShell = New-Object -ComObject Shell.Application
    $objFolder = $objShell.Namespace($TargetDir)
    $objFile = $objFolder.ParseName($FileName)
    $LengthStr = $objFolder.GetDetailsOf($objFile, $LengthColumn)

    $OldPath = [io.path]::combine($TargetDir, $FileName)

    if ($LengthStr -ne "")
    {
        $Length = [System.DateTime]$LengthStr
    }
    else
    {
        $Length = [System.DateTime]"01:00:00"
    }

    if ($Length -le $LongVideoLimit)
    {
        $NewPath = [io.path]::combine($ShortVideoDir, $FileName)
    }
    else
    {
        $NewPath = [io.path]::combine($LongVideoDir, $FileName)
    }

    mv $OldPath $NewPath
}

& "C:\Users\JunYing\OneDrive\Programming\OffPeakYtFetcher\CreateDate2Name.ps1"
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")