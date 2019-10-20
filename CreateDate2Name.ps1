"Added time created to file name"

$TargetDir="C:\Users\JunYing\Desktop\dwhelper"

Get-ChildItem $TargetDir | ForEach-Object {
    $OldName=$_.Name
    $OldNameSplit=$OldName -split "_"
    $OldTime=$OldNameSplit[0]

    $NewTime=$_.LastWriteTime.ToString("yyMMddHH")
    
    if ($NewTime -ne $OldTime) {
        $OldPath=$_.FullName
        $NewName=$_.LastWriteTime.ToString("yyMMddHH") + "_" + $_.Name
        "Rename " + $OldName + " to " + $NewName
        Rename-Item -Path $OldPath -NewName $NewName
    }
}