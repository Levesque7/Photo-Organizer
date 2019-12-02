#Settings
$Env:PhotoOrgLogDir = "$HOME\Photo Organizer Logs"
$Env:PhotoOrgRecycleBin = "$Env:PhotoOrgLogDir\Recycle_Bin"
$Env:PhotoOrgRnamedLog = "$Env:PhotoOrgLogDir\Renamed-Files.log"
$Env:PhotoOrgSessionLog = "$Env:PhotoOrgLogDir\Session.log"
$Env:PhotoOrgFileTypesLog = "$Env:PhotoOrgLogDir\File-Types.log"
$Env:PhotoOrgSpecialFilePrefixes = "SS,ZZ"

#Tools
Function Get-SpecialFilePrefixes {
    $specialFilePrefixes = $Env:PhotoOrgSpecialFilePrefixes.Split(",")
    $specialFilePrefixes
}
Function Update-FileTypes {
    Param([parameter(Mandatory)][string]$Destination)
    $types = Get-Content $Env:PhotoOrgFileTypesLog
    $files = Get-ChildItem -Path $Destination -File -Recurse
    foreach ($f in $files) {
        $ext = $f.Name.Split(".")[-1]
        if ($types -notcontains $ext) {
            Add-Content $Env:PhotoOrgFileTypesLog -Value $ext
            $types = Get-Content $Env:PhotoOrgFileTypesLog
        }
    }
}

Function Get-FileMetaData {
    Param([parameter(Mandatory)][string]$Path)

    $metaData = New-Object psobject
    

    $shell = New-Object -COMObject Shell.Application
    $folder = Split-Path $Path
    $file = Split-Path $Path -Leaf
    $shellfolder = $shell.Namespace($folder)
    $shellfile = $shellfolder.ParseName($file)

    for ($i = 0; $i -le 266; $i++) {
        $propertyName = $shellfolder.GetDetailsOf($shellfile.items, $i)
        $propertyValue = $shellfolder.GetDetailsOf($shellfile, $i)
        $property = @{$propertyName = $propertyValue }

        if ($propertyValue) { $metaData | Add-Member $property -Force }
        
    }

    $metaData
}

Function Get-FileID {
    Param([parameter(Mandatory)][string]$Path)

    $filehash = get-filehash $Path
    $ID = $filehash.hash.substring(54,10)
    $ID
}

Function Format-DateForRename {
    Param([parameter(Mandatory)][string]$Date)

    $split = ($Date.Split("/: ") -replace '[\W]','')
    $date = ("{0} {1} {2}" -f ($split[0..2] -join "/"),($split[3,4] -join ":"),$split[5])
    $date = Get-Date -Date $date -Format "yyyy_MM_dd_HHmm"
    $date
}

Function Rename-SeesawFile {
    Param(
        [parameter(Mandatory)][string]$FilePath
    )

    $dir = $FilePath | Split-Path
    $fileName = $FilePath | Split-Path -Leaf

    if (!(Test-Path -Path $FilePath)) { Throw "File $filename does not exist in $dir"}

    $split = $FileName.Split(" _-.")
    $ID = Get-FileID $FilePath
    $newFileName = ("SS_{0}_{1}_{2}_{3}{4}_{5}.jpg" -f $split[3],$split[2],$split[1],$split[4],$split[5],$ID,$split[7])
    $newFileName
}

Function Remove-DuplicateFiles {
    Param([parameter(Mandatory)][string]$Dir)

    $filesBefore =  Get-ChildItem $Dir -File -Recurse | Select-Object FullName
    $initialCount = $filesBefore.Length
    Get-ChildItem *.* -path $Dir -Recurse | Sort-Object -Property CreationTime | Get-FileHash | Group-Object -Property Hash | Where-Object { $_.count -gt 1 } | ForEach-Object { $_.group | Select-Object -skip 1 } | Move-Item -Destination $Env:PhotoOrgRecycleBin -Force
    $filesAfter =  Get-ChildItem $Dir -File -Recurse | Select-Object FullName
    $finalCount = $filesAfter.Length
    $duplicates = $initialCount - $finalCount
    $duplicates
}

Function Get-MonthName {
    Param([parameter(Mandatory)][string]$Month)

    $monthName = switch ($month)
    {
        "01" { "January" }
        "02" { "February" }
        "03" { "March" }
        "04" { "April" }
        "05" { "May" }
        "06" { "June" }
        "07" { "July" }
        "08" { "August" }
        "09" { "September" }
        "10" { "October" }
        "11" { "November" }
        "12" { "December" }
    }

    $monthName
}

Function Get-NewFileLocation {
    Param(
        [parameter(Mandatory)][string]$FilePath,
        [parameter(Mandatory)][string]$Destination
    )

    $fileName = Split-Path $FilePath -Leaf
    $split = $fileName.Split("_.")
    $specialFiles = Get-SpecialFilePrefixes
    if ($specialFiles -contains $split[0]) {
        $split = $split[1..5]
    }
    $monthName = Get-MonthName -Month $split[1]
    $newLocation = ("{0}\{1}\{2}_{3}" -f $Destination,$split[0],$split[1],$monthName)
    $newLocation
}

Function Get-TotalSize {
    Param([parameter(Mandatory)][string]$Destination)
    $files = Get-ChildItem -Path $Destination -File -Recurse
    $size = [Math]::Round((($files | Measure-Object -Property Length -Sum).Sum / 1GB),3)
    $size
}

Function Initialize-PhotoOrganizer {
    Param([parameter(Mandatory)][string]$Destination)

    $dirs = @($Destination,$Env:PhotoOrgLogDir,$Env:PhotoOrgRecycleBin)
    foreach ($d in $dirs) {
        if (!(Test-Path -Path $d)) { 
            New-Item -Path $d-ItemType Directory 
        }
    }
}
    
Function New-RenameLogEntry {
    Param(
        [parameter(Mandatory)][string]$OldFilePath,
        [parameter(Mandatory)][string]$NewFilePath
    )

    $oldFileName = Split-Path -Path $OldFilePath -Leaf
    $newFileName = Split-Path -Path $NewFilePath -Leaf

    Add-Content -Path $Env:PhotoOrgRnamedLog -Value "$OldFileName>$NewFileName"
}

Function New-DuplicateLogEntry {
    Param([parameter(Mandatory)][string]$PathOfDuplicate)
    Add-Content -Path $Env:PhotoOrgLogDir\$duplicateLog -Value $PathOfDuplicate
}

Function New-SessionLogEntry {
    Param(
        [parameter(Mandatory)][datetime]$StartDate,
        [parameter(Mandatory)][int]$Duplicates,
        [parameter(Mandatory)][int]$TotalFiles,
        [parameter(Mandatory)][string]$Destination    
    )
    $size = Get-TotalSize -Destination $Destination
    $endDate = Get-Date 
    $Date =  Get-Date -Date $StartDate -Format "yyyy/MM/dd-HHmm"
    $timespan = $endDate - $StartDate
    Add-Content -Path $Env:PhotoOrgSessionLog -Value "$Date, $totalFiles files processed in $($timespan.Days):$($timespan.Minutes):$($timespan.Seconds), $Duplicates duplicates removed, destination now contains $size GB of files"
}

Function Remove-EmptyFolders  {
    Param([parameter(Mandatory)][string]$Dir)
    do { 
        $a = Get-ChildItem $Dir -recurse | Where-Object {($_.PSIsContainer -eq $True) -and ($_.GetFiles().Count -eq 0) -and ($null -eq (Get-ChildItem -Path $_.FullName -Recurse))} | Remove-Item 
        $a = Get-ChildItem $Dir -recurse | Where-Object {($_.PSIsContainer -eq $True) -and ($_.GetFiles().Count -eq 0) -and ($null -eq (Get-ChildItem -Path $_.FullName -Recurse))}
    }
    until ($null -eq $a)  
}


# Main Functions
Function Rename-MediaFile {
    Param([parameter(Mandatory)][string]$Path)

    $metaData = Get-FileMetaData $Path
    $ID = Get-FileID $Path
    $ext = $metaData.'File extension'

    if ($metaData.'Date taken') { 
        $date = Format-DateForRename -Date $metaData.'Date Taken'
        $newFileName = ( "{0}_{1}{2}" -f $date,$ID,$ext )
    }
    elseif ($metaData.'Media Created') {
        $date = Format-DateForRename -Date $metaData.'Media Created'
        $newFileName = ( "{0}_{1}{2}" -f $date,$ID,$ext )
    }
    elseif ($metaData.Filename -like "Seesaw*") {
        $newFileName = Rename-SeesawFile -FilePath $metaData.Path
    }
    elseif ($metaData.'Date Created') {
        $date = Format-DateForRename -Date $metaData.'Date Created'
        $newFileName = ( "ZZ_{0}_{1}{2}" -f $date,$ID,$ext )
    }
    else {
        $creationDate = (Get-ChildItem -Path $Path).CreationTime
        $date = Get-Date -Date $creationDate -Format "yyyy_MM_dd_HHmm"
        $newFileName = ( "ZZ_{0}_{1}{2}" -f $date,$ID,$ext )
    }

    Rename-Item -Path $Path -NewName $newFileName
    $split = $newFileName.Split("_.")
    $specialFiles = Get-SpecialFilePrefixes
    if ($specialFiles -contains $split[0]) {
        New-RenameLogEntry -OldFilePath $Path -NewFilePath $newFileName
    }

    $newFileName
}

Function Move-MediaFile {
    Param(
        [parameter(Mandatory)][string]$FilePath,
        [parameter(Mandatory)][string]$Destination
    )

    if (Test-Path -Path $FilePath) {
        $newLocation = Get-NewFileLocation -FilePath $FilePath -Destination $Destination
        $exist? = Test-Path $newLocation
        if (!$exist?) { 
            New-Item -ItemType Directory -Path $newLocation -Force 
        }

        try {
            Move-Item -Path $FilePath -Destination $newLocation -ErrorAction Stop
        }
        catch [System.IO.IOException] {
            Move-Item -Path $FilePath -Destination $Env:PhotoOrgRecycleBin -Force
            $script:duplicates++
        }
    }
}

# Program Run
Function Start-PhotoOrganizer {
    Param(
        [parameter(Mandatory)][string]$OriginPath,
        [parameter(Mandatory)][string]$Destination
    )
    
    ## Initialize
    Write-Progress -Activity Initializing
    $startDate = Get-Date
    Initialize-PhotoOrganizer -Destination $Destination
    Write-Progress -Activity Initializing -CurrentOperation "Removing Duplicates..."
    $initialDups = Remove-DuplicateFiles -Dir $OriginPath
    $files = Get-ChildItem -Path $OriginPath -File -Recurse
    $totalFiles = ($files | Measure-Object).Count
    $destinationFilesBefore = (Get-ChildItem -Path $Destination -File -Recurse | Measure-Object).Count
    Write-Progress -Activity Initializing -Completed
    
    ## Renaming and moving Files
    $workingOn = 0  
    foreach ($f in $files) {
        $workingOn++
        $i = [Math]::Round(($workingOn/$totalFiles)*100)
        Write-Progress -Activity "Renaming and moving files" -Status "$i% Complete" -PercentComplete $i -CurrentOperation "Working on file $workingOn of $totalFiles"
        $fileName = Split-Path $f.FullName -Leaf
        $folder = Split-Path $f.FullName -Parent
        $split = $fileName.Split("_")
        $specialFiles = Get-SpecialFilePrefixes
        if ($specialFiles -notcontains $split[0]) {
            $newFileName = Rename-MediaFile -Path $f.FullName
        }
        else { 
            $newFileName = $fileName
        }
        Move-MediaFile -FilePath "$folder\$newFileName" -Destination $Destination
    }
    Write-Progress -Activity "Renaming and moving Files" -Completed

    ## Clean Up   
    Write-Progress -Activity "Finishing Up"
    $destinationFilesAfter = (Get-ChildItem -Path $Destination -File -Recurse | Measure-Object).Count
    $totalDuplicates = $initialDups + (($destinationFilesBefore + $totalFiles) - $destinationFilesAfter)
    Remove-EmptyFolders -Dir $OriginPath
    New-SessionLogEntry -StartDate $startDate -Duplicates $totalDuplicates -TotalFiles $totalFiles
    Update-FileTypes -Destination $Destination
    Write-Progress -Activity "Finishing Up" -Completed
}