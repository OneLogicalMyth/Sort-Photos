Function Sort-Photos
{
    <#
	.SYNOPSIS
	    Copies pictures to 2 locations (1 is optional) and renames them.

	.DESCRIPTION
        The Sort-Photos cmdlet copys the pictures to 2 locations and renames them
        using the date the picture was taken. The format can be adjusted to suit your
        taste.
        This cmdlet also uses a file hash to ensure all copies are the same.
		
	.PARAMETER SourcePath
	    Specifies the path to the folder where the pictures are located. Default is current location (Get-Location).

	.PARAMETER DestinationPath
	    Specifies the path to the folder where the pictures will be copied. Default is a sub folder named 'Sorted'.

    .EXAMPLE    
        PS C:\> Sort-Photos
 
        Description: 
        Copy all the pictures in folder you are in to a sub folder named 'Sorted'.

    .EXAMPLE    
        PS C:\> Sort-Photos -SourcePath C:\Folder\Pics\ 
 
        Description: 
        Renames all the pictures in the given folder path.
		
	.NOTES
		Author: Liam Glanfield
        E-mail: liam.glanfield@outlook.com
    #>
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param (
        [Parameter(Mandatory=$false)][string[]]$SourcePath = (Get-Location),
        [Parameter(Mandatory=$false)][string]$DestinationPath = 'Sorted',
        [parameter(Mandatory=$false)][switch]$SubFolders = $false,
        [parameter(Mandatory=$false)][string]$FileNameFormat = (Get-CultureDateFormat),
        [parameter(Mandatory=$false)][string]$FolderPathFormat = 'yyyy\MM(MMMM)-yyyy',
        [parameter(Mandatory=$false)][ValidateSet('SHA1','SHA256','SHA384','SHA512','MACTripleDES','MD5','RIPEMD160')][string]$HashAlgo = 'MD5'
    )

    <#
    dynamicparam {
        $attributes = new-object System.Management.Automation.ParameterAttribute
        $attributes.ParameterSetName = "__AllParameterSets"
        $attributes.Mandatory = $true
        $attributeCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
        $attributeCollection.Add($attributes)
        $values = Get-ChildItem -Directory | Resolve-Path -Relative
        $ValidateSet = new-object System.Management.Automation.ValidateSetAttribute($values)
        $attributeCollection.Add($ValidateSet)

        $dynParam1 = new-object -Type System.Management.Automation.RuntimeDefinedParameter("Type", [string], $attributeCollection)
        $paramDictionary = new-object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
        $paramDictionary.Add("Type", $dynParam1)
        return $paramDictionary 
    }
    #>

    Begin 
    {
        $OrgEA = $ErrorActionPreference
        $ErrorActionPreference = 'Stop'

        #Load required assembly
        [reflection.assembly]::LoadFile("C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll") | Out-Null

        #Build progress hash table for splating later to Write-Progress
        $Progress = @{
            Activity = "Copying and Sorting Pictures"
            CurrentOperation = ''
            Status = 'Validating Paths and Building a List of Pictures to Process'
            PercentComplete = 0
            Id = 1
        }
        
        #Get a list of pictures to process, display warnings of invalid paths
        try
        {
            #Check each path and ensure its valid
            $SourcePath | Resolve-Path -ErrorAction Stop | Out-Null
            #Add a * on to the end of each so that the Include parameter works on Get-ChildItem
            [string[]]$StaredPaths = $SourcePath | %{ Join-Path $_ '*' }
            #Using the verified and stared paths get all pictures possible
            $Pictures = $StaredPaths | Get-ChildItem -Include *.jpeg, *.jpg, *.tiff, *.tif -Recurse:$SubFolders
            #Count pictures found and set counter for progress
            $TotalPictures = $Pictures.Count
            $i = 1
        }
        catch
        {
            Write-Warning $_
        }

        #if total pictures found is 0 return null with error
        If($TotalPictures -eq 0){
            Write-Error 'No pictures found to process'
            Return $null
        }else{
            #File size checks - do we have enough space!
            $CopySize = ($Pictures | Measure-Object -Property Length -Sum).Sum
            $DestSize = (Get-Item $DestinationPath).PSDrive.Free

            if(($DestSize - $CopySize) -lt 0)
            {
                Write-Error 'Not enough space on destination path'
                Return $null
            }

            #Build status for Write-Progress
            $Progress.Status = "Calculating Source File Hashes"
            $Pictures = $Pictures | foreach {
                #Get image data from the picture
                $Progress.CurrentOperation = "Calculating hash for '$($_.FullName)' - Picture $i/$TotalPictures - Total Size $([System.Math]::Round($CopySize / 1GB,2))GB"
                $Percentage = ($i/$TotalPictures)*100
                [int]$Progress.PercentComplete = $Percentage
                Write-Progress @Progress
                Get-FileHash -Path $_.FullName -Algorithm $HashAlgo -ErrorAction Stop
                $i++
            }
        }

        $OutputCSV = @()

    }
    
    Process 
    {

        #Process each picture
        $i = 1
        foreach($Picture IN $Pictures){

            $PicFile = Get-Item $Picture.Path

            #Build output object
            $OutHeaders = @(
                            'SourceFile',
                            'SourceFileHash',
                            'DestinationPath',
                            'DestinationFileHash',
                            'BackupPath',
                            'BackupFileHash',
                            'CopiedOk',
                            'Errors'
                            )
            $ResultOut = '' | Select-Object $OutHeaders

            #Build status for Write-Progress
            $Progress.Status = "Copying and Sorting Picture '$($Picture.Path)'"
            $Percentage = ($i/$TotalPictures)*100

            #Get image data from the picture
            $Progress.CurrentOperation = 'Getting date taken from picture'
            [int]$Progress.PercentComplete = $Percentage
            Write-Progress @Progress
            
            try
            {
                #Create object for reading data from picture
                $PicData = New-Object System.Drawing.Bitmap($Picture.Path)
                #Get the date taken from picture in bytes
                [byte[]]$DateBytes = $PicData.GetPropertyItem(36867).Value
                #Convert bytes to a string
                [string]$DateRawString = [System.Text.Encoding]::ASCII.GetString($DateBytes)
                #Convert the string value of the date to a datetime object
                [datetime]$DateTaken = [datetime]::ParseExact($DateRawString,"yyyy:MM:dd HH:mm:ss`0",$Null)
                $Exif = $true
            }
            catch [System.Exception], [System.IO.IOException]
            {
                #Could not get date taken from picture process next picture
                $DateTaken = $PicFile.CreationTime
                $Exif = $false
            }

            #Process file and copy to primary location
            $Progress.CurrentOperation = "Creating destination path structure"
            [int]$Progress.PercentComplete = $Percentage
            Write-Progress @Progress

            try
            {
                #If path equals 'Sorted' then create a new or use existing folder
                #named 'Sorted' at the pictures current location
                if($DestinationPath -eq 'Sorted'){
                    
                    #Build path based on picture location
                    $PictureFolder = $Picture.Path
                    $DestinationPath = Join-Path $PictureFolder 'Sorted'

                    #Test folder for existance if not then create it
                    if(-not (Test-Path -LiteralPath $DestinationPath)){
                        New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop | Out-Null
                    }
                    
                }

                # Build folder path required and create folders if requried
                # The regex replace is used to double up a backslash if only 1 is used
                $FolderPathDate = $DateTaken.ToString(($FolderPathFormat -replace '\w\\{1}\w','\\'))
                

                # If no exif data then add to its own folder
                if($Exif)
                {
                    $OutputPath = Join-Path $DestinationPath $FolderPathDate
                }else{
                    $OutputPath = Join-Path $DestinationPath 'NoExif'
                    $OutputPath = Join-Path $OutputPath $FolderPathDate
                    
                }
                $OutputFile = Join-Path $OutputPath ($DateTaken.ToString($FileNameFormat) + $PicFile.Extension)

                #Build destination path if needed
                if(-not (New-RecursivePath $OutputPath)){
                    throw "unable to create destination path recursively for '$OutputPath'"
                }

            }
            catch
            {
                #Display warning detailing error
                Write-Warning "$_ - $DestinationPath"
                #Increment counter by 1
                $i++
                #Could not create folder structure continue on to next file
                continue
            }

            #Process file and copy to primary location
            $Progress.CurrentOperation = "Getting an hash of original file"
            [int]$Progress.PercentComplete = $Percentage
            Write-Progress @Progress

            try
            {
                if((Test-Path $OutputFile)){
                    
                    $OutputFileHash = Get-FileHash -LiteralPath $OutputFile -ErrorAction Stop -Algorithm $HashAlgo
                    
                    if($OutputFileHash.Hash -eq $Picture.Hash){
                        $CopyResult = 'Already Existing'
                    }else{
                        
                        $x = 1
                        While((Test-Path $OutputFile)){
                            $TempFileName = Split-Path $OutputFile -Leaf

                            $NewFileName = "$($TempFileName.split('.')[0])_DUPE$($i).$($TempFileName.split('.')[1])"
                            $OutputFile = Join-Path (Split-Path $OutputFile -Parent) $NewFileName.Replace("_DUPE$($i - 1)",'')
                            $x++
                        }

                        Copy-Item $Picture.Path $OutputFile -ErrorAction Stop
                        $CopyResult = 'Copied'

                    }
                }else{
                    Copy-Item $Picture.Path $OutputFile -ErrorAction Stop
                    $CopyResult = 'Copied'
                }
            }
            catch
            {
                Write-Warning $_
                $CopyResult = 'Copy failed'
            }


            #Process file and copy to primary location
            $Progress.CurrentOperation = "Comparing source and destination hash"
            [int]$Progress.PercentComplete = $Percentage
            Write-Progress @Progress

            try
            {
                $OutputFileHash = Get-FileHash -LiteralPath $OutputFile -ErrorAction Stop -Algorithm $HashAlgo
                if($CopyResult -eq 'Already Existing'){
                    $null
                }elseif($OutputFileHash.Hash -eq $Picture.Hash){
                    $CopyResult = 'Copied with matching hash'
                }else{
                    $CopyResult = 'Copy failed hash did not match!'
                }

            }
            catch
            {
                $CopyResult = 'Copy failed hash did not generate!'
            }

            #Increment counter by 1
            $i++
        
        $Out = '' | Select-Object SourceFile, DestFile, SourceHash, DestHash, Result
        $Out.SourceFile = $Picture.Path
        $Out.DestFile = $OutputFile
        $Out.SourceHash = $Picture.Hash
        $Out.DestHash = $OutputFileHash.Hash
        $Out.Result = $CopyResult
        $Out

        $OutputCSV += $Out
        
        }#end foreach
    
    }

    End
    {
        #Set error action preference back
        $ErrorActionPreference = $OrgEA

        #Output errors to text file
        $Error | Foreach{ $_.ToString() | Out-File (Join-Path $DestinationPath 'Copy-Pictures-Errors.txt') -Encoding ascii }

        # Save Results
        $Today = (Get-Date).ToString($FileNameFormat)
        $OutputCSV | Export-CSV (Join-Path $DestinationPath "$($Today)_Results.csv") -NoTypeInformation

    }
}
