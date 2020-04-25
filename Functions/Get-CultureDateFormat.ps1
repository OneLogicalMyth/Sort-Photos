Function Get-CultureDateFormat {
[cmdletbinding()]
Param([string]$Name=((Get-Culture).Name))

    # Grab culture date time formats
    $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Name)
    if($Culture.DisplayName -like 'Unknown Locale*'){
        throw 'Culture provided is not valid'
    }
    $CultureFormats = $Culture.DateTimeFormat

    # Get culture date and time, strip out everything but letters
    $ShortDate = $CultureFormats.ShortDatePattern -replace "\W+"
    $LongTime = $CultureFormats.LongTimePattern -replace "\W+"
    
    # Output string for default filename use
    Write-Output ($ShortDate + '_' + $LongTime)

}