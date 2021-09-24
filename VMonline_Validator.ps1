param (
    [Parameter(HelpMessage='Path of the file to process')]
    $InputFile,
    [Parameter(HelpMessage='Detect missing fields in the file')]
    [bool]$MissingFields
)
$OldTitle = $host.ui.RawUI.WindowTitle
$Version = "0.1 Public"
$host.ui.RawUI.WindowTitle = "VMonline Validator - $Version"

write-host "VMonline validator, Created by: CouncilWise Pty Ltd`r`nVersion - $Version"
if ((Get-Host).UI.RawUI.BufferSize.Width -lt 135){
    try {
        (Get-Host).UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size(135,(Get-Host).UI.RawUI.BufferSize.Height)
    } catch {}
}
if ((Get-Host).UI.RawUI.WindowSize.Width -lt 135) {
    try {
    (Get-Host).UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size(135,$(Get-Host).ui.rawui.WindowSize.Height)
    } catch {}
}
if ((Get-Host).UI.RawUI.MaxWindowSize.Width -gt 134) {
Write-Host "$(' '*66)@@   @@@@                             @@@                           
$(' '*66)@@   @@@@@                            @@@                           
$(' '*74)@@                                                          
$(' '*74)@@                                                          
$(' '*3)@@@@@@      @@@@@     @@       @@   @@ @@@@@@       @@@@@   @@@@@      @@   @@@      @@@      @@@ @@@@@*     @@@@@@@      @@@@@@   
$(' '*2)@@@*#@@@   @@@@*@@@    @@       @@   @@@@@ #@@@     @@@+@@@  @@@@@      @@    @@@     @@@@     @@  @@@@@@     @@@@@@@@    @@@@@@@@@ 
$(' '*1)@@:    @    @@     @@   @@       @@   @@@     @@    @@     @     @@      @@    @@@    @@@@@    @@@   @@@@@    @@@+  @@@   @@@@  @@@@ 
$(' '*1)@@         @@      @@   @@       @@   @@      #@    @@           @@      @@    @@@    @@@@@    @@@     @@@    @@@         @@@     @@@
$(' '*1)@@         @@       @@  @@       @@   @@       @   @@            @@      @@     @@@   @@ @@@  @@@      @@@    @@@@@       @@@@@@@@@@@
$(' '*1)@@         @@       @@  @@       @@   @@       @   @@            @@      @@      @@* @@@  @@  @@@      @@@       @@@@@@   @@@@@@@@@@@
$(' '*1)@@         @@       @@  @@       @@   @@       @   @@            @@      @@      @@@ @@   @@@ @@       @@@         @@@@   @@         
$(' '*1)@@         @@       @@  @@       @@   @@       @   @@            @@      @@      @@@,@@   @@@@@@       @@@           @@*  @@@        
$(' '*1)@@         @@      @@   @@.     @@@   @@       @    @@           @@      @@       @@@@@   @@@@@@       @@@    @@    .@@   @@@@       
$(' '*2)@@@@@@@@    @@@@@@@     @@@@@@@ @@   @@       @     @@@@@@@     @@@@@   @@@@@    @@@@     @@@@        @@@@@  @@@@@@@@      @@@@@@@@ 
$(' '*4)@@@@       @@@@@        @@@@  @@   @@       @      @@@@@      @@@@@   @@@@@     @@@     @@@@        @@@@@    @@@@@         @@@@@# "
}
Start-Sleep 1
if ($null -eq $InputFile -or [string]::IsNullOrWhiteSpace($InputFile) -eq $true) {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $FileBrowser.Title = "Please Select VMonline XML File"
        $FileBrowser.Filter = 'XML Files (*.xml)| *.xml'
    } catch {
        # Just putting this in to hide the error if one occours
    }
    if ($null -eq $(Get-ChildItem -include *.xml)) {
        try {
            #$FileBrowser.InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        } catch {}
    } else {
        $FileBrowser.InitialDirectory = Get-Location
    }
    
    $NoReally = $false
    If ($FileBrowser.ShowDialog() -eq "Cancel") {
        Write-Host "User canceled file select dialouge" -ForegroundColor Yellow
        $NoReally = $true
    } else {
        if ($FileBrowser.FileNames.count -gt 1) {
            Write-Host "Please only select 1x file" -ForegroundColor Yellow
            $NoReally = $true
        }
        $InputFile = $FileBrowser.FileName
    }
}
if ($NoReally -eq $true) {
    start-sleep 3
    $host.ui.RawUI.WindowTitle = $OldTitle
    break
}
try {
    $InputFile = Get-Item $InputFile
} catch {
    Write-Host "Some error occoured while trying to read the file, please check the path and try again" -ForegroundColor Red
    $NoReally = $true
}
if ($NoReally -eq $true) {
    start-sleep 3
    $host.ui.RawUI.WindowTitle = $OldTitle
    break
}
try {
    Write-Host "Reading data,... please wait..."
    [xml]$Data = Get-Content $InputFile.FullName
} catch {
    Write-Host "XML file didn't read propertly, this is possibly because of a bad formatting error" -ForegroundColor Red
    $NoReally = $true
}
if ($NoReally -eq $true) {
    start-sleep 3
    $host.ui.RawUI.WindowTitle = $OldTitle
    break
}

function Get-FileEncoding([Parameter(Mandatory=$True)]$Path) {
    $bytes = [byte[]](Get-Content $Path -Encoding byte -ReadCount 4 -TotalCount 4)

    if(!$bytes) { return 'utf8' }

    switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
        '^efbbbf'   {return 'utf8'}
        '^2b2f76'   {return 'utf7'}
        '^fffe'     {return 'unicode'}
        '^feff'     {return 'bigendianunicode'}
        '^0000feff' {return 'utf32'}
        default     {return 'ascii'}
    }
}

if ($(Get-FileEncoding $InputFile.FullName) -ne 'utf8') {
    Write-Host "File MAY have issues being processed as it's not formated in UTF8 and is currently $(Get-FileEncoding $InputFile.FullName)`r`nWe are unsure if this will cause issues however errors from VMonline have been experianced saying about a UTF8 requirement" -ForegroundColor Yellow
}

function Debug-DataType {
    Param(
        [ValidateSet('SaleTerms', 'Decimal', 'DateTime', 'String', 'Single', 'Int32', 'Int16', 'Int', 'Double', 'Bool', 'Area', 'OwnerType','PostCode')]
        [string]$Type,
        [string]$Data,
        [string]$DataName,
        [string]$Location,
        $Object_ID,
        [int]$Maxlength,
        [bool]$OnlyPositive,
        [bool]$NullAllowed
    )
    $ErrorItems = @()
    if ($Type -eq 'Decimal' -and $OnlyPositive -eq $true) {
        try {
            $temp123 = [decimal]$Data
            if ($temp123 -lt 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' is less than 0 and therefore invalid"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            }
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'PostCode') {
        if ([string]::IsNullOrWhiteSpace($data) -eq $true) {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName usually can not be blank"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        } else {
            try {
                $temp123 = [int]$Data
                $BadPostcode = $true
                if ([int]$Data -ge 1000 -and [int]$Data -le 2599) {$BadPostcode = $false} #NSW
                if ([int]$Data -ge 2620 -and [int]$Data -le 2899) {$BadPostcode = $false} #NSW
                if ([int]$Data -ge 2921 -and [int]$Data -le 2999) {$BadPostcode = $false} #NSW
                if ([int]$Data -ge 3000 -and [int]$Data -le 3999) {$BadPostcode = $false} #VIC
                if ([int]$Data -ge 8000 -and [int]$Data -le 8999) {$BadPostcode = $false} #VIC
                if ([int]$Data -ge 4000 -and [int]$Data -le 4999) {$BadPostcode = $false} #QLD
                if ([int]$Data -ge 9000 -and [int]$Data -le 9999) {$BadPostcode = $false} #QLD
                if ([int]$Data -ge 5000 -and [int]$Data -le 5999) {$BadPostcode = $false} #SA
                if ([int]$Data -ge 6000 -and [int]$Data -le 6999) {$BadPostcode = $false} #WA
                if ([int]$Data -ge 7000 -and [int]$Data -le 7999) {$BadPostcode = $false} #WA
                if ([int]$Data -ge 200 -and [int]$Data -le 299) {$BadPostcode = $false} #ACT
                if ([int]$Data -ge 2600 -and [int]$Data -le 2619) {$BadPostcode = $false} #ACT
                if ([int]$Data -ge 2900 -and [int]$Data -le 2920) {$BadPostcode = $false} #ACT
                if ([int]$Data -ge 800 -and [int]$Data -le 999) {$BadPostcode = $false} #NT
                if ($BadPostcode -eq $true){
                    $TempArray = New-Object System.Object
                    $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                    $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName is currently '$Data' but acording to AusPost that is invalid"
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                    $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                    $ErrorItems += $TempArray
                }
            } catch {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName cant not be converted to an INT so must be invalid, is currently '$Data'"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            }
        }
    }
    if ($Type -eq 'Decimal') {
        try {
            $temp123 = [decimal]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'DateTime') {
        try {
            $temp123 = [datetime]$Data
            if ($($temp123 -gt [datetime]"1753-01-01" -and $temp123 -lt [datetime]"2199-12-31") -eq $false) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' must be greater than '1753-01-01' and less than '2199-12-31'"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            }
        }
        catch {
            if ([string]::IsNullOrWhiteSpace($Data) -eq $false) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            } else {
                if ($NullAllowed -ne $true) {
                    $TempArray = New-Object System.Object
                    $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                    $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' is NULL and not supposed to be so it cannot be converted to a $Type object"
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                    $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                    $ErrorItems += $TempArray
                }
            }
        }
    }
    if ($Type -eq 'OwnerType') {
        if ($Data -eq 'Individual' -or $Data -eq 'Company') {
            #Do Nothing
        } else {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "'$DataName' must be 'Individual' or 'Company' and is currently '$Data'"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'SaleTerms') {
        $BadTerm = $true
        #This is a bad way to do this validation, however it's a little easier to read.
        if ($Data -eq 'C') {$BadTerm = $false}
        if ($Data -eq 'T') {$BadTerm = $false}
        if ([string]::IsNullOrWhiteSpace($Data) -eq $true) {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "'$DataName' must be 'C' or 'T' and is currently BLANK"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
        else {
            if ($BadTerm -eq $true) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "'$DataName' must be 'C' or 'T' and is currently '$Data'"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            }
        }
    }
    if ($Type -eq 'Area') {
        #This is a bad way to do this validation, however it's a little easier to read.
        if ([string]::IsNullOrWhiteSpace($Data) -eq $false) {
            $BadArea = $true
            if ($Data -eq 'm') {$BadArea = $false}
            if ($Data -eq 'M') {$BadArea = $false}
            if ($Data -eq 'h') {$BadArea = $false}
            if ($Data -eq 'H') {$BadArea = $false}
            if ($BadArea -eq $true) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "'$DataName' must be 'h' or 'H' or 'm' or 'M' and is currently '$Data'"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
                $ErrorItems += $TempArray
            }
        }
    }
    if ($Type -eq 'Bool') {
        if ($Data -ne 'true' -and $Data -ne 'false' -and [string]::IsNullOrWhiteSpace($Data) -eq $false) {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Incorrect value in '$DataName' should be true or false and currently is '$Data'"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'Double') {
        try {
            $temp123 = [Double]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'Int') {
        try {
            $temp123 = [Int]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'Int32') {
        try {
            $temp123 = [Int32]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'Int16') {
        try {
            $temp123 = [Int16]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'Single') {
        try {
            $temp123 = [Single]$Data
        }
        catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName '$Data' cannot be converted to a $Type object"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    if ($Type -eq 'String') {
        if ($($Data | Measure-Object -Character | Select-Object -ExpandProperty Characters) -gt $Maxlength) {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName is too long as it's limited to $Maxlength character length, and it's currently '$($Data | Measure-Object -Character | Select-Object -ExpandProperty Characters)'"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
        $enc = [System.Text.Encoding]::UTF8
        try {
            if ([string]::IsNullOrWhiteSpace($Data) -eq $false) {
                $null = $enc.GetBytes($Data)
            }
        } catch {
            $TempArray = New-Object System.Object
            $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $Object_ID
            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "$DataName is currently a type of $(Get-StringEncoding $Data) and unable to be converted to UTF8"
            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Data'
            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $Location
            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $Data
            $ErrorItems += $TempArray
        }
    }
    Return $ErrorItems
}

function Get-StringEncoding([Parameter(Mandatory=$True)]$String) {
    $bytes = [byte[]][char[]]$String

    if(!$bytes) { return 'utf8' }

    switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
        '^efbbbf'   {return 'utf8'}
        '^2b2f76'   {return 'utf7'}
        '^fffe'     {return 'unicode'}
        '^feff'     {return 'bigendianunicode'}
        '^0000feff' {return 'utf32'}
        default     {return 'ascii'}
    }
}

function Test-IsGuid
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$ObjectGuid
    )
    # Define verification regex
    [regex]$guidRegex = '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'
    # Check guid against regex
    return $ObjectGuid -match $guidRegex
}

$FileType = $Data | Get-Member -MemberType Property
if ($FileType.count -gt 2 -or $FileType.count -eq 0) {
    Write-Host "Something is not right with this file, It should only have 1 or 2 type's however it has $($FileType.count)" -ForegroundColor Red
    break
}
if ($FileType.Name -contains 'Payload' -or $FileType.Name -contains 'payload') {
    $FileType = $Data.Payload | Get-Member -MemberType Property | Where-Object {$_.Name -ne 'OneTimeToken'}
    if ($(Test-IsGuid $data.Payload.OneTimeToken) -eq $false) {
        Write-Host "OneTimeToken is NOT a valid GUID" -ForegroundColor Red
    } else {
        # do nothing as it's fine
    }
    $Payload = $data.Payload.$($FileType.Name)
    $xml = New-Object -TypeName xml
    $xml.AppendChild($xml.ImportNode($Payload, $true)) | Out-Null
    $Data = $xml
} else {
    Write-Host "Your XML file doesnt have a 'payload' element, this may be fine if your testing however VMonline will not accept it without one" -ForegroundColor Yellow
}

# Validate the XML file's root node
$FileNodes = @('Properties'; 'owners'; 'lands'; 'BuildingPermits'; 'Lease'; 'PlanningPermits'; 'Sales'; 'Valuations')
if ($null -eq $($FileNodes -contains $FileType.Name)) {
    Write-Host "The node of the XML file is currently $($FileType.Name) and should apparently be one of the following`r`n $(foreach ($item in $FileNodes) {" - $item`r`n"})" -ForegroundColor Yellow
}
else {
    Write-Host "XML file has a valid root node name of $($FileType.Name)" -ForegroundColor Green
}

#Validate that the file actually has child items
if ($($Data.$($FileType.Name).ChildNodes | Measure-Object | Select-Object -ExpandProperty Count) -eq 0) {
    Write-Host "Something is not right with this file, It appears that no items are present" -ForegroundColor Red
    $ProcessFile = $false
    break
}
else {
        write-host "File has $($Data.$($FileType.Name).ChildNodes | Measure-Object | Select-Object -ExpandProperty Count) items in it" -ForegroundColor Green
    $ProcessFile = $true
}

if ($MissingFields -eq $false) {
    Write-Host "Just an FYI, this will only check to see if the files contents are correct and will not detect missing fields as 'MissingFields' is to to FALSE" -ForegroundColor Yellow
} else {
    Write-Host "Just an FYI, this is going to check for missing fields in the file so it will take alot longer" -ForegroundColor Yellow
}

# Define how many secconds you want to wait before writing out the current status
$StatusSecconds = 1

# Validate each object
if ($ProcessFile -eq $true) {
    Write-Host "Processing a '$($FileType.Name)' file" -ForegroundColor Green
    if ($FileType.Name -eq 'Properties' -or $FileType.Name -eq 'properties') {
        $BadObjects = @()
        $Count = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($item in $Data.Properties.ChildNodes) {
            $count++
            $BadItems = @()
            if ($count -eq 1 -or $stopwatch.Elapsed.Seconds -ge $StatusSecconds) {
                $stopwatch.Restart()
                $Status = "Processing $count of $($Data.Properties.ChildNodes.Count) $(if ($BadObjects.count -gt 0) {"- Errors Found: $($BadObjects.count)"})"
                Write-Progress -Id 1 -Activity $Status -PercentComplete (($count / $($Data.Properties.ChildNodes.Count)) * 100)
            }
            # Check file base data
            if ($item.Name -ne 'Property') {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Incorrect object name, expecting 'Property' and is currently set to '$($item.Name)'"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Bad Object Name'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                $BadItems += $TempArray
            }

            $BadItems += Debug-DataType -Type Double -Data $item.AssessmentNumber -DataName 'AssessmentNumber' -Location 'Object root' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Status -DataName 'Status' -Location 'Object root' -Object_ID $item.AssessmentNumber -Maxlength 50

            if ($MissingFields -eq $true) {
                $FileObjects = @('FireServices'; 'Address'; 'CouncilInformation'; 'SiteInformation'; 'Owners')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }

            <#------------- Check 'FireServices' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('FireServiceLeviable'; 'FireServiceArea'; 'FireServiceSection20')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.FireServices | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'FireServices'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Bool -Data $item.FireServices.FireServiceLeviable -DataName 'FireServiceLeviable' -Location 'FireServices' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Address.FireServiceArea -DataName 'FireServiceArea' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type Bool -Data $item.FireServices.FireServiceSection20 -DataName 'FireServiceSection20' -Location 'FireServices' -Object_ID $item.AssessmentNumber
            <#--------------------------------------------------------#>

            <#------------- Check 'Address' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('PropertyName'; 'FloorNumber'; 'FloorType'; 'UnitNumber'; 'UnitType'; 'UnitNumSuffix', 'FromStNum', 'FromStNumPrefix', 'FromStNumSuffix', 'ToStNum', 'ToStNumPrefix', 'ToStNumSuffix', 'SourceStreet', 'SourceStreetType', 'SourceStreetPrefix', 'SourceStreetSuffix', 'SourceSuburb', 'PostCode', 'FormatPropAddress')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.Address | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Address'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }

            $BadItems += Debug-DataType -Type String -Data $item.Address.PropertyName -DataName 'PropertyName' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 4000
            $BadItems += Debug-DataType -Type String -Data $item.Address.FloorNumber -DataName 'FloorNumber' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.Address.FloorType -DataName 'FloorType' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type Int32 -Data $item.Address.UnitNumber -DataName 'UnitNumber' -Location 'Address' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Address.UnitType -DataName 'UnitType' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type String -Data $item.Address.UnitNumSuffix -DataName 'UnitNumSuffix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type Int16 -Data $item.Address.FromStNum -DataName 'FromStNum' -Location 'Address' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Address.FromStNumPrefix -DataName 'FromStNumPrefix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 1
            $BadItems += Debug-DataType -Type String -Data $item.Address.FromStNumSuffix -DataName 'FromStNumSuffix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 3
            $BadItems += Debug-DataType -Type Int16 -Data $item.Address.ToStNum -DataName 'ToStNum' -Location 'Address' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Address.ToStNumPrefix -DataName 'ToStNumPrefix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 1
            $BadItems += Debug-DataType -Type String -Data $item.Address.ToStNumSuffix -DataName 'ToStNumSuffix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 1
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceStreet -DataName 'SourceStreet' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceStreetType -DataName 'SourceStreetType' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceStreetPrefix -DataName 'SourceStreetPrefix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceStreetSuffix -DataName 'SourceStreetSuffix' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceSuburb -DataName 'SourceSuburb' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.Address.SourceSuburb -DataName 'SourceSuburb' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type PostCode -Data $item.Address.PostCode -DataName 'PostCode' -Location 'Address' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Address.FormatPropAddress -DataName 'FormatPropAddress' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 4000
            <#--------------------------------------------------------#>

            <#------------- Check 'CouncilInformation' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Rateable'; 'ParentCode'; 'CouncilNo'; 'RatingCode'; 'WaterAuthority'; 'Ward', 'GIS_Reference', 'Xref', 'GIS_Measure', 'GIS_Area', 'Zones', 'Overlays')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.CouncilInformation | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'CouncilInformation'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Bool -Data $item.CouncilInformation.Rateable -DataName 'Rateable' -Location 'CouncilInformation' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type Double -Data $item.CouncilInformation.ParentCode -DataName 'ParentCode' -Location 'CouncilInformation' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.CouncilNo -DataName 'CouncilNo' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.RatingCode -DataName 'RatingCode' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.WaterAuthority -DataName 'WaterAuthority' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 250
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.Ward -DataName 'Ward' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.GIS_Reference -DataName 'GIS_Reference' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.CouncilInformation.Xref -DataName 'Xref' -Location 'Address' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type Area -Data $item.CouncilInformation.GIS_Measure -DataName 'GIS_Measure' -Location 'Address' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type Double -Data $item.CouncilInformation.GIS_Area -DataName 'GIS_Area' -Location 'Address' -Object_ID $item.AssessmentNumber
    
            <#------------- Check 'CouncilInformation\Zones' object --------------#>
            if ($null -ne $item.CouncilInformation.Zones) {
                foreach ($Zone in $item.CouncilInformation.Zones) {
                    $BadItems += Debug-DataType -Type String -Data $Zone.Zone -DataName 'Zone' -Location 'CouncilInformation\Zones' -Object_ID $item.AssessmentNumber -Maxlength 10
                }
            }
            <#------------- Check 'CouncilInformation\Overlays' object --------------#>
            if ($null -ne $item.CouncilInformation.Overlays) {
                foreach ($Overlay in $item.CouncilInformation.Overlays) {
                    $BadItems += Debug-DataType -Type String -Data $Overlay.Overlay -DataName 'Overlay' -Location 'CouncilInformation\Overlays' -Object_ID $item.AssessmentNumber -Maxlength 50
                }
            }
            <#--------------------------------------------------------#>

            <#------------- Check 'SiteInformation' object --------------#>
    
            $FileObjects = @('HeaderID'; 'AVPCC'; 'PropertyArea'; 'Measure'; 'WidthFront'; 'WidthBack', 'DepthLeft', 'DepthRight')
            foreach ($obj in $FileObjects) {
                if ($null -eq $item.SiteInformation -and $MissingFields -eq $true) {
                    $TempArray = New-Object System.Object
                    #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                    $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object SiteInformation"
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Object'
                    $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                    $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                    $BadItems += $TempArray
                }
                else {
                    if ($MissingFields -eq $true) {
                        if ($null -eq $($item.SiteInformation | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                            $TempArray = New-Object System.Object
                            #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                            $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                            $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                            $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'SiteInformation'
                            $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                            $BadItems += $TempArray
                        }
                    }
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.HeaderID -DataName 'HeaderID' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type String -Data $item.SiteInformation.HeaderID -DataName 'AVPCC' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber -Maxlength 6
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.PropertyArea -DataName 'PropertyArea' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type Area -Data $item.SiteInformation.Measure -DataName 'Measure' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.WidthFront -DataName 'WidthFront' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.WidthBack -DataName 'WidthBack' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.DepthLeft -DataName 'DepthLeft' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                    $BadItems += Debug-DataType -Type Double -Data $item.SiteInformation.DepthRight -DataName 'DepthRight' -Location 'SiteInformation' -Object_ID $item.AssessmentNumber
                }
        
            }

            <#--------------------------------------------------------#>

            <#------------- Check 'Owners' object --------------#>
            # Will come back to this
            <#--------------------------------------------------------#>
            if ($BadItems.Count -ne 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.AssessmentNumber
                $TempArray | Add-Member -type NoteProperty -name "Possible_Errors" -Value $BadItems.Count
                $TempArray | Add-Member -type NoteProperty -name "Errors" -Value $BadItems
                $BadObjects += $TempArray
            }
        }
    }

    if ($FileType.Name -eq 'Valuations' -or $FileType.Name -eq 'valuations') {
        $BadObjects = @()
        $Count = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($item in $Data.Valuations.ChildNodes) {
            $count++
            $BadItems = @()
            if ($count -eq 1 -or $stopwatch.Elapsed.Seconds -ge $StatusSecconds) {
                $stopwatch.Restart()
                $Status = "Processing $count of $($Data.Valuations.ChildNodes.Count) $(if ($BadObjects.count -gt 0) {"- Errors Found: $($BadObjects.count)"})"
                Write-Progress -Id 1 -Activity $Status -PercentComplete (($count / $($Data.Valuations.ChildNodes.Count)) * 100)
            }
            if ($MissingFields -eq $true) {
                $FileObjects = @('AssessmentNumber'; 'CouncilSV'; 'CouncilCIV'; 'CouncilNAV'; 'OverrideCIV';'OverrideSV';'OverrideNAV';'EffectiveDate';'CouncilEffectiveDate')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Double -Data $item.AssessmentNumber -DataName 'AssessmentNumber' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.CouncilSV -DataName 'CouncilSV' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.CouncilCIV -DataName 'CouncilCIV' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.CouncilNAV -DataName 'CouncilNAV' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.OverrideCIV -DataName 'OverrideCIV' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.OverrideSV -DataName 'OverrideSV' -Location 'Object root'
            $BadItems += Debug-DataType -Type Decimal -Data $item.OverrideNAV -DataName 'OverrideNAV' -Location 'Object root'
            $BadItems += Debug-DataType -Type DateTime -Data $item.EffectiveDate -DataName 'EffectiveDate' -Location 'Object root'
            $BadItems += Debug-DataType -Type DateTime -Data $item.CouncilEffectiveDate -DataName 'CouncilEffectiveDate' -Location 'Object root'

            if ($BadItems.Count -ne 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.AssessmentNumber
                $TempArray | Add-Member -type NoteProperty -name "Possible_Errors" -Value $BadItems.Count
                $TempArray | Add-Member -type NoteProperty -name "Errors" -Value $BadItems
                $BadObjects += $TempArray
            }
        }
    }

    if ($FileType.Name -eq 'owners' -or $FileType.Name -eq 'Owners') {
        $BadObjects = @()
        $Count = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($item in $Data.owners.ChildNodes) {
            $count++
            $BadItems = @()
            if ($count -eq 1 -or $stopwatch.Elapsed.Seconds -ge $StatusSecconds) {
                $stopwatch.Restart()
                $Status = "Processing $count of $($Data.owners.ChildNodes.Count) $(if ($BadObjects.count -gt 0) {"- Errors Found: $($BadObjects.count)"})"
                Write-Progress -Id 1 -Activity $Status -PercentComplete (($count / $($Data.owners.ChildNodes.Count)) * 100)
            }
            if ($MissingFields -eq $true) {
                $FileObjects = @('OwnerNumber'; 'OwnerType'; 'SurName'; 'OtherNames'; 'FullAddress')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Int -Data $item.OwnerNumber -DataName 'OwnerNumber' -Location 'Object root' -Object_ID $item.OwnerNumber
            $BadItems += Debug-DataType -Type OwnerType -Data $item.OwnerType -DataName 'OwnerType' -Location 'Object root' -Object_ID $item.OwnerNumber
            $BadItems += Debug-DataType -Type String -Data $item.SurName -DataName 'SurName' -Location 'Object root' -Object_ID $item.OwnerNumber -Maxlength 255
            $BadItems += Debug-DataType -Type String -Data $item.OtherNames -DataName 'OtherNames' -Location 'Object root' -Object_ID $item.OwnerNumber -Maxlength 255
            $BadItems += Debug-DataType -Type String -Data $item.FullAddress -DataName 'FullAddress' -Location 'Object root' -Object_ID $item.OwnerNumber -Maxlength 250

            if ($BadItems.Count -ne 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.OwnerNumber
                $TempArray | Add-Member -type NoteProperty -name "Possible_Errors" -Value $BadItems.Count
                $TempArray | Add-Member -type NoteProperty -name "Errors" -Value $BadItems
                $BadObjects += $TempArray
            }
        }
    }

    if ($FileType.Name -eq 'Sales' -or $FileType.Name -eq 'sales') {
        $BadObjects = @()
        $Count = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($item in $Data.Sales.ChildNodes) {
            $count++
            $BadItems = @()
            if ($count -eq 1 -or $stopwatch.Elapsed.Seconds -ge $StatusSecconds) {
                $stopwatch.Restart()
                $Status = "Processing $count of $($Data.Sales.ChildNodes.Count) $(if ($BadObjects.count -gt 0) {"- Errors Found: $($BadObjects.count)"})"
                Write-Progress -Id 1 -Activity $Status -PercentComplete (($count / $($Data.Sales.ChildNodes.Count)) * 100)
            }
            if ($MissingFields -eq $true) {
                $FileObjects = @('AssessmentNumber'; 'SaleID'; 'TransactionNumber'; 'SalePrice'; 'ContractDate', 'InformationDate', 'SettlementDate', 'Comments', 'SalesDetails', 'Names')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Double -Data $item.AssessmentNumber -DataName 'AssessmentNumber' -Location 'Object root' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.SaleID -DataName 'SaleID' -Location 'Object root' -Object_ID $item.AssessmentNumber -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.TransactionNumber -DataName 'TransactionNumber' -Location 'Object root' -Object_ID $item.AssessmentNumber -Maxlength 10
            $BadItems += Debug-DataType -Type Decimal -Data $item.SalePrice -DataName 'SalePrice' -Location 'Object root' -Object_ID $item.AssessmentNumber -OnlyPositive $true
            $BadItems += Debug-DataType -Type DateTime -Data $item.ContractDate -DataName 'ContractDate' -Location 'Object root' -NullAllowed $true -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type DateTime -Data $item.InformationDate -DataName 'InformationDate' -Location 'Object root' -NullAllowed $true -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type DateTime -Data $item.SettlementDate -DataName 'SettlementDate' -Location 'Object root' -NullAllowed $true -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Comments -DataName 'Comments' -Location 'Object root' -Maxlength 255 -Object_ID $item.AssessmentNumber
        
            <#------------- Check 'SalesDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Terms'; 'DetailsOfSale'; 'TransferType')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.SalesDetails | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'SalesDetails'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }

            $BadItems += Debug-DataType -Type SaleTerms -Data $item.SalesDetails.Terms -DataName 'Terms' -Location 'SalesDetails' -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.SalesDetails.DetailsOfSale -DataName 'DetailsOfSale' -Location 'SalesDetails' -Maxlength 255 -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.SalesDetails.TransferType -DataName 'TransferType' -Location 'SalesDetails' -Maxlength 254 -Object_ID $item.AssessmentNumber
            <#--------------------------------------------------------#>

            <#------------- Check 'SalesDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('VendorName'; 'EmptorName')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.Names | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Names'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type String -Data $item.Names.VendorName -DataName 'VendorName' -Location 'Names' -Maxlength 254 -Object_ID $item.AssessmentNumber
            $BadItems += Debug-DataType -Type String -Data $item.Names.EmptorName -DataName 'EmptorName' -Location 'Names' -Maxlength 254 -Object_ID $item.AssessmentNumber
            <#--------------------------------------------------------#>

            if ($BadItems.Count -ne 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.AssessmentNumber
                $TempArray | Add-Member -type NoteProperty -name "Possible_Errors" -Value $BadItems.Count
                $TempArray | Add-Member -type NoteProperty -name "Errors" -Value $BadItems
                $BadObjects += $TempArray
            }
        }
    }

    if ($FileType.Name -eq 'lands' -or $FileType.Name -eq 'Lands') {
        $BadObjects = @()
        $Count = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($item in $Data.lands.ChildNodes) {
            $count++
            #if ($item.LandNumber -eq 42){break}
            $BadItems = @()
            if ($count -eq 1 -or $stopwatch.Elapsed.Seconds -ge $StatusSecconds) {
                $stopwatch.Restart()
                $Status = "Processing $count of $($Data.lands.ChildNodes.Count) $(if ($BadObjects.count -gt 0) {"- Errors Found: $($BadObjects.count)"})"
                Write-Progress -Id 1 -Activity "Processing an '$($FileType.Name)' file" -status $Status -PercentComplete (($count / $($Data.lands.ChildNodes.Count)) * 100)
            }
            if ($MissingFields -eq $true) {
                $FileObjects = @('LandNumber'; 'Description'; 'AssociatedAssessments'; 'PlanDetails'; 'CrownDetails', 'VolDetails', 'LandDetails')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        #$TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'Object root'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type String -Data $item.LandNumber -DataName 'LandNumber' -Location 'Object root' -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.Description -DataName 'Description' -Location 'Object root' -Maxlength 150
        
            <#------------- Check 'AssociatedAssessments' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Terms'; 'DetailsOfSale'; 'TransferType')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.AssociatedAssessments | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'AssociatedAssessments'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            if ([string]::IsNullOrWhiteSpace($($item.AssociatedAssessments -join ',')) -eq $false) {
                foreach ($Assesment in $item.AssociatedAssessments.AssessmentNumber) {
                    $BadItems += Debug-DataType -Type Double -Data $Assesment -DataName 'AssessmentNumber' -Location 'AssociatedAssessments'
                }
            } else {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Parcel missing a Assesment linkage"
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Data'
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'AssociatedAssessments'
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                $BadItems += $TempArray
            }
            <#--------------------------------------------------------#>

            <#------------- Check 'PlanDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Lot'; 'PartLot'; 'PlanType'; 'PlanNumber')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.PlanDetails | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'PlanDetails'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type String -Data $item.PlanDetails.Lot -DataName 'Lot' -Location 'PlanDetails' -Maxlength 30
            $BadItems += Debug-DataType -Type Bool -Data $item.PlanDetails.PartLot -DataName 'PartLot' -Location 'PlanDetails'
            $BadItems += Debug-DataType -Type String -Data $item.PlanDetails.PlanType -DataName 'PlanType' -Location 'PlanDetails' -Maxlength 250
            $BadItems += Debug-DataType -Type String -Data $item.PlanDetails.PlanNumber -DataName 'PlanNumber' -Location 'PlanDetails' -Maxlength 50
            <#--------------------------------------------------------#>

            <#------------- Check 'CrownDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Parish'; 'CrownAllot'; 'PartCrownAllot'; 'Section'; 'PartSection'; 'Portion'; 'PartPortion'; 'Book'; 'Block')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.CrownDetails | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'CrownDetails'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.Parish -DataName 'Parish' -Location 'CrownDetails' -Maxlength 50
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.CrownAllot -DataName 'CrownAllot' -Location 'CrownDetails' -Maxlength 12
            $BadItems += Debug-DataType -Type Bool -Data $item.CrownDetails.PartCrownAllot -DataName 'PartCrownAllot' -Location 'CrownDetails'
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.Section -DataName 'Section' -Location 'CrownDetails' -Maxlength 7
            $BadItems += Debug-DataType -Type Bool -Data $item.CrownDetails.PartSection -DataName 'PartSection' -Location 'CrownDetails'
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.Portion -DataName 'Portion' -Location 'CrownDetails' -Maxlength 4
            $BadItems += Debug-DataType -Type Bool -Data $item.CrownDetails.PartPortion -DataName 'PartPortion' -Location 'CrownDetails'
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.Book -DataName 'Book' -Location 'CrownDetails' -Maxlength 5
            $BadItems += Debug-DataType -Type String -Data $item.CrownDetails.Block -DataName 'Block' -Location 'CrownDetails' -Maxlength 2
            <#--------------------------------------------------------#>

            <#------------- Check 'VolDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('Volume'; 'PartVolume'; 'Folio'; 'PartFolio')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.VolDetails | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'VolDetails'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type String -Data $item.VolDetails.Volume -DataName 'Volume' -Location 'VolDetails' -Maxlength 50
            $BadItems += Debug-DataType -Type Bool -Data $item.VolDetails.PartVolume -DataName 'PartVolume' -Location 'VolDetails'
            $BadItems += Debug-DataType -Type String -Data $item.VolDetails.Folio -DataName 'Folio' -Location 'VolDetails' -Maxlength 50
            $BadItems += Debug-DataType -Type Bool -Data $item.VolDetails.PartFolio -DataName 'PartFolio' -Location 'VolDetails'
        
            <#--------------------------------------------------------#>

            <#------------- Check 'LandDetails' object --------------#>
            if ($MissingFields -eq $true) {
                $FileObjects = @('LandArea'; 'LandMeasure'; 'SPICode')
                foreach ($obj in $FileObjects) {
                    if ($null -eq $($item.LandDetails | Get-Member -MemberType Property | Where-Object 'Name' -eq $obj)) {
                        $TempArray = New-Object System.Object
                        $TempArray | Add-Member -type NoteProperty -name "Issue" -Value "Missing object property '$($obj.ToString())'"
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value 'Missing Feild'
                        $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value 'LandDetails'
                        $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value 'N/A'
                        $BadItems += $TempArray
                    }
                }
            }
            $BadItems += Debug-DataType -Type Double -Data $item.LandDetails.LandArea -DataName 'LandArea' -Location 'LandDetails'
            $BadItems += Debug-DataType -Type Area -Data $item.LandDetails.LandMeasure -DataName 'LandMeasure' -Location 'LandDetails'
            $BadItems += Debug-DataType -Type String -Data $item.LandDetails.SPICode -DataName 'SPICode' -Location 'LandDetails' -Maxlength 50
            <#--------------------------------------------------------#>


            if ($BadItems.Count -ne 0) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_Number" -Value $Count
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.LandNumber
                $TempArray | Add-Member -type NoteProperty -name "Possible_Errors" -Value $BadItems.Count
                $TempArray | Add-Member -type NoteProperty -name "Errors" -Value $BadItems
                $BadObjects += $TempArray
            }
        }
    }

    Write-Progress -Id 1 -Completed -Activity "Finishing up"
    $FileDate = ((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")
    if ($BadObjects.count -gt 0) {
        Write-Host "$($InputFile.FullName) has $($BadObjects.count) errors`r`nExporting data to the same location as the file you loaded" -ForegroundColor Red
        $BadObjects | ConvertTo-Json | Out-File $($InputFile.Directory.FullName + '\' + $InputFile.BaseName + '_' + $FileDate + '_ERRORS.json') -Force
        $FlatArray = @()
        foreach ($item in $BadObjects) {
            $Object_ID = $Null
            if ([string]::IsNullOrWhiteSpace($item.LandNumber) -eq $false) {
                $Object_ID = $item.LandNumber
            }
            if ([string]::IsNullOrWhiteSpace($item.AssessmentNumber) -eq $false) {
                $Object_ID = $item.AssessmentNumber
            }
            if ([string]::IsNullOrWhiteSpace($item.OwnerNumber) -eq $false) {
                $Object_ID = $item.OwnerNumber
            }
            
            Foreach ($obj in $item.Errors) {
                $TempArray = New-Object System.Object
                $TempArray | Add-Member -type NoteProperty -name "Object_ID" -Value $item.Object_ID
                $TempArray | Add-Member -type NoteProperty -name "Issue" -Value $obj.Issue
                $TempArray | Add-Member -type NoteProperty -name "Issue_Type" -Value $obj.Issue_Type
                $TempArray | Add-Member -type NoteProperty -name "Issue_Location" -Value $obj.Issue_Location
                $TempArray | Add-Member -type NoteProperty -name "Current_Data" -Value $obj.Current_Data
                $FlatArray += $TempArray
            }
        }
        $FlatArray | Export-Csv -NoTypeInformation $($InputFile.Directory.FullName + '\' + $InputFile.BaseName + '_' + $FileDate + '_ERRORS.csv')
        pause
    }
    else {
        Write-Host "$($InputFile.FullName) Passed validation" -ForegroundColor Green
        pause
    }
}
$host.ui.RawUI.WindowTitle = $OldTitle
# SIG # Begin signature block
# MIIhGAYJKoZIhvcNAQcCoIIhCTCCIQUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUKTS8z8+BQZxgcRhLilp8c7nF
# gpOgghw+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTEwggQZ
# oAMCAQICEAqhJdbWMht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTE2MDEwNzEyMDAwMFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnF
# OVQoV7YjSsQOB0UzURB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQA
# OPcuHjvuzKb2Mln+X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhis
# EeTwmQNtO4V8CdPuXciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQj
# MF287DxgaqwvB8z98OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+f
# MRTWrdXyZMt7HgXQhBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW
# /5MCAwEAAaOCAc4wggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/
# AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDBQBgNVHSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggEBAHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafD
# DiBCLK938ysfDCFaKrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6
# HHssIeLWWywUNUMEaLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4
# H9YLFKWA1xJHcLN11ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHK
# eZR+WfyMD+NvtQEmtmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIo
# xhhWz0E0tmZdtnR79VYzIi8iNrJLokqV2PWmjlIwggXYMIIDwKADAgECAhBMqvnK
# 22Nv4B/3TthbA4adMA0GCSqGSIb3DQEBDAUAMIGFMQswCQYDVQQGEwJHQjEbMBkG
# A1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYD
# VQQKExFDT01PRE8gQ0EgTGltaXRlZDErMCkGA1UEAxMiQ09NT0RPIFJTQSBDZXJ0
# aWZpY2F0aW9uIEF1dGhvcml0eTAeFw0xMDAxMTkwMDAwMDBaFw0zODAxMTgyMzU5
# NTlaMIGFMQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDEr
# MCkGA1UEAxMiQ09NT0RPIFJTQSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJHoVJLSClaxrA0k3cXPRGd0mSs3
# o30jcABxvFPfxPoqEo9LfxBWvZ9wcrdhf8lLDxenPeOwBGHu/xGXx/SGPgr6Plz5
# k+Y0etkUa+ecs4Wggnp2r3GQ1+z9DfqcbPrfsIL0FH75vsSmL09/mX+1/GdDcr0M
# ANaJ62ss0+2PmBwUq37l42782KjkkiTaQ2tiuFX96sG8bLaL8w6NmuSbbGmZ+HhI
# MEXVreENPEVg/DKWUSe8Z8PKLrZr6kbHxyCgsR9l3kgIuqROqfKDRjeE6+jMgUhD
# Z05yKptcvUwbKIpcInu0q5jZ7uBRg8MJRk5tPpn6lRfafDNXQTyNUe0LtlyvLGMa
# 31fIP7zpXcSbr0WZ4qNaJLS6qVY9z2+q/0lYvvCo//S4rek3+7q49As6+ehDQh6J
# 2ITLE/HZu+GJYLiMKFasFB2cCudx688O3T2plqFIvTz3r7UNIkzAEYHsVjv206Li
# W7eyBCJSlYCTaeiOTGXxkQMtcHQC6otnFSlpUgK7199QalVGv6CjKGF/cNDDoqos
# IapHziicBkV2v4IYJ7TVrrTLUOZr9EyGcTDppt8WhuDY/0Dd+9BCiH+jMzouXB5B
# EYFjzhhxayvspoq3MVw6akfgw3lZ1iAar/JqmKpyvFdK0kuduxD8sExB5e0dPV4o
# nZzMv7NR2qdH5YRTAgMBAAGjQjBAMB0GA1UdDgQWBBS7r34CPfqm8TyEjq3uOJjs
# 2TIy1DAOBgNVHQ8BAf8EBAMCAQYwDwYDVR0TAQH/BAUwAwEB/zANBgkqhkiG9w0B
# AQwFAAOCAgEACvHVRoS3rlG7bLJNQRQAk0ycy+XAVM+gJY4C+f2wog31IJg8Ey2s
# VqKw1n4Rkukuup4umnKxvRlEbGE1opq0FhJpWozh1z6kGugvA/SuYR0QGyqki3rF
# /gWm4cDWyP6ero8ruj2Z+NhzCVhGbqac9Ncn05XaN4NyHNNz4KJHmQM4XdVJeQAp
# HMfsmyAcByRpV3iyOfw6hKC1nHyNvy6TYie3OdoXGK69PAlo/4SbPNXWCwPjV54U
# 99HrT8i9hyO3tklDeYVcuuuSC6HG6GioTBaxGpkK6FMskruhCRh1DGWoe8sjtxrC
# KIXDG//QK2LvpHsJkZhnjBQBzWgGamMhdQOAiIpugcaF8qmkLef0pSQQR4PKzfSN
# eVixBpvnGirZnQHXlH3tA0rK8NvoqQE+9VaZyR6OST275Qm54E9Jkj0WgkDMzFnG
# 5jrtEi5pPGyVsf2qHXt/hr4eDjJG+/sTj3V/TItLRmP+ADRAcMHDuaHdpnDiBLNB
# vOmAkepknHrhIgOpnG5vDmVPbIeHXvNuoPl1pZtA6FOyJ51KucB3IY3/h/LevIzv
# F9+3SQvR8m4wCxoOTnbtEfz16Vayfb/HbQqTjKXQwLYdvjpOlKLXbmwLwop8+iDz
# xOTlzQ2oy5GSsXyF7LUUaWYOgufNzsgtplF/IcE1U4UGSl2frbsbX3QwggYBMIIE
# 6aADAgECAhBpS/3HyJmGXVmQNrYwJn0ZMA0GCSqGSIb3DQEBCwUAMIGRMQswCQYD
# VQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdT
# YWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDE3MDUGA1UEAxMuQ09N
# T0RPIFJTQSBFeHRlbmRlZCBWYWxpZGF0aW9uIENvZGUgU2lnbmluZyBDQTAeFw0x
# OTA4MDkwMDAwMDBaFw0yMjA4MDgyMzU5NTlaMIHmMRcwFQYDVQQFEw4yMyA2MTgg
# OTA2IDcwMDETMBEGCysGAQQBgjc8AgEDEwJBVTEdMBsGA1UEDxMUUHJpdmF0ZSBP
# cmdhbml6YXRpb24xCzAJBgNVBAYTAkFVMQ0wCwYDVQQRDAQ3MDE3MREwDwYDVQQI
# DAhUYXNtYW5pYTESMBAGA1UEBwwJT2xkIEJlYWNoMRYwFAYDVQQJDA0xIFRpdm9s
# aSBSb2FkMR0wGwYDVQQKDBRDb3VuY2lsIFdpc2UgUHR5IEx0ZDEdMBsGA1UEAwwU
# Q291bmNpbCBXaXNlIFB0eSBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQDXbZkLdlDB5eKFj+4TBiWiTT8JDmycwK/JKjU7OocPxh/sBqwQ44tLXgCk
# mOGLLeHTj98/D7mB7iXldhO90itDVayt6RJJpDeisKmdpn78SMQ1Gd7Lj2ue42Wz
# 8GWGb9u4lrN71gNoWBuDydpx9rzxsEnuqqDy4aQtbh0WtuHNQmgZEAi2aH2Cf3HY
# uY9wuv5flmrlwT0UO3muEFBzsJPVWZkpPeEnnNjPhe2yIolKR8AJt5y1sUGeLRYm
# b770S2fhF+EOSc9ocrOx26ByW3GyeD4dPK2Jxm5kgKMdCcgdISq9xpwHN/AQOcMV
# +YZ+jEZCA7dsWEKTTXjxWfy8qwnHAgMBAAGjggH8MIIB+DAfBgNVHSMEGDAWgBTf
# j/MgDOnKpgTYW1g3Kj2rRtyDSTAdBgNVHQ4EFgQUN9hPbfu/oxb+QWhhcihbV64a
# 1m0wDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEYGA1UdIAQ/MD0wOwYMKwYBBAGyMQEC
# AQYBMCswKQYIKwYBBQUHAgEWHWh0dHBzOi8vc2VjdXJlLmNvbW9kby5jb20vQ1BT
# MFUGA1UdHwROMEwwSqBIoEaGRGh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0NPTU9E
# T1JTQUV4dGVuZGVkVmFsaWRhdGlvbkNvZGVTaWduaW5nQ0EuY3JsMIGGBggrBgEF
# BQcBAQR6MHgwUAYIKwYBBQUHMAKGRGh0dHA6Ly9jcnQuY29tb2RvY2EuY29tL0NP
# TU9ET1JTQUV4dGVuZGVkVmFsaWRhdGlvbkNvZGVTaWduaW5nQ0EuY3J0MCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wSAYDVR0RBEEwP6AhBggr
# BgEFBQcIA6AVMBMMEUFVLTIzIDYxOCA5MDYgNzAwgRpzdXBwb3J0QGNvdW5jaWx3
# aXNlLmNvbS5hdTANBgkqhkiG9w0BAQsFAAOCAQEANy4t9P9WmyW+eJN1HA3Qe8s5
# ovL7IRCj4cD0ch0QhyTCtH0zPxgE9nyDflCFNiOUdA3yA9RVUfz6U1lLhRSlg6O3
# Pfv5y4iYPbjbDCKHsfmT4VYNQimJiRhxg7InS1Fjx3Pkii7l0V+TWdn6yN8bOMbD
# ZMUu2y7nEu7wmtaG8x7QMWB0juTwYZc199W2G7Mqb0yqyHmsz/9j4FY+T1aYMlRi
# qZOTJvOHTvt3AGGsgpfsx2mfjmN4EoWY6akj3hXM3x0gQB9UrzrP2zcZ+BP970ud
# nlAaPtf+16dA0QBAzZ+JW+5WwispPBdfZw873aYl53SvCy0M1aoEyhTHgQuLGzCC
# BiIwggQKoAMCAQICEG3UcusCrgQG492EP1/hReEwDQYJKoZIhvcNAQEMBQAwgYUx
# CzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNV
# BAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVkMSswKQYDVQQD
# EyJDT01PRE8gUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTE0MTIwMzAw
# MDAwMFoXDTI5MTIwMjIzNTk1OVowgZExCzAJBgNVBAYTAkdCMRswGQYDVQQIExJH
# cmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNP
# TU9ETyBDQSBMaW1pdGVkMTcwNQYDVQQDEy5DT01PRE8gUlNBIEV4dGVuZGVkIFZh
# bGlkYXRpb24gQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAiv29Q/A9yFUf81mK8Fq03JPRZBVKioSlLcsm+OBFOKO5AcVPEwhZ
# 0DFUys2QYaM+LPJNFVzU7sNqIpdI0QZDNAmZyc8wxJ9E/Vac7szng7mBzcjaCxwS
# SP9vouzEdcsJcM9R5buLn6q9eAZ9ldZhgfbaU8esnbMAuh7UvkBiCZmDPUXdTWWV
# BMz8+sdbeuIuDD1VNVc1SImJ8rlWpUtQGxzemJC98y7ciKnxdZuoPqIF2UG173et
# F8Ba9aPbTZ/RxLF7g7XuEJQrLnKuvu+VKZxSYsUsbSL3fUR6EF9jk2lN2X2ymrFO
# tVm//4X7vazs4Sum4yws6Nlu219NF3jLUwIDAQABo4IBfjCCAXowHwYDVR0jBBgw
# FoAUu69+Aj36pvE8hI6t7jiY7NkyMtQwHQYDVR0OBBYEFN+P8yAM6cqmBNhbWDcq
# PatG3INJMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMD4GA1UdIAQ3MDUwMwYEVR0gADArMCkGCCsGAQUFBwIB
# Fh1odHRwczovL3NlY3VyZS5jb21vZG8uY29tL0NQUzBMBgNVHR8ERTBDMEGgP6A9
# hjtodHRwOi8vY3JsLmNvbW9kb2NhLmNvbS9DT01PRE9SU0FDZXJ0aWZpY2F0aW9u
# QXV0aG9yaXR5LmNybDBxBggrBgEFBQcBAQRlMGMwOwYIKwYBBQUHMAKGL2h0dHA6
# Ly9jcnQuY29tb2RvY2EuY29tL0NPTU9ET1JTQUFkZFRydXN0Q0EuY3J0MCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZIhvcNAQEMBQAD
# ggIBAGZO7LcWd28R6Btdak7Z8otssVYoQIvAMcSZSCM9+A7ogJfvbSALHxPEhvsX
# NBXhjlT3wrgAcxXgKNnauvqCVML367/DNtAwn+WhHJTf73zo9ix4oqzPJmoVoRUx
# 1jE0mL1TT8SEg6PEllw92P7W+VT/Z5Nt+D4rayyiCHxWSIEyGLJurJDB2+TeOYuG
# 5ccYQFmk35ZHurJ/sfhXD4WAdDgOOlhiHv5S4+auUwmG/o+b21ZWzAewicEE8VML
# bG937LIf7PZbQENgDxurGFS0EASO+A7py4OxevI0TmpUTOmDKumwMCUczmKODuuF
# 5in+sUrj8q48kfVMob7IFw5cu0JN4xqKks0+IH7d6XWx6h90XJ5UwpQ3smHdBxZZ
# f5aAFuCZtdJusMkjBhWs0SP0M4vOdfDBhtP/4S76kE/+Rvm720+7t/7RDSsE8dLR
# lYUsii64hVbyw4RSoekzsetQyKGwn+PDizqHnudV09NtNBcwDWgiC9W57XM1csPt
# pzfN40OuRc00vyjKh2LtQ6Sv+ssxyyFYYUZetsZ6ph5TKqj4XFEfOloQDyjA5HSL
# dMYEqvhLJigKMonbnSpgcWrDlk4WuWO/YZVnjEsuu7BOg+lNMeWOJyL1PCZ7RJHT
# 1Frw03z0OL4UmpkOi7Fb6uSLDxGdd0KCHFw61NqriC+NVzBUMYIERDCCBEACAQEw
# gaYwgZExCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIx
# EDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVkMTcw
# NQYDVQQDEy5DT01PRE8gUlNBIEV4dGVuZGVkIFZhbGlkYXRpb24gQ29kZSBTaWdu
# aW5nIENBAhBpS/3HyJmGXVmQNrYwJn0ZMAkGBSsOAwIaBQCgQDAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAjBgkqhkiG9w0BCQQxFgQU8umhYGg9HIooN8pEjFs8
# lcLSCjowDQYJKoZIhvcNAQEBBQAEggEAf7wQgSZLu1XHbjoWnXOvfKfsNF/OkIXR
# Xg8lv3wA/nFEIHd63UWQ9RccbG/yAYXCMjKl383qWB44ThY7wV6v4wHmmUnaBCZq
# fvNPscOEIngAQtf91+m+/i1tHOaVOSsmgV86UxH6Oo2Jaq628ealtxHZnZQ04nre
# dP0cPbEww0F0/WHSHMQ3Dlj/6SrDbHAZXMgASGH8BHRf2eE5Z4NQIn7bMkm8Fq5Z
# qH8v5fArWmwG54Aw4CiWSciGlb7lVdq1aQ/D+nnNvwjdcj6dp2lp2DzSOO/CWV/j
# tApLgC60MlJSJFSweEnCM+RMhtihBktchk7CltLY0yP3II4radNTiKGCAjAwggIs
# BgkqhkiG9w0BCQYxggIdMIICGQIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYD
# VQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgVGltZXN0YW1waW5nIENBAhAN
# QkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsG
# CSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjEwOTI0MDUxMDMxWjAvBgkqhkiG
# 9w0BCQQxIgQgGB8gQc6xmUUrHNU0KzRHYPAtHnIZNO0cYumcWSdPH7owDQYJKoZI
# hvcNAQEBBQAEggEArj5RXyATPgtUVvG0QkFohrMqv4kGa2YA9F/6CFrtBVLdTu66
# JDEbW2LLIqAjF8Y1bMAP0QnXDojErvhuRwekzIKoGBWCEsGgYVRcuOpzcVXI8SqZ
# rtKBpQTPBQQXmTmHGYtgxAnt5Jk1vNb8nkN8e1y7lPYM5mh0d5BAXiKoaCMIz2Qa
# 8T5B9BMobKpP3V50FYs3G60mSA0kyAFWOFzzEHgrMXPAdRPMH+pOH28bnhoGrwSD
# edKFo5TtkLBdFAGQr6jtLv7xfn2O0aQJVgQtGMnznT/133Xnu8FF0/gZ5L+xR6nN
# tqOdrn1m3AvgcDI1QIc1iFnZUkRHjBecjak6qA==
# SIG # End signature block
