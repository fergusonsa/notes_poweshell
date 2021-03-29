
$VerbosePreference = "Continue"

if (-Not $NOTES_FILES_DIRECTORY) {
    New-Variable -Name 'NOTES_FILES_DIRECTORY' -Value 'C:\Users\tgf218\Documents\notes' -Option Constant
}
if (-Not (Test-Path $NOTES_FILES_DIRECTORY -PathType Container)) {
    New-Item -Path $NOTES_FILES_DIRECTORY -ItemType Directory
} 

if (-Not $REPORTS_DIRECTORY) {
    New-Variable -Name 'REPORTS_DIRECTORY' -Value 'C:\Users\tgf218\Documents\reports' -Option Constant
}
if (-Not (Test-Path $REPORTS_DIRECTORY -PathType Container)) {
    New-Item -Path $REPORTS_DIRECTORY -ItemType Directory
} 

function Setup-TranscriptLogging {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true)]
        [string]$baseName
    )
    if (-not $baseName) {
        $baseName = "powershell_ise"
    } 
    $VerbosePreference = "Continue"
    $LogPath = "C:\Users\tgf218\reports\"
    $LogPathName = Join-Path -Path $LogPath -ChildPath "$baseName-$(Get-Date -Format 'yyyyMMdd_hhmmss').log"
    Start-Transcript $LogPathName -Append
}

function Start-RequiredApplications{
    process {
        # Check for required applications running
        $AppExecutables = @{'notepad++' = 'C:\Program Files (x86)\Don_Ho_Notepad++_760\notepad++.exe'; # Notepad++
                            'Teams' = 'C:\Users\tgf218\AppData\Local\Microsoft\Teams\Update.exe --processStart "Teams.exe"' # MS Teams
                            }

        foreach ($process_name in $AppExecutables.Keys) {
            if (-not (get-process | ?{$_.name -eq $process_name})) {
                Write-Verbose -Message "Starting: $process_name"
                Start-Process -FilePath $AppExecutables[$process_name]
            } else {
                Write-Verbose -Message "Already started: $process_name"
            }
        }
        # MS Outlook
        if (-not (get-process | ?{$_.name -eq 'OUTLOOK'})) {
            Write-Verbose -Message "Starting: Outlook"
            start outlook
        } else {
            Write-Verbose -Message "Already started: Outlook"
        }
        
        # Browser
        # if (-not (get-process | ?{$_.name -eq 'MicrosoftEdge'})) {
        #     Write-Verbose -Message "Starting: microsoft-edge"
        #     Start-Process microsoft-edge:
        # } else {
        #     Write-Verbose -Message "Already started: microsoft-edge"
        # }
    }
}

function Check-VPN {
    process {
        $vpnCheck = [bool](Get-WmiObject -Query "Select Name,NetEnabled from Win32_NetworkAdapter where (Name like '%AnyConnect%' or Name like '%Juniper%' or Name like '%VPN%') and NetEnabled='True'")
        if ($vpnCheck) {
            Write-Verbose -Message "VPN is connected!"
            Add-Note -Note "VPN is connected!"
        } else {
            Write-Warning -Message "VPN is NOT connected!"
            Add-Note -Note "VPN is NOT connected!"
        }
    }
}

function Start-Day {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]$Date
    )
    process {
        if (-Not $Date) {
            $Date = get-date
        }
        $FileName = $Date.ToShortDateString() + '.txt' 
        $FilePath = Join-path $NOTES_FILES_DIRECTORY $FileName

        if (-Not (Test-Path -Path $FilePath -pathtype leaf)) {
            Set-Content -Path $FilePath -Value ($Date.ToShortDateString() + ' ' + $Date.ToShortTimeString() + '  Logged in from home') 
        }

        try{
          Stop-Transcript 
        }
        catch [System.InvalidOperationException]{}

        Setup-TranscriptLogging "powershell_ise"

        Start-RequiredApplications
        Check-VPN


        $FilePath
    }
}

function Open-Notes {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]$Date,
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 1)]
        [string]$FilePath,
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 2)]
        [Switch]$Yesterday
    )
    process {
        if ($Yesterday) {
            $Date = (get-Date).AddDays(-1)
        }
        if (-not $FilePath) {
            if (-Not $Date) {
                $Date = get-date
            }
            do {

                $FileName = $Date.ToShortDateString() + '.txt' 
                $FilePath = Join-path $NOTES_FILES_DIRECTORY $FileName
                $Date = $Date.AddDays(-1)
            } until ((test-path $FilePath) -or ($Date.Year -lt 2020))
        }
        if (Test-Path -Path $FilePath -pathtype leaf) {
            start 'C:\Program Files (x86)\Don_Ho_Notepad++_760\notepad++.exe' -ArgumentList $FilePath
        }

        $FilePath
    }
}

function Add-Note {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [string]$Note,

        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 1)]
        [datetime]$Date
    )
    process {
        if (-Not $Date) {
            $Date = get-date
        }
        $FileName = $Date.ToShortDateString() + '.txt' 
        $FilePath = Join-path $NOTES_FILES_DIRECTORY $FileName

        if (-Not (Test-Path -Path $FilePath -pathtype leaf)) {
            Set-Content -Path $FilePath -Value ($Date.ToShortDateString() + ' ' + $Date.ToShortTimeString() + '  ' + $Note)
        } else {
            add-content -Path $filePath -Value "$($Date.ToShortDateString()) $($Date.ToShortTimeString())  $Note" 
        }
        $FilePath
    }
}

function Start-Lunch {
    Start-OutOfOffice -Note "Starting lunch"
}

function End-Lunch {
    Add-Note -Note "Back from lunch"
}


function Start-OutOfOffice {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [string]$Note,

        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 1)]
        [datetime]$Date
    )
    process {
        if (-Not $Date) {
            $Date = get-date
        }
        if ($Note) {
            $Note = "Out of office - $Note"
        } else {
            $Note = "Out of office"
        }
        Add-Note -Note $Note -Date $Date
    }
}

function End-Day {
   [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]$Date
    )
    process {
        Add-Note -Note 'Finished for the day'

        try {
          Stop-Transcript 
        }
        catch [System.InvalidOperationException]{}
    }
}


function Get-NotesFilePathForDate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]$Date
    )
    process {
        if (-Not $Date) {
            $Date = get-date
        }
        $FileName = $Date.ToShortDateString() + '.txt' 
        $FilePath = Join-path $NOTES_FILES_DIRECTORY $FileName
        if (Test-Path -Path $FilePath -pathtype leaf) {
            $FilePath
        } else {
            $null
        }
    }
}

function Get-DayEntries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [string]$FilePath
    )
    process {
        $timestampLinePattern = "^(?<timestamp>\d{4}-\d{2}-\d{1,2} \d{1,2}:\d{2}(:\d{2})?(\s[AP]M)?\s*)?(?<line>.+)$"

        $entries = @{}
        $current_entry = $null
        $last_timestamp = $null
        ForEach($line in (get-content -path $FilePath)) {
            try {
                if ($line -eq '') {
                    if ($last_timestamp) {
                        $entries[$last_timestamp]  += $line
                    }
                } else {
                    $matches = $line | Select-String -Pattern $timestampLinePattern
                    $timestampString = $matches.Matches[0].Groups['timestamp'].Value
                    $line_content = $matches.Matches[0].Groups['line'].Value
           
                    if ($timestampString) {
                        $last_timestamp = [datetime] $timestampString
                        $entries[$last_timestamp] = @($line_content)
                    } else {                
                        $entries[$last_timestamp]  += $line_content
                        # Write-Verbose -Message "Added to existing entry '$line_content'"
                    }
                }
            } catch {
                Write-Warning "Unable to parse the line '$line'"
            }
        }

        $entries
    }
}


function Get-DayTimes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]
        $Date
    )
    process {
        $FilePath = Get-NotesFilePathForDate -Date $Date
        if ($FilePath) {
            $entries = Get-DayEntries -FilePath $FilePath 
            $sorted_entries = $entries.getEnumerator() | Sort-Object Name

            $periods = @{}
            $start_timestamp = $null
            
            foreach($e in $sorted_entries) {
                if (-not $start_timestamp -and $e.Name) {
                    $start_timestamp = $e.Name
                } else {
                    $last_timestamp = $e.Name
                }
                if ($e.Value[0] -like 'out of office*') {
                    $periods[$start_timestamp] = $last_timestamp 
                    $start_timestamp = $null
                    $last_timestamp = $null
                }
            }

            $current_timestamp = Get-Date
            if ($last_timestamp -eq $null -or ($last_timestamp.Date -eq $current_timestamp.Date -and $current_timestamp -gt $last_timestamp)) {
                $last_timestamp = $current_timestamp
            }
            if ($start_timestamp -and $last_timestamp) {
                $periods[$start_timestamp] = $last_timestamp
            }
            $start_timestamp = [datetime] ($periods.Keys | Measure-Object -Minimum).Minimum
            $last_timestamp = [datetime] ($periods.Values | Measure-Object -Maximum).Maximum

            # Write-Verbose -Message "Start time: $start_timestamp  End time: $last_timestamp"
            #foreach($periodKey in $periods.Keys) {
            #    Write-Verbose -Message "Period: $periodKey - $periods[$periodKey]"                
            #}
            # Write-Verbose -Message "Periods: $periods"
            Select-Object @{n='date'; e={$date.Date}}, @{n='start'; e={$start_timestamp}}, @{n='end'; e={$last_timestamp}}, @{n='hours'; e={($periods.Keys | % { $periods[$_] - $_ } | Measure-Object -sum -property TotalHours).Sum}} -InputObject ''
        }
    }
}

function Get-WeekTimes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)]
        [datetime]
        $WeekStartDate = (Get-Date)
    )
    process {
        # Find the week containing the date, Sat to Fri
        while ($WeekStartDate.DayOfWeek -ne 'Saturday') {$WeekStartDate = $WeekStartDate.AddDays(-1)}

        $DaysTimes = @{}
        0..6 | %{
            $changingDate = $WeekStartDate.AddDays($_)
            $DaysTimes[$changingDate.Date] = Get-DayTimes $changingDate
        }

        $DaysTimes
    }
}

function Get-Times {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [datetime] $StartDate = (Get-Date),
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ValidateSet('day', 'week', 'month')]
        [string] $Period = 'week'
    )
    process {
        switch($Period) {

            'day' {
                $numDays = 1
            } 
            'week' {
                $numDays = 7
                # Find the week containing the date, Sat to Fri
                while ($StartDate.DayOfWeek -ne 'Saturday') {$StartDate = $StartDate.AddDays(-1)}
            } 
            'month' {
                # Find the First day of the month
                $StartDate = $StartDate.AddDays(- $StartDate.Day + 1)
                $numDays = [datetime]::DaysInMonth($StartDate.Year,$StartDate.Month)
            } 
        }
        $DaysTimes = @{}
        0..($numDays - 1) | %{
            $changingDate = $StartDate.AddDays($_)
            $DaysTimes[$changingDate.Date] = Get-DayTimes $changingDate
        }

        $DaysTimes
    }
}


function Show-Times{ 
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [datetime] $WeekStartDate = (Get-Date), 
        [Parameter(Mandatory= $false, ValueFromPipelineByPropertyName = $true, Position = 2)]
        [Switch] $LastWeek, 
        [Parameter(Mandatory= $false, ValueFromPipelineByPropertyName = $true, Position = 3)]
        [Switch] $LastMonth

    )
    process {
        $period = 'week'
        if ($LastWeek) {
            $WeekStartDate = (Get-Date).AddDays(-7)
        } elseif ($LastMonth) {
            $WeekStartDate = (Get-Date).AddDays(- (Get-Date).Day)
            $period = 'month'
        }
        $DaysTimes = Get-Times -StartDate $WeekStartDate -Period $period
        $DaysTimes.getEnumerator() | Sort-Object -Property Name | Format-Table @{Label="Date"; Expression={($_.Name).ToShortDateString()}}, @{Label="Start"; Expression={(($_.Value).start).ToShortTimeString()}}, @{Label="End"; Expression={(($_.Value).end).ToShortTimeString()}}, @{Label="Hours"; Expression={($_.Value).hours.ToString("#.##")}} 
        
        $measurement = $DaysTimes.Keys | % {$DaysTimes[$_].hours} | Measure-Object -Sum
        $sumHours = $measurement.Sum
        $expected = $measurement.Count * 8.25 # 7.5 hours plus 0.75 hour lunch, this assumes that lunches are not marked as 'out of office'
        Write-Output "Total hours for the $period : $($sumHours.ToString("#.##"))  Number Days: $($measurement.Count)   Expected hours: $($expected.ToString("#.##"))"
    }
}
    
function Get-InternetExplorerHistory {

    #https://crucialsecurityblog.harris.com/2011/03/14/typedurls-part-1/
    # Based on https://github.com/rvrsh3ll/Misc-Powershell-Scripts/blob/master/Get-BrowserData.ps1


    $Null = New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS
    $Paths = Get-ChildItem 'HKU:\' -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-1-5-21-[0-9]+-[0-9]+-[0-9]+-[0-9]+$' }

    ForEach($Path in $Paths) {

        $User = ([System.Security.Principal.SecurityIdentifier] $Path.PSChildName).Translate( [System.Security.Principal.NTAccount]) | Select -ExpandProperty Value

        $Path = $Path | Select-Object -ExpandProperty PSPath

        $UserPath = "$Path\Software\Microsoft\Internet Explorer\TypedURLs"
        if (-not (Test-Path -Path $UserPath)) {
            Write-Verbose "[!] Could not find IE History for SID: $Path"
        } else {

            Get-Item -Path $UserPath -ErrorAction SilentlyContinue | ForEach-Object {
                $Key = $_
                $Key.GetValueNames() | ForEach-Object {
                    $Value = $Key.GetValue($_)
                    if ($Value -match $Search) {
                        New-Object -TypeName PSObject -Property @{
                            User = $UserName
                            Browser = 'IE'
                            DataType = 'History'
                            Data = $Value
                        }
                    }
                }
            }
        }
    }
}


