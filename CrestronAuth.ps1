<#

.DESCRIPTION
   Setup Authenticaiton on all Crestron Devices

.INPUTS
    Complete CSV with all Crestron IP Addresses

.OUTPUTS
    Password Change Results Log.csv

.NOTES
    Version: 1.0
    Creation Date: 2/1/2022
    Purpose/Change: Initial Release
#>

Write-Host -ForegroundColor Cyan @"

Crestron Authentication Setup

"@

# Global Variables

$CSV_Path = "KingCountyData.csv"
$User = "ftr_admin"
$Pass = "Fortherecord123!"
$Path = Get-Location

# Import Libraries
Import-Module PSCrestron

$passchanged = " "

# Stopwatch Feature
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Object for Results Data

$DeviceResultsData =@()

# Initilize Table

$DeviceResultsData | Out-GridView -Title "Device Status Results"

# Clear Device Error Court

$devicerror = " "

# Delete Old Results

Remove-Item -Path  "$Path\Password Change Results.csv" -ErrorAction SilentlyContinue 


function CrestronAuth([string]$IP)
{
    try
    {
        # Clear Connect Error
        $ConnectError = " "
        
        # New Data Object
        $DeviceResultItem = New-Object PSObject
        Write-Host -f Green "Auth Setup for: $IP"
        
        # Connect to Device via SSH w/ Default Credentials
        $SessionID = Open-CrestronSession -Device $IP -Secure -ErrorAction SilentlyContinue

        # Hostname
        $hostnameResponce = Invoke-CrestronSession $SessionID "hostname"
        $deviceHostname = [regex]::Match($hostnameResponce, "(?<=Host\sName:\s)[\w-]+").value
        Write-Host -f Green "Working on => $deviceHostname`n`n`n`n"

        # Set User/Password
        Invoke-CrestronSession $SessionID 'AUTH ON'
        Invoke-CrestronSession $SessionID "$User"
        Invoke-CrestronSession $SessionID "$Pass"
        Invoke-CrestronSession $SessionID "$Pass"
        
        #Reboot to Confirm Changes
        Invoke-CrestronSession $SessionID 'reboot'
        Close-CrestronSession $SessionID
        
        Write-host -f Green "$deviceHostname : Password Successfully Changed`n`n`n`n"
        $AuthMethod = 'Custom Password [New]'
    }
    catch
    {
        Write-host -f Yellow "`n-Default Password Unsuccessful`n"
    
        # Test for Custom Password
        Try
        {
            # Connect via SSH with New Credentials
            $SessionID = Open-CrestronSession -Device $IP -Secure -Username $User -Password $Pass -ErrorAction Continue

            # Get Device Hostname
            $hostnameResponce = Invoke-CrestronSession $SessionID "hostname"
            $deviceHostname = [regex]::Match($hostnameResponce, "(?<=Host\sName:\s)[\w-]+").value

            Write-Host -f Green "$deviceHostname ($d) - Password is already set`n`n`n`n"

            Close-CrestronSession $SessionID
            $AuthMethod = 'Custom Password'
        }

        #Catch for error connecting after default/Custom password
        catch 
        {
            $deviceHostname =" "
            $ConnectError = "Connection Attempts Unsuccessful"
            Write-Host -f Red "`n $d - Default & Custom Password Attempts Unsuccessful: Could NOT Connect!`n`n`n`n"
            $AuthMethod = 'Unknown Password/Error'
        }

    }

    #Current Date/Time
    $time = (get-date)
    #Build Table
    # Table Coulumn 1 - Time
    $DeviceResultItem | Add-Member -Name "Time" -MemberType NoteProperty -Value $time
    # Table Coulumn 2 - IP Address
    $DeviceResultItem | Add-Member -Name "IP Address" -MemberType NoteProperty -Value $IP
    # Table Coulumn 3 - Hostname
    $DeviceResultItem | Add-Member -Name "Hostname" -MemberType NoteProperty -Value $deviceHostname
    # Table Coulumn 4 - Authentication
    $DeviceResultItem | Add-Member -Name "Authentication" -MemberType NoteProperty -Value $AuthMethod
    # Table Coulumn 5 - Error
    $DeviceResultItem | Add-Member -Name "Error" -MemberType NoteProperty -Value $ConnectError
    #Add line to the report
    $DeviceResultsData += $DeviceResultItem
 
    #Append results to Password Change Results Document + Log 
    $DeviceResultsData | Export-Csv -Path "$Path\Password Change Results.csv" -NoTypeInformation -append
    $DeviceResultsData | Export-Csv -Path "$Path\Password Change Results Log.csv" -NoTypeInformation -append

    #Total time of script
    $stopwatch
        
}

class Courtroom
{
    [string]$RoomName
    [string]$ProcIP
    [string[]]$PanelIP

    [void] RunPanels()
    {
        foreach ($Panel in $this.PanelIP)
        {
            CrestronAuth $Panel
        }
    }
    [void]  RunProcessors()
    {  
        CrestronAuth $this.ProcIP
    }
}

# Import Device Data

try 
{
    $Data = Import-csv $CSV_Path
    Write-Host " "
}

catch
{
    Write-Host 'Error obtaining device data! Assure Data CSV is in same directory as script'
}

foreach($row in $Data)
{
    $c = new-object -TypeName Courtroom

    $c.RoomName = $row | Select-object -ExpandProperty RoomName

    $c.ProcIP = $row | Select-object -ExpandProperty ProcIP
    
    $c.PanelIP = $row | Select-object -ExpandProperty PanelIP
    $c.PanelIP = $c.PanelIP[0].split('~')

    
    # Run Functions
    foreach ($Panel in $c.PanelIP)
    {
        try
        {
            Write-Host -f Green "Starting Panel Auth"
            CrestronAuth $Panel
        }
        catch
        {
            Write-Host -f Red "Unable to Connect to Panel"
        }
    }
    
    try
    {
        Write-Host -f Green "Starting Proccessor Auth"
        CrestronAuth $c.ProcIP
    }
        catch
    {
        Write-Host -f Red "Unable to Connect to Processor"
    }
}