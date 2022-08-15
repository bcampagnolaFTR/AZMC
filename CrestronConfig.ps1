<#
    FTR Crestron PowerShell Script

    ver 1.whatever   


    Notes:
    In general in this script, bool values are positive indicators.
    e.g.    For errors / faults, the presence of an error or fault == $true
            For validity flags, such as  $isValidIPAddress, the bool positively indicates the validity ($true if valid)
#>

#####################################################################################################################
<#    Globals    #>
#####################################################################################################################
#region

# Debug
$global:DEBUG = $true

# This value needs to be copied while the script is running. It doesn't work if you run the script, and then try to reference the value from console
$global:scriptPath = $PSScriptRoot + '\'

# Log Files
$global:logFilePath = ("{0}Logs\" -f $global:scriptPath)
$global:logFilePathName = ""
$global:logFileName = ""

$global:logLevel = 1

# Credentials for Crestron device auth
$global:User = "ftr_admin"
$global:Pass = "Fortherecord123!"

# Spreadsheet Data
$global:SelectedFile = ""
$global:FileLoaded = $false

$global:sheet = ""
$global:sheetNumOfLines = 0

$global:rooms = @{}
$global:numOfRooms = 0

$global:defaults = @{}
$global:numOfDefaults = 0

$global:menuFunctions = @{}


#endregion


#####################################################################################################################
<#    Utils    #>
#####################################################################################################################
#region

# Connection Data Validation
#####################################################################################################################

function fValidateIPAddress ($s)
{
    return ($s -match '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')
}

function fValidateHostname ($s)
{
    if ($s.length -le 15)
    {
        return ($s -match '^[a-z0-9-]+$')
    }
    return $false
}

function fValidateGuid ($g)
{
    return ($g -match("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"))
}

function fValidateConnection ($t)
{
    if (fValidateIPAddress ($t))
    {
        return "ipaddress", "Invoke-CrestronCommand", "-device" 
    }
    elseif (fValidateGuid ($t))
    {
        return "guid", "Invoke-CrestronSession"
    }
    elseif (fValidateHostname ($t))
    {
        return "hostname", "Invoke-CrestronCommand", "-device"
    }
    else
    {
        return "none"
    }
}

function fGetCredsParams ([ref]$r)
{
    if($r.value.usesCreds)
    {
        return ("-username `"{0}`" -password `"{1}`"" -f $r.value.username, $r.value.password)
    }
    return ""
}

function fIPAddr([string]$a, [string]$b)
{
    # if $a (subnet) has any string value, then use it
    if($a)
    {
        $ipa = ("{0}.{1}" -f $a, $b)
    }
    # else if $a is empty, just return $b (else we will end up sending back (".{0}" -f $b)
    else
    {
        return $b
    }
    
    while($ipa -contains '..')
    {
        $ipa = $ipa.replace('..', '.')
    }
    return $ipa
}

function fDefault($room, $default)
{
    # this first condition is special behavior that allows the user to enter "0" in the spreadsheet to deliberately omit a value
    # use case e.g.: if widely using a default line, but a few select rooms don't have a DVD player, then those room data line values
    #     can receive "0" to override the default line.
    # Why not just leave blank? Because default values will override blank values. "0" actively overrides the default value with emptystring.
    if($room -ieq "0") 
    { 
        return "" 
    }
    elseif($room)
    {
        return $room
    }
    return $default
}


# Printing
#####################################################################################################################

function fPrettyPrintRooms([int]$h=20, [int]$w=15)
{
    $a = $global:rooms.keys | sort
    $col = [math]::ceiling($a.count / $h)
    $b = $a | ForEach-Object { ("{0:d3}. {1}"-f $_, $global:rooms[[int]$_].roomname).PadRight($w) }

    $b | Format-Wide {$_} -Force -Column $col 
}

function showAllRooms
{
    $e = $global:rooms.keys | sort

    foreach($r in $e)
    {
        write-host -b black -f green ("{0:d3}. {1}" -f $r, $global:rooms[$r].RoomName) 
    }
}

function fClear($ms = 100)
{
    clear
    start-sleep -Milliseconds 100
}

fClear


# Edit Text
#####################################################################################################################

function fRemoveWhitespace([string]$s)
{
    return ($s -replace ' ','')
}


# Script Management
#####################################################################################################################

function fatal
{
    Read-Host -Prompt "`n`nPress any key to exit"
    exit
}


# File Management
#####################################################################################################################

function fCheckFileExists($fileName)
{
    if([System.IO.File]::Exists($fileName))
    {
        return $true
    }
    return $false
}

function getFileName
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $global:ScriptPath }
    $null = $FileBrowser.ShowDialog()
    $fileName = $FileBrowser | Select-Object -ExpandProperty FileName
    $global:SelectedFile = $fileName
    return ($fileName)    
}

function importCSVFile([string]$fileName)
{
    if(-not (fCheckFileExists $fileName))
    {
        fErr ("Import: File import failed. File does not exist.`n    {0}" -f $fileName) $true
        return $null
    }
    elseif($fileName -notmatch ".csv$")
    {
        fErr ("Import: File import failed. This script only uses .csv files.`n    File selected  -  {0}" -f $fileName) $true
        return $null
    }
    try
    {
        $sheet = Import-csv $fileName
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count
        fErr ("Import: File opened successfully.`n    {0}" -f $fileName) $false
        fErr ("Import: Found {0} data lines to parse." -f $sheetLen) $false
        return $sheet
    }
    catch
    {
        fErr ("Import: File import failed.`n    {0}" -f $error) $true
        return $null
    }
}


# User Entry
#####################################################################################################################


function decodeRange([string]$s)
{
    $s = fRemoveWhitespace($s)
    
    $badChars = $s -replace "[0-9-,]",""
    if(-not [bool]$badChars)
    {
        $r = [int[]]@()
        $segments = $s.split(',')

        foreach($segment in $segments)
        {
            $v = [int[]]$segment.split('-')
            if($v.length -eq 2)
            {
                $r += $v[0]..$v[1]                    
            }
            elseif($v.length -eq 1)
            {
                $r += $v[0]
            }
            else
            {
                fErr ("DecodeRange: Syntax error in selection range. This is invalid: '{0}'" -f $segment) $true
                return $null
            }
        }

        return $r
    }
    else
    {
        fErr ("DecodeRange: User entered invalid characters '{0}'" -f $badChars) $true
        return $null
    }
}


# Time
#############################################################
function fGetDateTime
{
    return ("{0:yyyy.MM.dd_HH.mm.ss.fff}" -f (Get-Date))
}

#endregion


#####################################################################################################################
<#    Logging & Debug    #>
#####################################################################################################################
#region 

# Error
########################################################################
function fErr ($s, $c)
{
    if($global:logLevel -gt 0)
    {
        fLog $s
    }

    if($global:debug)
    {
        if($c -eq $true)
        {
            Write-Host -b black -f red $s
        }
        else
        {
            Write-Host -b black -f green $s
        }
    }
}

function err([string]$err)
{
    Write-Host $err
}
function err([int]$i)
{
    switch($i)
    {
        1 
        { 
            Write-Host -f Yellow -b DarkGray
            (
                "This script requires software called 'PSCrestron' (Crestron's PowerShell module).`n" +
                "Looks like you have a good internet connection, but something went wrong when I tried to download and install this software.`n" +
                
                "If you get this message again, try to manually download and install PSCrestron. Drop this into a web browser:`n" +
                "https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe`n" +
                "... and run the file, following any prompts to install the software.`n`n" +

                "If you're still having problems, you can email bcampagnola@fortherecord.com, with 'HALP!' in the subject line.`n`n" 
            ) 
            fatal
        } 
        2 
        {
            Write-Host -f Yellow -b DarkGray
            (
                "Hey!`n" +
                "This script requires software called 'PSCrestron' (Crestron's PowerShell module).`n" +
                "I attempted to download the software, but it appears that you don't have an internet connection right now.`n`n" +

                "Please connect to the web and then run this script again.`n`n" +

                "If you are connected to the web and still got this message, the following may be amiss:`n" + 
                "- there might be something wrong with the permissions on your machine`n" +
                "- Crestron's web service may be down`n" +
                "- you might be suffering from delusions`n`n`n" +


                "Can't figure it out? Email bcampagnola@fortherecord.com, with 'HALP!' in the subject line." 
            )        
            fatal        
        }
    }
}

function fFullError ($e)
{
    write-host -f cyan -b black InvocationInfo.MyCommand.Name = 
    write-host -f yellow $e.InvocationInfo.MyCommand.Name
    write-host -f cyan -b black ErrorDetails.Message =
    write-host -f yellow $e.ErrorDetails.Message
    #write-host -f cyan -b black InvocationInfo.PositionMessage = 
    #write-host -f yellow $e.InvocationInfo.PositionMessage
    write-host -f cyan -b black CategoryInfo.ToString = 
    write-host -f yellow $e.CategoryInfo.ToString()
    write-host -f cyan -b black FullyQualifiedErrorId =
    write-host -f yellow $e.FullyQualifiedErrorId
}


# Log File
########################################################################

function fLog ([string]$s, $i)
{
    if(fCheckFileExists $global:logFilePathName)
    {
        $dt = fGetDateTime
        $s = "{0}  -  {1}" -f [string]$dt, $s
        try
        {
            $s >> ("{0}{1}"-f $global:logFilePath, $global:logFileName)
        }
        catch
        {
            Write-Host -b black -f red "Logs: Failed to write error to log."
        }
    }
    else
    {
        Write-Host -f red -b black "Logs: Log file does not exist. Attempting to create a new file."
        if($i -lt 3)
        {
            fUpdateLogFileName
            fCreateLogPath
            fCreateLogFile
            $j = $i + 1
            fLog $s $i
        }
        else
        {
            Write-Host -f red -b black ("Logs: Failed to create a log file.`n    Most recent error: {0}" -f $error[0])
        }
    }
}

function fUpdateLogFileName
{
    $global:logFileName = fGetDateTime + ".log"
    $global:logFilePathName = $global:logFilePath + $global:logFileName
}

function fCreateLogPath
{
    if(-not(Test-Path -path $global:logFilePath))
    {
        try
        {
            New-Item -Path $global:scriptPath -Name "Logs\" -ItemType "directory" | out-null
        }
        catch
        {
            fErr ("Logs: Failed to make directory '\Logs'") $true
        }
    }
}

function fCreateLogFile
{
    try
    {
        New-Item -Path $global:logFilePath -ItemType "file" -Name $global:logFileName | out-null

        fErr ("Logs: Log file created.`n    {0}" -f $global:logFilePathName) $false
    }
    catch
    {
        fErr ("Logs: Failed to create log file.`n    {0}" -f $global:logFilePathName) $true        
    }
}

fUpdateLogFileName
fCreateLogPath
fCreateLogFile


#endregion


#####################################################################################################################
<#    Import Libs    #>
#####################################################################################################################
#region

Add-Type -AssemblyName System.Windows.Forms

function tryImportCrestron
{
    try
    {
        Import-Module PSCrestron -ErrorAction Stop
        fErr ("Init: Successfully imported PSCrestron module.") $false
        return 0
    }
    catch 
    { 
        fErr ("Init: Failed to import PSCrestron module.") $false
        return 1 
    }
}
   

$src = ""
$dst = ""
$err = tryImportCrestron

if($err -ne 0)
{
    $src = 'https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe'
    $dst = $env:temp + '\EDK_Setup_1.0.5.3.exe'

    if(Test-Connection "8.8.8.8" -Count 1 -Quiet)
    {
        Invoke-RestMethod $src -OutFile $dst -TimeoutSec 8
        Start-Process -Wait -FilePath "$dst" -ArgumentList "/S" -PassThru
    }
    else
    {
        err 2
    }

    $err = tryImportCrestron
    if($err -ne 0)
    {
        err 1
    }
}


function tryImportPoshSSH
{
    try
    {
        Import-Module -name Posh-SSH -ErrorAction Stop
        fErr ("Init: Successfully imported Posh-SSH module.") $false
        return 0
    }
    catch 
    { 
        fErr ("Init: Failed to import PSCrestron module.") $false
        return 1 
    }
}
   

$err = tryImportPoshSSH



#endregion


#####################################################################################################################
<#    Data & File Import    #>
#####################################################################################################################
#region

class Device
{
    [bool]$deviceUsable

    [string]$ipAddr
    [bool]$isIPValid
    
    [string]$fileName
    [bool]$isFileNameValid

    [int]$IPID

    Device ($ip_="")
    {
        #$this.setIPAddr $ip_
        $this.ipAddr = $ip_
        $this.isIPValid = fValidateIPAddress $this.ipAddr
    }    
    [void] setIPAddr ($ip_)
    {
        $this.ipAddr = $ip_
        $this.isIPValid = fValidateIPAddress $this.ipAddr
    }
    [void] setFileName ($fileName_)
    {
        $this.fileName = $fileName_
        $this.isFileNameValid = ($this.fileName -match ".lpz$") -or ($this.fileName -match ".cpz$") -or ($this.fileName -match ".vtz$")
    }
}

class Courtroom
{
    $index = 0
    $roomName = ""
    $facilityName = ""

    $isDef = $false
    $defIndex = 0

    $usesCreds = $false
    $username = ""
    $password = ""

    $subnet = ""

    $isVC4_System = $false
    $Processor
    $Panels = @()

    $ReporterWebSvc
    $Wyrestorm
    $Fixed_Cams = @()
    $DSPs = @()
    $RecorderSvrs = @()
    $DVD_Player 
    $Mute_GW
    $PTZ_Cams = @()

    $IPTable = @{}
    $data = @{}
    $ability = @{}

    Courtroom ($data_, [int]$i_)
    {
        $this.data = $data_
        $this.index = $i_

<#        foreach($f in $global:menuFunctions)
        {
            $this.ability[$f.key] = $false
        }
#>    }
<#
    [void] checkAbilities ()
    {
        if($this.Processor.isIPValid)
        {
            $this.Processor.deviceUsable = $true
            $this.ability[3] = $true

            if($this.Processor.isFileNameValid)
            {
                $this.ability[2] = $true
            }
        }
        foreach($pnl in $this.panels)
        {
            if($pnl.isIPValid)
            {
                $this.ability[5] = $true
                $pnl.deviceUsable = $true

                if($pnl.isFileNameValid)
                {
                    $this.ability[4] = $true
                }
            }           
        }

    }
    #>
}

function fParseLine($c)
{
    $row = $c.data         
    try
    {   
        $d = $global:defaults[$c.defIndex].data

        # Credentials
        $c.username, $c.password = (fDefault $row.Credentials $d.Credentials) -split ":"
        if([bool]$c.username -or [bool]$c.password)
        {
            $c.usesCreds = $true
        }

        # Subnet
        $c.subnet = fDefault $row.Subnet_Address $d.Subnet_Address

        # Processor
        $c.isVC4_System = [bool](fDefault $row.isVC4_System $d.isVC4_System)
        $ip = fDefault $row.Processor_IP $d.Processor_IP
        $c.Processor = [Device]::new((fIPAddr $c.subnet $ip))
        $c.Processor.setFileName((fDefault $row.FileName_LPZ $d.FileName_LPZ))

        # Panels
        $ips = (fDefault $row.Panel_IP $d.Panel_IP) -split "~"
        $files = (fDefault $row.FileName_VTZ $d.fileName_VTZ) -split "~"
        
        for($i = 0; $i -lt $ips.length; $i++)
        {
            $ip = fIPAddr $c.subnet $ips[$i]
            $pnl = [Device]::new($ip)

            # if there are as many .vtz file names as panel IPs, then assign each panel its own file name
            if($ips.length -eq $files.length) {    $pnl.setFileName($files[$i])    }
            # else all panels get the first (and probably the only) file name in $files
            else           {    $pnl.setFileName($files[0] )    }

            $c.panels += $pnl
        }

        # Reporter Web Service
        $ip = fDefault $row.IP_ReporterWebSvc $d.IP_ReporterWebSvc
        $c.ReporterWebSvc = [Device]::new((fIPAddr $c.subnet $ip))

        # Wyrestorm
        $ip = fDefault $row.IP_WyrestormCtrl $d.IP_WyrestormCtrl
        $c.Wyrestorm = [Device]::new((fIPAddr $c.subnet $ip))

        # Fixed Cams
        $ips = (fDefault $row.IP_FixedCams $d.IP_FixedCams) -split "~"
        foreach($ip in $ips)
        {
            $c.Fixed_Cams += [Device]::new((fIPAddr $c.subnet $ip))
        }

        # DSPs
        $ips = (fDefault $row.IP_DSPs $d.IP_DSPs) -split "~"
        foreach($ip in $ips)
        {
            $c.DSPs += [Device]::new((fIPAddr $c.subnet $ip))
        }
        
        # Recorder Servers
        $ips = (fDefault $row.IP_Recorders $d.IP_Recorders) -split "~"
        foreach($ip in $ips)
        {
            $c.RecorderSvrs += [Device]::new((fIPAddr $c.subnet $ip))
        }

        # DVD Player
        $ip = fDefault $row.IP_DVDPlayer $d.IP_DVDPlayer
        $c.DVD_Player += [Device]::new((fIPAddr $c.subnet $ip))
        
        # Mute Gateway
        $ip = fDefault $row.IP_AudicueGW $d.IP_AudicueGW
        $c.Mute_GW = [Device]::new((fIPAddr $c.subnet $ip))
    
        # PTZ Cams
        $ips = (fDefault $row.IP_PTZCams $d.IP_PTZCams) -split "~"
        foreach($ip in $ips)
        {
            $c.PTZ_Cams += [Device]::new((fIPAddr $c.subnet $ip))
        }
    }
    catch
    {
        fErr ("ParseLine: Failed to parse a value from data row {0}.`n {1}" -f $c.index, $error) $true
        return $true, $c
    }    
    
    return $false, $c
}

function fClearData
{
    $global:rooms = @{}
    $global:numOfRooms = 0
    $global:defaults = @{}
    $global:numOfDefaults = 0    
}

function fParseCSV($sheet)
{
    fClearData
    $i = 1
      
    foreach($row in $sheet)
    {     
        $i++

        try
        {
            $c = [Courtroom]::new($row, $i)

            # $i the room index, should align with the associated excel spreadsheet row numbers
            $c.Index = $i

            # Room_Name
            $c.RoomName = $row.Room_Name
        
            # Defaults
            $d = $row.Defaults
            $c.isDef = $d.contains("*")
            $c.defIndex = [int]($d -replace "[^0-9]", "")

            if($c.isDef)            # Default data line
            {
                if([bool]$c.defIndex)
                {
                    $global:defaults[$c.defIndex] = $c
                    $global:numOfDefaults++
                }
                else
                {
                    fErr ("Import: Line {0:d3} failed. Error in default index value." -f $i) $true
                    continue
                }
            }
            else               # Room data line
            {
                # $parseResults should be $false for success, $true or "error text" for fail
                $parseResults, $c = fParseLine $c $i

                if($parseResults)
                {
                    fErr ("DataParse: Failed to add room data from line {0}." -f $i)
                    continue
                }
                # add good room data to dict of rooms
                $global:rooms[[int]$c.index] = $c
                $global:numOfRooms++
            }
        }
        catch
        {
            fErr ("Import: File import failed for line {0:d3}.`n {1}" -f $i, $error) $true
        }
    }
    if($global:numOfRooms -gt 0)
    {
        $global:FileLoaded = $true
        fErr ("Import: File import success. Parsed {1} default lines and {0} rooms." -f $global:numOfRooms, $global:numOfDefaults) $false
    }
    else
    {
        $global:FileLoaded = $false
    }
}


#endregion


#####################################################################################################################
<#    CrestronSDK Functions    #>
#####################################################################################################################
#region

# Get Model Number of Device
# This module can take an IP address, hostname, or CrestronSession guid
function fCrestronGetModel ($target)
{
    $connection, $cmd, $params = fValidateConnection $target

    if($connection.contains("none"))
    {
        fErr ("CrestronSDK: Invalid connection data for CrestronGetModel.`n    {0}" -f $target) $true
        return     
    }

    try
    {
        Invoke-Expression ("{0} {1} `$target " -f $cmd, $params) -outvariable $response
        
        <#if ($connection.contains("ipaddr") -or $connection.contains("host"))
        {
            $response = Invoke-CrestronCommand -device $target -Command "`n"
        }
        elseif ($connection.contains("guid"))
        {
            $response = Invoke-CrestronSession $target -Command "`n"
        }
        #>
        
        $DeviceModel = ($response -replace "[^0-9]", "")
        return $DeviceModel
    }
    catch
    {
        fErr ("CrestronSDK: Failed to get Crestron model number.`n    Error: {0}`n" -f $error)
    }
}

function fCrestronStartSession($r, [string]$ipaddr)
{
    try
    {
        $creds = fGetCredsParams ([ref]$r)
        #write-host $creds -f yellow -b black

        $SessID = Open-CrestronSession -Device $ipaddr -Secure $creds     

        fErr ("PSCrestron: Open-CrestronSession for room {1:d3}- {2} successful. SessionID: {0}" -f $SessID, $r.index, $r.roomname) $False

        return $SessID
    }
    catch
    {
        fErr ("PSCrestron: Open-CrestronSession failed for room {0:d3} ({1}). `n    Connection param: {2} {3}" -f $r.index, $r.roomname, $ipaddr, $creds) $True  
        return $null
    }
}

function fCrestronRestartProg ($SessID, $target)
{
    # Reset Program
    try
    {
        Invoke-CrestronSession $SessID 'progreset -p:01'
        fErr ("PSCrestron: Restarting program for room {0:d3}." -f [int]$target) $False
    }
    catch
    {
        fErr ("PSCrestron: Program restart failed for room {0:d3}." -f [int]$target) $True
    }
}

function fCrestronCloseSession ($SessID, $target)
{
    # Close Session
    try
    {
        Close-CrestronSession $SessID
        fErr ("PSCrestron: Close-CrestronSession successful for room {0:d3}." -f [int]$target) $False
    }
    catch
    {
        fErr ("PSCrestron: Close-CrestronSession failed for room {0:d3}, `$SessID {1}." -f [int]$target, $SessID) $True
    }
}

#endregion

<#
updatepassword [currentpassword][newpassword][newpasswordagain]

[cr] (get device model)
ver
ver -v

adduser -n:ftr_admin -p:Fortherecord123!
addusertogroup -n:ftr_admin -g:Administrators


#>


<#  groups
        listgroups
        addgroup -n:groupname -l:accesslevel(A - admin; P - programmer; O - operator; U - user; C - connection only)
        adddomaingroup
        deletegroup
        deletedomaingroup

#>

<#   users
        listusers
        adduser -n:username -p:password
        addusertogroup -n:username -g:groupname
        removeuserfromgroup -n:username -g:groupname
        deleteuser username [/y]

#>

<#   blocked list
        addlockeduser username
        remlockeduser username
        addblockedip ipaddr
        remblockedip ipaddr
#>
# updatepassword  - update the current user's password
# resetpassword  - reset an existing user's password
# getpasswordrule  - display the current password rules

# setloginattempts  - blocks IP after n attempts
# setlockouttime  -  sets time an IP remains blocked
# setuserloginattempts  - blocks user after n attempts
# setuserlockouttime  - setse time a user remains blocked



# clearcsauthentication
# setcsauthentication

# err
# err plogcurrent
# err plogprevious
# clearerr


# ssl    - Display/Set SSL type
# sslverify   - Display/Set SSL certificate verification.



# SSHPORt                       Administrator       Enable/Disable and configure SSH port number
# AUTHentication                Administrator       Authentication on/off 

# USERPAGEAUTH                  Administrator       User page Authentication on/off
# TIMEZone                      Administrator       Get/Set the timezone 
# TIMEdate                      Programmer          Get the time and date
# SNTP                          Administrator       Configure network time synchronization 


<#
    SNMP                          Programmer          Enable/disable Simple Network Management Protocol
    SNMPAccess                    Programmer          Configure Access Rights for SNMP Communities
    SNMPALLOWall                  Programmer          Allows All SNMP Managers                
    SNMPCONtact                   Programmer          Configure an SNMP manager               
    SNMPLOCation                  Programmer          Configure an SNMP manager               
    SNMPMANager                   Programmer          Configure an SNMP manager 
#>



######################################################################################################
<#    User Commands    #>
######################################################################################################
#region


# User Range
######################################################################################################

function fGetRangeOfRooms
{   
    Write-host -f yellow "Which rooms do you want to target? (e.g. '3,9-12,17,16')`n"
    Write-host -f green "  *)  to target all rooms"
    write-host -f green "  b)  to go back`n"
    
    # Get input
    $input = fRemoveWhitespace(read-host " ")
    
    # Go Back
    if($input -ieq "b")
    {
        return
    }
    # Select All
    elseif($input -ieq "*")
    {
        $targets = ($global:rooms.keys | sort)
        fErr ("GetRange: User selected to target all rooms. Adding {0} rooms." -f $r.length) $false
    }
    # Parse Range
    else
    {
        $targets = decodeRange $input
    }

    $t = $targets | Where-Object -FilterScript { $global:rooms.ContainsKey($_) }
    if($t.length) { write-host -f yellow -b black ("GetRange: Found valid room selections - `n{0}" -f ($t -join ', ')) }

    return $t
}


# Show Info
######################################################################################################

function fShowInfo
{
    $s = "[no file loaded]"
    if($global:FileLoaded)
    {
        $s = $global:SelectedFile
    }
    $data = ""
    $data =  ("Info: .CSV file:  {0}`n" -f $s) 
    $data += ("Info:  Qty of rooms loaded:  {0}`n" -f $global:NumOfRooms)

    Write-host -b black -f green $data
}


# Send Processor IPT
######################################################################################################

function fSendProcIPT($SessID, [int]$ipid, $ipaddr, $target)
{
    try
    {
        $response = Invoke-CrestronSession $SessID -Command ("AddP {0:X} {1}" -f $IPID, $ipaddr)
        fErr ("ProcIPT: Successfully sent IPID {0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $false
        #write-host -b darkgray -f cyan (">>> {0}" -f $addPeer)
        #write-host -b darkgray -f yellow ("<<< {0}" -f $response)
        return $false
    }
    catch
    {
        fErr ("ProcIPT: Failed to commit IPID {0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $true         
        return $true
    }
}

function fSendProcIPT ([int[]]$targets)
{    
    foreach($target in $targets)
    {
        $r = $global:rooms[$target]

        # Connect to Device
        $SessID = fCrestronStartSession $r $r.processor.ipaddr
        if($SessID -eq $null) { continue } 



        # FTR ReporterWebSvc
        $ipid = 0x05
        $err = fSendProcIPT $SessID $ipid $r.ReporterWebSvc.ipaddr $target 
        
        # Wyrestorm Ctrl
        $ipid = 0x06
        $err = fSendProcIPT $SessID $ipid $r.wyrestorm.ipaddr $target 
        
        # Fixed Cams
        $ipid = 0x07
        for($i = 0; $i -lt $r.Fixed_Cams.length; $i++)
        {
            $err = fSendProcIPT $SessID ($ipid+$i) $r.Fixed_Cams[$i].ipAddr $target 
        }
        # DSPs
        $ipid = 0x0d
        for($i = 0; $i -lt $r.DSPs.length; $i++)
        {
            $err = fSendProcIPT $SessID ($ipid+$i) $r.DSPs[$i].ipAddr $target 
        }

        # FTR Recorders
        $ipid = 0x18
        for($i = 0; $i -lt $r.RecorderSvrs.length; $i++)
        {
            $err = fSendProcIPT $SessID ($ipid+$i) $r.RecorderSvrs[$i].ipAddr $target 
        }
        
        # DVD Player
        $ipid = 0x1a
        $err = fSendProcIPT$SessID $ipid $r.DVD_Player.ipaddr $target
        
        # Mute Gateways
        $ipid = 0x20
        $err = fSendProcIPT$SessID $ipid $r.Mute_GW.ipaddr $target 
        
        # PTZ Cams
        $ipid = 0x23
        for($i = 0; $i -lt $r.PTZ_Cams.length; $i++)
        { 
            $err = fSendProcIPT$SessID ($ipid+$i) $r.PTZ_Cams[$i].ipAddr $target 
            $ipid++
            $err = fSendProcIPT$SessID ($ipid+$i) $r.PTZ_Cams[$i].ipAddr $target 
        }

        fCrestronRestartProg $SessID $target
        fCrestronCloseSession $SessID $target
    }
}


# Send Panel IPT
######################################################################################################

function fSendPanelIPT ([int[]]$targets)
{    
    foreach($target in $targets)
    {
        $ipid = 0x03

        $r = $global:rooms[$target]

        foreach($pnl in $r.panels)
        {
            try
            {
                $creds = fGetCredsParams ([ref]$r)
                write-host $creds
                $response = Invoke-CrestronCommand -ShowProgress -Device $pnl.ipAddr -Command ("AddM {0:X} {1}" -f $ipid, $r.processor.ipAddr) -secure -port 22  -verbose -username "ftr_admin" -password "Fortherecord123!" #$creds
                fErr ("PanelIPT: Successfully sent IPID {0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $false
                #write-host -b darkgray -f cyan (">>> {0}" -f $addPeer)
                #write-host -b darkgray -f yellow ("<<< {0}" -f $response)               
            }
            catch
            {
                fErr ("PanelIPT: Failed to commit IPID {0:x2} {1} to the panel in room {2:d3}." -f $ipid, $r.processor.ipAddr, $target) $true         
                fErr ("    {0}" -f (fFullError $error[0])) $true         
                continue $true
            }
            $ipid++
        }
        return $false
    }
}


# Send .lpz File to Processors
######################################################################################################

function fGetSendIPTText ([bool]$b)
{
    if($b)
    {    return "out"    }
    return ""
}

function fSendProcLPZ ([int[]]$targets)
{
    if(-not $targets)
    {
        fErr "ProcLPZ: No rooms were targeted." $true
        return
    }

    write-host -f yellow "`nSend the default IP table?  y/n`n"
    write-host -f green "  b) to go back`n"
    $s = ""
    while(-not "ynbYNB" -contains $s)     {    $s = read-host " "    }
    if("bB".contains($s))                 {    return    }
    elseif("yY" -contains $s )            {    $sendIPT = $false    }
    elseif("nN" -contains $s )            {    $sendIPT = $true    }

    foreach($target in $targets)
    {
        $r = $global:rooms[$target]

        if(-not $r.processor.isIPValid)
        {
            fErr ("ProcLPZ: Invalid IP address for room {0}- {1}.`n    {2}" -f $target, $r.roomname, $r.processor.ipAddr)
            continue 
        }

        $lpzFile = $global:scriptPath + $r.processor.fileName

        $params = @{"device"   = $r.processor.ipAddr
                    "donotupdateIPTable"  = $sendIPT
                    "secure"=$true
                    "showprogress"=$true
                    "port"=22
                    "localfile"=$lpzFile
                    }
        if($r.usesCreds)
        {
            $params["password"] = $r.password
            $params["username"] = $r.username
        }

        # .lpz File
        try
        {
            $creds = fGetCredsParams ([ref]$r)
            Send-CrestronProgram @params
            fErr ("ProcLPZ: Successfully sent file '{0}' to room {1:d3} with{2} IP table." -f $lpzFile, $target, (fGetSendIPTText ([bool]$sendIPT))) $False   
        }
        catch
        {
            fErr ("ProcLPZ: Failed to send file`n    '{0}'`n    to room {1:d3}.`n    Error: {2}" -f $lpzFile, $target, $error[0]) $True
        }
    }
}

# Send .vtz File to Panels
######################################################################################################

function fSendPanelVTZ ($targets)
{
    if(-not $targets)
    {
        fErr "PanelVTZ: No rooms were targeted." $true
        return
    }

    write-host -f yellow "`nWhich panels do you want to load?`n"
    write-host -f green "  1) TSW-10xx (the 10`" Clerk / Judge panels)"
    write-host -f green "  2) TSW-7xx (the 7`" Counsel panels)`n"
    #write-host -f green "  3) Both`n`n"

    write-host -f green "  b) to go back`n"
    $s = ""
    while(-not "123bB" -contains $s)
    {
        $s = read-host " "
    }
    if("bB".contains($s))
    {
        return
    }
    else
    {
        $panel = [int]$s
    }

    foreach($target in [int[]]$targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr ("PanelVTZ: Target room {0:d3} is not in the list of rooms." -f $target) $True
            continue
        }

        $r = $null
        $d = $null

        $r = $global:rooms[$target]

        if($r.defIndex -gt 0)
        {
            if($global:defaults.ContainsKey($r.defIndex))
            {
                $d = $global:defaults[$r.defIndex]
            }
            else
            {
                fErr ("PanelVTZ: Attempting to use data from default line {2} with room {0} {1}. Default line does not exist." -f $target, $r.roomname, $r.defIndex) $true
            }
        }

        # ip address
        $ips = [int[]](fDefault $r.Panel_IP $d.Panel_IP)
        write-host
        if($panel -gt $ips.length)
        {
            fErr ("PanelVTZ: IP address for panel {0} does not exist.`n    Tried data from room {1} and default {2}.`n" -f $panel, $target, $r.defIndex) $true
            continue
        }
 
        $ipaddr = fIPAddr (fDefault $r.subnet $d.subnet) $ips[$panel-1]       
        if(-not (fValidateIPAddress $ipaddr))
        {
            fErr ("PanelVTZ: Invalid IP address for room {0} {1}, panel {3}.`n    {2}" -f $target, $r.roomname, $ipaddr, $panel) $true
            continue 
        }

        # .vtz File
        $vtzFiles = [string[]](fDefault $r.FileName_VTZ $d.FileName_VTZ)
        if(-not [bool]$vtzFiles)
        {
            fErr ("PanelVTZ: Neither the room data nor default data returned any .vtz file names.`n    Room {0:d3}" -f $target ) $true
            continue
        }
        if($panel -gt $vtzFiles.length)
        {
            fErr ("PanelVTZ: File name for panel {0} does not exist for room {1:d3}.`n" -f $panel, $target) $true
            continue
        }
        $vtzFile = $global:scriptPath + $vtzFiles[$panel-1]

        # send it
        try
        {
            Send-CrestronProject -ShowProgress -Device $ipaddr -LocalFile $vtzFile -Secure -password Crestron -username admin

            fErr ("PanelVTZ: Successfully sent file '{0}' to room {1:d3}, panel {2}." -f $vtzFile, $target, $panel) $false   
        }
        catch
        {
            fErr ("PanelVTZ: Failed to send file`n    '{0}'`n    to room {1:d3}, panel {3}.`n    Error: {2}" -f $vtzFile, $target, $Error, $panel) $true
        }
    }
    return
}


# Set Authentication
######################################################################################################

function fTestCredentialsProc ($ipaddr)
{
    try
    {
        $response = Invoke-CrestronCommand -device $ipaddr -command "`n" -secure
    }
    catch
    {
        $e = $error[0]
        if($e -contains "ermission denied (password)")
        {

        }
    }
}

function fTestCredentialsPanel ($ipaddr)
{
    try
    {
        $response = Invoke-CrestronCommand -device $ipaddr -command "`n" -secure
    }
    catch
    {
        $e = $error[0]
        if($e.FullyQualifiedErrorID -contains "ermission denied (password)")
        {
            
        }
    }
}

function fTestCredentials ($ipaddr, $type)
{
    # processor
    if($type -eq 1)
    {
        $result = fTestCredentialsProc $ipaddr 
    }
    # panel
    elseif($type -eq 2)
    {
        $result = fTestCredentialsPanel $ipaddr
    }
    return $result
}

function fSetAuthentication ($targets)
{
    if(-not $targets)
    {
        fErr "SetAuth: No rooms were targeted." $true
        return
    }

    write-host -f yellow "`nOn which devices do you want to set authentication?`n"
    write-host -f green "  1) Processors"
    write-host -f green "  2) Panels`n"
    #write-host -f green "  *) All`n"
    write-host -f green "  b) to go back`n"

    $s = ""
    while(-not "12bB" -contains $s)
    {
        $s = read-host " "
    }
    if("bB" -contains $s)
    {
        return
    }
    elseif($s.Length)
    {
        $types = [int]$s
    }
    else
    {
        write-host -f red -b black ("WTF SetAuth??  Response: {0}" -f $s)
        return
    }

    foreach($target in [int[]]$targets)
    {
        $r = $global:rooms[$target]

        if($types -eq 1)
        {
            try
            {
                    
            }
            catch
            {

            }
        }    
    }    
}

# Command Load .csv File
######################################################################################################

function fCommandImportCSV
{
    $sheet = importCSVFile (getFileName)
    if($sheet -eq $null)    {    return $false    }
    else
    {
        $global:sheet = $sheet 
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count  
    }   
    if($sheetLen -gt 1)
    {
        return (fParseCSV $global:sheet)
    }
    else
    {
        fErr ("ImportCSV: Skipping process - valid .csv files must have at least 2 lines of data.") $true
        return $null
    }
}

function fFileLoaded
{
    return $global:FileLoaded
}

function fLoadFileDirective
{
    Write-Host -f yellow "Please load a valid .csv file first."
}


# Send IP table to VC-4 server
######################################################################################################

function fSelectRunningProgramList ($rmname, $progs)
{
    $i = 0
    fClear
    $directive = "`nThere are multiple running programs on the {0} server.`nWhich one needs the IP table update?`n`n" -f $rmname
    write-host -f yellow $directive
    foreach ($p in $progs)
    {
        $i ++
        $item = "  {0}) {1}" -f $i, $p 
        write-host -f green $item
    }
    write-host -f green "`n  b) to cancel"
    Write-Host ""

    while ($true)
    {
        $u = read-host " "
        if("1234567890" -notmatch $u)
        {
            write-host "`n"
        }
        elseif([int]$u -gt $progs.count)
        {
            write-host "`n"
        }
        elseif($u -match "bB")
        {
            return $null
        }
        elseif($u -ieq 0)
        {
            return $null
        }
        else 
        {
            return [int]$u
        }
    }

    return $null
}
<#
function fSFTPGetItem ($sessID, $path)
{
    try
    {
        $ipt = Get-SFTPItem -SessionId $sessID -Path $path -Destination $env:
    }
    catch
    {
        fErr ("VC4IPT: Room {0}- {1}  failed to get item {2}." -f $target, $r.roomname, $path) $true
        return $null
    }
    return $ipt
}
#>
function fGetDipFile ($progName, $sessID)
{
    $dip = $null
    $p = "/opt/crestron/virtualcontrol/RunningPrograms/{0}/App/" -f $progName
    $files = Get-SFTPChildItem -SessionId $sessID $p
    $files | ForEach-Object -Process {if($_.Name -match ".dip"){ $dip = $_.Name; return; }}

    if(-not $dip)
    {
        return $null
    }

    #    need to get the dip file here
}


function fVC4AppendIPT ($ipt, $target)
{
    $tableIndex = 0
    $completeTable = @{}

    $addr = $null
    $ipid = $null

    $rows = $ipt.split('`n')
    foreach($r in $rows)
    {
        if($r.contains('=') -and -not $r.contains('$'))
        {
            $id, $value = $r.split('=')

            if($id -match "addr") { $addr = $value }
            elseif($id -match "id") { $ipid = [int32]"0x$value" }
            else { } #throw?

            $id = int($id -replace '\D+(\d+)','$1')
            if($id -gt $tableIndex) { $tableIndex = $id }
            
            if((-not $addr -is $null) -and (-not $ipid -is $null))
            {
                $completeTable[$ipid] = $addr
                $ipid = $null
                $addr = $null
            } 
        }
        else {}
    }   

    # FTR ReporterWebSvc
    $ipid = 0x05
    $tableIndex ++
    $completeTable[$ipid] = $r.ReporterWebSvc.ipaddr
    $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.ReporterWebSvc.ipaddr

    # Wyrestorm Ctrl
    $ipid = 0x06
    $tableIndex ++
    $completeTable[$ipid] = $r.wyrestorm.ipaddr
    $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.wyrestorm.ipaddr    

    # Fixed Cams
    $ipid = 0x07
    for($i = 0; $i -lt $r.Fixed_Cams.length; $i++)
    {
        $ipid += $i
        $tableIndex ++
        $completeTable[$ipid] = $r.Fixed_Cams[$i].ipAddr 
        $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.Fixed_Cams[$i].ipAddr    
    }

    # DSPs
    $ipid = 0x0d
    for($i = 0; $i -lt $r.DSPs.length; $i++)
    {
        $ipid += $i
        $tableIndex ++
        $completeTable[$ipid] = $r.DSPs[$i].ipAddr 
        $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.DSPs[$i].ipAddr    
    }

    # FTR Recorders
    $ipid = 0x18
    for($i = 0; $i -lt $r.RecorderSvrs.length; $i++)
    {
        $tableIndex ++
        $completeTable[$ipid+$i] = $r.RecorderSvrs[$i].ipAddr 
        $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid+$i, $r.RecorderSvrs[$i].ipAddr
    }
        
    # DVD Player
    $ipid = 0x1a
    $tableIndex ++
    $completeTable[$ipid] = $r.DVD_Player.ipaddr 
    $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.DVD_Player.ipaddr   
         
    # Mute Gateways
    $ipid = 0x20
    $tableIndex ++
    $completeTable[$ipid] = $r.Mute_GW.ipaddr 
    $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid, $r.Mute_GW.ipaddr         
    
    # PTZ Cams
    $ipid = 0x23
    for($i = 0; $i -lt $r.PTZ_Cams.length; $i++)
    { 
        $tableIndex ++
        $completeTable[$ipid+$i] = $r.PTZ_Cams[$i].ipAddr 
        $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid+$i, $r.PTZ_Cams[$i].ipAddr 

        $ipid++

        $tableIndex ++
        $completeTable[$ipid+$i] = $r.PTZ_Cams[$i].ipAddr 
        $ipt = "{0}`nid{1}={2:X}`naddr{1}={3}" -f $ipt, $tableIndex, $ipid+$i, $r.PTZ_Cams[$i].ipAddr     
    }

    $global:rooms[$target].$completeIPT = $completeTable

    return $ipt
}

function fSendVC4IPT  ($targets)
{
    if(-not $targets)
    {
        fErr "PanelVTZ: No rooms were targeted." $true
        return
    }

    foreach($target in $targets)
    {
        $ipt = $null
        $dip = $null
        $progSel = $null

        $r = $global:rooms[$target]

        if(-not $r.processor.isIPValid)
        {
            fErr ("VC4IPT: Invalid IP address for room {0}- {1}.`n    {2}" -f $target, $r.roomname, $r.processor.ipAddr) $true
            continue 
        }
        elseif(-not $r.isVC4_System)
        {
            fErr ("VC4IPT: Processor in room {0}- {1} is not marked as a VC-4 server. `nIf this IS a VC-4 system, please mark the appropriate column in the .csv file with a `'1`'." -f $target, $r.roomname) $true
            continue
        }

        # new SFTP session
        $ip = [string]$r.processor.ipAddr
        $usr = "ftr_admin"
        $pw = "Fortherecord123!"
        $spw = ConvertTo-SecureString -String $pw -AsPlainText -Force
        $creds = New-Object System.Management.Automation.PSCredential($usr, $spw)
        $sftp = New-SFTPSession -ComputerName $ip -Credential $creds

        [System.Collections.ArrayList]$progs = get-sftpchilditem $s.sessionid /opt/crestron/virtualcontrol/RunningPrograms | Select-Object -ExpandProperty FullName | foreach { $_.split('/')[-1] }
        $progs.remove("System")
        if($progs.count -eq 0)
        {
            fErr ("VC4IPT: VC-4 server in {0}- {1} does not have any programs running. `nFrom the VC-4 web service, use the `'Add Room`' function." -f $target, $r.roomname) $true
            Remove-SFTPSession $s.Session
            continue
        }
        elseif($progs.count -gt 1)
        {
            #show the rooms and allow for input selection
            fClear
            $progSel = fSelectRunningProgramList $r.roomname $progs

            if($progSel -is $null)
            {
                $sftp.Disconnect()
                continue
            }
            else
            {
                $progSel --
                $ipt = fGetDipFile($progs[$progSel], $sftp.SessionId)
            }
        }
        else
        {
            $progSel = 0
            $ipt = fGetDipFile($progs[$progSel], $sftp.SessionId)
        }

        if($ipt -is $null)
        {
            fErr ("VC4IPT: Couldn't find the IP table for room {0}- {1}, program `'{2}`'" -f $target, $r.roomname, $progs[$progSel]) $true
            continue
        }

        $ipt = fVC4AppendIPT $ipt

        if(-not $ipt -is $null)
        {
                
        }



        

        
        ##Get-SFTPItem -SessionId $sftp.SessionID -Path 
        #Set-SFTPItem 
        #Remove-SFTPSession


        <#
        write-host -f yellow "`nWhich panels do you want to load?`n"
        write-host -f green "  1) TSW-10xx (the 10`" Clerk / Judge panels)"
        write-host -f green "  2) TSW-7xx (the 7`" Counsel panels)`n"
        #write-host -f green "  3) Both`n`n"

        write-host -f green "  b) to go back`n"
        $s = ""
        #>
    }







    # Check if VC4?

    # Connect to VC4 server

    # poll the folder /opt/crestron/virtualcontrol/RunningPrograms/
    # parse the list of subfolders
    # remove the "/System" subfolder from our list
    # If there is only 1 RunningProgram, verify via input that it is the correct one
    # If >1, list the RunningPrograms, and allow input select
    # 
    # Once we have our target rooms:
    # Copy the .dip file via ftp
    # Check to be sure we aren't duplicating any entries
    # Add the appropriate lines
    # If successful, send back to the server via ftp, overwriting the existing file
    
    # On all-rooms complete, restart the service>   sudo systemctl restart virtualcontrol
    
     

}


# Get User Command
######################################################################################################

function getCommand
{
    $choice = Read-Host " "

    # Import CSV File
    if($choice -ieq '1')
    {
        fClear        
        $result = fCommandImportCSV
        continueScript
    }
    # Load Processor IP Code
    elseif($choice -ieq '2')
    {
        fClear
        if(fFileLoaded)
        {
            fPrettyPrintRooms
            write-host -f yellow -b black "`nLoad processor code (.lpz / .cpz file):"
        
            fSendProcLPZ (fGetRangeOfRooms)
        }
        else
        {
            fLoadFileDirective
        }

        continueScript 
    }
    # Send Processor IP Table
    elseif($choice -ieq '3')
    {
        fClear
        if(fFileLoaded)
        {
            fPrettyPrintRooms
            write-host -f yellow -b black "`nSend processor IP table:"
        
            fSendProcIPT (fGetRangeOfRooms)
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }
    # Load Touch Panel File
    elseif($choice -ieq '4')
    {
        fClear
        if(fFileLoaded)
        {
            fPrettyPrintRooms
            write-host -f yellow -b black "`nSend touch panel file (.vtz) :"
            fSendPanelVTZ (fGetRangeOfRooms)
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }
    # Send Touch Panel IP Table
    elseif($choice -ieq '5')
    {
        fClear
        if(fFileLoaded)
        {
            fPrettyPrintRooms
            write-host -f yellow -b black "`nSend touch panel IP table:"
        
            fSendPanelIPT (fGetRangeOfRooms)
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }
    <#
    # Set Device Hostnames
    elseif($choice -ieq '6')
    {
        fClear
        if(fFileLoaded)
        {
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }#>

    # Set Authentication Mode
    elseif($choice -ieq '6')
    {
        fClear
        if(fFileLoaded)
        {
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }

    # Get Device Status
    elseif($choice -ieq '7')
    {
        fClear
        if(fFileLoaded)
        {
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }

    # Send VC-4 IPT
    elseif($choice -ieq '9')
    {
        fClear
        if(fFileLoaded)
        {
            fPrettyPrintRooms
            write-host -f yellow -b black "`nSend VC-4 IP Table:"
            fSendVC4IPT (fGetRangeOfRooms)
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }

    # Get Info
    elseif($choice -ieq 'i')
    {
        if(fFileLoaded)
        {
            fClear
            $result = fShowInfo
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }
    # Exit Script
    elseif($choice -ieq 'x')
    {
        exit
    }
    # else
    else
    {
        continueScript
    }
}

function shellTextColor ([bool]$b)
{
    if($b) {return "green"}
    return "gray"
}

function updateShell
{
    $data = "`n`n"
    $data += "-------------------------------------------------"
    Write-Host -f Yellow $data

    $data  = "1) {0}" -f $global:menuFunctions[1]
    Write-Host -f Green $data
    
    $data  = "2) {0}`n" -f $global:menuFunctions[2]
    $data += "3) {0}`n" -f $global:menuFunctions[3]
    $data += "4) {0}`n" -f $global:menuFunctions[4]
    $data += "5) {0}`n" -f $global:menuFunctions[5]
    $data += "`n6) {0}`n" -f $global:menuFunctions[6]
    $data += "7) {0}`n" -f $global:menuFunctions[7]
    $data += "8) {0}`n" -f $global:menuFunctions[8]
    $data += "`n9) {0}`n" -f $global:menuFunctions[9]
    Write-Host -f (shellTextColor $FileLoaded) $data
    
    $data  = "`ni) info`n"
    $data += "`nx) exit`n"
    Write-Host -f Green $data
}


#endregion


function continueScript
{
    updateShell
    getCommand
}


$global:menuFunctions[1] = "Load .csv File"
$global:menuFunctions[2] = "Load Processor Code"
$global:menuFunctions[3] = "Send Processor IP Table"
$global:menuFunctions[4] = "Load Touch Panel File"
$global:menuFunctions[5] = "Send Panel IP Table"

$global:menuFunctions[6] = "Set Authentication"
$global:menuFunctions[7] = "Get Device Status"
$global:menuFunctions[8] = "Update Firmware"

$global:menuFunctions[9] = "Send VC-4 IP Table"




continueScript




