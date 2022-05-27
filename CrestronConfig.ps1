<#
    FTR Crestron PowerShell Script

    ver 1.whatever   
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
        return "ipaddress"
    }
    elseif (fValidateGuid ($t))
    {
        return "guid"
    }
    elseif (fValidateHostname ($t))
    {
        return "hostname"
    }
    else
    {
        return "none"
    }
}

function fIPAddr([string]$a, [string]$b)
{
    $ipa = ("{0}.{1}" -f $a, $b)
    while($ipa -contains '..')
    {
        $ipa = $ipa.replace('..', '.')
    }
    return $ipa
}

function fDefault($a, $b)
{
    if($a) { return $a }
    return $b
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
        $r = @()
        $segments = $s.split(',')

        foreach($segment in $segments)
        {
            $c = $segment.split('-').count
            if($c -eq 2)
            {
                $v = $segment.split('-')
                $r += [int]$v[0]..[int]$v[1]                    
            }
            elseif($c -eq 1)
            {
                $r += [int]$segment
            }
            else
            {
                fErr ("DecodeRange: Syntax error in selection range. This is invalid: '{0}'" -f $segment) $true
                return $null
            }
        }
        return [string[]]$r
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

function tryImport
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
$err = tryImport

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

    $err = tryImport
    if($err -ne 0)
    {
        err 1
    }
}

#endregion


#####################################################################################################################
<#    Data & File Import    #>
#####################################################################################################################
#region

class Courtroom
{
    [int]$Index
    [string]$RoomName
    [string]$Default
    [bool]$isDef
    [int]$defIndex
    [string]$FacilityName
    [string]$Subnet
    [string]$Processor_IP
    [string]$FileName_LPZ
    [bool]$localLPZFile
    [string[]]$Panel_IP
    [string[]]$FileName_VTZ
    [bool]$localVTZFile
    [string]$ReporterWebSvc_IP
    [string]$Wyrestorm_IP
    [string[]]$FixedCam_IP
    [string[]]$DSP_IP
    [string[]]$RecorderSvr_IP
    [string]$DVD_IP
    [string]$MuteGW_IP
    [string[]]$PTZCam_IP
    $IPTable = @{}
}

function parseLine($c, $i, $row)
{
    # the index is the key for $global:rooms dict entries
    # should also align with the excel spreadsheet row numbers
    $c.Index = $i           
    try
    {
        # should probably name every room uniquely, in case we want to create a $global:roomsByName hashtable
        $c.RoomName = $row | select-object -ExpandProperty Room_Name
        $c.Default = $row | select-object -ExpandProperty Defaults
        $c.isDef = $c.Default.contains("*")
        $c.defIndex = [int]($c.default -replace "[^0-9]", "")

        #$c.Host = $row | select-object -Expandproperty Hostname_Prefix

        $c.FacilityName = $row | select-object -ExpandProperty Facility_Name
        $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

        # only 1 Proc_IP and 1 LPZ per room
        $c.Processor_IP = $row | select-object -ExpandProperty Processor_IP
        $c.FileName_LPZ = $row | select-object -ExpandProperty FileName_LPZ

        # multiple panel_IPs and multiple fileName_VTZs are possible
        $c.Panel_IP = $row | select-object -ExpandProperty Panel_IP
        $c.FileName_VTZ = $row | select-object -ExpandProperty FileName_VTZ

        # x 1
        $c.ReporterWebSvc_IP = $row | select-object -ExpandProperty IP_ReporterWebSvc
        $c.Wyrestorm_IP = $row | select-object -ExpandProperty IP_WyrestormCtrl

        # x Multiple
        $c.FixedCam_IP = $row | select-object -ExpandProperty IP_FixedCams        
        $c.DSP_IP = $row | select-object -ExpandProperty IP_DSPs
        $c.RecorderSvr_IP = $row | select-object -ExpandProperty IP_Recorders
    
        # x 1
        $c.DVD_IP = $row | select-object -ExpandProperty IP_DVDPlayer
        $c.MuteGW_IP = $row | select-object -ExpandProperty IP_AudicueGW
    
        # x Multiple
        $c.PTZCam_IP = $row | select-object -ExpandProperty IP_PTZCams 
    }
    catch
    {
        fErr ("ParseLine: Failed to assign a value from data row {0}.`n {1}" -f $i, $error) $true
    }    
    
    try
    {
        # split the multiples by ~
        $c.Panel_IP = $c.Panel_IP[0].split('~')
        $c.FileName_VTZ = $c.FileName_VTZ[0].split('~')
        $c.FixedCam_IP = $c.FixedCam_IP[0].split('~')
        $c.DSP_IP = $c.DSP_IP[0].split('~')    
        $c.RecorderSvr_IP = $c.RecorderSvr_IP[0].split('~')
        $c.PTZCam_IP = $c.PTZCam_IP[0].split('~')

        $numOfPanels = $c.panel_IP.length
        $numOfVTZ = $c.FileName_VTZ.length

        # if there are multiple panel IP addresses, but there are fewer VTZ files listed,
        # take the first FileName_VTZ[0], and create an array of multiple VTZ files to match
        # the number of panel IPs  
        if($numOfPanels -gt $numOfVTZ)
        {
            $c.FileName_VTZ = [string[]]@($c.FileName_VTZ[0])*$numOfPanels
        }
    }
    catch
    {
        fErr ("ParseLine: Failed to split multiple-value column in data row {0}.`n {1}" -f $i, $error) $true
    }

    return $c
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
            $c = new-object -TypeName Courtroom
            $c = parseLine $c $i $row

            if($c.isDef -eq $true)
            {
                if([bool]$c.defIndex)
                {
                    $global:defaults[$c.defIndex] = $c
                    $global:numOfDefaults++
                }
                else
                {
                    fErr ("Import: Line {0:d3} failed." -f $i) $true
                    #continue
                }

            }
            else
            {
                $global:rooms[[int]$c.Index] = $c
                # $roomsByName[$c.RoomName] = $rooms[$c.Index]

                $global:numOfRooms++
                # fErr ("Import: Parsed data line {0:d3}" -f $i) $false
            }
        }
        catch
        {
            fErr ("Import: File import failed for line {0:d3}." -f $i) $true
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
    $connection = fValidateConnection $target

    if($connection.contains("none"))
    {
        fErr ("CrestronSDK: Invalid connection data for CrestronGetModel.`n    {0}"-f $target) $true
        return     
    }

    try
    {
        if ($connection.contains("ipaddr") -or $connection.contains("host"))
        {
            $response = Invoke-CrestronCommand -device $target -Command "`n"
        }
        elseif ($connection.contains("guid"))
        {
            $response = Invoke-CrestronSession $target -Command "`n"
        }
        
        $DeviceModel = ($response -replace "[^0-9]", "")
        return $DeviceModel
    }
    catch
    {
        fErr ("CrestronSDK: Failed to get Crestron model number.`n    Error: {0}`n" -f $error)
    }
}


function fCrestronStartSession($ipaddr, $target, [bool]$secure=$true)
{
    try
    {
        if($secure)
        {
            $SessID = Open-CrestronSession -Device $ipaddr -Secure # -Username $User -Password $Pass #-ErrorAction SilentlyContinue    
        }
        else
        {
            $SessID = Open-CrestronSession -Device $ipaddr # -Username $User -Password $Pass #-ErrorAction SilentlyContinue            
        }
        fErr ("PSCrestron: Open-CrestronSession for room {1:d3} successful. SessionID: {0}" -f $SessID, [int]$target) $False

        return $SessID
    }
    catch
    {
        fErr ("PSCrestron: Open-CrestronSession failed for room {0:d3}." -f [int]$target) $True  
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
        return $null
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

    return $targets
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

function fIPTSend($SessID, [int]$IPID, $ipaddr, $target)
{

    $addPeer = "AddP {0:X} {1}" -f $IPID, $ipaddr

    try
    {
        $response = Invoke-CrestronSession $SessID -Command ("{0}" -f $addPeer)
        fErr ("ProcIPT: Successfully sent IPID {0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $false
        write-host -b darkgray -f cyan (">>> {0}" -f $addPeer)
        write-host -b darkgray -f yellow ("<<< {0}" -f $response)
        return $false
    }
    catch
    {
        fErr ("ProcIPT: Failed to commit IPID {0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $true         
        return $true
    }
}

function fIPT ($SessID, [int]$ipid, $sub, $node, $target)
{
    # if the node address value is "0", or the value is [bool]-not, skip
    if($node -ieq "0" -or (-not [bool]$node))
    {
        write-host -f green -b black ("ProcIPT: Empty or null IP nodes for room {0} beginning IPID 0x{1:x2}" -f $target, [int]$ipid)
        return
    }

    foreach($n in $node)
    {
        $ipaddr = fIPAddr $sub $n
        $result = fIPTSend $SessID $ipid $ipaddr $target
        $ipid++
    }
}

function fSendProcIPT ([int[]]$targets)
{    
    # Write-Host -b red "Target rooms: $targets"
    foreach($target in $targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr ("ProcIPT: Target room {0:d3} is not in the list of rooms." -f $target) $True
            continue
        }
        $r = $global:rooms[$target]
        $sub = $r.Subnet

        $d = $null

        if($r.defIndex -gt 0)
        {
            if($global:defaults.ContainsKey($r.defIndex))
            {
                $d = $global:defaults[$r.defIndex]
            }
        }
        # Connect to Device -secure
        $ipaddr = fIPAddr (fDefault $r.subnet $d.subnet) (fDefault $r.Processor_IP $d.Processor_IP)
        $SessID = fCrestronStartSession $ipaddr $target

        if($SessID -eq $null) { continue } 


        # FTR ReporterWebSvc
        fIPT $SessID 0x05 $sub (fDefault $r.ReporterWebSvc_IP $d.ReporterWebSvc_IP) $target
        # Wyrestorm Ctrl
        fIPT $SessID 0x06 $sub (fDefault $r.Wyrestorm_IP $d.Wyrestorm_IP) $target        
        # Fixed Cams
        fIPT $SessID 0x07 $sub (fDefault $r.FixedCam_IP $d.FixedCam_IP) $target
        # DSPs
        fIPT $SessID 0x0d $sub (fDefault $r.DSP_IP $d.DSP_IP) $target
        # FTR Recorders
        fIPT $SessID 0x18 $sub (fDefault $r.RecorderSvr_IP $d.RecorderSvr_IP) $target
        # DVD Player
        fIPT $SessID 0x1a $sub (fDefault $r.DVD_IP $d.DVD_IP) $target
        # Mute Gateways
        fIPT $SessID 0x20 $sub (fDefault $r.MuteGW_IP $d.MuteGW_IP) $target
        #PTZ Cams
        if($r.PTZCam_IP -ieq "0")
        {
            # skip if marked as "0"
        }
        else
        {
            $l = fDefault $r.PTZCam_IP $d.PTZCam_IP
            if($l)
            {
                $ipid = 0x23
                foreach($n in $l)
                {
                    $ipaddr = fIPAddr $sub $n
                    $result = fIPTSend $SessID $ipid $ipaddr $target
                    $ipid++

                    $result = fIPTSend $SessID $ipid $ipaddr $target
                    $ipid++
                }
            }
        }

        fCrestronRestartProg $SessID $target
        fCrestronCloseSession $SessID $target
    }
}


# Send Panel IPT
######################################################################################################

function fSendPanelIPT ([int[]]$targets)
{    
    # Write-Host -b red "Target rooms: $targets"
    foreach($target in $targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr ("PanelIPT: Target room {0:d3} is not in the list of rooms." -f $target) $True
            continue
        }
        $r = $global:rooms[$target]
        $sub = $r.Subnet

        $d = $null

        if($r.defIndex -gt 0)
        {
            if($global:defaults.ContainsKey($r.defIndex))
            {
                $d = $global:defaults[$r.defIndex]
            }
        }
        # Connect to Device -secure
        $ipaddr = fIPAddr (fDefault $r.subnet $d.subnet) (fDefault $r.Processor_IP $d.Processor_IP)
        $SessID = fCrestronStartSession $ipaddr $target

        if($SessID -eq $null) { continue } 

        # FTR ReporterWebSvc
        fIPT $SessID 0x05 $sub (fDefault $r.ReporterWebSvc_IP $d.ReporterWebSvc_IP) $target
       
        fCrestronRestartProg $SessID $target
        fCrestronCloseSession $SessID $target
    }
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
    # write-host -f green "  3) Both`n"
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

        if($r.defIndex)
        {
            if($global:defaults.ContainsKey($r.defIndex))
            {
                $d = $global:defaults[$r.defIndex]
            }
            else
            {
                fErr ("SetAuth: Room {0} {1} code send is attempting to use data from default line {2}, which does not exist." -f $target, $r.roomname, $r.defIndex) $true
            }
        }

        if($types -eq 1)
        {
            $ipaddr = fIPAddr (fDefault $r.subnet $d.subnet) (fDefault $r.Processor_IP $d.Processor_IP)
            try
            {
                
            }
            catch
            {

            }
        }    
    }    
}


# Send .lpz File to Processors
######################################################################################################

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
    while(-not "ynbYNB" -contains $s)
    {
        $s = read-host " "
    }
    if("bB".contains($s))
    {
        return
    }
    elseif("yY".contains($s))
    {
        $sendIPT = $true
    }
    elseif("nN".contains($s))
    {
        $sendIPT = $false
    }
    else
    {
        write-host -f red -b black ("WTF sendProcLPZ??  Response: {0}" -f $s)
        return
    }

    foreach($target in $targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr ("ProcLPZ: Target room {0:d3} is not in the list of rooms." -f $target) $True
            continue
        }

        $r = $null
        $d = $null

        $r = $global:rooms[$target]

        if($r.defIndex)
        {
            if($global:defaults.ContainsKey($r.defIndex))
            {
                $d = $global:defaults[$r.defIndex]
            }
            else
            {
                fErr ("ProcLPZ: Room {0} {1} code send is attempting to use data from default line {2}, which does not exist." -f $target, $r.roomname, $r.defIndex) $true
            }
        }

        $ipaddr = fIPAddr (fDefault $r.subnet $d.subnet) (fDefault $r.Processor_IP $d.Processor_IP)
        if(-not (fValidateIPAddress $ipaddr))
        {
            fErr ("ProcLPZ: Invalid IP address for room {0} {1}.`n    {2}" -f $target, $r.roomname, $ipdadr)
            continue 
        }

        # .lpz File
        $lpzFile = $global:scriptPath + (fDefault $r.FileName_LPZ $d.FileName_LPZ)
        try
        {
            if($sendIPT)
            {
                Send-CrestronProgram -ShowProgress -Device $ipaddr -LocalFile $lpzFile -Secure
                fErr ("ProcLPZ: Successfully sent file '{0}' to room {1:d3} with IP table." -f $lpzFile, $target) $False   
            }
            else
            {
                Send-CrestronProgram -ShowProgress -Device $ipaddr -LocalFile $lpzFile -Secure -DoNotUpdateIPTable
                fErr ("ProcLPZ: Successfully sent file '{0}' to room {1:d3} without IP table." -f $lpzFile, $target) $False
            }
        }
        catch
        {
            fErr ("ProcLPZ: Failed to send file`n    '{0}'`n    to room {1:d3}.`n    Error: {2}" -f $lpzFile, $target, $Error) $True
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
        if(fFileLoaded)
        {
            # poll IP table for accuracy and connectivity
            exit
        }
        else
        {
            fLoadFileDirective
        }

        continueScript
    }
    else
    {
        continueScript
    }
}

function updateShell
{
    $data = "`n`n"
    $data += "-------------------------------------------------"
    Write-Host -f Yellow $data

    $data  = "1) Load .csv file"
    Write-Host -f Green $data
    
    $data  = "2) Load processor code`n"
    $data += "3) Send processor IP table`n"
    $data += "4) Load touch panel file`n"
    $data += "5) Send panel IP table`n"
    #$data += "6) Set hostnames`n"
    $data += "6) Set authentication`n"
    $data += "7) Get device status`n"
    if(fFileLoaded)
    {
        Write-Host -f Green $data
    }
    else
    {
        Write-Host -f Gray $data
    }

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

continueScript








