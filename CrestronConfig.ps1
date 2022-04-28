<#
    
#>


#####################################################################################################################
<#
    Globals
#>
#####################################################################################################################


$global:DEBUG = $true

$global:User = "ftr_admin"
$global:Pass = "Fortherecord123!"

# This value needs to be copied while the script is running. It doesn't work if you run the script, and then try to reference the value from console
$global:ScriptPath = $PSScriptRoot + '\'
$global:DataFileName = "AZMC_CourtroomData.csv"
$global:SelectedFile = ""
$global:FileLoaded = $false
$global:numOfRooms = 0

$global:logLevel = 1

$global:rooms = @{}

$global:defaults = @{}


#####################################################################################################################
<#
    Utils
#>
#####################################################################################################################

function removeWhitespace([string]$s)
{
    return ($s -replace ' ','')
}

function fClear($ms = 100)
{
    clear
    start-sleep -Milliseconds 100
}

fClear


function fatal
{
    Read-Host -Prompt "`n`nPress any key to exit"
    exit
}

function fDefault($a, $b)
{
    if($a)
    {
        return $a
    }
    return $b
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


# Error handling
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
            Write-Host -f Yellow 
            (
                "Hey!`n" + 
                "This script requires software called 'PSCrestron' (the PowerShellCrestron module).`n" +
                "I believe you have a good internet connection right now, but something else went wrong when I tried to download and install this software.`n" +
                
                "If you get this message again, try to manually download and install PSCrestron. Drop this into a web browser:`n" +
                "https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe`n" +
                "... and run the file, following any prompts to install the software.`n`n" +

                "If you're still having problems, you can email bcampagnola@fortherecord.com, with 'HALP!' in the subject line." 
            ) 
            fatal
        } 
        2 
        {
            Write-Host -f Yellow 
            (
                "Hey!`n" +
                "This script requires software called 'PSCrestron' (the PowerShellCrestron module).`n" +
                "I attempted to download the software, but it appears that this computer doesn't have an internet connection.`n`n" +

                "Try gettin' some internets, and then run the script again.`n`n" +

                "If you believe this is not the problem, there might be something wrong with the permissions on your machine, or you're having delusions of connectivity.`n`n`n" +


                "If you're online but still having problems, you can email bcampagnola@fortherecord.com, with 'HALP!' in the subject line." 
            )        
            fatal        
        }
        3
        {
            Write-Host -f Yellow 
            (
                "Hm.." + 
                "I couldn't find that file.`n" + 
                "Does this path look right?`n`n" +        
                
                "$ScriptPath" + "$DatafileName`n`n" + 

                "If not, try the options to edit the script path and .csv file name values.`n`n"
            )    
        }
    }
}


#####################################################################################################################
<#
    Import Libs
#>
#####################################################################################################################

Add-Type -AssemblyName System.Windows.Forms

function tryImport
{
    try
    {
        Import-Module PSCrestron -ErrorAction Stop
        return 0
    }
    catch 
    { 
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

#####################################################################################################################
<#
    Logging & Debug
#>
#####################################################################################################################

if(-not(Test-Path -path "Logs"))
{
    try
    {
        New-Item -Path . -Name "Logs" -ItemType "directory" | out-null
    }
    catch
    {
        fErr ("Logs: Failed to make directory '/Logs'") $true
    }
}

function fGetDateTime
{
    return ("{0:yyyy.MM.dd_HH.mm.ss.fff}" -f (Get-Date))
}


$global:logFileName = fGetDateTime
$global:logFileName += ".log"

try
{
    New-Item -Path .\Logs -ItemType "file" -Name $global:logFileName | out-null
}
catch
{
    fErr ("Logs: Failed to create '/Logs/{0}'" -f $global:logFileName) $true        
}

function fLog ([string]$s)
{
    $dt = fGetDateTime
    $s = "{0}  -  {1}" -f [string]$dt, $s
    $s >> .\Logs\$global:logFileName
}

function fErr ($s, $c)
{
    if($global:logLevel -gt 0)
    {
        fLog $s
    }

    if($c -eq $true)
    {
        if($global:debug)
        {
            Write-Host -b black -f red $s
        }
    }
    else
    {
        if($global:debug)
        {
            Write-Host -b black -f green $s
        }
    }
}



#####################################################################################################################
<#
    Data & File Import
#>
#####################################################################################################################


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

    # should probably name every room uniquely, in case we want to create a $global:roomsByName hashtable
    $c.RoomName = $row | select-object -ExpandProperty Room_Name
    $c.Default = $row | select-object -ExpandProperty Defaults
    $c.isDef = $c.Default.contains("*")
    $c.defIndex = [int]($c.default -replace "[^0-9]", "")

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

    return $c
}

function importFile([string]$fileName)
{
    if([System.IO.File]::Exists("$fileName"))
    {
        $global:sheet = Import-csv $fileName
        $sheetLen = $global:sheet | Measure-Object | Select-Object -ExpandProperty Count
        fErr ("Import: File opened successfully.") $false
        fErr ("Import: File name - {0}" -f $fileName) $false
        fErr ("Import: Found {0} data lines to parse." -f $sheetLen) $false
    }
    else
    {
        fErr ("Import: File import failed.") $true
        return
    }

    $global:rooms = @{}
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
                    fErr ("Import: Processed line {0:d3}: default index = {1}" -f $i, $c.defIndex) $false
                }
                else
                {
                    fErr ("Import: Line {0:d3} failed: default index = {1}." -f $i, $c.defIndex) $true
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
        $FileLoaded = $true
        fErr ("Import: File import success. {0} rooms parsed." -f $global:numOfRooms) $false
    }
}

function getFileName
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $ScriptPath }
    $null = $FileBrowser.ShowDialog()
    $fileName = $FileBrowser | Select-Object -ExpandProperty FileName
    $global:SelectedFile = $fileName
    return ($fileName)    
}

function selectAndImport
{  
    importFile (getFileName)
}

#####################################################################################################################
<#
    Menu Functions
#>
#####################################################################################################################


function showAllRooms
{
    $e = $global:rooms.keys | sort

    foreach($r in $e)
    {
        write-host -b black -f green ("{0:d3}. {1}" -f $r, $global:rooms[$r].RoomName) 
    }
}

function getRangeOfRooms
{
    Write-host -f yellow "Which rooms do you want to target? (e.g. '3,9-12,17,16')`n"
    Write-host -f green "  *  to target all rooms"
    write-host -f green "  b  to go back`n"
    return (read-host " ")
}

function verifyRangeChars([string]$s)
{
    # regex could improve this a lot
    # https://stackoverflow.com/questions/23697408/powershell-get-number-from-string
    $validChars = "0123456789,-".ToCharArray()
    $bad = ""

    foreach($c in $s.ToCharArray())
    {
        if($c -notin $validChars)
        {
            $bad += $c
        } 
    }
    return $bad
}

function decodeRange([string]$s)
{
    $s = removeWhitespace($s)
    
    if($s -ieq "b")
    {
        continueScript
    }
    elseif($s -ieq "*")
    {
        $r = @()
        foreach($k in $global:rooms.keys)
        {
            $r += $k
        }
        fErr ("Target: User selected to target all (`'*`') rooms. {0} rooms added." -f $r.length) $false
        return [string[]]$r
    }
    # target the specified rooms
    else
    {
        $badChars = verifyRangeChars($s)
        if($badChars.Length -eq 0)
        {
            $r = @()
            $segments = $s.split(',')

            foreach($segment in $segments)
            {
                if($segment.contains('-'))
                {
                    $v = $segment.split('-')
                    $r += [int]$v[0]..[int]$v[1]                    
                }
                else
                {
                    $r += [int]$segment
                }
            }
            # Write-Host $r
            return $r
        }
        else
        {
            Write-Host -f Red "Invalid characters:`n"   #throw the errors
            Write-Host -f Red $badChars 
            return ""
        }
    }
}

function fIPTSend($SessID, [int]$IPID, $ipaddr)
{
    $AddPeer = "AddP 0x{0:X} {1}" -f $IPID, $ipaddr
    # Write-Host -b black -f Green $addpeer

    try
    {
        $response = Invoke-CrestronSession $SessID -Command ("{0}" -f $AddPeer)
        return $false
    }
    catch
    {
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

        $result = fIPTSend $SessID $ipid $ipaddr
        if(-not $result)
        {
            fErr ("ProcIPT: Successfully sent IPID 0x{0:x2} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $false
        }
        else
        {
            fErr ("ProcIPT: Failed to commit IPID {0} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $true         
        }
        $ipid++
    }
}

function sendProcIPT
{    
    showAllRooms
    write-host -f yellow -b black "`nSend processor IP table:"

    $targets = decodeRange(getRangeOfRooms)
    if(-not $targets)
    {
        fErr "ProcIPT: No rooms were targeted." $true
        return
    }
    
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
        try
        {
            $ipaddr = fIPAddr $r.subnet (fDefault $r.Processor_IP $d.Processor_IP)
            $SessID = Open-CrestronSession -Device $ipaddr -Secure # -Username $User -Password $Pass #-ErrorAction SilentlyContinue

            fErr ("ProcIPT: Open-CrestronSession for room {1:d3} successful. SessionID: {0}" -f $SessID, $target) $False
        }
        catch
        {
            fErr ("ProcIPT: Open-CrestronSession failed for room {0:d3}." -f $target) $True  
            continue 
        }

        # Hostname
        try
        {
            $hostnameResponse = Invoke-CrestronSession $SessID "hostname"
            $deviceHostname = [regex]::Match($hostnameResponse, "(?<=Host\sName:\s)[\w-]+").value
            fErr ("ProcIPT: Retrieved hostname `'{0}`' from the {1:d3} processor." -f $deviceHostname, $target) $false
        }
        catch
        {
            fErr ("ProcIPT: Processor in room {0:d3} failed to respond with a valid hostname." -f $target) $True
            fErr ("ProcIPT: Attetmpting to continue with loading the IP table.") $True
        }

        # Send Crestron Program to Processor Secure
        # Send-CrestronProgram -ShowProgress -Device $ProcIP -LocalFile $LPZPath



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

                    $result = fIPTSend $SessID $ipid $ipaddr
                    if(-not $result)
                    {
                        fErr ("ProcIPT: Successfully sent IPID 0x{0:x2} IPAddr {1} to the processor in room {2:d3}." -f [int]$ipid, $ipaddr, $target) $false
                    }
                    else
                    {
                        fErr ("ProcIPT: Failed to commit IPID 0x{0:x2} IPAddr {1} to the processor in room {2:d3}." -f [int]$ipid, $ipaddr, $target) $true         
                    }
                    $ipid++

                    $result = fIPTSend $SessID $ipid $ipaddr
                    if(-not $result)
                    {
                        fErr ("ProcIPT: Successfully sent IPID 0x{0:x2} IPAddr {1} to the processor in room {2:d3}." -f [int]$ipid, $ipaddr, $target) $false
                    }
                    else
                    {
                        fErr ("ProcIPT: Failed to commit IPID 0x{0:x2} IPAddr {1} to the processor in room {2:d3}." -f [int]$ipid, $ipaddr, $target) $true         
                    }
                    $ipid++
                }
            }
        }

        # Reset Program
        try
        {
            Invoke-CrestronSession $SessID 'progreset -p:01'
            fErr ("ProcIPT: Restarting program for room {0:d3}." -f $target) $False
        }
        catch
        {
            fErr ("ProcIPT: Program restart failed for room {0:d3}." -f $target) $True
        }

        # Close Session
        try
        {
            Close-CrestronSession $SessID
            fErr ("ProcIPT: Close-CrestronSession successful for room {0:d3}." -f $target) $False
        }
        catch
        {
            fErr ("ProcIPT: Close-CrestronSession failed for room {0:d3}, `$SessID {1}." -f $target, $SessID) $True
        }
    }
}


function sendProcLPZ
{
    showAllRooms

    write-host -f yellow -b black "`nSend .lpz file:"
    $targets = decodeRange(getRangeOfRooms)
    if(-not $targets)
    {
        fErr "ProcLPZ: No rooms were targeted." $true
        return
    }
    else
    {
        write-host -f yellow "`nSend IP table?  y/n`n"
        write-host -f green "  b  to go back`n"
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
            write-host -f red -b black "WTF sendProcLPZ"
            return
        }
    }
    # Write-Host -b red "Target rooms: $targets"
    foreach($target in $targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr ("ProcLPZ: Target room {0:d3} is not in the list of rooms." -f $target) $True
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
        try
        {
            $ipaddr = fIPAddr $r.subnet (fDefault $r.Processor_IP $d.Processor_IP)
            $SessID = Open-CrestronSession -Device $ipaddr -Secure # -Username $User -Password $Pass #-ErrorAction SilentlyContinue

            fErr ("ProcLPZ: Open-CrestronSession for room {1:d3} successful. SessionID: {0}" -f $SessID, $target) $False
        }
        catch
        {
            fErr ("ProcLPZ: Open-CrestronSession failed for room {0:d3}." -f $target) $True  
            continue 
        }

        # Hostname
        try
        {
            $hostnameResponse = Invoke-CrestronSession $SessID "hostname"
            $deviceHostname = [regex]::Match($hostnameResponse, "(?<=Host\sName:\s)[\w-]+").value
            fErr ("ProcLPZ: Retrieved hostname `'{0}`' from the {1:d3} processor." -f $deviceHostname, $target) $false
        }
        catch
        {
            fErr ("ProcLPZ: Processor in room {0:d3} failed to respond with a valid hostname." -f $target) $True
            fErr ("ProcLPZ: Attetmpting to continue with .lpz load.") $True
        }

        $lpzFile = fDefault $r.FileName_LPZ $d.FileName_LPZ
        write-host -b black -f Yellow $ipaddr
        write-host -b black -f green $lpzFile
        # Send Crestron Program to Processor Secure
        try
        {
            if($sendIPT)
            {
                Send-CrestronProgram -ShowProgress -Device $ipaddr -LocalFile $lpzFile
                fErr ("ProcLPZ: Successfully sent file '{0}' to room {1:d3} with IP table." -f $lpzFile, $target) $False
            }
            else
            {
                Send-CrestronProgram -ShowProgress -Device $ipaddr -LocalFile $lpzFile -DoNotUpdateIPTable
                fErr ("ProcLPZ: Successfully sent file '{0}' to room {1:d3} without IP table." -f $lpzFile, $target) $False
            }
        }
        catch
        {
            write-warning $error[0]
            fErr ("ProcLPZ: Failed to send file '{0}' to room {1:d3}." -f $lpzFile, $target) $True
        }


        # Reset Program
        try
        {
            Invoke-CrestronSession $SessID 'progreset -p:01'
            fErr ("ProcLPZ: Restarting program for room {0:d3}." -f $target) $False
        }
        catch
        {
            fErr ("ProcLPZ: Program restart failed for room {0:d3}." -f $target) $True
        }

        # Close Session
        try
        {
            Close-CrestronSession $SessID
            fErr ("ProcLPZ: Close-CrestronSession successful for room {0:d3}." -f $target) $False
        }
        catch
        {
            fErr ("ProcLPZ: Close-CrestronSession failed for room {0:d3}, `$SessID {1}." -f $target, $SessID) $True
        }
    }
}


######################################################################################################
######################################################################################################
######################################################################################################



function showInfo
{
    $s = "[no file loaded]"
    if($global:FileLoaded)
    {
        $s = $global:SelectedFile
    }
    $data = ""
    $data = (".csv file: " + $s + "`n") 
    $data += ("num of rooms loaded: " + $NumOfRooms + "`n")

    Write-host -b black -f yellow $data
}

<######################################################################################################

<######################################################################################################>


function updateShell
{
    $data = "`n`n"
    $data += "-------------------------------------------------"
    Write-Host -f Yellow $data

    $data = "1) Load spreadsheet`n"
    $data += "2) Send Crestron processor code`n"
    $data += "3) Send Crestron processor IP table`n"
    $data += "4) Send Crestron touch panel code`n"
    $data += "5) Send Crestron panel IP table`n"

    $data += "`ni) info`n"
    $data += "`nx) exit`n"
    Write-Host -f Green $data
}


function getCommand
{
    $choice = Read-Host " "

    if($choice -ieq '1')
    {
        $result = selectAndImport
        continueScript
    }
    elseif($choice -ieq '2')
    {
        sendProcLPZ
        continueScript 
    }
    elseif($choice -ieq '3')
    {
        sendProcIPT
        continueScript
    }
    elseif($choice -ieq '4')
    {
        continueScript
    }
    elseif($choice -ieq '5')
    {
        continueScript
    }
    elseif($choice -ieq '6')
    {
        continueScript
    }
    elseif($choice -ieq 'i')
    {
        showInfo
        continueScript
    }
    elseif($choice -ieq 'x')
    {
        exit
    }
    else
    {
        continueScript
    }
}


function continueScript
{
    updateShell
    getCommand
}

continueScript








