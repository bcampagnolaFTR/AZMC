<#

.INSTRUCTIONS
    Read the README.txt file.
    Seriously.

.DESCRIPTION
    Pull in CSV of all Courtroom configuration details. Updates Display Project/Program, and Updates IP Tables.

.INPUTS
    Complete CSV with all Courtroom configuraiton details placed in local directory with script.

.NOTES
    Version: 1.01.00
    Creation Date: 2/1/2022
    Purpose/Change: Adding Addtional Logging
#>



$global:DEBUG = $true

$global:User = "ftr_admin"
$global:Pass = "Fortherecord123!"

# This value needs to be copied while the script is running. It doesn't work if you run the script, and then try to reference the value from console
$global:ScriptPath = $PSScriptRoot + '\'
$global:DataFileName = "AZMC_CourtroomData.csv"
$global:SelectedFile = ""
$global:FileLoaded = $false
$global:NumOfRooms = 0

$global:Shell01 = "" 
$global:Shell02 = ""
$global:Shell03 = ""
$global:Shell03Len = 30
$global:logLevel = 1
$global:roomDefaults = $null




function fClear($ms = 100)
{
    clear
    start-sleep -Milliseconds 100
}

fClear
Write-Host -f Yellow "Hi!`nI'm Jonks, your friendly neighborhood PowerShell script.`n`n`n"


function fatal
{
    Read-Host -Prompt "`n`nPress any key to exit"
    exit
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


# Import Libs
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
#####################################################################################################################

# Logging & Debug

if(-not(Test-Path -path "PS_Logs"))
{
    New-Item -Path . -Name "PS_Logs" -ItemType "directory"
}

function fGetDateTime
{
     return ("{0:yyyy.MM.dd_HH.mm.ss.fff}" -f (Get-Date))
}


$global:logFileName = fGetDateTime
$global:logFileName += ".log"

New-Item -Path .\PS_Logs -ItemType "file" -Name $global:logFileName

function fLog ([string]$s)
{
    $dt = fGetDateTime
    $s = "{0}  -  {1}" -f [string]$dt, $s
    $s >> .\PS_Logs\$global:logFileName
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
            Write-Host -b darkgray -f red $s
        }
    }
    else
    {
        if($global:debug)
        {
            Write-Host -b darkgray -f green $s
        }
    }
}






#####################################################################################################################
#####################################################################################################################


class Courtroom
{
    [int]$Index
    [string]$RoomName
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
}


function parseLine($c, $i, $row)
{
 
    # the index is the key for $global:rooms dict entries
    # should also align with the excel spreadsheet row numbers
    $c.Index = $i           

    # should probably name every room uniquely, in case we want to create a $global:roomsByName hashtable
    $c.RoomName = $row | Select-object -ExpandProperty Room_Name
    $c.FacilityName = $row | select-object -ExpandProperty Facility_Name
    $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

    # only 1 Proc_IP and 1 LPZ per room
    $c.Processor_IP = $row | Select-object -ExpandProperty Processor_IP
    $c.FileName_LPZ = $row | Select-object -ExpandProperty FileName_LPZ

    # multiple panel_IPs and multiple fileName_VTZs are possible
    $c.Panel_IP = $row | Select-object -ExpandProperty Panel_IP
    $c.FileName_VTZ = $row | Select-object -ExpandProperty FileName_VTZ

    # x 1
    $c.ReporterWebSvc_IP = $row | Select-object -ExpandProperty IP_ReporterWebSvc
    $c.Wyrestorm_IP = $row | Select-Object -ExpandProperty IP_WyrestormCtrl

    # x Multiple
    $c.FixedCam_IP = $row | Select-object -ExpandProperty IP_FixedCams        
    $c.DSP_IP = $row | Select-object -ExpandProperty IP_DSPs
    $c.RecorderSvr_IP = $row | Select-object -ExpandProperty IP_Recorders
    
    # x 1
    $c.DVD_IP = $row | Select-Object -ExpandProperty IP_DVDPlayer
    $c.MuteGW_IP = $row | Select-Object -ExpandProperty IP_AudicueGW
    
    # x Multiple
    $c.PTZCam_IP = $row | Select-Object -ExpandProperty IP_PTZCams 
    
    
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
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count
        fErr ("Import: File opened successfully.") $false
        fErr ("Import: File name - {0}" -f $fileName) $false
        fErr ("Import: Found {0} data lines to parse." -f $sheetLen) $false
    }
    else
    {
        fErr ("Import: File import failed.") $true
    }

    $global:rooms = @{}
    $global:roomsByName = @{}

    $i = 1
      
    foreach($row in $sheet)
    {     
        $i++

        try
        {
            $c = new-object -TypeName Courtroom
            if($i -eq 2)
            {
                $global:roomDefaults = new-object -TypeName Courtroom
                $global:roomDefaults = parseLine $c $i $row  
                fErr ("Import: Parsed defaults line {0:d3}" -f $i) $false
            }
            else
            {
                $c = parseLine $c $i $row

                $rooms[$c.Index] = $c
                # $roomsByName[$c.RoomName] = $rooms[$c.Index]

                $global:NumOfRooms++
                fErr ("Import: Parsed data line {0:d3}" -f $i) $false
            }
        }
        catch
        {
            fErr ("Import: File import failed for line {0:d3}." -f $i) $true
        }
    }
    if($NumOfRooms -gt 0)
    {
        $global:FileLoaded = $true

        fErr ("Import: File import success. {0} rooms parsed." -f $NumOfRooms) $false
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

######################################################################################################
######################################################################################################
######################################################################################################

function removeWhitespace([string]$s)
{
    return ($s -replace ' ','')
}

function showAllRooms
{
    Write-Host "`n`n"
    foreach($k in $global:rooms.Keys)
    {
        $s = "{0:d3}. {1}" -f $k, $global:rooms[$k].roomname
        Write-Host $s
    }
    Write-Host "`n`n"
}

function getRangeOfRooms
{
    # fClear
    Write-Host -f Yellow "Which rooms do you want to target? (e.g. '3,9-12,17,16')`n`n"
    Write-Host -f Yellow "  *  to target all rooms."
    Write-Host -f Yellow "  b  to go back.`n"
    
    return (Read-Host)
}

function verifyRangeChars([string]$s)
{
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

function fIPTSend([System.Guid]$SessID, [System.Int32]$IPID, [string]$sub, [string]$node, $target)
{
    $ipaddr = $sub+"."+$node

    # remove double dots .. (convenience)
    while(".." -in $ipaddr)
    {
        $ipaddr = $ipaddr.replace("..", ".")
    }

    $AddPeer = "AddP {0:X} {1}" -f $IPID, $ipaddr
    try
    {
        $response = Invoke-CrestronSession $SessID -Command ("{0}" -f $AddPeer)
        fErr ("ProcIPT: Successfully sent IPID {0} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $false
    }
    catch
    {
        fErr ("ProcIPT: Failed to commit IPID {0} IPAddr {1} to the processor in room {2:d3}." -f $ipid, $ipaddr, $target) $true
    }
}

function fIPT ([System.Guid]$SessID, [System.Int32]$IPID, [string]$sub, [string]$node, [string]$def, $target)
{
    # if the node address value is "0", skip
    if($node -ieq "0")
    {
        return
    }
    # elseif the node address value is populated on the spreadsheet, use it
    elseif($node)
    {
        foreach($n in $node)
        {
            fIPTSend $SessID, $IPID, $sub, $n
            $IPID ++
        }
    }
    # if not, check the $global:roomDefaults value
    elseif($def)
    {
        foreach($n in $def)
        {
            fIPTSend $SessID, $IPID, $sub, $n
            $IPID ++
        }
    }
    # if both are null / empty / zero, just skip this line
    else
    {
        return
    }
}

function sendProcIPT
{    
    showAllRooms
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
        
        # Connect to Device w/o SSH
        try
        {
            $SessID = Open-CrestronSession -Device $r.Processor_IP # -ErrorAction SilentlyContinue
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
            fErr ("ProcIPT: The {0:d3} processor failed to respond with a valid hostname." -f $target) $True
            fErr ("ProcIPT: Attetmpting to continue with loading the IP table.") $True
        }

        # Send Crestron Program to Processor Secure
        # Send-CrestronProgram -ShowProgress -Device $ProcIP -LocalFile $LPZPath

        $d = $global:roomDefaults

        # FTR ReporterWebSvc
        fIPT $SessID 0x05 $sub $r.ReporterWebSvc_IP $d.ReporterWebSvc_IP 

        # Wyrestorm Ctrl
        fIPT $SessID 0x06 $sub $r.WyrestormCtrl_IP $d.WyrestormCtrl_IP             

        # Fixed Cams
        fIPT $SessID 0x07 $sub $r.FixedCams_IP $d.FixedCams_IP
 
        # DSPs
        fIPT $SessID 0x0d $sub $r.DSP_IP $d.DSP_IP

        # FTR Recorders
        fIPT $SessID 0x18 $sub $r.RecorderSvr_IP $d.RecorderSvr_IP

        # DVD Player
        fIPT $SessID 0x1a $sub $r.DVD_IP $d.DVD_IP

        # Mute Gateways
        fIPT $SessID 0x20 $sub $r.MuteGW_IP $d.MuteGW_IP

        #PTZ Cams
        $ipid = 0x23
        if($r.PTZCam_IP -ieq "0")
        {

        }
        elseif($r.PTZCam_IP)
        {
            foreach($n in $r.PTZCam_IP)
            {
                fIPTSend $SessIP $ipid $sub $n
                $ipid++
                fIPTSend $SessIP $ipid $sub $n
                $ipid++
            }
        }
        elseif($d.PTZCam_IP)
        {
            foreach($n in $d.PTZCam_IP)
            {
                fIPTSend $SessIP $ipid $sub $n
                $ipid++
                fIPTSend $SessIP $ipid $sub $n
                $ipid++
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





######################################################################################################
######################################################################################################
######################################################################################################



function Shell01
{
    [string]$data = ""

    $data += "`n`n"

    $data += "Please make a selection.`n"
    $data += "`n"

    $global:Shell01 = $data
}

function Shell02
{
    [string]$data = ""
    $data += "a) load a .csv file`n"
    $data += "b) load Crestron processors IP table`n"
    $data += "c) load Crestron panels IP table`n"
    $data += "i) info`n"
    $data += "x) exit`n"

    $data += "`n"

    $global:Shell02 = $data
}

function Shell03([string]$s)
{
    $s = $global:Shell03 + $s 
    [string[]]$data = $s.Split("`n")
    if($data.Length > $global:Shell03Len)
    {
        $data = $data[(-$global:Shell03Len)..1]
    }
    $data += ">"
    $ofs = ""
    $global:Shell03 = "$data"
}

function showInfo
{
    $s = ""
    if($FileLoaded -eq $false)
    {
        $s1 = "[no file loaded]"
    }
    else
    {
        $s1 = $SelectedFile
    }
    [string]$data = ""
    $data += (".csv file: " + $s1 + "`n") 
    $data += ("num of rooms loaded: " + $NumOfRooms + "`n")

    Shell03 $data
}

function updateShell([bool]$clear)
{
    if($clear)
    {
        fClear
        Shell03
    }
    Write-Host -f Yellow $global:Shell01
    Write-Host -f Green $global:Shell02
    Write-Host -f White $global:Shell03
}


function getCommand
{
    [string]$choice = Read-Host

    if($choice -ieq 'a')
    {
        $result = selectAndImport
        continueScript $true

    }
    elseif($choice -ieq 'b')
    {
        sendProcIPT
        continueScript 
    }
    elseif($choice -ieq 'c')
    {

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
        continueScript $true
    }
}

function setShellAll
{
    Shell01
    Shell02
    Shell03
}

function continueScript([bool]$clear)
{
    setShellAll
    updateShell #($clear)
    getCommand
}

continueScript($true)
