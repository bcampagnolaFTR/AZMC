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


$global:logFileName = fGetDateTime + ".log"

New-Item -Path .\PS_Logs -ItemType "file" -Name $global:logFileName

function fLogWrite ($s)
{
    fGetDateTime$s >> .\PS_Logs\$global:logFileName
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
    # [string]$CommentLine
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


function parseLine()
{
 
            $c.Index = $i           
            # $c.CommentLine = $row | Select-object -ExpandProperty Ignore_Line

            $c.RoomName = $row | Select-object -ExpandProperty Room_Name
            $c.FacilityName = $row | select-object -ExpandProperty Facility_Name

            $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

            $c.Processor_IP = $row | Select-object -ExpandProperty Processor_IP
            $c.FileName_LPZ = $row | Select-object -ExpandProperty FileName_LPZ
            if($c.FileName_LPZ[0].length > 0)
            {
                $c.localLPZFile = $True
            }

            $c.Panel_IP = $row | Select-object -ExpandProperty Panel_IP
            $c.Panel_IP = $c.Panel_IP[0].split('~')
            $c.FileName_VTZ = $row | Select-object -ExpandProperty FileName_VTZ
            $c.FileName_VTZ = $c.FileName_VTZ[0].split('~')
            if($c.FileName_VTZ[0].length > 0)
            {
                $c.localVTZFile = $True
            }

            $c.ReporterWebSvc_IP = $row | Select-object -ExpandProperty IP_ReporterWebSvc

            $c.Wyrestorm_IP = $row | Select-Object -ExpandProperty IP_WyrestormCtrl

            $c.FixedCam_IP = $row | Select-object -ExpandProperty IP_FixedCams
            $c.FixedCam_IP = $c.FixedCam_IP[0].split('~')
            
            $c.DSP_IP = $row | Select-object -ExpandProperty IP_DSPs
            $c.DSP_IP = $c.DSP_IP[0].split('~')

            $c.RecorderSvr_IP = $row | Select-object -ExpandProperty IP_Recorders
            $c.RecorderSvr_IP = $c.RecorderSvr_IP[0].split('~')

            $c.DVD_IP = $row | Select-Object -ExpandProperty IP_DVDPlayer

            $c.MuteGW_IP = $row | Select-Object -ExpandProperty IP_AudicueGW

            $c.PTZCam_IP = $row | Select-Object -ExpandProperty IP_PTZCams 
            $c.PTZCam_IP = $c.PTZCam_IP[0].split('~')
}
function importFile([string]$fileName)
{
   
    if([System.IO.File]::Exists("$fileName"))
    {
        $global:sheet = Import-csv $fileName
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count
        fErr "File import successful. File name: $fileName." $false
        fErr "Number of data lines found: $sheetLen" $false
    }
    else
    {
        fErr "File import failed. File name: $filename." $true
    }

    $global:rooms = @{}
    $global:roomsByName = @{}


    $c = New-Object -TypeName Courtroom


    $i = 1

    

    foreach($row in $sheet)
    {
        # $i is initialized to 2. As this value will be the dict key, and we want that to match the .csv line number (for convenience),
        #     we will start the loop by adding 1.
        #     Ergo, the first dict entry will be     (2, $c) 
     
        $i += 1

        try
        {
            $c = new-object -TypeName Courtroom
 
            $c.Index = $i           
            # $c.CommentLine = $row | Select-object -ExpandProperty Ignore_Line

            $c.RoomName = $row | Select-object -ExpandProperty Room_Name
            $c.FacilityName = $row | select-object -ExpandProperty Facility_Name

            $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

            $c.Processor_IP = $row | Select-object -ExpandProperty Processor_IP
            $c.FileName_LPZ = $row | Select-object -ExpandProperty FileName_LPZ
            if($c.FileName_LPZ[0].length > 0)
            {
                $c.localLPZFile = $True
            }

            $c.Panel_IP = $row | Select-object -ExpandProperty Panel_IP
            $c.Panel_IP = $c.Panel_IP[0].split('~')
            $c.FileName_VTZ = $row | Select-object -ExpandProperty FileName_VTZ
            $c.FileName_VTZ = $c.FileName_VTZ[0].split('~')
            if($c.FileName_VTZ[0].length > 0)
            {
                $c.localVTZFile = $True
            }

            $c.ReporterWebSvc_IP = $row | Select-object -ExpandProperty IP_ReporterWebSvc

            $c.Wyrestorm_IP = $row | Select-Object -ExpandProperty IP_WyrestormCtrl

            $c.FixedCam_IP = $row | Select-object -ExpandProperty IP_FixedCams
            $c.FixedCam_IP = $c.FixedCam_IP[0].split('~')
            
            $c.DSP_IP = $row | Select-object -ExpandProperty IP_DSPs
            $c.DSP_IP = $c.DSP_IP[0].split('~')

            $c.RecorderSvr_IP = $row | Select-object -ExpandProperty IP_Recorders
            $c.RecorderSvr_IP = $c.RecorderSvr_IP[0].split('~')

            $c.DVD_IP = $row | Select-Object -ExpandProperty IP_DVDPlayer

            $c.MuteGW_IP = $row | Select-Object -ExpandProperty IP_AudicueGW

            $c.PTZCam_IP = $row | Select-Object -ExpandProperty IP_PTZCams 
            $c.PTZCam_IP = $c.PTZCam_IP[0].split('~')

            $rooms[$c.Index] = $c
            $roomsByName[$c.RoomName] = $rooms[$c.Index]

            $global:NumOfRooms += 1

        }
        catch
        {
            Write-Host -f Red ("import failed for line " + $i)
        }
    }
    if($NumOfRooms -gt 0)
    {
        $global:FileLoaded = $true

        return "It worked."
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



function fIPT ($SessID, $IPID, $sub, $node)
{
    $AddPeer = "AddP {0:X} {1}{2}" -f $IPID, $sub, $node
    $response = Invoke-CrestronSession $SessID -Command "$AddPeer"
}


function sendProcIPT
{    
    showAllRooms
    $targets = decodeRange(getRangeOfRooms)
    # $targets = [System.Int32[]]$targets
    if(-not $targets)
    {
        ferr 0 "No rooms were targeted." $true
        return
    }
    
    # Write-Host -b red "Target rooms: $targets"
    foreach($target in $targets)
    {
        # check for
        if($target -notin $global:rooms.keys)
        {
            fErr $target "Target room not in list of rooms." $True
            continue
        }
        $r = $global:rooms[$target]
        $sub = $r.Subnet
        
        # Connect to Device w/o SSH
        try
        {
            $SessID = Open-CrestronSession -Device $r.Processor_IP # -ErrorAction SilentlyContinue
            fErr $target "Open-CrestronSession successful. SessionID: {0}" -f $SessID, $False
        }
        catch
        {
            fErr $target "Open-CrestronSession failed." $True  
            fErr $target "Check the IP address in the spreadsheet. Be sure that you can ping the device from this machine.`n" $True
            continue 
        }

        # Hostname
        try
        {
            $hostnameResponse = Invoke-CrestronSession $SessID "hostname"
            $deviceHostname = [regex]::Match($hostnameResponse, "(?<=Host\sName:\s)[\w-]+").value
            fErr $target "Got hostname: $deviceHostname" $false
        }
        catch
        {
            fErr $target "Hostname failed to resolve: $hostnameResponse" $True
        }


        # Send Crestron Program to Processor Secure
        # Send-CrestronProgram -ShowProgress -Device $ProcIP -LocalFile $LPZPath



        # FTR ReporterWebSvc
        fIPT $SessID, 0x05, $sub, $r.ReporterWebSvc_IP

        # Wyrestorm Ctrl
        fIPT $SessID, 0x06, $sub, $r.WyrestormCtrl_IP

        # Fixed Cams
        $IPID = 0x07
        foreach($node in $r.FixedCams_IP)
        {
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
        }

        # DSPs
        $IPID = 0x0d
        foreach($node in $r.DSP_IP)
        {
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
        }

        # FTR Recorders
        $IPID = 0x18
        foreach($node in $r.RecorderWebSvr_IP)
        {
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
        }

        # DVD Player
        fIPT $SessID, 0x1a, $sub, $r.DVD_IP

        # Mute Gateways
        $IPID = 0x20
        foreach($node in $r.MuteGW_IP)
        {
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
        }

        #PTZ Cams
        $IPID = 0x23
        foreach($node in $r.PTZCam_IP)
        {
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
            fIPT $SessID, $IPID, $sub, $node
            $IPID += 1
        }
        
        # Reset Program
        try
        {
            Invoke-CrestronSession $SessID 'progreset -p:01'
            fErr $target, "Restarting program.", $False
        }
        catch
        {
            fErr $target, "Failed to restart program.", $True
        }
        finally{}

        # Close Session
        try
        {
            Close-CrestronSession $SessID
            fErr $target, "Close-CrestronSession successful. $SessID", $False
        }
        catch
        {
            fErr $target, "Close-CrestronSession failed. You should probably restart the PowerShell script.", $True
        }
        finally{}
    }
}

function sendPanelIPT
{
    showAllRooms
    [int[]]$targets = decodeRange(getRangeOfRooms)
    #foreach($target     
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
