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


Write-Host -ForegroundColor Cyan @"

FTR's Magic Crestron Configuration Script

"@

# MyShell


function fClear($ms = 100)
{
    clear
    start-sleep -Milliseconds 100
}

fClear
Write-Host -ForegroundColor Yellow "Hi!`nI'm Jonks, your friendly neighborhood PowerShell script.`n`n`n"


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
            Write-Host -ForegroundColor Yellow 
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
            Write-Host -ForegroundColor Yellow 
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
            Write-Host -ForegroundColor Yellow 
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
#####################################################################################################################


class Courtroom
{
    [int]$Index
    [string]$CommentLine
    [string]$RoomName
    [string]$FacilityName
    [string]$Subnet
    [string]$Processor_IP
    [string]$FileName_LPZ
    [string[]]$Panel_IP
    [string[]]$FileName_VTZ
    [string]$ReporterWebSvc_IP
    [string]$Wyrestorm_IP
    [string[]]$FixedCam_IP
    [string[]]$DSP_IP
    [string[]]$RecorderSvr_IP
    [string]$DVD_IP
    [string]$Audicue_IP
    [string[]]$PTZCam_IP

}



function importFile([string]$fileName)
{
   
    if([System.IO.File]::Exists("$fileName"))
    {
        $global:sheet = Import-csv $fileName
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count
        Shell03 "Ok, I imported the file $fileName.`nThere are $sheetLen rooms in the list.`n`n"
    }
    else
    {
        err 3
        return ""
    }

    $global:rooms = @{}
    $global:roomsByName = @{}

    $i = 1
    foreach($row in $sheet)
    {
        # $i is initialized to 1. As this value will be the dict key, and we want that to match the .csv line number (for convenience),
        #     we will start the loop by adding 1.
        #     Ergo, the first dict entry will be     (2, $c) 
     
        $i += 1

        try
        {
            $c = new-object -TypeName Courtroom
 
            $c.Index = $i           
            $c.CommentLine = $row | Select-object -ExpandProperty Ignore_Line

            $c.RoomName = $row | Select-object -ExpandProperty Room_Name
            $c.FacilityName = $row | select-object -ExpandProperty Facility_Name

            $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

            $c.Processor_IP = $row | Select-object -ExpandProperty Processor_IP
            $c.FileName_LPZ = $row | Select-object -ExpandProperty FileName_LPZ
            $c.Panel_IP = $row | Select-object -ExpandProperty Panel_IP
            $c.Panel_IP = $c.Panel_IP[0].split('~')
            $c.FileName_VTZ = $row | Select-object -ExpandProperty FileName_VTZ
            $c.FileName_VTZ = $c.FileName_VTZ[0].split('~')

            $c.ReporterWebSvc_IP = $row | Select-object -ExpandProperty IP_ReporterWebSvc

            $c.Wyrestorm_IP = $row | Select-Object -ExpandProperty IP_WyrestormCtrl

            $c.FixedCam_IP = $row | Select-object -ExpandProperty IP_FixedCams
            $c.FixedCam_IP = $c.FixedCam_IP[0].split('~')
            
            $c.DSP_IP = $row | Select-object -ExpandProperty IP_DSPs
            $c.DSP_IP = $c.DSP_IP[0].split('~')

            $c.RecorderSvr_IP = $row | Select-object -ExpandProperty IP_Recorders
            $c.RecorderSvr_IP = $c.RecorderSvr_IP[0].split('~')

            $c.DVD_IP = $row | Select-Object -ExpandProperty IP_DVDPlayer

            $c.Audicue_IP = $row | Select-Object -ExpandProperty IP_AudicueGW

            $c.PTZCam_IP = $row | Select-Object -ExpandProperty IP_PTZCams 
            $c.PTZCam_IP = $c.PTZCam_IP[0].split('~')

            $rooms[$c.Index] = $c
            $roomsByName[$c.RoomName] = $rooms[$c.Index]

            $global:NumOfRooms += 1
        }
        catch
        {
            Write-Host -ForegroundColor Red ("import failed for line " + $i)
        }
    }
    if($NumOfRooms -gt 0)
    {
        $global:FileLoaded = $true
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

function updateShell
{
    fClear
    Write-Host -ForegroundColor Yellow $global:Shell01
    Write-Host -ForegroundColor Green $global:Shell02
    Write-Host -ForegroundColor White $global:Shell03
}


function getCommand
{
    [string]$choice = Read-Host

    if($choice -ieq 'a')
    {
        selectAndImport
        continueScript
    }
    elseif($choice -ieq 'b')
    {
    
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
}

function setShellAll
{
    Shell01
    Shell02
    Shell03
}

function continueScript
{
    setShellAll
    updateShell
    getCommand
}

continueScript
